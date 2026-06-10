"""
Routeur d'images médicales

Déplace les fichiers déposés par le système d'acquisition dans le bon dossier
patient sur le partage réseau, puis insère une entrée dans le sous-formulaire
SFDoc du formulaire Access actif, via l'automatisation de l'interface win32com.

Chaîne de traitement :
  PollingObserver → file_queue → Worker → déplacement du fichier → insertion interface (win32com)

Dépendances : watchdog, pywin32, pythoncom, pystray, Pillow, psutil
"""

import os
import pythoncom
import queue
import shutil
import subprocess
import sys
import threading
import time
import ctypes
import logging
from datetime import datetime
from pathlib import Path
from watchdog.observers.polling import PollingObserver as Observer
from watchdog.events import FileSystemEventHandler

try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

try:
    import pystray
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False

import win32api
import win32event
import winerror
import psutil


# Configuration
BOX_NAME = "Studiovision"

SOURCE_DIR   = Path(r"??")  # Dossier de dépôt des acquisitions
ORPHAN_DIR   = Path(r"??")  # Destination des fichiers non rattachés
DEST_PHOTOS  = Path(r"??")  # Racine de l'archive photo réseau

STUDIO_VISION_CMD = [
    r"C:\Studiov2000-OM\svprog\msaccess.exe",
    "/runtime",
    r"C:\Studiov2000-OM\svprog\Ophprog.mde",
    "/wrkgrp",
    r"C:\Studiov2000-OM\config\system.mdw",
    "/User",
    "/Pwd",
    "/X",
    "demarrage",
]

WATCHED_EXTENSIONS: set[str] = {
    ".jpg", ".jpeg", ".jfif",
    ".png", ".bmp",
    ".tif", ".tiff",
    ".dcm",
    ".pdf", ".rtf", ".doc", ".docx", ".odt",
}

EXAM_DESCRIPTION: dict[str, str] = {
    ".jpg":  "Image",
    ".jpeg": "Image",
    ".jfif": "Image",
    ".png":  "Image",
    ".bmp":  "Image",
    ".tif":  "OCT",
    ".tiff": "OCT",
    ".dcm":  "DICOM",
    ".pdf":  "Document",
    ".rtf":  "Document",
    ".doc":  "Document",
    ".docx": "Document",
    ".odt":  "Document",
}

FILE_LOCK_RETRY_DELAY:  int = 3
FILE_LOCK_MAX_ATTEMPTS: int = 15
PATIENT_POLL_INTERVAL:  int = 3
PATIENT_WAIT_TIMEOUT:   int = 900
CATCHUP_INTERVAL:       int = 120

ACCESS_FIELD_CODE   = "Code patient"
ACCESS_FIELD_NOM    = "NOM"
ACCESS_FIELD_PRENOM = "Prénom"

SFDOC_SUBFORM_NAME = "SFDoc"
_AC_SUBFORM        = 112  # Constante Access de type de contrôle pour les sous-formulaires

_GUI_PRE_INSERT_DELAY = 0.3  # Secondes entre le déplacement du fichier et l'insertion interface
_UI_POST_INSERT_DELAY = 0.5  # Secondes entre l'insertion et Requery/MoveLast


# Journalisation
_LOG_DIR = Path(os.path.expanduser("~")) / "studiovision"
_LOG_DIR.mkdir(parents=True, exist_ok=True)
_LOG_FILE = _LOG_DIR / "image_router.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  [%(threadName)s]  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(_LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("image_router")


def _get_doctor_log_path() -> Path:
    return _LOG_DIR / f"doctor_report_{datetime.now().strftime('%Y-%m-%d')}.txt"


def log_doctor(message: str) -> None:
    """Ajoute une ligne horodatée au journal quotidien destiné au médecin."""
    timestamp = datetime.now().strftime("%H:%M")
    try:
        with open(_get_doctor_log_path(), "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {message}\n")
    except Exception as exc:
        log.warning(f"Impossible d'écrire dans le journal médecin : {exc}")


# État global
_NETWORK_SHARE_POLL = 10
_ICON_SIZE    = 64
_COLOR_READY  = (30, 144, 255)
_COLOR_ACTIVE = (50, 205, 50)

_icon: "pystray.Icon | None" = None
_status_text: str             = "Démarrage..."
_stop_event: threading.Event  = threading.Event()
_mutex_handle                 = None


# Utilitaires de la barre des tâches (system tray)
def _make_icon(color: tuple) -> "Image.Image":
    img  = Image.new("RGBA", (_ICON_SIZE, _ICON_SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    margin = 4
    draw.ellipse(
        [margin, margin, _ICON_SIZE - margin, _ICON_SIZE - margin],
        fill=color,
    )
    return img


def _set_status(text: str, processing: bool = False) -> None:
    global _status_text
    _status_text = text
    if _icon is not None:
        try:
            _icon.icon = _make_icon(_COLOR_ACTIVE if processing else _COLOR_READY)
            _icon.update_menu()
        except Exception as exc:
            log.debug(f"Échec de mise à jour de l'icône : {exc}")


def _notify(title: str, message: str = "") -> None:
    if _icon is not None:
        try:
            _icon.notify(message if message else title, title)
        except Exception as exc:
            log.debug(f"Échec de la notification : {exc}")


def _open_log_file(icon, item) -> None:  # noqa: ARG001
    try:
        os.startfile(str(_LOG_FILE))
    except Exception as exc:
        log.warning(f"Impossible d'ouvrir le fichier de log : {exc}")


def _open_doctor_log(icon, item) -> None:  # noqa: ARG001
    try:
        os.startfile(str(_get_doctor_log_path()))
    except Exception as exc:
        log.warning(f"Impossible d'ouvrir le journal médecin : {exc}")


def _quit(icon, item) -> None:  # noqa: ARG001
    log.info("Fermeture demandée depuis le menu de la barre des tâches.")
    _stop_event.set()
    icon.stop()


# Partage réseau
def wait_for_network_share() -> None:
    """Bloque jusqu'à ce que SOURCE_DIR soit accessible."""
    is_network = str(SOURCE_DIR).startswith("\\\\") or str(SOURCE_DIR).startswith("//")
    if not is_network:
        return
    attempt = 0
    while not SOURCE_DIR.is_dir():
        attempt += 1
        log.warning(
            f"Partage réseau inaccessible : {SOURCE_DIR}  "
            f"(tentative {attempt}, nouvel essai dans {_NETWORK_SHARE_POLL}s)"
        )
        time.sleep(_NETWORK_SHARE_POLL)
    if attempt:
        log.info(f"Partage réseau accessible après {attempt} tentative(s) : {SOURCE_DIR}")


# Résolution du dossier patient
def build_patient_relative_path(patient_code: str, last_name: str, first_name: str) -> str:
    """
    Renvoie le chemin de dossier relatif d'un patient.
    Format : <2premiers>.000\\<code><3derniers>.<3premiers>
    Exemple : code=1758511228, ABCDEF, DEFGH → "17.000\\1758511228abc.def"
    """
    prefix  = patient_code[:2]
    last_3  = last_name[:3].lower()
    first_3 = first_name[:3].lower()
    return f"{prefix}.000\\{patient_code}{last_3}.{first_3}"


def resolve_patient_folder(patient: dict) -> Path | None:
    """Résout et crée le dossier patient absolu. Renvoie None en cas d'échec."""
    try:
        rel    = build_patient_relative_path(patient["code"], patient["nom"], patient["prenom"])
        folder = DEST_PHOTOS / rel
        folder.mkdir(parents=True, exist_ok=True)
        log.info(f"Dossier patient résolu : {folder}")
        return folder
    except Exception as exc:
        log.error(f"Impossible de résoudre/créer le dossier patient : {exc}")
        return None


# Access COM — lecture du patient actif
def get_active_patient() -> dict | None:
    """Renvoie le code, le nom et le prénom du patient actif depuis le formulaire Access."""
    if not WIN32_AVAILABLE:
        return None
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            return None

        target: set[str] = {ACCESS_FIELD_CODE, ACCESS_FIELD_NOM, ACCESS_FIELD_PRENOM}
        data:   dict     = {}

        for i in range(form.Controls.Count):
            ctrl = form.Controls(i)
            try:
                name = str(ctrl.Name)
                if name in target:
                    data[name] = ctrl.Value
            except Exception:
                pass

        if not target.issubset(data.keys()):
            return None

        return {
            "code":   str(data[ACCESS_FIELD_CODE]),
            "nom":    str(data[ACCESS_FIELD_NOM]),
            "prenom": str(data[ACCESS_FIELD_PRENOM]),
        }
    except Exception as exc:
        log.debug(f"Erreur COM pendant la lecture du patient : {exc}")
        return None


# Access COM — recherche du sous-formulaire SFDoc
def _find_sfdoc(form) -> object | None:
    """Recherche récursivement le sous-formulaire SFDoc dans l'arborescence des contrôles du formulaire."""
    for i in range(form.Controls.Count):
        ctrl = form.Controls(i)
        try:
            if ctrl.ControlType != _AC_SUBFORM:
                continue
            if ctrl.Name == SFDOC_SUBFORM_NAME:
                return ctrl.Form
            found = _find_sfdoc(ctrl.Form)
            if found is not None:
                return found
        except Exception:
            pass
    return None


# Insertion dans l'interface
def gui_insert_document(patient: dict, relative_path: str, description: str) -> bool:
    """
    Insère une nouvelle entrée dans le sous-formulaire SFDoc via win32com.
    Renvoie True en cas de succès, False en cas d'erreur COM.
    """
    if not WIN32_AVAILABLE:
        log.error("win32com indisponible — insertion interface ignorée.")
        return False

    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            log.warning("Insertion interface ignorée : aucun formulaire actif dans Access.")
            return False

        current = get_active_patient()
        if not current or current["code"] != patient["code"]:
            log.warning(
                f"Insertion interface annulée : le patient a changé "
                f"(attendu={patient['code']}, "
                f"courant={current['code'] if current else 'aucun'})."
            )
            return False

        sfdoc = _find_sfdoc(form)
        if sfdoc is None:
            log.error(f"Sous-formulaire '{SFDOC_SUBFORM_NAME}' introuvable — insertion interface annulée.")
            return False

        rs = sfdoc.Recordset
        rs.AddNew()

        def _set_field(name: str, value) -> None:
            try:
                rs.Fields(name).Value = value
            except Exception as exc:
                log.warning(f"Échec d'écriture du champ '{name}' : {exc}")

        _set_field("code patient",  int(patient["code"]))
        _set_field("Date",          datetime.now())
        _set_field("DESCRIPTIONS",  description)
        _set_field("TEXTE",         relative_path)
        _set_field("Photo externe", relative_path)
        _set_field("TypeVW",        99)

        try:
            rs.Fields("NumDocExterne").Value = None
        except Exception:
            pass

        rs.Update()
        log.info(
            f"Insertion interface OK : patient={patient['code']} "
            f"desc='{description}' chemin='{relative_path}'"
        )

        time.sleep(_UI_POST_INSERT_DELAY)

        try:
            sfdoc.Requery()
            log.info(f"Requery() sur '{SFDOC_SUBFORM_NAME}'.")
        except Exception as exc:
            log.warning(f"Échec de Requery() : {exc}")
            try:
                sfdoc.Refresh()
                log.info(f"Refresh() de repli sur '{SFDOC_SUBFORM_NAME}'.")
            except Exception as exc2:
                log.warning(f"Le Refresh() de repli a aussi échoué : {exc2}")

        try:
            sfdoc.Recordset.MoveLast()
        except Exception as exc:
            log.debug(f"Échec de MoveLast() : {exc}")

        return True

    except Exception as exc:
        log.error(f"Échec de l'insertion interface : {exc}")
        return False


# Utilitaires fichiers
def wait_for_file(file: Path) -> bool:
    """Attend que le fichier ne soit plus verrouillé. Renvoie False si le délai est dépassé."""
    for attempt in range(1, FILE_LOCK_MAX_ATTEMPTS + 1):
        try:
            with file.open("rb"):
                return True
        except (PermissionError, OSError):
            log.debug(f"Fichier verrouillé ({attempt}/{FILE_LOCK_MAX_ATTEMPTS}), nouvel essai...")
            time.sleep(FILE_LOCK_RETRY_DELAY)
    log.error(f"Fichier toujours verrouillé après {FILE_LOCK_MAX_ATTEMPTS} tentatives : {file}")
    return False


def move_file(source: Path, dest_folder: Path, label: str = "") -> Path | None:
    """Déplace source vers dest_folder, en résolvant les conflits de nom par un suffixe horodaté."""
    dest_folder.mkdir(parents=True, exist_ok=True)
    dest = dest_folder / source.name
    if dest.exists():
        ts   = int(time.time())
        dest = dest_folder / f"{source.stem}_{ts}{source.suffix}"
        log.info(f"Conflit de nom résolu — renommé en {dest.name}")
    try:
        shutil.move(str(source), str(dest))
        tag = f"[{label}]  " if label else ""
        log.info(f"{tag}{source.name} → {dest}")
        return dest
    except Exception as exc:
        log.error(f"Échec du déplacement : {exc}")
        return None


def orphan_file(file: Path) -> None:
    """Déplace un fichier non traitable vers le dossier des orphelins."""
    log.warning(f"Mise en orphelin : {file.name}")
    move_file(file, ORPHAN_DIR, label="ORPHAN")


def prevent_sleep() -> None:
    """Empêche Windows de se mettre en veille pendant que le routeur est actif."""
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            0x80000000 |  # ES_CONTINUOUS
            0x00000001    # ES_SYSTEM_REQUIRED
        )
        log.info("Prévention de la veille active.")
    except Exception as exc:
        log.warning(f"Impossible de définir l'état d'exécution : {exc}")


# Démarrage / analyse de rattrapage
def _scan_source_for_missed_files(file_queue: queue.Queue) -> None:
    """Met en file les fichiers de SOURCE_DIR manqués pendant une interruption."""
    if not SOURCE_DIR.is_dir():
        log.warning(f"Analyse de rattrapage ignorée : {SOURCE_DIR} inaccessible.")
        return

    found: list[Path] = []
    try:
        for item in SOURCE_DIR.rglob("*"):
            if item.is_file() and item.suffix.lower() in WATCHED_EXTENSIONS:
                found.append(item)
    except Exception as exc:
        log.error(f"Erreur pendant l'analyse de rattrapage : {exc}")
        return

    if not found:
        log.info("Analyse de rattrapage : aucun fichier en attente.")
        return

    log.info(f"Analyse de rattrapage : {len(found)} fichier(s) trouvé(s) — mise en file.")
    for f in found:
        file_queue.put(f)
        log.info(f"  Rattrapage → mis en file : {f.name}")


def _catchup_loop(file_queue: queue.Queue) -> None:
    """Ré-analyse périodiquement SOURCE_DIR pour rattraper les fichiers manqués par l'observateur."""
    log.info(f"Fil de rattrapage démarré (intervalle : {CATCHUP_INTERVAL}s).")
    while not _stop_event.is_set():
        for _ in range(CATCHUP_INTERVAL):
            if _stop_event.is_set():
                break
            time.sleep(1)
        if _stop_event.is_set():
            break
        log.debug("Rattrapage : analyse de SOURCE_DIR...")
        _scan_source_for_missed_files(file_queue)
    log.info("Fil de rattrapage arrêté.")


# Fil de traitement (Worker)
def worker(file_queue: queue.Queue) -> None:
    """
    Consomme les fichiers de la file.
    Pour chaque fichier : déverrouillage → attente d'un patient ouvert → déplacement → insertion interface.
    """
    pythoncom.CoInitialize()
    log.info("Fil de traitement démarré.")

    needs_refresh:     bool       = False
    last_patient_code: str | None = None
    burst_count:       int        = 0

    try:
        while True:
            try:
                file: Path = file_queue.get(timeout=1.5)
            except queue.Empty:
                if needs_refresh:
                    log.info("Rafale terminée — tous les fichiers du lot ont été traités.")
                    _notify("Transfert terminé", f"{burst_count} fichier(s) traité(s)")
                    _set_status(f"{BOX_NAME} — Prêt", processing=False)
                    needs_refresh     = False
                    last_patient_code = None
                    burst_count       = 0
                continue
            except Exception as exc:
                log.error(f"Erreur de file : {exc}")
                continue

            log.info(f"Traitement : {file.name} ({file_queue.qsize()} en attente)")

            if burst_count == 0 and not needs_refresh:
                _notify("Transfert en cours", file.name)
            _set_status("Transfert en cours...", processing=True)

            if not file.exists():
                log.warning(f"Fichier disparu avant traitement : {file}")
                file_queue.task_done()
                continue

            if not wait_for_file(file):
                log.error(f"Abandon — verrou de fichier persistant : {file.name}")
                _notify("Erreur", f"Fichier verrouillé : {file.name}")
                file_queue.task_done()
                continue

            # Attendre qu'un patient soit ouvert dans Access
            patient    = None
            start_time = time.monotonic()
            first_log  = True

            while True:
                patient = get_active_patient()
                if patient:
                    break

                elapsed = time.monotonic() - start_time
                if elapsed >= PATIENT_WAIT_TIMEOUT:
                    orphan_file(file)
                    _notify("Fichier orphelin", file.name)
                    log_doctor(
                        f"ATTENTION : fichier '{file.name}' mis en orphelin "
                        f"(aucun patient ouvert après {PATIENT_WAIT_TIMEOUT // 60} min)."
                    )
                    file_queue.task_done()
                    patient = None
                    break

                if first_log:
                    log.info(
                        f"Aucun patient ouvert — attente "
                        f"(délai dans {PATIENT_WAIT_TIMEOUT // 60} min)"
                    )
                    first_log = False

                time.sleep(PATIENT_POLL_INTERVAL)

            if patient is None:
                continue

            log.info(
                f"Patient : {patient['nom']} {patient['prenom']} "
                f"(code {patient['code']})"
            )

            patient_folder = resolve_patient_folder(patient)
            if not patient_folder:
                log.error(f"Impossible de résoudre le dossier du patient {patient['code']}. Mise en orphelin.")
                orphan_file(file)
                _notify("Fichier orphelin", file.name)
                log_doctor(
                    f"ERREUR : impossible de créer le dossier pour '{file.name}' "
                    f"(patient {patient['nom']} {patient['prenom']}, "
                    f"code {patient['code']})."
                )
                file_queue.task_done()
                continue

            dest = move_file(file, patient_folder)
            if dest is None:
                file_queue.task_done()
                continue

            rel_path      = build_patient_relative_path(patient["code"], patient["nom"], patient["prenom"])
            relative_path = f"\\{rel_path}\\{dest.name}"
            description   = EXAM_DESCRIPTION.get(file.suffix.lower(), "Image")

            time.sleep(_GUI_PRE_INSERT_DELAY)

            if gui_insert_document(patient, relative_path, description):
                needs_refresh     = True
                last_patient_code = patient["code"]
                burst_count      += 1
                log.info(f"Entrée insérée : '{dest.name}' → {relative_path}")
                log_doctor(
                    f"SUCCÈS : '{dest.name}' ajouté au dossier de "
                    f"{patient['nom']} {patient['prenom']} (code {patient['code']})."
                )
            else:
                log.error(
                    f"Échec de l'insertion interface pour '{dest.name}' "
                    f"(patient {patient['code']}). Fichier situé : {dest}. Saisie manuelle nécessaire."
                )
                _notify("Erreur insertion", f"'{dest.name}' déplacé mais non inséré — voir le log.")
                log_doctor(
                    f"ERREUR : '{dest.name}' a été déplacé vers {dest} mais N'A PAS PU "
                    f"être inséré pour le patient {patient['nom']} {patient['prenom']} "
                    f"(code {patient['code']}). Saisie manuelle nécessaire."
                )

            file_queue.task_done()

    finally:
        _set_status(f"{BOX_NAME} — Arrêté")
        pythoncom.CoUninitialize()
        log.info("Fil de traitement arrêté.")


# Producteur Watchdog
class ImageProducer(FileSystemEventHandler):
    """Met en file les fichiers nouvellement créés détectés par l'observateur du système de fichiers."""

    def __init__(self, file_queue: queue.Queue) -> None:
        super().__init__()
        self._queue = file_queue

    def on_created(self, event) -> None:
        if event.is_directory:
            return
        file = Path(event.src_path)
        if file.suffix.lower() not in WATCHED_EXTENSIONS:
            return
        log.info(f"Mis en file : {file.name} (taille de la file : {self._queue.qsize() + 1})")
        self._queue.put(file)


# Fil d'observation en arrière-plan
def _run_background(file_queue: queue.Queue) -> None:
    """Démarre et surveille l'observateur du système de fichiers. Se reconnecte en cas de coupure réseau."""
    _RECONNECT_WAIT = 15

    def _start_observer() -> Observer:
        obs = Observer()
        obs.schedule(ImageProducer(file_queue), str(SOURCE_DIR), recursive=True)
        obs.start()
        log.info("Observateur démarré.")
        return obs

    observer = _start_observer()
    _set_status(f"{BOX_NAME} — Prêt", processing=False)

    try:
        while not _stop_event.is_set():
            if not observer.is_alive():
                log.warning("Observateur arrêté (coupure réseau ?). Reconnexion...")
                _set_status(f"{BOX_NAME} — Reconnexion...", processing=False)
                try:
                    observer.stop()
                    observer.join(timeout=5)
                except Exception:
                    pass
                wait_for_network_share()
                log.info(f"Attente de {_RECONNECT_WAIT}s avant redémarrage de l'observateur...")
                time.sleep(_RECONNECT_WAIT)
                observer = _start_observer()
                _set_status(f"{BOX_NAME} — Prêt", processing=False)
            time.sleep(1)
    finally:
        observer.stop()
        observer.join()
        remaining = file_queue.qsize()
        if remaining:
            log.info(f"Vidage des {remaining} fichier(s) restant(s) de la file...")
            file_queue.join()
        log.info("Fil d'arrière-plan arrêté.")
        if _icon is not None:
            _icon.stop()


# Cycle de vie de Studio Vision
_SV_POLL_INTERVAL    = 3   # Secondes entre deux vérifications de présence de msaccess.exe
_SV_STARTUP_TIMEOUT  = 30  # Secondes d'attente de l'apparition de msaccess.exe après lancement


def _get_msaccess_pids() -> set[int]:
    """Renvoie l'ensemble des PID des processus msaccess.exe en cours d'exécution."""
    pids: set[int] = set()
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            if (proc.info["name"] or "").lower() == "msaccess.exe":
                pids.add(proc.info["pid"])
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return pids


def _launch_studio_vision() -> None:
    """
    Lance Studio Vision et suit toute nouvelle instance de msaccess.exe.
    Gère correctement le relais de démarrage en 2 étapes de /runtime.
    Force l'arrêt des processus zombies nouvellement créés à la sortie.
    """
    log.info(f"Lancement de Studio Vision : {' '.join(STUDIO_VISION_CMD)}")

    pids_before: set[int] = _get_msaccess_pids()

    try:
        subprocess.Popen(STUDIO_VISION_CMD)
    except FileNotFoundError:
        log.critical(f"Exécutable Studio Vision introuvable. Arrêt.")
        _stop_event.set()
        return
    except Exception as exc:
        log.error(f"Impossible de lancer Studio Vision : {exc}. Arrêt.")
        _stop_event.set()
        return

    log.info(f"Attente de msaccess.exe (max {_SV_STARTUP_TIMEOUT}s)...")
    deadline = time.monotonic() + _SV_STARTUP_TIMEOUT

    while time.monotonic() < deadline and not _stop_event.is_set():
        if _get_msaccess_pids() - pids_before:
            log.info("Studio Vision démarre...")
            break
        time.sleep(1)
    else:
        if not _stop_event.is_set():
            log.error("msaccess.exe n'est pas apparu. Arrêt.")
            _stop_event.set()
        return

    consecutive_empty = 0
    _EMPTY_THRESHOLD  = 2
    tracked_pids: set[int] = set()

    try:
        while not _stop_event.is_set():
            time.sleep(_SV_POLL_INTERVAL)
            
            current_pids = _get_msaccess_pids() - pids_before
            
            tracked_pids.update(current_pids)

            if not current_pids:
                consecutive_empty += 1
                log.debug(f"Studio Vision absent ({consecutive_empty}/{_EMPTY_THRESHOLD}).")
                if consecutive_empty >= _EMPTY_THRESHOLD:
                    log.info("Studio Vision fermé par l'utilisateur. Lancement de l'arrêt.")
                    break
            else:
                consecutive_empty = 0
                
    except Exception as exc:
        log.error(f"Erreur pendant la surveillance de msaccess.exe : {exc}")
    finally:
        for pid in tracked_pids:
            try:
                p = psutil.Process(pid)
                if p.is_running():
                    p.kill()
                    log.info(f"Processus zombie msaccess.exe (PID {pid}) tué de force pour libérer les verrous COM.")
            except Exception:
                pass

        # Attendre que tous les PIDs tués aient vraiment disparu du système
        # avant de rendre la main — évite que le garde de redémarrage les détecte encore.
        _KILL_DRAIN_TIMEOUT = 10  # secondes max
        _KILL_DRAIN_POLL    = 0.5
        deadline = time.monotonic() + _KILL_DRAIN_TIMEOUT
        while time.monotonic() < deadline:
            still_alive = {pid for pid in tracked_pids if psutil.pid_exists(pid)}
            if not still_alive:
                break
            log.debug(f"Attente de la disparition complète des PID : {still_alive}")
            time.sleep(_KILL_DRAIN_POLL)

        _stop_event.set()
        if _icon is not None:
            try:
                _icon.stop()
            except Exception:
                pass


# Point d'entrée
def main() -> None:
    global _icon, _mutex_handle

    # Garde contre les instances multiples
    _mutex_handle = win32event.CreateMutex(None, False, "ImageRouter_StudioVision_Mutex")
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        router_alive = any(
            (p.info["name"] or "").lower() in ("python.exe", "pythonw.exe")
            and p.info["pid"] != os.getpid()
            for p in psutil.process_iter(["pid", "name"])
        )
        if router_alive:
            log.warning("Une autre instance tourne déjà. Sortie.")
            sys.exit(0)
        else:
            log.warning("Mutex périmé détecté (plantage précédent). Poursuite.")

    # Empêche un redémarrage manuel pendant que Studio Vision tourne
    try:
        parent_name = psutil.Process(os.getpid()).parent().name().lower()
    except Exception:
        parent_name = ""

    if parent_name == "explorer.exe":
        _SV_CLOSE_WAIT = 8   # secondes max pour que msaccess.exe disparaisse
        _SV_CLOSE_POLL = 0.5
        waited = 0.0
        while waited < _SV_CLOSE_WAIT:
            sv_running = any(
                (p.info["name"] or "").lower() == "msaccess.exe"
                for p in psutil.process_iter(["name"])
            )
            if not sv_running:
                break
            time.sleep(_SV_CLOSE_POLL)
            waited += _SV_CLOSE_POLL
        else:
            # msaccess.exe encore présent après le délai → bloquer
            ctypes.windll.user32.MessageBoxW(
                0,
                "Pour relancer le routeur d'images, veuillez fermer "
                "complètement puis relancer Studio Vision.",
                "Routeur d'images",
                0x30,
            )
            sys.exit(0)

    prevent_sleep()

    if not SOURCE_DIR.exists():
        log.critical(f"Dossier source introuvable : {SOURCE_DIR}")
        sys.exit(1)

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("=== Routeur d'images Studio Vision ===")
    log.info(f"  Dossier source : {SOURCE_DIR}")
    log.info(f"  Dossier photos : {DEST_PHOTOS}")
    log.info(f"  Dossier orphel.: {ORPHAN_DIR}")
    log.info(f"  Fichier log    : {_LOG_FILE}")
    log.info(f"  Délai patient  : {PATIENT_WAIT_TIMEOUT // 60} min")
    log.info(f"  Extensions     : {', '.join(sorted(WATCHED_EXTENSIONS))}")
    log.info(f"  Rattrapage     : toutes les {CATCHUP_INTERVAL}s")
    log.info(f"  Sous-form SFDoc: {SFDOC_SUBFORM_NAME}")
    log.info(f"  Journal médecin: {_get_doctor_log_path()}")

    log_doctor(f"Routeur d'images démarré — surveillance : {SOURCE_DIR}")

    file_queue: queue.Queue = queue.Queue()

    threading.Thread(target=worker,          args=(file_queue,), name="Worker",               daemon=True).start()
    threading.Thread(target=_catchup_loop,   args=(file_queue,), name="Catchup",              daemon=True).start()
    threading.Thread(target=_run_background, args=(file_queue,), name="Background",           daemon=True).start()

    log.info("Analyse de démarrage — recherche de fichiers en attente...")
    _scan_source_for_missed_files(file_queue)

    sv_thread = threading.Thread(target=_launch_studio_vision, name="StudioVisionLauncher", daemon=True)
    sv_thread.start()

    if not TRAY_AVAILABLE:
        log.warning("pystray/Pillow indisponible — exécution sans icône de barre des tâches.")
        try:
            while not _stop_event.is_set():
                time.sleep(1)
        except KeyboardInterrupt:
            log.info("Arrêt demandé par interruption clavier.")
        finally:
            _stop_event.set()
        log.info("Application arrêtée.")
        log_doctor("Routeur d'images arrêté.")
        return

    menu = pystray.Menu(
        pystray.MenuItem(
            text=lambda item: _status_text,
            action=None,
            enabled=False,
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Ouvrir le log technique",  _open_log_file),
        pystray.MenuItem("Ouvrir le journal médecin",  _open_doctor_log),
        pystray.MenuItem("Quitter",                _quit),
    )

    _icon = pystray.Icon(
        name=BOX_NAME,
        icon=_make_icon(_COLOR_READY),
        title=BOX_NAME,
        menu=menu,
    )

    log.info("Icône de barre des tâches démarrée.")
    _icon.run()

    _stop_event.set()
    log.info("Application arrêtée.")
    log_doctor("Routeur d'images arrêté.")


if __name__ == "__main__":
    main()
