"""
Routeur d'images médicales (Windows 7 / Python 3.8.10)

Déplace les fichiers déposés par le système d'acquisition dans le bon dossier
patient sur le partage réseau, puis insère une entrée dans le sous-formulaire
SFDoc du formulaire Access actif, via l'automatisation de l'interface win32com.

Chaîne de traitement :
  PollingObserver → file_queue → Worker → déplacement du fichier → insertion interface (win32com)

Un ensemble partagé (_enqueued_files) évite le double traitement entre le watchdog
et le fil de balayage périodique. Gère le relais de démarrage en 2 étapes de /runtime
et force l'arrêt des processus COM zombies à la sortie.

Dépendances : watchdog, pywin32, pythoncom, pystray, Pillow, psutil
"""

import ctypes
import logging
import logging.handlers
import os
import pythoncom
import queue
import shutil
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional, Set

import win32api
import win32event
import winerror
import psutil

from watchdog.events import FileSystemEventHandler
from watchdog.observers.polling import PollingObserver as Observer

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


# Configuration
BOX_NAME = "Windows 7"

SOURCE_DIR  = Path(r"??")  # Dossier de dépôt des acquisitions
ORPHAN_DIR  = Path(r"??")  # Destination des fichiers non rattachés
DEST_PHOTOS = Path(r"??")  # Racine de l'archive photo réseau

STUDIO_VISION_CMD = [
    r"C:\Studiov2000-W7\svprog\msaccess.exe",
    "/runtime",
    r"C:\Studiov2000-W7\svprog\Ophprog.mde",
    "/wrkgrp",
    r"C:\Studiov2000-W7\config\system.mdw",
    "/User",
    "/Pwd",
    "/X",
    "demarrage",
]

WATCHED_EXTENSIONS: Set[str] = {
    ".jpg", ".jpeg", ".jfif",
    ".png", ".bmp",
    ".tif", ".tiff",
    ".dcm",
    ".pdf", ".rtf", ".doc", ".docx", ".odt",
}

EXAM_DESCRIPTION: Dict[str, str] = {
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
SWEEP_INTERVAL_SECONDS: int = 300  # 5 minutes

ACCESS_FIELD_CODE   = "Code patient"
ACCESS_FIELD_NOM    = "NOM"
ACCESS_FIELD_PRENOM = "Prénom"
SFDOC_SUBFORM_NAME  = "SFDoc"
_AC_SUBFORM         = 112  # Constante Access ControlType pour le sous-formulaire

_GUI_PRE_INSERT_DELAY = 0.3  # Secondes entre le déplacement du fichier et l'insertion interface
_UI_POST_INSERT_DELAY = 0.5  # Secondes entre l'insertion et Requery/MoveLast


# Journalisation — deux gestionnaires à rotation temporelle :
#   image_router.log       : log technique complet (rotation 30 jours)
#   transferts_medecin.log : résumé en français clair pour les utilisateurs (90 jours)
_LOG_DIR = os.path.join(os.path.expanduser("~"), "studiovision")
os.makedirs(_LOG_DIR, exist_ok=True)

_LOG_FILE_TECH    = os.path.join(_LOG_DIR, "image_router.log")
_LOG_FILE_MEDECIN = os.path.join(_LOG_DIR, "transferts_medecin.log")

_tech_handler = logging.handlers.TimedRotatingFileHandler(
    _LOG_FILE_TECH, when="midnight", interval=1, backupCount=30, encoding="utf-8",
)
_tech_handler.setFormatter(
    logging.Formatter(
        "%(asctime)s  %(levelname)-8s  [%(threadName)s]  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
)

_console_handler = logging.StreamHandler(sys.stdout)
_console_handler.setFormatter(
    logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s", datefmt="%H:%M:%S")
)

logging.basicConfig(level=logging.INFO, handlers=[_tech_handler, _console_handler])
log = logging.getLogger("image_router")

_medecin_handler = logging.handlers.TimedRotatingFileHandler(
    _LOG_FILE_MEDECIN, when="midnight", interval=1, backupCount=90, encoding="utf-8",
)
_medecin_handler.setFormatter(logging.Formatter("%(message)s"))

_medecin_log = logging.getLogger("medecin")
_medecin_log.setLevel(logging.INFO)
_medecin_log.propagate = False  # Tenu séparé du log technique
_medecin_log.addHandler(_medecin_handler)


def _log_medecin(msg: str) -> None:
    """Écrit une ligne horodatée, en français clair, dans le journal destiné au médecin."""
    _medecin_log.info("%s - %s", datetime.now().strftime("%H:%M"), msg)


# État global
_ICON_SIZE    = 64
_COLOR_READY  = (30, 144, 255)
_COLOR_ACTIVE = (50, 205, 50)

_icon:         Optional["pystray.Icon"] = None
_status_text:  str                      = "Démarrage..."
_stop_event:   threading.Event          = threading.Event()
_mutex_handle                           = None


# Utilitaires de la barre des tâches (system tray)
def _make_icon(color: tuple) -> "Image.Image":
    img  = Image.new("RGBA", (_ICON_SIZE, _ICON_SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    m    = 4
    draw.ellipse([m, m, _ICON_SIZE - m, _ICON_SIZE - m], fill=color)
    return img


def _set_status(text: str, processing: bool = False) -> None:
    global _status_text
    _status_text = text
    if _icon is not None:
        try:
            _icon.icon = _make_icon(_COLOR_ACTIVE if processing else _COLOR_READY)
            _icon.update_menu()
        except Exception as exc:
            log.debug("Échec de mise à jour de l'icône : %s", exc)


def _notify(title: str, message: str = "") -> None:
    if _icon is not None:
        try:
            _icon.notify(message if message else title, title)
        except Exception as exc:
            log.debug("Échec de la notification : %s", exc)


def _open_tech_log(icon: object, item: object) -> None:
    try:
        os.startfile(_LOG_FILE_TECH)
    except Exception as exc:
        log.warning("Impossible d'ouvrir le log technique : %s", exc)


def _open_medecin_log(icon: object, item: object) -> None:
    try:
        os.startfile(_LOG_FILE_MEDECIN)
    except Exception as exc:
        log.warning("Impossible d'ouvrir le journal médecin : %s", exc)


def _quit(icon: object, item: object) -> None:
    log.info("Fermeture demandée depuis le menu de la barre des tâches.")
    _stop_event.set()
    icon.stop()


# Partage réseau
def wait_for_network_share() -> None:
    """Bloque jusqu'à ce que SOURCE_DIR soit accessible."""
    source_str = str(SOURCE_DIR)
    is_local   = not (source_str.startswith("\\\\") or source_str.startswith("//")) \
                 and len(source_str) >= 2 and source_str[1] == ":"
    if is_local:
        return
    first = True
    while True:
        try:
            if SOURCE_DIR.is_dir():
                if not first:
                    log.info("Partage réseau de nouveau accessible : %s", SOURCE_DIR)
                return
        except Exception:
            pass
        log.warning("Partage réseau inaccessible, nouvel essai dans 10s : %s", SOURCE_DIR)
        first = False
        time.sleep(10)


# Résolution du dossier patient
def build_patient_relative_path(patient_code: str, last_name: str, first_name: str) -> str:
    """
    Renvoie le chemin de dossier relatif selon la convention de nommage de Studio Vision.
    Format : <2premiers>.000\\<code><3derniers>.<3premiers>
    Exemple : code=1758511228, ABCDEF, DEFGH → "17.000\\1758511228abc.def"
    """
    prefix  = patient_code[:2]
    last_3  = last_name[:3].lower()
    first_3 = first_name[:3].lower()
    return "{0}.000\\{1}{2}.{3}".format(prefix, patient_code, last_3, first_3)


def resolve_patient_folder(patient: dict) -> Optional[Path]:
    """Résout et crée le dossier patient absolu sur le disque. Renvoie None en cas d'échec."""
    try:
        rel    = build_patient_relative_path(patient["code"], patient["nom"], patient["prenom"])
        folder = DEST_PHOTOS / rel
        folder.mkdir(parents=True, exist_ok=True)
        log.info("Dossier patient résolu : %s", folder)
        return folder
    except Exception as exc:
        log.error("Impossible de résoudre/créer le dossier patient : %s", exc)
        return None


# Access COM — lecture du patient actif
def get_active_patient() -> Optional[dict]:
    """Renvoie le code, le nom et le prénom du patient actif depuis le formulaire Access."""
    if not WIN32_AVAILABLE:
        return None
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            return None

        target: Set[str] = {ACCESS_FIELD_CODE, ACCESS_FIELD_NOM, ACCESS_FIELD_PRENOM}
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
        log.debug("Erreur COM dans get_active_patient : %s", exc)
        return None


# Access COM — recherche du sous-formulaire SFDoc
def _find_sfdoc(form: object) -> Optional[object]:
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
    Insère une nouvelle entrée dans le sous-formulaire SFDoc via l'automatisation win32com.
    Appelle AddNew(), remplit les champs, appelle Update(), puis Requery() + MoveLast().
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

        # Annuler si l'utilisateur a changé de patient
        current = get_active_patient()
        if not current or current["code"] != patient["code"]:
            log.warning(
                "Insertion interface annulée : le patient a changé (attendu=%s, courant=%s).",
                patient["code"], current["code"] if current else "none",
            )
            return False

        sfdoc = _find_sfdoc(form)
        if sfdoc is None:
            log.error("Sous-formulaire '%s' introuvable — insertion interface annulée.", SFDOC_SUBFORM_NAME)
            return False

        rs = sfdoc.Recordset
        rs.AddNew()

        def _set_field(name: str, value: object) -> None:
            try:
                rs.Fields(name).Value = value
            except Exception as exc:
                log.warning("Échec d'écriture du champ '%s' : %s", name, exc)

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
            "Insertion interface OK : patient=%s desc='%s' chemin='%s'",
            patient["code"], description, relative_path,
        )

        time.sleep(_UI_POST_INSERT_DELAY)

        try:
            sfdoc.Requery()
            log.info("Requery() sur '%s' OK.", SFDOC_SUBFORM_NAME)
        except Exception as exc:
            log.warning("Échec de Requery() : %s — tentative de Refresh().", exc)
            try:
                sfdoc.Refresh()
                log.info("Refresh() de repli sur '%s' OK.", SFDOC_SUBFORM_NAME)
            except Exception as exc2:
                log.warning("Le Refresh() de repli a aussi échoué : %s", exc2)

        try:
            sfdoc.Recordset.MoveLast()
        except Exception as exc:
            log.debug("Échec de MoveLast() : %s", exc)

        return True

    except Exception as exc:
        log.error("Échec de l'insertion interface : %s", exc)
        return False


# Utilitaires fichiers
def wait_for_file(file: Path) -> bool:
    """Bloque jusqu'à ce que le fichier soit lisible. Renvoie False si le nombre maximal de tentatives est dépassé."""
    for attempt in range(1, FILE_LOCK_MAX_ATTEMPTS + 1):
        try:
            with file.open("rb"):
                return True
        except (PermissionError, OSError):
            log.debug("Fichier verrouillé (%d/%d), nouvel essai...", attempt, FILE_LOCK_MAX_ATTEMPTS)
            time.sleep(FILE_LOCK_RETRY_DELAY)
    log.error("Fichier toujours verrouillé après %d tentatives : %s", FILE_LOCK_MAX_ATTEMPTS, file)
    return False


def move_file(source: Path, dest_folder: Path, label: str = "") -> Optional[Path]:
    """Déplace source vers dest_folder, en résolvant les conflits de nom par un suffixe horodaté."""
    dest_folder.mkdir(parents=True, exist_ok=True)
    dest = dest_folder / source.name
    if dest.exists():
        ts   = int(time.time())
        dest = dest_folder / "{0}_{1}{2}".format(source.stem, ts, source.suffix)
        log.info("Conflit de nom — renommé en %s", dest.name)
    try:
        shutil.move(str(source), str(dest))
        tag = "[{0}]  ".format(label) if label else ""
        log.info("%s%s -> %s", tag, source.name, dest)
        return dest
    except Exception as exc:
        log.error("Échec du déplacement : %s", exc)
        return None


def orphan_file(file: Path) -> None:
    """Déplace un fichier non traitable vers le dossier des orphelins."""
    log.warning("Mise en orphelin : %s", file.name)
    move_file(file, ORPHAN_DIR, label="ORPHAN")
    _log_medecin(
        "Fichier non attribué (aucun patient ouvert) : {0} — déplacé dans le dossier orphelins.".format(
            file.name
        )
    )


def prevent_sleep() -> None:
    """Empêche Windows de se mettre en veille pendant que le routeur est actif."""
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            0x80000000 |  # ES_CONTINUOUS
            0x00000001    # ES_SYSTEM_REQUIRED
        )
        log.info("Prévention de la veille active.")
    except Exception as exc:
        log.warning("Impossible de définir l'état d'exécution : %s", exc)


# Fil de traitement (Worker)
def worker(file_queue: queue.Queue) -> None:
    """
    Consomme les fichiers de la file.
    Pour chaque fichier : déverrouillage → attente d'un patient ouvert → déplacement → insertion interface.
    """
    pythoncom.CoInitialize()
    log.info("Fil de traitement démarré.")

    needs_refresh:     bool          = False
    last_patient_code: Optional[str] = None
    burst_count:       int           = 0

    try:
        while True:
            try:
                file: Path = file_queue.get(timeout=1.5)
            except queue.Empty:
                if needs_refresh:
                    log.info("Rafale terminée — tous les fichiers du lot ont été traités.")
                    _notify("Transfert terminé", "{0} fichier(s) traité(s)".format(burst_count))
                    _set_status("{0} — Prêt".format(BOX_NAME), processing=False)
                    needs_refresh     = False
                    last_patient_code = None
                    burst_count       = 0
                if _stop_event.is_set():
                    break
                continue

            if file is None:
                break

            log.info("Traitement : %s (%d en attente)", file.name, file_queue.qsize())

            if burst_count == 0 and not needs_refresh:
                _notify("Transfert en cours", file.name)
            _set_status("Transfert en cours...", processing=True)

            if not file.exists():
                log.warning("Fichier disparu avant traitement : %s", file)
                file_queue.task_done()
                continue

            if not wait_for_file(file):
                log.error("Abandon — verrou persistant : %s", file.name)
                _notify("Erreur", "Fichier verrouillé : {0}".format(file.name))
                _log_medecin(
                    "Erreur : le fichier {0} est verrouillé et n'a pas pu être transféré.".format(file.name)
                )
                file_queue.task_done()
                continue

            # Attendre qu'un patient soit ouvert dans Access
            patient: Optional[dict] = None
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
                    file_queue.task_done()
                    patient = None
                    break

                if first_log:
                    log.info("Aucun patient ouvert — attente (délai dans %d min).", PATIENT_WAIT_TIMEOUT // 60)
                    first_log = False

                time.sleep(PATIENT_POLL_INTERVAL)

            if patient is None:
                continue

            log.info("Patient : %s %s (code %s)", patient["nom"], patient["prenom"], patient["code"])

            patient_folder = resolve_patient_folder(patient)
            if not patient_folder:
                log.error("Impossible de résoudre le dossier du patient %s — mise en orphelin.", patient["code"])
                orphan_file(file)
                _notify("Fichier orphelin", file.name)
                _log_medecin(
                    "Erreur : impossible de créer le dossier pour {0} {1} "
                    "(code {2}) — fichier orphelin.".format(
                        patient["nom"].upper(), patient["prenom"], patient["code"]
                    )
                )
                file_queue.task_done()
                continue

            dest = move_file(file, patient_folder)
            if dest is None:
                file_queue.task_done()
                continue

            rel_path      = build_patient_relative_path(patient["code"], patient["nom"], patient["prenom"])
            relative_path = "\\{0}\\{1}".format(rel_path, dest.name)
            description   = EXAM_DESCRIPTION.get(file.suffix.lower(), "Image")

            time.sleep(_GUI_PRE_INSERT_DELAY)

            if gui_insert_document(patient, relative_path, description):
                needs_refresh     = True
                last_patient_code = patient["code"]
                burst_count      += 1
                log.info("Entrée insérée : '%s' -> %s", dest.name, relative_path)
                _log_medecin(
                    "Image transférée avec succès pour le patient {0} {1} ({2}).".format(
                        patient["nom"].upper(), patient["prenom"], description,
                    )
                )
            else:
                log.error(
                    "Échec de l'insertion interface pour '%s' (patient %s). Fichier situé : %s. Saisie manuelle nécessaire.",
                    dest.name, patient["code"], dest,
                )
                _notify("Erreur insertion", "'{0}' déplacé mais non inséré — voir logs.".format(dest.name))
                _log_medecin(
                    "Erreur : l'image {0} a été déplacée vers {1} mais n'a PAS pu "
                    "être insérée pour le patient {2} {3} (code {4}). "
                    "Saisie manuelle requise.".format(
                        dest.name, dest,
                        patient["nom"].upper(), patient["prenom"], patient["code"],
                    )
                )

            file_queue.task_done()

    finally:
        _set_status("{0} — Arrêté".format(BOX_NAME))
        pythoncom.CoUninitialize()
        log.info("Fil de traitement arrêté.")


# Balayage de rattrapage — registre partagé pour éviter le double traitement
_enqueued_files: Set[Path] = set()
_enqueued_lock  = threading.Lock()


def _sweep_source_dir(file_queue: queue.Queue) -> None:
    """Analyse SOURCE_DIR et met en file tout fichier valide non encore suivi."""
    try:
        found = list(SOURCE_DIR.rglob("*"))
    except Exception as exc:
        log.warning("Balayage : impossible de lister SOURCE_DIR : %s", exc)
        return

    for path in found:
        if not path.is_file() or path.suffix.lower() not in WATCHED_EXTENSIONS:
            continue
        with _enqueued_lock:
            if path in _enqueued_files:
                continue
            _enqueued_files.add(path)

        log.info("Balayage : remise en file du fichier manqué : %s", path.name)
        _log_medecin(
            "Fichier détecté lors du balayage périodique "
            "(non capturé en temps réel) : {0}.".format(path.name)
        )
        file_queue.put(path)


def _run_sweep(file_queue: queue.Queue) -> None:
    """Fil d'arrière-plan : balayage de rattrapage périodique."""
    log.info("Fil de balayage démarré — intervalle : %d s.", SWEEP_INTERVAL_SECONDS)
    while not _stop_event.wait(timeout=SWEEP_INTERVAL_SECONDS):
        log.debug("Balayage : recherche de fichiers manqués dans SOURCE_DIR...")
        _sweep_source_dir(file_queue)
    log.info("Fil de balayage arrêté.")


# Producteur Watchdog
class ImageProducer(FileSystemEventHandler):
    """Met en file les fichiers nouvellement créés détectés par l'observateur.
    Les enregistre aussi dans _enqueued_files pour éviter une remise en file par le balayage."""

    def __init__(self, file_queue: queue.Queue) -> None:
        super().__init__()
        self._queue = file_queue

    def on_created(self, event: object) -> None:
        if event.is_directory:
            return
        file = Path(event.src_path)
        if file.suffix.lower() not in WATCHED_EXTENSIONS:
            return
        with _enqueued_lock:
            if file in _enqueued_files:
                return
            _enqueued_files.add(file)
        log.info("Mis en file (watchdog) : %s (file : %d)", file.name, self._queue.qsize() + 1)
        self._queue.put(file)


# Fil d'observation en arrière-plan
def _run_background(file_queue: queue.Queue) -> None:
    """Démarre et surveille l'observateur du système de fichiers. Se reconnecte en cas de coupure réseau."""
    _RECONNECT_DELAY = 15

    def _start_observer() -> Observer:
        obs = Observer()
        obs.schedule(ImageProducer(file_queue), str(SOURCE_DIR), recursive=True)
        obs.start()
        log.info("Observateur démarré — surveillance : %s", SOURCE_DIR)
        return obs

    observer = _start_observer()
    _set_status("{0} — Prêt".format(BOX_NAME), processing=False)

    try:
        while not _stop_event.is_set():
            time.sleep(1)
            if not observer.is_alive():
                log.warning(
                    "Observateur arrêté (coupure réseau possible) — attente %ds avant reconnexion.",
                    _RECONNECT_DELAY,
                )
                _set_status("{0} — Reconnexion...".format(BOX_NAME))
                try:
                    observer.stop()
                    observer.join(timeout=5)
                except Exception:
                    pass
                wait_for_network_share()
                time.sleep(_RECONNECT_DELAY)
                observer = _start_observer()
                _set_status("{0} — Prêt".format(BOX_NAME), processing=False)
    finally:
        observer.stop()
        observer.join()
        remaining = file_queue.qsize()
        if remaining:
            log.info("Attente des %d fichier(s) restant(s)...", remaining)
            file_queue.join()
        log.info("Fil d'arrière-plan arrêté.")
        if _icon is not None:
            _icon.stop()


# Cycle de vie de Studio Vision
_SV_POLL_INTERVAL   = 3   # Secondes entre deux vérifications de présence de msaccess.exe
_SV_STARTUP_TIMEOUT = 30  # Secondes d'attente de l'apparition de msaccess.exe après lancement


def _get_msaccess_pids() -> Set[int]:
    """Renvoie l'ensemble des PID des processus msaccess.exe en cours d'exécution."""
    pids: Set[int] = set()
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            if (proc.info["name"] or "").lower() == "msaccess.exe":
                pids.add(proc.info["pid"])
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return pids


def _launch_studio_vision() -> None:
    """
    Lance Studio Vision et surveille tous les processus msaccess.exe.
    Gère le relais en 2 étapes de /runtime : le processus lanceur se termine ~2 s
    après le démarrage et engendre le vrai processus de travail sous un nouveau PID.
    On suit l'ensemble des nouveaux PID plutôt qu'un seul.
    Force l'arrêt de tout processus suivi à la sortie pour libérer les verrous COM.
    """
    log.info("Lancement de Studio Vision : %s", " ".join(STUDIO_VISION_CMD))

    pids_before: Set[int] = _get_msaccess_pids()

    try:
        subprocess.Popen(STUDIO_VISION_CMD)
    except FileNotFoundError:
        log.critical("Exécutable Studio Vision introuvable. Arrêt.")
        _stop_event.set()
        return
    except Exception as exc:
        log.error("Impossible de lancer Studio Vision : %s. Arrêt.", exc)
        _stop_event.set()
        return

    log.info("Attente de msaccess.exe (max %ds)...", _SV_STARTUP_TIMEOUT)
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
    tracked_pids: Set[int] = set()

    try:
        while not _stop_event.is_set():
            time.sleep(_SV_POLL_INTERVAL)
            current_pids = _get_msaccess_pids() - pids_before
            tracked_pids.update(current_pids)

            if not current_pids:
                consecutive_empty += 1
                log.debug("Studio Vision absent (%d/%d).", consecutive_empty, _EMPTY_THRESHOLD)
                if consecutive_empty >= _EMPTY_THRESHOLD:
                    log.info("Studio Vision fermé par l'utilisateur. Lancement de l'arrêt.")
                    break
            else:
                consecutive_empty = 0

    except Exception as exc:
        log.error("Erreur pendant la surveillance de msaccess.exe : %s", exc)
    finally:
        for pid in tracked_pids:
            try:
                p = psutil.Process(pid)
                if p.is_running():
                    p.kill()
                    log.info("Processus zombie msaccess.exe (PID %d) tué de force pour libérer les verrous COM.", pid)
            except Exception:
                pass

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
    _mutex_handle = win32event.CreateMutex(None, False, "ImageRouter_Windows7_Mutex")
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
        sv_running = any(
            (p.info["name"] or "").lower() == "msaccess.exe"
            for p in psutil.process_iter(["name"])
        )
        if sv_running:
            ctypes.windll.user32.MessageBoxW(
                0,
                "Pour relancer le routeur d'images, veuillez fermer "
                "complètement puis relancer Studio Vision.",
                "Routeur d'images",
                0x30,  # MB_ICONWARNING | MB_OK
            )
            sys.exit(0)

    prevent_sleep()

    log.info("Vérification de la disponibilité du partage réseau...")
    wait_for_network_share()

    if not SOURCE_DIR.exists():
        log.critical("Dossier source introuvable : %s", SOURCE_DIR)
        sys.exit(1)

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("=== Routeur d'images Studio Vision (Windows 7) ===")
    log.info("  Dossier source : %s", SOURCE_DIR)
    log.info("  Dossier photos : %s", DEST_PHOTOS)
    log.info("  Dossier orphel.: %s", ORPHAN_DIR)
    log.info("  Log technique  : %s", _LOG_FILE_TECH)
    log.info("  Journal médecin: %s", _LOG_FILE_MEDECIN)
    log.info("  Délai patient  : %d min", PATIENT_WAIT_TIMEOUT // 60)
    log.info("  Extensions     : %s", ", ".join(sorted(WATCHED_EXTENSIONS)))
    log.info("  Balayage       : toutes les %d s", SWEEP_INTERVAL_SECONDS)
    log.info("  Sous-form SFDoc: %s", SFDOC_SUBFORM_NAME)

    _log_medecin("Routeur d'images démarré (version 6) — surveillance active.")

    file_queue: queue.Queue = queue.Queue()

    log.info("Analyse de démarrage — recherche de fichiers en attente...")
    _sweep_source_dir(file_queue)

    threading.Thread(target=worker,          args=(file_queue,), name="Worker",     daemon=True).start()
    threading.Thread(target=_run_background, args=(file_queue,), name="Background", daemon=True).start()
    threading.Thread(target=_run_sweep,      args=(file_queue,), name="Sweep",      daemon=True).start()

    sv_thread = threading.Thread(target=_launch_studio_vision, name="StudioVisionLauncher", daemon=True)
    sv_thread.start()

    if not TRAY_AVAILABLE:
        log.warning("pystray/Pillow indisponible — exécution sans icône de barre des tâches.")
        try:
            while not _stop_event.is_set():
                time.sleep(1)
        except KeyboardInterrupt:
            log.info("Arrêt demandé par le clavier.")
        finally:
            _stop_event.set()
        _log_medecin("Routeur d'images arrêté.")
        log.info("Application arrêtée.")
        return

    menu = pystray.Menu(
        pystray.MenuItem(
            text=lambda item: _status_text,
            action=None,
            enabled=False,
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Ouvrir le log technique", _open_tech_log),
        pystray.MenuItem("Ouvrir le log médecin",   _open_medecin_log),
        pystray.MenuItem("Quitter",                 _quit),
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
    _log_medecin("Routeur d'images arrêté.")
    log.info("Application arrêtée.")


if __name__ == "__main__":
    main()