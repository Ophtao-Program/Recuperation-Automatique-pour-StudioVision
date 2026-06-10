"""
Fronto — Pont série réfractomètre pour StudioVision
Lit les trames du réfractomètre sur COM6 et injecte les valeurs
dans le formulaire REFRACTION ouvert dans Access via win32com.
"""

from typing import Dict, Optional, Set
import win32com.client
import pythoncom
import serial
import threading
import subprocess
import sys
import re
import time
from pathlib import Path
import psutil
import logging

try:
    import pystray
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False

# Journalisation — écrit dans fronto.log et stdout
logging.basicConfig(
    level=logging.DEBUG,  # Utiliser INFO en production
    format="%(asctime)s [%(levelname)-8s] %(threadName)-15s: %(message)s",
    handlers=[
        logging.FileHandler("fronto.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("Fronto")

# Configuration
APP_NAME    = "Fronto"
SERIAL_PORT = "COM6"
BAUD_RATE   = 9600
BYTESIZE    = serial.EIGHTBITS
PARITY      = serial.PARITY_NONE
STOPBITS    = serial.STOPBITS_ONE

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

_SV_POLL_INTERVAL   = 3   # Secondes entre deux vérifications de présence de msaccess.exe
_SV_STARTUP_TIMEOUT = 30  # Secondes d'attente de l'apparition de msaccess.exe

# État global
_stop_event  = threading.Event()
_icon        = None
_status_text = "Démarrage..."

_ICON_SIZE    = 64
_COLOR_READY  = (30, 144, 255)
_COLOR_ACTIVE = (50, 205, 50)
_COLOR_ERROR  = (220, 50, 50)


# Utilitaires de la barre des tâches (system tray)
def _make_icon(color):
    img  = Image.new("RGBA", (_ICON_SIZE, _ICON_SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    margin = 4
    draw.ellipse([margin, margin, _ICON_SIZE - margin, _ICON_SIZE - margin], fill=color)
    return img


def _set_status(text, color=None):
    global _status_text
    _status_text = text
    if _icon is not None:
        try:
            if color is not None:
                _icon.icon = _make_icon(color)
            _icon.update_menu()
        except Exception as e:
            logger.debug("Échec de mise à jour de l'icône : %s", e)


def _notify(title, message=""):
    if _icon is not None:
        try:
            _icon.notify(message if message else title, title)
        except Exception as e:
            logger.debug("Échec de la notification : %s", e)


def _quit(icon, item):
    logger.info("Fermeture demandée depuis le menu de la barre des tâches.")
    _stop_event.set()
    icon.stop()


def parse_trame(line):
    # type: (str) -> Optional[Dict[str, str]]
    """
    Analyse une trame série du réfractomètre.
    Examples:
      [01]RSPH=-11.25;CYL=-02.50;AXS=028;AD1=+01.75;Phx=+31.27;[17]
      [02]LSPH=-12.50;CYL=-01.00;AXS=148;AD1=+03.50;Phx=+32.46;IDP=43288;[17]
    Renvoie un dict avec 'eye' ('OD' ou 'OG') et les valeurs analysées, ou None.
    """
    line = line.strip()
    logger.debug("Analyse de la trame : '%s'", line)

    # [01] = œil droit (OD), [02] = œil gauche (OG)
    eye_match = re.match(r'^\[0([12])\]', line)
    if not eye_match:
        logger.warning("Préfixe d'œil non reconnu — [01] ou [02] attendu. Ignoré.")
        return None

    eye = "OD" if eye_match.group(1) == "1" else "OG"
    result = {"eye": eye}
    logger.debug("Œil détecté : %s", eye)

    m = re.search(r'[RL]SPH=([+-]?\d+\.\d+)', line)
    if m:
        result["SPH"] = m.group(1)

    m = re.search(r'CYL=([+-]?\d+\.\d+)', line)
    if m:
        result["CYL"] = m.group(1)

    m = re.search(r'AXS=(\d+)', line)
    if m:
        result["AXS"] = str(int(m.group(1)))  # retire les zéros de tête

    m = re.search(r'AD1=([+-]?\d+\.\d+)', line)
    if m:
        result["ADD"] = m.group(1)

    if len(result) > 1:
        logger.info("Trame décodée : %s", result)
        return result
    else:
        logger.warning("Aucune valeur clinique (SPH, CYL, AXS, ADD) trouvée dans la trame.")
        return None


def inject_into_access(data):
    # type: (Dict[str, str]) -> None
    """Injecte les valeurs réfractomètre analysées dans le formulaire REFRACTION d'Access."""
    logger.info("Injection dans Access...")
    pythoncom.CoInitialize()
    try:
        access = win32com.client.GetActiveObject("Access.Application")
    except Exception as e:
        logger.error("Impossible de se connecter à StudioVision via COM : %s", e)
        _set_status("{} — Erreur Access".format(APP_NAME), _COLOR_ERROR)
        return

    try:
        form = access.Forms("REFRACTION")
    except Exception as e:
        logger.error("Formulaire 'REFRACTION' non ouvert dans Access : %s", e)
        _set_status("{} — Formulaire introuvable".format(APP_NAME), _COLOR_ERROR)
        return

    eye = data["eye"]
    mapping = {
        "SPH": "SPHERE {}".format(eye),
        "CYL": "CYLINDRE {}".format(eye),
        "AXS": "AXE {}".format(eye),
        "ADD": "ADD {}".format(eye),
    }

    success_count = 0
    for key, field_name in mapping.items():
        if key in data:
            try:
                form.Controls(field_name).Value = data[key]
                logger.info("  %s = %s", field_name, data[key])
                success_count += 1
            except Exception as e:
                logger.error("  Échec d'écriture du champ '%s' : %s", field_name, e)

    logger.info("Injection terminée (%d champ(s) écrit(s) pour %s).", success_count, eye)
    _set_status("{} — Données envoyées ({})".format(APP_NAME, eye), _COLOR_ACTIVE)
    _notify("Réfractomètre", "{} valeur(s) injectée(s) pour {}".format(success_count, eye))
    time.sleep(2)
    _set_status("{} — En attente".format(APP_NAME), _COLOR_READY)


def monitor_serial():
    """Lit en continu le port série et transmet les trames analysées à Access."""
    logger.info("Ouverture de %s à %d bps...", SERIAL_PORT, BAUD_RATE)
    while not _stop_event.is_set():
        try:
            with serial.Serial(
                port=SERIAL_PORT,
                baudrate=BAUD_RATE,
                bytesize=BYTESIZE,
                parity=PARITY,
                stopbits=STOPBITS,
                timeout=1,
            ) as ser:
                logger.info("%s ouvert — en attente de trames...", SERIAL_PORT)
                _set_status("{} — En attente".format(APP_NAME), _COLOR_READY)
                buffer = ""
                while not _stop_event.is_set():
                    raw = ser.read(256)
                    if raw:
                        buffer += raw.decode("ascii", errors="replace")
                        while "\n" in buffer:
                            line, buffer = buffer.split("\n", 1)
                            line = line.strip()
                            if not line:
                                continue
                            logger.info("Trame reçue : %s", line)
                            data = parse_trame(line)
                            if data:
                                # Injection dans un fil séparé pour ne pas bloquer la lecture série
                                t = threading.Thread(
                                    target=inject_into_access,
                                    args=(data,),
                                    name="Inject_{}".format(data["eye"]),
                                )
                                t.daemon = True
                                t.start()

        except serial.SerialException as e:
            logger.error("Erreur du port série sur %s : %s — nouvel essai dans 5s...", SERIAL_PORT, e)
            _set_status("{} — Erreur port série".format(APP_NAME), _COLOR_ERROR)
            time.sleep(5)
        except Exception as e:
            logger.critical("Erreur inattendue dans monitor_serial : %s", e)
            _set_status("{} — Erreur inattendue".format(APP_NAME), _COLOR_ERROR)
            time.sleep(5)

    logger.info("Fil série arrêté.")


def _get_msaccess_pids():
    """Renvoie l'ensemble des PID des processus msaccess.exe en cours d'exécution."""
    pids = set()
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            if (proc.info["name"] or "").lower() == "msaccess.exe":
                pids.add(proc.info["pid"])
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return pids


def _launch_studio_vision():
    """
    Lance StudioVision et surveille msaccess.exe.
    Déclenche l'arrêt à la fermeture de StudioVision.
    Force l'arrêt des processus zombies msaccess.exe à la sortie pour libérer les verrous COM.
    """
    logger.info("Lancement de StudioVision...")
    _set_status("{} — Lancement de StudioVision...".format(APP_NAME), _COLOR_READY)

    pids_before = _get_msaccess_pids()

    try:
        subprocess.Popen(STUDIO_VISION_CMD)
    except FileNotFoundError:
        logger.error("Exécutable introuvable : %s. Arrêt.", STUDIO_VISION_CMD[0])
        _set_status("{} — Exécutable introuvable".format(APP_NAME), _COLOR_ERROR)
        _stop_event.set()
        return
    except Exception as e:
        logger.error("Impossible de lancer StudioVision : %s", e)
        _set_status("{} — Erreur de lancement".format(APP_NAME), _COLOR_ERROR)
        _stop_event.set()
        return

    logger.info("Attente de msaccess.exe (max %ds)...", _SV_STARTUP_TIMEOUT)
    deadline = time.monotonic() + _SV_STARTUP_TIMEOUT

    while time.monotonic() < deadline and not _stop_event.is_set():
        if _get_msaccess_pids() - pids_before:
            logger.info("StudioVision est lancé.")
            _set_status("{} — StudioVision lancé".format(APP_NAME), _COLOR_READY)
            break
        time.sleep(1)
    else:
        if not _stop_event.is_set():
            logger.error("msaccess.exe n'est pas apparu. Arrêt.")
            _set_status("{} — SV n'a pas démarré".format(APP_NAME), _COLOR_ERROR)
            _stop_event.set()
        return

    consecutive_empty = 0
    _EMPTY_THRESHOLD  = 2
    tracked_pids = set()

    try:
        while not _stop_event.is_set():
            time.sleep(_SV_POLL_INTERVAL)
            current_pids = _get_msaccess_pids() - pids_before
            tracked_pids.update(current_pids)

            if not current_pids:
                consecutive_empty += 1
                logger.debug("StudioVision absent (%d/%d).", consecutive_empty, _EMPTY_THRESHOLD)
                if consecutive_empty >= _EMPTY_THRESHOLD:
                    logger.info("StudioVision fermé par l'utilisateur. Lancement de l'arrêt.")
                    break
            else:
                consecutive_empty = 0
    except Exception as e:
        logger.error("Erreur pendant la surveillance de msaccess.exe : %s", e)
    finally:
        for pid in tracked_pids:
            try:
                p = psutil.Process(pid)
                if p.is_running():
                    p.kill()
                    logger.warning("Processus zombie msaccess.exe (PID %d) tué de force.", pid)
            except Exception:
                pass

        _stop_event.set()
        if _icon is not None:
            try:
                _icon.stop()
            except Exception:
                pass
        logger.info("Fil de cycle de vie arrêté.")


def main():
    global _icon

    logger.info("Fronto démarre.")

    sv_thread = threading.Thread(target=_launch_studio_vision, name="SV_Launcher")
    sv_thread.daemon = True
    sv_thread.start()

    serial_thread = threading.Thread(target=monitor_serial, name="SerialMonitor")
    serial_thread.daemon = True
    serial_thread.start()

    if not TRAY_AVAILABLE:
        logger.warning("pystray/Pillow indisponible — exécution sans interface (headless).")
        try:
            while not _stop_event.is_set():
                time.sleep(1)
        except KeyboardInterrupt:
            logger.info("Interruption clavier.")
        finally:
            _stop_event.set()
        logger.info("Application arrêtée.")
        return

    menu = pystray.Menu(
        pystray.MenuItem(text=lambda item: _status_text, action=None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Quitter", _quit),
    )

    _icon = pystray.Icon(
        name=APP_NAME,
        icon=_make_icon(_COLOR_READY),
        title=APP_NAME,
        menu=menu,
    )

    logger.info("Icône de barre des tâches démarrée.")
    _icon.run()

    _stop_event.set()
    sv_thread.join(timeout=15)
    serial_thread.join(timeout=5)
    logger.info("Application arrêtée.")


if __name__ == "__main__":
    main()