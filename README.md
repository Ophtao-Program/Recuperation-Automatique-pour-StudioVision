# Récupération Automatique pour StudioVision

Routage automatique des images et documents médicaux vers le bon dossier patient de **StudioVision**.

Lorsqu'un appareil d'imagerie dépose un fichier, le programme le détecte, identifie le patient ouvert dans StudioVision, déplace le fichier dans le dossier patient correspondant sur le lecteur réseau, puis insère directement une entrée dans le formulaire Access (sous-formulaire `SFDoc`) via l'automatisation COM — de sorte que l'image apparaît immédiatement dans le dossier du patient, sans aucune saisie manuelle.

---

## Aperçu du fonctionnement

1. Un **observateur** (`PollingObserver`) surveille le dossier d'acquisition (`SOURCE_DIR`).
2. Chaque fichier détecté est placé dans une file et traité par un fil de travail en arrière-plan.
3. Le fil attend que le fichier soit entièrement écrit (vérification du verrou, avec réessais).
4. Il interroge le formulaire Access actif de StudioVision via COM pour récupérer le patient courant (code, nom, prénom).
5. Il résout le dossier du patient sur le réseau selon la convention de nommage de Studio Vision.
6. Il y déplace le fichier (en ajoutant un suffixe horodaté en cas de conflit de nom).
7. Il insère une nouvelle entrée dans le sous-formulaire `SFDoc` (`AddNew → Update → Requery`).
8. Si aucun patient n'est ouvert avant l'expiration du délai, le fichier est déplacé vers le dossier des orphelins.

---

## Les scripts (`src/`)

| Fichier | Rôle |
|---|---|
| `routeur_images.py` | Routeur principal des images et documents. Insertion dans Access via `win32com`. À utiliser sur les postes récents (Python 3.10+). |
| `routeur_images_windows7.py` | Même routeur, variante compatible **Windows 7 / Python 3.8.10**. Gère le relais de démarrage `/runtime` en deux étapes et un balayage de rattrapage partagé. |
| `pont_refractometre.py` | Pont série pour le **réfractomètre**. Lit les trames sur le port série (COM6 par défaut) et injecte les valeurs dans le formulaire `REFRACTION` ouvert dans Access. |

> Les anciennes versions (V1 à V5) ont été retirées : seule la version de production est conservée.

---

## Prérequis

- **Windows uniquement** — automatisation COM via `pywin32`.
- Python **3.10+** (`routeur_images.py`, `pont_refractometre.py`) ou **3.8.10+** (`routeur_images_windows7.py`).
- `pystray` et `Pillow` pour l'icône de la barre des tâches (bascule en mode sans interface s'ils sont absents).
- `psutil` pour le suivi du cycle de vie des processus.
- `pyserial` uniquement pour `pont_refractometre.py`.

### Installation

```powershell
pip install -r requirements.txt
python -m pywin32_postinstall -install
```

---

## Configuration

Renseignez les constantes en haut du script que vous lancez :

| Variable | Description |
|---|---|
| `SOURCE_DIR` | Dossier surveillé pour les nouveaux fichiers (dossier de dépôt de l'appareil d'imagerie) |
| `ORPHAN_DIR` | Destination des fichiers qui n'ont pas pu être rattachés à un patient |
| `DEST_PHOTOS` | Racine des dossiers photos des patients sur le lecteur réseau |
| `STUDIO_VISION_CMD` | Commande de lancement de StudioVision (`msaccess.exe /runtime ...`) |

Autres constantes ajustables :

| Constante | Valeur par défaut | Description |
|---|---|---|
| `FILE_LOCK_RETRY_DELAY` | `3` s | Délai entre deux essais quand un fichier est encore verrouillé |
| `FILE_LOCK_MAX_ATTEMPTS` | `15` | Nombre maximal d'essais avant abandon sur un fichier verrouillé |
| `PATIENT_POLL_INTERVAL` | `3` s | Fréquence d'interrogation d'Access pour détecter un patient ouvert |
| `PATIENT_WAIT_TIMEOUT` | `900` s | Délai avant mise en orphelin si aucun patient n'est trouvé (15 min) |
| `CATCHUP_INTERVAL` / `SWEEP_INTERVAL_SECONDS` | `120` s / `300` s | Intervalle entre deux analyses de rattrapage du dossier source |
| `SFDOC_SUBFORM_NAME` | `"SFDoc"` | Nom du sous-formulaire Access listant les documents |

---

## Extensions surveillées

`.jpg`, `.jpeg`, `.jfif`, `.png`, `.bmp`, `.tif`, `.tiff`, `.dcm`, `.pdf`, `.rtf`, `.doc`, `.docx`, `.odt`

La description insérée dans StudioVision est déduite de l'extension : `Image`, `OCT` (`.tif`/`.tiff`), `DICOM` (`.dcm`) ou `Document`.

---

## Exécution

```powershell
# Poste récent (Python 3.10+)
pythonw "src/routeur_images.py"

# Poste Windows 7 (Python 3.8.10)
pythonw "src/routeur_images_windows7.py"

# Pont réfractomètre (formulaire REFRACTION)
pythonw "src/pont_refractometre.py"
```

L'arrêt se fait via l'entrée **Quitter** du menu de la barre des tâches, ou par `Ctrl+C` en mode sans interface.

---

## Convention de nommage des dossiers patients

Le dossier patient est déduit des champs d'identité lus dans le formulaire Access actif, selon la convention de Studio Vision :

```
DEST_PHOTOS\<2 premiers chiffres>.000\<code><3 premières lettres du nom>.<3 premières lettres du prénom>\
```

Exemple : code `0042`, nom `Dupont`, prénom `Marie` → `DEST_PHOTOS\00.000\0042dup.mar\`

---

## Journaux

Les journaux sont écrits dans `~/studiovision/`.

- **Journal technique** (`image_router.log`) — trace détaillée pour le diagnostic.
- **Journal médecin** — résumé en français clair, lisible par le praticien, indiquant pour chaque fichier s'il a bien été ajouté au dossier du bon patient (ou les éventuels problèmes à traiter).

Les deux sont accessibles directement depuis le menu de la barre des tâches.

---

## Fichiers orphelins

Un fichier est déplacé vers `ORPHAN_DIR` lorsque :

- aucun patient n'est ouvert dans StudioVision avant l'expiration du délai ;
- le dossier du patient n'a pas pu être résolu sur le disque.

Tous les événements « orphelin » sont consignés comme avertissements et doivent être traités manuellement.

---

## Notes techniques

- `pythoncom.CoInitialize()` / `CoUninitialize()` sont appelés sur le fil de travail : les objets COM ne peuvent pas être partagés entre fils.
- Lorsque plusieurs routeurs tournent en même temps sur la même machine, chacun se lie à sa propre instance `Access.Application` en parcourant la *Running Object Table* (ROT) plutôt qu'en utilisant `GetActiveObject`, qui renvoie toujours la première instance enregistrée.
- Le routeur lance StudioVision via `subprocess.Popen`, suit les nouveaux PID `msaccess.exe`, et les force à s'arrêter à la sortie pour libérer les verrous COM.
- En mode barre des tâches, `pystray` requiert le fil principal ; l'observateur, le fil de travail et le fil de cycle de vie tournent en tâches de fond (daemons).

---

## Documentation

- **Guide d'utilisation.pdf** — guide d'installation et d'utilisation pas à pas.
- **Rapport_Technique.pdf** — rapport technique détaillant l'architecture et les choix d'implémentation.
