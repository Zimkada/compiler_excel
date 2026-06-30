"""
Résolution des chemins de ressources, compatible exécution normale ET packagée.

En développement, les ressources (icônes) sont à la racine du projet et lues
via un chemin relatif au répertoire courant. Une fois l'application packagée
avec PyInstaller (mode one-folder ou one-file), ces ressources sont extraites
dans un dossier temporaire pointé par ``sys._MEIPASS`` : le répertoire courant
n'est alors plus celui de l'application, et ``Path("icon.ico")`` échoue.

``resource_path()`` résout ce problème en cherchant la ressource d'abord dans
le bundle PyInstaller (s'il existe), puis à la racine du projet (développement).
"""

import sys
from pathlib import Path


def resource_path(relative: str) -> Path:
    """Retourne le chemin absolu d'une ressource embarquée.

    Args:
        relative: Chemin de la ressource relatif à la racine du projet
            (ex. ``"icon.ico"``).

    Returns:
        Chemin absolu vers la ressource, qu'on tourne depuis les sources ou
        depuis un exécutable PyInstaller. Le fichier n'est pas garanti
        d'exister : l'appelant vérifie via ``.exists()`` comme avant.
    """
    base = getattr(sys, "_MEIPASS", None)
    if base:
        # Exécution packagée : ressources extraites par PyInstaller.
        return Path(base) / relative
    # Exécution depuis les sources : racine du projet (parent de utils/).
    return Path(__file__).resolve().parent.parent / relative
