"""
Worker de vérification de mise à jour (appel réseau non-bloquant).

Effectue l'appel HTTP à l'API GitHub dans un thread séparé pour ne jamais
figer l'interface, et n'émet un signal QUE si une mise à jour est réellement
disponible. Tout échec (hors-ligne, timeout, API indisponible) est silencieux :
la vérification de mise à jour ne doit jamais gêner l'utilisateur.
"""

import urllib.request
import urllib.error

from PyQt6.QtCore import QThread, pyqtSignal

from utils.update_checker import (
    LATEST_RELEASE_URL, evaluate_release, UpdateInfo,
)
from utils import logger


class UpdateCheckWorker(QThread):
    """Vérifie en arrière-plan si une version plus récente existe.

    Émet ``update_found`` (avec un UpdateInfo) uniquement quand une mise à jour
    est disponible. En cas d'absence de mise à jour ou d'échec réseau, aucun
    signal n'est émis (silencieux).
    """

    update_found = pyqtSignal(object)  # UpdateInfo

    def __init__(self, current_version: str, timeout: float = 5.0, parent=None):
        super().__init__(parent)
        self.current_version = current_version
        self.timeout = timeout

    def run(self):
        try:
            req = urllib.request.Request(
                LATEST_RELEASE_URL,
                headers={
                    "Accept": "application/vnd.github+json",
                    "User-Agent": "ExcelCompiler-UpdateCheck",
                },
            )
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                body = resp.read().decode("utf-8", errors="replace")
        except (urllib.error.URLError, OSError, ValueError) as e:
            # Hors-ligne, timeout, DNS, etc. : silencieux, jamais bloquant.
            logger.info(f"Vérification de mise à jour ignorée: {e}")
            return
        except Exception as e:  # garde-fou ultime, jamais fatal
            logger.info(f"Vérification de mise à jour: erreur inattendue ignorée: {e}")
            return

        info: UpdateInfo = evaluate_release(body, self.current_version)
        if info.update_available:
            logger.info(
                f"Mise à jour disponible: {info.latest_version} "
                f"(actuelle {info.current_version})"
            )
            self.update_found.emit(info)
