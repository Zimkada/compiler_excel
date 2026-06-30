"""
Point d'entrée de l'application ExcelCompiler v3.2
Auteur: GOUNOU N'GOBI Chabi Zimé
"""

import sys
import logging
import warnings
from pathlib import Path

# openpyxl avertit qu'il ne sait pas relire les formes/dessins DrawingML d'un
# .xlsx. On ne lit que les données des cellules : ce message est sans effet sur
# la compilation. On le masque (ciblé) pour ne pas polluer la console utilisateur.
warnings.filterwarnings(
    "ignore",
    message="DrawingML support is incomplete",
    category=UserWarning,
    module="openpyxl",
)

from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QIcon

from ui.main_window import MainWindow
from utils import logger


def _log_dir() -> Path:
    """Dossier des logs, inscriptible aussi bien en dev qu'une fois packagé.

    En développement on garde ``./logs``. Packagée, l'app peut être installée
    dans un emplacement non inscriptible (Program Files) : on bascule alors sur
    ``%LOCALAPPDATA%\\ExcelCompiler\\logs``.
    """
    if getattr(sys, "frozen", False):
        import os
        base = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "ExcelCompiler"
        return base / "logs"
    return Path("logs")


def setup_logging():
    """Configure le logging pour l'application"""
    # Créer dossier logs s'il n'existe pas (emplacement inscriptible)
    log_dir = _log_dir()
    log_dir.mkdir(parents=True, exist_ok=True)

    # Configuration du logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_dir / 'excelcompiler.log', encoding='utf-8'),
            logging.StreamHandler()
        ]
    )


def main():
    """
    Fonction principale de l'application
    """
    # Setup logging
    setup_logging()
    logger.info("=" * 80)
    logger.info("Démarrage ExcelCompiler v3.2")
    logger.info("=" * 80)

    # Créer l'application Qt
    app = QApplication(sys.argv)
    app.setApplicationName("ExcelCompiler")
    app.setApplicationVersion("3.2")
    app.setOrganizationName("GOUNOU N'GOBI Chabi Zimé")

    # Restaurer le thème mémorisé (clair par défaut) avant de styler l'UI
    from PyQt6.QtCore import QSettings
    from ui.styles import build_stylesheet
    from ui.styles import theme as T

    saved_theme = QSettings(
        "GOUNOU N'GOBI Chabi Zimé", "ExcelCompiler"
    ).value("theme", "light")
    T.set_theme(saved_theme if saved_theme in ("light", "dark") else "light")

    # Appliquer le design system premium (feuille de style globale)
    app.setStyleSheet(build_stylesheet())

    # Définir l'icône de l'application (si disponible)
    from utils import resource_path
    icon_path = resource_path("icon.png")
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))

    # Écran de démarrage (splash) pendant le chargement
    from ui.splash import make_splash
    splash = make_splash()
    splash.show()
    app.processEvents()

    # Créer et afficher la fenêtre principale
    window = MainWindow()
    window.show()
    splash.finish(window)

    logger.info("Fenêtre principale affichée")

    # Lancer la boucle d'événements
    exit_code = app.exec()

    logger.info(f"Application terminée (code: {exit_code})")
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
