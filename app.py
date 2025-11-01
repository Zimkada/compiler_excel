"""
Point d'entrée de l'application ExcelCompiler v3.2
Auteur: GOUNOU N'GOBI Chabi Zimé
"""

import sys
import logging
from pathlib import Path

from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QIcon

from ui.main_window import MainWindow
from utils import logger


def setup_logging():
    """Configure le logging pour l'application"""
    # Créer dossier logs s'il n'existe pas
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)

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

    # Définir l'icône de l'application (si disponible)
    icon_path = Path("icon.png")
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))

    # Créer et afficher la fenêtre principale
    window = MainWindow()
    window.show()

    logger.info("Fenêtre principale affichée")

    # Lancer la boucle d'événements
    exit_code = app.exec()

    logger.info(f"Application terminée (code: {exit_code})")
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
