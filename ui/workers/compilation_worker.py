"""
Worker thread pour la compilation (wrapper ExcelCompiler v3.2)
Auteur: GOUNOU N'GOBI Chabi Zimé
"""

from PyQt6.QtCore import QThread, pyqtSignal
from pathlib import Path
from typing import List

from core.compilation import (
    ExcelCompiler,
    CompilationOptions,
    CompilationResult,
    OutputFormat
)
from utils import logger


class CompilationWorker(QThread):
    """
    Thread de compilation qui utilise ExcelCompiler v3.2

    Signaux émis:
    - progress_update: (int progress, str message)
    - compilation_finished: (CompilationResult result)
    - error_occurred: (str error_message)
    """

    # Signaux
    progress_update = pyqtSignal(int, str)  # (pourcentage, message)
    compilation_finished = pyqtSignal(object)  # CompilationResult
    error_occurred = pyqtSignal(str)  # Message d'erreur
    compilation_cancelled = pyqtSignal()  # Annulation volontaire (pas une erreur)

    def __init__(self,
                 file_paths: List[str],
                 output_file: str,
                 options: CompilationOptions,
                 output_format: OutputFormat = OutputFormat.XLSX):
        """
        Initialise le worker de compilation

        Args:
            file_paths: Liste chemins fichiers à compiler
            output_file: Chemin fichier de sortie
            options: Options de compilation
            output_format: Format de sortie (XLSX, CSV, TSV)
        """
        super().__init__()
        self.file_paths = file_paths
        self.output_file = output_file
        self.options = options
        self.output_format = output_format
        self._is_cancelled = False

    def run(self):
        """Exécute la compilation dans un thread séparé"""
        try:
            logger.info(f"Début compilation de {len(self.file_paths)} fichiers")

            # Émettre progression: démarrage
            self.progress_update.emit(0, "Initialisation de la compilation...")

            # Créer le compilateur
            compiler = ExcelCompiler(self.options)

            # Compiler les fichiers en transmettant la progression réelle
            # et la possibilité d'annuler. Les callbacks sont exécutés dans
            # ce thread; émettre un signal Qt depuis ici est thread-safe.
            result = compiler.compile_files(
                self.file_paths,
                self.output_file,
                self.output_format,
                progress_callback=lambda pct, msg: self.progress_update.emit(pct, msg),
                cancel_check=lambda: self._is_cancelled,
            )

            # La compilation s'est arrêtée proprement suite à une annulation:
            # ce n'est pas une erreur, on utilise un signal dédié.
            if result.cancelled or self._is_cancelled:
                logger.info("Compilation annulée par l'utilisateur")
                self.compilation_cancelled.emit()
                return

            # Vérifier le résultat
            if result.success:
                logger.info(f"Compilation réussie: {result.successful_files}/{result.total_files} fichiers")
                self.compilation_finished.emit(result)
            else:
                error_msg = "Compilation échouée"
                if result.warnings:
                    error_msg += f": {', '.join(result.warnings[:3])}"
                logger.error(error_msg)
                self.error_occurred.emit(error_msg)

        except Exception as e:
            error_msg = f"Erreur compilation: {str(e)}"
            logger.error(error_msg, exc_info=True)
            self.error_occurred.emit(error_msg)

    def cancel(self):
        """Annule la compilation en cours"""
        self._is_cancelled = True
        logger.info("Demande d'annulation de compilation")
