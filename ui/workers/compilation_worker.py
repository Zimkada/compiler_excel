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

            # Émettre progression: détection
            if self.options.auto_detect_structure:
                self.progress_update.emit(10, f"Détection automatique de structure sur {len(self.file_paths)} fichiers...")

            # Compiler les fichiers
            result = compiler.compile_files(
                self.file_paths,
                self.output_file,
                self.output_format
            )

            # Vérifier si annulé pendant la compilation
            if self._is_cancelled:
                logger.info("Compilation annulée par l'utilisateur")
                self.error_occurred.emit("Compilation annulée")
                return

            # Émettre progression: finalisation
            self.progress_update.emit(90, "Finalisation...")

            # Vérifier le résultat
            if result.success:
                logger.info(f"Compilation réussie: {result.successful_files}/{result.total_files} fichiers")
                self.progress_update.emit(100, "Compilation terminée avec succès!")
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
