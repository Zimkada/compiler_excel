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
                # N'afficher que ce qui explique VRAIMENT l'échec. Les
                # avertissements informatifs (colonnes parasites ignorées,
                # colonne inconnue ajoutée, lignes de total exclues…) accompagnent
                # aussi les compilations réussies : les présenter comme la cause
                # de l'échec envoyait l'utilisateur vérifier ses colonnes alors
                # que le problème était ailleurs (aucune donnée exploitable).
                error_msg = self._build_failure_message(result)
                logger.error(error_msg)
                self.error_occurred.emit(error_msg)

        except Exception as e:
            error_msg = f"Erreur compilation: {str(e)}"
            logger.error(error_msg, exc_info=True)
            self.error_occurred.emit(error_msg)

    @staticmethod
    def _build_failure_message(result) -> str:
        """Compose un message d'échec qui nomme la CAUSE, pas les à-côtés.

        Les avertissements sont de deux natures très différentes :
        - informatifs (colonnes parasites ignorées, colonne inconnue ajoutée,
          lignes de total ou de signature exclues) — ils accompagnent aussi
          les compilations parfaitement réussies ;
        - bloquants (« Aucune donnée compilée », « Erreur fatale : … ») — eux
          seuls expliquent un échec.

        On ne remonte que les seconds. À défaut, on dit simplement qu'aucune
        donnée exploitable n'a été trouvée, ce qui oriente l'utilisateur vers
        la bonne question (ligne d'en-tête, fichiers sélectionnés) plutôt que
        vers ses colonnes.
        """
        blocking_markers = ("Aucune donnée", "Erreur fatale", "annulée")
        blocking = [
            w for w in getattr(result, 'warnings', [])
            if any(m.lower() in w.lower() for m in blocking_markers)
        ]

        msg = "Compilation échouée"
        if blocking:
            return f"{msg} : {' · '.join(blocking[:3])}"

        failed = [
            f for f in getattr(result, 'file_results', [])
            if not getattr(f, 'success', True) and getattr(f, 'error_message', None)
        ]
        if failed:
            details = ', '.join(
                f"{Path(f.file_path).name} ({f.error_message})" for f in failed[:3]
            )
            return f"{msg} : {details}"

        return (
            f"{msg} : aucune donnée exploitable n'a été trouvée. "
            "Vérifiez la ligne d'en-tête indiquée et les fichiers sélectionnés."
        )

    def cancel(self):
        """Annule la compilation en cours"""
        self._is_cancelled = True
        logger.info("Demande d'annulation de compilation")
