"""
Widget d'affichage des résultats de compilation
Auteur: GOUNOU N'GOBI Chabi Zimé
"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QTextEdit, QPushButton,
    QHBoxLayout, QLabel
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from pathlib import Path
import subprocess
import platform

from core.compilation import CompilationResult
from ui.styles import theme as T


class ResultsWidget(QWidget):
    """
    Widget pour afficher les résultats et statistiques de compilation
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.last_output_file = None
        self._last_result = None  # mémorise le dernier rendu pour re-thématiser
        self.setup_ui()
        T.manager.theme_changed.connect(self.apply_theme)

    def setup_ui(self):
        """Configure l'interface du widget"""
        layout = QVBoxLayout(self)

        # Group box principal
        group = QGroupBox("✅ Résultats de compilation")
        group.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        group_layout = QVBoxLayout()

        # Statistiques (label résumé)
        self.label_stats = QLabel("")
        self.label_stats.setFont(QFont("Segoe UI", 9))
        self.label_stats.setWordWrap(True)
        self.label_stats.setTextFormat(Qt.TextFormat.RichText)
        group_layout.addWidget(self.label_stats)

        # Zone de texte pour logs détaillés
        self.text_logs = QTextEdit()
        self.text_logs.setReadOnly(True)
        self.text_logs.setFont(QFont("Consolas", 8))
        self.text_logs.setMinimumHeight(150)
        self.text_logs.setMaximumHeight(200)
        group_layout.addWidget(self.text_logs)

        # Boutons d'action
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.button_open_folder = QPushButton("📁 Ouvrir le dossier")
        self.button_open_folder.setFont(QFont("Segoe UI", 9))
        self.button_open_folder.setProperty("variant", "primary")
        self.button_open_folder.setCursor(Qt.CursorShape.PointingHandCursor)
        self.button_open_folder.setEnabled(False)
        self.button_open_folder.clicked.connect(self.open_output_folder)

        self.button_clear = QPushButton("🗑️ Effacer")
        self.button_clear.setFont(QFont("Segoe UI", 9))
        self.button_clear.clicked.connect(self.clear)

        button_layout.addWidget(self.button_open_folder)
        button_layout.addWidget(self.button_clear)
        button_layout.addStretch()
        group_layout.addLayout(button_layout)

        group.setLayout(group_layout)
        layout.addWidget(group)

        self.apply_theme()

        # Initialiser avec message vide
        self.clear()

    def apply_theme(self):
        """(Ré)applique les styles inline et re-rend le contenu HTML au thème."""
        self.label_stats.setStyleSheet(
            f"color: {T.TEXT_PRIMARY}; padding: 12px; "
            f"background-color: {T.BG_SUBTLE}; border-radius: {T.RADIUS_MD}px;"
        )
        self.text_logs.setStyleSheet(
            f"QTextEdit {{ border: 1px solid {T.BORDER}; "
            f"border-radius: {T.RADIUS_SM}px; background-color: {T.BG_SUBTLE}; "
            f"color: {T.TEXT_PRIMARY}; padding: 8px; }}"
        )
        # Re-rendre le résumé HTML avec les couleurs du thème courant
        state, payload = self._last_result or ("empty", None)
        if state == "result":
            self._render_stats(payload)
        elif state == "error":
            self._render_error(payload)
        else:
            self._render_empty()

    def display_result(self, result: CompilationResult):
        """
        Affiche les résultats de compilation

        Args:
            result: Résultat de compilation
        """
        # Sauvegarder le fichier de sortie
        self.last_output_file = result.output_file
        self._last_result = ("result", result)

        # Afficher statistiques (rendu thème-aware)
        self._render_stats(result)
        self.button_open_folder.setEnabled(bool(result.success))

        # Afficher logs détaillés
        self.text_logs.clear()
        self.text_logs.append("=== RÉSUMÉ DÉTAILLÉ ===\n")
        self.text_logs.append(result.get_summary())

        # Afficher warnings si présents
        if result.warnings:
            self.text_logs.append("\n=== AVERTISSEMENTS ===\n")
            for warning in result.warnings:
                self.text_logs.append(f"⚠️  {warning}")

        # Scroll vers le haut
        self.text_logs.verticalScrollBar().setValue(0)

    def _render_stats(self, result: CompilationResult):
        """Construit le résumé HTML des statistiques aux couleurs du thème."""
        if result.success:
            stats_html = f"""
            <div style='color: {T.SUCCESS};'>
                <b>✅ Compilation réussie</b>
            </div>
            <hr>
            <b>Statistiques:</b><br>
            • Fichiers traités: <b>{result.successful_files}/{result.total_files}</b><br>
            • Lignes compilées: <b>{result.total_rows}</b><br>
            • Temps total: <b>{result.total_processing_time:.2f}s</b><br>
            • Fichier de sortie: <b>{Path(result.output_file).name}</b>
            """
            if result.auto_detection_used:
                stats_html += f"<br>• Détection auto: <b>{result.detection_success_rate:.0%}</b> réussite"
        else:
            stats_html = f"""
            <div style='color: {T.DANGER};'>
                <b>❌ Compilation échouée</b>
            </div>
            <hr>
            • Fichiers réussis: {result.successful_files}/{result.total_files}<br>
            • Fichiers échoués: {result.failed_files}
            """
        self.label_stats.setText(stats_html)

    def _render_error(self, error_message: str):
        """Construit le message d'erreur HTML aux couleurs du thème."""
        self.label_stats.setText(f"""
        <div style='color: {T.DANGER};'>
            <b>❌ Erreur</b>
        </div>
        <hr>
        {error_message}
        """)

    def _render_empty(self):
        """Affiche l'état vide engageant aux couleurs du thème."""
        self.label_stats.setText(
            "<div style='text-align:center; padding:18px;'>"
            "<span style='font-size:30pt;'>✅</span><br>"
            "<b style='font-size:12pt;'>Aucune compilation pour l'instant</b><br>"
            f"<span style='color:{T.TEXT_SECONDARY};'>Lancez une compilation depuis l'onglet "
            "« Compilation » : les statistiques et détails s'afficheront ici.</span>"
            "</div>"
        )

    def display_error(self, error_message: str):
        """
        Affiche un message d'erreur

        Args:
            error_message: Message d'erreur
        """
        self._last_result = ("error", error_message)
        self._render_error(error_message)

        self.text_logs.clear()
        self.text_logs.append("=== ERREUR ===\n")
        self.text_logs.append(error_message)

        self.button_open_folder.setEnabled(False)

    def open_output_folder(self):
        """Ouvre le dossier contenant le fichier de sortie"""
        if not self.last_output_file:
            return

        output_path = Path(self.last_output_file)
        if not output_path.exists():
            return

        folder_path = output_path.parent

        # Ouvrir le dossier selon l'OS
        try:
            if platform.system() == "Windows":
                subprocess.run(['explorer', str(folder_path)])
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(['open', str(folder_path)])
            else:  # Linux
                subprocess.run(['xdg-open', str(folder_path)])
        except Exception as e:
            self.text_logs.append(f"\n❌ Erreur ouverture dossier: {e}")

    def clear(self):
        """Efface les résultats et affiche un état vide engageant."""
        self._last_result = ("empty", None)
        self._render_empty()
        self.text_logs.clear()
        self.text_logs.setPlaceholderText("Les détails de compilation apparaîtront ici…")
        self.button_open_folder.setEnabled(False)
        self.last_output_file = None
