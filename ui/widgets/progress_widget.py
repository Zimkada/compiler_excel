"""
Widget de progression avec possibilité d'annulation
Auteur: GOUNOU N'GOBI Chabi Zimé
"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QProgressBar, QLabel, QPushButton
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QFont

from ui.styles import theme as T


class ProgressWidget(QWidget):
    """
    Widget affichant la progression de la compilation
    avec bouton d'annulation
    """

    cancel_requested = pyqtSignal()  # Émis quand l'utilisateur annule

    def __init__(self, parent=None):
        super().__init__(parent)
        self._status_state = "idle"  # idle | success | error
        self.setup_ui()
        self.hide()  # Caché par défaut
        T.manager.theme_changed.connect(self.apply_theme)

    def setup_ui(self):
        """Configure l'interface du widget"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)

        # Label de status
        self.label_status = QLabel("")
        self.label_status.setFont(QFont("Segoe UI", 9))
        self.label_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.label_status)

        # Barre de progression
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        layout.addWidget(self.progress_bar)

        # Bouton annuler
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.button_cancel = QPushButton("❌ Annuler")
        self.button_cancel.setFont(QFont("Segoe UI", 9))
        self.button_cancel.clicked.connect(self.on_cancel_clicked)

        button_layout.addWidget(self.button_cancel)
        button_layout.addStretch()
        layout.addLayout(button_layout)

        self.apply_theme()

    def apply_theme(self):
        """(Ré)applique les styles inline selon le thème actif."""
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{
                border: 1px solid {T.BORDER_STRONG};
                border-radius: 5px;
                text-align: center;
                height: 25px;
                background-color: {T.BG_SUBTLE};
                color: {T.TEXT_SECONDARY};
            }}
            QProgressBar::chunk {{
                background-color: {T.ACCENT};
                border-radius: 3px;
            }}
        """)
        self.button_cancel.setStyleSheet(f"""
            QPushButton {{
                background-color: {T.DANGER};
                color: #FFFFFF;
                padding: 6px 12px;
                border: none;
                border-radius: {T.RADIUS_SM}px;
            }}
            QPushButton:hover {{
                background-color: {T.DANGER_SOFT};
                color: {T.DANGER};
            }}
            QPushButton:disabled {{
                background-color: {T.BG_SUBTLE};
                color: {T.TEXT_MUTED};
            }}
        """)
        self._apply_status_style()

    def _apply_status_style(self):
        """Couleur du label de statut selon l'état courant et le thème."""
        if self._status_state == "success":
            self.label_status.setStyleSheet(
                f"color: {T.SUCCESS}; font-weight: bold; padding: 5px;")
        elif self._status_state == "error":
            self.label_status.setStyleSheet(
                f"color: {T.DANGER}; font-weight: bold; padding: 5px;")
        else:
            self.label_status.setStyleSheet(
                f"color: {T.TEXT_SECONDARY}; padding: 5px;")

    def start(self, message: str = "Compilation en cours..."):
        """
        Démarre l'affichage de progression

        Args:
            message: Message à afficher
        """
        self._status_state = "idle"
        self._apply_status_style()
        self.label_status.setText(message)
        self.progress_bar.setValue(0)
        self.button_cancel.setEnabled(True)
        self.show()

    def update_progress(self, value: int, message: str = ""):
        """
        Met à jour la progression

        Args:
            value: Pourcentage (0-100)
            message: Message optionnel
        """
        self.progress_bar.setValue(value)
        if message:
            self.label_status.setText(message)

    def finish_success(self, message: str = "Compilation terminée!"):
        """
        Termine avec succès

        Args:
            message: Message de succès
        """
        self.progress_bar.setValue(100)
        self.label_status.setText(f"✅ {message}")
        self._status_state = "success"
        self._apply_status_style()
        self.button_cancel.setEnabled(False)

    def finish_error(self, message: str = "Erreur lors de la compilation"):
        """
        Termine avec erreur

        Args:
            message: Message d'erreur
        """
        self.label_status.setText(f"❌ {message}")
        self._status_state = "error"
        self._apply_status_style()
        self.button_cancel.setEnabled(False)

    def reset(self):
        """Réinitialise le widget"""
        self.progress_bar.setValue(0)
        self.label_status.setText("")
        self._status_state = "idle"
        self._apply_status_style()
        self.button_cancel.setEnabled(True)
        self.hide()

    def on_cancel_clicked(self):
        """Appelé quand l'utilisateur clique sur Annuler"""
        self.button_cancel.setEnabled(False)
        self.label_status.setText("Annulation en cours...")
        self.cancel_requested.emit()
