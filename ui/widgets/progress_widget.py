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

from ui.styles import EXCEL_GREEN, OFFICE_ORANGE, NEUTRAL_DARK


class ProgressWidget(QWidget):
    """
    Widget affichant la progression de la compilation
    avec bouton d'annulation
    """

    cancel_requested = pyqtSignal()  # Émis quand l'utilisateur annule

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        self.hide()  # Caché par défaut

    def setup_ui(self):
        """Configure l'interface du widget"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)

        # Label de status
        self.label_status = QLabel("")
        self.label_status.setFont(QFont("Segoe UI", 9))
        self.label_status.setStyleSheet(f"color: {NEUTRAL_DARK}; padding: 5px;")
        self.label_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.label_status)

        # Barre de progression
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{
                border: 2px solid #d0d0d0;
                border-radius: 5px;
                text-align: center;
                height: 25px;
                background-color: white;
            }}
            QProgressBar::chunk {{
                background-color: {EXCEL_GREEN};
                border-radius: 3px;
            }}
        """)
        layout.addWidget(self.progress_bar)

        # Bouton annuler
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.button_cancel = QPushButton("❌ Annuler")
        self.button_cancel.setFont(QFont("Segoe UI", 9))
        self.button_cancel.setStyleSheet(f"""
            QPushButton {{
                background-color: {OFFICE_ORANGE};
                color: white;
                padding: 6px 12px;
                border: none;
                border-radius: 4px;
            }}
            QPushButton:hover {{
                background-color: #c03301;
            }}
            QPushButton:disabled {{
                background-color: #cccccc;
                color: #666666;
            }}
        """)
        self.button_cancel.clicked.connect(self.on_cancel_clicked)

        button_layout.addWidget(self.button_cancel)
        button_layout.addStretch()
        layout.addLayout(button_layout)

    def start(self, message: str = "Compilation en cours..."):
        """
        Démarre l'affichage de progression

        Args:
            message: Message à afficher
        """
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
        self.label_status.setStyleSheet(f"color: {EXCEL_GREEN}; font-weight: bold; padding: 5px;")
        self.button_cancel.setEnabled(False)

    def finish_error(self, message: str = "Erreur lors de la compilation"):
        """
        Termine avec erreur

        Args:
            message: Message d'erreur
        """
        self.label_status.setText(f"❌ {message}")
        self.label_status.setStyleSheet(f"color: {OFFICE_ORANGE}; font-weight: bold; padding: 5px;")
        self.button_cancel.setEnabled(False)

    def reset(self):
        """Réinitialise le widget"""
        self.progress_bar.setValue(0)
        self.label_status.setText("")
        self.label_status.setStyleSheet(f"color: {NEUTRAL_DARK}; padding: 5px;")
        self.button_cancel.setEnabled(True)
        self.hide()

    def on_cancel_clicked(self):
        """Appelé quand l'utilisateur clique sur Annuler"""
        self.button_cancel.setEnabled(False)
        self.label_status.setText("Annulation en cours...")
        self.cancel_requested.emit()
