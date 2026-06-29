"""
Widget de sélection de fichiers
Auteur: GOUNOU N'GOBI Chabi Zimé
"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGroupBox,
    QPushButton, QListWidget, QLabel, QCheckBox,
    QFileDialog, QListWidgetItem
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QFont
from pathlib import Path
from typing import List

from ui.styles import EXCEL_GREEN, NEUTRAL_DARK


class FileSelectorWidget(QWidget):
    """
    Widget pour sélectionner les fichiers à compiler

    Signaux:
    - files_selected: Émis quand la sélection de fichiers change
    """

    files_selected = pyqtSignal(list)  # List[str] des fichiers sélectionnés

    def __init__(self, parent=None):
        super().__init__(parent)
        self.directory = ""
        self.all_files = []  # Tous les fichiers du dossier
        self.setup_ui()

    def setup_ui(self):
        """Configure l'interface du widget"""
        layout = QVBoxLayout(self)

        # Group box principal
        group = QGroupBox("📁 Sélection des fichiers")
        group.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        group_layout = QVBoxLayout()

        # Sélection du dossier
        dir_layout = QHBoxLayout()
        self.label_directory = QLabel("Aucun dossier sélectionné")
        self.label_directory.setStyleSheet(f"color: {NEUTRAL_DARK}; padding: 5px;")

        self.button_choose_directory = QPushButton("📂 Choisir un dossier")
        self.button_choose_directory.setFont(QFont("Segoe UI", 9))
        self.button_choose_directory.setProperty("variant", "primary")
        self.button_choose_directory.setCursor(Qt.CursorShape.PointingHandCursor)
        self.button_choose_directory.clicked.connect(self.choose_directory)

        dir_layout.addWidget(self.label_directory, 1)
        dir_layout.addWidget(self.button_choose_directory)
        group_layout.addLayout(dir_layout)

        # Barre de sélection
        selection_bar = QHBoxLayout()
        self.checkbox_select_all = QCheckBox("Tout sélectionner")
        self.checkbox_select_all.stateChanged.connect(self.toggle_select_all)

        self.label_file_count = QLabel("0 fichier sélectionné")
        self.label_file_count.setStyleSheet(f"color: {NEUTRAL_DARK}; font-weight: bold;")

        selection_bar.addWidget(self.checkbox_select_all)
        selection_bar.addStretch()
        selection_bar.addWidget(self.label_file_count)
        group_layout.addLayout(selection_bar)

        # Liste des fichiers
        self.list_files = QListWidget()
        self.list_files.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.list_files.setMinimumHeight(150)
        self.list_files.setMaximumHeight(250)
        self.list_files.itemSelectionChanged.connect(self.on_selection_changed)
        # Style hérité du design system global (QListWidget)

        group_layout.addWidget(self.list_files)

        group.setLayout(group_layout)
        layout.addWidget(group)

    def choose_directory(self):
        """Ouvre dialogue de sélection de dossier"""
        directory = QFileDialog.getExistingDirectory(
            self,
            "Sélectionner un dossier contenant des fichiers Excel/CSV",
            "",
            QFileDialog.Option.ShowDirsOnly
        )

        if directory:
            self.directory = directory
            self.load_files_from_directory(directory)

    def load_files_from_directory(self, directory: str):
        """
        Charge les fichiers Excel/CSV du dossier

        Args:
            directory: Chemin du dossier
        """
        self.list_files.clear()
        self.all_files = []

        # Extensions supportées
        extensions = ['.xlsx', '.xls', '.xlsm', '.csv', '.tsv']

        # Chercher fichiers
        path = Path(directory)
        for ext in extensions:
            for file_path in path.glob(f"*{ext}"):
                if file_path.is_file():
                    self.all_files.append(str(file_path))

        # Trier par nom
        self.all_files.sort(key=lambda x: Path(x).name.lower())

        # Ajouter à la liste
        for file_path in self.all_files:
            item = QListWidgetItem(Path(file_path).name)
            item.setData(Qt.ItemDataRole.UserRole, file_path)  # Stocker chemin complet
            self.list_files.addItem(item)

        # Mettre à jour le label du dossier
        short_path = str(Path(directory).name)
        self.label_directory.setText(f"📁 {short_path}")

        # Mettre à jour le compteur
        self.update_file_count()

        # Auto-sélectionner tous les fichiers par défaut
        if self.all_files:
            self.list_files.selectAll()
            self.checkbox_select_all.setChecked(True)

    def toggle_select_all(self, state):
        """Sélectionne/désélectionne tous les fichiers"""
        if state == Qt.CheckState.Checked.value:
            self.list_files.selectAll()
        else:
            self.list_files.clearSelection()

    def on_selection_changed(self):
        """Appelé quand la sélection change"""
        self.update_file_count()

        # Émettre signal avec les fichiers sélectionnés
        selected_files = self.get_selected_files()
        self.files_selected.emit(selected_files)

    def update_file_count(self):
        """Met à jour le compteur de fichiers sélectionnés"""
        selected_count = len(self.list_files.selectedItems())
        total_count = self.list_files.count()

        if selected_count == 0:
            self.label_file_count.setText("Aucun fichier sélectionné")
        elif selected_count == 1:
            self.label_file_count.setText("1 fichier sélectionné")
        else:
            self.label_file_count.setText(f"{selected_count}/{total_count} fichiers sélectionnés")

        # Mettre à jour la checkbox "tout sélectionner"
        if selected_count == total_count and total_count > 0:
            self.checkbox_select_all.setChecked(True)
        else:
            self.checkbox_select_all.setChecked(False)

    def get_selected_files(self) -> List[str]:
        """
        Retourne la liste des fichiers sélectionnés

        Returns:
            Liste des chemins complets des fichiers sélectionnés
        """
        selected_files = []
        for item in self.list_files.selectedItems():
            file_path = item.data(Qt.ItemDataRole.UserRole)
            selected_files.append(file_path)
        return selected_files

    def has_files_selected(self) -> bool:
        """Vérifie si au moins un fichier est sélectionné"""
        return len(self.list_files.selectedItems()) > 0
