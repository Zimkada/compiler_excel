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

from ui.styles import theme as T

SUPPORTED_EXTENSIONS = ['.xlsx', '.xlsm', '.csv', '.tsv']


class _PassthroughListWidget(QListWidget):
    """QListWidget qui laisse remonter la molette au QScrollArea parent
    quand son propre contenu n'a rien à défiler dans la direction voulue."""

    def wheelEvent(self, event):
        bar = self.verticalScrollBar()
        at_top = bar.value() == bar.minimum()
        at_bottom = bar.value() == bar.maximum()
        scrolling_up = event.angleDelta().y() > 0
        if (at_top and scrolling_up) or (at_bottom and not scrolling_up):
            event.ignore()
            return
        super().wheelEvent(event)


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
        self.setAcceptDrops(True)  # Activer le glisser-déposer de fichiers
        T.get_manager().theme_changed.connect(self.apply_theme)

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
        self.label_directory.setStyleSheet(f"color: {T.TEXT_SECONDARY}; padding: 5px;")

        self.button_choose_directory = QPushButton("📂 Choisir un dossier")
        self.button_choose_directory.setFont(QFont("Segoe UI", 9))
        self.button_choose_directory.setProperty("variant", "primary")
        self.button_choose_directory.setCursor(Qt.CursorShape.PointingHandCursor)
        self.button_choose_directory.clicked.connect(self.choose_directory)

        dir_layout.addWidget(self.label_directory, 1)
        dir_layout.addWidget(self.button_choose_directory)
        group_layout.addLayout(dir_layout)

        # Indice glisser-déposer
        self.drop_hint = QLabel("⤓  Glissez-déposez vos fichiers Excel/CSV ici")
        self.drop_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        group_layout.addWidget(self.drop_hint)

        # Barre de sélection
        selection_bar = QHBoxLayout()
        self.checkbox_select_all = QCheckBox("Tout sélectionner")
        self.checkbox_select_all.stateChanged.connect(self.toggle_select_all)

        self.label_file_count = QLabel("0 fichier sélectionné")
        self.label_file_count.setStyleSheet(f"color: {T.TEXT_SECONDARY}; font-weight: bold;")

        selection_bar.addWidget(self.checkbox_select_all)
        selection_bar.addStretch()
        selection_bar.addWidget(self.label_file_count)
        group_layout.addLayout(selection_bar)

        # Liste des fichiers
        self.list_files = _PassthroughListWidget()
        self.list_files.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.list_files.setMinimumHeight(150)
        self.list_files.setMaximumHeight(250)
        self.list_files.itemSelectionChanged.connect(self.on_selection_changed)
        # Style hérité du design system global (QListWidget)

        group_layout.addWidget(self.list_files)

        group.setLayout(group_layout)
        layout.addWidget(group)

        self.apply_theme()

    def apply_theme(self):
        """(Ré)applique les styles inline dépendant du thème actif."""
        self.label_directory.setStyleSheet(
            f"color: {T.TEXT_SECONDARY}; padding: 5px;")
        self.label_file_count.setStyleSheet(
            f"color: {T.TEXT_SECONDARY}; font-weight: bold;")
        self._reset_drop_hint_style()

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
        path = Path(directory)
        files = []
        for ext in SUPPORTED_EXTENSIONS:
            for file_path in path.glob(f"*{ext}"):
                if file_path.is_file():
                    files.append(str(file_path))

        self._populate_list(files)
        self.label_directory.setText(f"📁 {path.name}")

    def add_files(self, paths: List[str]):
        """Ajoute des fichiers (ex. via glisser-déposer), en filtrant les
        extensions supportées et en évitant les doublons."""
        combined = list(self.all_files)
        for p in paths:
            fp = Path(p)
            if fp.is_file() and fp.suffix.lower() in SUPPORTED_EXTENSIONS:
                if str(fp) not in combined:
                    combined.append(str(fp))
        self._populate_list(combined)
        if combined and not self.directory:
            self.label_directory.setText(f"📁 {Path(combined[0]).parent.name}")

    def _populate_list(self, files: List[str]):
        """Remplit la liste avec les fichiers donnés (triés), met à jour les
        compteurs, l'indice de glisser-déposer et la sélection par défaut."""
        self.list_files.clear()
        self.all_files = sorted(files, key=lambda x: Path(x).name.lower())

        for file_path in self.all_files:
            item = QListWidgetItem(Path(file_path).name)
            item.setData(Qt.ItemDataRole.UserRole, file_path)
            self.list_files.addItem(item)

        # Masquer l'indice de drop dès qu'il y a des fichiers
        self.drop_hint.setVisible(not bool(self.all_files))

        self.update_file_count()

        if self.all_files:
            self.list_files.selectAll()
            self.checkbox_select_all.setChecked(True)

    # ── Glisser-déposer ──────────────────────────────────────────────────
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.drop_hint.setStyleSheet(
                f"color: {T.ACCENT}; font-size: {T.FONT_SIZE_SM}pt; font-weight: 700; "
                f"border: 1.5px dashed {T.ACCENT}; border-radius: {T.RADIUS_MD}px; "
                f"padding: 14px; margin-top: 4px; background-color: {T.ACCENT_SOFT};"
            )
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self._reset_drop_hint_style()

    def dropEvent(self, event):
        paths = [u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()]
        self._reset_drop_hint_style()
        if not paths:
            return
        # Si un dossier est déposé, charger son contenu ; sinon ajouter les fichiers
        dirs = [p for p in paths if Path(p).is_dir()]
        files = [p for p in paths if Path(p).is_file()]
        if dirs:
            self.directory = dirs[0]
            self.load_files_from_directory(dirs[0])
        if files:
            self.add_files(files)
        event.acceptProposedAction()

    def _reset_drop_hint_style(self):
        self.drop_hint.setStyleSheet(
            f"color: {T.TEXT_MUTED}; font-size: {T.FONT_SIZE_SM}pt; "
            f"border: 1.5px dashed {T.BORDER_STRONG}; border-radius: {T.RADIUS_MD}px; "
            f"padding: 14px; margin-top: 4px;"
        )

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
