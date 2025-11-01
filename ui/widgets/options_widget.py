"""
Widget de configuration des options de compilation
Auteur: GOUNOU N'GOBI Chabi Zimé
"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QGroupBox,
    QCheckBox, QLabel, QSpinBox, QLineEdit, QComboBox
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QFont

from core.compilation import CompilationOptions, FilenameOption, OutputFormat
from ui.styles import EXCEL_GREEN, NEUTRAL_DARK, OFFICE_ORANGE


class OptionsWidget(QWidget):
    """
    Widget pour configurer les options de compilation

    Mode AUTO par défaut avec possibilité de passer en mode MANUEL
    """

    options_changed = pyqtSignal()  # Émis quand les options changent

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        """Configure l'interface du widget"""
        layout = QVBoxLayout(self)

        # Group box principal
        group = QGroupBox("⚙️ Options de compilation")
        group.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        group_layout = QVBoxLayout()

        # === MODE DÉTECTION AUTOMATIQUE ===
        self.checkbox_auto_detect = QCheckBox("✨ Détection automatique (recommandé)")
        self.checkbox_auto_detect.setChecked(True)  # COCHÉ PAR DÉFAUT
        self.checkbox_auto_detect.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        self.checkbox_auto_detect.setStyleSheet(f"color: {EXCEL_GREEN};")
        self.checkbox_auto_detect.toggled.connect(self.on_auto_detect_toggled)
        group_layout.addWidget(self.checkbox_auto_detect)

        # Label explicatif
        help_label = QLabel("Le système détecte automatiquement la structure de chaque fichier")
        help_label.setStyleSheet("color: #666; font-size: 8pt; font-style: italic; padding-left: 25px;")
        group_layout.addWidget(help_label)

        # === CONFIGURATION MANUELLE (cachée par défaut) ===
        self.manual_group = QGroupBox("Configuration manuelle")
        self.manual_group.setVisible(False)  # CACHÉ par défaut
        manual_layout = QGridLayout()

        # Ligne début en-têtes
        self.label_header_start = QLabel("Ligne début en-têtes:")
        self.spinbox_header_start = QSpinBox()
        self.spinbox_header_start.setMinimum(1)
        self.spinbox_header_start.setMaximum(20)
        self.spinbox_header_start.setValue(1)
        self.spinbox_header_start.valueChanged.connect(self.options_changed.emit)

        # Nombre lignes en-tête
        self.label_header_rows = QLabel("Nombre lignes en-tête:")
        self.spinbox_header_rows = QSpinBox()
        self.spinbox_header_rows.setMinimum(1)
        self.spinbox_header_rows.setMaximum(15)
        self.spinbox_header_rows.setValue(1)
        self.spinbox_header_rows.valueChanged.connect(self.options_changed.emit)

        # Warning manuel
        warning_label = QLabel("⚠️ Mode manuel: assurez-vous que tous les fichiers ont la même structure")
        warning_label.setStyleSheet(f"color: {OFFICE_ORANGE}; font-size: 8pt; padding: 5px;")
        warning_label.setWordWrap(True)

        manual_layout.addWidget(self.label_header_start, 0, 0)
        manual_layout.addWidget(self.spinbox_header_start, 0, 1)
        manual_layout.addWidget(self.label_header_rows, 1, 0)
        manual_layout.addWidget(self.spinbox_header_rows, 1, 1)
        manual_layout.addWidget(warning_label, 2, 0, 1, 2)

        self.manual_group.setLayout(manual_layout)
        group_layout.addWidget(self.manual_group)

        # === OPTIONS COMMUNES ===
        common_layout = QGridLayout()

        # Nom de fichier
        label_filename_option = QLabel("📋 Nom de fichier:")
        self.combo_filename_option = QComboBox()
        self.combo_filename_option.addItem("Aucun", FilenameOption.NONE)
        self.combo_filename_option.addItem("Avec extension", FilenameOption.WITH_EXTENSION)
        self.combo_filename_option.addItem("Sans extension", FilenameOption.WITHOUT_EXTENSION)
        self.combo_filename_option.setCurrentIndex(2)  # Sans extension par défaut
        self.combo_filename_option.currentIndexChanged.connect(self.options_changed.emit)

        # Fichier de sortie
        label_output = QLabel("📄 Fichier de sortie:")
        self.lineedit_output = QLineEdit("compilation.xlsx")
        self.lineedit_output.setPlaceholderText("Nom du fichier de sortie")
        self.lineedit_output.textChanged.connect(self.options_changed.emit)

        # Format de sortie
        label_format = QLabel("💾 Format:")
        self.combo_format = QComboBox()
        self.combo_format.addItem("Excel (.xlsx)", OutputFormat.XLSX)
        self.combo_format.addItem("CSV (.csv)", OutputFormat.CSV)
        self.combo_format.addItem("TSV (.tsv)", OutputFormat.TSV)
        self.combo_format.currentIndexChanged.connect(self.options_changed.emit)

        # Répéter en-têtes
        self.checkbox_repeat_headers = QCheckBox("Répéter en-têtes entre fichiers")
        self.checkbox_repeat_headers.setChecked(False)
        self.checkbox_repeat_headers.toggled.connect(self.options_changed.emit)

        # Supprimer lignes vides
        self.checkbox_remove_empty = QCheckBox("Supprimer lignes vides")
        self.checkbox_remove_empty.setChecked(True)
        self.checkbox_remove_empty.toggled.connect(self.options_changed.emit)

        common_layout.addWidget(label_filename_option, 0, 0)
        common_layout.addWidget(self.combo_filename_option, 0, 1)
        common_layout.addWidget(label_output, 1, 0)
        common_layout.addWidget(self.lineedit_output, 1, 1)
        common_layout.addWidget(label_format, 2, 0)
        common_layout.addWidget(self.combo_format, 2, 1)
        common_layout.addWidget(self.checkbox_repeat_headers, 3, 0, 1, 2)
        common_layout.addWidget(self.checkbox_remove_empty, 4, 0, 1, 2)

        group_layout.addLayout(common_layout)

        group.setLayout(group_layout)
        layout.addWidget(group)

    def on_auto_detect_toggled(self, checked):
        """Affiche/masque les options manuelles"""
        self.manual_group.setVisible(not checked)
        self.options_changed.emit()

    def get_compilation_options(self) -> CompilationOptions:
        """
        Construit les options de compilation selon le mode choisi

        Returns:
            CompilationOptions configuré
        """
        # Récupérer le filename_option
        filename_option = self.combo_filename_option.currentData()

        if self.checkbox_auto_detect.isChecked():
            # MODE AUTO
            return CompilationOptions(
                auto_detect_structure=True,
                enable_cross_validation=True,
                detection_confidence_threshold=0.65,
                filename_option=filename_option,
                repeat_headers=self.checkbox_repeat_headers.isChecked(),
                remove_empty_rows=self.checkbox_remove_empty.isChecked()
            )
        else:
            # MODE MANUEL
            return CompilationOptions(
                auto_detect_structure=False,
                manual_header_start_row=self.spinbox_header_start.value(),
                manual_header_rows=self.spinbox_header_rows.value(),
                filename_option=filename_option,
                repeat_headers=self.checkbox_repeat_headers.isChecked(),
                remove_empty_rows=self.checkbox_remove_empty.isChecked()
            )

    def get_output_file(self) -> str:
        """Retourne le nom du fichier de sortie"""
        return self.lineedit_output.text().strip()

    def get_output_format(self) -> OutputFormat:
        """Retourne le format de sortie"""
        return self.combo_format.currentData()

    def is_auto_detect_enabled(self) -> bool:
        """Vérifie si la détection auto est activée"""
        return self.checkbox_auto_detect.isChecked()
