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

        # === MODE 1: FICHIER DE RÉFÉRENCE (Semi-automatique) ===
        self.checkbox_reference_mode = QCheckBox("✨ Utiliser fichier de référence (recommandé)")
        self.checkbox_reference_mode.setChecked(True)  # COCHÉ PAR DÉFAUT
        self.checkbox_reference_mode.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        self.checkbox_reference_mode.setStyleSheet(f"color: {EXCEL_GREEN};")
        self.checkbox_reference_mode.toggled.connect(self.on_mode_changed)
        group_layout.addWidget(self.checkbox_reference_mode)

        # Label explicatif mode référence
        help_ref_label = QLabel(
            "Spécifiez l'en-tête du premier fichier, "
            "le système trouvera automatiquement les en-têtes similaires dans les autres"
        )
        help_ref_label.setStyleSheet("color: #666; font-size: 8pt; font-style: italic; padding-left: 25px;")
        help_ref_label.setWordWrap(True)
        group_layout.addWidget(help_ref_label)

        # Options mode référence
        self.reference_group = QGroupBox("Configuration du fichier de référence")
        self.reference_group.setVisible(True)  # VISIBLE par défaut
        ref_layout = QGridLayout()

        ref_label1 = QLabel("Ligne d'en-tête (1er fichier):")
        self.spinbox_ref_header = QSpinBox()
        self.spinbox_ref_header.setMinimum(1)
        self.spinbox_ref_header.setMaximum(50)
        self.spinbox_ref_header.setValue(1)
        self.spinbox_ref_header.valueChanged.connect(self.options_changed.emit)

        ref_label2 = QLabel("Nombre de lignes d'en-tête:")
        self.spinbox_ref_header_lines = QSpinBox()
        self.spinbox_ref_header_lines.setMinimum(1)
        self.spinbox_ref_header_lines.setMaximum(10)
        self.spinbox_ref_header_lines.setValue(1)
        self.spinbox_ref_header_lines.valueChanged.connect(self.options_changed.emit)

        ref_info_label = QLabel("💡 Le premier fichier sélectionné servira de référence")
        ref_info_label.setStyleSheet("color: #217346; font-size: 8pt; padding: 5px;")
        ref_info_label.setWordWrap(True)

        ref_layout.addWidget(ref_label1, 0, 0)
        ref_layout.addWidget(self.spinbox_ref_header, 0, 1)
        ref_layout.addWidget(ref_label2, 1, 0)
        ref_layout.addWidget(self.spinbox_ref_header_lines, 1, 1)
        ref_layout.addWidget(ref_info_label, 2, 0, 1, 2)

        self.reference_group.setLayout(ref_layout)
        group_layout.addWidget(self.reference_group)

        # === MODE 2: CONFIGURATION MANUELLE COMPLÈTE ===
        self.checkbox_manual_mode = QCheckBox("⚙️ Configuration manuelle pour tous les fichiers")
        self.checkbox_manual_mode.setChecked(False)
        self.checkbox_manual_mode.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        self.checkbox_manual_mode.toggled.connect(self.on_mode_changed)
        group_layout.addWidget(self.checkbox_manual_mode)

        # Label explicatif mode manuel
        help_manual_label = QLabel("Tous les fichiers utiliseront exactement la même configuration")
        help_manual_label.setStyleSheet("color: #666; font-size: 8pt; font-style: italic; padding-left: 25px;")
        help_manual_label.setWordWrap(True)
        group_layout.addWidget(help_manual_label)

        # Options mode manuel
        self.manual_group = QGroupBox("Configuration manuelle")
        self.manual_group.setVisible(False)  # CACHÉ par défaut
        manual_layout = QGridLayout()

        manual_label1 = QLabel("Ligne d'en-tête:")
        self.spinbox_header_start = QSpinBox()
        self.spinbox_header_start.setMinimum(1)
        self.spinbox_header_start.setMaximum(50)
        self.spinbox_header_start.setValue(1)
        self.spinbox_header_start.valueChanged.connect(self.options_changed.emit)

        manual_label2 = QLabel("Nombre de lignes d'en-tête:")
        self.spinbox_header_rows = QSpinBox()
        self.spinbox_header_rows.setMinimum(1)
        self.spinbox_header_rows.setMaximum(10)
        self.spinbox_header_rows.setValue(1)
        self.spinbox_header_rows.valueChanged.connect(self.options_changed.emit)

        # Warning manuel
        warning_label = QLabel("⚠️ Assurez-vous que tous les fichiers ont exactement la même structure")
        warning_label.setStyleSheet(f"color: {OFFICE_ORANGE}; font-size: 8pt; padding: 5px;")
        warning_label.setWordWrap(True)

        manual_layout.addWidget(manual_label1, 0, 0)
        manual_layout.addWidget(self.spinbox_header_start, 0, 1)
        manual_layout.addWidget(manual_label2, 1, 0)
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

    def on_mode_changed(self, checked):
        """Gère le changement de mode (référence/manuel)"""
        sender = self.sender()

        if sender == self.checkbox_reference_mode and checked:
            # Mode référence activé → désactiver manuel
            self.checkbox_manual_mode.blockSignals(True)
            self.checkbox_manual_mode.setChecked(False)
            self.checkbox_manual_mode.blockSignals(False)
            self.reference_group.setVisible(True)
            self.manual_group.setVisible(False)
        elif sender == self.checkbox_manual_mode and checked:
            # Mode manuel activé → désactiver référence
            self.checkbox_reference_mode.blockSignals(True)
            self.checkbox_reference_mode.setChecked(False)
            self.checkbox_reference_mode.blockSignals(False)
            self.reference_group.setVisible(False)
            self.manual_group.setVisible(True)
        elif not self.checkbox_reference_mode.isChecked() and not self.checkbox_manual_mode.isChecked():
            # Aucun mode sélectionné → forcer mode référence
            self.checkbox_reference_mode.blockSignals(True)
            self.checkbox_reference_mode.setChecked(True)
            self.checkbox_reference_mode.blockSignals(False)
            self.reference_group.setVisible(True)
            self.manual_group.setVisible(False)

        self.options_changed.emit()

    def get_compilation_options(self) -> CompilationOptions:
        """
        Construit les options de compilation selon le mode choisi

        Returns:
            CompilationOptions configuré
        """
        # Récupérer le filename_option
        filename_option = self.combo_filename_option.currentData()

        if self.checkbox_reference_mode.isChecked():
            # MODE 1: FICHIER DE RÉFÉRENCE (semi-automatique)
            return CompilationOptions(
                auto_detect_structure=False,  # Pas de détection hybride
                use_reference_mode=True,  # Nouveau mode
                reference_header_row=self.spinbox_ref_header.value(),
                reference_header_lines=self.spinbox_ref_header_lines.value(),
                filename_option=filename_option,
                repeat_headers=self.checkbox_repeat_headers.isChecked(),
                remove_empty_rows=self.checkbox_remove_empty.isChecked()
            )
        else:
            # MODE 2: CONFIGURATION MANUELLE COMPLÈTE
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

    def is_reference_mode_enabled(self) -> bool:
        """Vérifie si le mode référence est activé"""
        return self.checkbox_reference_mode.isChecked()

    def is_manual_mode_enabled(self) -> bool:
        """Vérifie si le mode manuel est activé"""
        return self.checkbox_manual_mode.isChecked()
