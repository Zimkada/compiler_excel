"""
Widget de configuration des options de compilation
Auteur: GOUNOU N'GOBI Chabi Zimé
"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QGroupBox,
    QCheckBox, QLabel, QSpinBox, QLineEdit, QComboBox,
    QRadioButton, QButtonGroup
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QFont

from core.compilation import CompilationOptions, FilenameOption, OutputFormat
from ui.styles import theme as T


class OptionsWidget(QWidget):
    """
    Widget pour configurer les options de compilation

    Mode AUTO par défaut avec possibilité de passer en mode MANUEL
    """

    options_changed = pyqtSignal()  # Émis quand les options changent

    # Organisation / application pour QSettings (mêmes clés que le thème).
    _SETTINGS_ORG = "GOUNOU N'GOBI Chabi Zimé"
    _SETTINGS_APP = "ExcelCompiler"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        self.restore_settings()
        T.get_manager().theme_changed.connect(self.apply_theme)

    def _settings(self):
        from PyQt6.QtCore import QSettings
        return QSettings(self._SETTINGS_ORG, self._SETTINGS_APP)

    def save_settings(self):
        """Mémorise les options courantes (mode, cases, format, date…) pour
        les restaurer au prochain lancement. Silencieux en cas d'échec."""
        try:
            s = self._settings()
            s.beginGroup("options")
            s.setValue("mode", self._current_mode_key())
            s.setValue("ref_header_row", self.spinbox_ref_header.value())
            s.setValue("ref_header_lines", self.spinbox_ref_header_lines.value())
            s.setValue("man_header_start", self.spinbox_header_start.value())
            s.setValue("man_header_rows", self.spinbox_header_rows.value())
            s.setValue("filename_option", self.combo_filename_option.currentIndex())
            s.setValue("output_format", self.combo_format.currentIndex())
            s.setValue("date_format", self.combo_date_format.currentIndex())
            s.setValue("output_name", self.lineedit_output.text())
            s.setValue("sort_column", self.lineedit_sort_column.text())
            for key, cb in self._persistent_checkboxes().items():
                s.setValue(key, cb.isChecked())
            s.endGroup()
        except Exception:
            pass

    def restore_settings(self):
        """Restaure les options mémorisées. En l'absence de réglages sauvés
        (premier lancement), conserve les valeurs par défaut de l'UI."""
        try:
            s = self._settings()
            s.beginGroup("options")
            if not s.childKeys() and not s.childGroups():
                s.endGroup()
                return
            mode = s.value("mode", "reference")
            self._apply_mode_key(mode)
            self.spinbox_ref_header.setValue(int(s.value("ref_header_row", 1)))
            self.spinbox_ref_header_lines.setValue(int(s.value("ref_header_lines", 1)))
            self.spinbox_header_start.setValue(int(s.value("man_header_start", 1)))
            self.spinbox_header_rows.setValue(int(s.value("man_header_rows", 1)))
            self.combo_filename_option.setCurrentIndex(int(s.value("filename_option", 2)))
            self.combo_format.setCurrentIndex(int(s.value("output_format", 0)))
            self.combo_date_format.setCurrentIndex(int(s.value("date_format", 0)))
            saved_name = s.value("output_name")
            if saved_name:
                self.lineedit_output.setText(str(saved_name))
            saved_sort = s.value("sort_column")
            if saved_sort is not None:
                self.lineedit_sort_column.setText(str(saved_sort))
            for key, cb in self._persistent_checkboxes().items():
                val = s.value(key)
                if val is not None:
                    cb.setChecked(val in (True, "true", "True", 1, "1"))
            s.endGroup()
            # Refléter le mode restauré sur la visibilité des panneaux.
            self.reference_group.setVisible(self.checkbox_reference_mode.isChecked())
            self.manual_group.setVisible(self.checkbox_manual_mode.isChecked())
            self.lineedit_sort_column.setEnabled(self.checkbox_sort_data.isChecked())
        except Exception:
            pass

    def _persistent_checkboxes(self):
        return {
            "repeat_headers": self.checkbox_repeat_headers,
            "remove_empty": self.checkbox_remove_empty,
            "unmerge_cells": self.checkbox_unmerge_cells,
            "flatten_headers": self.checkbox_flatten_headers,
            "align_columns": self.checkbox_align_columns,
            "drop_subtotals": self.checkbox_drop_subtotals,
            "mark_subtotals": self.checkbox_mark_subtotals,
            "remove_duplicates": self.checkbox_remove_duplicates,
            "sort_data": self.checkbox_sort_data,
        }

    def _current_mode_key(self) -> str:
        if self.checkbox_auto_mode.isChecked():
            return "auto"
        if self.checkbox_manual_mode.isChecked():
            return "manual"
        return "reference"

    def _apply_mode_key(self, key):
        target = {
            "auto": self.checkbox_auto_mode,
            "manual": self.checkbox_manual_mode,
        }.get(key, self.checkbox_reference_mode)
        target.setChecked(True)

    def setup_ui(self):
        """Configure l'interface du widget"""
        layout = QVBoxLayout(self)

        # === SECTION 1: MODE DE DÉTECTION ===
        detection_group = QGroupBox("📋 Mode de détection")
        detection_group.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        detection_layout = QVBoxLayout()

        # Groupe de boutons radio : exclusivité NATIVE des 3 modes (un seul
        # actif à la fois), au lieu de 3 cases à cocher pilotées à la main —
        # les cases suggéraient à tort qu'on pouvait en cocher plusieurs.
        self.mode_group = QButtonGroup(self)
        self.mode_group.setExclusive(True)

        # === MODE 0: DÉTECTION AUTOMATIQUE (sans aucune saisie) ===
        self.checkbox_auto_mode = QRadioButton("🤖 Détection automatique")
        self.checkbox_auto_mode.setChecked(False)
        self.checkbox_auto_mode.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        self.checkbox_auto_mode.toggled.connect(self.on_mode_changed)
        self.mode_group.addButton(self.checkbox_auto_mode)
        detection_layout.addWidget(self.checkbox_auto_mode)

        self.help_auto_label = QLabel(
            "Le système détecte seul les en-têtes et les données de chaque "
            "fichier. Aucune saisie nécessaire."
        )
        self.help_auto_label.setWordWrap(True)
        detection_layout.addWidget(self.help_auto_label)

        # === MODE 1: FICHIER DE RÉFÉRENCE (Semi-automatique) ===
        self.checkbox_reference_mode = QRadioButton("✨ Utiliser fichier de référence (recommandé)")
        self.checkbox_reference_mode.setChecked(True)  # SÉLECTIONNÉ PAR DÉFAUT
        self.checkbox_reference_mode.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        self.checkbox_reference_mode.toggled.connect(self.on_mode_changed)
        self.mode_group.addButton(self.checkbox_reference_mode)
        detection_layout.addWidget(self.checkbox_reference_mode)

        # Label explicatif mode référence
        self.help_ref_label = QLabel(
            "Spécifiez l'en-tête du premier fichier, "
            "le système trouvera automatiquement les en-têtes similaires dans les autres"
        )
        self.help_ref_label.setWordWrap(True)
        detection_layout.addWidget(self.help_ref_label)

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

        self.ref_info_label = QLabel("💡 Le premier fichier sélectionné servira de référence")
        self.ref_info_label.setWordWrap(True)

        ref_layout.addWidget(ref_label1, 0, 0)
        ref_layout.addWidget(self.spinbox_ref_header, 0, 1)
        ref_layout.addWidget(ref_label2, 1, 0)
        ref_layout.addWidget(self.spinbox_ref_header_lines, 1, 1)
        ref_layout.addWidget(self.ref_info_label, 2, 0, 1, 2)

        self.reference_group.setLayout(ref_layout)
        detection_layout.addWidget(self.reference_group)

        # === MODE 2: CONFIGURATION MANUELLE COMPLÈTE ===
        self.checkbox_manual_mode = QRadioButton("⚙️ Configuration manuelle pour tous les fichiers")
        self.checkbox_manual_mode.setChecked(False)
        self.checkbox_manual_mode.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        self.checkbox_manual_mode.toggled.connect(self.on_mode_changed)
        self.mode_group.addButton(self.checkbox_manual_mode)
        detection_layout.addWidget(self.checkbox_manual_mode)

        # Label explicatif mode manuel
        self.help_manual_label = QLabel("Tous les fichiers utiliseront exactement la même configuration")
        self.help_manual_label.setWordWrap(True)
        detection_layout.addWidget(self.help_manual_label)

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
        self.manual_warning_label = QLabel("⚠️ Assurez-vous que tous les fichiers ont exactement la même structure")
        self.manual_warning_label.setWordWrap(True)

        manual_layout.addWidget(manual_label1, 0, 0)
        manual_layout.addWidget(self.spinbox_header_start, 0, 1)
        manual_layout.addWidget(manual_label2, 1, 0)
        manual_layout.addWidget(self.spinbox_header_rows, 1, 1)
        manual_layout.addWidget(self.manual_warning_label, 2, 0, 1, 2)

        self.manual_group.setLayout(manual_layout)
        detection_layout.addWidget(self.manual_group)

        detection_group.setLayout(detection_layout)
        layout.addWidget(detection_group)

        # === SECTION 2: OPTIONS DE COMPILATION ===
        compilation_group = QGroupBox("⚙️ Options de compilation")
        compilation_group.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        compilation_layout = QGridLayout()

        # Colonne nom de fichier source
        label_filename_option = QLabel("📋 Colonne nom de fichier source:")
        self.combo_filename_option = QComboBox()
        self.combo_filename_option.addItem("Ne pas ajouter", FilenameOption.NONE)
        self.combo_filename_option.addItem("Avec extension (.xlsx)", FilenameOption.WITH_EXTENSION)
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

        compilation_layout.addWidget(label_filename_option, 0, 0)
        compilation_layout.addWidget(self.combo_filename_option, 0, 1)
        compilation_layout.addWidget(label_output, 1, 0)
        compilation_layout.addWidget(self.lineedit_output, 1, 1)
        compilation_layout.addWidget(label_format, 2, 0)
        compilation_layout.addWidget(self.combo_format, 2, 1)
        compilation_layout.addWidget(self.checkbox_repeat_headers, 3, 0, 1, 2)
        compilation_layout.addWidget(self.checkbox_remove_empty, 4, 0, 1, 2)

        compilation_group.setLayout(compilation_layout)
        layout.addWidget(compilation_group)

        # === SECTION 2bis: ROBUSTESSE DES TABLEAUX ===
        robustness_group = QGroupBox("🧩 Robustesse des tableaux")
        robustness_group.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        robustness_layout = QVBoxLayout()

        # Dé-fusion des cellules fusionnées
        self.checkbox_unmerge_cells = QCheckBox("Propager les cellules fusionnées")
        self.checkbox_unmerge_cells.setChecked(True)
        self.checkbox_unmerge_cells.setToolTip(
            "Recopie la valeur d'une cellule fusionnée sur toute sa plage "
            "(ex. un département fusionné verticalement est répété sur chaque ligne)."
        )
        self.checkbox_unmerge_cells.toggled.connect(self.options_changed.emit)

        # Aplatissement des en-têtes multi-lignes
        self.checkbox_flatten_headers = QCheckBox("Fusionner les en-têtes multi-lignes")
        self.checkbox_flatten_headers.setChecked(True)
        self.checkbox_flatten_headers.setToolTip(
            "Fusionne un en-tête réparti sur plusieurs lignes en un seul libellé "
            "par colonne (ex. « NOMBRE DE CAS - 1er cycle »)."
        )
        self.checkbox_flatten_headers.toggled.connect(self.options_changed.emit)

        # Alignement des colonnes par libellé
        self.checkbox_align_columns = QCheckBox("Aligner les colonnes par libellé")
        self.checkbox_align_columns.setChecked(True)
        self.checkbox_align_columns.setToolTip(
            "Aligne les colonnes des fichiers sur leur libellé (et non leur "
            "position) : évite d'empiler une colonne sous une autre quand l'ordre "
            "diffère. Colonnes manquantes laissées vides, jamais inventées."
        )
        self.checkbox_align_columns.toggled.connect(self.options_changed.emit)

        # Lignes de sous-total / total
        self.checkbox_drop_subtotals = QCheckBox("Exclure les lignes de total / sous-total")
        self.checkbox_drop_subtotals.setChecked(True)
        self.checkbox_drop_subtotals.setToolTip(
            "Retire les lignes d'agrégat (ENSEMBLE, TOTAL…) pour éviter le double "
            "comptage. Décochez pour les conserver."
        )
        self.checkbox_drop_subtotals.toggled.connect(self.on_drop_subtotals_toggled)

        self.checkbox_mark_subtotals = QCheckBox(
            "    ↳ Si conservées, ajouter une colonne « Type de ligne »")
        self.checkbox_mark_subtotals.setChecked(True)
        self.checkbox_mark_subtotals.setToolTip(
            "Quand les totaux sont conservés, ajoute une colonne marquant chaque "
            "ligne (détail / sous-total / total) pour un filtrage propre."
        )
        self.checkbox_mark_subtotals.setEnabled(False)  # actif seulement si on conserve
        self.checkbox_mark_subtotals.toggled.connect(self.options_changed.emit)

        robustness_layout.addWidget(self.checkbox_unmerge_cells)
        robustness_layout.addWidget(self.checkbox_flatten_headers)
        robustness_layout.addWidget(self.checkbox_align_columns)
        robustness_layout.addWidget(self.checkbox_drop_subtotals)
        robustness_layout.addWidget(self.checkbox_mark_subtotals)

        robustness_group.setLayout(robustness_layout)
        layout.addWidget(robustness_group)

        # === SECTION 3: OPTIONS AVANCÉES ===
        advanced_group = QGroupBox("🔧 Options avancées")
        advanced_group.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        advanced_layout = QGridLayout()

        # Supprimer doublons
        self.checkbox_remove_duplicates = QCheckBox("Supprimer les doublons")
        self.checkbox_remove_duplicates.setChecked(False)
        self.checkbox_remove_duplicates.setToolTip("Élimine les lignes identiques (compare toutes les colonnes)")
        self.checkbox_remove_duplicates.toggled.connect(self.options_changed.emit)

        # Trier les données
        self.checkbox_sort_data = QCheckBox("Trier les données")
        self.checkbox_sort_data.setChecked(False)
        self.checkbox_sort_data.setToolTip("Trie les données par la colonne spécifiée")
        self.checkbox_sort_data.toggled.connect(self.on_sort_toggled)

        label_sort_column = QLabel("Colonne de tri:")
        self.lineedit_sort_column = QLineEdit("A")
        self.lineedit_sort_column.setPlaceholderText("A, B, C ou 1, 2, 3")
        self.lineedit_sort_column.setMaximumWidth(100)
        self.lineedit_sort_column.setEnabled(False)
        self.lineedit_sort_column.textChanged.connect(self.options_changed.emit)

        # Format de date
        label_date_format = QLabel("📅 Format de date:")
        self.combo_date_format = QComboBox()
        self.combo_date_format.addItem("Français (JJ/MM/AAAA)", "FRENCH")
        self.combo_date_format.addItem("Standard (AAAA-MM-JJ)", "STANDARD")
        self.combo_date_format.addItem("Américain (MM/JJ/AAAA)", "US")
        self.combo_date_format.addItem("Date et heure (JJ/MM/AAAA HH:MM:SS)", "DATETIME_FRENCH")
        self.combo_date_format.setCurrentIndex(0)  # Français par défaut
        self.combo_date_format.currentIndexChanged.connect(self.options_changed.emit)

        advanced_layout.addWidget(self.checkbox_remove_duplicates, 0, 0, 1, 2)
        advanced_layout.addWidget(self.checkbox_sort_data, 1, 0, 1, 2)
        advanced_layout.addWidget(label_sort_column, 2, 0)
        advanced_layout.addWidget(self.lineedit_sort_column, 2, 1)
        advanced_layout.addWidget(label_date_format, 3, 0)
        advanced_layout.addWidget(self.combo_date_format, 3, 1)

        advanced_group.setLayout(advanced_layout)
        layout.addWidget(advanced_group)

        self.apply_theme()

    def apply_theme(self):
        """(Ré)applique les styles inline dépendant du thème actif."""
        # Cases « mode » : libellé en accent
        accent_checkbox = f"color: {T.ACCENT};"
        self.checkbox_auto_mode.setStyleSheet(accent_checkbox)
        self.checkbox_reference_mode.setStyleSheet(accent_checkbox)

        # Libellés d'aide discrets
        help_style = (
            f"color: {T.TEXT_MUTED}; font-size: 8pt; "
            f"font-style: italic; padding-left: 25px;"
        )
        self.help_auto_label.setStyleSheet(help_style)
        self.help_ref_label.setStyleSheet(help_style)
        self.help_manual_label.setStyleSheet(help_style)

        # Astuce mode référence (teinte accent)
        self.ref_info_label.setStyleSheet(
            f"color: {T.ACCENT}; font-size: 8pt; padding: 5px;")

        # Avertissement mode manuel (teinte alerte)
        self.manual_warning_label.setStyleSheet(
            f"color: {T.WARNING}; font-size: 8pt; padding: 5px;")

    def on_mode_changed(self, checked):
        """Met à jour l'affichage quand le mode de détection change.

        L'exclusivité des 3 modes est assurée nativement par le QButtonGroup
        (un seul radio actif). On se contente donc d'ajuster la visibilité des
        panneaux de configuration et d'émettre le signal. ``toggled`` est émis
        deux fois lors d'un changement (l'ancien passe à False, le nouveau à
        True) ; on n'agit que sur la transition vers True pour ne pas émettre
        ni recalculer deux fois.
        """
        if not checked:
            return
        self.reference_group.setVisible(self.checkbox_reference_mode.isChecked())
        self.manual_group.setVisible(self.checkbox_manual_mode.isChecked())
        self.options_changed.emit()

    def on_sort_toggled(self, checked):
        """Active/désactive le champ de colonne de tri"""
        self.lineedit_sort_column.setEnabled(checked)
        self.options_changed.emit()

    def on_drop_subtotals_toggled(self, checked):
        """La case « marquer » n'a de sens que si on CONSERVE les totaux.

        Quand on exclut (case cochée), marquer est sans objet -> désactivé.
        """
        self.checkbox_mark_subtotals.setEnabled(not checked)
        self.options_changed.emit()

    def get_compilation_options(self) -> CompilationOptions:
        """
        Construit les options de compilation selon le mode choisi

        Returns:
            CompilationOptions configuré
        """
        # Récupérer les options communes
        filename_option = self.combo_filename_option.currentData()

        # Colonne de tri : convertir la saisie en index 0-based. Une saisie
        # invalide (ex. « 1A ») renvoie None ; on retombe alors sur 0 ici, mais
        # main_window valide AVANT de compiler (via parse_sort_column) et refuse
        # de lancer sur une saisie invalide plutôt que de trier en silence.
        sort_column_index = 0
        if self.checkbox_sort_data.isChecked():
            parsed = self.parse_sort_column(self.lineedit_sort_column.text())
            if parsed is not None:
                sort_column_index = parsed

        # Récupérer le format de date
        date_format_str = self.combo_date_format.currentData()
        from core.compilation.compilation_models import DateFormat
        date_format_map = {
            "FRENCH": DateFormat.FRENCH,
            "STANDARD": DateFormat.ISO,
            "US": DateFormat.AMERICAN,
            "DATETIME_FRENCH": DateFormat.DATETIME_FRENCH
        }
        date_format = date_format_map.get(date_format_str, DateFormat.FRENCH)

        # Options communes à tous les modes
        common = dict(
            filename_option=filename_option,
            repeat_headers=self.checkbox_repeat_headers.isChecked(),
            remove_empty_rows=self.checkbox_remove_empty.isChecked(),
            remove_duplicates=self.checkbox_remove_duplicates.isChecked(),
            sort_data=self.checkbox_sort_data.isChecked(),
            sort_column=sort_column_index,
            date_format=date_format,
            # Robustesse des tableaux (v3.2)
            unmerge_cells=self.checkbox_unmerge_cells.isChecked(),
            flatten_multiindex_headers=self.checkbox_flatten_headers.isChecked(),
            align_columns_by_label=self.checkbox_align_columns.isChecked(),
            drop_subtotal_rows=self.checkbox_drop_subtotals.isChecked(),
            mark_subtotal_rows=self.checkbox_mark_subtotals.isChecked(),
        )

        if self.checkbox_auto_mode.isChecked():
            # MODE 0: DÉTECTION AUTOMATIQUE (hybride, sans saisie)
            return CompilationOptions(
                auto_detect_structure=True,
                use_reference_mode=False,
                **common,
            )
        elif self.checkbox_manual_mode.isChecked():
            # MODE 2: CONFIGURATION MANUELLE COMPLÈTE
            return CompilationOptions(
                auto_detect_structure=False,
                use_reference_mode=False,
                manual_header_start_row=self.spinbox_header_start.value(),
                manual_header_rows=self.spinbox_header_rows.value(),
                **common,
            )
        else:
            # MODE 1: FICHIER DE RÉFÉRENCE (semi-automatique, défaut)
            return CompilationOptions(
                auto_detect_structure=False,
                use_reference_mode=True,
                reference_header_row=self.spinbox_ref_header.value(),
                reference_header_lines=self.spinbox_ref_header_lines.value(),
                **common,
            )

    @staticmethod
    def parse_sort_column(text: str):
        """Convertit une saisie de colonne de tri en index 0-based, ou None.

        Accepte la notation tableur (A, B, …, Z, AA, …) OU un numéro (1, 2, …).
        Une saisie mixte ou vide (« 1A », « A1 », «  », « ! ») renvoie None :
        l'appelant refuse alors de trier plutôt que de retomber en silence sur
        la colonne A (tri erroné invisible pour l'utilisateur).
        """
        s = (text or "").strip().upper()
        if not s:
            return None
        if s.isalpha():
            idx = 0
            for ch in s:
                idx = idx * 26 + (ord(ch) - ord('A') + 1)
            return idx - 1
        if s.isdigit():
            n = int(s)
            return n - 1 if n >= 1 else None
        return None

    def sort_column_is_valid(self) -> bool:
        """Vrai si le tri est désactivé, ou activé avec une colonne valide."""
        if not self.checkbox_sort_data.isChecked():
            return True
        return self.parse_sort_column(self.lineedit_sort_column.text()) is not None

    def get_output_file(self) -> str:
        """Retourne le nom du fichier de sortie"""
        return self.lineedit_output.text().strip()

    def get_output_format(self) -> OutputFormat:
        """Retourne le format de sortie"""
        return self.combo_format.currentData()

    def is_auto_mode_enabled(self) -> bool:
        """Vérifie si le mode détection automatique est activé"""
        return self.checkbox_auto_mode.isChecked()

    def is_reference_mode_enabled(self) -> bool:
        """Vérifie si le mode référence est activé"""
        return self.checkbox_reference_mode.isChecked()

    def is_manual_mode_enabled(self) -> bool:
        """Vérifie si le mode manuel est activé"""
        return self.checkbox_manual_mode.isChecked()
