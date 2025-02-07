import sys
import os
import pandas as pd
from PyQt6.QtWidgets import *
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QIcon, QFont, QPalette, QColor
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import openpyxl
import logging
from datetime import datetime

# Configuration du logging
logging.basicConfig(
    filename=f'excel_compiler_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class CompilationWorker(QThread):
    def __init__(self, files, directory, header_start_row, header_rows, add_filename=True, sort_data=False, sort_column=0, repeat_headers=False, remove_empty_rows=False, parent=None):
        super().__init__(parent)
        self.files = files
        self.directory = directory
        self.header_start_row = header_start_row
        self.header_rows = header_rows
        self.add_filename = add_filename
        self.sort_data = sort_data
        self.sort_column = sort_column
        self.repeat_headers = repeat_headers
        self.remove_empty_rows = remove_empty_rows

    progress = pyqtSignal(int)
    error = pyqtSignal(str)      
    finished = pyqtSignal(tuple) 

    def run(self):
        combined_data = []
        headers = None
        merged_cells = []
        preliminary_info = []

        for i, file in enumerate(self.files):
            try:
                file_path = os.path.join(self.directory, file)
                
                # Pour le premier fichier, on a besoin des merged_cells donc on ne met pas read_only
                if i == 0:
                    wb = openpyxl.load_workbook(file_path, data_only=True)
                else:
                    wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
                    
                ws = wb.active
                
                # Capture des informations préliminaires du premier fichier
                if i == 0 and self.header_start_row > 1:
                    for row in range(1, self.header_start_row):
                        row_data = []
                        for cell in ws[row]:
                            row_data.append(cell.value)
                        preliminary_info.append(row_data)
                
                # Capture des en-têtes du premier fichier
                if headers is None:
                    headers = []
                    for row in range(self.header_start_row, self.header_start_row + self.header_rows):
                        header_row = []
                        for cell in ws[row]:
                            header_row.append(cell.value)
                        headers.append(header_row)
                    
                    # Capture des merged_cells seulement pour le premier fichier
                    if hasattr(ws, 'merged_cells'):
                        for merged_range in ws.merged_cells.ranges:
                            if merged_range.min_row <= (self.header_start_row + self.header_rows):
                                merged_cells.append(merged_range)
                
                # Ajout des données
                data = []
                
                # Si répétition des en-têtes est activée et ce n'est pas le premier fichier
                if self.repeat_headers and combined_data:
                    data.append([None] * len(headers[-1]))
                    for header_row in headers:
                        row_data = header_row.copy()
                        if self.add_filename:
                            row_data.append(None)
                        data.append(row_data)
                
                # Ajout des données du fichier
                for row in ws.iter_rows(min_row=self.header_start_row + self.header_rows):
                    row_data = [cell.value for cell in row]
                    if self.add_filename:
                        row_data.append(file)
                    
                    # Vérifier si la ligne n'est pas vide avant de l'ajouter
                    if not self.remove_empty_rows or not all(cell is None or str(cell).strip() == "" for cell in row_data[:-1]):  # Exclure la colonne du nom de fichier
                        data.append(row_data)
                
                combined_data.extend(data)
                self.progress.emit(i + 1)
                
                # Fermer le workbook pour libérer la mémoire
                wb.close()
                
            except Exception as e:
                logging.error(f"Erreur lors du traitement du fichier {file}: {str(e)}")
                self.error.emit(f"Erreur avec le fichier {file}: {str(e)}")
                continue

        if combined_data and headers:
            if self.add_filename:
                headers[-1].append("Fichier source")
            
            # Tri des données si l'option est activée
            if self.sort_data and self.sort_column < len(headers[-1]):
                try:
                    sort_idx = self.sort_column - 1
                    if self.repeat_headers:
                        sections = []
                        current_section = []
                        for row in combined_data:
                            if all(cell is None for cell in row):
                                if current_section:
                                    sections.append(current_section)
                                current_section = []
                            else:
                                current_section.append(row)
                        if current_section:
                            sections.append(current_section)
                        
                        for section in sections:
                            section.sort(key=lambda x: (x[sort_idx] is None, x[sort_idx]))
                        
                        combined_data = []
                        for i, section in enumerate(sections):
                            if i > 0:
                                combined_data.append([None] * len(headers[-1]))
                            combined_data.extend(section)
                    else:
                        combined_data.sort(key=lambda x: (x[sort_idx] is None, x[sort_idx]))
                except Exception as e:
                    logging.warning(f"Erreur lors du tri : {str(e)}")
            
            self.finished.emit((preliminary_info, headers, combined_data, merged_cells))


class ModernExcelCompilerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setup_ui()
        self.setup_variables()
        self.create_main_layout()
        self.connect_signals()
        logging.info("Application démarrée")

    def setup_ui(self):
        self.setWindowTitle("Compilateur Excel Professionnel")
        self.setGeometry(50, 50, 1200, 700)
        self.setWindowIcon(QIcon("excel.png"))
        self.setMinimumSize(800, 600)
        self.apply_stylesheet()

    def setup_variables(self):
        self.directory = ""
        self.files = []
        self.compilation_worker = None

    def apply_stylesheet(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #2e7d32;
                border-radius: 8px;
                margin-top: 12px;
                padding: 15px;
                background-color: white;
            }
            QGroupBox::title {
                color: #2e7d32;
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QPushButton {
                background-color: #2e7d32;
                color: white;
                padding: 8px 15px;
                border-radius: 4px;
                font-weight: bold;
                min-height: 30px;
            }
            QPushButton:hover {
                background-color: #1b5e20;
            }
            QPushButton:disabled {
                background-color: #c8e6c9;
            }
            QTabWidget::pane {
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                background-color: white;
            }
            QTabBar::tab {
                background-color: #f5f5f5;
                border: 1px solid #e0e0e0;
                padding: 8px 15px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background-color: #2e7d32;
                color: white;
            }
            QTabBar::tab:hover {
                background-color: #4caf50;
                color: white;
            }
        """)

    def create_main_layout(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        
        # Création des onglets
        self.tabs = QTabWidget()
        self.create_compilation_tab()
        self.create_advanced_options_tab()
        self.create_help_tab()
        
        main_layout.addWidget(self.tabs)

    def create_compilation_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        # Groupe sélection des fichiers
        files_group = self.create_files_group()
        layout.addWidget(files_group)

        # Groupe options de compilation
        options_group = self.create_options_group()
        layout.addWidget(options_group)

        # Barre de progression
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Bouton de compilation
        compile_layout = QHBoxLayout()
        self.button_compile = QPushButton("Lancer la compilation")
        self.button_compile.setFont(QFont("Segoe UI", 12))
        self.button_compile.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay))
        compile_layout.addStretch()
        compile_layout.addWidget(self.button_compile)
        compile_layout.addStretch()
        layout.addLayout(compile_layout)

        # Status label
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status_label)

        tab.setLayout(layout)
        self.tabs.addTab(tab, "Compilation")

    def create_files_group(self):
        group = QGroupBox("Sélection des fichiers")
        layout = QVBoxLayout()

        # Date et heure
        self.datetime_label = QLabel()
        self.update_datetime()
        
        # Timer pour mettre à jour la date et l'heure
        self.datetime_timer = QTimer()
        self.datetime_timer.timeout.connect(self.update_datetime)
        self.datetime_timer.start(1000)  # Mise à jour toutes les secondes

        # Sélection du répertoire
        dir_layout = QHBoxLayout()
        self.label_directory = QLabel("Aucun répertoire sélectionné")
        self.button_choose_directory = QPushButton("Choisir un répertoire")
        self.button_choose_directory.setFont(QFont("Segoe UI", 12))
        self.button_choose_directory.setIcon(QIcon("folder.png"))
        dir_layout.addWidget(self.label_directory)
        dir_layout.addWidget(self.button_choose_directory)
        layout.addLayout(dir_layout)

        # Liste des fichiers
        self.list_files = QListWidget()
        self.list_files.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        
        # Options de sélection
        
        selection_layout = QHBoxLayout()
        self.checkbox_all_files = QCheckBox("Sélectionner tous les fichiers")
        self.label_file_count = QLabel("0 fichier(s) sélectionné(s)")
        selection_layout.addWidget(self.checkbox_all_files)
        selection_layout.addStretch()
        selection_layout.addWidget(self.label_file_count)
        selection_layout.addStretch()
        selection_layout.addWidget(self.datetime_label)
        
        layout.addLayout(selection_layout)
        layout.addWidget(self.list_files)
        group.setLayout(layout)
        return group

    def create_options_group(self):
        group = QGroupBox("Options de compilation")
        grid_layout = QGridLayout()
        
        # Première colonne : Spinboxes et labels
        # Ligne 0
        self.label_header_start = QLabel("Ligne de début des en-têtes :")
        self.spinbox_header_start = QSpinBox()
        self.spinbox_header_start.setMinimum(1)
        self.spinbox_header_start.setMaximum(20)
        self.spinbox_header_start.setValue(1)
        header_start_container = QHBoxLayout()
        header_start_container.addWidget(self.label_header_start)
        header_start_container.addWidget(self.spinbox_header_start)
        header_start_container.addStretch()
        
        # Ligne 1
        self.label_header = QLabel("Nombre de lignes d'en-tête :")
        self.spinbox_header = QSpinBox()
        self.spinbox_header.setMinimum(1)
        self.spinbox_header.setMaximum(15)
        self.spinbox_header.setValue(1)
        header_container = QHBoxLayout()
        header_container.addWidget(self.label_header)
        header_container.addWidget(self.spinbox_header)
        header_container.addStretch()
        
        # Deuxième colonne : Checkboxes pour répéter et fusionner
        self.checkbox_repeat_header = QCheckBox("Répéter les en-têtes pour chaque fichier")
        self.checkbox_merge_headers = QCheckBox("Fusionner les en-têtes multi-niveaux")
        
        # Troisième colonne : Nom de fichier et sortie
        self.checkbox_add_filename = QCheckBox("Ajouter les noms des fichiers sources")
        self.checkbox_add_filename.setChecked(False)
        
        self.label_output_name = QLabel("Nom du fichier de sortie :")
        self.lineedit_output_name = QLineEdit("compilation.xlsx")
        output_container = QHBoxLayout()
        output_container.addWidget(self.label_output_name)
        output_container.addWidget(self.lineedit_output_name)
        
        # Ajout des widgets au QGridLayout
        # Première colonne (0)
        grid_layout.addLayout(header_start_container, 0, 0)
        grid_layout.addLayout(header_container, 1, 0)
        
        # Deuxième colonne (1)
        grid_layout.addWidget(self.checkbox_repeat_header, 0, 1)
        grid_layout.addWidget(self.checkbox_merge_headers, 1, 1)
        
        # Troisième colonne (2)
        grid_layout.addWidget(self.checkbox_add_filename, 0, 2)
        grid_layout.addLayout(output_container, 1, 2)
        
        # Définir l'espacement et les marges
        grid_layout.setSpacing(20)  # Espacement entre les éléments
        grid_layout.setContentsMargins(20, 20, 20, 20)  # Marges autour de la grille
        
        # Définir les colonnes pour qu'elles aient la même largeur
        grid_layout.setColumnStretch(0, 1)
        grid_layout.setColumnStretch(1, 1)
        grid_layout.setColumnStretch(2, 1)
        
        # Alignement vertical des éléments
        grid_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        
        group.setLayout(grid_layout)
        return group

    def update_datetime(self):
        current_datetime = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        self.datetime_label.setText(f"Date et heure : {current_datetime}")


    def create_advanced_options_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        # Groupe traitement des données
        data_group = QGroupBox("Traitement des données")
        data_layout = QVBoxLayout()

        self.checkbox_remove_duplicates = QCheckBox("Supprimer les doublons")
        self.checkbox_remove_duplicates.setChecked(True)
        
        self.checkbox_remove_empty_rows = QCheckBox("Supprimer les lignes entièrement vides")
        self.checkbox_remove_empty_rows.setChecked(False)

        sort_layout = QVBoxLayout()
        self.checkbox_sort_data = QCheckBox("Trier les données")
        
        sort_options = QHBoxLayout()
        self.label_sort_column = QLabel("Colonne de tri (ex: A, B, C) :")
        self.lineedit_sort_column = QLineEdit("A")
        self.lineedit_sort_column.setEnabled(False)
        sort_options.addWidget(self.label_sort_column)
        sort_options.addWidget(self.lineedit_sort_column)
        sort_options.addStretch()

        sort_layout.addWidget(self.checkbox_sort_data)
        sort_layout.addLayout(sort_options)

        data_layout.addWidget(self.checkbox_remove_duplicates)
        data_layout.addWidget(self.checkbox_remove_empty_rows)
        data_layout.addLayout(sort_layout)
        data_group.setLayout(data_layout)

        # Groupe formatage
        format_group = QGroupBox("Formatage")
        format_layout = QVBoxLayout()

        self.checkbox_auto_width = QCheckBox("Ajuster automatiquement la largeur des colonnes")
        self.checkbox_auto_width.setChecked(True)
        
        self.checkbox_freeze_header = QCheckBox("Figer la première ligne")
        self.checkbox_freeze_header.setChecked(False)

        format_layout.addWidget(self.checkbox_auto_width)
        format_layout.addWidget(self.checkbox_freeze_header)
        format_group.setLayout(format_layout)

        layout.addWidget(data_group)
        layout.addWidget(format_group)
        layout.addStretch()
        tab.setLayout(layout)
        self.tabs.addTab(tab, "Options avancées")

    def create_help_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        help_text = """
        <style>
        body {
            font-family: "Segoe UI", sans-serif;
            margin: 0 20px;
            line-height: 1.6;
        }

        h2, h3 {
            color: #2e7d32;
            font-weight: bold;
            margin: 25px 0 15px 0;
        }

        h4 {
            margin: 12px 0;
            padding-left: 20px;
            font-weight: normal;
        }

        ul {
            margin-left: 50px;
            line-height: 1.5;
        }

        li {
            margin: 15px 0;
        }

        .section-content {
            padding-left: 30px;
        }

        b {
            color: #1b5e20;
        }

        /* Améliorations pour l'expérience utilisateur */
        @media (max-width: 768px) {
            body {
                margin: 0 10px;
            }
            
            .section-content {
                padding-left: 15px;
            }
        }

        /* Animation subtile au survol des sections */
        h3:hover {
            transform: translateX(5px);
            transition: transform 0.2s ease;
        }
        </style>

        <h2>Guide d'utilisation du Compilateur Excel</h2>
        
        <h3>  I. Sélection des fichiers</h3>
        <h4>  - Cliquez sur "Choisir un répertoire" pour sélectionner le dossier contenant vos fichiers Excel<h4>
        <h4>  - Sélectionnez les fichiers à compiler dans la liste<h4>
        <h4>  - Utilisez la case "Sélectionner tous les fichiers" pour tout sélectionner/désélectionner<h4>
        <h4>  - Le nombre de fichiers sélectionnés est affiché en temps réel<h4>
        
        <h3>  II. Options de compilation</h3>
        <h4>  - <b>Ligne de début de l'en-tête</b> : Spécifiez à quelle ligne commence l'en-tête dans vos fichiers<h4> 
        <h4>  - <b>Nombre de lignes d'en-tête</b> : Indiquez combien de lignes constituent l'en-tête<h4> 
        <h4>  - <b>Répéter les en-têtes</b> : Réinsère les en-têtes entre chaque fichier dans la compilation<h4> 
        <h4>  - <b>Fusionner les en-têtes multi-niveaux</b> : Conserve la fusion des cellules d'en-tête<h4> 
        <h4>  - <b>Ajouter les noms des fichiers sources</b> : Ajoute une colonne avec les noms des fichiers d'origine<h4> 
        <h4>  - <b>Nom du fichier de sortie</b> : Définissez le nom du fichier compilé (.xlsx sera ajouté automatiquement)<h4> 
        
        <h3>  III. Options avancées</h3>
        
              Traitement des données :
        <h4>  - <b>Supprimer les doublons</b> : Élimine les lignes identiques<h4> 
        <h4>  - <b>Supprimer les lignes vides</b> : Retire les lignes ne contenant aucune donnée<h4> 
        <h4>  - <b>Trier les données</b> : Trie le contenu selon une colonne spécifique<h4> 
        <h4>  - <b>Colonne de tri</b> : Spécifiez la colonne pour le tri (ex: A pour première colonne)<h4> 
        
              Formatage :
        <h4>  - <b>Ajuster la largeur des colonnes</b> : Adapte automatiquement la largeur selon le contenu<h4> 
        <h4>  - <b>Figer la première ligne</b> : Maintient l'en-tête visible lors du défilement<h4> 
        
        <h3>  IV. Informations préliminaires</h3>
        <h4>  - Pour le premier fichier uniquement, les informations situées avant l'en-tête sont préservées<h4> 
        <h4>  - Ces informations sont copiées dans le fichier final<h4> 
        
        <h3>  V. Résolution des problèmes courants</h3>
        <h4>  - Assurez-vous que tous les fichiers ont la même structure d'en-tête<h4>
        <h4>  -  Vérifiez que la ligne de début d'en-tête est correctement définie<h4>
        <h4>  -  Fermez les fichiers Excel avant la compilation<h4>
        <h4>  -  En cas d'erreur, consultez le fichier log généré dans le répertoire de l'application<h4>
        
        <h3>  VI. Bonnes pratiques</h3>
        <h4>  -  Faites une sauvegarde de vos fichiers avant la compilation<h4>
        <h4>  -  Vérifiez le résultat après compilation<h4>
        <h4>  -  Pour les gros fichiers, évitez de sélectionner trop de fichiers à la fois<h4>
        <h4>  -  Utilisez des noms de fichiers explicites pour le fichier de sortie<h4>

        """

        help_label = QLabel(help_text)
        help_label.setWordWrap(True)
        help_label.setTextFormat(Qt.TextFormat.RichText)
        help_label.setStyleSheet("QLabel { line-height: 1.5; }")
        
        scroll = QScrollArea()
        scroll.setWidget(help_label)
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { background-color: white; }")
        
        layout.addWidget(scroll)
        tab.setLayout(layout)
        self.tabs.addTab(tab, "Aide")

    def connect_signals(self):
        self.button_choose_directory.clicked.connect(self.choose_directory)
        self.checkbox_all_files.stateChanged.connect(self.toggle_file_selection)
        self.button_compile.clicked.connect(self.compile_files)
        self.list_files.itemSelectionChanged.connect(self.update_file_count)
        self.checkbox_sort_data.stateChanged.connect(
            lambda: self.lineedit_sort_column.setEnabled(self.checkbox_sort_data.isChecked())
        )

    def choose_directory(self):
        self.directory = QFileDialog.getExistingDirectory(
            self,
            "Choisir un répertoire",
            "",
            QFileDialog.Option.ShowDirsOnly
        )
        if self.directory:
            self.label_directory.setText(f"Répertoire : {self.directory}")
            self.files = [
                f for f in os.listdir(self.directory)
                if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')
            ]
            self.list_files.clear()
            self.list_files.addItems(self.files)
            self.checkbox_all_files.setChecked(True)
            logging.info(f"Répertoire sélectionné : {self.directory}")

    def update_file_count(self):
        count = len(self.list_files.selectedItems())
        self.label_file_count.setText(f"{count} fichier(s) sélectionné(s)")

    def toggle_file_selection(self):
        for i in range(self.list_files.count()):
            self.list_files.item(i).setSelected(self.checkbox_all_files.isChecked())

    def compile_files(self):
            
            if not self._validate_compilation():
                return
            self.button_compile.setEnabled(False)
            self.progress_bar.setVisible(True)
            self.status_label.setText("Compilation en cours...")
            
            selected_files = [item.text() for item in self.list_files.selectedItems()]
            self.progress_bar.setMaximum(len(selected_files))
            
            sort_column = 1
            if self.checkbox_sort_data.isChecked():
                col_letter = self.lineedit_sort_column.text().upper()
                sort_column = self._column_letter_to_number(col_letter)

            self.compilation_worker = CompilationWorker(
                files=selected_files,
                directory=self.directory,
                header_start_row=self.spinbox_header_start.value(),
                header_rows=self.spinbox_header.value(),
                add_filename=self.checkbox_add_filename.isChecked(),
                sort_data=self.checkbox_sort_data.isChecked(),
                sort_column=sort_column,
                repeat_headers=self.checkbox_repeat_header.isChecked(),
                remove_empty_rows=self.checkbox_remove_empty_rows.isChecked(),
                parent=self
            )
            self.compilation_worker.progress.connect(self.update_progress)
            self.compilation_worker.finished.connect(self.save_compilation)
            self.compilation_worker.error.connect(self.show_error)
            self.compilation_worker.start()
    
    def _column_letter_to_number(self, column_letter):
        result = 0
        for i, letter in enumerate(reversed(column_letter.strip())):
            result += (ord(letter) - ord('A') + 1) * (26 ** i)
        return result

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def show_error(self, message):
        QMessageBox.warning(self, "Erreur", message)
        self.button_compile.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.status_label.setText("La compilation a échoué")
        logging.error(message)

    def _validate_compilation(self):
        if not self.directory:
            QMessageBox.warning(self, "Erreur", "Veuillez choisir un répertoire.")
            return False
        
        selected_files = [item.text() for item in self.list_files.selectedItems()]
        if not selected_files:
            QMessageBox.warning(self, "Erreur", "Veuillez sélectionner au moins un fichier.")
            return False
        
        output_name = self.lineedit_output_name.text()
        if not output_name:
            QMessageBox.warning(self, "Erreur", "Veuillez spécifier un nom pour le fichier de sortie.")
            return False
            
        return True

    def save_compilation(self, compilation_data):
            try:
                preliminary_info, headers, data, merged_cells = compilation_data
                output_name = self.lineedit_output_name.text()
                if not output_name.lower().endswith(('.xlsx', '.xls')):
                    output_name += '.xlsx'
                
                output_path = os.path.join(self.directory, output_name)
                
                # Création du nouveau workbook
                wb = openpyxl.Workbook()
                ws = wb.active
                
                # Écriture des informations préliminaires
                current_row = 1
                for row_data in preliminary_info:
                    for col_idx, value in enumerate(row_data, 1):
                        cell = ws.cell(row=current_row, column=col_idx, value=value)
                        cell.font = Font(italic=True)
                    current_row += 1
                
                # Écriture des en-têtes
                header_start = current_row
                for header_row in headers:
                    for col_idx, value in enumerate(header_row, 1):
                        cell = ws.cell(row=current_row, column=col_idx, value=value)
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill(start_color='2E7D32', end_color='2E7D32', fill_type='solid')
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    current_row += 1
                
                # Application des fusions de cellules
                if self.checkbox_merge_headers.isChecked():
                    for merged_range in merged_cells:
                        # Ajuster les numéros de ligne pour tenir compte des informations préliminaires
                        adjusted_range = openpyxl.worksheet.cell_range.CellRange(
                            min_col=merged_range.min_col,
                            min_row=merged_range.min_row - self.spinbox_header_start.value() + header_start,
                            max_col=merged_range.max_col,
                            max_row=merged_range.max_row - self.spinbox_header_start.value() + header_start
                        )
                        ws.merge_cells(range_string=adjusted_range.coord)
                
                # Écriture des données
                for row_data in data:
                    for col_idx, value in enumerate(row_data, 1):
                        cell = ws.cell(row=current_row, column=col_idx, value=value)
                        cell.border = Border(
                            left=Side(border_style='thin'),
                            right=Side(border_style='thin'),
                            top=Side(border_style='thin'),
                            bottom=Side(border_style='thin')
                        )
                    current_row += 1
                
                # Ajustement automatique des colonnes
                if self.checkbox_auto_width.isChecked():
                    for column in ws.columns:
                        max_length = 0
                        column_letter = get_column_letter(column[0].column)
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = (max_length + 2)
                        ws.column_dimensions[column_letter].width = min(adjusted_width, 50)
                
                # Figer les volets
                if self.checkbox_freeze_header.isChecked():
                    ws.freeze_panes = ws.cell(row=len(preliminary_info) + len(headers) + 1, column=1)
                
                wb.save(output_path)
                
                self.status_label.setText("Compilation terminée avec succès!")
                QMessageBox.information(
                    self,
                    "Succès",
                    f"Compilation terminée.\nFichier enregistré : {output_path}"
                )
                logging.info(f"Compilation réussie. Fichier créé : {output_path}")
                
            except Exception as e:
                error_msg = f"Erreur lors de la sauvegarde : {str(e)}"
                QMessageBox.critical(self, "Erreur", error_msg)
                logging.error(error_msg)
            
            finally:
                self.button_compile.setEnabled(True)
                self.progress_bar.setVisible(False)

#export default CompilationWorker

if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        window = ModernExcelCompilerApp()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        logging.critical(f"Erreur critique de l'application: {str(e)}")
        raise