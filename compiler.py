import sys
import os
import pandas as pd
from PyQt6.QtWidgets import *
from PyQt6.QtCore import Qt, QThread, pyqtSignal
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
    progress = pyqtSignal(int)
    finished = pyqtSignal(tuple)
    error = pyqtSignal(str)

    def __init__(self, files, directory, header_rows, parent=None):
        super().__init__(parent)
        self.files = files
        self.directory = directory
        self.header_rows = header_rows
        #self.merge_headers = merge_headers

    def run(self):
        try:
            combined_data = []
            headers = None
            merged_cells = []

            for i, file in enumerate(self.files):
                try:
                    file_path = os.path.join(self.directory, file)
                    
                    # Lecture avec openpyxl pour préserver la structure des en-têtes
                    wb = openpyxl.load_workbook(file_path)
                    ws = wb.active
                    
                    # Capture des informations sur les cellules fusionnées dans les en-têtes
                    if headers is None:
                        headers = []
                        for row in range(1, self.header_rows + 1):
                            header_row = []
                            for cell in ws[row]:
                                header_row.append(cell.value)
                            headers.append(header_row)
                        
                        # Capture des cellules fusionnées dans les en-têtes
                        for merged_range in ws.merged_cells.ranges:
                            if merged_range.min_row <= self.header_rows:
                                merged_cells.append(merged_range)
                    
                    # Lecture des données
                    data = []
                    for row in ws.iter_rows(min_row=self.header_rows + 1):
                        row_data = [cell.value for cell in row]
                        row_data.append(file)  # Ajout du nom du fichier à la fin
                        data.append(row_data)
                    
                    combined_data.extend(data)
                    self.progress.emit(i + 1)
                    
                except Exception as e:
                    logging.error(f"Erreur lors du traitement du fichier {file}: {str(e)}")
                    self.error.emit(f"Erreur avec le fichier {file}: {str(e)}")
                    continue

            if combined_data and headers:
                # Ajout du nom de la colonne pour les noms de fichiers
                headers[-1].append("Fichier source")
                self.finished.emit((headers, combined_data, merged_cells))
            else:
                self.error.emit("Aucun fichier n'a pu être traité correctement.")
                
        except Exception as e:
            logging.error(f"Erreur générale: {str(e)}")
            self.error.emit(f"Erreur générale: {str(e)}")

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
        self.setGeometry(100, 100, 1200, 800)
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
        
        layout.addLayout(selection_layout)
        layout.addWidget(self.list_files)
        group.setLayout(layout)
        return group

    def create_options_group(self):
        group = QGroupBox("Options de compilation")
        layout = QVBoxLayout()

        # En-têtes
        header_layout = QHBoxLayout()
        self.label_header = QLabel("Lignes d'en-tête :")
        self.spinbox_header = QSpinBox()
        self.spinbox_header.setMinimum(1)
        self.spinbox_header.setMaximum(15)
        self.spinbox_header.setValue(1)
        header_layout.addWidget(self.label_header)
        header_layout.addWidget(self.spinbox_header)
        header_layout.addStretch()

        # Options des en-têtes
        self.checkbox_repeat_header = QCheckBox("Répéter les en-têtes pour chaque fichier")
        self.checkbox_merge_headers = QCheckBox("Fusionner les en-têtes multi-niveaux")
        
        # Nom du fichier de sortie
        output_layout = QHBoxLayout()
        self.label_output_name = QLabel("Nom du fichier de sortie :")
        self.lineedit_output_name = QLineEdit("compilation.xlsx")
        output_layout.addWidget(self.label_output_name)
        output_layout.addWidget(self.lineedit_output_name)

        layout.addLayout(header_layout)
        layout.addWidget(self.checkbox_repeat_header)
        layout.addWidget(self.checkbox_merge_headers)
        layout.addLayout(output_layout)
        group.setLayout(layout)
        return group

    def create_advanced_options_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        # Groupe traitement des données
        data_group = QGroupBox("Traitement des données")
        data_layout = QVBoxLayout()

        self.checkbox_remove_duplicates = QCheckBox("Supprimer les doublons")
        self.checkbox_remove_duplicates.setChecked(True)

        sort_layout = QVBoxLayout()
        self.checkbox_sort_data = QCheckBox("Trier les données")
        
        sort_options = QHBoxLayout()
        self.label_sort_column = QLabel("Colonne de tri :")
        self.lineedit_sort_column = QLineEdit("1")
        self.lineedit_sort_column.setEnabled(False)
        sort_options.addWidget(self.label_sort_column)
        sort_options.addWidget(self.lineedit_sort_column)
        sort_options.addStretch()

        sort_layout.addWidget(self.checkbox_sort_data)
        sort_layout.addLayout(sort_options)

        data_layout.addWidget(self.checkbox_remove_duplicates)
        data_layout.addLayout(sort_layout)
        data_group.setLayout(data_layout)

        # Groupe formatage
        format_group = QGroupBox("Formatage")
        format_layout = QVBoxLayout()

        self.checkbox_auto_width = QCheckBox("Ajuster automatiquement la largeur des colonnes")
        self.checkbox_auto_width.setChecked(True)
        
        self.checkbox_freeze_header = QCheckBox("Figer la première ligne")
        self.checkbox_freeze_header.setChecked(True)

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
        <h2>Guide d'utilisation</h2>
        
        <h3>1. Sélection des fichiers</h3>
        - Cliquez sur "Choisir un répertoire" pour sélectionner le dossier contenant vos fichiers Excel
        - Sélectionnez les fichiers à compiler dans la liste
        - Utilisez la case "Sélectionner tous les fichiers" pour tout sélectionner
        
        <h3>2. Options de base</h3>
        - Spécifiez le nombre de lignes d'en-tête
        - Choisissez si vous voulez répéter les en-têtes
        - Définissez le nom du fichier de sortie
        
        <h3>3. Options avancées</h3>
        - Suppression des doublons
        - Tri des données
        - Formatage automatique
        
        <h3>4. Résolution des problèmes courants</h3>
        - Assurez-vous que tous les fichiers ont la même structure
        - Vérifiez que les en-têtes sont cohérents
        - Fermez les fichiers Excel avant la compilation
        """

        help_label = QLabel(help_text)
        help_label.setWordWrap(True)
        help_label.setTextFormat(Qt.TextFormat.RichText)
        
        scroll = QScrollArea()
        scroll.setWidget(help_label)
        scroll.setWidgetResizable(True)
        
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
        
        self.compilation_worker = CompilationWorker(
            files=selected_files,
            directory=self.directory,
            header_rows=self.spinbox_header.value(),
            #merge_headers=self.checkbox_merge_headers.isChecked(),
            parent=self
        )
        self.compilation_worker.progress.connect(self.update_progress)
        self.compilation_worker.finished.connect(self.save_compilation)
        self.compilation_worker.error.connect(self.show_error)
        self.compilation_worker.start()

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
                headers, data, merged_cells = compilation_data
                output_name = self.lineedit_output_name.text()
                if not output_name.lower().endswith(('.xlsx', '.xls')):
                    output_name += '.xlsx'
                
                output_path = os.path.join(self.directory, output_name)
                
                # Création du nouveau workbook
                wb = openpyxl.Workbook()
                ws = wb.active
                
                # Écriture des en-têtes
                for row_idx, header_row in enumerate(headers, 1):
                    for col_idx, value in enumerate(header_row, 1):
                        cell = ws.cell(row=row_idx, column=col_idx, value=value)
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill(start_color='2E7D32', end_color='2E7D32', fill_type='solid')
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # Application des fusions de cellules si l'option est activée
                if self.checkbox_merge_headers.isChecked():
                    for merged_range in merged_cells:
                        ws.merge_cells(
                            start_row=merged_range.min_row,
                            start_column=merged_range.min_col,
                            end_row=merged_range.max_row,
                            end_column=merged_range.max_col
                        )
                
                # Écriture des données
                for row_idx, row_data in enumerate(data, len(headers) + 1):
                    for col_idx, value in enumerate(row_data, 1):
                        cell = ws.cell(row=row_idx, column=col_idx, value=value)
                        cell.border = Border(
                            left=Side(border_style='thin'),
                            right=Side(border_style='thin'),
                            top=Side(border_style='thin'),
                            bottom=Side(border_style='thin')
                        )
                
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
                    ws.freeze_panes = ws.cell(row=len(headers) + 1, column=1)
                
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

if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        window = ModernExcelCompilerApp()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        logging.critical(f"Erreur critique de l'application: {str(e)}")
        raise