"""
Excel Compiler Application
Version: 2.0
Auteur: GOUNOU N'GOBI Chabi Zimé (Data Manager, Data Analyst)
Améliorations: Mars 2025

Application pour compiler plusieurs fichiers Excel en un seul fichier avec diverses options de formatage.
"""

import sys
import os
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, QTabWidget,
    QGroupBox, QLabel, QSpinBox, QCheckBox, QLineEdit, QPushButton, QListWidget,
    QProgressBar, QMessageBox, QFileDialog, QListWidgetItem, QStyle, QProgressDialog,
    QTableWidget, QTableWidgetItem, QHeaderView, QScrollArea, QDialog, QComboBox,
    QStyledItemDelegate, QStyleOptionButton, QStyle, QGridLayout
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QSize
from PyQt6.QtGui import QIcon, QFont, QPalette, QColor, QBrush
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import openpyxl
import logging
import csv
import unittest
from datetime import datetime
from typing import List, Tuple, Dict, Optional, Any, Union
import traceback

# Constants for styles
COLORS = {
    "PRIMARY": "2e7d32",
    "PRIMARY_DARK": "2e7d32",
    #"PRIMARY_DARK": "1b5e20", 
    "PRIMARY_LIGHT": "4caf50",  
    "ACCENT": "c8e6c9",
    #"BACKGROUND": "f5f5f5",
    "BACKGROUND": "2e7d32",
    "LIGHT_TEXT": "FFFFFF",
    "DARK_TEXT": "212121",
    "BORDER": "e0e0e0",
    "WARNING": "f44336",
    "SUCCESS": "4CAF50",
    "INFO": "2196F3"
}

# Pour openpyxl, ajoutez FF au début pour l'opacité
EXCEL_COLORS = {
    "PRIMARY": "FF2e7d32",
    #"PRIMARY_DARK": "FF1b5e20",
    "PRIMARY_DARK": "FF2e7d32",
    "PRIMARY_LIGHT": "FF4caf50",
    "ACCENT": "FFc8e6c9",
    "BACKGROUND":"FF2e7d32",
    #"BACKGROUND": "FFf5f5f5",
    "LIGHT_TEXT": "FFFFFFFF",
    "DARK_TEXT": "FF212121",
    "BORDER": "FFe0e0e0",
    "WARNING": "FFf44336",
    "SUCCESS": "FF4CAF50",
    "INFO": "FF2196F3"
}

FONT_SIZES = {
    "SMALL": 9,
    "NORMAL": 10,
    "LARGE": 12,
    "HEADER": 14
}

# Configuration du logging
LOG_FILE = f'excel_compiler_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)


class FileVerification:
    """
    Classe utilitaire pour vérifier la compatibilité des fichiers Excel et CSV.
    """
    
    @staticmethod
    def verify_excel_file(file_path: str, header_start_row: int, header_rows: int) -> Tuple[bool, str]:
        """
        Vérifie si un fichier Excel est compatible pour la compilation.
        
        Args:
            file_path: Chemin complet du fichier à vérifier
            header_start_row: Ligne de début des en-têtes
            header_rows: Nombre de lignes d'en-tête
            
        Returns:
            Tuple[bool, str]: (est_compatible, message_d'erreur)
        """
        try:
            # Essai d'ouverture du fichier
            wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            ws = wb.active
            
            # Vérification du nombre de lignes
            if ws.max_row < header_start_row + header_rows:
                return False, f"Structure d'en-tête incompatible: le fichier n'a que {ws.max_row} lignes"
                
            # Vérification de la protection du fichier
            if hasattr(ws, 'protection') and ws.protection.sheet:
                return False, "Le fichier est protégé en écriture"
                
            # Autres vérifications possibles...
            
            wb.close()
            return True, "Fichier compatible"
            
        except PermissionError:
            return False, "Le fichier est ouvert dans une autre application"
        except Exception as e:
            return False, f"Erreur lors de la vérification: {str(e)}"
    
    @staticmethod
    def verify_csv_file(file_path: str, header_start_row: int, header_rows: int) -> Tuple[bool, str]:
        """
        Vérifie si un fichier CSV est compatible pour la compilation.
        
        Args:
            file_path: Chemin complet du fichier à vérifier
            header_start_row: Ligne de début des en-têtes
            header_rows: Nombre de lignes d'en-tête
            
        Returns:
            Tuple[bool, str]: (est_compatible, message_d'erreur)
        """
        try:
            # Compter les lignes du fichier
            with open(file_path, 'r', newline='', encoding='utf-8-sig') as csvfile:
                reader = csv.reader(csvfile)
                row_count = sum(1 for _ in reader)
                
            if row_count < header_start_row + header_rows:
                return False, f"Structure d'en-tête incompatible: le fichier n'a que {row_count} lignes"
                
            return True, "Fichier compatible"
            
        except UnicodeDecodeError:
            # Essayer avec une autre encodage
            try:
                with open(file_path, 'r', newline='', encoding='latin-1') as csvfile:
                    reader = csv.reader(csvfile)
                    row_count = sum(1 for _ in reader)
                    
                if row_count < header_start_row + header_rows:
                    return False, f"Structure d'en-tête incompatible: le fichier n'a que {row_count} lignes"
                    
                return True, "Fichier compatible (encodage latin-1)"
            except Exception as e:
                return False, f"Erreur d'encodage: {str(e)}"
                
        except PermissionError:
            return False, "Le fichier est ouvert dans une autre application"
        except Exception as e:
            return False, f"Erreur lors de la vérification: {str(e)}"


class CompilationWorker(QThread):
    """
    Thread de travail qui gère la compilation des fichiers Excel et CSV.
    Permet de ne pas bloquer l'interface utilisateur pendant le traitement.
    """
    
    progress = pyqtSignal(int)  # Signal de progression
    error = pyqtSignal(str)     # Signal d'erreur
    finished = pyqtSignal(tuple)  # Signal de fin avec les données compilées
    
    def __init__(self, files, directory, header_start_row, header_rows, add_filename=True,
                 sort_data=False, sort_column=0, repeat_headers=False, remove_empty_rows=False,
                 remove_duplicates=False, parent=None):
        """
        Initialisation du worker de compilation.
        
        Args:
            files: Liste des noms de fichiers à compiler
            directory: Répertoire contenant les fichiers
            header_start_row: Ligne de début des en-têtes
            header_rows: Nombre de lignes d'en-tête
            add_filename: Ajouter le nom du fichier source comme colonne
            sort_data: Trier les données
            sort_column: Colonne de tri (index)
            repeat_headers: Répéter les en-têtes pour chaque fichier
            remove_empty_rows: Supprimer les lignes vides
            remove_duplicates: Supprimer les doublons
        """
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
        self.remove_duplicates = remove_duplicates
        
    def run(self):
        """
        Méthode principale exécutée dans le thread.
        Réalise la compilation des fichiers.
        """
        combined_data = []
        headers = None
        merged_cells = []
        preliminary_info = []
        
        successful_files = []
        failed_files = []
        
        for i, file in enumerate(self.files):
            try:
                file_path = os.path.join(self.directory, file)
                
                # Traitement différent selon le type de fichier
                if file.lower().endswith(('.xlsx', '.xls')):
                    result = self._process_excel_file(file_path, i, headers, preliminary_info)
                elif file.lower().endswith('.csv'):
                    result = self._process_csv_file(file_path, i, headers, preliminary_info)
                else:
                    raise ValueError(f"Format de fichier non pris en charge: {file}")
                    
                if result:
                    file_headers, file_data, file_merged_cells = result
                    
                    # Mise à jour des en-têtes globales si nécessaire
                    if headers is None:
                        headers = file_headers
                        merged_cells = file_merged_cells
                        
                    # Ajout des données au résultat combiné    
                    combined_data.extend(file_data)
                    successful_files.append(file)
                
                self.progress.emit(i + 1)
                
            except Exception as e:
                logging.error(f"Erreur lors du traitement du fichier {file}: {str(e)}")
                logging.error(traceback.format_exc())
                failed_files.append((file, str(e)))
                self.error.emit(f"Erreur avec le fichier {file}: {str(e)}")
                continue
        
        # Traitement post-compilation
        if combined_data and headers:
            # Suppression des doublons si demandé
            if self.remove_duplicates:
                combined_data = self._remove_duplicate_rows(combined_data)
                
            # Tri des données si l'option est activée
            if self.sort_data and self.sort_column < len(headers[-1]):
                combined_data = self._sort_data(combined_data, headers)
            
            # Ajout du nom du fichier source à l'en-tête si nécessaire    
            if self.add_filename and headers:
                headers[-1].append("Fichier source")
        
        # Émission du signal de fin avec toutes les données
        self.finished.emit((preliminary_info, headers, combined_data, merged_cells, successful_files, failed_files))
            
    def _process_excel_file(self, file_path, file_index, global_headers, preliminary_info):
        """
        Traite un fichier Excel et extrait ses données.
        
        Args:
            file_path: Chemin du fichier
            file_index: Index du fichier dans la liste
            global_headers: En-têtes déjà établies (pour les fichiers suivants)
            preliminary_info: Informations préliminaires à collecter du premier fichier
            
        Returns:
            Tuple contenant les en-têtes, données et cellules fusionnées du fichier
        """
        # Pour le premier fichier, on a besoin des merged_cells donc on ne met pas read_only
        if file_index == 0:
            wb = openpyxl.load_workbook(file_path, data_only=True)
        else:
            wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            
        ws = wb.active
        file_merged_cells = []
        file_headers = []
        
        # Capture des informations préliminaires du premier fichier
        if file_index == 0 and self.header_start_row > 1:
            for row in range(1, self.header_start_row):
                row_data = []
                for cell in ws[row]:
                    row_data.append(cell.value)
                preliminary_info.append(row_data)
        
        # Capture des en-têtes du premier fichier ou utilisation des en-têtes globales
        if global_headers is None:
            for row in range(self.header_start_row, self.header_start_row + self.header_rows):
                header_row = []
                for cell in ws[row]:
                    header_row.append(cell.value)
                file_headers.append(header_row)
                
                # Capture des merged_cells seulement pour le premier fichier
                if hasattr(ws, 'merged_cells'):
                    for merged_range in ws.merged_cells.ranges:
                        if merged_range.min_row <= (self.header_start_row + self.header_rows):
                            file_merged_cells.append(merged_range)
        else:
            file_headers = global_headers
            
        # Ajout des données du fichier
        file_data = self._extract_excel_data(ws, file_path.split(os.sep)[-1])
        
        # Fermer le workbook pour libérer la mémoire
        wb.close()
        
        return file_headers, file_data, file_merged_cells
        
    def _extract_excel_data(self, worksheet, filename):
        """
        Extrait les données d'une feuille Excel.
        
        Args:
            worksheet: Feuille de calcul à traiter
            filename: Nom du fichier source
            
        Returns:
            Liste des données extraites
        """
        data = []
        
        # Si répétition des en-têtes est activée et ce n'est pas le premier fichier
        if self.repeat_headers and len(data) > 0:
            # Ajouter une ligne vide comme séparateur
            data.append([None] * (len(worksheet[self.header_start_row]) + (1 if self.add_filename else 0)))
            
            # Ajouter les en-têtes
            for row in range(self.header_start_row, self.header_start_row + self.header_rows):
                row_data = [cell.value for cell in worksheet[row]]
                if self.add_filename:
                    row_data.append(None)  # Pas de nom de fichier dans l'en-tête répété
                data.append(row_data)
        
        # Ajout des données
        for row in worksheet.iter_rows(min_row=self.header_start_row + self.header_rows):
            row_data = [cell.value for cell in row]
            
            # Ajouter le nom du fichier si demandé
            if self.add_filename:
                row_data.append(filename)
            
            # Vérifier si la ligne n'est pas vide avant de l'ajouter
            if not self.remove_empty_rows or not all(
                cell is None or str(cell).strip() == "" 
                for cell in row_data[:-1 if self.add_filename else None]
            ):
                data.append(row_data)
                
        return data
        
    def _process_csv_file(self, file_path, file_index, global_headers, preliminary_info):
        """
        Traite un fichier CSV et extrait ses données.
        
        Args:
            file_path: Chemin du fichier
            file_index: Index du fichier dans la liste
            global_headers: En-têtes déjà établies (pour les fichiers suivants)
            preliminary_info: Informations préliminaires à collecter du premier fichier
            
        Returns:
            Tuple contenant les en-têtes et données du fichier
        """
        # Détection de l'encodage du fichier
        encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    # Détection du délimiteur
                    sample = f.read(4096)
                    sniffer = csv.Sniffer()
                    dialect = sniffer.sniff(sample)
                    delimiter = dialect.delimiter
                    break
            except Exception:
                continue
        else:
            raise ValueError(f"Impossible de déterminer l'encodage du fichier CSV: {file_path}")
        
        # Lecture du CSV avec pandas
        df = pd.read_csv(file_path, delimiter=delimiter, header=None, encoding=encoding)
        
        file_headers = []
        file_merged_cells = []  # Toujours vide pour les CSV
        
        # Capture des informations préliminaires du premier fichier
        if file_index == 0 and self.header_start_row > 1:
            for row in range(0, self.header_start_row - 1):
                if row < len(df):
                    preliminary_info.append(df.iloc[row].tolist())
        
        # Capture des en-têtes du premier fichier ou utilisation des en-têtes globales
        if global_headers is None:
            for row in range(self.header_start_row - 1, self.header_start_row - 1 + self.header_rows):
                if row < len(df):
                    file_headers.append(df.iloc[row].tolist())
        else:
            file_headers = global_headers
            
        # Extraction des données
        file_data = []
        
        # Si répétition des en-têtes est activée et ce n'est pas le premier fichier
        if self.repeat_headers and len(file_data) > 0:
            # Ajouter une ligne vide comme séparateur
            file_data.append([None] * (len(file_headers[-1]) + (1 if self.add_filename else 0)))
            
            # Ajouter les en-têtes
            for header_row in file_headers:
                row_data = header_row.copy()
                if self.add_filename:
                    row_data.append(None)
                file_data.append(row_data)
        
        # Ajout des données
        for row in range(self.header_start_row - 1 + self.header_rows, len(df)):
            row_data = df.iloc[row].tolist()
            
            # Ajouter le nom du fichier si demandé
            if self.add_filename:
                row_data.append(os.path.basename(file_path))
            
            # Vérifier si la ligne n'est pas vide avant de l'ajouter
            if not self.remove_empty_rows or not all(
                pd.isna(cell) or str(cell).strip() == "" 
                for cell in row_data[:-1 if self.add_filename else None]
            ):
                file_data.append(row_data)
                
        return file_headers, file_data, file_merged_cells
    
    def _remove_duplicate_rows(self, data):
        """
        Supprime les lignes en double dans les données.
        
        Args:
            data: Liste des données à filtrer
            
        Returns:
            Liste des données sans doublons
        """
        unique_data = []
        seen = set()
        
        for row in data:
            # Convertir la ligne en tuple pour pouvoir l'ajouter à un set
            row_tuple = tuple(str(cell) if cell is not None else None for cell in row)
            
            if row_tuple not in seen:
                seen.add(row_tuple)
                unique_data.append(row)
                
        return unique_data
        
    def _sort_data(self, data, headers):
        """
        Trie les données selon la colonne spécifiée.
        
        Args:
            data: Données à trier
            headers: En-têtes pour déterminer le nombre de colonnes
            
        Returns:
            Données triées
        """
        try:
            sort_idx = self.sort_column - 1
            
            if self.repeat_headers:
                # Cas spécial: préserver les sections avec en-têtes répétés
                sections = []
                current_section = []
                
                for row in data:
                    # Une ligne entièrement vide indique un séparateur de section
                    if all(cell is None for cell in row):
                        if current_section:
                            sections.append(current_section)
                        current_section = [row]  # Garder la ligne vide
                    else:
                        current_section.append(row)
                
                if current_section:
                    sections.append(current_section)
                
                # Trier chaque section individuellement
                for section in sections:
                    # Séparer l'en-tête et les données
                    header_rows = []
                    data_rows = []
                    
                    for i, row in enumerate(section):
                        # La première ligne est le séparateur, puis viennent les en-têtes
                        if i < self.header_rows + 1:
                            header_rows.append(row)
                        else:
                            data_rows.append(row)
                    
                    # Trier uniquement les données
                    data_rows.sort(key=lambda x: (x[sort_idx] is None, x[sort_idx]))
                    
                    # Recombiner
                    section.clear()
                    section.extend(header_rows)
                    section.extend(data_rows)
                
                # Aplatir les sections
                sorted_data = []
                for section in sections:
                    sorted_data.extend(section)
                
                return sorted_data
            else:
                # Cas simple: trier toutes les données
                return sorted(data, key=lambda x: (x[sort_idx] is None, x[sort_idx]))
                
        except Exception as e:
            logging.warning(f"Erreur lors du tri : {str(e)}")
            return data


class ExcelFormatter:
    """
    Classe responsable du formatage du fichier Excel de sortie.
    """
    
    # Styles prédéfinis
    HEADER_STYLE = {
        'font': Font(bold=True, color=EXCEL_COLORS["LIGHT_TEXT"]),
        'fill': PatternFill(start_color=EXCEL_COLORS["PRIMARY"], end_color=EXCEL_COLORS["PRIMARY"], fill_type='solid'),
        'alignment': Alignment(horizontal='center', vertical='center')
    }
    
    PRELIMINARY_STYLE = {
        'font': Font(italic=True)
    }
    
    DATA_BORDER = Border(
        left=Side(border_style='thin'),
        right=Side(border_style='thin'),
        top=Side(border_style='thin'),
        bottom=Side(border_style='thin')
    )
    
    @staticmethod
    def write_preliminary_info(worksheet, preliminary_info):
        """
        Écrit les informations préliminaires dans la feuille de calcul.
        
        Args:
            worksheet: Feuille de calcul à modifier
            preliminary_info: Données préliminaires à écrire
            
        Returns:
            Ligne courante après écriture
        """
        current_row = 1
        
        for row_data in preliminary_info:
            for col_idx, value in enumerate(row_data, 1):
                cell = worksheet.cell(row=current_row, column=col_idx, value=value)
                cell.font = ExcelFormatter.PRELIMINARY_STYLE['font']
            current_row += 1
            
        return current_row
    
    @staticmethod
    def write_headers(worksheet, headers, start_row):
        """
        Écrit les en-têtes dans la feuille de calcul avec le style approprié.
        
        Args:
            worksheet: Feuille de calcul à modifier
            headers: Données d'en-tête à écrire
            start_row: Ligne de début pour écrire les en-têtes
            
        Returns:
            Ligne courante après écriture
        """
        current_row = start_row
        
        for header_row in headers:
            for col_idx, value in enumerate(header_row, 1):
                cell = worksheet.cell(row=current_row, column=col_idx, value=value)
                cell.font = ExcelFormatter.HEADER_STYLE['font']
                cell.fill = ExcelFormatter.HEADER_STYLE['fill']
                cell.alignment = ExcelFormatter.HEADER_STYLE['alignment']
            current_row += 1
            
        return current_row
    
    @staticmethod
    def write_data(worksheet, data, start_row):
        """
        Écrit les données dans la feuille de calcul avec le style approprié.
        
        Args:
            worksheet: Feuille de calcul à modifier
            data: Données à écrire
            start_row: Ligne de début pour écrire les données
            
        Returns:
            Ligne courante après écriture
        """
        current_row = start_row
        
        for row_data in data:
            for col_idx, value in enumerate(row_data, 1):
                cell = worksheet.cell(row=current_row, column=col_idx, value=value)
                cell.border = ExcelFormatter.DATA_BORDER
            current_row += 1
            
        return current_row
    
    @staticmethod
    def apply_merged_cells(worksheet, merged_ranges, header_start_original, header_start_new):
        """
        Applique les fusions de cellules en ajustant les numéros de ligne.
        
        Args:
            worksheet: Feuille de calcul à modifier
            merged_ranges: Plages de cellules à fusionner
            header_start_original: Ligne de début originale des en-têtes
            header_start_new: Nouvelle ligne de début des en-têtes
        """
        for merged_range in merged_ranges:
            # Ajuster les numéros de ligne
            adjusted_range = openpyxl.worksheet.cell_range.CellRange(
                min_col=merged_range.min_col,
                min_row=merged_range.min_row - header_start_original + header_start_new,
                max_col=merged_range.max_col,
                max_row=merged_range.max_row - header_start_original + header_start_new
            )
            worksheet.merge_cells(range_string=adjusted_range.coord)
    
    @staticmethod
    def adjust_column_widths(worksheet):
        """
        Ajuste la largeur des colonnes en fonction du contenu.
        
        Args:
            worksheet: Feuille de calcul à modifier
        """
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
                    
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column_letter].width = min(adjusted_width, 50)
    
    @staticmethod
    def freeze_header(worksheet, freeze_row):
        """
        Fige les volets à la ligne spécifiée.
        
        Args:
            worksheet: Feuille de calcul à modifier
            freeze_row: Ligne à partir de laquelle figer les volets
        """
        worksheet.freeze_panes = worksheet.cell(row=freeze_row, column=1)


class VerificationReportDialog(QDialog):
    """
    Boîte de dialogue affichant le rapport de vérification des fichiers avant compilation.
    """
    
    def __init__(self, parent, compatible_files, incompatible_files):
        """
        Initialise la boîte de dialogue avec les résultats de la vérification.
        
        Args:
            parent: Widget parent
            compatible_files: Liste des fichiers compatibles
            incompatible_files: Liste des fichiers incompatibles avec raisons
        """
        super().__init__(parent)
        self.setWindowTitle("Rapport de vérification des fichiers")
        self.setMinimumSize(800, 600)
        self.compatible_files = compatible_files
        self.incompatible_files = incompatible_files
        self.continue_with_compatible = False
        
        self.init_ui()
    
    def init_ui(self):
        """Initialise l'interface utilisateur de la boîte de dialogue."""
        layout = QVBoxLayout()
        
        # En-tête avec statistiques
        total_files = len(self.compatible_files) + len(self.incompatible_files)
        header_label = QLabel(f"<h2>Vérification de {total_files} fichiers</h2>")
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(header_label)
        
        compat_percent = len(self.compatible_files) * 100 / total_files if total_files > 0 else 0
        incompat_percent = len(self.incompatible_files) * 100 / total_files if total_files > 0 else 0
        
        stats_label = QLabel(
            f"<div style='text-align:center; margin:10px 0;'>"
            f"<span style='color:{COLORS['SUCCESS']}; font-weight:bold;'>Compilables: {len(self.compatible_files)} "
            f"({compat_percent:.1f}%)</span> | "
            f"<span style='color:{COLORS['WARNING']}; font-weight:bold;'>Non compilables: {len(self.incompatible_files)} "
            f"({incompat_percent:.1f}%)</span>"
            f"</div>"
        )
        stats_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(stats_label)
        
        # Tableau des fichiers incompatibles
        if self.incompatible_files:
            group_incompatible = QGroupBox("Fichiers non compilables")
            group_layout = QVBoxLayout()
            
            # Instructions pour résoudre les problèmes
            help_label = QLabel(
                "Ces fichiers ne peuvent pas être compilés pour les raisons indiquées. "
                "Voici quelques conseils pour résoudre les problèmes courants :"
                "<ul>"
                "<li><b>Fichier ouvert</b>: Fermez le fichier dans Excel et réessayez</li>"
                "<li><b>Fichier protégé</b>: Désactivez la protection dans Excel (Révision > Protéger la feuille)</li>"
                "<li><b>Structure d'en-tête incompatible</b>: Vérifiez que le nombre de lignes d'en-tête est correct</li>"
                "<li><b>Erreur d'encodage</b>: Réenregistrez le fichier CSV avec l'encodage UTF-8</li>"
                "</ul>"
            )
            help_label.setWordWrap(True)
            group_layout.addWidget(help_label)
            
            table = QTableWidget()
            table.setColumnCount(2)
            table.setHorizontalHeaderLabels(["Nom du fichier", "Problème détecté"])
            table.setRowCount(len(self.incompatible_files))
            table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
            
            for row, (file_name, reason) in enumerate(self.incompatible_files):
                table.setItem(row, 0, QTableWidgetItem(file_name))
                table.setItem(row, 1, QTableWidgetItem(reason))
                # Colorer la ligne en rouge pâle pour mettre en évidence
                for col in range(2):
                    table.item(row, col).setBackground(QBrush(QColor("#ffebee")))
            
            group_layout.addWidget(table)
            group_incompatible.setLayout(group_layout)
            layout.addWidget(group_incompatible)
        
        # Liste des fichiers compatibles
        if self.compatible_files:
            group_compatible = QGroupBox("Fichiers compilables")
            group_layout = QVBoxLayout()
        
            list_widget = QListWidget()
            for file in self.compatible_files:
                item = QListWidgetItem(file)
                item.setBackground(QBrush(QColor("#e8f5e9")))  # Vert pâle
                list_widget.addItem(item)
            
            group_layout.addWidget(list_widget)
            group_compatible.setLayout(group_layout)
            layout.addWidget(group_compatible)
        
        # Boutons
        button_layout = QHBoxLayout()

        if self.incompatible_files and self.compatible_files:
            ignore_button = QPushButton("Ignorer les non compilables et compiler")
            ignore_button.setStyleSheet(f"background-color: #{COLORS['INFO']}; color: white;")
            ignore_button.clicked.connect(self.continue_with_compatible_only)
            button_layout.addWidget(ignore_button)

        """"
        if self.compatible_files:
            compile_button = QPushButton("Compiler les fichiers compatibles")
            compile_button.setStyleSheet(f"background-color: #{COLORS['SUCCESS']}; color: white;")
            compile_button.clicked.connect(self.accept)
            compile_button.setDefault(True)
            button_layout.addWidget(compile_button)

        """
        cancel_button = QPushButton("Annuler")
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)
    
    def continue_with_compatible_only(self):
        """
        Méthode appelée quand l'utilisateur choisit d'ignorer les fichiers incompatibles.
        """
        self.continue_with_compatible = True
        self.accept()


class CompilationReportDialog(QDialog):
    """
    Boîte de dialogue affichant le rapport après compilation.
    """
    
    def __init__(self, parent, successful_files, failed_files, output_path):
        """
        Initialise la boîte de dialogue avec les résultats de la compilation.
        
        Args:
            parent: Widget parent
            successful_files: Liste des fichiers compilés avec succès
            failed_files: Liste des fichiers échoués avec raisons
            output_path: Chemin du fichier de sortie
        """
        super().__init__(parent)
        self.setWindowTitle("Rapport de compilation")
        self.setMinimumSize(800, 600)
        self.successful_files = successful_files
        self.failed_files = failed_files
        self.output_path = output_path
        
        self.init_ui()
    
    def init_ui(self):
        """Initialise l'interface utilisateur de la boîte de dialogue."""
        layout = QVBoxLayout()
        
        # En-tête avec statistiques
        total_files = len(self.successful_files) + len(self.failed_files)
        header_label = QLabel(f"<h2>Résultat de la compilation</h2>")
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(header_label)
        
        # Affichage du chemin du fichier de sortie
        output_label = QLabel(f"<div style='text-align:center; margin:5px 0;'>"
                             f"<b>Fichier généré:</b> {self.output_path}"
                             f"</div>")
        output_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(output_label)
        
        # Statistiques
        percentage = len(self.successful_files) * 100 / total_files if total_files > 0 else 0
        stats_label = QLabel(f"<div style='text-align:center; margin:10px 0; font-size:16px;'>"
                            f"<b>Taux de compilation: {percentage:.1f}%</b> ({len(self.successful_files)}/{total_files})"
                            f"</div>")
        stats_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(stats_label)
        
        # Icône de succès si tout a été compilé
        if percentage == 100:
            success_icon = QLabel()
            pixmap = self.style().standardIcon(QStyle.StandardPixmap.SP_DialogApplyButton).pixmap(64, 64)
            success_icon.setPixmap(pixmap)
            success_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(success_icon)
            
            success_message = QLabel("<div style='text-align:center; color:green; font-weight:bold;'>"
                                    "Tous les fichiers ont été compilés avec succès !"
                                    "</div>")
            success_message.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(success_message)
        
        # Liste des fichiers compilés
        if self.successful_files:
            success_group = QGroupBox(f"Fichiers compilés ({len(self.successful_files)})")
            success_layout = QVBoxLayout()
            
            list_widget = QListWidget()
            for file in self.successful_files:
                item = QListWidgetItem(file)
                item.setBackground(QBrush(QColor("#e8f5e9")))  # Vert pâle
                list_widget.addItem(item)
            
            success_layout.addWidget(list_widget)
            success_group.setLayout(success_layout)
            layout.addWidget(success_group)
        
        # Tableau des échecs si existants
        if self.failed_files:
            failed_group = QGroupBox(f"Fichiers non compilés ({len(self.failed_files)})")
            failed_layout = QVBoxLayout()
            
            table = QTableWidget()
            table.setColumnCount(2)
            table.setHorizontalHeaderLabels(["Nom du fichier", "Erreur"])
            table.setRowCount(len(self.failed_files))
            table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
            
            for row, (file_name, reason) in enumerate(self.failed_files):
                table.setItem(row, 0, QTableWidgetItem(file_name))
                table.setItem(row, 1, QTableWidgetItem(reason))
                # Colorer la ligne en rouge pâle
                for col in range(2):
                    table.item(row, col).setBackground(QBrush(QColor("#ffebee")))
            
            failed_layout.addWidget(table)
            failed_group.setLayout(failed_layout)
            layout.addWidget(failed_group)
        
        # Bouton OK
        button = QPushButton("OK")
        button.clicked.connect(self.accept)
        layout.addWidget(button, alignment=Qt.AlignmentFlag.AlignCenter)
        
        self.setLayout(layout)


class IPWarningDialog(QDialog):
    """
    Boîte de dialogue d'avertissement concernant la propriété intellectuelle.
    """
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Propriété Intellectuelle - Avertissement")
        self.setMinimumSize(800, 600)
        
        self.init_ui()
    
    def init_ui(self):
        """Initialise l'interface utilisateur de la boîte de dialogue."""
        layout = QVBoxLayout()
        
        # Titre
        title_label = QLabel("<h1>Avertissement</h1>")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)
        
        # Contenu
        content = QLabel(
            "<p style='font-size: 14px; line-height: 1.5;'>"
            "Cette application <b>Compilateur Excel Professionnel</b> est la propriété intellectuelle exclusive de:<br><br>"
            "<div style='text-align: center; font-weight: bold; font-size: 16px; margin: 20px 0;'>"
            "GOUNOU N'GOBI Chabi Zimé<br>"
            "Data Manager & Data Analyst<br><br>"
            "</div>"
            "Tous droits réservés. Cette application est protégée par les lois sur le droit d'auteur "
            "et les traités internationaux sur la propriété intellectuelle.<br><br>"
            
            "<span style='color: #f44336; font-weight: bold;'>AVERTISSEMENT IMPORTANT:</span><br>"
            "<ul>"
            "<li>Toute reproduction non autorisée, distribution ou modification de cette application "
            "est strictement interdite.</li>"
            "<li>L'utilisation de cette application est soumise aux termes de la licence accordée par l'auteur.</li>"
            "</ul><br>"
            
            "Pour toute question concernant les droits d'utilisation ou pour signaler une violation "
            "de la propriété intellectuelle, veuillez contacter l'auteur à l'adresse : zimkada@gmail.com."
            
            "</p>"
        )
        content.setWordWrap(True)
        content.setTextFormat(Qt.TextFormat.RichText)
        
        scroll = QScrollArea()
        scroll.setWidget(content)
        scroll.setWidgetResizable(True)
        layout.addWidget(scroll)
        
        # Boutons
        button_layout = QHBoxLayout()
        
        accept_button = QPushButton("J'accepte ces conditions")
        accept_button.clicked.connect(self.accept)
        accept_button.setDefault(True)
        
        exit_button = QPushButton("Quitter l'application")
        exit_button.clicked.connect(self.reject)
        
        button_layout.addWidget(accept_button)
        button_layout.addWidget(exit_button)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)


class ModernExcelCompilerApp(QMainWindow):
    """
    Classe principale de l'application de compilation Excel.
    """
    def __init__(self):
        """Initialise l'application."""
        super().__init__()
        self.setup_ui()
        self.setup_variables()
        self.create_main_layout()  
        self.connect_signals()
        logging.info("Application démarrée")
        
        # Afficher l'avertissement de propriété intellectuelle au démarrage
        self.show_ip_warning()
    
    def create_main_layout(self):
        """Crée la mise en page principale de l'application."""
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        
        # Création des onglets
        self.tabs = QTabWidget()
        self.create_compilation_tab()
        self.create_advanced_options_tab()
        self.create_help_tab()
        self.create_about_tab()
        
        # Appliquer un style spécifique à la barre d'onglets pour qu'elle soit verte
        self.tabs.setStyleSheet(f"""
            QTabWidget::pane {{
                border: 1px solid #{COLORS["PRIMARY"]};
            }}
            QTabBar {{
                background-color: #{COLORS["PRIMARY"]};
            }}
            QTabBar::tab {{
                background-color: #{COLORS["PRIMARY"]};
                color: white;
                padding: 8px 15px;
            }}
            QTabBar::tab:selected {{
                background-color: #{COLORS["PRIMARY_DARK"]};
                font-weight: bold;
            }}
            QTabBar::tab:!selected {{
                margin-top: 2px;
            }}
        """)
        
        main_layout.addWidget(self.tabs)



    def setup_ui(self):
        """Configure l'interface utilisateur principale."""
        self.setWindowTitle("Compilateur Excel Professionnel")
        self.setGeometry(50, 50, 1200, 700)
        self.setWindowIcon(QIcon("icon.jpg"))
        self.setMinimumSize(800, 600)
        self.apply_stylesheet()
    
    def setup_variables(self):
        """Initialise les variables de l'application."""
        self.directory = ""
        self.files = []
        self.compilation_worker = None
        self.verification_enabled = True  # Par défaut, la vérification préliminaire est activée
    
    def apply_stylesheet(self):
        """Applique le style CSS à l'application."""
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: {COLORS["BACKGROUND"]};
            }}
            
            /* Style pour la barre de menus */
            QMenuBar {{
                background-color: #{COLORS["PRIMARY"]};
                color: white;
                border: none;
            }}
            
            QMenuBar::item {{
                background-color: #{COLORS["PRIMARY"]};
                color: white;
            }}
            
            QMenuBar::item:selected {{
                background-color: #{COLORS["PRIMARY_DARK"]};
            }}
            
            QMenu {{
                background-color: white;
                border: 1px solid #{COLORS["PRIMARY"]};
            }}
            
            QMenu::item:selected {{
                background-color: #{COLORS["PRIMARY_LIGHT"]};
                color: white;
            }}
            
            /* Style pour les onglets */
            QTabWidget::pane {{
                border: 1px solid #{COLORS["PRIMARY"]};
                border-radius: 4px;
                background-color: white;
            }}
            
            QTabBar {{
                background-color: #{COLORS["PRIMARY"]};
            }}
            
            QTabBar::tab {{
                background-color: #{COLORS["PRIMARY"]};
                color: white;
                border: 1px solid #{COLORS["PRIMARY_DARK"]};
                padding: 8px 15px;
                margin-right: 2px;
            }}
            
            QTabBar::tab:selected {{
                background-color: #{COLORS["PRIMARY_DARK"]};
                color: white;
            }}
            
            QTabBar::tab:hover {{
                background-color: #{COLORS["PRIMARY_LIGHT"]};
                color: white;
            }}
            
            /* Style pour les groupes et bordures */
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #{COLORS["PRIMARY"]};
                border-radius: 8px;
                margin-top: 12px;
                padding: 15px;
                background-color: white;
            }}
            
            QGroupBox::title {{
                color: #{COLORS["PRIMARY"]};
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }}
            
            /* Style pour les boutons */
            QPushButton {{
                background-color: #{COLORS["PRIMARY"]};
                color: white;
                padding: 8px 15px;
                border-radius: 4px;
                font-weight: bold;
                min-height: 30px;
            }}
            
            QPushButton:hover {{
                background-color: #{COLORS["PRIMARY_DARK"]};
            }}
            
            QPushButton:disabled {{
                background-color: #a0a0a0;
                color: #d0d0d0;
            }}
            
            /* Style pour les tableaux */
            QTableWidget {{
                border: 1px solid #{COLORS["PRIMARY"]};
                gridline-color: #{COLORS["PRIMARY"]};
            }}
            
            QTableWidget QHeaderView::section {{
                background-color: #{COLORS["PRIMARY"]};
                color: white;
                padding: 5px;
                border: 1px solid white;
            }}
            
            /* Style pour la barre de progression */
            QProgressBar {{
                border: 1px solid #{COLORS["PRIMARY"]};
                border-radius: 4px;
                text-align: center;
            }}
            
            QProgressBar::chunk {{
                background-color: #{COLORS["PRIMARY"]};
                width: 10px;
                margin: 0.5px;
            }}
            
            /* Style pour les checkboxes */
            QCheckBox::indicator:checked {{
                background-color: #{COLORS["PRIMARY"]};
                border: 1px solid #{COLORS["PRIMARY_DARK"]};
            }}
            
            /* Style pour les étiquettes de titre */
            QLabel[accessibleName="title"] {{
                color: #{COLORS["PRIMARY"]};
                font-weight: bold;
            }}
            
            /* Style pour les en-têtes de vue */
            QHeaderView::section {{
                background-color: #{COLORS["PRIMARY"]};
                color: white;
            }}
            
            /* Style pour tous les bordures et contours */
            * {{
                border-color: #{COLORS["PRIMARY"]};
            }}
        """)

    
    def create_compilation_tab(self):
        """Crée l'onglet principal de compilation."""
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
        self.button_compile.setFont(QFont("Segoe UI", FONT_SIZES["LARGE"]))
        self.button_compile.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay))
        self.button_compile.setStyleSheet(f"background-color: #{COLORS['PRIMARY']}; color: white;")
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
        """Crée le groupe d'éléments pour la sélection des fichiers."""
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
        self.button_choose_directory.setFont(QFont("Segoe UI", FONT_SIZES["NORMAL"]))
        self.button_choose_directory.setIcon(QIcon("folder.png"))
        self.button_choose_directory.setStyleSheet(f"background-color: #{COLORS['PRIMARY']}; color: white;")

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
        """Crée le groupe d'éléments pour les options de compilation."""
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
        self.checkbox_add_filename.setChecked(True)
        
        self.label_output_name = QLabel("Nom du fichier de sortie :")
        self.lineedit_output_name = QLineEdit("compilation.xlsx")
        output_container = QHBoxLayout()
        output_container.addWidget(self.label_output_name)
        output_container.addWidget(self.lineedit_output_name)
        
        # Option de vérification préliminaire
        verification_layout = QHBoxLayout()
        self.checkbox_verify_files = QCheckBox("Activer la vérification préliminaire des fichiers")
        self.checkbox_verify_files.setChecked(True)
        verification_layout.addWidget(self.checkbox_verify_files)
        verification_layout.addStretch()
        
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
        
        # Ligne supplémentaire pour la vérification
        grid_layout.addLayout(verification_layout, 2, 0, 1, 3)
        
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
        """Met à jour l'affichage de la date et de l'heure."""
        current_datetime = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        self.datetime_label.setText(f"Date et heure : {current_datetime}")
    
    def create_advanced_options_tab(self):
        """Crée l'onglet des options avancées."""
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
        
        self.checkbox_freeze_header = QCheckBox("Figer les en-têtes")
        self.checkbox_freeze_header.setChecked(True)
        
        format_layout.addWidget(self.checkbox_auto_width)
        format_layout.addWidget(self.checkbox_freeze_header)
        format_group.setLayout(format_layout)
        
        # Groupe formats des fichiers
        format_files_group = QGroupBox("Formats de fichiers supportés")
        format_files_layout = QVBoxLayout()
        
        self.checkbox_excel = QCheckBox("Fichiers Excel (.xlsx, .xls)")
        self.checkbox_excel.setChecked(True)
        self.checkbox_excel.setEnabled(False)  # Toujours activé
        
        self.checkbox_csv = QCheckBox("Fichiers CSV (.csv)")
        self.checkbox_csv.setChecked(True)
        
        format_files_layout.addWidget(self.checkbox_excel)
        format_files_layout.addWidget(self.checkbox_csv)
        format_files_group.setLayout(format_files_layout)
        
        layout.addWidget(data_group)
        layout.addWidget(format_group)
        layout.addWidget(format_files_group)
        layout.addStretch()
        tab.setLayout(layout)
        self.tabs.addTab(tab, "Options avancées")
    
    def create_help_tab(self):
        """Crée l'onglet d'aide."""
        tab = QWidget()
        layout = QVBoxLayout()
        
        help_text = """
        <style>
        body {
            font-family: "Segoe UI", sans-serif;
            margin: 0 20px;
            line-height: 1.6;
            color: #333;
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

        .highlight {
            background-color: #e8f5e9;
            padding: 2px 5px;
            border-radius: 3px;
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
        
        <h3>I. Sélection des fichiers</h3>
        <div class="section-content">
            <h4>• Cliquez sur <b>"Choisir un répertoire"</b> pour sélectionner le dossier contenant vos fichiers Excel et CSV</h4>
            <h4>• Sélectionnez les fichiers à compiler dans la liste</h4>
            <h4>• Utilisez la case <b>"Sélectionner tous les fichiers"</b> pour tout sélectionner/désélectionner</h4>
            <h4>• Le nombre de fichiers sélectionnés est affiché en temps réel</h4>
        </div>
        
        <h3>II. Options de compilation</h3>
        <div class="section-content">
            <h4>• <b>Ligne de début de l'en-tête</b> : Spécifiez à quelle ligne commence l'en-tête dans vos fichiers</h4> 
            <h4>• <b>Nombre de lignes d'en-tête</b> : Indiquez combien de lignes constituent l'en-tête</h4> 
            <h4>• <b>Répéter les en-têtes</b> : Réinsère les en-têtes entre chaque fichier dans la compilation</h4> 
            <h4>• <b>Fusionner les en-têtes multi-niveaux</b> : Conserve la fusion des cellules d'en-tête</h4> 
            <h4>• <b>Ajouter les noms des fichiers sources</b> : Ajoute une colonne avec les noms des fichiers d'origine</h4> 
            <h4>• <b>Nom du fichier de sortie</b> : Définissez le nom du fichier compilé (.xlsx sera ajouté automatiquement)</h4>
            <h4>• <b>Vérification préliminaire</b> : Active ou désactive l'analyse des fichiers avant compilation</h4>
        </div>
        
        <h3>III. Options avancées</h3>
        <div class="section-content">
            <h4><u>Traitement des données :</u></h4>
            <h4>• <b>Supprimer les doublons</b> : Élimine les lignes identiques</h4> 
            <h4>• <b>Supprimer les lignes vides</b> : Retire les lignes ne contenant aucune donnée</h4> 
            <h4>• <b>Trier les données</b> : Trie le contenu selon une colonne spécifique</h4> 
            <h4>• <b>Colonne de tri</b> : Spécifiez la colonne pour le tri (ex: A pour première colonne)</h4> 
            
            <h4><u>Formatage :</u></h4>
            <h4>• <b>Ajuster la largeur des colonnes</b> : Adapte automatiquement la largeur selon le contenu</h4> 
            <h4>• <b>Figer les en-têtes</b> : Maintient l'en-tête visible lors du défilement</h4>
            
            <h4><u>Formats de fichiers supportés :</u></h4>
            <h4>• <b>Excel</b> : Traite les fichiers .xlsx et .xls</h4>
            <h4>• <b>CSV</b> : Traite les fichiers .csv avec détection automatique du délimiteur</h4>
        </div>
        
        <h3>IV. Informations préliminaires</h3>
        <div class="section-content">
            <h4>• Pour le premier fichier uniquement, les informations situées avant l'en-tête sont préservées</h4> 
            <h4>• Ces informations sont copiées dans le fichier final</h4>
        </div>
        
        <h3>V. Vérification préliminaire des fichiers</h3>
        <div class="section-content">
            <h4>• Avant la compilation, l'application analyse tous les fichiers sélectionnés</h4>
            <h4>• Un rapport détaillé identifie les fichiers compatibles et incompatibles</h4>
            <h4>• Pour chaque fichier problématique, un motif précis est affiché</h4>
            <h4>• Vous pouvez choisir d'ignorer les fichiers incompatibles ou d'annuler la compilation</h4>
            <h4>• Cette vérification peut être désactivée dans les options de compilation</h4>
        </div>
        
        <h3>VI. Résolution des problèmes courants</h3>
        <div class="section-content">
            <h4>• <span class="highlight">Fichier ouvert dans Excel</span> : Fermez le fichier Excel avant la compilation</h4>
            <h4>• <span class="highlight">Structure d'en-tête incompatible</span> : Vérifiez que la ligne de début d'en-tête et le nombre de lignes sont corrects</h4>
            <h4>• <span class="highlight">Fichier protégé</span> : Désactivez la protection du fichier dans Excel</h4>
            <h4>• <span class="highlight">Problèmes d'encodage CSV</span> : Enregistrez le fichier en UTF-8</h4>
            <h4>• <span class="highlight">Erreurs de compilation</span> : Consultez le fichier log généré dans le répertoire de l'application</h4>
        </div>
        
        <h3>VII. Bonnes pratiques</h3>
        <div class="section-content">
            <h4>• Faites une sauvegarde de vos fichiers avant la compilation</h4>
            <h4>• Vérifiez que tous les fichiers ont une structure similaire</h4>
            <h4>• Pour les gros fichiers, traitez-les par lots</h4>
            <h4>• Utilisez des noms explicites pour les fichiers de sortie</h4>
            <h4>• Activez la vérification préliminaire pour éviter les erreurs</h4>
        </div>" \
            """

        # Création de l'étiquette avec le texte d'aide
        help_label = QLabel(help_text)
        help_label.setTextFormat(Qt.TextFormat.RichText)
        help_label.setWordWrap(True)
        
        # Ajout d'un ScrollArea pour permettre le défilement
        scroll = QScrollArea()
        scroll.setWidget(help_label)
        scroll.setWidgetResizable(True)
        
        layout.addWidget(scroll)
        tab.setLayout(layout)
        self.tabs.addTab(tab, "Aide")


    def create_about_tab(self):
        """Crée l'onglet À propos."""
        tab = QWidget()
        layout = QVBoxLayout()
        
        about_text = """
        <div style="text-align: center; margin: 50px 20px;">
            <h1 style="color: #2e7d32; margin-bottom: 30px;">Compilateur Excel Professionnel</h1>
            <h2>Version 2.0</h2>
            <p style="font-size: 16px; margin: 30px 0;">
                Développé par:<br>
                <strong style="font-size: 20px; color: #1b5e20;">GOUNOU N'GOBI Chabi Zimé</strong><br>
                Data Manager & Data Analyst
            <h2>Email : zimkada@gmail.com</h2>
            
            </p>
            <p style="margin-top: 40px; color: #555; font-style: italic;">
                Copyright © 2025. Tous droits réservés.
            </p>
        </div>
        """
        
        label = QLabel(about_text)
        label.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(label)
        
        tab.setLayout(layout)
        self.tabs.addTab(tab, "À propos")
    
    def connect_signals(self):
        """Connecte les signaux aux slots."""
        # Boutons
        self.button_choose_directory.clicked.connect(self.choose_directory)
        self.button_compile.clicked.connect(self.compile_files)
        
        # Changements de sélection
        self.list_files.itemSelectionChanged.connect(self.update_selection_count)
        self.checkbox_all_files.stateChanged.connect(self.toggle_all_files)
        
        # Options avancées
        self.checkbox_sort_data.stateChanged.connect(self.toggle_sort_options)
        
        # Option de vérification préliminaire
        self.checkbox_verify_files.stateChanged.connect(self.toggle_verification)
    
    def toggle_verification(self, state):
        """Active ou désactive la vérification préliminaire des fichiers."""
        self.verification_enabled = state == Qt.CheckState.Checked.value
    
    def choose_directory(self):
        """Ouvre un dialogue pour choisir un répertoire et affiche les fichiers Excel et CSV."""
        directory = QFileDialog.getExistingDirectory(self, "Choisir un répertoire")
        if directory:
            self.directory = directory
            self.label_directory.setText(directory)
            self.load_files()
    
    def load_files(self):
        """Charge les fichiers Excel et CSV du répertoire sélectionné."""
        self.list_files.clear()
        
        if self.directory and os.path.isdir(self.directory):
            for file in os.listdir(self.directory):
                if self.checkbox_excel.isChecked() and file.lower().endswith(('.xlsx', '.xls')):
                    self.list_files.addItem(file)
                elif self.checkbox_csv.isChecked() and file.lower().endswith('.csv'):
                    self.list_files.addItem(file)
    
    def toggle_all_files(self, state):
        """Sélectionne ou désélectionne tous les fichiers."""
        for i in range(self.list_files.count()):
            self.list_files.item(i).setSelected(state == Qt.CheckState.Checked.value)
    
    def update_selection_count(self):
        """Met à jour le compteur de fichiers sélectionnés."""
        selected_count = len(self.list_files.selectedItems())
        self.label_file_count.setText(f"{selected_count} fichier(s) sélectionné(s)")
        self.button_compile.setEnabled(selected_count > 0)
    
    def toggle_sort_options(self, state):
        """Active ou désactive les options de tri."""
        self.lineedit_sort_column.setEnabled(state == Qt.CheckState.Checked.value)
    
    def show_ip_warning(self):
        """Affiche l'avertissement de propriété intellectuelle."""
        dialog = IPWarningDialog(self)
        if dialog.exec() == QDialog.DialogCode.Rejected:
            sys.exit(0)
    
    def verify_files(self, files):
        """
        Vérifie la compatibilité des fichiers avant la compilation.
        
        Args:
            files: Liste des noms de fichiers à vérifier
            
        Returns:
            Tuple (compatible_files, incompatible_files)
        """
        compatible_files = []
        incompatible_files = []
        
        header_start_row = self.spinbox_header_start.value()
        header_rows = self.spinbox_header.value()
        
        for file in files:
            file_path = os.path.join(self.directory, file)
            
            try:
                if file.lower().endswith(('.xlsx', '.xls')):
                    is_compatible, message = FileVerification.verify_excel_file(
                        file_path, header_start_row, header_rows
                    )
                elif file.lower().endswith('.csv'):
                    is_compatible, message = FileVerification.verify_csv_file(
                        file_path, header_start_row, header_rows
                    )
                else:
                    is_compatible, message = False, "Format de fichier non pris en charge"
                
                if is_compatible:
                    compatible_files.append(file)
                else:
                    incompatible_files.append((file, message))
                    
            except Exception as e:
                logging.error(f"Erreur lors de la vérification du fichier {file}: {str(e)}")
                incompatible_files.append((file, f"Erreur: {str(e)}"))
        
        return compatible_files, incompatible_files
    
    def compile_files(self):
        """Lance la compilation des fichiers sélectionnés."""
        selected_items = self.list_files.selectedItems()
        
        if not selected_items:
            QMessageBox.warning(self, "Aucun fichier sélectionné",
                                "Veuillez sélectionner au moins un fichier à compiler.")
            return
        
        selected_files = [item.text() for item in selected_items]
        
        # Vérification préliminaire des fichiers si activée
        if self.verification_enabled:
            compatible_files, incompatible_files = self.verify_files(selected_files)
            
            if incompatible_files:
                dialog = VerificationReportDialog(self, compatible_files, incompatible_files)
                result = dialog.exec()
                
                if result == QDialog.DialogCode.Rejected:
                    return
                    
                # Si l'utilisateur a choisi de continuer uniquement avec les fichiers compatibles
                if dialog.continue_with_compatible:
                    selected_files = compatible_files
        
        # Préparation des paramètres
        header_start_row = self.spinbox_header_start.value()
        header_rows = self.spinbox_header.value()
        add_filename = self.checkbox_add_filename.isChecked()
        sort_data = self.checkbox_sort_data.isChecked()
        sort_column = self.get_sort_column_index()
        repeat_headers = self.checkbox_repeat_header.isChecked()
        remove_empty_rows = self.checkbox_remove_empty_rows.isChecked()
        remove_duplicates = self.checkbox_remove_duplicates.isChecked()
        
        # Configuration de la barre de progression
        self.progress_bar.setMaximum(len(selected_files))
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.status_label.setText("Compilation en cours...")
        
        # Création et démarrage du worker de compilation
        self.compilation_worker = CompilationWorker(
            selected_files, self.directory, header_start_row, header_rows,
            add_filename, sort_data, sort_column, repeat_headers, 
            remove_empty_rows, remove_duplicates
        )
        
        self.compilation_worker.progress.connect(self.update_progress)
        self.compilation_worker.error.connect(self.show_error_message)
        self.compilation_worker.finished.connect(self.compilation_finished)
        
        # Désactiver l'interface pendant la compilation
        self.set_ui_enabled(False)
        
        # Démarrer la compilation
        self.compilation_worker.start()
    
    def get_sort_column_index(self):
        """
        Convertit la lettre de colonne Excel en index (0-indexé).
        Ex: A -> 0, B -> 1, etc.
        """
        sort_column = 0
        if self.checkbox_sort_data.isChecked():
            col_name = self.lineedit_sort_column.text().strip().upper()
            if col_name:
                # Conversion de la lettre de colonne Excel en index
                if len(col_name) == 1 and 'A' <= col_name <= 'Z':
                    sort_column = ord(col_name) - ord('A')
                elif len(col_name) > 1:
                    # Gestion des colonnes au-delà de Z (AA, AB, etc.)
                    sort_column = 0
                    for char in col_name:
                        if 'A' <= char <= 'Z':
                            sort_column = sort_column * 26 + (ord(char) - ord('A') + 1)
                    sort_column -= 1  # Ajuster pour 0-indexé
        
        return max(0, sort_column)
    
    def update_progress(self, value):
        """Met à jour la barre de progression."""
        self.progress_bar.setValue(value)
    
    def show_error_message(self, message):
        """Affiche un message d'erreur."""
        QMessageBox.warning(self, "Erreur de compilation", message)
    
    def set_ui_enabled(self, enabled):
        """Active ou désactive l'interface utilisateur."""
        self.button_compile.setEnabled(enabled)
        self.button_choose_directory.setEnabled(enabled)
        self.list_files.setEnabled(enabled)
        self.checkbox_all_files.setEnabled(enabled)
        self.tabs.setTabEnabled(1, enabled)  # Onglet options avancées
    
    def compilation_finished(self, result):
        """
        Gestionnaire appelé lorsque la compilation est terminée.
        
        Args:
            result: Tuple contenant les données compilées
        """
        preliminary_info, headers, combined_data, merged_cells, successful_files, failed_files = result
        
        # Réactiver l'interface
        self.set_ui_enabled(True)
        self.progress_bar.setVisible(False)
        
        # Vérifier s'il y a des données à écrire
        if not combined_data or not headers:
            self.status_label.setText("Échec: Aucune donnée à écrire.")
            QMessageBox.critical(self, "Échec de la compilation", 
                                 "Aucune donnée valide n'a pu être compilée.")
            return
        
        # Obtenir le nom du fichier de sortie
        output_name = self.lineedit_output_name.text().strip()
        if not output_name.lower().endswith('.xlsx'):
            output_name += '.xlsx'
        
        output_path = os.path.join(self.directory, output_name)
        
        try:
            # Création du workbook de sortie
            wb = openpyxl.Workbook()
            ws = wb.active
            
            # Écrire les informations préliminaires
            current_row = ExcelFormatter.write_preliminary_info(ws, preliminary_info)
            
            # Écrire les en-têtes
            header_row = ExcelFormatter.write_headers(ws, headers, current_row)
            
            # Écrire les données
            ExcelFormatter.write_data(ws, combined_data, header_row)
            
            # Appliquer les cellules fusionnées si l'option est activée
            if self.checkbox_merge_headers.isChecked() and merged_cells:
                ExcelFormatter.apply_merged_cells(
                    ws, merged_cells, self.spinbox_header_start.value(), current_row
                )
            
            # Ajuster la largeur des colonnes si l'option est activée
            if self.checkbox_auto_width.isChecked():
                ExcelFormatter.adjust_column_widths(ws)
            
            # Figer les volets si l'option est activée
            if self.checkbox_freeze_header.isChecked():
                ExcelFormatter.freeze_header(ws, header_row)
            
            # Sauvegarder le fichier
            wb.save(output_path)
            
            # Mettre à jour le statut
            self.status_label.setText(f"Compilation terminée. Fichier enregistré : {output_name}")
            
            # Afficher le rapport de compilation
            dialog = CompilationReportDialog(self, successful_files, failed_files, output_path)
            dialog.exec()
            
        except PermissionError:
            self.status_label.setText("Échec: Le fichier de sortie est ouvert dans une autre application.")
            QMessageBox.critical(self, "Erreur de compilation", 
                                 "Le fichier de sortie est ouvert dans une autre application. Veuillez le fermer et réessayer.")
        except Exception as e:
            self.status_label.setText(f"Échec: {str(e)}")
            logging.error(f"Erreur lors de la création du fichier de sortie: {str(e)}")
            logging.error(traceback.format_exc())
            QMessageBox.critical(self, "Erreur de compilation", 
                                 f"Une erreur est survenue lors de la création du fichier de sortie:\n{str(e)}")


# Tests unitaires pour les classes
class TestFileVerification(unittest.TestCase):
    """Tests unitaires pour la classe FileVerification."""
    
    def test_verify_excel_file(self):
        """Teste la vérification des fichiers Excel."""
        # Ce test nécessiterait des fichiers Excel de test
        pass
    
    def test_verify_csv_file(self):
        """Teste la vérification des fichiers CSV."""
        # Ce test nécessiterait des fichiers CSV de test
        pass


class TestCompilationWorker(unittest.TestCase):
    """Tests unitaires pour la classe CompilationWorker."""
    
    def test_sort_data(self):
        """Teste le tri des données."""
        worker = CompilationWorker([], "", 1, 1)
        data = [["B", 2], ["A", 1], ["C", 3]]
        headers = [["Col1", "Col2"]]
        worker.sort_column = 0
        sorted_data = worker._sort_data(data, headers)
        self.assertEqual(sorted_data[0][0], "A")
        self.assertEqual(sorted_data[1][0], "B")
        self.assertEqual(sorted_data[2][0], "C")
    
    def test_remove_duplicate_rows(self):
        """Teste la suppression des doublons."""
        worker = CompilationWorker([], "", 1, 1)
        data = [["A", 1], ["B", 2], ["A", 1], ["C", 3]]
        deduplicated = worker._remove_duplicate_rows(data)
        self.assertEqual(len(deduplicated), 3)

# Point d'entrée de l'application
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # Style moderne
    window = ModernExcelCompilerApp()
    window.show()
    sys.exit(app.exec())