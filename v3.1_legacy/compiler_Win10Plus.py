"""
Excel Compiler Application
Version: 3.1
Auteur: GOUNOU N'GOBI Chabi Zimé (Data Manager, Data Analyst)
Améliorations: Juin 2025 - Support étendu formats Excel et texte

Application pour compiler plusieurs fichiers Excel en un seul fichier avec diverses options de formatage.
Nouvelles fonctionnalités:
- Prévisualisation des données
- Internationalisation complète
- Options de format de date
"""

import sys
import os
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, QTabWidget,
    QGroupBox, QLabel, QSpinBox, QCheckBox, QLineEdit, QPushButton, QListWidget,
    QProgressBar, QMessageBox, QFileDialog, QListWidgetItem, QStyle, QProgressDialog,
    QTableWidget, QTableWidgetItem, QHeaderView, QScrollArea, QDialog, QComboBox,
    QStyledItemDelegate, QStyleOptionButton, QGridLayout, QRadioButton,
    QButtonGroup, QSplitter, QToolBar, QMenu, QMenuBar, QSizePolicy,
    QWizard, QWizardPage, QTextEdit, QFrame
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QSize, QSettings, QTranslator, QLocale, QMetaObject, Q_ARG, QObject
from PyQt6.QtGui import QIcon, QFont, QPalette, QColor, QBrush, QKeySequence, QAction
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment, numbers
from openpyxl.utils import get_column_letter
import openpyxl
import logging
import csv
import json
import unittest
import tempfile
import shutil
from unittest.mock import Mock, patch, MagicMock
from datetime import datetime
from typing import List, Tuple, Dict, Optional, Any, Union, Set, Iterator, Generator
import traceback
import gc
import psutil
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
from collections import defaultdict
import weakref
import signal
import time
from logging.handlers import RotatingFileHandler
from pathlib import Path
import re
import math
import hashlib
import mimetypes
from urllib.parse import unquote
import queue
import platform
from dataclasses import dataclass, field
from enum import Enum
import json
import configparser
from contextlib import contextmanager

# Import du système de monitoring
try:
    from enhanced_monitoring import start_monitoring, stop_monitoring, log_event, increment_counter, set_gauge, time_operation, get_monitoring_report
    MONITORING_AVAILABLE = True
except ImportError:
    MONITORING_AVAILABLE = False
    # Stubs pour éviter les erreurs
    def start_monitoring(): pass
    def stop_monitoring(): pass
    def log_event(*args, **kwargs): pass
    def increment_counter(*args, **kwargs): pass
    def set_gauge(*args, **kwargs): pass
    def time_operation(name): return contextmanager(lambda: (yield))()
    def get_monitoring_report(): return {}

# Constantes
MAX_PREVIEW_ROWS = 200
SUPPORTED_EXCEL_EXTENSIONS = ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xls']
SUPPORTED_TEXT_EXTENSIONS = ['.csv', '.tsv', '.txt']
DEFAULT_ENCODINGS = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']

# Constantes pour les performances
CHUNK_SIZE = 10000  # Nombre de lignes par chunk
MAX_MEMORY_USAGE = 1024 * 1024 * 512  # 512MB max
MAX_THREADS = min(4, os.cpu_count() or 1)  # Limite le nombre de threads

# Constantes de sécurité
MAX_FILE_SIZE = 100 * 1024 * 1024  # 100MB par fichier
MAX_SESSION_SIZE = 1024 * 1024 * 1024  # 1GB par session
MAX_FILENAME_LENGTH = 255  # Limite nom de fichier
MAX_PATH_LENGTH = 4096  # Limite chemin complet
FORBIDDEN_FILENAME_CHARS = ['<', '>', ':', '"', '|', '?', '*', '\0', '\\', '/']
FORBIDDEN_EXTENSIONS = ['.exe', '.bat', '.cmd', '.com', '.scr', '.pif', '.vbs', '.js']
ALLOWED_MIME_TYPES = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',  # .xlsx
    'application/vnd.ms-excel',  # .xls
    'text/csv',  # .csv
    'text/tab-separated-values',  # .tsv
    'text/plain'  # .txt
]

# Délimiteurs supportés pour CSV/TSV/TXT
SUPPORTED_DELIMITERS = {
    'auto': None,  # Détection automatique
    'comma': ',',  # Virgule (CSV standard)
    'semicolon': ';',  # Point-virgule (CSV européen)
    'tab': '\t',  # Tabulation (TSV)
    'pipe': '|',  # Pipe
    'space': ' ',  # Espace
    'colon': ':',  # Deux points
    'custom': None  # Délimiteur personnalisé défini par l'utilisateur
}

# Configuration de responsivité
SCREEN_BREAKPOINTS = {
    'small': 1024,   # Petit écran (tablettes)
    'medium': 1366,  # Écran moyen (laptops)
    'large': 1920,   # Grand écran (desktop)
    'xlarge': 2560   # Très grand écran (4K)
}

RESPONSIVE_FONT_SCALES = {
    'small': 0.85,
    'medium': 1.0,
    'large': 1.1,
    'xlarge': 1.2
}

# Constantes de surveillance et monitoring
HEALTH_CHECK_INTERVAL = 30  # secondes
METRICS_RETENTION_DAYS = 7  # jours
LOG_ROTATION_SIZE = 10 * 1024 * 1024  # 10MB
LOG_BACKUP_COUNT = 5
ALERT_MEMORY_THRESHOLD = 80  # pourcentage
ALERT_CPU_THRESHOLD = 90  # pourcentage
PERFORMANCE_METRICS_BUFFER = 1000  # nombre d'échantillons

# Niveaux d'alerte
class AlertLevel(Enum):
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    CRITICAL = "critical"

# Constantes de robustesse et configuration
CONFIG_FILE_NAME = "excelcompiler_config.json"
BACKUP_INTERVAL = 300  # secondes (5 minutes)
MAX_BACKUP_FILES = 10
AUTO_RECOVERY_ENABLED = True
CRASH_REPORT_DIR = "crash_reports"
VALIDATION_TIMEOUT = 30  # secondes

# Types d'erreurs pour la récupération
class ErrorType(Enum):
    FILE_ACCESS = "file_access"
    MEMORY_ERROR = "memory_error"
    VALIDATION_ERROR = "validation_error"
    NETWORK_ERROR = "network_error"
    PERMISSION_ERROR = "permission_error"
    CORRUPT_DATA = "corrupt_data"
    TIMEOUT_ERROR = "timeout_error"

# Stratégies de récupération
class RecoveryStrategy(Enum):
    RETRY = "retry"
    SKIP = "skip"
    FALLBACK = "fallback"
    USER_CHOICE = "user_choice"
    AUTO_FIX = "auto_fix"

class ThreadSafeGUIHelper:
    """Helper pour les opérations GUI thread-safe"""
    
    @staticmethod
    def is_main_thread() -> bool:
        """Vérifie si on est dans le thread principal"""
        try:
            return QThread.currentThread() == QApplication.instance().thread()
        except (AttributeError, RuntimeError) as e:
            # AttributeError: QApplication pas initialisée
            # RuntimeError: Qt context error
            logging.warning(f"Impossible de vérifier le thread principal: {e}")
            return False
    
    @staticmethod
    def invoke_in_main_thread(obj, method_name: str, *args, **kwargs):
        """Invoke une méthode dans le thread principal de manière thread-safe"""
        if ThreadSafeGUIHelper.is_main_thread():
            # Déjà dans le thread principal, appel direct
            return getattr(obj, method_name)(*args, **kwargs)
        else:
            # Utilisation de QMetaObject.invokeMethod pour thread safety
            if args:
                qt_args = [Q_ARG("QVariant", arg) for arg in args]
                return QMetaObject.invokeMethod(obj, method_name, Qt.ConnectionType.QueuedConnection, *qt_args)
            else:
                return QMetaObject.invokeMethod(obj, method_name, Qt.ConnectionType.QueuedConnection)
    
    @staticmethod
    def assert_main_thread(method_name: str = "unknown"):
        """Assert pour vérifier qu'on est dans le thread principal (debug)"""
        if not ThreadSafeGUIHelper.is_main_thread():
            import traceback
            stack_trace = traceback.format_stack()
            logging.warning(f"GUI method '{method_name}' called from non-main thread!\nStack trace:\n{''.join(stack_trace[-3:])}")
            raise RuntimeError(f"GUI method '{method_name}' must be called from main thread")


class PreviewWorker(QThread):
    """Worker asynchrone pour charger les prévisualisations sans bloquer l'UI"""
    
    preview_loaded = pyqtSignal(str, list, list)  # file_path, headers, data
    preview_error = pyqtSignal(str, str)  # file_path, error_message
    
    def __init__(self, file_path, header_start_row, header_rows, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.header_start_row = header_start_row
        self.header_rows = header_rows
        self.cancelled = False
    
    def cancel(self):
        """Annule le chargement en cours"""
        self.cancelled = True
    
    def run(self):
        """Charge la prévisualisation en arrière-plan"""
        if self.cancelled:
            return
            
        try:
            headers, data = self._load_preview_data()
            if not self.cancelled:
                self.preview_loaded.emit(self.file_path, headers, data)
        except Exception as e:
            if not self.cancelled:
                self.preview_error.emit(self.file_path, str(e))
    
    def _load_preview_data(self):
        """Charge les données de prévisualisation selon le type de fichier"""
        if self.file_path.lower().endswith(('.xlsx', '.xls', '.xlsm', '.xltx', '.xltm')):
            return self._load_excel_preview()
        else:
            return self._load_csv_preview()
    
    def _load_excel_preview(self):
        """Charge la prévisualisation Excel de manière optimisée"""
        wb = openpyxl.load_workbook(self.file_path, data_only=True, read_only=True)
        ws = wb.active
        
        headers = []
        data = []
        
        # Limiter à 100 lignes pour la prévisualisation (plus rapide)
        MAX_PREVIEW_ROWS = 100
        
        # Récupérer les en-têtes
        for row in range(self.header_start_row, self.header_start_row + self.header_rows):
            if self.cancelled:
                break
            header_row = []
            for cell in ws[row]:
                header_row.append(cell.value)
            headers.append(header_row)
        
        # Récupérer les données (limité)
        row_count = 0
        for row in ws.iter_rows(min_row=self.header_start_row + self.header_rows):
            if self.cancelled or row_count >= MAX_PREVIEW_ROWS:
                break
                
            row_data = [cell.value for cell in row]
            data.append(row_data)
            row_count += 1
        
        wb.close()
        return headers, data
    
    def _load_csv_preview(self):
        """Charge la prévisualisation CSV de manière optimisée"""
        encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']
        
        for encoding in encodings:
            if self.cancelled:
                break
                
            try:
                # Lire seulement les premières lignes pour la prévisualisation
                df = pd.read_csv(self.file_path, encoding=encoding, nrows=100)
                
                headers = []
                header_start = self.header_start_row - 1
                header_rows = self.header_rows
                
                for row in range(header_start, min(header_start + header_rows, len(df))):
                    if self.cancelled:
                        break
                    if row < len(df):
                        headers.append(df.iloc[row].tolist())
                
                # Données après les en-têtes
                data = []
                for row in range(header_start + header_rows, len(df)):
                    if self.cancelled:
                        break
                    data.append(df.iloc[row].tolist())
                
                return headers, data
                
            except UnicodeDecodeError:
                continue
        
        raise ValueError("Impossible de décoder le fichier avec les encodages supportés")


class PreviewCache:
    """Cache intelligent pour les prévisualisations de fichiers"""
    
    def __init__(self, max_size=10):
        self.cache = {}
        self.access_times = {}
        self.max_size = max_size
    
    def get(self, file_path, file_mtime):
        """Récupère une prévisualisation en cache si elle est valide"""
        if file_path in self.cache:
            cached_data, cached_mtime = self.cache[file_path]
            if cached_mtime >= file_mtime:
                # Mettre à jour l'heure d'accès
                self.access_times[file_path] = time.time()
                return cached_data
            else:
                # Fichier modifié, supprimer du cache
                self.remove(file_path)
        return None
    
    def put(self, file_path, data, file_mtime):
        """Ajoute une prévisualisation au cache"""
        # Nettoyer le cache si nécessaire
        if len(self.cache) >= self.max_size:
            self._evict_lru()
        
        self.cache[file_path] = (data, file_mtime)
        self.access_times[file_path] = time.time()
    
    def remove(self, file_path):
        """Supprime une entrée du cache"""
        self.cache.pop(file_path, None)
        self.access_times.pop(file_path, None)
    
    def _evict_lru(self):
        """Supprime l'élément le moins récemment utilisé"""
        if not self.access_times:
            return
        
        lru_path = min(self.access_times.keys(), key=lambda k: self.access_times[k])
        self.remove(lru_path)
    
    def clear(self):
        """Vide le cache"""
        self.cache.clear()
        self.access_times.clear()


class TutorialWizard(QWizard):
    """Assistant tutoriel intégré pour guider les nouveaux utilisateurs"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Assistant de démarrage - ExcelCompiler")
        self.setWindowIcon(QIcon("icon.png") if os.path.exists("icon.png") else self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon))
        self.setWizardStyle(QWizard.WizardStyle.ModernStyle)
        self.resize(800, 600)
        
        # Pages du tutoriel
        self.addPage(self.create_welcome_page())
        self.addPage(self.create_files_page())
        self.addPage(self.create_configuration_page())
        self.addPage(self.create_compilation_page())
        self.addPage(self.create_finish_page())
        
        # Styles CSS
        self.setStyleSheet("""
            QWizard {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                                          stop: 0 #f0f0f0, stop: 1 #e0e0e0);
            }
            QWizardPage {
                background: white;
                border-radius: 8px;
                margin: 10px;
            }
            QLabel {
                color: #333;
                font-size: 11pt;
            }
            .title {
                font-size: 16pt;
                font-weight: bold;
                color: #2c3e50;
                margin-bottom: 10px;
            }
            .subtitle {
                font-size: 12pt;
                color: #7f8c8d;
                margin-bottom: 20px;
            }
            .step {
                background: #ecf0f1;
                border-left: 4px solid #3498db;
                padding: 15px;
                margin: 10px 0;
                border-radius: 4px;
            }
            .warning {
                background: #fff3cd;
                border-left: 4px solid #ffc107;
                padding: 15px;
                margin: 10px 0;
                border-radius: 4px;
            }
            .success {
                background: #d4edda;
                border-left: 4px solid #28a745;
                padding: 15px;
                margin: 10px 0;
                border-radius: 4px;
            }
        """)
    
    def create_welcome_page(self):
        """Page de bienvenue"""
        page = QWizardPage()
        page.setTitle("Bienvenue dans ExcelCompiler")
        page.setSubTitle("Cet assistant vous guidera dans l'utilisation de l'application")
        
        layout = QVBoxLayout()
        
        # Logo et titre
        title = QLabel("🎯 ExcelCompiler v3.1")
        title.setObjectName("title")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        
        # Description
        desc = QLabel("""
        ExcelCompiler est un outil puissant pour fusionner plusieurs fichiers Excel et CSV en un seul fichier.
        
        <b>Fonctionnalités principales :</b>
        • Support des formats Excel (.xlsx, .xlsm, .xls) et texte (.csv, .tsv, .txt)
        • Détection automatique des encodages et délimiteurs
        • Options avancées de formatage et de tri
        • Interface responsive adaptée à votre écran
        • Validation en temps réel des paramètres
        • Système d'annulation des opérations
        """)
        desc.setWordWrap(True)
        desc.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(desc)
        
        # Avantages
        advantages = QFrame()
        advantages.setObjectName("step")
        adv_layout = QVBoxLayout(advantages)
        adv_layout.addWidget(QLabel("<b>✨ Pourquoi utiliser ExcelCompiler ?</b>"))
        adv_layout.addWidget(QLabel("• Gain de temps considérable sur les tâches répétitives"))
        adv_layout.addWidget(QLabel("• Interface intuitive et guidée"))
        adv_layout.addWidget(QLabel("• Gestion robuste des erreurs"))
        adv_layout.addWidget(QLabel("• Prévisualisation des données avant compilation"))
        layout.addWidget(advantages)
        
        layout.addStretch()
        page.setLayout(layout)
        return page
    
    def create_files_page(self):
        """Page de sélection des fichiers"""
        page = QWizardPage()
        page.setTitle("Sélection des fichiers")
        page.setSubTitle("Apprenez à sélectionner et gérer vos fichiers")
        
        layout = QVBoxLayout()
        
        # Étape 1
        step1 = QFrame()
        step1.setObjectName("step")
        step1_layout = QVBoxLayout(step1)
        step1_layout.addWidget(QLabel("<b>📁 Étape 1 : Choisir le dossier source</b>"))
        step1_layout.addWidget(QLabel("1. Cliquez sur le bouton 'Parcourir' dans la section 'Répertoire'"))
        step1_layout.addWidget(QLabel("2. Naviguez vers le dossier contenant vos fichiers"))
        step1_layout.addWidget(QLabel("3. Sélectionnez le dossier et validez"))
        layout.addWidget(step1)
        
        # Étape 2
        step2 = QFrame()
        step2.setObjectName("step")
        step2_layout = QVBoxLayout(step2)
        step2_layout.addWidget(QLabel("<b>📋 Étape 2 : Sélectionner les fichiers</b>"))
        step2_layout.addWidget(QLabel("1. La liste des fichiers compatibles s'affiche automatiquement"))
        step2_layout.addWidget(QLabel("2. Cochez les fichiers que vous souhaitez compiler"))
        step2_layout.addWidget(QLabel("3. Utilisez 'Tout sélectionner' ou 'Tout désélectionner' si nécessaire"))
        layout.addWidget(step2)
        
        # Conseils
        tips = QFrame()
        tips.setObjectName("warning")
        tips_layout = QVBoxLayout(tips)
        tips_layout.addWidget(QLabel("<b>💡 Conseils utiles</b>"))
        tips_layout.addWidget(QLabel("• Vérifiez que tous vos fichiers ont la même structure (mêmes colonnes)"))
        tips_layout.addWidget(QLabel("• Les fichiers Excel et CSV peuvent être mélangés"))
        tips_layout.addWidget(QLabel("• L'ordre de sélection n'affecte pas le résultat final"))
        layout.addWidget(tips)
        
        layout.addStretch()
        page.setLayout(layout)
        return page
    
    def create_configuration_page(self):
        """Page de configuration"""
        page = QWizardPage()
        page.setTitle("Configuration de la compilation")
        page.setSubTitle("Paramétrez votre compilation selon vos besoins")
        
        layout = QVBoxLayout()
        
        # En-têtes
        headers = QFrame()
        headers.setObjectName("step")
        headers_layout = QVBoxLayout(headers)
        headers_layout.addWidget(QLabel("<b>📊 Configuration des en-têtes</b>"))
        headers_layout.addWidget(QLabel("• <b>Ligne de début :</b> Ligne où commencent les en-têtes (généralement 1)"))
        headers_layout.addWidget(QLabel("• <b>Nombre de lignes :</b> Combien de ligne(s) forment l'en-tête"))
        headers_layout.addWidget(QLabel("• <b>Répéter les en-têtes :</b> Inclure les en-têtes pour chaque fichier"))
        layout.addWidget(headers)
        
        # Options avancées
        advanced = QFrame()
        advanced.setObjectName("step")
        advanced_layout = QVBoxLayout(advanced)
        advanced_layout.addWidget(QLabel("<b>⚙️ Options avancées</b>"))
        advanced_layout.addWidget(QLabel("• <b>Ajouter nom de fichier :</b> Ajoute une colonne avec le nom du fichier source"))
        advanced_layout.addWidget(QLabel("• <b>Trier les données :</b> Trie par la colonne spécifiée"))
        advanced_layout.addWidget(QLabel("• <b>Supprimer lignes vides :</b> Élimine les lignes sans données"))
        advanced_layout.addWidget(QLabel("• <b>Supprimer doublons :</b> Élimine les lignes identiques"))
        layout.addWidget(advanced)
        
        # Format de date
        date_format = QFrame()
        date_format.setObjectName("step")
        date_format_layout = QVBoxLayout(date_format)
        date_format_layout.addWidget(QLabel("<b>📅 Format de date</b>"))
        date_format_layout.addWidget(QLabel("• <b>Français :</b> JJ/MM/AAAA (recommandé pour la France)"))
        date_format_layout.addWidget(QLabel("• <b>Anglais :</b> MM/JJ/AAAA (format américain)"))
        date_format_layout.addWidget(QLabel("• <b>ISO :</b> AAAA-MM-JJ (standard international)"))
        layout.addWidget(date_format)
        
        # Attention
        warning = QFrame()
        warning.setObjectName("warning")
        warning_layout = QVBoxLayout(warning)
        warning_layout.addWidget(QLabel("<b>⚠️ Points d'attention</b>"))
        warning_layout.addWidget(QLabel("• Vérifiez que la ligne de début correspond à vos fichiers"))
        warning_layout.addWidget(QLabel("• Le tri peut considérablement ralentir le traitement"))
        warning_layout.addWidget(QLabel("• La suppression des doublons compare toutes les colonnes"))
        layout.addWidget(warning)
        
        layout.addStretch()
        page.setLayout(layout)
        return page
    
    def create_compilation_page(self):
        """Page de compilation"""
        page = QWizardPage()
        page.setTitle("Lancement de la compilation")
        page.setSubTitle("Démarrez et suivez le processus de compilation")
        
        layout = QVBoxLayout()
        
        # Avant compilation
        before = QFrame()
        before.setObjectName("step")
        before_layout = QVBoxLayout(before)
        before_layout.addWidget(QLabel("<b>🚀 Avant de lancer la compilation</b>"))
        before_layout.addWidget(QLabel("1. Vérifiez vos paramètres dans l'onglet 'Configuration'"))
        before_layout.addWidget(QLabel("2. Consultez la prévisualisation si nécessaire"))
        before_layout.addWidget(QLabel("3. Assurez-vous d'avoir suffisamment d'espace disque"))
        layout.addWidget(before)
        
        # Pendant compilation
        during = QFrame()
        during.setObjectName("step")
        during_layout = QVBoxLayout(during)
        during_layout.addWidget(QLabel("<b>⏳ Pendant la compilation</b>"))
        during_layout.addWidget(QLabel("• La barre de progression affiche l'avancement"))
        during_layout.addWidget(QLabel("• Le temps estimé restant est calculé automatiquement"))
        during_layout.addWidget(QLabel("• Vous pouvez annuler à tout moment avec le bouton 'Annuler'"))
        during_layout.addWidget(QLabel("• L'interface reste réactive pendant le traitement"))
        layout.addWidget(during)
        
        # Après compilation
        after = QFrame()
        after.setObjectName("success")
        after_layout = QVBoxLayout(after)
        after_layout.addWidget(QLabel("<b>✅ Après la compilation</b>"))
        after_layout.addWidget(QLabel("• Le fichier de sortie est automatiquement ouvert"))
        after_layout.addWidget(QLabel("• Un résumé des opérations est affiché"))
        after_layout.addWidget(QLabel("• Les erreurs éventuelles sont détaillées"))
        after_layout.addWidget(QLabel("• Vous pouvez sauvegarder le fichier où vous le souhaitez"))
        layout.addWidget(after)
        
        # Dépannage
        troubleshooting = QFrame()
        troubleshooting.setObjectName("warning")
        troubleshooting_layout = QVBoxLayout(troubleshooting)
        troubleshooting_layout.addWidget(QLabel("<b>🔧 En cas de problème</b>"))
        troubleshooting_layout.addWidget(QLabel("• Vérifiez que les fichiers ne sont pas ouverts dans Excel"))
        troubleshooting_layout.addWidget(QLabel("• Consultez les logs d'erreur pour plus de détails"))
        troubleshooting_layout.addWidget(QLabel("• Réduisez le nombre de fichiers si la mémoire est insuffisante"))
        layout.addWidget(troubleshooting)
        
        layout.addStretch()
        page.setLayout(layout)
        return page
    
    def create_finish_page(self):
        """Page de fin"""
        page = QWizardPage()
        page.setTitle("Félicitations !")
        page.setSubTitle("Vous êtes maintenant prêt à utiliser ExcelCompiler")
        
        layout = QVBoxLayout()
        
        # Félicitations
        congrats = QLabel("🎉 Vous avez terminé le tutoriel !")
        congrats.setObjectName("title")
        congrats.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(congrats)
        
        # Récapitulatif
        summary = QFrame()
        summary.setObjectName("success")
        summary_layout = QVBoxLayout(summary)
        summary_layout.addWidget(QLabel("<b>📋 Ce que vous avez appris :</b>"))
        summary_layout.addWidget(QLabel("✓ Comment sélectionner vos fichiers"))
        summary_layout.addWidget(QLabel("✓ Comment configurer la compilation"))
        summary_layout.addWidget(QLabel("✓ Comment lancer et suivre le processus"))
        summary_layout.addWidget(QLabel("✓ Comment gérer les problèmes courants"))
        layout.addWidget(summary)
        
        # Conseils finaux
        final_tips = QFrame()
        final_tips.setObjectName("step")
        final_tips_layout = QVBoxLayout(final_tips)
        final_tips_layout.addWidget(QLabel("<b>💡 Conseils pour bien commencer :</b>"))
        final_tips_layout.addWidget(QLabel("• Commencez par des petits tests avec quelques fichiers"))
        final_tips_layout.addWidget(QLabel("• Gardez des sauvegardes de vos fichiers originaux"))
        final_tips_layout.addWidget(QLabel("• Utilisez la prévisualisation pour vérifier la structure"))
        final_tips_layout.addWidget(QLabel("• N'hésitez pas à relancer ce tutoriel si besoin"))
        layout.addWidget(final_tips)
        
        # Ressources
        resources = QFrame()
        resources.setObjectName("step")
        resources_layout = QVBoxLayout(resources)
        resources_layout.addWidget(QLabel("<b>📚 Ressources utiles :</b>"))
        resources_layout.addWidget(QLabel("• Menu Aide > Guide utilisateur"))
        resources_layout.addWidget(QLabel("• Menu Aide > À propos (informations version)"))
        resources_layout.addWidget(QLabel("• Logs d'erreur dans le dossier de l'application"))
        layout.addWidget(resources)
        
        # Checkbox pour ne plus afficher
        self.dont_show_again = QCheckBox("Ne plus afficher ce tutoriel au démarrage")
        layout.addWidget(self.dont_show_again)
        
        layout.addStretch()
        page.setLayout(layout)
        return page
    
    def accept(self):
        """Gère la fermeture du wizard"""
        # Sauvegarder la préférence de l'utilisateur
        if hasattr(self, 'dont_show_again') and self.dont_show_again.isChecked():
            settings = QSettings('ExcelCompiler', 'ExcelCompiler')
            settings.setValue('tutorial/show_on_startup', False)
        
        super().accept()


class TutorialManager:
    """Gestionnaire pour l'assistant tutoriel"""
    
    @staticmethod
    def should_show_tutorial() -> bool:
        """Détermine si le tutoriel doit être affiché"""
        settings = QSettings('ExcelCompiler', 'ExcelCompiler')
        return settings.value('tutorial/show_on_startup', True, type=bool)
    
    @staticmethod
    def show_tutorial(parent=None):
        """Affiche l'assistant tutoriel"""
        wizard = TutorialWizard(parent)
        return wizard.exec()
    
    @staticmethod
    def reset_tutorial_settings():
        """Remet les paramètres du tutoriel à zéro"""
        settings = QSettings('ExcelCompiler', 'ExcelCompiler')
        settings.setValue('tutorial/show_on_startup', True)


class SecurityError(Exception):
    """Exception levée pour les violations de sécurité"""
    pass


class SecurityManager:
    """Gestionnaire de sécurité pour validation des fichiers et chemins"""
    
    def __init__(self):
        self.session_size = 0
        self.processed_files = set()
        
    def reset_session(self):
        """Remet à zéro les compteurs de session"""
        self.session_size = 0
        self.processed_files.clear()
    
    def sanitize_file_path(self, file_path: str) -> str:
        """
        Sécurise un chemin de fichier contre les attaques path traversal
        
        Args:
            file_path: Chemin du fichier à sécuriser
            
        Returns:
            Chemin sécurisé
            
        Raises:
            SecurityError: Si le chemin est dangereux
        """
        if not file_path:
            raise SecurityError("Chemin de fichier vide")
        
        # Décoder les caractères URL-encodés
        try:
            decoded_path = unquote(file_path)
        except Exception:
            raise SecurityError("Chemin de fichier invalide (encodage)")
        
        # Normaliser le chemin
        normalized_path = os.path.normpath(decoded_path)
        
        # Vérifier la longueur
        if len(normalized_path) > MAX_PATH_LENGTH:
            raise SecurityError(f"Chemin trop long (max {MAX_PATH_LENGTH} caractères)")
        
        # Détecter les tentatives de path traversal
        if ".." in normalized_path.split(os.sep):
            raise SecurityError("Tentative de path traversal détectée")
        
        # Vérifier que c'est un chemin absolu ou relatif sécurisé
        if normalized_path.startswith(os.sep) and not self._is_safe_absolute_path(normalized_path):
            raise SecurityError("Chemin absolu non autorisé")
        
        # Vérifier les caractères interdits dans le chemin (en tenant compte de Windows)
        forbidden_chars = ['<', '>', '"', '|', '?', '*', '\0']
        
        # Pour Windows, autoriser ':' uniquement après une lettre de lecteur
        if ':' in normalized_path:
            # Vérifier si c'est un chemin Windows valide (ex: C:\ ou C:/)
            import re
            if not re.match(r'^[A-Za-z]:[/\\]', normalized_path):
                forbidden_chars.append(':')
        
        if any(char in normalized_path for char in forbidden_chars):
            raise SecurityError("Caractères interdits dans le chemin")
        
        return normalized_path
    
    def _is_safe_absolute_path(self, path: str) -> bool:
        """Vérifie si un chemin absolu est sécurisé"""
        # Autoriser seulement les chemins dans certains répertoires
        safe_prefixes = [
            os.path.expanduser("~"),  # Répertoire utilisateur
            tempfile.gettempdir(),    # Répertoire temporaire
        ]
        
        # Ajouter le répertoire de l'application
        try:
            app_dir = os.path.dirname(os.path.abspath(__file__))
            safe_prefixes.append(app_dir)
        except (OSError, ValueError, AttributeError) as e:
            # OSError: Problème accès fichier système
            # ValueError: Chemin invalide
            # AttributeError: __file__ non disponible
            logging.warning(f"Impossible de déterminer le répertoire de l'application: {e}")
            # Ne pas ajouter de répertoire par défaut pour sécurité
        
        return any(path.startswith(prefix) for prefix in safe_prefixes)
    
    def validate_filename(self, filename: str) -> bool:
        """
        Valide un nom de fichier
        
        Args:
            filename: Nom du fichier à valider
            
        Returns:
            True si le nom est valide
            
        Raises:
            SecurityError: Si le nom est dangereux
        """
        if not filename:
            raise SecurityError("Nom de fichier vide")
        
        # Vérifier la longueur
        if len(filename) > MAX_FILENAME_LENGTH:
            raise SecurityError(f"Nom de fichier trop long (max {MAX_FILENAME_LENGTH} caractères)")
        
        # Vérifier les caractères interdits
        for char in FORBIDDEN_FILENAME_CHARS:
            if char in filename:
                raise SecurityError(f"Caractère interdit '{char}' dans le nom de fichier")
        
        # Vérifier les noms réservés Windows
        reserved_names = [
            'CON', 'PRN', 'AUX', 'NUL',
            'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
            'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'
        ]
        
        base_name = os.path.splitext(filename)[0].upper()
        if base_name in reserved_names:
            raise SecurityError(f"Nom de fichier réservé: {filename}")
        
        # Vérifier l'extension
        _, ext = os.path.splitext(filename)
        if ext.lower() in FORBIDDEN_EXTENSIONS:
            raise SecurityError(f"Extension de fichier interdite: {ext}")
        
        return True
    
    def check_file_size(self, file_path: str) -> bool:
        """
        Vérifie la taille d'un fichier
        
        Args:
            file_path: Chemin du fichier
            
        Returns:
            True si la taille est acceptable
            
        Raises:
            SecurityError: Si le fichier est trop gros
        """
        try:
            file_size = os.path.getsize(file_path)
        except OSError as e:
            raise SecurityError(f"Impossible de lire la taille du fichier: {e}")
        
        # Vérifier la taille individuelle
        if file_size > MAX_FILE_SIZE:
            raise SecurityError(f"Fichier trop volumineux: {file_size / (1024*1024):.1f}MB (max {MAX_FILE_SIZE / (1024*1024):.0f}MB)")
        
        # Vérifier la taille de session
        if self.session_size + file_size > MAX_SESSION_SIZE:
            raise SecurityError(f"Taille totale de session dépassée (max {MAX_SESSION_SIZE / (1024*1024*1024):.0f}GB)")
        
        return True
    
    def verify_file_integrity(self, file_path: str) -> Dict[str, Any]:
        """
        Vérifie l'intégrité d'un fichier
        
        Args:
            file_path: Chemin du fichier
            
        Returns:
            Dict avec informations d'intégrité
            
        Raises:
            SecurityError: Si le fichier est corrompu ou dangereux
        """
        try:
            # Vérifier que le fichier existe et est lisible
            if not os.path.isfile(file_path):
                raise SecurityError("Le fichier n'existe pas ou n'est pas un fichier régulier")
            
            # Calculer le checksum SHA-256
            sha256_hash = hashlib.sha256()
            with open(file_path, "rb") as f:
                for chunk in iter(lambda: f.read(4096), b""):
                    sha256_hash.update(chunk)
            
            file_hash = sha256_hash.hexdigest()
            
            # Détecter le type MIME
            mime_type, _ = mimetypes.guess_type(file_path)
            
            # Vérifier que le type MIME est autorisé
            if mime_type and mime_type not in ALLOWED_MIME_TYPES:
                # Autoriser les types indéterminés pour les CSV/TXT
                _, ext = os.path.splitext(file_path)
                if ext.lower() not in ['.csv', '.tsv', '.txt']:
                    raise SecurityError(f"Type de fichier non autorisé: {mime_type}")
            
            # Vérification basique de la structure pour les fichiers texte
            if mime_type and mime_type.startswith('text/'):
                self._verify_text_file_structure(file_path)
            
            return {
                'path': file_path,
                'size': os.path.getsize(file_path),
                'sha256': file_hash,
                'mime_type': mime_type,
                'is_safe': True,
                'verified_at': datetime.now().isoformat()
            }
            
        except Exception as e:
            if isinstance(e, SecurityError):
                raise
            raise SecurityError(f"Erreur lors de la vérification d'intégrité: {e}")
    
    def _verify_text_file_structure(self, file_path: str):
        """Vérifie la structure d'un fichier texte"""
        try:
            # Lire les premiers 1024 caractères pour vérification
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                sample = f.read(1024)
            
            # Vérifier qu'il n'y a pas de contenu suspect
            suspicious_patterns = [
                r'<script', r'javascript:', r'vbscript:', r'onload=', r'onerror=',
                r'eval\(', r'exec\(', r'system\(', r'__import__'
            ]
            
            for pattern in suspicious_patterns:
                if re.search(pattern, sample, re.IGNORECASE):
                    raise SecurityError(f"Contenu suspect détecté dans le fichier")
                    
        except UnicodeDecodeError:
            # Si le fichier n'est pas en UTF-8, essayer d'autres encodages
            for encoding in ['latin-1', 'cp1252']:
                try:
                    with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                        sample = f.read(1024)
                    break
                except (OSError, IOError, UnicodeDecodeError) as e:
                    # OSError/IOError: Problème accès fichier
                    # UnicodeDecodeError: Encore un problème d'encodage
                    logging.debug(f"Échec encodage {encoding} pour {file_path}: {e}")
                    continue
            else:
                raise SecurityError("Impossible de décoder le fichier texte")
    
    def add_processed_file(self, file_path: str, file_size: int):
        """Ajoute un fichier traité au compteur de session"""
        self.processed_files.add(file_path)
        self.session_size += file_size
    
    def get_session_stats(self) -> Dict[str, Any]:
        """Retourne les statistiques de la session actuelle"""
        return {
            'files_count': len(self.processed_files),
            'total_size': self.session_size,
            'size_percentage': (self.session_size / MAX_SESSION_SIZE) * 100,
            'remaining_size': MAX_SESSION_SIZE - self.session_size
        }


class SecurityValidator:
    """Validateur de sécurité pour l'interface utilisateur"""
    
    def __init__(self):
        self.security_manager = SecurityManager()
    
    def validate_directory_path(self, directory_path: str) -> Tuple[bool, str]:
        """
        Valide un chemin de répertoire
        
        Returns:
            Tuple (is_valid, error_message)
        """
        try:
            if not directory_path:
                return False, "Aucun répertoire spécifié"
            
            # Sécuriser le chemin
            safe_path = self.security_manager.sanitize_file_path(directory_path)
            
            # Vérifier que c'est un répertoire existant
            if not os.path.isdir(safe_path):
                return False, "Le répertoire n'existe pas"
            
            # Vérifier les permissions de lecture
            if not os.access(safe_path, os.R_OK):
                return False, "Pas d'autorisation de lecture sur ce répertoire"
            
            return True, ""
            
        except SecurityError as e:
            return False, f"Erreur de sécurité: {e}"
        except Exception as e:
            return False, f"Erreur: {e}"
    
    def validate_file_selection(self, files: List[str], directory: str) -> Tuple[bool, str, List[Dict]]:
        """
        Valide une sélection de fichiers
        
        Returns:
            Tuple (is_valid, error_message, file_details)
        """
        try:
            if not files:
                return False, "Aucun fichier sélectionné", []
            
            # Réinitialiser la session
            self.security_manager.reset_session()
            
            file_details = []
            total_size = 0
            
            for filename in files:
                # Valider le nom de fichier
                self.security_manager.validate_filename(filename)
                
                # Construire le chemin complet
                file_path = os.path.join(directory, filename)
                safe_path = self.security_manager.sanitize_file_path(file_path)
                
                # Vérifier la taille
                self.security_manager.check_file_size(safe_path)
                
                # Vérifier l'intégrité
                integrity_info = self.security_manager.verify_file_integrity(safe_path)
                
                file_size = integrity_info['size']
                total_size += file_size
                
                file_details.append({
                    'filename': filename,
                    'path': safe_path,
                    'size': file_size,
                    'size_mb': file_size / (1024 * 1024),
                    'sha256': integrity_info['sha256'],
                    'mime_type': integrity_info['mime_type'],
                    'is_safe': True
                })
                
                # Ajouter au compteur de session
                self.security_manager.add_processed_file(safe_path, file_size)
            
            return True, "", file_details
            
        except SecurityError as e:
            return False, f"Erreur de sécurité: {e}", []
        except Exception as e:
            return False, f"Erreur: {e}", []


@dataclass
class HealthCheckResult:
    """Résultat d'un health check"""
    timestamp: str
    component: str
    status: str  # 'healthy', 'warning', 'critical'
    message: str
    metrics: Dict[str, Any] = field(default_factory=dict)


@dataclass
class PerformanceMetric:
    """Métrique de performance"""
    timestamp: str
    metric_name: str
    value: float
    unit: str
    tags: Dict[str, str] = field(default_factory=dict)


@dataclass
class Alert:
    """Alerte système"""
    timestamp: str
    level: AlertLevel
    component: str
    message: str
    resolved: bool = False
    resolved_at: Optional[str] = None


class HealthMonitor:
    """Moniteur de santé système pour surveillance continue"""
    
    def __init__(self):
        self.health_checks = {}
        self.last_check_time = time.time()
        self.alert_queue = queue.Queue()
        self.is_monitoring = False
        self.monitoring_thread = None
        
    def register_health_check(self, name: str, check_function: callable, interval: int = HEALTH_CHECK_INTERVAL):
        """Enregistre un health check"""
        self.health_checks[name] = {
            'function': check_function,
            'interval': interval,
            'last_run': 0,
            'last_result': None
        }
        
    def start_monitoring(self):
        """Démarre la surveillance continue"""
        if self.is_monitoring:
            return
            
        self.is_monitoring = True
        self.monitoring_thread = threading.Thread(target=self._monitoring_loop, daemon=True)
        self.monitoring_thread.start()
        
    def stop_monitoring(self):
        """Arrête la surveillance"""
        self.is_monitoring = False
        if self.monitoring_thread:
            self.monitoring_thread.join(timeout=1.0)
            
    def _monitoring_loop(self):
        """Boucle principale de surveillance"""
        while self.is_monitoring:
            try:
                current_time = time.time()
                
                for name, check_info in self.health_checks.items():
                    if current_time - check_info['last_run'] >= check_info['interval']:
                        try:
                            result = check_info['function']()
                            check_info['last_result'] = result
                            check_info['last_run'] = current_time
                            
                            # Générer des alertes si nécessaire
                            if result.status in ['warning', 'critical']:
                                alert = Alert(
                                    timestamp=result.timestamp,
                                    level=AlertLevel.WARNING if result.status == 'warning' else AlertLevel.CRITICAL,
                                    component=result.component,
                                    message=result.message
                                )
                                self.alert_queue.put(alert)
                                
                        except Exception as e:
                            logging.error(f"Erreur dans health check '{name}': {e}")
                            
                time.sleep(1)  # Attendre 1 seconde avant la prochaine vérification
                
            except Exception as e:
                logging.error(f"Erreur dans la boucle de surveillance: {e}")
                time.sleep(5)
                
    def get_health_status(self) -> Dict[str, Any]:
        """Retourne l'état de santé actuel"""
        status = {
            'overall_status': 'healthy',
            'checks': {},
            'last_update': datetime.now().isoformat()
        }
        
        critical_count = 0
        warning_count = 0
        
        for name, check_info in self.health_checks.items():
            result = check_info['last_result']
            if result:
                status['checks'][name] = {
                    'status': result.status,
                    'message': result.message,
                    'last_check': result.timestamp,
                    'metrics': result.metrics
                }
                
                if result.status == 'critical':
                    critical_count += 1
                elif result.status == 'warning':
                    warning_count += 1
        
        # Déterminer le statut global
        if critical_count > 0:
            status['overall_status'] = 'critical'
        elif warning_count > 0:
            status['overall_status'] = 'warning'
            
        return status
        
    def get_recent_alerts(self, limit: int = 10) -> List[Alert]:
        """Récupère les alertes récentes"""
        alerts = []
        temp_queue = queue.Queue()
        
        # Extraire les alertes de la queue
        while not self.alert_queue.empty() and len(alerts) < limit:
            try:
                alert = self.alert_queue.get_nowait()
                alerts.append(alert)
                temp_queue.put(alert)
            except queue.Empty:
                break
                
        # Remettre les alertes dans la queue
        while not temp_queue.empty():
            self.alert_queue.put(temp_queue.get_nowait())
            
        return alerts[-limit:]  # Retourner les plus récentes


class PerformanceMonitor:
    """Moniteur de performance avec métriques en temps réel"""
    
    def __init__(self):
        self.metrics_buffer = []
        self.start_time = time.time()
        self.operation_counts = defaultdict(int)
        self.operation_durations = defaultdict(list)
        
    def record_metric(self, name: str, value: float, unit: str = "", tags: Dict[str, str] = None):
        """Enregistre une métrique de performance"""
        metric = PerformanceMetric(
            timestamp=datetime.now().isoformat(),
            metric_name=name,
            value=value,
            unit=unit,
            tags=tags or {}
        )
        
        self.metrics_buffer.append(metric)
        
        # Limiter la taille du buffer
        if len(self.metrics_buffer) > PERFORMANCE_METRICS_BUFFER:
            self.metrics_buffer = self.metrics_buffer[-PERFORMANCE_METRICS_BUFFER:]
            
    def start_operation(self, operation_name: str) -> str:
        """Démarre le chronométrage d'une opération"""
        operation_id = f"{operation_name}_{time.time()}"
        self.operation_counts[operation_name] += 1
        return operation_id
        
    def end_operation(self, operation_name: str, operation_id: str):
        """Termine le chronométrage d'une opération"""
        start_time = float(operation_id.split('_')[-1])
        duration = time.time() - start_time
        
        self.operation_durations[operation_name].append(duration)
        self.record_metric(f"{operation_name}_duration", duration, "seconds")
        
        # Nettoyer les anciennes mesures (garder seulement les 100 dernières)
        if len(self.operation_durations[operation_name]) > 100:
            self.operation_durations[operation_name] = self.operation_durations[operation_name][-100:]
            
    def get_system_metrics(self) -> Dict[str, Any]:
        """Récupère les métriques système actuelles"""
        try:
            # Métriques mémoire
            memory = psutil.virtual_memory()
            
            # Métriques CPU
            cpu_percent = psutil.cpu_percent(interval=1)
            
            # Métriques disque
            disk = psutil.disk_usage('/')
            
            # Métriques réseau (si applicable)
            net_io = psutil.net_io_counters()
            
            return {
                'memory': {
                    'total': memory.total,
                    'available': memory.available,
                    'percent': memory.percent,
                    'used': memory.used
                },
                'cpu': {
                    'percent': cpu_percent,
                    'count': psutil.cpu_count()
                },
                'disk': {
                    'total': disk.total,
                    'used': disk.used,
                    'free': disk.free,
                    'percent': (disk.used / disk.total) * 100
                },
                'network': {
                    'bytes_sent': net_io.bytes_sent,
                    'bytes_recv': net_io.bytes_recv
                },
                'uptime': time.time() - self.start_time
            }
        except Exception as e:
            logging.error(f"Erreur lors de la récupération des métriques système: {e}")
            return {}
            
    def get_performance_summary(self) -> Dict[str, Any]:
        """Retourne un résumé des performances"""
        summary = {
            'operations': {},
            'recent_metrics': self.metrics_buffer[-50:] if self.metrics_buffer else [],
            'system_metrics': self.get_system_metrics()
        }
        
        # Statistiques des opérations
        for operation, durations in self.operation_durations.items():
            if durations:
                summary['operations'][operation] = {
                    'count': self.operation_counts[operation],
                    'avg_duration': sum(durations) / len(durations),
                    'min_duration': min(durations),
                    'max_duration': max(durations),
                    'last_duration': durations[-1] if durations else 0
                }
                
        return summary


class StructuredLogger:
    """Logger structuré avec rotation et formatage JSON"""
    
    def __init__(self, name: str = "ExcelCompiler", log_dir: str = "logs"):
        self.name = name
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(exist_ok=True)
        
        # Configuration du logger principal
        self.logger = logging.getLogger(name)
        self.logger.setLevel(logging.INFO)
        
        # Éviter les doublons
        if not self.logger.handlers:
            self._setup_handlers()
            
        self.session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        
    def _setup_handlers(self):
        """Configure les handlers de logging"""
        
        # Handler pour fichier avec rotation
        log_file = self.log_dir / f"{self.name.lower()}.log"
        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=LOG_ROTATION_SIZE,
            backupCount=LOG_BACKUP_COUNT,
            encoding='utf-8'
        )
        
        # Handler pour erreurs séparées
        error_file = self.log_dir / f"{self.name.lower()}_errors.log"
        error_handler = RotatingFileHandler(
            error_file,
            maxBytes=LOG_ROTATION_SIZE,
            backupCount=LOG_BACKUP_COUNT,
            encoding='utf-8'
        )
        error_handler.setLevel(logging.ERROR)
        
        # Formateur structuré
        formatter = logging.Formatter(
            '%(asctime)s | %(levelname)s | %(name)s | %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        file_handler.setFormatter(formatter)
        error_handler.setFormatter(formatter)
        
        self.logger.addHandler(file_handler)
        self.logger.addHandler(error_handler)
        
    def log_structured(self, level: str, event: str, **kwargs):
        """Log un événement structuré"""
        log_entry = {
            'session_id': self.session_id,
            'event': event,
            'timestamp': datetime.now().isoformat(),
            **kwargs
        }
        
        message = json.dumps(log_entry, ensure_ascii=False, separators=(',', ':'))
        
        if level.upper() == 'DEBUG':
            logging.debug(message)
        elif level.upper() == 'INFO':
            logging.info(message)
        elif level.upper() == 'WARNING':
            logging.warning(message)
        elif level.upper() == 'ERROR':
            logging.error(message)
        elif level.upper() == 'CRITICAL':
            logging.critical(message)
            
    def log_operation(self, operation: str, status: str, duration: float = None, **kwargs):
        """Log une opération avec ses métriques"""
        self.log_structured(
            'INFO',
            'operation_completed',
            operation=operation,
            status=status,
            duration_seconds=duration,
            **kwargs
        )
        
    def log_error(self, error: Exception, context: str = "", **kwargs):
        """Log une erreur avec son contexte"""
        self.log_structured(
            'ERROR',
            'error_occurred',
            error_type=type(error).__name__,
            error_message=str(error),
            context=context,
            traceback=traceback.format_exc(),
            **kwargs
        )
        
    def log_security_event(self, event_type: str, severity: str, details: Dict[str, Any]):
        """Log un événement de sécurité"""
        self.log_structured(
            'WARNING' if severity == 'warning' else 'ERROR',
            'security_event',
            event_type=event_type,
            severity=severity,
            **details
        )


@dataclass
class ConfigurationSchema:
    """Schéma de configuration de l'application"""
    # Paramètres généraux
    auto_save_enabled: bool = True
    auto_save_interval: int = 300  # secondes
    backup_enabled: bool = True
    max_backup_files: int = 10
    
    # Paramètres de performance
    max_memory_usage: int = 512 * 1024 * 1024  # 512MB
    max_threads: int = 4
    chunk_size: int = 10000
    
    # Paramètres de sécurité
    security_level: str = "high"  # low, medium, high
    validate_file_integrity: bool = True
    max_file_size: int = 100 * 1024 * 1024  # 100MB
    
    # Paramètres de l'interface
    language: str = "fr"
    theme: str = "default"
    show_preview: bool = True
    enable_validation: bool = True
    
    # Paramètres de surveillance
    health_checks_enabled: bool = True
    logging_level: str = "INFO"
    metrics_retention_days: int = 7
    
    # Paramètres de récupération
    auto_recovery_enabled: bool = True
    crash_reporting: bool = True
    recovery_timeout: int = 30


class ConfigurationManager:
    """Gestionnaire de configuration externalisée avec validation"""
    
    def __init__(self, config_file: str = CONFIG_FILE_NAME):
        self.config_file = Path(config_file)
        self.config = ConfigurationSchema()
        self._backup_config = None
        
    def load_configuration(self) -> ConfigurationSchema:
        """Charge la configuration depuis le fichier"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                
                # Valider et appliquer la configuration
                validated_config = self._validate_configuration(config_data)
                self.config = ConfigurationSchema(**validated_config)
                
                logging.info(f"Configuration chargée depuis {self.config_file}")
            else:
                # Créer la configuration par défaut
                self.save_configuration()
                logging.info("Configuration par défaut créée")
                
        except Exception as e:
            logging.error(f"Erreur lors du chargement de la configuration: {e}")
            # Utiliser la configuration par défaut en cas d'erreur
            self.config = ConfigurationSchema()
            
        return self.config
    
    def save_configuration(self) -> bool:
        """Sauvegarde la configuration dans le fichier"""
        try:
            # Créer le répertoire si nécessaire
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            
            # Sauvegarder avec indentation pour lisibilité
            config_dict = self._config_to_dict()
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_dict, f, indent=4, ensure_ascii=False)
            
            logging.info(f"Configuration sauvegardée dans {self.config_file}")
            return True
            
        except Exception as e:
            logging.error(f"Erreur lors de la sauvegarde de la configuration: {e}")
            return False
    
    def _validate_configuration(self, config_data: Dict[str, Any]) -> Dict[str, Any]:
        """Valide les données de configuration"""
        validated = {}
        defaults = ConfigurationSchema()
        
        # Validation avec valeurs par défaut
        for field_name, default_value in defaults.__dict__.items():
            if field_name in config_data:
                try:
                    # Validation du type
                    if isinstance(default_value, bool):
                        validated[field_name] = bool(config_data[field_name])
                    elif isinstance(default_value, int):
                        validated[field_name] = int(config_data[field_name])
                    elif isinstance(default_value, str):
                        validated[field_name] = str(config_data[field_name])
                    else:
                        validated[field_name] = config_data[field_name]
                        
                    # Validation des valeurs spécifiques
                    validated[field_name] = self._validate_field_value(field_name, validated[field_name])
                    
                except (ValueError, TypeError):
                    logging.warning(f"Valeur invalide pour {field_name}, utilisation de la valeur par défaut")
                    validated[field_name] = default_value
            else:
                validated[field_name] = default_value
        
        return validated
    
    def _validate_field_value(self, field_name: str, value: Any) -> Any:
        """Valide la valeur d'un champ spécifique"""
        # Validation des langues supportées
        if field_name == "language" and value not in ["fr", "en", "es", "de"]:
            return "fr"
        
        # Validation du niveau de sécurité
        if field_name == "security_level" and value not in ["low", "medium", "high"]:
            return "high"
        
        # Validation des niveaux de logging
        if field_name == "logging_level" and value not in ["DEBUG", "INFO", "WARNING", "ERROR"]:
            return "INFO"
        
        # Validation des valeurs numériques positives
        if field_name in ["auto_save_interval", "max_backup_files", "max_threads", "chunk_size", "recovery_timeout"]:
            return max(1, value)
        
        # Validation des tailles en bytes
        if field_name in ["max_memory_usage", "max_file_size"]:
            return max(1024 * 1024, value)  # Minimum 1MB
        
        return value
    
    def _config_to_dict(self) -> Dict[str, Any]:
        """Convertit la configuration en dictionnaire"""
        return {
            field_name: getattr(self.config, field_name)
            for field_name in self.config.__dict__
        }
    
    def get_setting(self, key: str, default=None):
        """Récupère une valeur de configuration spécifique"""
        return getattr(self.config, key, default)
    
    def set_setting(self, key: str, value: Any):
        """Définit une valeur de configuration"""
        if hasattr(self.config, key):
            setattr(self.config, key, value)
        else:
            logging.warning(f"Clé de configuration inconnue: {key}")
    
    def create_backup(self) -> bool:
        """Crée une sauvegarde de la configuration actuelle"""
        try:
            backup_file = self.config_file.with_suffix('.backup')
            if self.config_file.exists():
                shutil.copy2(self.config_file, backup_file)
                return True
        except Exception as e:
            logging.error(f"Erreur lors de la création de la sauvegarde de configuration: {e}")
        return False
    
    def restore_backup(self) -> bool:
        """Restaure la configuration depuis la sauvegarde"""
        try:
            backup_file = self.config_file.with_suffix('.backup')
            if backup_file.exists():
                shutil.copy2(backup_file, self.config_file)
                self.load_configuration()
                return True
        except Exception as e:
            logging.error(f"Erreur lors de la restauration de la configuration: {e}")
        return False


class AdvancedErrorHandler:
    """Gestionnaire d'erreurs avancé avec stratégies de récupération"""
    
    def __init__(self):
        self.error_history = []
        self.recovery_strategies = {
            ErrorType.FILE_ACCESS: [RecoveryStrategy.RETRY, RecoveryStrategy.USER_CHOICE],
            ErrorType.MEMORY_ERROR: [RecoveryStrategy.AUTO_FIX, RecoveryStrategy.RETRY],
            ErrorType.VALIDATION_ERROR: [RecoveryStrategy.AUTO_FIX, RecoveryStrategy.SKIP],
            ErrorType.PERMISSION_ERROR: [RecoveryStrategy.USER_CHOICE, RecoveryStrategy.FALLBACK],
            ErrorType.CORRUPT_DATA: [RecoveryStrategy.AUTO_FIX, RecoveryStrategy.SKIP],
            ErrorType.TIMEOUT_ERROR: [RecoveryStrategy.RETRY, RecoveryStrategy.SKIP],
        }
        
    def handle_error(self, error: Exception, context: Dict[str, Any] = None) -> Tuple[bool, str, Any]:
        """
        Gère une erreur avec stratégie de récupération
        
        Returns:
            Tuple (success, message, recovery_data)
        """
        error_type = self._classify_error(error)
        context = context or {}
        
        # Enregistrer l'erreur
        error_record = {
            'timestamp': datetime.now().isoformat(),
            'error_type': error_type.value,
            'error_message': str(error),
            'context': context,
            'recovery_attempted': False
        }
        self.error_history.append(error_record)
        
        # Tenter la récupération
        return self._attempt_recovery(error, error_type, context, error_record)
    
    def _classify_error(self, error: Exception) -> ErrorType:
        """Classifie le type d'erreur"""
        if isinstance(error, (FileNotFoundError, IsADirectoryError, NotADirectoryError)):
            return ErrorType.FILE_ACCESS
        elif isinstance(error, MemoryError):
            return ErrorType.MEMORY_ERROR
        elif isinstance(error, PermissionError):
            return ErrorType.PERMISSION_ERROR
        elif isinstance(error, TimeoutError):
            return ErrorType.TIMEOUT_ERROR
        elif "corrupt" in str(error).lower() or "invalid format" in str(error).lower():
            return ErrorType.CORRUPT_DATA
        elif "validation" in str(error).lower():
            return ErrorType.VALIDATION_ERROR
        else:
            return ErrorType.FILE_ACCESS  # Par défaut
    
    def _attempt_recovery(self, error: Exception, error_type: ErrorType, 
                         context: Dict[str, Any], error_record: Dict) -> Tuple[bool, str, Any]:
        """Tente la récupération selon les stratégies définies"""
        strategies = self.recovery_strategies.get(error_type, [RecoveryStrategy.USER_CHOICE])
        
        for strategy in strategies:
            try:
                success, message, data = self._execute_recovery_strategy(
                    strategy, error, error_type, context
                )
                
                if success:
                    error_record['recovery_attempted'] = True
                    error_record['recovery_strategy'] = strategy.value
                    error_record['recovery_success'] = True
                    logging.info(f"Récupération réussie avec stratégie {strategy.value}")
                    return True, message, data
                    
            except Exception as recovery_error:
                logging.error(f"Erreur lors de la récupération {strategy.value}: {recovery_error}")
                continue
        
        # Aucune récupération n'a fonctionné
        error_record['recovery_attempted'] = True
        error_record['recovery_success'] = False
        return False, f"Récupération impossible pour l'erreur: {error}", None
    
    def _execute_recovery_strategy(self, strategy: RecoveryStrategy, error: Exception,
                                 error_type: ErrorType, context: Dict[str, Any]) -> Tuple[bool, str, Any]:
        """Exécute une stratégie de récupération spécifique"""
        
        if strategy == RecoveryStrategy.RETRY:
            return self._strategy_retry(error, context)
        elif strategy == RecoveryStrategy.AUTO_FIX:
            return self._strategy_auto_fix(error, error_type, context)
        elif strategy == RecoveryStrategy.FALLBACK:
            return self._strategy_fallback(error, context)
        elif strategy == RecoveryStrategy.SKIP:
            return self._strategy_skip(error, context)
        elif strategy == RecoveryStrategy.USER_CHOICE:
            return self._strategy_user_choice(error, context)
        else:
            return False, "Stratégie de récupération inconnue", None
    
    def _strategy_retry(self, error: Exception, context: Dict[str, Any]) -> Tuple[bool, str, Any]:
        """Stratégie de retry automatique"""
        max_retries = context.get('max_retries', 3)
        current_retry = context.get('current_retry', 0)
        
        if current_retry < max_retries:
            context['current_retry'] = current_retry + 1
            time.sleep(min(2 ** current_retry, 10))  # Backoff exponentiel
            return True, f"Tentative {current_retry + 1}/{max_retries}", context
        
        return False, "Nombre maximum de tentatives atteint", None
    
    def _strategy_auto_fix(self, error: Exception, error_type: ErrorType, 
                          context: Dict[str, Any]) -> Tuple[bool, str, Any]:
        """Stratégie de correction automatique"""
        if error_type == ErrorType.MEMORY_ERROR:
            # Libérer de la mémoire
            gc.collect()
            return True, "Mémoire libérée automatiquement", {"memory_cleaned": True}
        
        elif error_type == ErrorType.VALIDATION_ERROR:
            # Correction de données simples
            if "empty" in str(error).lower():
                return True, "Lignes vides ignorées automatiquement", {"skip_empty": True}
        
        elif error_type == ErrorType.CORRUPT_DATA:
            # Tentative de nettoyage des données
            return True, "Données corrompues ignorées", {"data_cleaned": True}
        
        return False, "Correction automatique non disponible", None
    
    def _strategy_fallback(self, error: Exception, context: Dict[str, Any]) -> Tuple[bool, str, Any]:
        """Stratégie de fallback vers une méthode alternative"""
        fallback_method = context.get('fallback_method')
        if fallback_method and callable(fallback_method):
            try:
                result = fallback_method()
                return True, "Méthode alternative utilisée", result
            except Exception:
                pass
        
        return False, "Aucune méthode alternative disponible", None
    
    def _strategy_skip(self, error: Exception, context: Dict[str, Any]) -> Tuple[bool, str, Any]:
        """Stratégie d'ignore/skip de l'élément problématique"""
        return True, "Élément problématique ignoré", {"skipped": True}
    
    def _strategy_user_choice(self, error: Exception, context: Dict[str, Any]) -> Tuple[bool, str, Any]:
        """Stratégie demandant le choix à l'utilisateur"""
        # Cette stratégie sera implémentée avec des dialogues GUI
        return False, "Intervention utilisateur requise", {"user_choice_needed": True}
    
    def get_error_statistics(self) -> Dict[str, Any]:
        """Retourne les statistiques d'erreurs"""
        if not self.error_history:
            return {"total_errors": 0}
        
        total_errors = len(self.error_history)
        error_types = defaultdict(int)
        recovery_success_count = 0
        
        for error_record in self.error_history:
            error_types[error_record['error_type']] += 1
            if error_record.get('recovery_success', False):
                recovery_success_count += 1
        
        return {
            "total_errors": total_errors,
            "error_types": dict(error_types),
            "recovery_success_rate": recovery_success_count / total_errors if total_errors > 0 else 0,
            "recent_errors": self.error_history[-10:]  # 10 dernières erreurs
        }


class AutoSaveManager:
    """Gestionnaire de sauvegarde automatique du travail en cours"""
    
    def __init__(self, backup_dir: str = "backups"):
        self.backup_dir = Path(backup_dir)
        self.backup_dir.mkdir(exist_ok=True)
        self.is_enabled = True
        self.backup_interval = BACKUP_INTERVAL
        self.max_backup_files = MAX_BACKUP_FILES
        self.last_backup_time = 0
        self.current_session_data = {}
        
    def enable_auto_save(self, enabled: bool = True):
        """Active/désactive la sauvegarde automatique"""
        self.is_enabled = enabled
        logging.info(f"Sauvegarde automatique {'activée' if enabled else 'désactivée'}")
    
    def set_backup_interval(self, interval_seconds: int):
        """Définit l'intervalle de sauvegarde"""
        self.backup_interval = max(60, interval_seconds)  # Minimum 1 minute
    
    def save_session_state(self, session_data: Dict[str, Any]) -> bool:
        """Sauvegarde l'état de la session actuelle"""
        if not self.is_enabled:
            return False
        
        try:
            current_time = time.time()
            
            # Vérifier si il faut faire une sauvegarde
            if current_time - self.last_backup_time < self.backup_interval:
                return False
            
            # Préparer les données de sauvegarde
            backup_data = {
                'timestamp': datetime.now().isoformat(),
                'session_id': datetime.now().strftime("%Y%m%d_%H%M%S"),
                'application_state': session_data,
                'version': '3.1'
            }
            
            # Créer le fichier de sauvegarde
            backup_filename = f"session_backup_{backup_data['session_id']}.json"
            backup_path = self.backup_dir / backup_filename
            
            with open(backup_path, 'w', encoding='utf-8') as f:
                json.dump(backup_data, f, indent=2, ensure_ascii=False)
            
            self.last_backup_time = current_time
            self.current_session_data = session_data.copy()
            
            # Nettoyer les anciennes sauvegardes
            self._cleanup_old_backups()
            
            logging.info(f"Session sauvegardée: {backup_filename}")
            return True
            
        except Exception as e:
            logging.error(f"Erreur lors de la sauvegarde automatique: {e}")
            return False
    
    def load_latest_backup(self) -> Optional[Dict[str, Any]]:
        """Charge la sauvegarde la plus récente"""
        try:
            backup_files = list(self.backup_dir.glob("session_backup_*.json"))
            if not backup_files:
                return None
            
            # Trier par date de modification (plus récent en premier)
            latest_backup = max(backup_files, key=lambda f: f.stat().st_mtime)
            
            with open(latest_backup, 'r', encoding='utf-8') as f:
                backup_data = json.load(f)
            
            logging.info(f"Sauvegarde chargée: {latest_backup.name}")
            return backup_data
            
        except Exception as e:
            logging.error(f"Erreur lors du chargement de la sauvegarde: {e}")
            return None
    
    def list_available_backups(self) -> List[Dict[str, Any]]:
        """Liste les sauvegardes disponibles"""
        backups = []
        
        try:
            backup_files = list(self.backup_dir.glob("session_backup_*.json"))
            
            for backup_file in backup_files:
                try:
                    with open(backup_file, 'r', encoding='utf-8') as f:
                        backup_data = json.load(f)
                    
                    backups.append({
                        'filename': backup_file.name,
                        'timestamp': backup_data.get('timestamp', 'Unknown'),
                        'session_id': backup_data.get('session_id', 'Unknown'),
                        'size': backup_file.stat().st_size,
                        'path': str(backup_file)
                    })
                except Exception:
                    continue
            
            # Trier par timestamp (plus récent en premier)
            backups.sort(key=lambda x: x['timestamp'], reverse=True)
            
        except Exception as e:
            logging.error(f"Erreur lors de la liste des sauvegardes: {e}")
        
        return backups
    
    def _cleanup_old_backups(self):
        """Nettoie les anciennes sauvegardes"""
        try:
            backup_files = list(self.backup_dir.glob("session_backup_*.json"))
            
            if len(backup_files) > self.max_backup_files:
                # Trier par date de modification (plus ancien en premier)
                backup_files.sort(key=lambda f: f.stat().st_mtime)
                
                # Supprimer les fichiers les plus anciens
                files_to_remove = backup_files[:-self.max_backup_files]
                for file_to_remove in files_to_remove:
                    file_to_remove.unlink()
                    logging.info(f"Ancienne sauvegarde supprimée: {file_to_remove.name}")
                    
        except Exception as e:
            logging.error(f"Erreur lors du nettoyage des sauvegardes: {e}")
    
    def create_manual_backup(self, session_data: Dict[str, Any], name: str = None) -> bool:
        """Crée une sauvegarde manuelle avec un nom personnalisé"""
        try:
            if name:
                backup_filename = f"manual_backup_{name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            else:
                backup_filename = f"manual_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            
            backup_data = {
                'timestamp': datetime.now().isoformat(),
                'type': 'manual',
                'name': name or 'Manual Backup',
                'application_state': session_data,
                'version': '3.1'
            }
            
            backup_path = self.backup_dir / backup_filename
            
            with open(backup_path, 'w', encoding='utf-8') as f:
                json.dump(backup_data, f, indent=2, ensure_ascii=False)
            
            logging.info(f"Sauvegarde manuelle créée: {backup_filename}")
            return True
            
        except Exception as e:
            logging.error(f"Erreur lors de la sauvegarde manuelle: {e}")
            return False


class UserGuideDialog(QDialog):
    """Dialogue affichant le guide utilisateur complet"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Guide utilisateur - ExcelCompiler v3.1")
        self.setWindowIcon(QIcon("icon.png") if os.path.exists("icon.png") else self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView))
        self.resize(900, 700)
        
        # Configuration du dialogue
        self.setModal(True)
        layout = QVBoxLayout(self)
        
        # Zone de texte avec le guide
        self.text_area = QTextEdit()
        self.text_area.setReadOnly(True)
        self.text_area.setHtml(self.get_user_guide_content())
        
        # Style CSS pour le guide
        self.text_area.setStyleSheet("""
            QTextEdit {
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 11pt;
                line-height: 1.6;
                background-color: #ffffff;
                border: 1px solid #ddd;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        
        layout.addWidget(self.text_area)
        
        # Boutons
        button_layout = QHBoxLayout()
        
        # Bouton Imprimer (si possible)
        print_button = QPushButton("🖨️ Imprimer")
        print_button.clicked.connect(self.print_guide)
        button_layout.addWidget(print_button)
        
        button_layout.addStretch()
        
        # Bouton Fermer
        close_button = QPushButton("Fermer")
        close_button.clicked.connect(self.accept)
        close_button.setDefault(True)
        button_layout.addWidget(close_button)
        
        layout.addLayout(button_layout)
    
    def get_user_guide_content(self) -> str:
        """Retourne le contenu HTML du guide utilisateur"""
        return """
        <html>
        <head>
            <style>
                body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; }
                h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 10px; }
                h2 { color: #34495e; margin-top: 30px; border-left: 4px solid #3498db; padding-left: 15px; }
                h3 { color: #7f8c8d; margin-top: 20px; }
                .feature { background: #ecf0f1; padding: 15px; margin: 10px 0; border-radius: 8px; border-left: 4px solid #3498db; }
                .warning { background: #fff3cd; padding: 15px; margin: 10px 0; border-radius: 8px; border-left: 4px solid #ffc107; }
                .success { background: #d4edda; padding: 15px; margin: 10px 0; border-radius: 8px; border-left: 4px solid #28a745; }
                .step { background: #f8f9fa; padding: 10px; margin: 5px 0; border-radius: 4px; }
                .code { background: #f4f4f4; padding: 2px 6px; border-radius: 3px; font-family: monospace; }
                ul { margin-left: 20px; }
                li { margin-bottom: 8px; }
                table { border-collapse: collapse; width: 100%; margin: 15px 0; }
                th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                th { background-color: #f2f2f2; font-weight: bold; }
            </style>
        </head>
        <body>
            <h1>🎯 Guide Utilisateur ExcelCompiler v3.1</h1>
            
            <div class="feature">
                <h2>📋 Table des matières</h2>
                <ul>
                    <li><a href="#introduction">Introduction</a></li>
                    <li><a href="#installation">Installation et démarrage</a></li>
                    <li><a href="#interface">Interface utilisateur</a></li>
                    <li><a href="#utilisation">Utilisation pas-à-pas</a></li>
                    <li><a href="#fonctionnalites">Fonctionnalités avancées</a></li>
                    <li><a href="#troubleshooting">Dépannage</a></li>
                    <li><a href="#faq">FAQ</a></li>
                </ul>
            </div>
            
            <h2 id="introduction">🚀 Introduction</h2>
            <p>ExcelCompiler est un outil professionnel conçu pour <strong>fusionner rapidement plusieurs fichiers Excel et CSV</strong> en un seul fichier consolidé. Il s'adresse aux analystes de données, gestionnaires de projet et professionnels travaillant régulièrement avec des données tabulaires.</p>
            
            <div class="success">
                <h3>✨ Avantages clés</h3>
                <ul>
                    <li><strong>Gain de temps :</strong> Automatise des tâches qui prendraient des heures manuellement</li>
                    <li><strong>Fiabilité :</strong> Gestion robuste des erreurs et validation des données</li>
                    <li><strong>Flexibilité :</strong> Support de multiples formats et options de configuration</li>
                    <li><strong>Simplicité :</strong> Interface intuitive avec assistant intégré</li>
                </ul>
            </div>
            
            <h2 id="installation">💾 Installation et démarrage</h2>
            
            <h3>Prérequis système</h3>
            <table>
                <tr><th>Composant</th><th>Minimum requis</th><th>Recommandé</th></tr>
                <tr><td>OS</td><td>Windows 10, macOS 10.14, Linux Ubuntu 18.04</td><td>Windows 11, macOS 12+, Ubuntu 22.04+</td></tr>
                <tr><td>RAM</td><td>4 GB</td><td>8 GB ou plus</td></tr>
                <tr><td>Espace disque</td><td>500 MB</td><td>2 GB pour les gros fichiers</td></tr>
                <tr><td>Python</td><td>3.8+</td><td>3.10+ (si installation depuis source)</td></tr>
            </table>
            
            <h3>Installation</h3>
            <div class="step">
                <strong>Option 1 : Exécutable (Recommandé)</strong><br>
                1. Téléchargez <span class="code">ExcelCompiler_Setup.exe</span><br>
                2. Exécutez l'installateur en tant qu'administrateur<br>
                3. Suivez les instructions à l'écran<br>
                4. Lancez depuis le menu Démarrer ou le raccourci bureau
            </div>
            
            <div class="step">
                <strong>Option 2 : Depuis les sources</strong><br>
                1. Installez Python 3.8+ et pip<br>
                2. Installez les dépendances : <span class="code">pip install -r requirements.txt</span><br>
                3. Lancez : <span class="code">python compiler.py</span>
            </div>
            
            <h2 id="interface">🖥️ Interface utilisateur</h2>
            
            <h3>Vue d'ensemble</h3>
            <p>L'interface d'ExcelCompiler est organisée en onglets pour une navigation intuitive :</p>
            
            <div class="feature">
                <h4>📁 Onglet Fichiers</h4>
                <ul>
                    <li><strong>Répertoire :</strong> Sélection du dossier contenant vos fichiers</li>
                    <li><strong>Liste des fichiers :</strong> Affichage et sélection des fichiers à compiler</li>
                    <li><strong>Filtres :</strong> Options de tri et de filtrage des fichiers</li>
                </ul>
            </div>
            
            <div class="feature">
                <h4>⚙️ Onglet Configuration</h4>
                <ul>
                    <li><strong>En-têtes :</strong> Configuration des lignes d'en-tête</li>
                    <li><strong>Options :</strong> Paramètres de traitement des données</li>
                    <li><strong>Format :</strong> Options de formatage et de tri</li>
                </ul>
            </div>
            
            <div class="feature">
                <h4>👁️ Onglet Prévisualisation</h4>
                <ul>
                    <li><strong>Aperçu :</strong> Visualisation des données avant compilation</li>
                    <li><strong>Validation :</strong> Vérification de la cohérence des structures</li>
                    <li><strong>Statistiques :</strong> Informations sur les fichiers sélectionnés</li>
                </ul>
            </div>
            
            <h2 id="utilisation">📖 Utilisation pas-à-pas</h2>
            
            <h3>Étape 1 : Sélection des fichiers</h3>
            <div class="step">
                1. <strong>Cliquez sur "Parcourir"</strong> dans la section Répertoire<br>
                2. <strong>Naviguez</strong> vers le dossier contenant vos fichiers Excel/CSV<br>
                3. <strong>Sélectionnez le dossier</strong> et cliquez sur "Sélectionner un dossier"<br>
                4. <strong>Cochez les fichiers</strong> que vous souhaitez compiler dans la liste
            </div>
            
            <div class="warning">
                <strong>⚠️ Important :</strong> Assurez-vous que tous vos fichiers ont une structure similaire (mêmes colonnes dans le même ordre) pour un résultat optimal.
            </div>
            
            <h3>Étape 2 : Configuration des paramètres</h3>
            <div class="step">
                1. <strong>Ligne de début :</strong> Indiquez à quelle ligne commencent vos en-têtes (généralement 1)<br>
                2. <strong>Nombre de lignes d'en-tête :</strong> Combien de lignes forment votre en-tête (généralement 1)<br>
                3. <strong>Options avancées :</strong> Configurez selon vos besoins :
                <ul>
                    <li>Ajouter le nom de fichier source</li>
                    <li>Trier les données par colonne</li>
                    <li>Supprimer les lignes vides</li>
                    <li>Éliminer les doublons</li>
                </ul>
            </div>
            
            <h3>Étape 3 : Prévisualisation (optionnel)</h3>
            <div class="step">
                1. <strong>Cliquez sur l'onglet "Prévisualisation"</strong><br>
                2. <strong>Vérifiez</strong> que la structure des données est correcte<br>
                3. <strong>Contrôlez</strong> les statistiques de fichiers<br>
                4. <strong>Ajustez</strong> les paramètres si nécessaire
            </div>
            
            <h3>Étape 4 : Compilation</h3>
            <div class="step">
                1. <strong>Cliquez sur "Compiler"</strong><br>
                2. <strong>Suivez la progression</strong> dans la barre de statut<br>
                3. <strong>Patientez</strong> pendant le traitement (vous pouvez annuler si besoin)<br>
                4. <strong>Sauvegardez</strong> le fichier résultat à l'emplacement de votre choix
            </div>
            
            <h2 id="fonctionnalites">🔧 Fonctionnalités avancées</h2>
            
            <h3>Détection automatique des formats</h3>
            <p>ExcelCompiler détecte automatiquement :</p>
            <ul>
                <li><strong>Encodages :</strong> UTF-8, Latin-1, CP1252, etc.</li>
                <li><strong>Délimiteurs CSV :</strong> Virgule, point-virgule, tabulation, pipe</li>
                <li><strong>Formats de date :</strong> Français, anglais, ISO</li>
            </ul>
            
            <h3>Gestion des erreurs</h3>
            <div class="feature">
                <ul>
                    <li><strong>Récupération d'erreurs :</strong> Continue le traitement même si certains fichiers échouent</li>
                    <li><strong>Journalisation :</strong> Logs détaillés de toutes les opérations</li>
                    <li><strong>Validation :</strong> Vérification en temps réel des paramètres</li>
                    <li><strong>Annulation :</strong> Possibilité d'interrompre à tout moment</li>
                </ul>
            </div>
            
            <h3>Options de performance</h3>
            <table>
                <tr><th>Option</th><th>Description</th><th>Recommandation</th></tr>
                <tr><td>Traitement par chunks</td><td>Lecture des gros fichiers par petits blocs</td><td>Activé automatiquement</td></tr>
                <tr><td>Gestion mémoire</td><td>Nettoyage automatique de la mémoire</td><td>Activé automatiquement</td></tr>
                <tr><td>Cache métadonnées</td><td>Mise en cache des informations de fichiers</td><td>Activé automatiquement</td></tr>
                <tr><td>Threads multiples</td><td>Traitement parallèle des fichiers</td><td>Activé automatiquement</td></tr>
            </table>
            
            <h2 id="troubleshooting">🔧 Dépannage</h2>
            
            <h3>Problèmes courants</h3>
            
            <div class="warning">
                <h4>❌ "Fichier en cours d'utilisation"</h4>
                <strong>Cause :</strong> Le fichier est ouvert dans Excel ou une autre application<br>
                <strong>Solution :</strong> Fermez tous les fichiers Excel avant de lancer la compilation
            </div>
            
            <div class="warning">
                <h4>❌ "Erreur de mémoire insuffisante"</h4>
                <strong>Cause :</strong> Trop de gros fichiers traités simultanément<br>
                <strong>Solution :</strong> Réduisez le nombre de fichiers ou utilisez un ordinateur avec plus de RAM
            </div>
            
            <div class="warning">
                <h4>❌ "Structure de données incohérente"</h4>
                <strong>Cause :</strong> Les fichiers n'ont pas les mêmes colonnes<br>
                <strong>Solution :</strong> Vérifiez que tous vos fichiers ont la même structure d'en-têtes
            </div>
            
            <div class="warning">
                <h4>❌ "Encodage non reconnu"</h4>
                <strong>Cause :</strong> Fichier CSV avec un encodage exotique<br>
                <strong>Solution :</strong> Convertissez le fichier en UTF-8 avec un éditeur de texte
            </div>
            
            <h3>Optimisation des performances</h3>
            <div class="success">
                <h4>💡 Conseils pour de meilleures performances</h4>
                <ul>
                    <li>Évitez de trier de très gros fichiers (>100MB)</li>
                    <li>Fermez les autres applications gourmandes en mémoire</li>
                    <li>Utilisez des fichiers sur le disque local plutôt que réseau</li>
                    <li>Préférez les fichiers .xlsx aux .xls pour de meilleures performances</li>
                </ul>
            </div>
            
            <h2 id="faq">❓ FAQ - Foire Aux Questions</h2>
            
            <h3>Q: Puis-je mélanger des fichiers Excel et CSV ?</h3>
            <p><strong>R:</strong> Oui, ExcelCompiler supporte le mélange de formats. Assurez-vous simplement que la structure des colonnes est cohérente.</p>
            
            <h3>Q: Quelle est la taille maximale de fichier supportée ?</h3>
            <p><strong>R:</strong> Il n'y a pas de limite théorique, mais la performance dépend de votre RAM. Fichiers testés jusqu'à 500MB avec 8GB de RAM.</p>
            
            <h3>Q: Puis-je annuler une compilation en cours ?</h3>
            <p><strong>R:</strong> Oui, cliquez sur le bouton "Annuler" dans la barre de progression. L'annulation est propre et n'endommage pas les données.</p>
            
            <h3>Q: Les formules Excel sont-elles préservées ?</h3>
            <p><strong>R:</strong> Non, seules les valeurs sont copiées. Les formules sont converties en leurs résultats calculés.</p>
            
            <h3>Q: Comment signaler un bug ou suggérer une amélioration ?</h3>
            <p><strong>R:</strong> Contactez le développeur à zimkada@gmail.com avec une description détaillée et les fichiers de log.</p>
            
            <h3>Q: ExcelCompiler fonctionne-t-il hors ligne ?</h3>
            <p><strong>R:</strong> Oui, aucune connexion Internet n'est requise pour le fonctionnement normal.</p>
            
            <div class="success">
                <h3>🎉 Félicitations !</h3>
                <p>Vous avez maintenant toutes les clés pour utiliser efficacement ExcelCompiler. N'hésitez pas à explorer les fonctionnalités avancées et à consulter l'assistant de démarrage pour un accompagnement interactif.</p>
            </div>
            
            <hr>
            <p><em>Guide utilisateur ExcelCompiler v3.1 - Dernière mise à jour : Juillet 2025</em></p>
        </body>
        </html>
        """
    
    def print_guide(self):
        """Imprime le guide utilisateur"""
        try:
            from PyQt6.QtPrintSupport import QPrinter, QPrintDialog
            printer = QPrinter()
            dialog = QPrintDialog(printer, self)
            if dialog.exec() == QPrintDialog.DialogCode.Accepted:
                self.text_area.print(printer)
        except ImportError:
            QMessageBox.information(
                self, 
                "Impression", 
                "La fonctionnalité d'impression n'est pas disponible.\nVous pouvez copier le contenu et l'imprimer depuis un traitement de texte."
            )

class ResponsiveManager:
    """Gestionnaire de responsivité pour adapter l'interface à la taille d'écran"""
    
    def __init__(self):
        self.screen_size = None
        self.scale_factor = 1.0
        self.font_scale = 1.0
        self.current_breakpoint = 'medium'
        
    def detect_screen_size(self, widget) -> Dict[str, Any]:
        """Détecte la taille d'écran disponible"""
        if hasattr(widget, 'screen'):
            screen = widget.screen()
            geometry = screen.availableGeometry()
            width = geometry.width()
            height = geometry.height()
        else:
            # Fallback pour les anciens systèmes
            width = 1366
            height = 768
            
        self.screen_size = {'width': width, 'height': height}
        
        # Déterminer le breakpoint
        if width <= SCREEN_BREAKPOINTS['small']:
            self.current_breakpoint = 'small'
        elif width <= SCREEN_BREAKPOINTS['medium']:
            self.current_breakpoint = 'medium'
        elif width <= SCREEN_BREAKPOINTS['large']:
            self.current_breakpoint = 'large'
        else:
            self.current_breakpoint = 'xlarge'
            
        # Calculer les facteurs d'échelle
        self.scale_factor = min(width / 1366, height / 768)  # Base de référence
        self.font_scale = RESPONSIVE_FONT_SCALES[self.current_breakpoint]
        
        return {
            'width': width,
            'height': height,
            'breakpoint': self.current_breakpoint,
            'scale_factor': self.scale_factor,
            'font_scale': self.font_scale
        }
    
    def get_responsive_size(self, base_width: int, base_height: int) -> Tuple[int, int]:
        """Calcule une taille responsive basée sur l'écran"""
        if not self.screen_size:
            return base_width, base_height
            
        screen_width = self.screen_size['width']
        screen_height = self.screen_size['height']
        
        # Calculer la taille optimale (70-80% de l'écran)
        optimal_width = min(int(screen_width * 0.8), int(base_width * self.scale_factor))
        optimal_height = min(int(screen_height * 0.8), int(base_height * self.scale_factor))
        
        # Assurer des minimums
        min_width = min(800, screen_width - 100)
        min_height = min(600, screen_height - 100)
        
        return max(optimal_width, min_width), max(optimal_height, min_height)
    
    def get_responsive_font_size(self, base_size: int) -> int:
        """Calcule une taille de police responsive"""
        return max(8, int(base_size * self.font_scale))
    
    def get_responsive_margin(self, base_margin: int) -> int:
        """Calcule une marge responsive"""
        return max(5, int(base_margin * self.scale_factor))
    
    def get_responsive_icon_size(self, base_size: int) -> int:
        """Calcule une taille d'icône responsive"""
        return max(16, int(base_size * self.scale_factor))

class FlexibleLayout:
    """Utilitaires pour créer des layouts flexibles et adaptatifs"""
    
    @staticmethod
    def create_responsive_grid(parent, items_per_row_config: Dict[str, int], spacing: int = 10):
        """
        Crée une grille qui s'adapte selon la taille d'écran.
        items_per_row_config: {'small': 1, 'medium': 2, 'large': 3, 'xlarge': 4}
        """
        layout = QGridLayout()
        layout.setSpacing(spacing)
        layout.setContentsMargins(spacing, spacing, spacing, spacing)
        
        # Stocker la configuration pour le redimensionnement
        layout.items_per_row_config = items_per_row_config
        layout.child_widgets = []
        
        return layout
    
    @staticmethod
    def update_grid_layout(layout, responsive_manager: ResponsiveManager):
        """Met à jour une grille responsive selon la taille d'écran actuelle"""
        if not hasattr(layout, 'items_per_row_config') or not hasattr(layout, 'child_widgets'):
            return
            
        breakpoint = responsive_manager.current_breakpoint
        items_per_row = layout.items_per_row_config.get(breakpoint, 2)
        
        # Réorganiser les widgets dans la grille
        for i, widget in enumerate(layout.child_widgets):
            row = i // items_per_row
            col = i % items_per_row
            layout.addWidget(widget, row, col)
    
    @staticmethod
    def create_adaptive_splitter(orientation=Qt.Orientation.Horizontal, sizes: List[int] = None):
        """Crée un splitter adaptatif avec des tailles proportionnelles"""
        splitter = QSplitter(orientation)
        splitter.setChildrenCollapsible(False)
        
        if sizes:
            splitter.setSizes(sizes)
            
        # Politique de redimensionnement
        splitter.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        
        return splitter
    
    @staticmethod
    def make_widget_responsive(widget, min_size: Tuple[int, int] = None, 
                             size_policy: Tuple[QSizePolicy.Policy, QSizePolicy.Policy] = None):
        """Rend un widget responsive"""
        if min_size:
            widget.setMinimumSize(QSize(min_size[0], min_size[1]))
            
        if size_policy:
            widget.setSizePolicy(QSizePolicy(size_policy[0], size_policy[1]))
        else:
            widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        
        return widget

class ResponsiveTableWidget(QTableWidget):
    """TableWidget adaptatif qui ajuste automatiquement ses colonnes"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.responsive_manager = None
        self.setup_responsive_behavior()
    
    def setup_responsive_behavior(self):
        """Configure le comportement responsive du tableau"""
        header = self.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        
        # Politique de taille
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        
        # Connecter au redimensionnement
        self.horizontalHeader().sectionResized.connect(self.on_section_resized)
    
    def set_responsive_manager(self, manager: ResponsiveManager):
        """Associe un gestionnaire de responsivité"""
        self.responsive_manager = manager
        
    def resizeEvent(self, event):
        """Gestionnaire de redimensionnement adaptatif"""
        super().resizeEvent(event)
        self.adjust_column_sizes()
    
    def adjust_column_sizes(self):
        """Ajuste automatiquement les tailles de colonnes"""
        if self.columnCount() == 0:
            return
            
        available_width = self.viewport().width()
        
        # Calculer la largeur idéale par colonne
        ideal_width = available_width // self.columnCount()
        min_width = 100 if self.responsive_manager else 80
        
        # Ajuster selon le gestionnaire de responsivité
        if self.responsive_manager:
            min_width = self.responsive_manager.get_responsive_margin(min_width)
            
        # Appliquer les tailles
        for col in range(self.columnCount()):
            current_width = max(ideal_width, min_width)
            self.setColumnWidth(col, current_width)
    
    def on_section_resized(self, logical_index, old_size, new_size):
        """Gestionnaire de redimensionnement de section"""
        # Empêcher les colonnes de devenir trop petites
        min_width = 80 if not self.responsive_manager else self.responsive_manager.get_responsive_margin(80)
        if new_size < min_width:
            self.setColumnWidth(logical_index, min_width)

class CancellableProgressWidget(QWidget):
    """Widget de progression avec bouton d'annulation intégré"""
    
    cancel_requested = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        self.is_cancelling = False
        
    def setup_ui(self):
        """Configure l'interface du widget de progression"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # Label de statut
        self.status_label = QLabel("Prêt")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status_label)
        
        # Barre de progression
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)
        
        # Layout pour les informations détaillées
        info_layout = QHBoxLayout()
        
        # Label de progression détaillée
        self.detail_label = QLabel("")
        self.detail_label.setStyleSheet("color: #666666; font-size: 12px;")
        info_layout.addWidget(self.detail_label)
        
        info_layout.addStretch()
        
        # Label de temps estimé
        self.time_label = QLabel("")
        self.time_label.setStyleSheet("color: #666666; font-size: 12px;")
        self.time_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        info_layout.addWidget(self.time_label)
        
        layout.addLayout(info_layout)
        
        # Bouton d'annulation
        self.cancel_button = QPushButton("Annuler")
        self.cancel_button.setEnabled(False)
        self.cancel_button.clicked.connect(self.request_cancel)
        layout.addWidget(self.cancel_button)
        
        # Variables pour le calcul du temps
        self.start_time = None
        self.last_progress = 0
        
    def start_operation(self, operation_name: str = "Opération en cours"):
        """Démarre une nouvelle opération"""
        self.status_label.setText(operation_name)
        self.progress_bar.setValue(0)
        self.detail_label.setText("")
        self.time_label.setText("")
        self.cancel_button.setEnabled(True)
        self.cancel_button.setText("Annuler")
        self.is_cancelling = False
        self.start_time = time.time()
        self.last_progress = 0
        
    def update_progress(self, value: int, detail: str = ""):
        """Met à jour la progression de manière thread-safe"""
        ThreadSafeGUIHelper.assert_main_thread("CancellableProgressWidget.update_progress")
        
        if self.is_cancelling:
            return
            
        self.progress_bar.setValue(value)
        self.last_progress = value
        
        if detail:
            self.detail_label.setText(detail)
            
        # Calculer et afficher le temps estimé avec plus de précision
        if self.start_time and value > 0:
            elapsed = time.time() - self.start_time
            if value < 100:
                estimated_total = elapsed * 100 / value
                remaining = estimated_total - elapsed
                # Amélioration : estimation plus précise avec lissage
                if hasattr(self, '_last_estimates'):
                    self._last_estimates.append(remaining)
                    if len(self._last_estimates) > 5:
                        self._last_estimates.pop(0)
                    # Moyenne mobile pour stabiliser l'estimation
                    remaining = sum(self._last_estimates) / len(self._last_estimates)
                else:
                    self._last_estimates = [remaining]
                
                self.time_label.setText(f"Temps restant: {self.format_time(remaining)}")
            else:
                self.time_label.setText(f"Terminé en {self.format_time(elapsed)}")
    
    def update_stage(self, stage_name: str, stage_progress: int = None):
        """Met à jour l'étape actuelle avec un message informatif"""
        ThreadSafeGUIHelper.assert_main_thread("CancellableProgressWidget.update_stage")
        
        if self.is_cancelling:
            return
            
        # Messages informatifs selon l'étape
        stage_messages = {
            "loading_files": "📂 Chargement des fichiers...",
            "analyzing_headers": "📋 Analyse des en-têtes...",
            "processing_data": "⚙️ Traitement des données...",
            "removing_duplicates": "🔄 Suppression des doublons...", 
            "sorting_data": "📊 Tri des données...",
            "generating_output": "📄 Génération du fichier de sortie...",
            "finalizing": "✅ Finalisation...",
            "verification": "🔍 Vérification préliminaire...",
            "memory_cleanup": "🧹 Nettoyage mémoire..."
        }
        
        friendly_message = stage_messages.get(stage_name, stage_name)
        self.set_status(friendly_message)
        
        if stage_progress is not None:
            self.update_progress(stage_progress, f"Étape: {friendly_message}")
    
    def set_status(self, status: str):
        """Met à jour le statut"""
        self.status_label.setText(status)
    
    def show_file_info(self, current_file: int, total_files: int, filename: str):
        """Affiche des informations sur le fichier en cours de traitement"""
        ThreadSafeGUIHelper.assert_main_thread("CancellableProgressWidget.show_file_info")
        
        if self.is_cancelling:
            return
            
        progress = int((current_file / total_files) * 100) if total_files > 0 else 0
        detail = f"📄 Fichier {current_file}/{total_files}: {filename}"
        
        self.update_progress(progress, detail)
    
    def show_memory_info(self, used_mb: float, available_mb: float):
        """Affiche des informations sur l'utilisation mémoire"""
        ThreadSafeGUIHelper.assert_main_thread("CancellableProgressWidget.show_memory_info")
        
        if self.is_cancelling:
            return
            
        detail = f"💾 Mémoire: {used_mb:.1f}MB utilisés, {available_mb:.1f}MB disponibles"
        self.detail_label.setText(detail)
    
    def request_cancel(self):
        """Demande l'annulation de l'opération"""
        if not self.is_cancelling:
            self.is_cancelling = True
            self.cancel_button.setText("Annulation en cours...")
            self.cancel_button.setEnabled(False)
            self.status_label.setText("Annulation en cours...")
            self.detail_label.setText("Arrêt des opérations, veuillez patienter...")
            self.cancel_requested.emit()
    
    def finish_operation(self, success: bool = True, message: str = ""):
        """Termine l'opération"""
        if success:
            self.progress_bar.setValue(100)
            self.status_label.setText(message or "Terminé avec succès")
            self.detail_label.setText("")
            if self.start_time:
                elapsed = time.time() - self.start_time
                self.time_label.setText(f"Terminé en {self.format_time(elapsed)}")
        else:
            self.status_label.setText(message or "Opération échouée")
            self.detail_label.setText("")
            self.time_label.setText("")
            
        self.cancel_button.setEnabled(False)
        self.cancel_button.setText("Annuler")
        self.is_cancelling = False
    
    def reset(self):
        """Remet le widget à zéro"""
        self.progress_bar.setValue(0)
        self.status_label.setText("Prêt")
        self.detail_label.setText("")
        self.time_label.setText("")
        self.cancel_button.setEnabled(False)
        self.cancel_button.setText("Annuler")
        self.is_cancelling = False
        self.start_time = None
        self.last_progress = 0
    
    def format_time(self, seconds: float) -> str:
        """Formate le temps en chaîne lisible"""
        if seconds < 60:
            return f"{int(seconds)}s"
        elif seconds < 3600:
            minutes = int(seconds // 60)
            secs = int(seconds % 60)
            return f"{minutes}m {secs}s"
        else:
            hours = int(seconds // 3600)
            minutes = int((seconds % 3600) // 60)
            return f"{hours}h {minutes}m"

class ValidationIndicator(QWidget):
    """Indicateur visuel pour l'état de validation d'un champ"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.validation_state = "neutral"  # neutral, valid, invalid, warning
        self.setup_ui()
        
    def setup_ui(self):
        """Configure l'interface de l'indicateur"""
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)
        
        # Icône d'état
        self.icon_label = QLabel()
        self.icon_label.setFixedSize(16, 16)
        layout.addWidget(self.icon_label)
        
        # Message de validation
        self.message_label = QLabel()
        self.message_label.setStyleSheet("font-size: 12px;")
        layout.addWidget(self.message_label)
        
        layout.addStretch()
        
        self.update_display()
    
    def set_state(self, state: str, message: str = ""):
        """Met à jour l'état de validation"""
        self.validation_state = state
        self.message_label.setText(message)
        self.update_display()
    
    def update_display(self):
        """Met à jour l'affichage selon l'état"""
        if self.validation_state == "valid":
            self.icon_label.setText("✓")
            self.icon_label.setStyleSheet("color: #4CAF50; font-weight: bold;")
            self.message_label.setStyleSheet("color: #4CAF50; font-size: 12px;")
        elif self.validation_state == "invalid":
            self.icon_label.setText("✗")
            self.icon_label.setStyleSheet("color: #f44336; font-weight: bold;")
            self.message_label.setStyleSheet("color: #f44336; font-size: 12px;")
        elif self.validation_state == "warning":
            self.icon_label.setText("⚠")
            self.icon_label.setStyleSheet("color: #FF9800; font-weight: bold;")
            self.message_label.setStyleSheet("color: #FF9800; font-size: 12px;")
        else:  # neutral
            self.icon_label.setText("")
            self.icon_label.setStyleSheet("")
            self.message_label.setStyleSheet("color: #666666; font-size: 12px;")

class RealTimeValidator(QObject):
    """Validateur en temps réel pour les champs de l'interface"""
    
    validation_changed = pyqtSignal(str, bool, str)  # field_name, is_valid, message
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.validators = {}
        self.validation_states = {}
        
    def add_field_validator(self, field_name: str, widget, validation_func, 
                          indicator: ValidationIndicator = None):
        """Ajoute un validateur pour un champ"""
        self.validators[field_name] = {
            'widget': widget,
            'validation_func': validation_func,
            'indicator': indicator
        }
        self.validation_states[field_name] = {"valid": True, "message": ""}
        
        # Connecter les signaux selon le type de widget
        if isinstance(widget, QLineEdit):
            widget.textChanged.connect(lambda: self._validate_field(field_name))
        elif isinstance(widget, QSpinBox):
            widget.valueChanged.connect(lambda: self._validate_field(field_name))
        elif isinstance(widget, QComboBox):
            widget.currentTextChanged.connect(lambda: self._validate_field(field_name))
            
        # Validation initiale
        self._validate_field(field_name)
    
    def _validate_field(self, field_name: str):
        """Valide un champ spécifique"""
        if field_name not in self.validators:
            return
            
        validator_info = self.validators[field_name]
        widget = validator_info['widget']
        validation_func = validator_info['validation_func']
        indicator = validator_info['indicator']
        
        # Obtenir la valeur selon le type de widget
        if isinstance(widget, QLineEdit):
            value = widget.text()
        elif isinstance(widget, QSpinBox):
            value = widget.value()
        elif isinstance(widget, QComboBox):
            value = widget.currentText()
        else:
            value = None
        
        # Exécuter la validation
        try:
            is_valid, message = validation_func(value)
        except Exception as e:
            is_valid, message = False, f"Erreur de validation: {str(e)}"
        
        # Mettre à jour l'état
        self.validation_states[field_name] = {"valid": is_valid, "message": message}
        
        # Mettre à jour l'indicateur visuel
        if indicator:
            if is_valid:
                if message:
                    indicator.set_state("valid", message)
                else:
                    indicator.set_state("neutral", "")
            else:
                indicator.set_state("invalid", message)
        
        # Mettre à jour le style du widget
        self._update_widget_style(widget, is_valid)
        
        # Émettre le signal
        self.validation_changed.emit(field_name, is_valid, message)
    
    def _update_widget_style(self, widget, is_valid: bool):
        """Met à jour le style visuel du widget selon l'état de validation"""
        if is_valid:
            # Style normal
            widget.setStyleSheet("")
        else:
            # Style d'erreur
            widget.setStyleSheet("""
                border: 2px solid #f44336;
                background-color: #ffebee;
            """)
    
    def is_all_valid(self) -> bool:
        """Vérifie si tous les champs sont valides"""
        return all(state["valid"] for state in self.validation_states.values())
    
    def get_invalid_fields(self) -> List[str]:
        """Retourne la liste des champs invalides"""
        return [field for field, state in self.validation_states.items() if not state["valid"]]
    
    def get_validation_summary(self) -> str:
        """Retourne un résumé des erreurs de validation"""
        invalid_fields = self.get_invalid_fields()
        if not invalid_fields:
            return "Tous les champs sont valides"
        
        messages = []
        for field in invalid_fields:
            message = self.validation_states[field]["message"]
            messages.append(f"• {field}: {message}")
        
        return "\n".join(messages)

# Fonctions de validation spécifiques
def validate_output_filename(filename: str) -> Tuple[bool, str]:
    """Valide le nom de fichier de sortie"""
    if not filename:
        return False, "Le nom de fichier est requis"
    
    if len(filename.strip()) == 0:
        return False, "Le nom de fichier ne peut pas être vide"
    
    # Caractères interdits dans les noms de fichiers
    invalid_chars = r'[<>:"/\\|?*]'
    if re.search(invalid_chars, filename):
        return False, "Caractères interdits: < > : \" / \\ | ? *"
    
    # Vérifier la longueur
    if len(filename) > 255:
        return False, "Nom trop long (max 255 caractères)"
    
    # Extension
    if not filename.endswith('.xlsx'):
        return True, "Extension .xlsx sera ajoutée automatiquement"
    
    return True, "Nom de fichier de sortie valide"

def validate_header_start_row(value: int) -> Tuple[bool, str]:
    """Valide la ligne de début des en-têtes"""
    if value < 1:
        return False, "La ligne de début doit être >= 1"
    
    if value > 1000:
        return False, "Ligne trop élevée (max 1000)"
    
    return True, ""

def validate_header_rows(value: int) -> Tuple[bool, str]:
    """Valide le nombre de lignes d'en-tête"""
    if value < 1:
        return False, "Au moins 1 ligne d'en-tête requise"
    
    if value > 10:
        return False, "Trop de lignes d'en-tête (max 10)"
    
    return True, ""

def validate_sort_column(value: str) -> Tuple[bool, str]:
    """Valide la colonne de tri (accepte lettres A,B,C ou nombres 1,2,3)"""
    if not value:
        return True, ""  # Optionnel
    
    value = value.strip().upper()
    
    # Format lettre (A, B, C, etc.)
    if len(value) == 1 and 'A' <= value <= 'Z':
        col_index = ord(value) - ord('A') + 1  # A=1, B=2, etc.
        return True, f"Tri sur colonne {value} (position {col_index})"
    
    # Format lettre multi-colonnes (AA, AB, etc.)
    if len(value) <= 3 and value.isalpha():
        try:
            # Conversion Excel-style (A=1, AA=27, etc.)
            col_index = 0
            for char in value:
                col_index = col_index * 26 + (ord(char) - ord('A') + 1)
            if col_index <= 100:
                return True, f"Tri sur colonne {value} (position {col_index})"
            else:
                return False, "Colonne trop élevée (max Z ou 100)"
        except:
            return False, "Format de colonne invalide"
    
    # Format numérique (1, 2, 3, etc.)
    try:
        col_num = int(value)
        if col_num < 1:
            return False, "Numéro de colonne doit être >= 1"
        if col_num > 100:
            return False, "Numéro de colonne trop élevé (max 100)"
        # Convertir en lettre pour affichage
        if col_num <= 26:
            letter = chr(ord('A') + col_num - 1)
            return True, f"Tri sur colonne {letter} (position {col_num})"
        else:
            return True, f"Tri sur colonne {col_num}"
    except ValueError:
        pass
    
    return False, "Format invalide. Utilisez A,B,C ou 1,2,3"

def validate_directory_selection(app_instance) -> Tuple[bool, str]:
    """Valide la sélection du répertoire et des fichiers"""
    if not app_instance.directory:
        return False, "Aucun répertoire sélectionné"
    
    if not os.path.exists(app_instance.directory):
        return False, "Le répertoire n'existe plus"
    
    if not os.access(app_instance.directory, os.R_OK):
        return False, "Permissions insuffisantes pour lire le répertoire"
    
    # Vérifier les fichiers sélectionnés
    if hasattr(app_instance, 'list_files'):
        selected_files = [item.text() for item in app_instance.list_files.selectedItems()]
        if not selected_files:
            return False, "Aucun fichier sélectionné"
        
        # Vérifier que les fichiers existent
        for file in selected_files:
            file_path = os.path.join(app_instance.directory, file)
            if not os.path.exists(file_path):
                return False, f"Fichier introuvable: {file}"
        
        return True, f"{len(selected_files)} fichier(s) sélectionné(s)"
    
    return True, "Répertoire valide"

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class MemoryManager:
    """Gestionnaire de mémoire pour surveiller et optimiser l'utilisation RAM"""
    
    def __init__(self, max_memory_mb: int = 512):
        self.max_memory_bytes = max_memory_mb * 1024 * 1024
        self.process = psutil.Process()
        
    def get_memory_usage(self) -> int:
        """Retourne l'utilisation mémoire actuelle en bytes"""
        return self.process.memory_info().rss
    
    def is_memory_limit_reached(self) -> bool:
        """Vérifie si la limite mémoire est atteinte"""
        return self.get_memory_usage() > self.max_memory_bytes
    
    def force_garbage_collection(self):
        """Force le garbage collection pour libérer la mémoire"""
        gc.collect()

class ValidationCache:
    """Cache pour les résultats de validation des fichiers"""
    
    def __init__(self, max_size=50):
        self._cache = {}
        self._access_times = {}
        self._lock = threading.Lock()
        self.max_size = max_size
    
    def get_validation_result(self, file_path: str) -> Optional[Tuple[bool, List[str]]]:
        """Récupère le résultat de validation du cache"""
        with self._lock:
            try:
                stat = os.stat(file_path)
                cache_key = f"{file_path}_{stat.st_mtime}_{stat.st_size}"
                
                if cache_key in self._cache:
                    self._access_times[cache_key] = time.time()
                    return self._cache[cache_key]
                return None
            except (OSError, IOError):
                return None
    
    def set_validation_result(self, file_path: str, is_valid: bool, errors: List[str]):
        """Stocke le résultat de validation dans le cache"""
        with self._lock:
            try:
                # Nettoyer le cache si nécessaire
                if len(self._cache) >= self.max_size:
                    self._evict_lru()
                
                stat = os.stat(file_path)
                cache_key = f"{file_path}_{stat.st_mtime}_{stat.st_size}"
                self._cache[cache_key] = (is_valid, errors)
                self._access_times[cache_key] = time.time()
            except (OSError, IOError):
                pass
    
    def _evict_lru(self):
        """Supprime l'élément le moins récemment utilisé"""
        if not self._access_times:
            return
        
        lru_key = min(self._access_times.keys(), key=lambda k: self._access_times[k])
        self._cache.pop(lru_key, None)
        self._access_times.pop(lru_key, None)
    
    def clear(self):
        """Vide le cache"""
        with self._lock:
            self._cache.clear()
            self._access_times.clear()


class SimilarityDetector:
    """Détecteur de similarité entre fichiers pour optimisation"""
    
    def __init__(self):
        self._fingerprints = {}
        self._lock = threading.Lock()
    
    def get_file_fingerprint(self, file_path: str) -> str:
        """Génère une empreinte du fichier basée sur sa structure"""
        try:
            with self._lock:
                # Utiliser les 100 premiers octets + taille + extension comme empreinte
                with open(file_path, 'rb') as f:
                    first_bytes = f.read(100)
                
                stat = os.stat(file_path)
                extension = os.path.splitext(file_path)[1].lower()
                
                fingerprint_data = f"{first_bytes.hex()}_{stat.st_size}_{extension}"
                return hashlib.md5(fingerprint_data.encode()).hexdigest()
        except (OSError, IOError):
            return ""
    
    def find_similar_files(self, file_path: str, file_list: List[str]) -> List[str]:
        """Trouve les fichiers similaires dans la liste"""
        target_fingerprint = self.get_file_fingerprint(file_path)
        if not target_fingerprint:
            return []
        
        similar_files = []
        for other_file in file_list:
            if other_file != file_path:
                other_fingerprint = self.get_file_fingerprint(other_file)
                if other_fingerprint == target_fingerprint:
                    similar_files.append(other_file)
        
        return similar_files
    
    def clear(self):
        """Vide le cache d'empreintes"""
        with self._lock:
            self._fingerprints.clear()


class FileMetadataCache:
    """Cache intelligent pour les métadonnées des fichiers"""
    
    def __init__(self):
        self._cache = {}
        self._lock = threading.Lock()
        
    def get_metadata(self, file_path: str) -> Optional[Dict]:
        """Récupère les métadonnées du cache"""
        with self._lock:
            stat = os.stat(file_path)
            cache_key = f"{file_path}_{stat.st_mtime}_{stat.st_size}"
            return self._cache.get(cache_key)
    
    def set_metadata(self, file_path: str, metadata: Dict):
        """Stocke les métadonnées dans le cache"""
        with self._lock:
            stat = os.stat(file_path)
            cache_key = f"{file_path}_{stat.st_mtime}_{stat.st_size}"
            self._cache[cache_key] = metadata
    
    def clear(self):
        """Vide le cache"""
        with self._lock:
            self._cache.clear()

class ChunkedFileReader:
    """Lecteur de fichiers par chunks pour optimiser la mémoire"""
    
    def __init__(self, chunk_size: int = CHUNK_SIZE):
        self.chunk_size = chunk_size
        self.memory_manager = MemoryManager()
    
    def read_excel_chunks(self, file_path: str, start_row: int = 1) -> Generator[List[List], None, None]:
        """Lit un fichier Excel par chunks"""
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            ws = wb.active
            
            current_chunk = []
            row_count = 0
            
            for row in ws.iter_rows(min_row=start_row):
                if self.memory_manager.is_memory_limit_reached():
                    self.memory_manager.force_garbage_collection()
                
                row_data = [cell.value for cell in row]
                current_chunk.append(row_data)
                row_count += 1
                
                if row_count >= self.chunk_size:
                    yield current_chunk
                    current_chunk = []
                    row_count = 0
            
            if current_chunk:
                yield current_chunk
                
        finally:
            wb.close()
            del wb
            gc.collect()
    
    def read_csv_chunks(self, file_path: str, delimiter: str = ',', encoding: str = 'utf-8', 
                       skip_rows: int = 0) -> Generator[List[List], None, None]:
        """Lit un fichier CSV par chunks"""
        try:
            chunk_iter = pd.read_csv(
                file_path, 
                delimiter=delimiter,
                encoding=encoding,
                header=None,
                skiprows=skip_rows,
                chunksize=self.chunk_size,
                low_memory=True
            )
            
            for chunk_df in chunk_iter:
                if self.memory_manager.is_memory_limit_reached():
                    self.memory_manager.force_garbage_collection()
                
                chunk_data = chunk_df.values.tolist()
                yield chunk_data
                
                del chunk_df
                
        except Exception as e:
            logging.error(f"Erreur lors de la lecture par chunks de {file_path}: {e}")
            raise

class AdvancedFileDetector:
    """Détecteur avancé pour encodages et délimiteurs de fichiers texte"""
    
    def __init__(self, logger: Optional['AdvancedLogger'] = None):
        self.custom_logger = logger
        
    def detect_encoding(self, file_path: str, sample_size: int = 8192) -> str:
        """
        Détecte l'encodage d'un fichier avec plusieurs stratégies.
        """
        try:
            # Essayer chardet pour une détection automatique
            import chardet
            
            with open(file_path, 'rb') as f:
                raw_data = f.read(sample_size)
                
            detection = chardet.detect(raw_data)
            detected_encoding = detection.get('encoding')
            confidence = detection.get('confidence', 0)
            
            if logging:
                logging.info(f"Encodage détecté: {detected_encoding} (confiance: {confidence:.2f})")
            
            # Si la confiance est élevée, utiliser l'encodage détecté
            if confidence > 0.8 and detected_encoding:
                # Normaliser certains encodages
                encoding_map = {
                    'ascii': 'utf-8',
                    'windows-1252': 'cp1252',
                    'iso-8859-1': 'latin-1'
                }
                return encoding_map.get(detected_encoding.lower(), detected_encoding)
                
        except ImportError:
            if logging:
                logging.warning("chardet non disponible, utilisation des encodages par défaut")
        except Exception as e:
            if logging:
                logging.warning(f"Erreur lors de la détection d'encodage: {e}")
        
        # Fallback : tester les encodages courants
        return self._fallback_encoding_detection(file_path, sample_size)
    
    def _fallback_encoding_detection(self, file_path: str, sample_size: int) -> str:
        """
        Détection d'encodage par tentatives successives.
        """
        encodings_to_try = [
            'utf-8-sig',  # UTF-8 avec BOM
            'utf-8',      # UTF-8 standard
            'cp1252',     # Windows-1252 (Western European)
            'latin-1',    # ISO-8859-1
            'cp850',      # Code page 850 (Western European)
            'utf-16',     # UTF-16 avec BOM
            'cp437'       # Code page 437 (US)
        ]
        
        for encoding in encodings_to_try:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    f.read(sample_size)
                if logging:
                    logging.info(f"Encodage {encoding} validé par test de lecture")
                return encoding
            except (UnicodeDecodeError, UnicodeError):
                continue
            except Exception:
                continue
        
        # Si rien ne fonctionne, retourner utf-8 par défaut
        if logging:
            logging.warning("Aucun encodage détecté, utilisation d'utf-8 par défaut")
        return 'utf-8'
    
    def detect_delimiter(self, file_path: str, encoding: str = 'utf-8', sample_lines: int = 10) -> str:
        """
        Détecte automatiquement le délimiteur d'un fichier CSV.
        """
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                sample_text = ""
                for i, line in enumerate(f):
                    if i >= sample_lines:
                        break
                    sample_text += line
            
            if not sample_text.strip():
                return ','  # Défaut si fichier vide
            
            # Utiliser csv.Sniffer pour la détection
            sniffer = csv.Sniffer()
            
            # Délimiteurs à tester en priorité
            delimiters_to_test = [',', ';', '\t', '|', ':', ' ']
            
            try:
                dialect = sniffer.sniff(sample_text, delimiters=''.join(delimiters_to_test))
                detected_delimiter = dialect.delimiter
                
                if logging:
                    delimiter_name = self._get_delimiter_name(detected_delimiter)
                    logging.info(f"Délimiteur détecté: {delimiter_name} ('{detected_delimiter}')")
                
                return detected_delimiter
                
            except csv.Error:
                # Si Sniffer échoue, analyser manuellement
                return self._manual_delimiter_detection(sample_text)
                
        except Exception as e:
            if logging:
                logging.warning(f"Erreur lors de la détection de délimiteur: {e}")
            return ','  # Retourner virgule par défaut
    
    def _manual_delimiter_detection(self, sample_text: str) -> str:
        """
        Détection manuelle du délimiteur en comptant les occurrences.
        """
        delimiters = {
            ',': 'virgule',
            ';': 'point-virgule', 
            '\t': 'tabulation',
            '|': 'pipe',
            ':': 'deux-points',
            ' ': 'espace'
        }
        
        lines = sample_text.strip().split('\n')
        if len(lines) < 2:
            return ','
        
        delimiter_scores = {}
        
        for delimiter, name in delimiters.items():
            # Compter les occurrences dans chaque ligne
            counts = [line.count(delimiter) for line in lines[:5]]  # Tester 5 premières lignes
            
            if not counts or max(counts) == 0:
                delimiter_scores[delimiter] = 0
                continue
            
            # Un bon délimiteur a un nombre cohérent d'occurrences par ligne
            avg_count = sum(counts) / len(counts)
            variance = sum((c - avg_count) ** 2 for c in counts) / len(counts)
            
            # Score basé sur la moyenne et la cohérence
            if avg_count > 0:
                consistency = 1 / (1 + variance)  # Plus la variance est faible, mieux c'est
                delimiter_scores[delimiter] = avg_count * consistency
            else:
                delimiter_scores[delimiter] = 0
        
        # Retourner le délimiteur avec le meilleur score
        best_delimiter = max(delimiter_scores, key=delimiter_scores.get)
        
        if logging:
            delimiter_name = self._get_delimiter_name(best_delimiter)
            score = delimiter_scores[best_delimiter]
            logging.info(f"Délimiteur choisi par analyse: {delimiter_name} (score: {score:.2f})")
        
        return best_delimiter if delimiter_scores[best_delimiter] > 0 else ','
    
    def _get_delimiter_name(self, delimiter: str) -> str:
        """Retourne le nom lisible d'un délimiteur."""
        delimiter_names = {
            ',': 'virgule',
            ';': 'point-virgule',
            '\t': 'tabulation',
            '|': 'pipe',
            ':': 'deux-points',
            ' ': 'espace'
        }
        return delimiter_names.get(delimiter, f"'{delimiter}'")
    
    def validate_delimiter(self, delimiter: str) -> bool:
        """
        Valide qu'un délimiteur personnalisé est utilisable.
        """
        if not delimiter:
            return False
        
        # Caractères interdits (qui casseraient le parsing CSV)
        forbidden_chars = ['"', "'", '\n', '\r']
        
        if any(char in delimiter for char in forbidden_chars):
            return False
        
        # Délimiteur trop long (plus de 3 caractères est suspect)
        if len(delimiter) > 3:
            return False
            
        return True

class CancellationToken:
    """Token d'annulation pour interrompre les opérations en cours"""
    
    def __init__(self):
        self._cancelled = threading.Event()
        self._lock = threading.Lock()
    
    def cancel(self):
        """Demande l'annulation"""
        with self._lock:
            self._cancelled.set()
    
    def is_cancelled(self) -> bool:
        """Vérifie si l'annulation a été demandée"""
        return self._cancelled.is_set()
    
    def throw_if_cancelled(self):
        """Lève une exception si l'annulation a été demandée"""
        if self.is_cancelled():
            raise InterruptedError("Opération annulée par l'utilisateur")
    
    def reset(self):
        """Remet le token à zéro"""
        with self._lock:
            self._cancelled.clear()

class AdvancedLogger:
    """Système de logging avancé avec rotation automatique"""
    
    def __init__(self, log_dir: str = "logs", max_bytes: int = 10*1024*1024, backup_count: int = 5):
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(exist_ok=True)
        
        # Configuration du logger principal
        self.logger = logging.getLogger('ExcelCompiler')
        self.logger.setLevel(logging.DEBUG)
        
        # Éviter les handlers dupliqués
        if not self.logger.handlers:
            # Handler pour fichier avec rotation
            log_file = self.log_dir / 'excel_compiler.log'
            file_handler = RotatingFileHandler(
                log_file, maxBytes=max_bytes, backupCount=backup_count
            )
            file_handler.setLevel(logging.DEBUG)
            
            # Handler pour console
            console_handler = logging.StreamHandler()
            console_handler.setLevel(logging.INFO)
            
            # Format des logs
            formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - [%(threadName)s] - %(message)s'
            )
            file_handler.setFormatter(formatter)
            console_handler.setFormatter(formatter)
            
            self.logger.addHandler(file_handler)
            self.logger.addHandler(console_handler)
    
    def debug(self, message: str, **kwargs):
        """Log de débogage"""
        self.logger.debug(message, extra=kwargs)
    
    def info(self, message: str, **kwargs):
        """Log d'information"""
        self.logger.info(message, extra=kwargs)
    
    def warning(self, message: str, **kwargs):
        """Log d'avertissement"""
        self.logger.warning(message, extra=kwargs)
    
    def error(self, message: str, **kwargs):
        """Log d'erreur"""
        self.logger.error(message, extra=kwargs)
    
    def critical(self, message: str, **kwargs):
        """Log critique"""
        self.logger.critical(message, extra=kwargs)

class ErrorRecoveryManager:
    """Gestionnaire de récupération d'erreurs pour continuer le traitement"""
    
    def __init__(self, logger=None):
        # Utiliser le logger standard de Python
        pass
        self.error_statistics = defaultdict(int)
        self.failed_files = []
        self.recovery_strategies = {
            'encoding_error': self._handle_encoding_error,
            'permission_error': self._handle_permission_error,
            'file_not_found': self._handle_file_not_found,
            'memory_error': self._handle_memory_error,
            'format_error': self._handle_format_error
        }
    
    def handle_error(self, error: Exception, context: Dict[str, Any]) -> bool:
        """
        Gère une erreur et tente une récupération
        Returns: True si le traitement peut continuer, False sinon
        """
        error_type = self._classify_error(error)
        self.error_statistics[error_type] += 1
        
        file_path = context.get('file_path', 'unknown')
        logging.error(f"Erreur {error_type} dans {file_path}: {str(error)}")
        
        # Tenter une stratégie de récupération
        if error_type in self.recovery_strategies:
            try:
                return self.recovery_strategies[error_type](error, context)
            except Exception as recovery_error:
                logging.error(f"Échec de récupération pour {error_type}: {str(recovery_error)}")
        
        # Ajouter à la liste des fichiers échoués
        self.failed_files.append({
            'file': file_path,
            'error_type': error_type,
            'error_message': str(error),
            'timestamp': datetime.now()
        })
        
        return True  # Continuer le traitement des autres fichiers
    
    def _classify_error(self, error: Exception) -> str:
        """Classifie le type d'erreur"""
        if isinstance(error, UnicodeDecodeError):
            return 'encoding_error'
        elif isinstance(error, PermissionError):
            return 'permission_error'
        elif isinstance(error, FileNotFoundError):
            return 'file_not_found'
        elif isinstance(error, MemoryError):
            return 'memory_error'
        elif 'format' in str(error).lower() or 'corrupt' in str(error).lower():
            return 'format_error'
        else:
            return 'unknown_error'
    
    def _handle_encoding_error(self, error: Exception, context: Dict) -> bool:
        """Gère les erreurs d'encodage"""
        file_path = context.get('file_path')
        if file_path and file_path.endswith('.csv'):
            # Essayer différents encodages
            for encoding in ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']:
                try:
                    context['suggested_encoding'] = encoding
                    logging.info(f"Tentative avec encodage {encoding} pour {file_path}")
                    return True
                except (UnicodeDecodeError, LookupError) as e:
                    # UnicodeDecodeError: Problème décodage
                    # LookupError: Encodage non reconnu
                    logging.debug(f"Échec encodage {encoding}: {e}")
                    continue
        return False
    
    def _handle_permission_error(self, error: Exception, context: Dict) -> bool:
        """Gère les erreurs de permission"""
        logging.warning("Fichier ignoré à cause des permissions insuffisantes")
        return True  # Continuer avec les autres fichiers
    
    def _handle_file_not_found(self, error: Exception, context: Dict) -> bool:
        """Gère les fichiers introuvables"""
        logging.warning("Fichier ignoré car introuvable")
        return True  # Continuer avec les autres fichiers
    
    def _handle_memory_error(self, error: Exception, context: Dict) -> bool:
        """Gère les erreurs de mémoire"""
        logging.warning("Erreur mémoire - tentative de nettoyage")
        gc.collect()
        return True
    
    def _handle_format_error(self, error: Exception, context: Dict) -> bool:
        """Gère les erreurs de format"""
        logging.warning("Format de fichier non supporté ou corrompu")
        return True
    
    def get_summary(self) -> Dict[str, Any]:
        """Retourne un résumé des erreurs rencontrées"""
        return {
            'error_statistics': dict(self.error_statistics),
            'failed_files': self.failed_files,
            'total_errors': sum(self.error_statistics.values())
        }

class InputValidator:
    """Validateur strict des paramètres utilisateur"""
    
    def __init__(self, logger=None):
        # Utiliser le logger standard de Python
        pass
    
    def validate_files(self, files: List[str], directory: str) -> Tuple[bool, List[str]]:
        """Valide la liste des fichiers"""
        errors = []
        
        if not files:
            errors.append("Aucun fichier sélectionné")
            return False, errors
        
        if not os.path.isdir(directory):
            errors.append(f"Répertoire inexistant: {directory}")
            return False, errors
        
        valid_files = []
        for file in files:
            file_path = os.path.join(directory, file)
            
            # Vérifier l'existence
            if not os.path.exists(file_path):
                errors.append(f"Fichier inexistant: {file}")
                continue
            
            # Vérifier les permissions
            if not os.access(file_path, os.R_OK):
                errors.append(f"Permissions insuffisantes: {file}")
                continue
            
            # Vérifier l'extension
            ext = os.path.splitext(file)[1].lower()
            if ext not in SUPPORTED_EXCEL_EXTENSIONS + SUPPORTED_TEXT_EXTENSIONS:
                errors.append(f"Format non supporté: {file} ({ext})")
                continue
            
            # Vérifier la taille (limite à 100MB par fichier)
            size = os.path.getsize(file_path)
            if size > 100 * 1024 * 1024:
                errors.append(f"Fichier trop volumineux: {file} ({size / (1024*1024):.1f}MB)")
                continue
            
            valid_files.append(file)
        
        if not valid_files:
            errors.append("Aucun fichier valide trouvé")
            return False, errors
        
        return True, []
    
    def validate_parameters(self, params: Dict[str, Any]) -> Tuple[bool, List[str]]:
        """Valide les paramètres de compilation"""
        errors = []
        
        # Validation des lignes d'en-tête
        header_start = params.get('header_start_row', 1)
        if not isinstance(header_start, int) or header_start < 1:
            errors.append("La ligne de début doit être un entier >= 1")
        
        header_rows = params.get('header_rows', 1)
        if not isinstance(header_rows, int) or header_rows < 1:
            errors.append("Le nombre de lignes d'en-tête doit être >= 1")
        
        # Validation du nom de fichier de sortie
        output_file = params.get('output_file', '')
        if not output_file:
            errors.append("Nom de fichier de sortie requis")
        elif not output_file.endswith('.xlsx'):
            errors.append("Le fichier de sortie doit avoir l'extension .xlsx")
        elif not self._is_valid_filename(output_file):
            errors.append("Nom de fichier invalide (caractères interdits)")
        
        # Validation des options
        filename_option = params.get('filename_option', 'none')
        if filename_option not in ['none', 'with_extension', 'without_extension']:
            errors.append("Option de nom de fichier invalide")
        
        sort_column = params.get('sort_column', 0)
        if not isinstance(sort_column, int) or sort_column < 0:
            errors.append("La colonne de tri doit être un entier >= 0")
        
        return len(errors) == 0, errors
    
    def _is_valid_filename(self, filename: str) -> bool:
        """Vérifie si le nom de fichier est valide"""
        # Caractères interdits dans les noms de fichiers
        invalid_chars = r'[<>:"/\\|?*]'
        return not re.search(invalid_chars, filename)

# Constants for styles
COLORS = {
    "PRIMARY": "2e7d32",
    "PRIMARY_DARK": "2e7d32",
    "PRIMARY_LIGHT": "4caf50",  
    "ACCENT": "c8e6c9",
    "BACKGROUND": "2e7d32",
    "LIGHT_TEXT": "FFFFFF",
    "DARK_TEXT": "212121",
    "BORDER": "e0e0e0",
    "WARNING": "f44336",
    "SUCCESS": "4CAF50",
    "INFO": "2196F3"
}

EXCEL_COLORS = {
    "PRIMARY": "FF2e7d32",
    "PRIMARY_DARK": "FF2e7d32",
    "PRIMARY_LIGHT": "FF4caf50",
    "ACCENT": "FFc8e6c9",
    "BACKGROUND":"FF2e7d32",
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

# Formats de date disponibles
DATE_FORMATS = {
    "STANDARD": {"format": "yyyy-MM-dd", "code": "yyyy-mm-dd", "excel_format": "yyyy-mm-dd"},
    "FRENCH": {"format": "dd/MM/yyyy", "code": "dd/mm/yyyy", "excel_format": "dd/mm/yyyy"},
    "US": {"format": "MM/dd/yyyy", "code": "mm/dd/yyyy", "excel_format": "mm/dd/yyyy"},
    "DATETIME": {"format": "yyyy-MM-dd HH:mm:ss", "code": "yyyy-mm-dd hh:mm:ss", "excel_format": "yyyy-mm-dd hh:mm:ss"},
    "DATETIME_FRENCH": {"format": "dd/MM/yyyy HH:mm:ss", "code": "dd/mm/yyyy hh:mm:ss", "excel_format": "dd/mm/yyyy hh:mm:ss"},
    "DATE_ONLY": {"format": "yyyy-MM-dd", "code": "yyyy-mm-dd", "excel_format": "yyyy-mm-dd"},
    "TIME_ONLY": {"format": "HH:mm:ss", "code": "hh:mm:ss", "excel_format": "hh:mm:ss"},
    "SHORT": {"format": "dd/MM/yy", "code": "dd/mm/yy", "excel_format": "dd/mm/yy"},
    "CUSTOM": {"format": "", "code": "", "excel_format": ""}
}


# Configuration du logging avec gestion d'erreurs
def setup_logging():
    """Configure le logging avec gestion robuste des erreurs."""
    try:
        # Tentative de création du fichier de log
        logging.basicConfig(
            level=logging.ERROR,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('error.log', mode='w', encoding='utf-8'),
                logging.StreamHandler()  # ✅ AUSSI dans la console
            ]
        )
        print("✅ Logging configuré avec fichier error.log")
    except (PermissionError, OSError) as e:
        # ✅ FALLBACK : Log uniquement dans la console
        logging.basicConfig(
            level=logging.ERROR,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        print(f"⚠️ Impossible de créer error.log ({e}), logs en console uniquement")

# Dictionnaires pour l'internationalisation
TRANSLATIONS = {
    "fr": {
        "app_title": "Compilateur Excel Professionnel",
        "file_selection": "Sélection des fichiers",
        "compilation_options": "Options de compilation",
        "advanced_options": "Options avancées",
        "help": "Aide",
        "about": "À propos",
        "preview": "Aperçu",
        "date_format": "Format de date",
        "languages": "Langues",
        "no_directory": "Aucun répertoire sélectionné",
        "choose_directory": "Choisir un répertoire",
        "select_all_files": "Sélectionner tous les fichiers",
        "files_selected": "{} fichier(s) sélectionné(s)",
        "date_time": "Date et heure : {}",
        "header_start_row": "Ligne de début des en-têtes :",
        "header_rows": "Nombre de lignes d'en-tête :",
        "repeat_headers": "Répéter les en-têtes pour chaque fichier",
        "merge_headers": "Fusionner les en-têtes multi-niveaux",
        "add_filename": "Ajouter les noms des fichiers sources",
        "preliminary_info": "Informations de début du fichier de sortie (titre et métadonnées)",
        "include_preliminary": "Inclure titre et métadonnées dans fichier de sortie",
        "preliminary_source_file": "Source (titre et métadonnées) :",
        "no_preliminary_lines": "Aucune information de début trouvée dans ce fichier",
        "preliminary_preview": "Aperçu des informations qui seront copiées :",
        "output_filename": "Nom du fichier de sortie :",
        "enable_verification": "Activer la vérification préliminaire des fichiers",
        "start_compilation": "Lancer la compilation",
        "remove_duplicates": "Supprimer les doublons",
        "remove_empty_rows": "Supprimer les lignes entièrement vides",
        "sort_data": "Trier les données",
        "sort_column": "Colonne de tri (ex: A, B, C) :",
        "auto_width": "Ajuster automatiquement la largeur des colonnes",
        "freeze_headers": "Figer les en-têtes",
        "file_formats": "Formats de fichiers supportés",
        "excel_files": "Fichiers Excel (.xlsx, .xlsm, .xltx, .xltm, .xls)",
        "text_files": "Fichiers texte (.csv, .tsv, .txt)",
        "preview_data": "Prévisualiser les données avant compilation",
        "refresh_preview": "Actualiser l'aperçu",
        "preview_limited": "Aperçu limité aux {} premières lignes",
        "date_format_options": "Options de format de date",
        "date_format_standard": "Standard (AAAA-MM-JJ)",
        "date_format_french": "Français (JJ/MM/AAAA)",
        "date_format_us": "Américain (MM/JJ/AAAA)",
        "date_format_datetime": "Date et heure (AAAA-MM-JJ HH:MM:SS)",
        "date_format_datetime_french": "Date et heure française (JJ/MM/AAAA HH:MM:SS)",
        "date_format_date_only": "Date uniquement (AAAA-MM-JJ)",
        "date_format_time_only": "Heure uniquement (HH:MM:SS)",
        "date_format_short": "Format court (JJ/MM/AA)",
        "date_format_custom": "Personnalisé :",
        "language": "Langue",
        "french": "Français",
        "english": "Anglais",
        "spanish": "Espagnol",
        "german": "Allemand",
        "apply": "Appliquer",
        "cancel": "Annuler",
        "ok": "OK",
        "error": "Erreur",
        "success": "Succès",
        "warning": "Avertissement",
        "no_files_selected": "Aucun fichier sélectionné",
        "select_files_message": "Veuillez sélectionner au moins un fichier à compiler.",
        "compilation_in_progress": "Compilation en cours...",
        "compilation_complete": "Compilation terminée. Fichier enregistré : {}",
        "compilation_failed": "Échec: {}",
        "no_data": "Aucune donnée à écrire.",
        "no_data_message": "Aucune donnée valide n'a pu être compilée.",
        "file_open_error": "Le fichier de sortie est ouvert dans une autre application.",
        "file_open_error_message": "Le fichier de sortie est ouvert dans une autre application. Veuillez le fermer et réessayer.",
        "verification_title": "Rapport de vérification des fichiers",
        "verification_header": "Vérification de {} fichiers",
        "compilable": "Compilables: {} ({}%)",
        "not_compilable": "Non compilables: {} ({}%)",
        "non_compilable_files": "Fichiers non compilables",
        "compilable_files": "Fichiers compilables",
        "error_resolution_tips": "Ces fichiers ne peuvent pas être compilés pour les raisons indiquées. Voici quelques conseils pour résoudre les problèmes courants :",
        "open_file_tip": "Fichier ouvert: Fermez le fichier dans Excel et réessayez",
        "protected_file_tip": "Fichier protégé: Désactivez la protection dans Excel (Révision > Protéger la feuille)",
        "header_structure_tip": "Structure d'en-tête incompatible: Vérifiez que le nombre de lignes d'en-tête est correct",
        "encoding_error_tip": "Erreur d'encodage: Réenregistrez le fichier CSV avec l'encodage UTF-8",
        "ignore_non_compilable": "Ignorer les non compilables et compiler",
        "filename": "Nom du fichier",
        "filename_option": "Colonne nom de fichier :",
        "filename_none": "Ne pas ajouter",
        "filename_with_extension": "Ajouter avec extension",
        "filename_without_extension": "Ajouter sans extension",
        "detected_issue": "Problème détecté",
        "compilation_report": "Rapport de compilation",
        "compilation_result": "Résultat de la compilation",
        "generated_file": "Fichier généré:",
        "compilation_rate": "Taux de compilation: {}% ({}/{})",
        "compiled_files": "Fichiers compilés ({})",
        "not_compiled_files": "Fichiers non compilés ({})",
        "all_files_compiled": "Tous les fichiers ont été compilés avec succès !",
        "ip_warning_title": "Propriété Intellectuelle - Avertissement",
        "ip_warning_content": "Cette application <b>Compilateur Excel Professionnel</b> est la propriété intellectuelle exclusive de:",
        "developer_name": "GOUNOU N'GOBI Chabi Zimé",
        "developer_title": "Data Manager & Data Analyst",
        "copyright_notice": "Tous droits réservés. Cette application est protégée par les lois sur le droit d'auteur et les traités internationaux sur la propriété intellectuelle.",
        "warning_important": "AVERTISSEMENT IMPORTANT:",
        "unauthorized_reproduction": "Toute reproduction non autorisée, distribution ou modification de cette application est strictement interdite.",
        "license_terms": "L'utilisation de cette application est soumise aux termes de la licence accordée par l'auteur.",
        "contact_info": "Pour toute question concernant les droits d'utilisation ou pour signaler une violation de la propriété intellectuelle, veuillez contacter l'auteur à l'adresse : zimkada@gmail.com.",
        "accept_conditions": "J'accepte ces conditions",
        "quit_app": "Quitter l'application",
        "file_menu": "Fichier",
        "edit_menu": "Édition",
        "tools_menu": "Outils",
        "language_menu": "Langue",
        "help_menu": "Aide",
        "open_dir": "Ouvrir un répertoire...",
        "save_settings": "Enregistrer les paramètres",
        "load_settings": "Charger les paramètres",
        "exit": "Quitter",
        "select_all": "Sélectionner tout",
        "deselect_all": "Désélectionner tout",
        "invert_selection": "Inverser la sélection",
        "preview_tool": "Aperçu des données",
        "compile_tool": "Compiler les fichiers",
        "about_help": "À propos...",
        "settings_saved": "Paramètres sauvegardés avec succès",
        "settings_loaded": "Paramètres chargés avec succès",
        "tutorial_assistant": "Assistant de démarrage",
        "user_guide": "Guide utilisateur",
        "thread_safety": "Sécurité thread",
        "unit_tests": "Tests unitaires",
        "performance": "Performance optimisée",
        "responsive_design": "Interface responsive",
        "advanced_validation": "Validation avancée",
        "cancellation_system": "Système d'annulation",
        "memory_management": "Gestion mémoire intelligente",
        "file_detection": "Détection avancée de fichiers",
        "error_recovery": "Récupération d'erreurs robuste",
        "timeout_title": "Délai d'attente dépassé",
        "timeout_message": "La compilation prend plus de temps que prévu ({} minutes). Voulez-vous prolonger le délai ou annuler ?"
    },
    "en": {
        "app_title": "Professional Excel Compiler",
        "file_selection": "File Selection",
        "compilation_options": "Compilation Options",
        "advanced_options": "Advanced Options",
        "help": "Help",
        "about": "About",
        "preview": "Preview",
        "date_format": "Date Format",
        "languages": "Languages",
        "no_directory": "No directory selected",
        "choose_directory": "Choose directory",
        "select_all_files": "Select all files",
        "files_selected": "{} file(s) selected",
        "date_time": "Date and time: {}",
        "header_start_row": "Header start row:",
        "header_rows": "Number of header rows:",
        "repeat_headers": "Repeat headers for each file",
        "merge_headers": "Merge multi-level headers",
        "add_filename": "Add source filenames",
        "preliminary_info": "Output file header information (title and metadata)",
        "include_preliminary": "Include title and metadata in output file",
        "preliminary_source_file": "Source file (title and metadata):",
        "no_preliminary_lines": "No header information found in this file",
        "preliminary_preview": "Preview of information to be copied:",
        "output_filename": "Output filename:",
        "enable_verification": "Enable preliminary file verification",
        "start_compilation": "Start compilation",
        "remove_duplicates": "Remove duplicates",
        "remove_empty_rows": "Remove empty rows",
        "sort_data": "Sort data",
        "sort_column": "Sort column (e.g. A, B, C):",
        "auto_width": "Auto-adjust column width",
        "freeze_headers": "Freeze headers",
        "file_formats": "Supported file formats",
        "excel_files": "Excel files (.xlsx, .xlsm, .xltx, .xltm, .xls)",
        "text_files": "Text files (.csv, .tsv, .txt)",
        "preview_data": "Preview data before compilation",
        "refresh_preview": "Refresh preview",
        "preview_limited": "Preview limited to first {} rows",
        "date_format_options": "Date format options",
        "date_format_standard": "Standard (YYYY-MM-DD)",
        "date_format_french": "French (DD/MM/YYYY)",
        "date_format_us": "US (MM/DD/YYYY)",
        "date_format_datetime": "Date and time (YYYY-MM-DD HH:MM:SS)",
        "date_format_datetime_french": "French date and time (DD/MM/YYYY HH:MM:SS)",
        "date_format_date_only": "Date only (YYYY-MM-DD)",
        "date_format_time_only": "Time only (HH:MM:SS)",
        "date_format_short": "Short format (DD/MM/YY)",
        "date_format_custom": "Custom:",
        "language": "Language",
        "french": "French",
        "english": "English",
        "spanish": "Spanish",
        "german": "German",
        "apply": "Apply",
        "cancel": "Cancel",
        "ok": "OK",
        "error": "Error",
        "success": "Success",
        "warning": "Warning",
        "no_files_selected": "No files selected",
        "select_files_message": "Please select at least one file to compile.",
        "compilation_in_progress": "Compilation in progress...",
        "compilation_complete": "Compilation complete. File saved: {}",
        "compilation_failed": "Failed: {}",
        "no_data": "No data to write.",
        "no_data_message": "No valid data could be compiled.",
        "file_open_error": "The output file is open in another application.",
        "file_open_error_message": "The output file is open in another application. Please close it and try again.",
        "verification_title": "File Verification Report",
        "verification_header": "Verification of {} files",
        "compilable": "Compilable: {} ({}%)",
        "not_compilable": "Not compilable: {} ({}%)",
        "non_compilable_files": "Non-compilable files",
        "compilable_files": "Compilable files",
        "error_resolution_tips": "These files cannot be compiled for the reasons indicated. Here are some tips to resolve common problems:",
        "open_file_tip": "Open file: Close the file in Excel and try again",
        "protected_file_tip": "Protected file: Disable protection in Excel (Review > Protect Sheet)",
        "header_structure_tip": "Incompatible header structure: Check that the number of header rows is correct",
        "encoding_error_tip": "Encoding error: Resave the CSV file with UTF-8 encoding",
        "ignore_non_compilable": "Ignore non-compilable and compile",
        "filename": "Filename",
        "filename_option": "Filename column:",
        "filename_none": "Don't add",
        "filename_with_extension": "Add with extension", 
        "filename_without_extension": "Add without extension",
        "detected_issue": "Detected issue",
        "compilation_report": "Compilation Report",
        "compilation_result": "Compilation Result",
        "generated_file": "Generated file:",
        "compilation_rate": "Compilation rate: {}% ({}/{})",
        "compiled_files": "Compiled files ({})",
        "not_compiled_files": "Files not compiled ({})",
        "all_files_compiled": "All files were successfully compiled!",
        "ip_warning_title": "Intellectual Property - Warning",
        "ip_warning_content": "This <b>Professional Excel Compiler</b> application is the exclusive intellectual property of:",
        "developer_name": "GOUNOU N'GOBI Chabi Zimé",
        "developer_title": "Data Manager & Data Analyst",
        "copyright_notice": "All rights reserved. This application is protected by copyright laws and international intellectual property treaties.",
        "warning_important": "IMPORTANT WARNING:",
        "unauthorized_reproduction": "Any unauthorized reproduction, distribution or modification of this application is strictly prohibited.",
        "license_terms": "Use of this application is subject to the terms of the license granted by the author.",
        "contact_info": "For any questions regarding usage rights or to report a violation of intellectual property, please contact the author at: zimkada@gmail.com.",
        "accept_conditions": "I accept these conditions",
        "quit_app": "Quit application",
        "file_menu": "File",
        "edit_menu": "Edit",
        "tools_menu": "Tools",
        "language_menu": "Language",
        "help_menu": "Help",
        "open_dir": "Open directory...",
        "save_settings": "Save settings",
        "load_settings": "Load settings",
        "exit": "Exit",
        "select_all": "Select all",
        "deselect_all": "Deselect all",
        "invert_selection": "Invert selection",
        "preview_tool": "Data preview",
        "compile_tool": "Compile files",
        "about_help": "About...",
        "settings_saved": "Settings saved successfully",
        "settings_loaded": "Settings loaded successfully",
        "tutorial_assistant": "Getting Started Assistant",
        "user_guide": "User Guide",
        "thread_safety": "Thread Safety",
        "unit_tests": "Unit Tests",
        "performance": "Optimized Performance",
        "responsive_design": "Responsive Interface",
        "advanced_validation": "Advanced Validation",
        "cancellation_system": "Cancellation System",
        "memory_management": "Smart Memory Management",
        "file_detection": "Advanced File Detection",
        "error_recovery": "Robust Error Recovery",
        "timeout_title": "Timeout Exceeded",
        "timeout_message": "Compilation is taking longer than expected ({} minutes). Do you want to extend the timeout or cancel?"
    },
    "es": {
        "app_title": "Compilador Excel Profesional",
        "file_selection": "Selección de Archivos",
        "compilation_options": "Opciones de Compilación",
        "advanced_options": "Opciones Avanzadas",
        "help": "Ayuda",
        "about": "Acerca de",
        "preview": "Vista Previa",
        "date_format": "Formato de Fecha",
        "languages": "Idiomas",
        "no_directory": "Ningún directorio seleccionado",
        "choose_directory": "Elegir directorio",
        "select_all_files": "Seleccionar todos los archivos",
        "files_selected": "{} archivo(s) seleccionado(s)",
        "date_time": "Fecha y hora: {}",
        "header_start_row": "Fila de inicio de encabezados:",
        "header_rows": "Número de filas de encabezado:",
        "repeat_headers": "Repetir encabezados para cada archivo",
        "merge_headers": "Fusionar encabezados multinivel",
        "add_filename": "Añadir nombres de archivos fuente",
        "start_compilation": "Iniciar compilación",
        "remove_duplicates": "Eliminar duplicados",
        "remove_empty_rows": "Eliminar filas vacías",
        "sort_data": "Ordenar datos",
        "sort_column": "Columna de ordenación (ej: A, B, C):",
        "file_formats": "Formatos de archivo soportados",
        "excel_files": "Archivos Excel (.xlsx, .xlsm, .xltx, .xltm, .xls)",
        "text_files": "Archivos de texto (.csv, .tsv, .txt)",
        "preview_data": "Vista previa de datos antes de compilar",
        "refresh_preview": "Actualizar vista previa",
        "language": "Idioma",
        "french": "Francés",
        "english": "Inglés",
        "spanish": "Español",
        "german": "Alemán",
        "apply": "Aplicar",
        "cancel": "Cancelar",
        "ok": "Aceptar",
        "error": "Error",
        "success": "Éxito",
        "warning": "Advertencia",
        "no_files_selected": "Ningún archivo seleccionado",
        "compilation_complete": "Compilación completada. Archivo guardado: {}",
        "compilation_failed": "Falló: {}",
        "developer_name": "GOUNOU N'GOBI Chabi Zimé",
        "developer_title": "Gestor de Datos y Analista de Datos",
        "tutorial_assistant": "Asistente de Inicio",
        "user_guide": "Guía del Usuario",
        "thread_safety": "Seguridad de Hilos",
        "unit_tests": "Pruebas Unitarias",
        "performance": "Rendimiento Optimizado",
        "responsive_design": "Interfaz Responsiva",
        "advanced_validation": "Validación Avanzada",
        "cancellation_system": "Sistema de Cancelación",
        "memory_management": "Gestión Inteligente de Memoria",
        "file_detection": "Detección Avanzada de Archivos",
        "error_recovery": "Recuperación Robusta de Errores",
        "timeout_title": "Tiempo de Espera Excedido",
        "timeout_message": "La compilación está tomando más tiempo del esperado ({} minutos). ¿Desea extender el tiempo o cancelar?"
    },
    "de": {
        "app_title": "Professioneller Excel-Compiler",
        "file_selection": "Dateiauswahl",
        "compilation_options": "Kompilierungsoptionen",
        "advanced_options": "Erweiterte Optionen",
        "help": "Hilfe",
        "about": "Über",
        "preview": "Vorschau",
        "date_format": "Datumsformat",
        "languages": "Sprachen",
        "no_directory": "Kein Verzeichnis ausgewählt",
        "choose_directory": "Verzeichnis auswählen",
        "select_all_files": "Alle Dateien auswählen",
        "files_selected": "{} Datei(en) ausgewählt",
        "date_time": "Datum und Uhrzeit: {}",
        "header_start_row": "Startreihe der Kopfzeilen:",
        "header_rows": "Anzahl der Kopfzeilen:",
        "repeat_headers": "Kopfzeilen für jede Datei wiederholen",
        "merge_headers": "Mehrstufige Kopfzeilen zusammenführen",
        "add_filename": "Quelldateinamen hinzufügen",
        "start_compilation": "Kompilierung starten",
        "remove_duplicates": "Duplikate entfernen",
        "remove_empty_rows": "Leere Zeilen entfernen",
        "sort_data": "Daten sortieren",
        "sort_column": "Sortierspalte (z.B: A, B, C):",
        "file_formats": "Unterstützte Dateiformate",
        "excel_files": "Excel-Dateien (.xlsx, .xlsm, .xltx, .xltm, .xls)",
        "text_files": "Textdateien (.csv, .tsv, .txt)",
        "preview_data": "Datenvorschau vor Kompilierung",
        "refresh_preview": "Vorschau aktualisieren",
        "language": "Sprache",
        "french": "Französisch",
        "english": "Englisch",
        "spanish": "Spanisch",
        "german": "Deutsch",
        "apply": "Anwenden",
        "cancel": "Abbrechen",
        "ok": "OK",
        "error": "Fehler",
        "success": "Erfolg",
        "warning": "Warnung",
        "no_files_selected": "Keine Dateien ausgewählt",
        "compilation_complete": "Kompilierung abgeschlossen. Datei gespeichert: {}",
        "compilation_failed": "Fehlgeschlagen: {}",
        "developer_name": "GOUNOU N'GOBI Chabi Zimé",
        "developer_title": "Datenmanager und Datenanalyst",
        "tutorial_assistant": "Einrichtungsassistent",
        "user_guide": "Benutzerhandbuch",
        "thread_safety": "Thread-Sicherheit",
        "unit_tests": "Unit-Tests",
        "performance": "Optimierte Leistung",
        "responsive_design": "Responsive Benutzeroberfläche",
        "advanced_validation": "Erweiterte Validierung",
        "cancellation_system": "Abbruchsystem",
        "memory_management": "Intelligente Speicherverwaltung",
        "file_detection": "Erweiterte Dateierkennung",
        "error_recovery": "Robuste Fehlerwiederherstellung",
        "timeout_title": "Timeout Überschritten",
        "timeout_message": "Die Kompilierung dauert länger als erwartet ({} Minuten). Möchten Sie das Timeout verlängern oder abbrechen?"
    }
}

class TranslationManager:
    """
    Gestionnaire de traduction pour l'application.
    Permet de changer la langue de l'interface utilisateur.
    """
    _instance = None
    
    def __new__(cls):
        """Implémentation du pattern Singleton."""
        if cls._instance is None:
            cls._instance = super(TranslationManager, cls).__new__(cls)
            cls._instance._initialize()
        return cls._instance
    
    def _initialize(self):
        """Initialise le gestionnaire de traduction avec détection automatique."""
        # Détection automatique de la langue système
        self.current_language = self._detect_system_language()
        self.translations = TRANSLATIONS
        self.language_changed_callbacks = []
        self.available_languages = {
            "fr": "Français",
            "en": "English", 
            "es": "Español",
            "de": "Deutsch"
        }
    
    def set_language(self, language_code):
        """
        Change la langue courante.
        
        Args:
            language_code: Code de la langue (fr, en, es, de)
        """
        if language_code in self.translations:
            self.current_language = language_code
            
            # Appeler tous les callbacks enregistrés
            for callback in self.language_changed_callbacks:
                callback()
    
    def get_text(self, key, *args):
        """
        Obtient le texte traduit pour une clé donnée.
        
        Args:
            key: Clé de traduction
            *args: Arguments de formatage optionnels
            
        Returns:
            Texte traduit
        """
        if key in self.translations[self.current_language]:
            text = self.translations[self.current_language][key]
            if args:
                return text.format(*args)
            return text
        return key
    
    def register_language_changed_callback(self, callback):
        """
        Enregistre un callback à appeler lorsque la langue change.
        
        Args:
            callback: Fonction à appeler
        """
        if callback not in self.language_changed_callbacks:
            self.language_changed_callbacks.append(callback)
    
    def unregister_language_changed_callback(self, callback):
        """
        Supprime un callback enregistré.
        
        Args:
            callback: Fonction à supprimer
        """
        if callback in self.language_changed_callbacks:
            self.language_changed_callbacks.remove(callback)
    
    def _detect_system_language(self) -> str:
        """Détecte automatiquement la langue du système"""
        try:
            import locale
            # Utiliser getlocale() au lieu de getdefaultlocale() (déprécié)
            system_locale = locale.getlocale()[0]
            
            if system_locale:
                # Extraire le code de langue (les 2 premiers caractères)
                lang_code = system_locale[:2].lower()
                
                # Mapper vers nos langues supportées
                if lang_code in self.available_languages if hasattr(self, 'available_languages') else TRANSLATIONS:
                    return lang_code
                elif lang_code == 'es':  # Espagnol
                    return 'es'
                elif lang_code == 'de':  # Allemand
                    return 'de'
                elif lang_code in ['en', 'us']:  # Anglais
                    return 'en'
                else:
                    return 'fr'  # Français par défaut
            
        except Exception as e:
            logging.warning(f"Impossible de détecter la langue système: {e}")
        
        return 'fr'  # Français par défaut en cas d'erreur
    
    def get_available_languages(self) -> Dict[str, str]:
        """Retourne la liste des langues disponibles"""
        return self.available_languages.copy()
    
    def get_current_language(self) -> str:
        """Retourne la langue courante"""
        return self.current_language
    
    def get_language_name(self, language_code: str) -> str:
        """Retourne le nom d'affichage d'une langue"""
        return self.available_languages.get(language_code, language_code)
    
    def save_language_preference(self):
        """Sauvegarde la préférence de langue dans les paramètres"""
        try:
            settings = QSettings('ExcelCompiler', 'ExcelCompiler')
            settings.setValue('language/current', self.current_language)
        except Exception as e:
            logging.warning(f"Impossible de sauvegarder la préférence de langue: {e}")
    
    def load_language_preference(self):
        """Charge la préférence de langue depuis les paramètres"""
        try:
            settings = QSettings('ExcelCompiler', 'ExcelCompiler')
            saved_language = settings.value('language/current', self.current_language)
            if saved_language in self.translations:
                self.set_language(saved_language)
        except Exception as e:
            logging.warning(f"Impossible de charger la préférence de langue: {e}")

class SettingsManager:
    """
    Gestionnaire des paramètres de l'application.
    Permet de sauvegarder et charger les préférences utilisateur.
    """
    _instance = None
    
    def __new__(cls):
        """Implémentation du pattern Singleton."""
        if cls._instance is None:
            cls._instance = super(SettingsManager, cls).__new__(cls)
            cls._instance._initialize()
        return cls._instance
    
    def _initialize(self):
        """Initialise le gestionnaire de paramètres."""
        self.settings = QSettings("GounouNGobi", "ExcelCompiler")
    
    def save_settings(self, window):
        """
        Sauvegarde les paramètres actuels de l'application.
        
        Args:
            window: Fenêtre principale de l'application
        """
        settings = self.settings
        
        # Sauvegarde des options générales
        settings.setValue("language", TranslationManager().current_language)
        settings.setValue("headerStartRow", window.spinbox_header_start.value())
        settings.setValue("headerRows", window.spinbox_header.value())
        settings.setValue("repeatHeaders", window.checkbox_repeat_header.isChecked())
        settings.setValue("mergeHeaders", window.checkbox_merge_headers.isChecked())
        settings.setValue("filenameOption", window.combo_filename_option.currentData())
        settings.setValue("outputFilename", window.lineedit_output_name.text())
        settings.setValue("verifyFiles", window.checkbox_verify_files.isChecked())
        settings.setValue("includePreliminary", window.checkbox_preliminary.isChecked())
        settings.setValue("preliminarySourceFile", window.preliminary_source_file)
        
        # Options avancées
        settings.setValue("removeDuplicates", window.checkbox_remove_duplicates.isChecked())
        settings.setValue("removeEmptyRows", window.checkbox_remove_empty_rows.isChecked())
        settings.setValue("sortData", window.checkbox_sort_data.isChecked())
        settings.setValue("sortColumn", window.lineedit_sort_column.text())
        settings.setValue("autoWidth", window.checkbox_auto_width.isChecked())
        settings.setValue("freezeHeader", window.checkbox_freeze_header.isChecked())
        settings.setValue("csvSupport", window.checkbox_csv.isChecked())
        
        # Option de format de date
        settings.setValue("dateFormat", window.date_format)
        
        # Sauvegarde du dernier répertoire utilisé
        if window.directory:
            settings.setValue("lastDirectory", window.directory)
        
        settings.sync()
    
    def load_settings(self, window):
        """
        Charge les paramètres sauvegardés.
        
        Args:
            window: Fenêtre principale de l'application
            
        Returns:
            bool: True si des paramètres ont été chargés, False sinon
        """
        settings = self.settings
        
        # Vérifier si des paramètres existent
        if not settings.contains("headerStartRow"):
            return False
        
        # Chargement de la langue
        language = settings.value("language", "fr")
        TranslationManager().set_language(language)
        
        # Chargement des options générales
        window.spinbox_header_start.setValue(int(settings.value("headerStartRow", 1)))
        window.spinbox_header.setValue(int(settings.value("headerRows", 1)))
        window.checkbox_repeat_header.setChecked(self._to_bool(settings.value("repeatHeaders", False)))
        window.checkbox_merge_headers.setChecked(self._to_bool(settings.value("mergeHeaders", False)))
        window.lineedit_output_name.setText(settings.value("outputFilename", "compilation.xlsx"))
        window.checkbox_verify_files.setChecked(self._to_bool(settings.value("verifyFiles", True)))
        window.checkbox_preliminary.setChecked(self._to_bool(settings.value("includePreliminary", False)))
        window.preliminary_source_file = settings.value("preliminarySourceFile", "")

        # Mettre à jour l'état des contrôles
        window.toggle_preliminary_options(Qt.CheckState.Checked.value if window.checkbox_preliminary.isChecked() else Qt.CheckState.Unchecked.value)

        # Sélectionner le bon fichier dans le combo si défini
        if window.preliminary_source_file:
            index = window.combo_preliminary_source.findText(window.preliminary_source_file)
            if index >= 0:
                window.combo_preliminary_source.setCurrentIndex(index)

        filename_option = settings.value("filenameOption", "none")
        window.filename_option = filename_option
        index = window.combo_filename_option.findData(filename_option)
        if index >= 0:
            window.combo_filename_option.setCurrentIndex(index)
        
        # Options avancées
        window.checkbox_remove_duplicates.setChecked(self._to_bool(settings.value("removeDuplicates", False)))
        window.checkbox_remove_empty_rows.setChecked(self._to_bool(settings.value("removeEmptyRows", False)))
        window.checkbox_sort_data.setChecked(self._to_bool(settings.value("sortData", False)))
        window.lineedit_sort_column.setText(settings.value("sortColumn", "A"))
        window.lineedit_sort_column.setEnabled(window.checkbox_sort_data.isChecked())
        window.checkbox_auto_width.setChecked(self._to_bool(settings.value("autoWidth", True)))
        window.checkbox_freeze_header.setChecked(self._to_bool(settings.value("freezeHeader", False)))
        window.checkbox_csv.setChecked(self._to_bool(settings.value("csvSupport", True)))
        
        # Option de format de date
        window.date_format = settings.value("dateFormat", "FRENCH")
    
        
        # Chargement du dernier répertoire utilisé
        last_directory = settings.value("lastDirectory", "")
        if last_directory and os.path.isdir(last_directory):
            window.directory = last_directory
            window.label_directory.setText(last_directory)
            window.load_files()
        
        return True
    
    def _to_bool(self, value):
        """
        Convertit une valeur en booléen.
        
        Args:
            value: Valeur à convertir
            
        Returns:
            bool: Valeur convertie
        """
        if isinstance(value, bool):
            return value
        return value.lower() in ("true", "1", "yes", "y", "t")
    
class FileVerification:
    """
    Classe utilitaire pour vérifier la compatibilité des fichiers Excel et CSV.
    """
    
    @staticmethod
    def verify_excel_file(file_path: str, header_start_row: int, header_rows: int) -> Tuple[bool, str]:
        """
        Vérifie si un fichier Excel est compatible pour la compilation.
        """
        try:
            # Support des nouveaux formats Excel
            if file_path.lower().endswith(('.xlsx', '.xlsm', '.xltx', '.xltm')):
                wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
                ws = wb.active
                
                # Vérification du nombre de lignes
                if ws.max_row < header_start_row + header_rows:
                    return False, f"Structure d'en-tête incompatible: le fichier n'a que {ws.max_row} lignes"
                    
                # Vérification de la protection du fichier
                if hasattr(ws, 'protection') and ws.protection.sheet:
                    return False, "Le fichier est protégé en écriture"
                    
                wb.close()
                return True, "Fichier compatible"
            
            elif file_path.lower().endswith('.xls'):
                # Ancien format Excel - vérification basique
                try:
                    wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
                    ws = wb.active
                    
                    if ws.max_row < header_start_row + header_rows:
                        return False, f"Structure d'en-tête incompatible: le fichier n'a que {ws.max_row} lignes"
                    
                    wb.close()
                    return True, "Fichier compatible"
                except (PermissionError, FileNotFoundError, ValueError, Exception) as e:
                    # PermissionError: Pas d'accès au fichier
                    # FileNotFoundError: Fichier supprimé pendant vérification
                    # ValueError: Format Excel corrompu
                    # Exception: Autres erreurs openpyxl
                    logging.warning(f"Impossible de vérifier le fichier Excel {file_path}: {e}")
                    # Fallback pour très anciens fichiers .xls ou fichiers protégés
                    return True, f"Fichier probablement compatible (vérification échouée: {type(e).__name__})"
            
            else:
                return False, "Format Excel non supporté"
                
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

    @staticmethod
    def verify_text_file(file_path: str, header_start_row: int, header_rows: int) -> Tuple[bool, str]:
        """
        Vérifie si un fichier texte délimité est compatible pour la compilation.
        
        Args:
            file_path: Chemin complet du fichier à vérifier
            header_start_row: Ligne de début des en-têtes
            header_rows: Nombre de lignes d'en-tête
            
        Returns:
            Tuple[bool, str]: (est_compatible, message_d'erreur)
        """
        try:
            # Détection du délimiteur
            file_ext = os.path.splitext(file_path)[1].lower()
            
            if file_ext == '.tsv':
                delimiter = '\t'
            else:
                # Détection automatique pour .txt
                encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']
                delimiter = ','
                
                for encoding in encodings:
                    try:
                        with open(file_path, 'r', encoding=encoding) as f:
                            sample = f.read(4096)
                            sniffer = csv.Sniffer()
                            delimiter = sniffer.sniff(sample).delimiter
                            break
                    except Exception:
                        continue
            
            # Compter les lignes du fichier
            for encoding in ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']:
                try:
                    with open(file_path, 'r', newline='', encoding=encoding) as file:
                        reader = csv.reader(file, delimiter=delimiter)
                        row_count = sum(1 for _ in reader)
                        break
                except Exception:
                    continue
            else:
                return False, "Erreur d'encodage: impossible de lire le fichier"
                    
            if row_count < header_start_row + header_rows:
                return False, f"Structure d'en-tête incompatible: le fichier n'a que {row_count} lignes"
                
            return True, "Fichier compatible"
            
        except PermissionError:
            return False, "Le fichier est ouvert dans une autre application"
        except Exception as e:
            return False, f"Erreur lors de la vérification: {str(e)}"


class CompilationWorker(QThread):
    """
    Thread de travail qui gère la compilation des fichiers Excel et CSV.
    Permet de ne pas bloquer l'interface utilisateur pendant le traitement.
    Thread-safe avec signaux robustes pour communication avec l'interface principale.
    """
    
    # Signaux thread-safe pour communication avec l'interface principale
    progress = pyqtSignal(int)  # Signal de progression
    progress_detail = pyqtSignal(int, str)  # Signal de progression avec détail
    error = pyqtSignal(str)     # Signal d'erreur
    warning = pyqtSignal(str)   # Signal d'avertissement
    info = pyqtSignal(str)      # Signal d'information
    finished = pyqtSignal(tuple)  # Signal de fin avec les données compilées
    status_update = pyqtSignal(str)  # Signal de mise à jour du statut
    
    def __init__(self, files, directory, header_start_row, header_rows, filename_option="none",
             sort_data=False, sort_column=0, repeat_headers=False, remove_empty_rows=False,
             remove_duplicates=False, date_format="FRENCH", include_preliminary=False,
             preliminary_source_file="", parent=None):
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
            date_format: Format de date à utiliser
            include_preliminary: Inclure les informations préliminaires
            preliminary_source_file: Fichier source pour les informations préliminaires
            parent: Widget parent
        """
        # Validations
        if not files:
            raise ValueError("La liste de fichiers ne peut pas être vide")
        if not os.path.isdir(directory):
            raise ValueError(f"Le répertoire {directory} n'existe pas")
        if header_start_row < 1:
            raise ValueError("La ligne de début doit être >= 1")
        if filename_option not in ["none", "with_extension", "without_extension"]:
            raise ValueError(f"Option filename invalide: {filename_option}")

        # Assignation des paramètres
        super().__init__(parent)
        self.files = files
        self.directory = directory
        self.header_start_row = header_start_row
        self.header_rows = header_rows
        self.filename_option = filename_option
        self.sort_data = sort_data
        self.sort_column = sort_column
        self.repeat_headers = repeat_headers
        self.remove_empty_rows = remove_empty_rows
        self.remove_duplicates = remove_duplicates
        self.date_format = date_format
        self.include_preliminary = include_preliminary 
        self.preliminary_source_file = preliminary_source_file
        
        # Nouveaux composants de performance
        self.memory_manager = MemoryManager()
        
        # Détection de similarité pour optimiser les fichiers répétitifs
        self.similarity_detector = SimilarityDetector()
        self.file_processing_cache = {}  # Cache des résultats de traitement
        
        # Attributs de robustesse manquants
        self.cancellation_token = CancellationToken()
        self.error_recovery = ErrorRecoveryManager()
        self.input_validator = InputValidator()  
        self.file_detector = AdvancedFileDetector()
        self.metadata_cache = FileMetadataCache()
        self.chunked_reader = ChunkedFileReader()
        self.executor = ThreadPoolExecutor(max_workers=MAX_THREADS)
        
        # Optimisations mémoire
        self.last_memory_check = 0
        self.cached_available_memory = 0
        self.memory_check_interval = 5.0
        self._memory_check_count = 0
        
        # Attributs manquants détectés
        self._access_times = {}
        self._backup_config = {}
        self._cache = {}
        self._cancelled = False
        self._current_operation_id = None
        self._fingerprints = {}
        self._last_estimates = {}
        self._last_warning = None
        self._lock = threading.Lock()
        self._options_group_original_visible = True
    
    def force_garbage_collection(self):
        """Force le garbage collection"""
        import gc
        gc.collect()
    
    def is_memory_limit_reached(self, new_data=None, combined_data=None):
        """Vérifie si la limite mémoire serait atteinte"""
        try:
            # Estimation simplifiée de la mémoire
            memory_info = psutil.virtual_memory()
            available_memory = memory_info.available
            
            # Si appelé sans arguments, vérifier juste l'usage mémoire actuel
            if new_data is None and combined_data is None:
                # Vérifier si plus de 85% de la RAM est utilisée
                return memory_info.percent > 85.0
            
            # Estimer la taille des nouvelles données
            if isinstance(new_data, list) and new_data:
                estimated_new_size = len(str(new_data)) * 8  # Estimation approximative
            else:
                estimated_new_size = 1024  # Valeur par défaut
            
            # Estimer la taille des données actuelles
            if isinstance(combined_data, list) and combined_data:
                estimated_current_size = len(str(combined_data)) * 8
            else:
                estimated_current_size = 0
            
            # Vérifier si on a assez de mémoire (seuil de sécurité 80%)
            total_needed = estimated_new_size + estimated_current_size
            return total_needed > (available_memory * 0.8)
            
        except Exception as e:
            logging.warning(f"Impossible de vérifier la mémoire: {e}")
            return False  # En cas d'erreur, continuer le traitement
    
    def _check_file_similarity_cache(self, file_path: str, file_index: int) -> Optional[Tuple]:
        """Vérifie si un fichier similaire a déjà été traité"""
        try:
            # Générer l'empreinte du fichier
            fingerprint = self.similarity_detector.get_file_fingerprint(file_path)
            if not fingerprint:
                return None
            
            # Vérifier si nous avons déjà traité un fichier avec cette empreinte
            if fingerprint in self.file_processing_cache:
                cached_result = self.file_processing_cache[fingerprint]
                logging.info(f"Fichier similaire trouvé en cache: {file_path}")
                
                # Cloner les données pour éviter les modifications partagées
                headers, data, merged_cells = cached_result
                return (
                    [row[:] for row in headers] if headers else None,  # Deep copy headers
                    [row[:] for row in data] if data else [],         # Deep copy data
                    merged_cells[:]                                   # Copy merged cells
                )
            
            return None
        except Exception as e:
            logging.warning(f"Erreur lors de la vérification de similarité: {e}")
            return None
    
    def _cache_file_result(self, file_path: str, headers, data, merged_cells):
        """Met en cache le résultat du traitement d'un fichier"""
        try:
            fingerprint = self.similarity_detector.get_file_fingerprint(file_path)
            if fingerprint:
                # Limiter la taille du cache (garder seulement les 20 derniers)
                if len(self.file_processing_cache) >= 20:
                    # Supprimer le plus ancien (simplistic, pourrait être amélioré avec LRU)
                    first_key = next(iter(self.file_processing_cache))
                    del self.file_processing_cache[first_key]
                
                # Cloner les données avant mise en cache
                self.file_processing_cache[fingerprint] = (
                    [row[:] for row in headers] if headers else None,
                    [row[:] for row in data] if data else [],
                    merged_cells[:]
                )
                logging.debug(f"Résultat mis en cache pour: {file_path}")
        except Exception as e:
            logging.warning(f"Erreur lors de la mise en cache: {e}")

    def _combine_headers_for_data(self, headers):
        """
        Combine les en-têtes multi-lignes en un seul en-tête pour la compilation.
        
        Args:
            headers: Liste des lignes d'en-têtes
            
        Returns:
            Liste des en-têtes combinées
        """
        if not headers or len(headers) == 1:
            return headers[0] if headers else []
        
        # Déterminer le nombre de colonnes maximum
        max_cols = max(len(row) for row in headers) if headers else 0
        combined = []
        
        for col_idx in range(max_cols):
            combined_cell = []
            for row in headers:
                if col_idx < len(row) and row[col_idx] not in [None, '']:
                    cell_value = str(row[col_idx]).strip()
                    if cell_value and cell_value not in combined_cell:
                        combined_cell.append(cell_value)
            
            # Combiner les parties non vides avec un séparateur
            combined.append(' | '.join(combined_cell) if combined_cell else f'Col_{col_idx+1}')
        
        return combined
    
    def _check_memory_before_extend(self, new_data, combined_data):
        """
        Vérifie la mémoire disponible avant d'ajouter des données (optimisé avec cache).
        
        Args:
            new_data: Nouvelles données à ajouter
            combined_data: Données déjà en mémoire
            
        Raises:
            MemoryError: Si la mémoire est insuffisante
        """
        try:
            current_time = time.time()
            
            # Utiliser le cache si la vérification est récente
            if (current_time - self.last_memory_check) < self.memory_check_interval:
                available_memory = self.cached_available_memory
            else:
                # Nouvelle vérification mémoire
                memory_info = psutil.virtual_memory()
                available_memory = memory_info.available
                
                # Mettre à jour le cache
                self.last_memory_check = current_time
                self.cached_available_memory = available_memory
            
            # Incrémenter compteur pour statistiques
            self._memory_check_count += 1
            
            # Estimer la taille des nouvelles données (optimisé)
            if isinstance(new_data, list):
                # Estimation plus précise basée sur échantillonnage
                if len(new_data) > 100:
                    # Échantillonner les 100 premières lignes pour estimation
                    sample_size = sum(len(str(row)) for row in new_data[:100])
                    estimated_new_size = (sample_size * len(new_data)) // 100
                else:
                    estimated_new_size = sum(len(str(row)) for row in new_data)
                estimated_new_size *= 2  # Facteur de sécurité pour overhead Python
            else:
                estimated_new_size = len(str(new_data)) * 8
            
            # Estimer la taille des données actuelles
            estimated_current_size = len(combined_data) * 512 if combined_data else 0
            
            # Taille totale estimée après ajout
            estimated_total_size = estimated_current_size + estimated_new_size
            
            # Seuil adaptatif basé sur la RAM totale
            total_ram = psutil.virtual_memory().total
            if total_ram > 16 * (1024**3):  # Plus de 16GB de RAM
                safe_factor = 0.9  # Plus permissif
            elif total_ram > 8 * (1024**3):  # Plus de 8GB de RAM
                safe_factor = 0.8  # Équilibré
            else:
                safe_factor = 0.7  # Plus conservateur
            
            safe_limit = available_memory * safe_factor
            
            # Logger pour debugging si nécessaire (moins fréquent)
            if estimated_total_size > safe_limit * 0.6 and (current_time - getattr(self, '_last_warning', 0)) > 10:
                logging.info(
                    f"Utilisation mémoire élevée: {estimated_total_size/(1024*1024):.1f}MB "
                    f"sur {available_memory/(1024*1024):.1f}MB disponible (seuil: {safe_factor*100:.0f}%)"
                )
                self._last_warning = current_time
            
            if estimated_total_size > safe_limit:
                # Convertir en MB pour affichage
                needed_mb = estimated_total_size / (1024 * 1024)
                available_mb = available_memory / (1024 * 1024)
                current_mb = estimated_current_size / (1024 * 1024)
                total_ram_gb = total_ram / (1024**3)
                
                raise MemoryError(
                    f"Mémoire insuffisante pour traiter ce volume de données.\n"
                    f"Mémoire nécessaire: {needed_mb:.1f}MB\n"
                    f"Mémoire disponible: {available_mb:.1f}MB (RAM totale: {total_ram_gb:.1f}GB)\n"
                    f"Données déjà chargées: {current_mb:.1f}MB\n"
                    f"Seuil sécurité: {safe_factor*100:.0f}% ({safe_limit/(1024*1024):.1f}MB)\n"
                    f"Recommandation: Réduisez le nombre de fichiers sélectionnés."
                )
                
        except Exception as e:
            # En cas d'erreur de vérification, logger mais ne pas bloquer
            if not isinstance(e, MemoryError):
                logging.warning(f"Impossible de vérifier la mémoire: {e}")
        
        # Initialisation des composants de performance et robustesse
        self.metadata_cache = FileMetadataCache()
        self.chunked_reader = ChunkedFileReader()
        self.executor = ThreadPoolExecutor(max_workers=MAX_THREADS)
        
        # Cache pour optimisation vérification mémoire
        self.last_memory_check = 0
        self.cached_available_memory = 0
        self.memory_check_interval = 5.0  # Vérifier toutes les 5 secondes max
        self._memory_check_count = 0  # Compteur pour statistiques
        
        # Composants de robustesse
        self.cancellation_token = CancellationToken()
        # Utiliser le logger standard
        self.error_recovery = ErrorRecoveryManager()
        self.input_validator = InputValidator()  
        self.file_detector = AdvancedFileDetector()
        
        # Cache pour métadonnées
        self.metadata_cache = FileMetadataCache()
        
        # Gestionnaire de mémoire
        # Note: self.memory_manager pointe vers self car CompilationWorker 
        # implémente les méthodes de gestion mémoire
        self.memory_manager = self
    
    def force_garbage_collection(self):
        """Force le garbage collection"""
        import gc
        gc.collect() 
    
    def cancel(self):
        """Annule la compilation en cours"""
        self.cancellation_token.cancel()
        logging.info("Annulation demandée par l'utilisateur")
        
    def run(self):
        """
        Méthode principale exécutée dans le thread avec optimisations de performance.
        Utilise le pool de threads et le streaming par chunks.
        """
        start_time = time.time()
        logging.info(f"Début de compilation de {len(self.files)} fichiers")
        
        # Signal d'étape : Démarrage
        self.status_update.emit("🚀 Initialisation de la compilation...")
        
        try:
            # Signal d'étape : Validation
            self.status_update.emit("🔍 Validation des paramètres...")
            
            # Validation des entrées
            validation_params = {
                'header_start_row': self.header_start_row,
                'header_rows': self.header_rows,
                'filename_option': self.filename_option,
                'sort_column': self.sort_column,
                'output_file': 'compilation.xlsx'  # Valeur par défaut
            }
            
            is_valid, validation_errors = self.input_validator.validate_parameters(validation_params)
            if not is_valid:
                error_msg = "Erreurs de validation: " + "; ".join(validation_errors)
                logging.error(error_msg)
                self.error.emit(error_msg)
                return
            
            is_valid, validation_errors = self.input_validator.validate_files(self.files, self.directory)
            if not is_valid:
                error_msg = "Erreurs de fichiers: " + "; ".join(validation_errors)
                logging.error(error_msg)
                self.error.emit(error_msg)
                return
            
            combined_data = []
            headers = None
            merged_cells = []
            preliminary_info = []
            
            successful_files = []
            failed_files = []
            
            # Vérifier l'annulation
            self.cancellation_token.throw_if_cancelled()
            
            # Charger les informations préliminaires si demandées
            if self.include_preliminary and self.preliminary_source_file:
                self.status_update.emit("📋 Chargement des informations préliminaires...")
                try:
                    preliminary_info = self._load_preliminary_info()
                    logging.info("Informations préliminaires chargées")
                except FileNotFoundError:
                    logging.warning(f"Fichier préliminaire non trouvé: {self.preliminary_source_file}")
                except PermissionError:
                    logging.warning(f"Impossible d'accéder au fichier préliminaire: {self.preliminary_source_file}")
                except (ValueError, KeyError) as e:
                    logging.warning(f"Format invalide du fichier préliminaire {self.preliminary_source_file}: {str(e)}")
                except MemoryError:
                    logging.warning(f"Fichier préliminaire trop volumineux: {self.preliminary_source_file}")
                except Exception as e:
                    logging.warning(f"Erreur inattendue lors du chargement préliminaire {self.preliminary_source_file}: {type(e).__name__}: {str(e)}")
                    if not self.error_recovery.handle_error(e, {'context': 'preliminary_info'}):
                        return
            
            # Signal d'étape : Début du traitement des fichiers
            self.status_update.emit(f"📂 Traitement de {len(self.files)} fichiers...")
            
            # Traitement parallèle des fichiers avec pool de threads
            futures = []
            
            # Déterminer le mode de traitement selon les options
            order_dependent_options = (
                self.repeat_headers or          # En-têtes répétés nécessitent global_headers
                self.include_preliminary or     # Métadonnées du premier fichier
                len(merged_cells) > 0          # Cellules fusionnées dépendent de l'ordre
            )
            
            use_sequential = (
                len(self.files) <= 3 or         # Petits nombres de fichiers
                order_dependent_options         # Options incompatibles avec parallélisme
            )
            
            if use_sequential:
                logging.info(f"Mode SÉQUENTIEL choisi: {len(self.files)} fichiers, options_dépendantes={order_dependent_options}")
            else:
                logging.info(f"Mode PARALLÈLE choisi: {len(self.files)} fichiers, aucune option dépendante")
            
            if use_sequential:
                for i, file in enumerate(self.files):
                    try:
                        # Vérifier l'annulation avant chaque fichier
                        self.cancellation_token.throw_if_cancelled()
                        
                        logging.info(f"Traitement du fichier {i+1}/{len(self.files)}: {file}")
                        
                        result = self._process_single_file(file, i, headers, preliminary_info)
                        if result:
                            headers = self._handle_file_result(result, combined_data, headers, merged_cells, 
                                                   successful_files, i == 0) or headers
                            successful_files.append(file)
                            logging.info(f"Fichier {file} traité avec succès")
                        
                        self.progress.emit(i + 1)
                        
                    except InterruptedError:
                        logging.info("Compilation annulée par l'utilisateur")
                        return
                    except Exception as e:
                        context = {'file_path': os.path.join(self.directory, file)}
                        if self.error_recovery.handle_error(e, context):
                            failed_files.append((file, str(e)))
                            self.error.emit(f"Erreur avec le fichier {file}: {str(e)}")
                            continue  # Continuer avec le fichier suivant
                        else:
                            raise  # Arrêter si récupération impossible
            else:
                # Traitement parallèle pour de nombreux fichiers
                with self.executor:
                    # Soumettre les tâches de traitement de fichiers
                    for i, file in enumerate(self.files):
                        self.cancellation_token.throw_if_cancelled()
                        future = self.executor.submit(self._process_single_file_robust, file, i, headers, preliminary_info)
                        futures.append((future, file, i))
                    
                    # Traiter les résultats au fur et à mesure
                    for future, file, file_index in futures:
                        try:
                            self.cancellation_token.throw_if_cancelled()
                            
                            result = future.result(timeout=300)  # 5 minutes max par fichier
                            if result:
                                headers = self._handle_file_result(result, combined_data, headers, merged_cells, 
                                                       successful_files, file_index == 0) or headers
                            successful_files.append(file)
                            
                            # Signal informatif sur le fichier traité
                            progress_detail = f"📄 Fichier {file_index + 1}/{len(self.files)}: {file} ✅"
                            self.progress_detail.emit(file_index + 1, progress_detail)
                            self.progress.emit(file_index + 1)
                            
                            # Vérification de la mémoire et nettoyage si nécessaire
                            if self.memory_manager.is_memory_limit_reached():
                                logging.warning("Limite mémoire atteinte - nettoyage en cours")
                                self.memory_manager.force_garbage_collection()
                        
                        except InterruptedError:
                            logging.info("Compilation annulée par l'utilisateur")
                            return
                        except Exception as e:
                            context = {'file_path': os.path.join(self.directory, file)}
                            if self.error_recovery.handle_error(e, context):
                                failed_files.append((file, str(e)))
                                self.error.emit(f"Erreur avec le fichier {file}: {str(e)}")
                                continue
                            else:
                                raise
            
            # Vérifier l'annulation avant le post-traitement
            self.cancellation_token.throw_if_cancelled()
            
            # Traitement post-compilation optimisé
            if combined_data and headers:
                self.status_update.emit("⚙️ Post-traitement des données...")
                logging.info("Début du post-traitement des données")
                
                # Suppression des doublons par chunks pour économiser la mémoire
                if self.remove_duplicates:
                    self.status_update.emit("🔄 Suppression des doublons...")
                    logging.info("Suppression des doublons")
                    self.cancellation_token.throw_if_cancelled()
                    old_data = combined_data
                    combined_data = self._remove_duplicate_rows_chunked_preserve_headers(combined_data)
                    # Libération explicite des anciennes données après déduplication
                    del old_data
                    self.memory_manager.force_garbage_collection()
                    
                # Tri des données par chunks si l'option est activée
                if self.sort_data and headers and self.sort_column < len(headers[-1]):
                    self.status_update.emit("📊 Tri des données...")
                    logging.info("Tri des données en cours...")
                    self.cancellation_token.throw_if_cancelled()
                    combined_data = self._sort_data_chunked(combined_data, headers)
                
                # Ajout du nom du fichier source à l'en-tête si nécessaire    
                if self.filename_option != "none" and headers:
                    # Ajouter la colonne à toutes les lignes d'en-têtes pour maintenir la cohérence
                    for header_row in headers:
                        if header_row == headers[-1]:
                            header_row.append("Fichier source")
                        else:
                            header_row.append(None)  # Colonne vide pour les autres lignes d'en-têtes
                
                # Option : combiner les en-têtes multi-lignes en un seul en-tête
                # Pour une meilleure compatibilité avec les outils d'analyse de données
                if len(headers) > 1:
                    combined_header = self._combine_headers_for_data(headers)
                    # Optionnel : remplacer les en-têtes multi-lignes par l'en-tête combiné
                    # headers = [combined_header]
            
            # Nettoyage final avant émission des données
            self.status_update.emit("🧹 Finalisation et nettoyage...")
            logging.info("Nettoyage mémoire final avant génération du fichier")
            self.memory_manager.force_garbage_collection()
            
            # Statistiques finales
            compilation_time = time.time() - start_time
            error_summary = self.error_recovery.get_summary()
            
            logging.info(f"Compilation terminée en {compilation_time:.2f} secondes")
            logging.info(f"Fichiers traités avec succès: {len(successful_files)}")
            logging.info(f"Fichiers échoués: {len(failed_files)}")
            logging.info(f"Total d'erreurs: {error_summary['total_errors']}")
            
            # Émission du signal de fin avec toutes les données
            self.finished.emit((preliminary_info, headers, combined_data, merged_cells, successful_files, failed_files))
            
        except InterruptedError:
            logging.info("Compilation annulée par l'utilisateur")
            self.error.emit("Compilation annulée par l'utilisateur")
        except MemoryError as e:
            logging.critical(f"Mémoire insuffisante lors de la compilation: {str(e)}")
            self.error.emit(f"Mémoire insuffisante: Réduisez le nombre de fichiers sélectionnés")
        except PermissionError as e:
            logging.critical(f"Problème d'autorisation lors de la compilation: {str(e)}")
            self.error.emit(f"Erreur d'autorisation: Vérifiez que les fichiers ne sont pas ouverts")
        except FileNotFoundError as e:
            logging.critical(f"Fichier manquant lors de la compilation: {str(e)}")
            self.error.emit(f"Fichier manquant: Un fichier sélectionné n'existe plus")
        except Exception as e:
            logging.critical(f"Erreur critique inattendue: {type(e).__name__}: {str(e)}")
            self.error.emit(f"Erreur critique inattendue: {type(e).__name__}: {str(e)}")
        finally:
            # Nettoyage final
            self.metadata_cache.clear()
            self.memory_manager.force_garbage_collection()
            logging.info("Nettoyage final terminé")
    
    def _process_single_file_robust(self, file: str, file_index: int, global_headers: Optional[List], preliminary_info: List):
        """
        Version robuste du traitement d'un fichier avec gestion d'erreurs intégrée et cache de similarité.
        """
        try:
            self.cancellation_token.throw_if_cancelled()
            
            # Vérifier d'abord le cache de similarité
            file_path = os.path.join(self.directory, file)
            cached_result = self._check_file_similarity_cache(file_path, file_index)
            
            if cached_result is not None:
                # Fichier similaire trouvé en cache, utiliser le résultat mis en cache
                headers, data, merged_cells = cached_result
                result = (headers, data, merged_cells)
                
                # Adapter les données pour le nom de fichier si nécessaire
                if self.filename_option != "none" and data:
                    for row in data:
                        if self.filename_option == "with_extension":
                            row.append(file)
                        elif self.filename_option == "without_extension":
                            filename_without_ext = os.path.splitext(file)[0]
                            row.append(filename_without_ext)
                
                return result
            
            # Traitement normal si pas de cache
            result = self._process_single_file(file, file_index, global_headers, preliminary_info)
            
            # Mettre en cache le résultat pour les fichiers futurs similaires
            if result:
                headers, data, merged_cells = result
                self._cache_file_result(file_path, headers, data, merged_cells)
            
            return result
        except InterruptedError:
            raise  # Propager l'interruption
        except Exception as e:
            context = {'file_path': os.path.join(self.directory, file)}
            if self.error_recovery.handle_error(e, context):
                return None  # Continuer avec les autres fichiers
            else:
                raise  # Arrêter si récupération impossible
    
    def _process_single_file(self, file: str, file_index: int, global_headers: Optional[List], preliminary_info: List):
        """
        Traite un seul fichier avec optimisations.
        """
        file_path = os.path.join(self.directory, file)
        capture_preliminary = (file_index == 0 and self.include_preliminary and not self.preliminary_source_file)
        
        # Vérifier le cache des métadonnées
        cached_metadata = self.metadata_cache.get_metadata(file_path)
        if cached_metadata and not capture_preliminary:
            # Utiliser les données en cache si disponibles
            return cached_metadata
        
        # Traitement différent selon le type de fichier avec streaming
        if file.lower().endswith(('.xlsx', '.xlsm', '.xltx', '.xltm', '.xls')):
            result = self._process_excel_file_chunked(file_path, file_index, global_headers, 
                                                    preliminary_info, capture_preliminary)
        elif file.lower().endswith('.csv'):
            result = self._process_csv_file_chunked(file_path, file_index, global_headers, 
                                                  preliminary_info, capture_preliminary)
        elif file.lower().endswith(('.tsv', '.txt')):
            result = self._process_text_file_chunked(file_path, file_index, global_headers, 
                                                   preliminary_info, capture_preliminary)
        else:
            raise ValueError(f"Format de fichier non pris en charge: {file}")
        
        # Mettre en cache les métadonnées si approprié
        if result and not capture_preliminary:
            self.metadata_cache.set_metadata(file_path, result)
        
        return result
    
    def _handle_file_result(self, result, combined_data: List, headers: Optional[List], 
                          merged_cells: List, successful_files: List, is_first_file: bool) -> Optional[List]:
        """
        Traite le résultat d'un fichier et met à jour les structures de données.
        
        Returns:
            Headers mises à jour (ou None si pas de modification)
        """
        file_headers, file_data, file_merged_cells = result
        
        # Mise à jour des en-têtes globales si nécessaire
        updated_headers = headers
        if headers is None and file_headers:
            updated_headers = file_headers
            merged_cells.extend(file_merged_cells)
            
        # Vérification mémoire avant ajout des données
        self._check_memory_before_extend(file_data, combined_data)
        
        # Ajout des données au résultat combiné    
        combined_data.extend(file_data)
        
        # Nettoyage de la mémoire après chaque fichier
        del file_data
        if self.memory_manager.is_memory_limit_reached():
            self.memory_manager.force_garbage_collection()
            
        return updated_headers
    
    def _process_excel_file_chunked(self, file_path: str, file_index: int, global_headers: Optional[List], 
                                   preliminary_info: List, capture_preliminary: bool = True):
        """
        Traite un fichier Excel par chunks pour optimiser la mémoire.
        """
        try:
            file_headers = []
            file_merged_cells = []
            file_data = []
            
            # Pour le premier fichier, charger normalement pour obtenir les merged_cells
            if file_index == 0:
                wb = openpyxl.load_workbook(file_path, data_only=True)
                ws = wb.active
                
                # Capture des informations préliminaires
                if self.header_start_row > 1 and capture_preliminary:
                    for row in range(1, self.header_start_row):
                        row_data = [cell.value for cell in ws[row]]
                        preliminary_info.append(row_data)
                
                # Capture des en-têtes et merged_cells
                if global_headers is None:
                    for row in range(self.header_start_row, self.header_start_row + self.header_rows):
                        header_row = [cell.value for cell in ws[row]]
                        file_headers.append(header_row)
                    
                    if hasattr(ws, 'merged_cells'):
                        for merged_range in ws.merged_cells.ranges:
                            # Ne capturer que les cellules fusionnées dans la zone des en-têtes
                            if (merged_range.min_row >= self.header_start_row and 
                                merged_range.max_row <= (self.header_start_row + self.header_rows - 1)):
                                file_merged_cells.append(merged_range)
                else:
                    file_headers = global_headers
                
                wb.close()
                del wb
            
            # Déterminer quels en-têtes utiliser pour la répétition
            headers_to_repeat = file_headers if file_headers else global_headers
            
            
            # Si répétition des en-têtes est activée et ce n'est pas le premier fichier
            if self.repeat_headers and file_index > 0 and headers_to_repeat:
                # Ajouter les en-têtes
                for header_row in headers_to_repeat:
                    row_data = header_row.copy()
                    if self.filename_option != "none":
                        row_data.append(None)
                    file_data.append(row_data)
            
            # Utiliser le lecteur par chunks pour les données
            filename = os.path.basename(file_path)
            start_row = self.header_start_row + self.header_rows
            
            for chunk in self.chunked_reader.read_excel_chunks(file_path, start_row):
                # Vérifier l'annulation à chaque chunk
                self.cancellation_token.throw_if_cancelled()
                
                processed_chunk = []
                
                for row_data in chunk:
                    # Ajouter le nom du fichier selon l'option choisie
                    if self.filename_option == "with_extension":
                        row_data.append(filename)
                    elif self.filename_option == "without_extension":
                        filename_without_ext = os.path.splitext(filename)[0]
                        row_data.append(filename_without_ext)
                    
                    # Vérifier si la ligne n'est pas vide
                    if not self.remove_empty_rows or not all(
                        cell is None or str(cell).strip() == "" 
                        for cell in row_data[:-1 if self.filename_option != "none" else None]
                    ):
                        processed_chunk.append(row_data)
                
                # Vérification mémoire avant ajout de chunk
                self._check_memory_before_extend(processed_chunk, file_data)
                file_data.extend(processed_chunk)
                
                # Nettoyage de la mémoire si nécessaire
                if self.memory_manager.is_memory_limit_reached():
                    self.memory_manager.force_garbage_collection()
            
            return file_headers, file_data, file_merged_cells
            
        except Exception as e:
            logging.error(f"Erreur lors du traitement par chunks du fichier Excel {file_path}: {e}")
            raise
    
    def _process_csv_file_chunked(self, file_path: str, file_index: int, global_headers: Optional[List], 
                                 preliminary_info: List, capture_preliminary: bool = True):
        """
        Traite un fichier CSV par chunks pour optimiser la mémoire.
        """
        try:
            # Détection de l'encodage et du délimiteur
            encoding, delimiter = self._detect_csv_format(file_path)
            
            file_headers = []
            file_merged_cells = []  # Toujours vide pour CSV
            file_data = []
            
            # Traitement des en-têtes et informations préliminaires pour le premier fichier
            if file_index == 0 or global_headers is None:
                df_sample = pd.read_csv(file_path, delimiter=delimiter, encoding=encoding, 
                                      nrows=max(self.header_start_row + self.header_rows, 10), header=None)
                
                # Capture des informations préliminaires
                if self.header_start_row > 1 and capture_preliminary:
                    for row in range(0, self.header_start_row - 1):
                        if row < len(df_sample):
                            preliminary_info.append(df_sample.iloc[row].tolist())
                
                # Capture des en-têtes
                if global_headers is None:
                    for row in range(self.header_start_row - 1, self.header_start_row - 1 + self.header_rows):
                        if row < len(df_sample):
                            file_headers.append(df_sample.iloc[row].tolist())
                else:
                    file_headers = global_headers
                
                del df_sample
            
            # Déterminer quels en-têtes utiliser pour la répétition
            headers_to_repeat = file_headers if file_headers else global_headers
            
            
            # Si répétition des en-têtes est activée et ce n'est pas le premier fichier
            if self.repeat_headers and file_index > 0 and headers_to_repeat:
                # Ajouter les en-têtes
                for header_row in headers_to_repeat:
                    row_data = header_row.copy()
                    if self.filename_option != "none":
                        row_data.append(None)
                    file_data.append(row_data)
            
            # Traitement par chunks des données
            filename = os.path.basename(file_path)
            skip_rows = self.header_start_row - 1 + self.header_rows
            
            for chunk in self.chunked_reader.read_csv_chunks(file_path, delimiter, encoding, skip_rows):
                # Vérifier l'annulation à chaque chunk
                self.cancellation_token.throw_if_cancelled()
                
                processed_chunk = []
                
                for row_data in chunk:
                    # Ajouter le nom du fichier si demandé
                    if self.filename_option != "none":
                        if self.filename_option == "with_extension":
                            row_data.append(filename)
                        elif self.filename_option == "without_extension":
                            filename_without_ext = os.path.splitext(filename)[0]
                            row_data.append(filename_without_ext)
                    
                    # Vérifier si la ligne n'est pas vide
                    if not self.remove_empty_rows or not all(
                        cell is None or str(cell).strip() == "" 
                        for cell in row_data[:-1 if self.filename_option != "none" else None]
                    ):
                        processed_chunk.append(row_data)
                
                # Vérification mémoire avant ajout de chunk
                self._check_memory_before_extend(processed_chunk, file_data)
                file_data.extend(processed_chunk)
                
                if self.memory_manager.is_memory_limit_reached():
                    self.memory_manager.force_garbage_collection()
            
            return file_headers, file_data, file_merged_cells
            
        except Exception as e:
            logging.error(f"Erreur lors du traitement par chunks du fichier CSV {file_path}: {e}")
            raise
    
    def _process_text_file_chunked(self, file_path: str, file_index: int, global_headers: Optional[List], 
                                  preliminary_info: List, capture_preliminary: bool = True):
        """
        Traite un fichier texte (TSV/TXT) par chunks pour optimiser la mémoire.
        Utilise la détection automatique améliorée des délimiteurs.
        """
        try:
            # Détection automatique pour fichiers .txt
            if file_path.lower().endswith('.txt'):
                encoding = self.file_detector.detect_encoding(file_path)
                delimiter = self.file_detector.detect_delimiter(file_path, encoding)
                logging.info(f"Fichier TXT - Délimiteur détecté: '{delimiter}'")
            
            # Pour TSV, forcer la tabulation mais vérifier l'encodage
            elif file_path.lower().endswith('.tsv'):
                encoding = self.file_detector.detect_encoding(file_path)
                delimiter = '\t'
                logging.info(f"Fichier TSV - Délimiteur: tabulation, Encodage: {encoding}")
                
                # Vérifier que le fichier contient bien des tabulations
                try:
                    with open(file_path, 'r', encoding=encoding) as f:
                        sample = f.read(1024)
                        if '\t' not in sample and ',' in sample:
                            logging.warning("Fichier .tsv ne contient pas de tabulations, basculement vers virgule")
                            delimiter = ','
                except Exception:
                    pass
            
            else:
                # Fallback sur CSV standard
                encoding, delimiter = self._detect_csv_format(file_path)
            
            # Traiter comme un fichier CSV avec l'encodage et délimiteur détectés
            return self._process_csv_file_chunked_with_format(file_path, file_index, global_headers, 
                                                            preliminary_info, capture_preliminary, 
                                                            encoding, delimiter)
            
        except Exception as e:
            logging.warning(f"Erreur lors de la détection format texte: {e}, fallback CSV")
            return self._process_csv_file_chunked(file_path, file_index, global_headers, 
                                                preliminary_info, capture_preliminary)
    
    def _process_csv_file_chunked_with_format(self, file_path: str, file_index: int, global_headers: Optional[List], 
                                            preliminary_info: List, capture_preliminary: bool, 
                                            encoding: str, delimiter: str):
        """
        Version de _process_csv_file_chunked qui utilise l'encodage et délimiteur fournis.
        """
        try:
            file_headers = []
            file_merged_cells = []  # Toujours vide pour CSV
            file_data = []
            
            # Traitement des en-têtes et informations préliminaires pour le premier fichier
            if file_index == 0 or global_headers is None:
                df_sample = pd.read_csv(file_path, delimiter=delimiter, encoding=encoding, 
                                      nrows=max(self.header_start_row + self.header_rows, 10), header=None)
                
                # Capture des informations préliminaires
                if self.header_start_row > 1 and capture_preliminary:
                    for row in range(0, self.header_start_row - 1):
                        if row < len(df_sample):
                            preliminary_info.append(df_sample.iloc[row].tolist())
                
                # Capture des en-têtes
                if global_headers is None:
                    for row in range(self.header_start_row - 1, self.header_start_row - 1 + self.header_rows):
                        if row < len(df_sample):
                            file_headers.append(df_sample.iloc[row].tolist())
                else:
                    file_headers = global_headers
                
                del df_sample
            
            # Déterminer quels en-têtes utiliser pour la répétition
            headers_to_repeat = file_headers if file_headers else global_headers
            
            
            # Si répétition des en-têtes est activée et ce n'est pas le premier fichier
            if self.repeat_headers and file_index > 0 and headers_to_repeat:
                # Ajouter les en-têtes
                for header_row in headers_to_repeat:
                    row_data = header_row.copy()
                    if self.filename_option != "none":
                        row_data.append(None)
                    file_data.append(row_data)
            
            # Traitement par chunks des données
            filename = os.path.basename(file_path)
            skip_rows = self.header_start_row - 1 + self.header_rows
            
            for chunk in self.chunked_reader.read_csv_chunks(file_path, delimiter, encoding, skip_rows):
                # Vérifier l'annulation à chaque chunk
                self.cancellation_token.throw_if_cancelled()
                
                processed_chunk = []
                
                for row_data in chunk:
                    # Ajouter le nom du fichier si demandé
                    if self.filename_option != "none":
                        if self.filename_option == "with_extension":
                            row_data.append(filename)
                        elif self.filename_option == "without_extension":
                            filename_without_ext = os.path.splitext(filename)[0]
                            row_data.append(filename_without_ext)
                    
                    # Vérifier si la ligne n'est pas vide
                    if not self.remove_empty_rows or not all(
                        cell is None or str(cell).strip() == "" 
                        for cell in row_data[:-1 if self.filename_option != "none" else None]
                    ):
                        processed_chunk.append(row_data)
                
                # Vérification mémoire avant ajout de chunk
                self._check_memory_before_extend(processed_chunk, file_data)
                file_data.extend(processed_chunk)
                
                if self.memory_manager.is_memory_limit_reached():
                    self.memory_manager.force_garbage_collection()
            
            return file_headers, file_data, file_merged_cells
            
        except Exception as e:
            logging.error(f"Erreur lors du traitement avec format spécifique: {e}")
            raise
    
    def _detect_csv_format(self, file_path: str) -> Tuple[str, str]:
        """
        Détecte l'encodage et le délimiteur d'un fichier CSV avec détection avancée.
        """
        try:
            # Détection d'encodage avancée
            encoding = self.file_detector.detect_encoding(file_path)
            
            # Détection de délimiteur avancée
            delimiter = self.file_detector.detect_delimiter(file_path, encoding)
            
            logging.info(f"Format CSV détecté - Encodage: {encoding}, Délimiteur: {delimiter}")
            
            return encoding, delimiter
            
        except Exception as e:
            logging.warning(f"Erreur lors de la détection du format CSV: {e}")
            
            # Fallback vers l'ancienne méthode
            for encoding in DEFAULT_ENCODINGS:
                try:
                    with open(file_path, 'r', encoding=encoding) as f:
                        sample = f.read(4096)
                        sniffer = csv.Sniffer()
                        dialect = sniffer.sniff(sample)
                        return encoding, dialect.delimiter
                except Exception:
                    continue
            
            raise ValueError(f"Impossible de déterminer l'encodage du fichier: {file_path}")
    
    def _remove_duplicate_rows_chunked_preserve_headers(self, data: List[List]) -> List[List]:
        """
        Supprime les doublons par chunks en préservant les en-têtes répétés.
        """
        if not data:
            return data
        
        # Si les en-têtes répétés sont activés, traitement spécial
        if self.repeat_headers:
            return self._remove_duplicates_preserve_headers(data)
        
        # Traitement normal pour les autres cas
        return self._remove_duplicate_rows_chunked(data)
    
    def _remove_duplicates_preserve_headers(self, data: List[List]) -> List[List]:
        """
        Supprime les doublons en préservant les en-têtes répétés.
        Détecte automatiquement les en-têtes répétés et les préserve.
        """
        if not data:
            return data
        
        if len(data) < self.header_rows:
            return data  # Pas assez de données pour avoir des en-têtes
            
        # Identifier le pattern des en-têtes originaux
        original_headers = []
        for i in range(self.header_rows):
            if i < len(data):
                header_tuple = tuple(str(cell) if cell is not None else '' for cell in data[i])
                original_headers.append(header_tuple)
        
        unique_data = []
        seen_data_rows = set()  # Pour tracking des vraies données
        i = 0
        
        while i < len(data):
            current_row = data[i]
            current_tuple = tuple(str(cell) if cell is not None else '' for cell in current_row)
            
            # Vérifier si c'est le début d'un bloc d'en-têtes répétés
            is_header_block = False
            if len(original_headers) > 0 and current_tuple == original_headers[0]:
                # Vérifier si les lignes suivantes correspondent aux en-têtes
                is_complete_header = True
                for j in range(1, len(original_headers)):
                    if i + j >= len(data):
                        is_complete_header = False
                        break
                    next_row_tuple = tuple(str(cell) if cell is not None else '' for cell in data[i + j])
                    if next_row_tuple != original_headers[j]:
                        is_complete_header = False
                        break
                
                if is_complete_header:
                    is_header_block = True
            
            if is_header_block:
                # C'est un bloc d'en-têtes répétés - toujours le préserver
                for j in range(len(original_headers)):
                    if i + j < len(data):
                        self._check_memory_before_append(data[i + j], unique_data)
                        unique_data.append(data[i + j])
                i += len(original_headers)  # Passer tout le bloc d'en-têtes
            else:
                # C'est une ligne de données - vérifier les doublons
                if current_tuple not in seen_data_rows:
                    seen_data_rows.add(current_tuple)
                    self._check_memory_before_append(current_row, unique_data)
                    unique_data.append(current_row)
                i += 1
        
        return unique_data
    
    def _check_memory_before_append(self, row: List, target_list: List[List]):
        """Vérifie la mémoire avant d'ajouter une ligne"""
        try:
            if self.is_memory_limit_reached([row], target_list):
                self.force_garbage_collection()
        except Exception as e:
            logging.warning(f"Erreur vérification mémoire: {e}")

    def _remove_duplicate_rows_chunked(self, data: List[List]) -> List[List]:
        """
        Supprime les doublons par chunks pour économiser la mémoire.
        """
        if not data:
            return data
        
        # Pour de petites listes, utiliser la méthode classique
        if len(data) < CHUNK_SIZE:
            return self._remove_duplicate_rows(data)
        
        # Traitement par chunks pour les grandes listes
        unique_data = []
        seen_hashes = set()
        
        for i in range(0, len(data), CHUNK_SIZE):
            chunk = data[i:i + CHUNK_SIZE]
            
            for row in chunk:
                row_hash = hash(tuple(str(cell) if cell is not None else '' for cell in row))
                if row_hash not in seen_hashes:
                    seen_hashes.add(row_hash)
                    unique_data.append(row)
            
            # Libération explicite du chunk après traitement
            del chunk
            
            # Nettoyage périodique pour éviter l'accumulation
            if i % (CHUNK_SIZE * 5) == 0:  # Tous les 5 chunks
                self.memory_manager.force_garbage_collection()
            elif self.memory_manager.is_memory_limit_reached():
                self.memory_manager.force_garbage_collection()
        
        return unique_data
    
    def _sort_data_chunked(self, data: List[List], headers: List[List]) -> List[List]:
        """
        Trie les données par chunks pour économiser la mémoire.
        """
        if not data or self.sort_column >= len(headers[-1]):
            return data
        
        # Pour de petites listes, utiliser la méthode classique
        if len(data) < CHUNK_SIZE:
            return self._sort_data(data, headers)
        
        # Tri externe pour les grandes listes avec gestion mémoire
        try:
            # Nettoyage préventif avant le tri (opération coûteuse)
            self.memory_manager.force_garbage_collection()
            
            # Séparer en-têtes et données
            header_count = len(headers) if headers else 0
            headers_part = data[:header_count] if header_count > 0 else []
            data_part = data[header_count:] if header_count > 0 else data
            
            # Trier seulement les données, pas les en-têtes
            sorted_data_part = sorted(data_part, key=lambda row: row[self.sort_column] if self.sort_column < len(row) and row[self.sort_column] is not None else '')
            
            # Recombiner : en-têtes + données triées
            sorted_data = headers_part + sorted_data_part
            
            # Libération explicite des données originales après tri
            del data
            
            # Nettoyage post-tri
            self.memory_manager.force_garbage_collection()
            
            return sorted_data
        except MemoryError:
            logging.warning("Mémoire insuffisante pour le tri, données non triées")
            return data
        except Exception as e:
            logging.warning(f"Erreur lors du tri: {e}, données non triées")
            return data
            
    def _process_excel_file(self, file_path, file_index, global_headers, preliminary_info, capture_preliminary=True):
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
        if file_index == 0 and self.header_start_row > 1 and capture_preliminary:
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
        if self.repeat_headers and file_index > 0:
            # Ajouter une ligne vide comme séparateur
            data.append([None] * (len(worksheet[self.header_start_row]) + (1 if self.filename_option != "none" else 0)))
            
            # Ajouter les en-têtes
            for row in range(self.header_start_row, self.header_start_row + self.header_rows):
                row_data = [cell.value for cell in worksheet[row]]
                if self.filename_option != "none":
                    row_data.append(None)  # Pas de nom de fichier dans l'en-tête répété
                data.append(row_data)
        
        # Ajout des données
        for row in worksheet.iter_rows(min_row=self.header_start_row + self.header_rows):
            row_data = [cell.value for cell in row]
            
            # Ajouter le nom du fichier selon l'option choisie
            if self.filename_option == "with_extension":
                row_data.append(filename)
            elif self.filename_option == "without_extension":
                # Enlever l'extension du nom de fichier
                filename_without_ext = os.path.splitext(filename)[0]
                row_data.append(filename_without_ext)
            # Si "none", on n'ajoute rien
            
            # Vérifier si la ligne n'est pas vide avant de l'ajouter
            if not self.remove_empty_rows or not all(
                cell is None or str(cell).strip() == "" 
                for cell in row_data[:-1 if self.filename_option != "none" else None]
            ):
                data.append(row_data)
                
        return data
        
    def _process_csv_file(self, file_path, file_index, global_headers, preliminary_info, capture_preliminary=True):
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
        if file_index == 0 and self.header_start_row > 1 and capture_preliminary:
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
        if self.repeat_headers and file_index > 0:
            # Ajouter une ligne vide comme séparateur
            file_data.append([None] * (len(file_headers[-1]) + (1 if self.filename_option != "none" else 0)))
            
            # Ajouter les en-têtes
            for header_row in file_headers:
                row_data = header_row.copy()
                if self.filename_option != "none":
                    row_data.append(None)
                file_data.append(row_data)
        
        # Ajout des données
        for row in range(self.header_start_row - 1 + self.header_rows, len(df)):
            row_data = df.iloc[row].tolist()
            
            # Ajouter le nom du fichier si demandé
            if self.filename_option != "none":
                if self.filename_option == "with_extension":
                    row_data.append(os.path.basename(file_path))
                elif self.filename_option == "without_extension":
                    filename_without_ext = os.path.splitext(os.path.basename(file_path))[0]
                    row_data.append(filename_without_ext)
            
            # Vérifier si la ligne n'est pas vide avant de l'ajouter
            if not self.remove_empty_rows or not all(
                pd.isna(cell) or str(cell).strip() == "" 
                for cell in row_data[:-1 if self.filename_option != "none" else None]
            ):
                file_data.append(row_data)
                
        return file_headers, file_data, file_merged_cells
    
    
    def _process_text_file(self, file_path, file_index, global_headers, preliminary_info, capture_preliminary=True):
        """
        Traite un fichier texte délimité (.tsv, .txt).
        
        Args:
            file_path: Chemin du fichier
            file_index: Index du fichier dans la liste
            global_headers: En-têtes déjà établies
            preliminary_info: Informations préliminaires
            
        Returns:
            Tuple contenant les en-têtes et données du fichier
        """
        # Détection du délimiteur selon l'extension
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext == '.tsv':
            delimiter = '\t'
        else:  # .txt
            # Détection automatique du délimiteur
            encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']
            delimiter = ','  # Par défaut
            
            for encoding in encodings:
                try:
                    with open(file_path, 'r', encoding=encoding) as f:
                        sample = f.read(4096)
                        sniffer = csv.Sniffer()
                        delimiter = sniffer.sniff(sample).delimiter
                        break
                except Exception:
                    continue
        
        # Lecture du fichier avec pandas
        encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']
        df = None
        
        for encoding in encodings:
            try:
                df = pd.read_csv(file_path, delimiter=delimiter, header=None, encoding=encoding)
                break
            except Exception:
                continue
        
        if df is None:
            raise ValueError(f"Impossible de lire le fichier avec les encodages disponibles")
        
        file_headers = []
        file_merged_cells = []  # Toujours vide pour les fichiers texte
        
        # Capture des informations préliminaires du premier fichier
        if file_index == 0 and self.header_start_row > 1 and capture_preliminary:
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
        if self.repeat_headers and file_index > 0:
            # Ajouter une ligne vide comme séparateur
            file_data.append([None] * (len(file_headers[-1]) + (1 if self.filename_option != "none" else 0)))
            
            # Ajouter les en-têtes
            for header_row in file_headers:
                row_data = header_row.copy()
                if self.filename_option != "none":
                    row_data.append(None)
                file_data.append(row_data)
        
        # Ajout des données
        for row in range(self.header_start_row - 1 + self.header_rows, len(df)):
            row_data = df.iloc[row].tolist()
            
            # Ajouter le nom du fichier si demandé
            if self.filename_option != "none":
                if self.filename_option == "with_extension":
                    row_data.append(os.path.basename(file_path))
                elif self.filename_option == "without_extension":
                    filename_without_ext = os.path.splitext(os.path.basename(file_path))[0]
                    row_data.append(filename_without_ext)
            
            # Vérifier si la ligne n'est pas vide avant de l'ajouter
            if not self.remove_empty_rows or not all(
                pd.isna(cell) or str(cell).strip() == "" 
                for cell in row_data[:-1 if self.filename_option != "none" else None]
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
            sort_idx = self.sort_column  # Pas de -1 car sort_column est déjà 0-based
            
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
                # Tri de toutes les lignes comme des données
                return sorted(data, key=lambda x: (x[sort_idx] is None, x[sort_idx]) if sort_idx < len(x) else (''))
                
        except Exception as e:
            logging.warning(f"Erreur lors du tri : {str(e)}")
            return data

    def _load_preliminary_info(self):
        """
        Charge les informations préliminaires depuis le fichier source spécifié.
        
        Returns:
            Liste des lignes préliminaires
        """
        preliminary_info = []
        file_path = os.path.join(self.directory, self.preliminary_source_file)
        
        if self.preliminary_source_file.lower().endswith(('.xlsx', '.xlsm', '.xltx', '.xltm', '.xls')):
            # Fichier Excel
            wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            ws = wb.active
            
            for row in range(1, self.header_start_row):
                row_data = []
                for cell in ws[row]:
                    row_data.append(cell.value)
                preliminary_info.append(row_data)
            
            wb.close()
            
        elif self.preliminary_source_file.lower().endswith(('.csv', '.tsv', '.txt')):
            # Fichier CSV/texte
            encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']
            
            # Détection du délimiteur
            delimiter = ','
            if self.preliminary_source_file.lower().endswith('.tsv'):
                delimiter = '\t'
            else:
                for encoding in encodings:
                    try:
                        with open(file_path, 'r', encoding=encoding) as f:
                            sample = f.read(4096)
                            sniffer = csv.Sniffer()
                            delimiter = sniffer.sniff(sample).delimiter
                            break
                    except Exception:
                        continue
            
            # Lecture des lignes préliminaires
            df = pd.read_csv(file_path, delimiter=delimiter, header=None, encoding=encoding)
            
            for row in range(0, self.header_start_row - 1):
                if row < len(df):
                    preliminary_info.append(df.iloc[row].tolist())
        
        return preliminary_info

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
    def write_data(worksheet, data, start_row, date_format="FRENCH"):
        """
        Écrit les données dans la feuille de calcul avec le style approprié.
        
        Args:
            worksheet: Feuille de calcul à modifier
            data: Données à écrire
            start_row: Ligne de début pour écrire les données
            date_format: Format de date à utiliser (par défaut FRENCH)
            
        Returns:
            Ligne courante après écriture
        """
        current_row = start_row
        
        excel_date_format = DATE_FORMATS.get(date_format, DATE_FORMATS["FRENCH"])["excel_format"]
        
        for row_data in data:
            for col_idx, value in enumerate(row_data, 1):
                cell = worksheet.cell(row=current_row, column=col_idx, value=value)
                cell.border = ExcelFormatter.DATA_BORDER
                
                # Appliquer le format de date si nécessaire
                if isinstance(value, datetime):
                    cell.number_format = excel_date_format
                    
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
            # Calculer le décalage en lignes
            row_offset = header_start_new - header_start_original
            
            # Ajuster les numéros de ligne
            adjusted_min_row = max(1, merged_range.min_row + row_offset)
            adjusted_max_row = max(1, merged_range.max_row + row_offset)
            
            # Vérifier que la plage ajustée est valide
            if adjusted_min_row <= adjusted_max_row:
                adjusted_range = openpyxl.worksheet.cell_range.CellRange(
                    min_col=merged_range.min_col,
                    min_row=adjusted_min_row,
                    max_col=merged_range.max_col,
                    max_row=adjusted_max_row
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
                    if cell.value is not None and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except (TypeError, AttributeError, ValueError) as e:
                    # TypeError: cell.value non stringifiable
                    # AttributeError: cell.value n'a pas la méthode attendue
                    # ValueError: Problème de conversion string
                    # Ignorer silencieusement ces cellules
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

class PreviewDialog(QDialog):
    """
    Boîte de dialogue pour prévisualiser les données avant compilation.
    """
    
    def __init__(self, parent, directory, files, header_start_row, header_rows):
        """
        Initialise la boîte de dialogue de prévisualisation.
        
        Args:
            parent: Widget parent
            directory: Répertoire contenant les fichiers
            files: Liste des fichiers à prévisualiser
            header_start_row: Ligne de début des en-têtes
            header_rows: Nombre de lignes d'en-tête
        """
        super().__init__(parent)
        self.translate = TranslationManager().get_text
        self.setWindowTitle(self.translate("preview"))
        self.resize(1000, 700)
        
        self.directory = directory
        self.files = files
        self.header_start_row = header_start_row
        self.header_rows = header_rows
        
        self.selected_file = None
        self.preview_data = None
        self.max_preview_rows = 200  # Limite de lignes pour la prévisualisation
        
        self.init_ui()
        
        # Charger le premier fichier s'il existe
        if files:
            self.file_combo.setCurrentIndex(0)
            self.selected_file = files[0]
            self.load_preview()
    
    def init_ui(self):
        """Initialise l'interface utilisateur de la boîte de dialogue."""
        layout = QVBoxLayout()
        
        # Sélection du fichier à prévisualiser
        file_layout = QHBoxLayout()
        file_label = QLabel(self.translate("filename") + ":")
        self.file_combo = QComboBox()
        self.file_combo.addItems(self.files)
        self.file_combo.currentIndexChanged.connect(self.on_file_changed)
        
        self.refresh_button = QPushButton(self.translate("refresh_preview"))
        self.refresh_button.clicked.connect(self.load_preview)
        
        file_layout.addWidget(file_label)
        file_layout.addWidget(self.file_combo, 1)
        file_layout.addWidget(self.refresh_button)
        
        layout.addLayout(file_layout)
        
        # Tableau de prévisualisation
        self.preview_table = self.create_responsive_table()
        self.preview_table.setAlternatingRowColors(True)
        layout.addWidget(self.preview_table)
        
        # Informations sur la prévisualisation
        self.info_label = QLabel(self.translate("preview_limited", self.max_preview_rows))
        layout.addWidget(self.info_label)
        
        # Bouton de fermeture
        button_layout = QHBoxLayout()
        close_button = QPushButton(self.translate("ok"))
        close_button.clicked.connect(self.accept)
        button_layout.addStretch()
        button_layout.addWidget(close_button)
        button_layout.addStretch()
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def on_file_changed(self, index):
        """
        Appelé lorsque l'utilisateur change de fichier dans le combobox.
        
        Args:
            index: Index du fichier sélectionné
        """
        if index >= 0 and index < len(self.files):
            self.selected_file = self.files[index]
            self.load_preview()
    
    def load_preview(self):
        """Charge et affiche la prévisualisation du fichier sélectionné."""
        if not self.selected_file:
            return
        
        file_path = os.path.join(self.directory, self.selected_file)
        
        try:
            if self.selected_file.lower().endswith(('.xlsx', '.xls')):
                self.load_excel_preview(file_path)
            elif self.selected_file.lower().endswith('.csv'):
                self.load_csv_preview(file_path)
        except Exception as e:
            QMessageBox.warning(
                self,
                self.translate("error"),
                f"{self.translate('error')}: {str(e)}"
            )
    
    def load_excel_preview(self, file_path):
        """
        Charge la prévisualisation d'un fichier Excel.
        
        Args:
            file_path: Chemin du fichier Excel à prévisualiser
        """
        try:
            # Charger le fichier Excel
            wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            ws = wb.active
            
            # Obtenir les données (limité au nombre max de lignes)
            preview_data = []
            headers = []
            
            # Récupérer les en-têtes
            for row in range(self.header_start_row, self.header_start_row + self.header_rows):
                header_row = []
                for cell in ws[row]:
                    header_row.append(cell.value)
                headers.append(header_row)
            
            # Récupérer les données
            row_count = 0
            for row in ws.iter_rows(min_row=self.header_start_row + self.header_rows):
                if row_count >= self.max_preview_rows:
                    break
                
                row_data = [cell.value for cell in row]
                preview_data.append(row_data)
                row_count += 1
            
            # Afficher les données dans le tableau
            self.display_preview(headers, preview_data)
            
            wb.close()
        except Exception as e:
            raise ValueError(f"Erreur lors de la lecture du fichier Excel: {str(e)}")
    
    def load_csv_preview(self, file_path):
        """
        Charge la prévisualisation d'un fichier CSV.
        
        Args:
            file_path: Chemin du fichier CSV à prévisualiser
        """
        try:
            # Détection de l'encodage du fichier
            encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']
            delimiter = ','
            
            for encoding in encodings:
                try:
                    with open(file_path, 'r', encoding=encoding) as f:
                        sample = f.read(4096)
                        sniffer = csv.Sniffer()
                        delimiter = sniffer.sniff(sample).delimiter
                        break
                except Exception:
                    continue
            
            # Lecture du CSV avec pandas
            df = pd.read_csv(file_path, delimiter=delimiter, header=None, encoding=encoding)
            
            # Limiter le nombre de lignes
            if len(df) > self.max_preview_rows + self.header_start_row + self.header_rows:
                df = df.iloc[:(self.max_preview_rows + self.header_start_row + self.header_rows)]
            
            # Récupérer les en-têtes
            headers = []
            for row in range(self.header_start_row - 1, self.header_start_row - 1 + self.header_rows):
                if row < len(df):
                    headers.append(df.iloc[row].tolist())
            
            # Récupérer les données
            preview_data = []
            for row in range(self.header_start_row - 1 + self.header_rows, len(df)):
                preview_data.append(df.iloc[row].tolist())
            
            # Afficher les données dans le tableau
            self.display_preview(headers, preview_data)
            
        except Exception as e:
            raise ValueError(f"Erreur lors de la lecture du fichier CSV: {str(e)}")
    
    def _combine_multi_headers(self, headers):
        """
        Combine les en-têtes multi-lignes en un seul en-tête pour l'affichage.
        
        Args:
            headers: Liste des lignes d'en-têtes
            
        Returns:
            Liste des en-têtes combinées
        """
        if not headers or len(headers) == 1:
            return headers[0] if headers else []
        
        # Déterminer le nombre de colonnes maximum
        max_cols = max(len(row) for row in headers)
        combined = []
        
        for col_idx in range(max_cols):
            combined_cell = []
            for row in headers:
                if col_idx < len(row) and row[col_idx] not in [None, '']:
                    cell_value = str(row[col_idx]).strip()
                    if cell_value and cell_value not in combined_cell:
                        combined_cell.append(cell_value)
            
            # Combiner les parties non vides avec un séparateur
            combined.append(' - '.join(combined_cell) if combined_cell else '')
        
        return combined
    
    def display_preview(self, headers, data):
        """
        Affiche les données dans le tableau de prévisualisation.
        
        Args:
            headers: Données d'en-tête
            data: Données à afficher
        """
        if not headers or not headers[0]:
            return
        
        # Configurer le tableau
        self.preview_table.clear()
        self.preview_table.setRowCount(len(data))
        
        # Combiner les en-têtes multi-lignes pour l'affichage
        combined_headers = self._combine_multi_headers(headers)
        self.preview_table.setColumnCount(len(combined_headers))
        
        # Définir les en-têtes des colonnes combinées
        self.preview_table.setHorizontalHeaderLabels([str(h) if h is not None else "" for h in combined_headers])
        
        # Ajouter les données
        for row_idx, row_data in enumerate(data):
            for col_idx, value in enumerate(row_data[:len(combined_headers)]):  # Limiter aux colonnes d'en-tête
                # Créer un élément de tableau avec la valeur formatée
                item = QTableWidgetItem(str(value) if value is not None else "")
                self.preview_table.setItem(row_idx, col_idx, item)
        
        # Ajuster la taille des colonnes
        self.preview_table.resizeColumnsToContents()
        
        # Mettre à jour l'étiquette d'information
        if len(data) >= self.max_preview_rows:
            self.info_label.setText(self.translate("preview_limited", self.max_preview_rows))
        else:
            self.preview_info_label.setText(f"{len(data)} lignes affichées")


class DateFormatDialog(QDialog):
    """
    Boîte de dialogue pour choisir le format de date.
    """
    
    def __init__(self, parent, current_format="FRENCH"):
        """
        Initialise la boîte de dialogue de format de date.
        
        Args:
            parent: Widget parent
            current_format: Format de date actuellement sélectionné
        """
        super().__init__(parent)
        self.translate = TranslationManager().get_text
        self.setWindowTitle(self.translate("date_format"))
        self.resize(500, 400)
        
        self.current_format = current_format
        self.custom_format = ""
        self.selected_format = current_format
        
        self.init_ui()
    
    def init_ui(self):
        """Initialise l'interface utilisateur de la boîte de dialogue."""
        layout = QVBoxLayout()
        
        # Groupe de formats prédéfinis
        group = QGroupBox(self.translate("date_format_options"))
        group_layout = QVBoxLayout()
        
        # Créer les boutons radio pour chaque format prédéfini
        self.radio_group = QButtonGroup(self)
        formats = [
            ("STANDARD", "date_format_standard"),
            ("FRENCH", "date_format_french"),
            ("US", "date_format_us"),
            ("DATETIME", "date_format_datetime"),
            ("DATETIME_FRENCH", "date_format_datetime_french"),
            ("DATE_ONLY", "date_format_date_only"),
            ("TIME_ONLY", "date_format_time_only"),
            ("SHORT", "date_format_short"),
            ("CUSTOM", "date_format_custom")
        ]
        
        self.radio_buttons = {}
        
        for i, (format_key, label_key) in enumerate(formats):
            radio = QRadioButton(self.translate(label_key))
            self.radio_group.addButton(radio, i)
            self.radio_buttons[format_key] = radio
            
            if format_key == "CUSTOM":
                custom_layout = QHBoxLayout()
                custom_layout.addWidget(radio)
                self.custom_edit = QLineEdit()
                self.custom_edit.setPlaceholderText("dd/MM/yyyy HH:mm:ss")
                self.custom_edit.setEnabled(False)
                custom_layout.addWidget(self.custom_edit)
                group_layout.addLayout(custom_layout)
            else:
                group_layout.addWidget(radio)
        
        group.setLayout(group_layout)
        layout.addWidget(group)
        
        # Exemple avec la date actuelle
        example_layout = QHBoxLayout()
        example_layout.addWidget(QLabel(self.translate("preview") + ":"))
        self.example_label = QLabel()
        self.update_example()
        example_layout.addWidget(self.example_label)
        layout.addLayout(example_layout)
        
        # Boutons OK/Annuler
        button_layout = QHBoxLayout()
        ok_button = QPushButton(self.translate("ok"))
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton(self.translate("cancel"))
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
        # Connecter les signaux
        self.radio_group.buttonClicked.connect(self.on_format_changed)
        self.custom_edit.textChanged.connect(self.on_custom_format_changed)
        
        # Sélectionner le format actuel
        if self.current_format in self.radio_buttons:
            self.radio_buttons[self.current_format].setChecked(True)
            if self.current_format == "CUSTOM":
                self.custom_edit.setEnabled(True)
    
    def on_format_changed(self, button):
        """
        Appelé lorsque l'utilisateur change de format.
        
        Args:
            button: Bouton radio sélectionné
        """
        for format_key, radio in self.radio_buttons.items():
            if radio == button:
                self.selected_format = format_key
                if format_key == "CUSTOM":
                    self.custom_edit.setEnabled(True)
                else:
                    self.custom_edit.setEnabled(False)
                break
        
        self.update_example()
    
    def on_custom_format_changed(self, text):
        """
        Appelé lorsque l'utilisateur modifie le format personnalisé.
        
        Args:
            text: Nouveau texte du format personnalisé
        """
        self.custom_format = text
        self.update_example()
    
    def update_example(self):
        """Met à jour l'exemple de format de date."""
        now = datetime.now()
        
        if self.selected_format == "CUSTOM":
            try:
                # Convertir le format de l'utilisateur en format de date Python
                user_format = self.custom_format
                # Remplacer les tokens de format
                py_format = user_format.replace("dd", "%d").replace("MM", "%m").replace("yyyy", "%Y")
                py_format = py_format.replace("HH", "%H").replace("mm", "%M").replace("ss", "%S")
                py_format = py_format.replace("yy", "%y")
                
                formatted_date = now.strftime(py_format)
                self.example_label.setText(formatted_date)
            except Exception:
                self.example_label.setText("Format invalide")
        else:
            # Utiliser le format prédéfini
            date_format = DATE_FORMATS[self.selected_format]["format"]
            
            # Convertir en format Python
            py_format = date_format.replace("dd", "%d").replace("MM", "%m").replace("yyyy", "%Y")
            py_format = py_format.replace("HH", "%H").replace("mm", "%M").replace("ss", "%S")
            py_format = py_format.replace("yy", "%y")
            
            formatted_date = now.strftime(py_format)
            self.example_label.setText(formatted_date)
    
    def get_selected_format(self):
        """
        Obtient le format de date sélectionné.
        
        Returns:
            str: Clé du format sélectionné
        """
        if self.selected_format == "CUSTOM":
            DATE_FORMATS["CUSTOM"]["format"] = self.custom_format
            
            # Générer un format Excel personnalisé
            excel_format = self.custom_format
            excel_format = excel_format.replace("dd", "dd").replace("MM", "mm").replace("yyyy", "yyyy")
            excel_format = excel_format.replace("HH", "hh").replace("mm", "mm").replace("ss", "ss")
            excel_format = excel_format.replace("yy", "yy")
            
            DATE_FORMATS["CUSTOM"]["excel_format"] = excel_format
            
        return self.selected_format


class LanguageDialog(QDialog):
    """
    Boîte de dialogue pour choisir la langue de l'interface.
    """
    
    def __init__(self, parent):
        """
        Initialise la boîte de dialogue de choix de langue.
        
        Args:
            parent: Widget parent
        """
        super().__init__(parent)
        self.translate = TranslationManager().get_text
        self.setWindowTitle(self.translate("languages"))
        self.resize(300, 200)
        
        self.selected_language = TranslationManager().current_language
        
        self.init_ui()
    
    def init_ui(self):
        """Initialise l'interface utilisateur de la boîte de dialogue."""
        layout = QVBoxLayout()
        
        # Groupe des langues
        group = QGroupBox(self.translate("language"))
        group_layout = QVBoxLayout()
        
        # Créer les boutons radio pour chaque langue
        self.radio_group = QButtonGroup(self)
        
        # Français
        self.radio_fr = QRadioButton(self.translate("french"))
        self.radio_group.addButton(self.radio_fr, 0)
        group_layout.addWidget(self.radio_fr)
        
        # Anglais
        self.radio_en = QRadioButton(self.translate("english"))
        self.radio_group.addButton(self.radio_en, 1)
        group_layout.addWidget(self.radio_en)
        
        # Espagnol
        self.radio_es = QRadioButton(self.translate("spanish"))
        self.radio_group.addButton(self.radio_es, 2)
        group_layout.addWidget(self.radio_es)
        
        # Allemand
        self.radio_de = QRadioButton(self.translate("german"))
        self.radio_group.addButton(self.radio_de, 3)
        group_layout.addWidget(self.radio_de)
        
        group.setLayout(group_layout)
        layout.addWidget(group)
        
        # Boutons OK/Annuler
        button_layout = QHBoxLayout()
        apply_button = QPushButton(self.translate("apply"))
        apply_button.clicked.connect(self.accept)
        cancel_button = QPushButton(self.translate("cancel"))
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(apply_button)
        button_layout.addWidget(cancel_button)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
        # Sélectionner la langue actuelle
        if self.selected_language == "fr":
            self.radio_fr.setChecked(True)
        elif self.selected_language == "en":
            self.radio_en.setChecked(True)
        elif self.selected_language == "es":
            self.radio_es.setChecked(True)
        elif self.selected_language == "de":
            self.radio_de.setChecked(True)
        
        # Connecter les signaux
        self.radio_group.buttonClicked.connect(self.on_language_changed)
    
    def on_language_changed(self, button):
        """
        Appelé lorsque l'utilisateur change de langue.
        
        Args:
            button: Bouton radio sélectionné
        """
        if button == self.radio_fr:
            self.selected_language = "fr"
        elif button == self.radio_en:
            self.selected_language = "en"
        elif button == self.radio_es:
            self.selected_language = "es"
        elif button == self.radio_de:
            self.selected_language = "de"
    
    def get_selected_language(self):
        """
        Obtient la langue sélectionnée.
        
        Returns:
            str: Code de la langue sélectionnée
        """
        return self.selected_language


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
        self.translate = TranslationManager().get_text
        self.setWindowTitle(self.translate("verification_title"))
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
        header_label = QLabel(f"<h2>{self.translate('verification_header', total_files)}</h2>")
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(header_label)
        
        compat_percent = len(self.compatible_files) * 100 / total_files if total_files > 0 else 0
        incompat_percent = len(self.incompatible_files) * 100 / total_files if total_files > 0 else 0
        
        stats_label = QLabel(
            f"<div style='text-align:center; margin:10px 0;'>"
            f"<span style='color:#{COLORS['SUCCESS']}; font-weight:bold;'>{self.translate('compilable', len(self.compatible_files), compat_percent)}</span> | "
            f"<span style='color:#{COLORS['WARNING']}; font-weight:bold;'>{self.translate('not_compilable', len(self.incompatible_files), incompat_percent)}</span>"
            f"</div>"
        )
        stats_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(stats_label)
        
        # Tableau des fichiers incompatibles
        if self.incompatible_files:
            group_incompatible = QGroupBox(self.translate("non_compilable_files"))
            group_layout = QVBoxLayout()
            
            # Instructions pour résoudre les problèmes
            help_label = QLabel(
                f"{self.translate('error_resolution_tips')}"
                "<ul>"
                f"<li><b>{self.translate('open_file_tip')}</b></li>"
                f"<li><b>{self.translate('protected_file_tip')}</b></li>"
                f"<li><b>{self.translate('header_structure_tip')}</b></li>"
                f"<li><b>{self.translate('encoding_error_tip')}</b></li>"
                "</ul>"
            )
            help_label.setWordWrap(True)
            group_layout.addWidget(help_label)
            
            table = QTableWidget()
            table.setColumnCount(2)
            table.setHorizontalHeaderLabels([self.translate("filename"), self.translate("detected_issue")])
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
            group_compatible = QGroupBox(self.translate("compilable_files"))
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
            ignore_button = QPushButton(self.translate("ignore_non_compilable"))
            ignore_button.setStyleSheet(f"background-color: #{COLORS['INFO']}; color: white;")
            ignore_button.clicked.connect(self.continue_with_compatible_only)
            button_layout.addWidget(ignore_button)

        cancel_button = QPushButton(self.translate("cancel"))
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
        self.translate = TranslationManager().get_text
        self.setWindowTitle(self.translate("compilation_report"))
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
        header_label = QLabel(f"<h2>{self.translate('compilation_result')}</h2>")
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(header_label)
        
        # Affichage du chemin du fichier de sortie
        output_label = QLabel(f"<div style='text-align:center; margin:5px 0;'>"
                             f"<b>{self.translate('generated_file')}</b> {self.output_path}"
                             f"</div>")
        output_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(output_label)
        
        # Statistiques
        percentage = len(self.successful_files) * 100 / total_files if total_files > 0 else 0
        stats_label = QLabel(f"<div style='text-align:center; margin:10px 0; font-size:16px;'>"
                            f"<b>{self.translate('compilation_rate', percentage, len(self.successful_files), total_files)}</b>"
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
            
            success_message = QLabel(f"<div style='text-align:center; color:green; font-weight:bold;'>"
                                    f"{self.translate('all_files_compiled')}"
                                    f"</div>")
            success_message.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(success_message)
        
        # Liste des fichiers compilés
        if self.successful_files:
            success_group = QGroupBox(self.translate("compiled_files", len(self.successful_files)))
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
            failed_group = QGroupBox(self.translate("not_compiled_files", len(self.failed_files)))
            failed_layout = QVBoxLayout()
            
            table = QTableWidget()
            table.setColumnCount(2)
            table.setHorizontalHeaderLabels([self.translate("filename"), self.translate("error")])
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
        button = QPushButton(self.translate("ok"))
        button.clicked.connect(self.accept)
        layout.addWidget(button, alignment=Qt.AlignmentFlag.AlignCenter)
        
        self.setLayout(layout)


class IPWarningDialog(QDialog):
    """
    Boîte de dialogue d'avertissement concernant la propriété intellectuelle.
    """
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.translate = TranslationManager().get_text
        self.setWindowTitle(self.translate("ip_warning_title"))
        self.setMinimumSize(800, 600)
        
        self.init_ui()
    
    def init_ui(self):
        """Initialise l'interface utilisateur de la boîte de dialogue."""
        layout = QVBoxLayout()
        
        # Titre
        title_label = QLabel("<h1>" + self.translate("warning") + "</h1>")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)
        
        # Contenu
        content = QLabel(
            "<p style='font-size: 14px; line-height: 1.5;'>"
            f"{self.translate('ip_warning_content')}<br><br>"
            "<div style='text-align: center; font-weight: bold; font-size: 16px; margin: 20px 0;'>"
            f"{self.translate('developer_name')}<br>"
            f"{self.translate('developer_title')}<br><br>"
            "</div>"
            f"{self.translate('copyright_notice')}<br><br>"
            
            f"<span style='color: #{COLORS['WARNING']}; font-weight: bold;'>{self.translate('warning_important')}</span><br>"
            "<ul>"
            f"<li>{self.translate('unauthorized_reproduction')}</li>"
            f"<li>{self.translate('license_terms')}</li>"
            "</ul><br>"
            
            f"{self.translate('contact_info')}"
            
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
        
        accept_button = QPushButton(self.translate("accept_conditions"))
        accept_button.clicked.connect(self.accept)
        accept_button.setDefault(True)
        
        exit_button = QPushButton(self.translate("quit_app"))
        exit_button.clicked.connect(self.reject)
        
        button_layout.addWidget(accept_button)
        button_layout.addWidget(exit_button)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)


class ModernExcelCompilerApp(QMainWindow):
    """
    Classe principale de l'application de compilation Excel.
    """
    # Signaux pour communication thread-safe
    progress_update_signal = pyqtSignal(int, str)  # progress, detail
    compilation_error_signal = pyqtSignal(str)  # error_message
    compilation_success_signal = pyqtSignal(str, int, int, str)  # message, successful, failed, output_path
    progress_finish_signal = pyqtSignal(bool, str)  # success, message
    progress_reset_signal = pyqtSignal()
    button_enable_signal = pyqtSignal(bool)  # enable/disable compile button
    
    def __init__(self):
        """Initialise l'application."""
        super().__init__()
        
        # Initialiser le gestionnaire de responsivité
        self.responsive_manager = ResponsiveManager()
        screen_info = self.responsive_manager.detect_screen_size(self)
        logging.info(f"Écran détecté: {screen_info['width']}x{screen_info['height']} - Breakpoint: {screen_info['breakpoint']}")
        
        # Initialiser le validateur en temps réel
        self.real_time_validator = RealTimeValidator(self)
        self.real_time_validator.validation_changed.connect(self.on_validation_changed)
        
        # Initialiser le gestionnaire de traduction
        self.translate = TranslationManager().get_text
        TranslationManager().register_language_changed_callback(self.update_ui_language)
        
        # Initialiser les variables
        self.setup_variables()
        
        # Configuration timeout pour workers
        self.worker_timeout = 300  # 5 minutes par défaut
        self.timeout_timer = None
        
        # Configurer l'interface utilisateur
        self.setup_ui()
        self.create_menu()
        self.create_toolbar()
        self.create_main_layout()
        self.connect_signals()
        
        # Charger les paramètres sauvegardés
        SettingsManager().load_settings(self)
        
        # Configurer la validation en temps réel
        self.setup_real_time_validation()
        
        logging.info("Application démarrée")
        
        # Afficher l'avertissement de propriété intellectuelle au démarrage
        self.show_ip_warning()
        
        # Afficher le tutoriel si c'est la première utilisation
        QTimer.singleShot(1000, self.check_and_show_tutorial)
    
    def resizeEvent(self, event):
        """Gestionnaire d'événement de redimensionnement pour la responsivité."""
        super().resizeEvent(event)
        
        # Redétecter la taille d'écran si nécessaire
        new_size = event.size()
        if new_size.width() != self.size().width() or new_size.height() != self.size().height():
            # Mise à jour du gestionnaire de responsivité si la fenêtre change beaucoup
            old_breakpoint = self.responsive_manager.current_breakpoint
            self.responsive_manager.detect_screen_size(self)
            
            # Si le breakpoint a changé, reappliquer les styles
            if old_breakpoint != self.responsive_manager.current_breakpoint:
                self.apply_responsive_stylesheet()
                logging.info(f"Breakpoint changé: {old_breakpoint} -> {self.responsive_manager.current_breakpoint}")
    
    def create_responsive_table(self, parent=None) -> ResponsiveTableWidget:
        """Crée un tableau responsive configuré avec le gestionnaire de responsivité."""
        table = ResponsiveTableWidget(parent)
        table.set_responsive_manager(self.responsive_manager)
        return table
    
    def setup_variables(self):
        """Initialise les variables de l'application."""
        self.directory = ""
        self.files = []
        self.compilation_worker = None
        self.verification_enabled = True  # Par défaut, la vérification préliminaire est activée
        self.date_format = "FRENCH"  # Format de date par défaut
        self.filename_option = "none"  # variable : "none", "with_extension", "without_extension"
        self.include_preliminary = False  # Inclure les lignes préliminaires
        self.preliminary_source_file = ""  # Fichier source pour les lignes préliminaires
        
        # Initialiser le système de prévisualisation optimisé
        self.preview_cache = PreviewCache(max_size=15)
        self.preview_worker = None
        
        # Initialiser les caches de performance pour répétitions
        self.validation_cache = ValidationCache(max_size=100)
        self.similarity_detector = SimilarityDetector()
        
        # Initialiser les systèmes de surveillance
        self.setup_monitoring_systems()

    def setup_monitoring_systems(self):
        """Initialise les systèmes de surveillance et monitoring"""
        try:
            # Initialiser le logger structuré
            self.structured_logger = StructuredLogger("ExcelCompiler")
            self.structured_logger.log_structured("INFO", "application_started", 
                                                version="3.1", 
                                                system_info=platform.platform())
            
            # Initialiser le moniteur de performance
            self.performance_monitor = PerformanceMonitor()
            
            # Initialiser le moniteur de santé
            self.health_monitor = HealthMonitor()
            
            # Enregistrer les health checks de base
            self.register_health_checks()
            
            # Démarrer la surveillance
            self.health_monitor.start_monitoring()
            
            # Démarrer le monitoring global
            if MONITORING_AVAILABLE:
                start_monitoring()
                log_event("application_started", {
                    "version": "3.1",
                    "system_info": platform.platform(),
                    "timestamp": datetime.now().isoformat()
                })
            
            # Initialiser les systèmes de robustesse
            self.setup_robustness_systems()
            
            logging.info("Systèmes de surveillance initialisés avec succès")
            
        except Exception as e:
            logging.error(f"Erreur lors de l'initialisation de la surveillance: {e}")
            # Continuer sans surveillance en cas d'erreur
            self.structured_logger = None
            self.performance_monitor = None
            self.health_monitor = None
    
    def register_health_checks(self):
        """Enregistre les health checks automatiques"""
        if not self.health_monitor:
            return
            
        # Health check mémoire
        self.health_monitor.register_health_check(
            "memory_usage", 
            self._check_memory_health, 
            interval=60  # Toutes les 60 secondes
        )
        
        # Health check espace disque
        self.health_monitor.register_health_check(
            "disk_space", 
            self._check_disk_health, 
            interval=300  # Toutes les 5 minutes
        )
        
        # Health check GUI responsiveness
        self.health_monitor.register_health_check(
            "gui_responsiveness", 
            self._check_gui_health, 
            interval=30  # Toutes les 30 secondes
        )
    
    def _check_memory_health(self) -> HealthCheckResult:
        """Vérifie l'utilisation mémoire"""
        try:
            memory = psutil.virtual_memory()
            percent_used = memory.percent
            
            if percent_used >= ALERT_MEMORY_THRESHOLD:
                status = "critical" if percent_used >= 95 else "warning"
                message = f"Utilisation mémoire élevée: {percent_used:.1f}%"
            else:
                status = "healthy"
                message = f"Utilisation mémoire normale: {percent_used:.1f}%"
            
            return HealthCheckResult(
                timestamp=datetime.now().isoformat(),
                component="memory",
                status=status,
                message=message,
                metrics={
                    "percent_used": percent_used,
                    "total_gb": memory.total / (1024**3),
                    "available_gb": memory.available / (1024**3)
                }
            )
        except Exception as e:
            return HealthCheckResult(
                timestamp=datetime.now().isoformat(),
                component="memory",
                status="critical",
                message=f"Erreur lors de la vérification mémoire: {e}"
            )
    
    def _check_disk_health(self) -> HealthCheckResult:
        """Vérifie l'espace disque disponible"""
        try:
            disk = psutil.disk_usage('/')
            percent_used = (disk.used / disk.total) * 100
            free_gb = disk.free / (1024**3)
            
            if percent_used >= 95:
                status = "critical"
                message = f"Espace disque critique: {percent_used:.1f}% utilisé"
            elif percent_used >= 85:
                status = "warning"
                message = f"Espace disque faible: {percent_used:.1f}% utilisé"
            else:
                status = "healthy"
                message = f"Espace disque suffisant: {free_gb:.1f}GB disponible"
            
            return HealthCheckResult(
                timestamp=datetime.now().isoformat(),
                component="disk",
                status=status,
                message=message,
                metrics={
                    "percent_used": percent_used,
                    "free_gb": free_gb,
                    "total_gb": disk.total / (1024**3)
                }
            )
        except Exception as e:
            return HealthCheckResult(
                timestamp=datetime.now().isoformat(),
                component="disk",
                status="critical",
                message=f"Erreur lors de la vérification disque: {e}"
            )
    
    def _check_gui_health(self) -> HealthCheckResult:
        """Vérifie la responsivité de l'interface"""
        try:
            # Vérifier que le thread principal répond
            if not ThreadSafeGUIHelper.is_main_thread():
                # Nous ne sommes pas dans le thread principal, ce qui est normal pour ce check
                pass
            
            # Vérifier l'état de l'application
            app = QApplication.instance()
            if not app:
                status = "critical"
                message = "Application QT non accessible"
            else:
                status = "healthy"
                message = "Interface utilisateur responsive"
            
            return HealthCheckResult(
                timestamp=datetime.now().isoformat(),
                component="gui",
                status=status,
                message=message,
                metrics={
                    "main_thread_active": ThreadSafeGUIHelper.is_main_thread(),
                    "qt_app_active": app is not None
                }
            )
        except Exception as e:
            return HealthCheckResult(
                timestamp=datetime.now().isoformat(),
                component="gui",
                status="warning",
                message=f"Vérification GUI partielle: {e}"
            )

    def setup_robustness_systems(self):
        """Initialise les systèmes de robustesse et gestion d'erreurs"""
        try:
            # Initialiser le gestionnaire de configuration
            self.config_manager = ConfigurationManager()
            self.app_config = self.config_manager.load_configuration()
            
            # Appliquer la configuration chargée
            self.apply_configuration()
            
            # Initialiser le gestionnaire d'erreurs avancé
            self.error_handler = AdvancedErrorHandler()
            
            # Initialiser le gestionnaire de sauvegarde automatique
            self.auto_save_manager = AutoSaveManager()
            self.auto_save_manager.enable_auto_save(self.app_config.auto_save_enabled)
            self.auto_save_manager.set_backup_interval(self.app_config.auto_save_interval)
            
            # Vérifier s'il y a une sauvegarde de session à récupérer
            if self.app_config.auto_recovery_enabled:
                self.check_recovery_needed()
            
            logging.info("Systèmes de robustesse initialisés avec succès")
            
        except Exception as e:
            logging.error(f"Erreur lors de l'initialisation de la robustesse: {e}")
            # Continuer avec des systèmes par défaut
            self.config_manager = None
            self.error_handler = None
            self.auto_save_manager = None
    
    def apply_configuration(self):
        """Applique la configuration chargée à l'application"""
        if not self.app_config:
            return
        
        try:
            # Appliquer les paramètres de performance
            global MAX_THREADS, CHUNK_SIZE
            MAX_THREADS = self.app_config.max_threads
            CHUNK_SIZE = self.app_config.chunk_size
            
            # Appliquer les paramètres de sécurité
            if hasattr(self, 'verification_enabled'):
                self.verification_enabled = self.app_config.enable_validation
            
            # Configurer le niveau de logging
            logger = logging.getLogger()
            level = getattr(logging, self.app_config.logging_level, logging.INFO)
            logger.setLevel(level)
            
            logging.info(f"Configuration appliquée - Performance: {MAX_THREADS} threads, chunks de {CHUNK_SIZE}")
            
        except Exception as e:
            logging.error(f"Erreur lors de l'application de la configuration: {e}")
    
    def check_recovery_needed(self):
        """Vérifie s'il y a une session à récupérer"""
        if not self.auto_save_manager:
            return
        
        try:
            latest_backup = self.auto_save_manager.load_latest_backup()
            if latest_backup:
                # Vérifier si la sauvegarde est récente (moins de 24h)
                backup_time = datetime.fromisoformat(latest_backup['timestamp'])
                time_diff = datetime.now() - backup_time
                
                if time_diff.total_seconds() < 24 * 3600:  # 24 heures
                    # Proposer la récupération à l'utilisateur
                    QTimer.singleShot(2000, lambda: self.offer_session_recovery(latest_backup))
                    
        except Exception as e:
            logging.error(f"Erreur lors de la vérification de récupération: {e}")
    
    def offer_session_recovery(self, backup_data: Dict[str, Any]):
        """Propose la récupération de session à l'utilisateur"""
        try:
            backup_time = backup_data['timestamp']
            session_id = backup_data.get('session_id', 'Unknown')
            
            msg = QMessageBox(self)
            msg.setWindowTitle("Récupération de session")
            msg.setIcon(QMessageBox.Icon.Question)
            msg.setText("Une session précédente a été détectée.")
            msg.setInformativeText(f"Session du {backup_time}\nID: {session_id}\n\nVoulez-vous récupérer cette session ?")
            
            recover_button = msg.addButton("Récupérer", QMessageBox.ButtonRole.AcceptRole)
            ignore_button = msg.addButton("Ignorer", QMessageBox.ButtonRole.RejectRole)
            msg.setDefaultButton(recover_button)
            
            result = msg.exec()
            
            if msg.clickedButton() == recover_button:
                self.restore_session_from_backup(backup_data)
            else:
                logging.info("Récupération de session ignorée par l'utilisateur")
                
        except Exception as e:
            logging.error(f"Erreur lors de l'offre de récupération: {e}")
    
    def restore_session_from_backup(self, backup_data: Dict[str, Any]):
        """Restaure la session depuis les données de sauvegarde"""
        try:
            app_state = backup_data.get('application_state', {})
            
            # Restaurer le répertoire de travail
            if 'directory' in app_state and os.path.exists(app_state['directory']):
                self.directory = app_state['directory']
                self.label_directory.setText(self.directory)
                
            # Restaurer les paramètres de compilation
            if 'compilation_settings' in app_state:
                settings = app_state['compilation_settings']
                
                if 'header_start_row' in settings:
                    self.spinbox_header_start.setValue(max(1, settings['header_start_row']))
                if 'header_rows' in settings:
                    self.spinbox_header.setValue(max(1, settings['header_rows']))
                if 'date_format' in settings:
                    self.date_format = settings['date_format']
                if 'filename_option' in settings:
                    self.filename_option = settings['filename_option']
                if 'include_preliminary' in settings:
                    self.include_preliminary = settings['include_preliminary']
                    self.checkbox_preliminary.setChecked(self.include_preliminary)
                    self.toggle_preliminary_options(Qt.CheckState.Checked.value if self.include_preliminary else Qt.CheckState.Unchecked.value)
            
            # Recharger les fichiers si le répertoire existe
            if self.directory and os.path.exists(self.directory):
                self.load_files()
            
            # Afficher un message de confirmation
            self.status_label.setText("Session restaurée avec succès")
            
            logging.info("Session restaurée depuis la sauvegarde")
            
        except Exception as e:
            logging.error(f"Erreur lors de la restauration de session: {e}")
            QMessageBox.warning(
                self,
                "Erreur de récupération",
                f"Impossible de restaurer la session:\n{e}"
            )
    
    def save_current_session(self):
        """Sauvegarde l'état actuel de la session"""
        if not self.auto_save_manager:
            return False
        
        try:
            session_data = {
                'directory': self.directory,
                'compilation_settings': {
                    'header_start_row': self.spinbox_header_start.value() if hasattr(self, 'spinbox_header_start') else 1,
                    'header_rows': self.spinbox_header.value() if hasattr(self, 'spinbox_header') else 1,
                    'date_format': self.date_format,
                    'filename_option': self.filename_option,
                    'verification_enabled': self.verification_enabled,
                    'include_preliminary': self.include_preliminary
                },
                'window_state': {
                    'geometry': self.geometry().getRect(),
                    'maximized': self.isMaximized()
                }
            }
            
            return self.auto_save_manager.save_session_state(session_data)
            
        except Exception as e:
            logging.error(f"Erreur lors de la sauvegarde de session: {e}")
            return False

    def setup_ui(self):
        """Configure l'interface utilisateur principale avec responsivité."""
        self.setWindowTitle(self.translate("app_title"))
        
        # Calculer la taille responsive
        responsive_width, responsive_height = self.responsive_manager.get_responsive_size(1200, 700)
        
        # Calculer la position centrée
        if self.responsive_manager.screen_size:
            screen_width = self.responsive_manager.screen_size['width']
            screen_height = self.responsive_manager.screen_size['height']
            x = (screen_width - responsive_width) // 2
            y = (screen_height - responsive_height) // 2
        else:
            x, y = 50, 50
            
        self.setGeometry(x, y, responsive_width, responsive_height)
        self.setWindowIcon(QIcon(resource_path("icon.ico")))
        
        # Taille minimale responsive
        min_width = min(800, self.responsive_manager.screen_size['width'] - 100) if self.responsive_manager.screen_size else 800
        min_height = min(600, self.responsive_manager.screen_size['height'] - 100) if self.responsive_manager.screen_size else 600
        self.setMinimumSize(min_width, min_height)
        
        self.apply_responsive_stylesheet()
    
    def apply_responsive_stylesheet(self):
        """Applique le style CSS responsive à l'application."""
        # Calculer les tailles responsives
        font_size_normal = self.responsive_manager.get_responsive_font_size(FONT_SIZES["NORMAL"])
        font_size_large = self.responsive_manager.get_responsive_font_size(FONT_SIZES["LARGE"])
        font_size_header = self.responsive_manager.get_responsive_font_size(FONT_SIZES["HEADER"])
        margin_normal = self.responsive_manager.get_responsive_margin(5)
        margin_large = self.responsive_manager.get_responsive_margin(10)
        icon_size = self.responsive_manager.get_responsive_icon_size(32)
        
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: #{COLORS["BACKGROUND"]};
                font-size: {font_size_normal}px;
            }}
            
            /* Style pour les labels */
            QLabel {{
                color: #{COLORS["DARK_TEXT"]};
                font-weight: bold;
                font-size: {font_size_normal}px;
            }}
            
            /* Style pour les champs de texte */
            QLineEdit, QTextEdit {{
                background-color: white;
                border: 1px solid #{COLORS["PRIMARY_LIGHT"]};
                border-radius: 4px;
                padding: {margin_normal}px;
                font-size: {font_size_normal}px;
                min-height: {font_size_normal + margin_normal * 2}px;
            }}
            QLineEdit:focus, QTextEdit:focus {{
                border-color: #{COLORS["PRIMARY"]};
                border: 2px solid #{COLORS["PRIMARY"]};
            }}
            
            /* Style pour les boutons */
            QPushButton {{
                background-color: #{COLORS["PRIMARY"]};
                color: #{COLORS["LIGHT_TEXT"]};
                border: none;
                border-radius: 4px;
                padding: {margin_normal}px {margin_large}px;
                font-weight: bold;
                font-size: {font_size_normal}px;
                min-height: {font_size_normal + margin_normal * 4}px;
            }}
            QPushButton:hover {{
                background-color: #{COLORS["PRIMARY_DARK"]};
            }}
            QPushButton:pressed {{
                background-color: #{COLORS["PRIMARY_LIGHT"]};
            }}
            QPushButton:disabled {{
                background-color: #cccccc;
                color: #666666;
            }}
            
            /* Style pour les checkbox */
            QCheckBox {{
                color: #{COLORS["DARK_TEXT"]};
                font-size: {font_size_normal}px;
                spacing: {margin_normal}px;
            }}
            QCheckBox::indicator {{
                width: {font_size_normal + 4}px;
                height: {font_size_normal + 4}px;
            }}
            
            /* Style pour les combobox */
            QComboBox {{
                background-color: white;
                border: 1px solid #{COLORS["PRIMARY_LIGHT"]};
                border-radius: 4px;
                padding: {margin_normal}px;
                font-size: {font_size_normal}px;
                min-height: {font_size_normal + margin_normal * 2}px;
                min-width: {150 * self.responsive_manager.scale_factor}px;
            }}
            
            /* Style pour les spinbox */
            QSpinBox {{
                background-color: white;
                border: 1px solid #{COLORS["PRIMARY_LIGHT"]};
                border-radius: 4px;
                padding: {margin_normal}px;
                font-size: {font_size_normal}px;
                min-height: {font_size_normal + margin_normal * 2}px;
            }}
            
            /* Style pour les onglets */
            QTabWidget::pane {{
                border: 1px solid #{COLORS["BORDER"]};
                background-color: white;
            }}
            QTabBar::tab {{
                background-color: #{COLORS["ACCENT"]};
                color: #{COLORS["DARK_TEXT"]};
                padding: {margin_normal}px {margin_large}px;
                font-size: {font_size_normal}px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                margin-right: 2px;
            }}
            QTabBar::tab:selected {{
                background-color: #{COLORS["PRIMARY"]};
                color: #{COLORS["LIGHT_TEXT"]};
            }}
            
            /* Style pour les groupes */
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #{COLORS["PRIMARY_LIGHT"]};
                border-radius: 8px;
                margin: {margin_normal}px 0;
                padding-top: {font_size_large}px;
                font-size: {font_size_normal}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {margin_large}px;
                padding: 0 {margin_normal}px 0 {margin_normal}px;
                color: #{COLORS["PRIMARY"]};
                font-size: {font_size_large}px;
            }}
            
            /* Style pour les listes */
            QListWidget {{
                background-color: white;
                border: 1px solid #{COLORS["BORDER"]};
                border-radius: 4px;
                font-size: {font_size_normal}px;
            }}
            QListWidget::item {{
                padding: {margin_normal}px;
                border-bottom: 1px solid #{COLORS["BORDER"]};
            }}
            QListWidget::item:selected {{
                background-color: #{COLORS["PRIMARY_LIGHT"]};
                color: #{COLORS["LIGHT_TEXT"]};
            }}
            
            /* Style pour les tableaux */
            QTableWidget {{
                background-color: white;
                border: 1px solid #{COLORS["BORDER"]};
                border-radius: 4px;
                font-size: {font_size_normal}px;
                gridline-color: #{COLORS["BORDER"]};
            }}
            QTableWidget::item {{
                padding: {margin_normal}px;
                border: none;
            }}
            QTableWidget::item:selected {{
                background-color: #{COLORS["PRIMARY_LIGHT"]};
                color: #{COLORS["LIGHT_TEXT"]};
            }}
            QHeaderView::section {{
                background-color: #{COLORS["PRIMARY"]};
                color: #{COLORS["LIGHT_TEXT"]};
                padding: {margin_normal}px;
                border: none;
                font-weight: bold;
                font-size: {font_size_normal}px;
            }}
            
            /* Style pour la barre de progression */
            QProgressBar {{
                border: 2px solid #{COLORS["BORDER"]};
                border-radius: 4px;
                text-align: center;
                font-size: {font_size_normal}px;
                min-height: {font_size_normal + margin_normal * 4}px;
            }}
            QProgressBar::chunk {{
                background-color: #{COLORS["SUCCESS"]};
                border-radius: 2px;
            }}
            
            /* Style pour la barre d'outils */
            QToolBar {{
                background-color: #{COLORS["PRIMARY"]};
                border: none;
                spacing: {margin_normal}px;
                font-size: {font_size_normal}px;
            }}
            QToolBar QToolButton {{
                background-color: transparent;
                color: #{COLORS["LIGHT_TEXT"]};
                border: none;
                padding: {margin_normal}px;
                border-radius: 4px;
                font-size: {font_size_normal}px;
            }}
            QToolBar QToolButton:hover {{
                background-color: #{COLORS["PRIMARY_DARK"]};
            }}
            
            /* Style pour les menus */
            QMenuBar {{
                background-color: #{COLORS["PRIMARY"]};
                color: #{COLORS["LIGHT_TEXT"]};
                font-size: {font_size_normal}px;
            }}
            QMenuBar::item {{
                padding: {margin_normal}px {margin_large}px;
                background-color: transparent;
            }}
            QMenuBar::item:selected {{
                background-color: #{COLORS["PRIMARY_DARK"]};
            }}
            QMenu {{
                background-color: white;
                border: 1px solid #{COLORS["BORDER"]};
                font-size: {font_size_normal}px;
            }}
            QMenu::item {{
                padding: {margin_normal}px {margin_large * 2}px;
            }}
            QMenu::item:selected {{
                background-color: #{COLORS["PRIMARY_LIGHT"]};
                color: #{COLORS["LIGHT_TEXT"]};
            }}
            
            /* Style responsive pour petits écrans */
            """ + (f"""
            QTabWidget QWidget {{
                padding: {margin_normal // 2}px;
            }}
            QGroupBox {{
                margin: {margin_normal // 2}px 0;
            }}
            """ if self.responsive_manager.current_breakpoint == 'small' else "") + """
        """)
    
    def apply_stylesheet(self):
        """Applique le style CSS à l'application."""
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: #{COLORS["BACKGROUND"]};
            }}
            
            /* Style pour les labels */
            QLabel {{
                color: #{COLORS["DARK_TEXT"]};
                font-weight: bold;
            }}
            
            /* Style pour les champs de texte */
            QLineEdit, QTextEdit {{
                background-color: white;
                border: 1px solid #{COLORS["PRIMARY_LIGHT"]};
                border-radius: 4px;
                padding: 5px;
            }}
            QLineEdit:focus, QTextEdit:focus {{
                border-color: #{COLORS["PRIMARY"]};
            }}
            
            /* Style pour les combobox */
            /* Style simple uniforme */
            QComboBox {{
                background-color: white;
                border: 1px solid #{COLORS["PRIMARY_LIGHT"]};
                border-radius: 4px;
                padding: 5px;
                font-weight: bold;
                color: #{COLORS["DARK_TEXT"]};
            }}

            QComboBox:hover {{
                border-color: #{COLORS["PRIMARY_DARK"]};
            }}

            QComboBox:focus {{
                border-color: #{COLORS["PRIMARY"]};
            }}

            /* Forcer le style du popup/menu déroulant */
            QComboBox QAbstractItemView {{
                background-color: white !important;
                border: 2px solid #{COLORS["PRIMARY"]} !important;
                border-radius: 4px;
                font-weight: bold;
                outline: none;
                selection-background-color: #{COLORS["SUCCESS"]} !important;
                selection-color: white !important;
            }}

            /* Style des items avec !important pour forcer */
            QComboBox QAbstractItemView::item {{
                padding: 8px 12px;
                border-bottom: 1px solid #{COLORS["PRIMARY_LIGHT"]};
                color: black !important;
                background-color: white !important;
                min-height: 20px;
            }}

            /* Hover avec couleurs forcées */
            QComboBox QAbstractItemView::item:hover {{
                background-color: #{COLORS["ACCENT"]} !important;
                color: black !important;
            }}

            /* Sélection avec couleurs forcées */
            QComboBox QAbstractItemView::item:selected {{
                background-color: #{COLORS["SUCCESS"]} !important;
                color: white !important;
            }}

            /* AJOUT : Style spécifique pour QListView (certains combos utilisent ça) */
            QComboBox QListView {{
                background-color: white !important;
                border: 2px solid #{COLORS["PRIMARY"]} !important;
                selection-background-color: #{COLORS["SUCCESS"]} !important;
            }}

            QComboBox QListView::item {{
                color: black !important;
                background-color: white !important;
                padding: 8px 12px;
                border-bottom: 1px solid #{COLORS["PRIMARY_LIGHT"]};
            }}

            QComboBox QListView::item:hover {{
                background-color: #{COLORS["ACCENT"]} !important;
                color: black !important;
            }}

            QComboBox QListView::item:selected {{
                background-color: #{COLORS["SUCCESS"]} !important;
                color: white !important;
            }}

            /* AJOUT : Pour forcer sur tous les types de popup */
            QComboBox * {{
                selection-background-color: #{COLORS["SUCCESS"]} !important;
                selection-color: white !important;
            }}

            /* Style pour la liste des fichiers */
            QListWidget::item {{
                background-color: white;
                border: 1px solid #{COLORS["PRIMARY_LIGHT"]};
                border-radius: 3px;
                margin: 1px;
                padding: 2px 4px;
                min-height: 10px;
            }}

            QListWidget::item:focus {{
                border: 2px solid #{COLORS["PRIMARY"]};
                outline: none;
            }}

            QListWidget::item:selected {{
                background-color: #{COLORS["ACCENT"]};
                border: 2px solid #{COLORS["SUCCESS"]};
                color: #{COLORS["DARK_TEXT"]};
                font-weight: bold;
            }}

            QListWidget::item:selected:hover {{
                background-color: #{COLORS["PRIMARY_LIGHT"]};
                color: white;
            }}

            QListWidget::item:!selected {{
                background-color: white;
                color: #{COLORS["DARK_TEXT"]};
            }}

            QListWidget::item:!selected:hover {{
                background-color: #{COLORS["PRIMARY_LIGHT"]};
                color: #{COLORS["DARK_TEXT"]};
            }}

            QListWidget::item:!selected:pressed {{
                background-color: #{COLORS["PRIMARY_DARK"]};
                color: white;
            }}

            /* Style pour les tableaux */
            QTableWidget {{
                border: 1px solid #{COLORS["PRIMARY"]};
                gridline-color: #{COLORS["PRIMARY"]};
                background-color: white;
            }}

            QTableWidget::item {{
                border-bottom: 1px solid #{COLORS["PRIMARY_LIGHT"]};
            }}

            QTableWidget::item:selected {{
                background-color: #{COLORS["PRIMARY"]};
                color: white;
            }}

            QTableWidget::item:hover {{
                background-color: #{COLORS["PRIMARY_LIGHT"]};
                color: #{COLORS["DARK_TEXT"]};
            }}

            QTableWidget QHeaderView::section {{
                background-color: #{COLORS["PRIMARY"]};
                color: white;
                padding: 5px;
                border: 1px solid white;
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

            QPushButton:pressed {{
                background-color: #{COLORS["PRIMARY_LIGHT"]};
                color: #{COLORS["DARK_TEXT"]};
            }}
            
            QPushButton:disabled {{
                background-color: #a0a0a0;
                color: #d0d0d0;
            }}

            /* Style pour les barres de progression */
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

            /* Style pour les barres de défilement */
            QScrollBar:vertical {{
                border: 1px solid #{COLORS["PRIMARY_LIGHT"]};
                background-color: #{COLORS["ACCENT"]};
                width: 18px;
                border-radius: 9px;
                margin: 0px;
            }}

            QScrollBar::handle:vertical {{
                background-color: #{COLORS["PRIMARY"]};
                border-radius: 8px;
                min-height: 30px;
                margin: 2px;
            }}

            QScrollBar::handle:vertical:hover {{
                background-color: #{COLORS["PRIMARY_DARK"]};
            }}

            QScrollBar::handle:vertical:pressed {{
                background-color: #{COLORS["PRIMARY_LIGHT"]};
            }}

            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0px;
            }}

            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
                background: transparent;
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

            /* Style pour la barre d'outils */
            QToolBar {{
                background-color: #{COLORS["PRIMARY"]};
                border: none;
                spacing: 10px;
                padding: 5px;
            }}
            
            QToolButton {{
                background-color: transparent;
                color: white;
                border-radius: 4px;
                padding: 5px;
            }}
            
            QToolButton:hover {{
                background-color: #{COLORS["PRIMARY_DARK"]};
            }}
            
            /* Style pour les onglets */
            QTabWidget::pane {{
                border: 1px solid #{COLORS["PRIMARY"]};
                border-radius: 4px;
                background-color: white;
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
            
            /* Style pour les checkboxes */
            /* Style pour tous les checkboxes - texte en gras */
            QCheckBox {{
                font-weight: bold;
                color: #{COLORS["DARK_TEXT"]};
            }}
            
            QCheckBox::indicator {{
                width: 18px;
                height: 18px;
                border: 2px solid #{COLORS["PRIMARY"]};
                border-radius: 3px;
                background-color: white;
            }}

            QCheckBox::indicator:checked {{
                background-color: #{COLORS["SUCCESS"]};
                border: 2px solid #{COLORS["SUCCESS"]};
                image: none;
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
    
    def create_menu(self):
        """Crée le menu de l'application."""
        menubar = self.menuBar()
        
        # Menu Fichier
        file_menu = menubar.addMenu(self.translate("file_menu"))
        
        # Action Ouvrir un répertoire
        open_action = QAction(self.translate("open_dir"), self)
        open_action.setShortcut(QKeySequence.StandardKey.Open)
        open_action.triggered.connect(self.choose_directory)
        file_menu.addAction(open_action)
        
        file_menu.addSeparator()
        
        # Action Enregistrer les paramètres
        save_settings_action = QAction(self.translate("save_settings"), self)
        save_settings_action.triggered.connect(self.save_settings)
        file_menu.addAction(save_settings_action)
        
        # Action Charger les paramètres
        load_settings_action = QAction(self.translate("load_settings"), self)
        load_settings_action.triggered.connect(self.load_settings)
        file_menu.addAction(load_settings_action)
        
        file_menu.addSeparator()
        
        # Action Quitter
        exit_action = QAction(self.translate("exit"), self)
        exit_action.setShortcut(QKeySequence.StandardKey.Quit)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Menu Édition
        edit_menu = menubar.addMenu(self.translate("edit_menu"))
        
        # Action Sélectionner tout
        select_all_action = QAction(self.translate("select_all"), self)
        select_all_action.triggered.connect(self.select_all_files)
        edit_menu.addAction(select_all_action)
        
        # Action Désélectionner tout
        deselect_all_action = QAction(self.translate("deselect_all"), self)
        deselect_all_action.triggered.connect(self.deselect_all_files)
        edit_menu.addAction(deselect_all_action)
        
        # Action Inverser la sélection
        invert_selection_action = QAction(self.translate("invert_selection"), self)
        invert_selection_action.triggered.connect(self.invert_file_selection)
        edit_menu.addAction(invert_selection_action)
        
        # Menu Outils
        tools_menu = menubar.addMenu(self.translate("tools_menu"))
        
        # Action Aperçu des données
        preview_action = QAction(self.translate("preview_tool"), self)
        preview_action.triggered.connect(self.show_preview)
        tools_menu.addAction(preview_action)
        
        # Action Format de date
        date_format_action = QAction(self.translate("date_format"), self)
        date_format_action.triggered.connect(self.show_date_format_dialog)
        tools_menu.addAction(date_format_action)
        
        # Menu Langue
        language_menu = menubar.addMenu(self.translate("language_menu"))
        
        # Action Changer la langue
        change_language_action = QAction(self.translate("language"), self)
        change_language_action.triggered.connect(self.show_language_dialog)
        language_menu.addAction(change_language_action)
        
        # Menu Aide
        about_menu = menubar.addMenu("📚 Aide && Support")
        
        # Action Assistant tutoriel
        tutorial_action = QAction("🎯 Assistant de démarrage", self)
        tutorial_action.triggered.connect(self.show_tutorial)
        about_menu.addAction(tutorial_action)
        
        about_menu.addSeparator()
        
        # Action Guide utilisateur
        user_guide_action = QAction("📖 Guide utilisateur", self)
        user_guide_action.triggered.connect(self.show_user_guide)
        about_menu.addAction(user_guide_action)
        
        about_menu.addSeparator()
        
        # Action Aide
        about_action = QAction(self.translate("about"), self)
        about_action.triggered.connect(self.show_about_dialog)
        about_menu.addAction(about_action)
    
    def create_toolbar(self):
        """Crée la barre d'outils de l'application."""
        toolbar = QToolBar(self)
        toolbar.setMovable(False)
        toolbar.setIconSize(QSize(32, 32))
        
        # Action Ouvrir un répertoire
        open_action = QAction(self.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon), self.translate("open_dir"), self)
        open_action.triggered.connect(self.choose_directory)
        toolbar.addAction(open_action)
        
        toolbar.addSeparator()
        
        # Action Aperçu des données
        preview_action = QAction(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogContentsView), self.translate("preview_tool"), self)
        preview_action.triggered.connect(self.show_preview)
        toolbar.addAction(preview_action)
        
        # Action Compiler
        compile_action = QAction(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay), self.translate("compile_tool"), self)
        compile_action.triggered.connect(self.compile_files)
        toolbar.addAction(compile_action)
        
        self.addToolBar(toolbar)
    
    def create_main_layout(self):
        """Crée la mise en page principale de l'application."""
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        
        # Création des onglets
        self.tabs = QTabWidget()
        self.create_compilation_tab()
        self.create_advanced_options_tab()
        self.create_date_format_tab()
        self.create_preview_tab()
        self.create_help_tab()
        self.create_about_tab()
        
        # Appliquer un style spécifique à la barre d'onglets
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
    

    def create_compilation_tab(self):
        """Crée l'onglet principal de compilation avec layout responsive."""
        tab = QWidget()
        layout = QVBoxLayout()
        
        # ============ ZONE PRINCIPALE (flexible) ============
        main_content_widget = QWidget()
        main_content_layout = QVBoxLayout(main_content_widget)
        main_content_layout.setContentsMargins(0, 0, 0, 0)
        
        # Groupe sélection des fichiers
        files_group = self.create_files_group()
        main_content_layout.addWidget(files_group, 2)  # (plus grand)
        
        # Groupe options de compilation
        self.options_group = self.create_options_group()
        main_content_layout.addWidget(self.options_group, 1)  # (plus petit)
        
        # Ajouter la zone principale au layout avec stretch
        layout.addWidget(main_content_widget, 1)  # Zone flexible
        
        # ============ ZONE COMPILATION (fixe en bas) ============
        compile_zone = QWidget()
        compile_zone_layout = QVBoxLayout(compile_zone)
        compile_zone_layout.setContentsMargins(5, 5, 5, 5)
        compile_zone_layout.setSpacing(5)
        
        # Widget de progression avec annulation (dans zone dédiée)
        self.progress_widget = CancellableProgressWidget()
        self.progress_widget.setVisible(False)
        self.progress_widget.cancel_requested.connect(self.cancel_compilation)
        compile_zone_layout.addWidget(self.progress_widget)
        
        # Bouton de compilation
        compile_layout = QHBoxLayout()
        self.button_compile = QPushButton(self.translate("start_compilation"))
        self.button_compile.setFont(QFont("Segoe UI", FONT_SIZES["LARGE"]))
        self.button_compile.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay))
        self.button_compile.setStyleSheet(f"background-color: #{COLORS['PRIMARY']}; color: white;")
        compile_layout.addStretch()
        compile_layout.addWidget(self.button_compile)
        compile_layout.addStretch()
        compile_zone_layout.addLayout(compile_layout)
        
        # Ajouter la zone de compilation (taille fixe)
        layout.addWidget(compile_zone, 0)  # Pas de stretch = taille fixe
        
        # Status label
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status_label)
        
        tab.setLayout(layout)
        self.tabs.addTab(tab, self.translate("compilation_options"))
        
        # Initialiser les références pour le mode compact
        self.main_content_widget = main_content_widget
        self.compile_zone = compile_zone
        self.is_compact_mode = False
    
    def enter_compact_mode(self):
        """Bascule vers le mode compact pendant la compilation"""
        if self.is_compact_mode:
            return
            
        self.is_compact_mode = True
        
        # Sauvegarder l'état original de la zone d'options avant de la masquer
        if hasattr(self, 'options_group') and self.options_group:
            self._options_group_original_visible = self.options_group.isVisible()
            # Masquer complètement la zone d'options pendant la compilation
            self.options_group.setVisible(False)
        
        # S'assurer que la zone de progression est visible
        self.progress_widget.setVisible(True)
        
        # Log du passage en mode compact
        screen_info = self.responsive_manager.detect_screen_size(self)
        if hasattr(self, 'structured_logger') and self.structured_logger:
            self.structured_logger.log_structured("INFO", "compact_mode_enabled",
                                                breakpoint=screen_info['breakpoint'])
    
    def exit_compact_mode(self):
        """Quitte le mode compact après la compilation"""
        if not self.is_compact_mode:
            return
            
        self.is_compact_mode = False
        
        # Restaurer la zone d'options à son état original
        if hasattr(self, 'options_group') and self.options_group:
            if hasattr(self, '_options_group_original_visible'):
                self.options_group.setVisible(self._options_group_original_visible)
                delattr(self, '_options_group_original_visible')
            else:
                # Par défaut, la rendre visible
                self.options_group.setVisible(True)
        
        # Masquer la barre de progression
        self.progress_widget.setVisible(False)
        
        # Log de la sortie du mode compact
        if hasattr(self, 'structured_logger') and self.structured_logger:
            self.structured_logger.log_structured("INFO", "compact_mode_disabled")

    def create_files_group(self):
        """Crée le groupe d'éléments pour la sélection des fichiers."""
        group = QGroupBox(self.translate("file_selection"))
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
        self.label_directory = QLabel(self.translate("no_directory"))
        self.button_choose_directory = QPushButton(self.translate("choose_directory"))
        self.button_choose_directory.setFont(QFont("Segoe UI", FONT_SIZES["NORMAL"]))
        self.button_choose_directory.setIcon(QIcon(resource_path("folder.ico")))
        self.button_choose_directory.setStyleSheet(f"background-color: #{COLORS['PRIMARY']}; color: white;")
        dir_layout.addWidget(self.label_directory)
        dir_layout.addWidget(self.button_choose_directory)
        layout.addLayout(dir_layout)
        
        # Liste des fichiers
        self.list_files = QListWidget()
        self.list_files.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        
        # Options de sélection
        selection_layout = QHBoxLayout()
        self.checkbox_all_files = QCheckBox(self.translate("select_all_files"))
        self.label_file_count = QLabel(self.translate("files_selected", 0))
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
        group = QGroupBox(self.translate("compilation_options"))
        grid_layout = QGridLayout()
        
        # ===== LIGNE 1 : Structure des données =====
        
        # Colonne 1 - Ligne de début des en-têtes
        self.label_header_start = QLabel(self.translate("header_start_row"))
        self.spinbox_header_start = QSpinBox()
        self.spinbox_header_start.setMinimum(1)
        self.spinbox_header_start.setMaximum(20)
        self.spinbox_header_start.setValue(1)
        header_start_container = QHBoxLayout()
        header_start_container.addWidget(self.label_header_start)
        header_start_container.addWidget(self.spinbox_header_start)
        header_start_container.addStretch()
        
        # Colonne 2 - Nombre de lignes d'en-tête
        self.label_header = QLabel(self.translate("header_rows"))
        self.spinbox_header = QSpinBox()
        self.spinbox_header.setMinimum(1)
        self.spinbox_header.setMaximum(15)
        self.spinbox_header.setValue(1)
        header_container = QHBoxLayout()
        header_container.addWidget(self.label_header)
        header_container.addWidget(self.spinbox_header)
        header_container.addStretch()
        
        # Colonne 3 - Répéter les en-têtes
        self.checkbox_repeat_header = QCheckBox(self.translate("repeat_headers"))
        
        # ===== LIGNE 2 : Options de traitement =====
        
        # Colonne 1 - Inclure les informations de début de fichier
        self.checkbox_preliminary = QCheckBox(self.translate("include_preliminary"))
        self.checkbox_preliminary.setChecked(False)
        
        # Colonne 2 - Fichier source (étendu sur toute la largeur de la colonne)
        self.label_preliminary_source = QLabel(self.translate("preliminary_source_file"))
        self.combo_preliminary_source = QComboBox()
        self.combo_preliminary_source.setEnabled(False)  # Désactivé par défaut
        self.combo_preliminary_source.setMinimumWidth(150)  # RÉDUIT : de 200 à 150
    

        preliminary_source_container = QHBoxLayout()
        preliminary_source_container.addWidget(self.label_preliminary_source)
        preliminary_source_container.addWidget(self.combo_preliminary_source)
        preliminary_source_container.addStretch() # pour aligner avec les autres containers
        
        # Colonne 3 - Fusionner les en-têtes multi-niveaux
        self.checkbox_merge_headers = QCheckBox(self.translate("merge_headers"))
        
        # ===== LIGNE 3 : Sortie et finalisation =====
        
        # Colonne 1 - Option nom de fichier
        self.label_filename_option = QLabel(self.translate("filename_option"))
        self.combo_filename_option = QComboBox()
        self.combo_filename_option.addItem(self.translate("filename_none"), "none")
        self.combo_filename_option.addItem(self.translate("filename_with_extension"), "with_extension") 
        self.combo_filename_option.addItem(self.translate("filename_without_extension"), "without_extension")
        self.combo_filename_option.setCurrentIndex(0)  # Par défaut "none"

        filename_container = QHBoxLayout()
        filename_container.addWidget(self.label_filename_option)
        filename_container.addWidget(self.combo_filename_option)
        filename_container.addStretch()
        
        # Colonne 2 - Nom du fichier de sortie
        self.label_output_name = QLabel(self.translate("output_filename"))
        self.lineedit_output_name = QLineEdit("compilation.xlsx")
        
        # Container horizontal pour le label et le champ de saisie (ligne 3)
        output_container = QHBoxLayout()
        output_container.addWidget(self.label_output_name)
        output_container.addWidget(self.lineedit_output_name)
        output_container.addStretch()
        
        # Créer l'indicateur de validation pour le nom de fichier de sortie (ligne 4)
        self.output_indicator = ValidationIndicator()
        
        # Colonne 3 - Activer la vérification préliminaire
        self.checkbox_verify_files = QCheckBox(self.translate("enable_verification"))
        self.checkbox_verify_files.setChecked(True)
        
        # ===== PLACEMENT DANS LA GRILLE (3 lignes x 3 colonnes) =====
        
        # Ligne 1 : Structure des données
        grid_layout.addLayout(header_start_container, 0, 0)      # Ligne début en-têtes
        grid_layout.addLayout(header_container, 0, 1)            # Nombre lignes en-tête
        grid_layout.addWidget(self.checkbox_repeat_header, 0, 2) # Répéter en-têtes
        
        # Ligne 2 : Options de traitement
        grid_layout.addWidget(self.checkbox_preliminary, 1, 0)       # Inclure infos début
        grid_layout.addLayout(preliminary_source_container, 1, 1)    # Fichier source (étendu)
        grid_layout.addWidget(self.checkbox_merge_headers, 1, 2)     # Fusionner en-têtes
        
        # Ligne 3 : Sortie et finalisation
        grid_layout.addLayout(filename_container, 2, 0)          # Option nom fichier
        grid_layout.addLayout(output_container, 2, 1)            # Nom fichier sortie
        grid_layout.addWidget(self.checkbox_verify_files, 2, 2)  # Vérification préliminaire
        
        # Ligne 4 : Indicateur de validation (aligné sous "Nom du fichier de sortie")
        grid_layout.addWidget(self.output_indicator, 3, 1)       # Validation nom fichier sortie
        
        # ===== CONFIGURATION DU LAYOUT (identique à l'existant) =====
        
        # Définir l'espacement et les marges (garder les mêmes valeurs)
        grid_layout.setSpacing(20)
        grid_layout.setContentsMargins(20, 20, 20, 20)
        
        # Définir les colonnes pour qu'elles aient la même largeur
        grid_layout.setColumnStretch(0, 1)
        grid_layout.setColumnStretch(1, 1)
        grid_layout.setColumnStretch(2, 1)
        
        # Alignement vertical des éléments
        grid_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        
        group.setLayout(grid_layout)
        return group
    
    def create_advanced_options_tab(self):
        """Crée l'onglet des options avancées."""
        tab = QWidget()
        layout = QVBoxLayout()
        
        # Groupe traitement des données
        data_group = QGroupBox(self.translate("compilation_options"))
        data_layout = QVBoxLayout()
        
        self.checkbox_remove_duplicates = QCheckBox(self.translate("remove_duplicates"))
        self.checkbox_remove_duplicates.setChecked(False)
        
        self.checkbox_remove_empty_rows = QCheckBox(self.translate("remove_empty_rows"))
        self.checkbox_remove_empty_rows.setChecked(False)
        
        sort_layout = QVBoxLayout()
        self.checkbox_sort_data = QCheckBox(self.translate("sort_data"))
        
        sort_options = QHBoxLayout()
        self.label_sort_column = QLabel(self.translate("sort_column"))
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
        format_group = QGroupBox(self.translate("compilation_options"))
        format_layout = QVBoxLayout()
        
        self.checkbox_auto_width = QCheckBox(self.translate("auto_width"))
        self.checkbox_auto_width.setChecked(True)
        
        self.checkbox_freeze_header = QCheckBox(self.translate("freeze_headers"))
        self.checkbox_freeze_header.setChecked(False)
        
        format_layout.addWidget(self.checkbox_auto_width)
        format_layout.addWidget(self.checkbox_freeze_header)
        format_group.setLayout(format_layout)
        
        # Groupe formats des fichiers
        format_files_group = QGroupBox(self.translate("file_formats"))
        format_files_layout = QVBoxLayout()
        
        self.checkbox_excel = QCheckBox(self.translate("excel_files"))
        self.checkbox_excel.setChecked(True)
        self.checkbox_excel.setEnabled(False)
        
        self.checkbox_csv = QCheckBox(self.translate("text_files"))
        self.checkbox_csv.setChecked(True)
        
        format_files_layout.addWidget(self.checkbox_excel)
        format_files_layout.addWidget(self.checkbox_csv)
        format_files_group.setLayout(format_files_layout)
        
        layout.addWidget(data_group)
        layout.addWidget(format_group)
        layout.addWidget(format_files_group)
        layout.addStretch()
        tab.setLayout(layout)
        self.tabs.addTab(tab, self.translate("advanced_options"))
    
    def create_date_format_tab(self):
        """Crée l'onglet de format de date."""
        tab = QWidget()
        layout = QVBoxLayout()
        
        # Groupe format de date
        group = QGroupBox(self.translate("date_format_options"))
        group_layout = QVBoxLayout()
        
        # Créer les boutons radio pour chaque format prédéfini
        self.date_radio_group = QButtonGroup(self)
        formats = [
            ("STANDARD", "date_format_standard"),
            ("FRENCH", "date_format_french"),
            ("US", "date_format_us"),
            ("DATETIME", "date_format_datetime"),
            ("DATETIME_FRENCH", "date_format_datetime_french"),
            ("DATE_ONLY", "date_format_date_only"),
            ("TIME_ONLY", "date_format_time_only"),
            ("SHORT", "date_format_short"),
            ("CUSTOM", "date_format_custom")
        ]
        
        self.date_radio_buttons = {}
        
        for i, (format_key, label_key) in enumerate(formats):
            radio = QRadioButton(self.translate(label_key))
            self.date_radio_group.addButton(radio, i)
            self.date_radio_buttons[format_key] = radio
            
            if format_key == "CUSTOM":
                custom_layout = QHBoxLayout()
                custom_layout.addWidget(radio)
                self.date_custom_edit = QLineEdit()
                self.date_custom_edit.setPlaceholderText("dd/MM/YYYY HH:mm:ss")
                self.date_custom_edit.setEnabled(False)
                custom_layout.addWidget(self.date_custom_edit)
                group_layout.addLayout(custom_layout)
            else:
                group_layout.addWidget(radio)
        
        group.setLayout(group_layout)
        layout.addWidget(group)
        
        # Exemple avec la date actuelle
        example_layout = QHBoxLayout()
        example_layout.addWidget(QLabel(self.translate("preview") + ":"))
        self.date_example_label = QLabel()
        self.update_date_example()
        example_layout.addWidget(self.date_example_label)
        layout.addLayout(example_layout)
        
        # Bouton Appliquer
        button_layout = QHBoxLayout()
        apply_button = QPushButton(self.translate("apply"))
        apply_button.clicked.connect(self.apply_date_format)
        button_layout.addStretch()
        button_layout.addWidget(apply_button)
        button_layout.addStretch()
        
        layout.addLayout(button_layout)
        layout.addStretch()
        
        tab.setLayout(layout)
        self.tabs.addTab(tab, self.translate("date_format"))
        
        # Connecter les signaux
        self.date_radio_group.buttonClicked.connect(self.on_date_format_changed)
        self.date_custom_edit.textChanged.connect(self.on_date_custom_format_changed)
        
        # Sélectionner le format actuel
        if self.date_format in self.date_radio_buttons:
            self.date_radio_buttons[self.date_format].setChecked(True)
            if self.date_format == "CUSTOM":
                self.date_custom_edit.setEnabled(True)
                self.date_custom_edit.setText(DATE_FORMATS["CUSTOM"]["format"])
    
    def create_preview_tab(self):
        """Crée l'onglet de prévisualisation des données."""
        tab = QWidget()
        layout = QVBoxLayout()
        
        # Options de prévisualisation
        options_layout = QHBoxLayout()
        
        # Sélection du fichier
        file_label = QLabel(self.translate("filename") + ":")
        self.preview_combo = QComboBox()
        self.preview_combo.currentIndexChanged.connect(self.on_preview_file_changed)
        
        # Bouton Actualiser
        self.preview_refresh_button = QPushButton(self.translate("refresh_preview"))
        self.preview_refresh_button.clicked.connect(self.refresh_preview)
        
        options_layout.addWidget(file_label)
        options_layout.addWidget(self.preview_combo, 1)
        options_layout.addWidget(self.preview_refresh_button)
        
        layout.addLayout(options_layout)
        
        # Tableau de prévisualisation
        self.preview_table = self.create_responsive_table()
        self.preview_table.setAlternatingRowColors(True)
        layout.addWidget(self.preview_table)
        
        # Informations sur la prévisualisation
        self.preview_info_label = QLabel(self.translate("preview_limited", 200))
        layout.addWidget(self.preview_info_label)
        
        tab.setLayout(layout)
        self.tabs.addTab(tab, self.translate("preview"))

    def create_help_tab(self):
        """Crée l'onglet d'aide."""
        tab = QWidget()
        layout = QVBoxLayout()
        
        help_text = """
        <style>
        body {
            font-family: "Segoe UI", sans-serif;
            margin: 0 40px;
            line-height: 1.7;
            color: #333;
        }

        h2, h3 {
            color: #2e7d32;
            font-weight: bold;
            margin: 30px 0 20px 0;
        }

        h4 {
            margin: 15px 0;
            padding-left: 60px;
            padding-right: 30px;
            font-weight: normal;
        }

        ul {
            margin-left: 80px;
            line-height: 1.5;
        }

        li {
            margin: 15px 0;
        }

        .section-content {
            padding: 20px 60px 20px 60px;
            background-color: #fafafa;
            border-radius: 8px;
            margin-bottom: 20px;
        }

        b {
            color: #1b5e20;
        }

        .highlight {
            background-color: #e8f5e9;
            padding: 2px 5px;
            border-radius: 3px;
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
            <h4>• <b>Inclure titre et métadonnées du fichier de sortie</b> : Choisit d'inclure ou non les informations situées avant l'en-tête</h4>
            <h4>&nbsp;&nbsp;&nbsp;&nbsp;- <i>Non cochée</i> : Aucune information préliminaire dans le fichier final</h4>
            <h4>&nbsp;&nbsp;&nbsp;&nbsp;- <i>Cochée</i> : Copie les lignes situées avant l'en-tête du fichier sélectionné</h4>
            <h4>&nbsp;&nbsp;&nbsp;&nbsp;<span style='background-color: #e3f2fd; padding: 2px 5px; border-radius: 3px;'><i>💡 Exemples :</i> Titre du rapport, date de génération, métadonnées</span></h4>
            <h4>• <b>Fichier source (titre et métadonnées)</b> : Sélectionnez le fichier duquel copier les informations préliminaires</h4>
            <h4>• <b>Colonne nom de fichier</b> : Choisit comment ajouter les noms des fichiers sources</h4>
            <h4>&nbsp;&nbsp;&nbsp;&nbsp;- <i>Ne pas ajouter</i> : Aucune colonne de nom de fichier</h4>
            <h4>&nbsp;&nbsp;&nbsp;&nbsp;- <i>Avec extension</i> : Ajoute une colonne avec les noms complets (ex: fichier.xlsx)</h4>
            <h4>&nbsp;&nbsp;&nbsp;&nbsp;- <i>Sans extension</i> : Ajoute une colonne avec les noms sans extension (ex: fichier)</h4>
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
            <h4>• <b>Fichiers Excel</b> : .xlsx, .xlsm (avec macros), .xltx/.xltm (modèles), .xls (format classique)</h4>
            <h4>• <b>Fichiers texte</b> : .csv, .tsv (tabulations), .txt (délimiteur auto-détecté)</h4>
        </div>
        
        <h3>IV. Format de date</h3>
        <div class="section-content">
            <h4>• Sélectionnez le format de date à utiliser pour les cellules contenant des dates</h4>
            <h4>• Formats prédéfinis disponibles: standard, français, américain, etc.</h4>
            <h4>• Option de format personnalisé pour des besoins spécifiques</h4>
            <h4>• L'aperçu montre comment la date actuelle apparaîtra dans ce format</h4>
        </div>
        
        <h3>V. Prévisualisation</h3>
        <div class="section-content">
            <h4>• Examine le contenu des fichiers avant de les compiler</h4>
            <h4>• Sélectionnez un fichier dans la liste déroulante pour le prévisualiser</h4>
            <h4>• Vérifie la structure des données, les en-têtes, etc.</h4>
            <h4>• L'aperçu est limité aux 200 premières lignes pour des raisons de performance</h4>
        </div>
        
        <h3>VI. Gestion des informations préliminaires</h3>
        <div class="section-content">
            <h4><u>Principe de fonctionnement :</u></h4>
            <h4>• Les lignes situées <b>avant l'en-tête</b> sont considérées comme des informations préliminaires</h4>
            <h4>• Ces lignes contiennent généralement le titre, la date, des métadonnées ou des informations contextuelles</h4>
            
            <h4><u>Options disponibles :</u></h4>
            <h4>• <b>Sans informations préliminaires</b> : Le fichier compilé commence directement par les en-têtes de colonnes</h4>
            <h4>• <b>Avec informations préliminaires</b> : Copie les lignes préliminaires du fichier source choisi au début du fichier compilé</h4>
            
            <h4><u>Exemple pratique :</u></h4>
            <h4><i>Fichier source :</i></h4>
            <h4>&nbsp;&nbsp;&nbsp;&nbsp;Ligne 1 : "Rapport mensuel des ventes"</h4>
            <h4>&nbsp;&nbsp;&nbsp;&nbsp;Ligne 2 : "Généré le 23/06/2025"</h4>
            <h4>&nbsp;&nbsp;&nbsp;&nbsp;Ligne 3 : "Produit | Quantité | Prix" ← En-tête (ligne de début = 3)</h4>
            <h4>&nbsp;&nbsp;&nbsp;&nbsp;Ligne 4 : "Produit A | 100 | 25€" ← Données</h4>
            
            <h4><i>Fichier compilé :</i></h4>
            <h4>&nbsp;&nbsp;&nbsp;&nbsp;✅ Lignes 1-2 : Informations préliminaires conservées</h4>
            <h4>&nbsp;&nbsp;&nbsp;&nbsp;✅ Ligne 3 : En-têtes de colonnes</h4>
            <h4>&nbsp;&nbsp;&nbsp;&nbsp;✅ Lignes suivantes : Données de tous les fichiers</h4>
        </div>
        
        <h3>VII. Internationalisation</h3>
        <div class="section-content">
            <h4>• L'application est disponible en français et en anglais</h4>
            <h4>• Changez la langue via le menu Langue</h4>
            <h4>• La langue choisie est sauvegardée dans les préférences</h4>
        </div>
        
        <h3>VIII. Vérification préliminaire des fichiers</h3>
        <div class="section-content">
            <h4>• Avant la compilation, l'application analyse tous les fichiers sélectionnés</h4>
            <h4>• Un rapport détaillé identifie les fichiers compatibles et incompatibles</h4>
            <h4>• Pour chaque fichier problématique, un motif précis est affiché</h4>
            <h4>• Vous pouvez choisir d'ignorer les fichiers incompatibles ou d'annuler la compilation</h4>
        </div>
        
        <h3>IX. Bonnes pratiques</h3>
        <div class="section-content">
            <h4>• Faites une sauvegarde de vos fichiers avant la compilation</h4>
            <h4>• Vérifiez que tous les fichiers ont une structure similaire</h4>
            <h4>• Pour les gros fichiers, traitez-les par lots</h4>
            <h4>• Utilisez des noms explicites pour les fichiers de sortie</h4>
            <h4>• Activez la vérification préliminaire pour éviter les erreurs</h4>
            <h4>• Prévisualisez les données avant compilation pour vérifier leur structure</h4>
            <h4>• Testez avec quelques fichiers avant de compiler des lots importants</h4>
        </div>
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
        self.tabs.addTab(tab, self.translate("help"))
    
    def create_about_tab(self):
        """Crée l'onglet À propos."""
        tab = QWidget()
        layout = QVBoxLayout()
        
        about_text = f"""
        <div style="text-align: center; margin: 50px 20px;">
            <h1 style="color: #{COLORS['PRIMARY']}; margin-bottom: 30px;">{self.translate("app_title")}</h1>
            <h2>Version 3.1</h2>
            <p style="font-size: 16px; margin: 30px 0;">
                Développé par:<br>
                <strong style="font-size: 20px; color: #{COLORS['PRIMARY_DARK']};">{self.translate("developer_name")}</strong><br>
                {self.translate("developer_title")}<br>
            <h2>Email : zimkada@gmail.com</h2>
            
            </p>
            <p style="margin-top: 40px; color: #555; font-style: italic;">
                Copyright © 2025. {self.translate("copyright_notice")}
            </p>
        </div>
        """
        
        label = QLabel(about_text)
        label.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(label)
        
        tab.setLayout(layout)
        self.tabs.addTab(tab, self.translate("about"))
    
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
        
        # Format de date
        self.date_radio_group.buttonClicked.connect(self.on_date_format_changed)
        self.date_custom_edit.textChanged.connect(self.on_date_custom_format_changed)
        
        # Signaux thread-safe pour communication avec workers
        self.progress_update_signal.connect(self._handle_progress_update)
        self.compilation_error_signal.connect(self._handle_compilation_error)
        self.compilation_success_signal.connect(self._handle_compilation_success)
        self.progress_finish_signal.connect(self._handle_progress_finish)
        self.progress_reset_signal.connect(self._handle_progress_reset)
        self.button_enable_signal.connect(self._handle_button_enable)
        
        # Option nom de fichier
        self.combo_filename_option.currentIndexChanged.connect(self.on_filename_option_changed)
        
        # Options informations préliminaires
        self.checkbox_preliminary.stateChanged.connect(self.toggle_preliminary_options)
        self.combo_preliminary_source.currentIndexChanged.connect(self.on_preliminary_source_changed)

    
    def on_filename_option_changed(self, index):
        """
        Appelé lorsque l'option de nom de fichier change.
        
        Args:
            index: Index sélectionné dans le combo
        """
        if index >= 0:
            self.filename_option = self.combo_filename_option.itemData(index)

    def toggle_preliminary_options(self, state):
        """
        Active ou désactive les options d'informations préliminaires.
        
        Args:
            state: État de la checkbox
        """
        is_enabled = state == Qt.CheckState.Checked.value
        self.include_preliminary = is_enabled
        self.combo_preliminary_source.setEnabled(is_enabled)
        
        if is_enabled and self.combo_preliminary_source.count() == 0:
            # Charger les fichiers dans le combo si pas encore fait
            self.update_preliminary_source_combo()

    def on_preliminary_source_changed(self, index):
        """
        Appelé lorsque l'utilisateur change le fichier source pour les informations préliminaires.
        
        Args:
            index: Index du fichier sélectionné
        """
        if index >= 0:
            self.preliminary_source_file = self.combo_preliminary_source.currentText()

    def update_preliminary_source_combo(self):
        """
        Met à jour la liste des fichiers dans le combo des sources préliminaires.
        """
        self.combo_preliminary_source.clear()
        if not self.files:  
            return
            
        for file in self.files:
            self.combo_preliminary_source.addItem(file)
        
        # Sélectionner le premier fichier par défaut
        if self.files:
            self.combo_preliminary_source.setCurrentIndex(0)
            self.preliminary_source_file = self.files[0]

    
    def update_ui_language(self):
        """Met à jour la langue de l'interface utilisateur."""
        # Mettre à jour le titre de la fenêtre
        self.setWindowTitle(self.translate("app_title"))
        
        # Mettre à jour les onglets
        self.tabs.setTabText(0, self.translate("compilation_options"))
        self.tabs.setTabText(1, self.translate("advanced_options"))
        self.tabs.setTabText(2, self.translate("date_format"))
        self.tabs.setTabText(3, self.translate("preview"))
        self.tabs.setTabText(4, self.translate("help"))
        self.tabs.setTabText(5, self.translate("about"))
        
        # Mettre à jour les groupes
        for group in self.findChildren(QGroupBox):
            if group.title() == "Sélection des fichiers" or group.title() == "File Selection":
                group.setTitle(self.translate("file_selection"))
            elif group.title() == "Options de compilation" or group.title() == "Compilation Options":
                group.setTitle(self.translate("compilation_options"))
            elif group.title() == "Formats de fichiers supportés" or group.title() == "Supported File Formats":
                group.setTitle(self.translate("file_formats"))
            elif group.title().startswith("Options de format de date") or group.title().startswith("Date Format Options"):
                group.setTitle(self.translate("date_format_options"))
        
        # Mettre à jour les labels
        self.label_directory.setText(self.translate("no_directory") if not self.directory else self.directory)
        self.label_header_start.setText(self.translate("header_start_row"))
        self.label_header.setText(self.translate("header_rows"))
        self.label_preliminary_source.setText(self.translate("preliminary_source_file"))
        self.label_output_name.setText(self.translate("output_filename"))
        self.label_sort_column.setText(self.translate("sort_column"))
        self.update_datetime()
        self.update_selection_count()
        
        # Mettre à jour les checkboxes
        self.checkbox_all_files.setText(self.translate("select_all_files"))
        self.checkbox_repeat_header.setText(self.translate("repeat_headers"))
        self.checkbox_merge_headers.setText(self.translate("merge_headers"))
        self.label_filename_option.setText(self.translate("filename_option"))
        self.checkbox_preliminary.setText(self.translate("include_preliminary"))
        self.checkbox_verify_files.setText(self.translate("enable_verification"))
        self.checkbox_remove_duplicates.setText(self.translate("remove_duplicates"))
        self.checkbox_remove_empty_rows.setText(self.translate("remove_empty_rows"))
        self.checkbox_sort_data.setText(self.translate("sort_data"))
        self.checkbox_auto_width.setText(self.translate("auto_width"))
        self.checkbox_freeze_header.setText(self.translate("freeze_headers"))
        self.checkbox_excel.setText(self.translate("excel_files"))
        self.checkbox_csv.setText(self.translate("text_files"))
        
        
        # Mettre à jour les boutons
        self.button_choose_directory.setText(self.translate("choose_directory"))
        self.button_compile.setText(self.translate("start_compilation"))
        self.preview_refresh_button.setText(self.translate("refresh_preview"))


         # Mettre à jour les items du combo filename
        current_data = self.combo_filename_option.currentData()
        self.combo_filename_option.clear()
        self.combo_filename_option.addItem(self.translate("filename_none"), "none")
        self.combo_filename_option.addItem(self.translate("filename_with_extension"), "with_extension")
        self.combo_filename_option.addItem(self.translate("filename_without_extension"), "without_extension")
        
        # Restaurer la sélection
        index = self.combo_filename_option.findData(current_data)
        if index >= 0:
            self.combo_filename_option.setCurrentIndex(index)
        
            
        # Mettre à jour les boutons radio du format de date
        for format_key, label_key in [
            ("STANDARD", "date_format_standard"),
            ("FRENCH", "date_format_french"),
            ("US", "date_format_us"),
            ("DATETIME", "date_format_datetime"),
            ("DATETIME_FRENCH", "date_format_datetime_french"),
            ("DATE_ONLY", "date_format_date_only"),
            ("TIME_ONLY", "date_format_time_only"),
            ("SHORT", "date_format_short"),
            ("CUSTOM", "date_format_custom")
        ]:
            if format_key in self.date_radio_buttons:
                self.date_radio_buttons[format_key].setText(self.translate(label_key))
        
        # Recréer le menu
        menubar = self.menuBar()
        menubar.clear()
        self.create_menu()
        
        # Mettre à jour les infos de prévisualisation
        self.preview_info_label.setText(self.translate("preview_limited", 200))
        
        # Recréer la barre d'outils
        for toolbar in self.findChildren(QToolBar):
            self.removeToolBar(toolbar)
        self.create_toolbar()
        
        # Recharger les onglets d'aide et à propos
        self.tabs.removeTab(5)  # À propos
        self.tabs.removeTab(4)  # Aide
        self.create_help_tab()
        self.create_about_tab()
    
    def update_datetime(self):
        """Met à jour l'affichage de la date et de l'heure."""
        current_datetime = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        self.datetime_label.setText(self.translate("date_time", current_datetime))
    
    def toggle_verification(self, state):
        """Active ou désactive la vérification préliminaire des fichiers."""
        self.verification_enabled = state == Qt.CheckState.Checked.value
    
    def toggle_all_files(self, state):
        """Sélectionne ou désélectionne tous les fichiers."""
        for i in range(self.list_files.count()):
            self.list_files.item(i).setSelected(state == Qt.CheckState.Checked.value)
    
    def select_all_files(self):
        """Sélectionne tous les fichiers."""
        self.checkbox_all_files.setChecked(True)
    
    def deselect_all_files(self):
        """Désélectionne tous les fichiers."""
        self.checkbox_all_files.setChecked(False)
    
    def invert_file_selection(self):
        """Inverse la sélection des fichiers."""
        for i in range(self.list_files.count()):
            item = self.list_files.item(i)
            item.setSelected(not item.isSelected())
    
    def update_selection_count(self):
        """Met à jour le compteur de fichiers sélectionnés."""
        selected_count = len(self.list_files.selectedItems())
        self.label_file_count.setText(self.translate("files_selected", selected_count))
        self.button_compile.setEnabled(selected_count > 0)
    
    def toggle_sort_options(self, state):
        """Active ou désactive les options de tri."""
        self.lineedit_sort_column.setEnabled(state == Qt.CheckState.Checked.value)
    
    def on_date_format_changed(self, button):
        """
        Appelé lorsque l'utilisateur change de format de date.
        
        Args:
            button: Bouton radio sélectionné
        """
        for format_key, radio in self.date_radio_buttons.items():
            if radio == button:
                self.date_custom_edit.setEnabled(format_key == "CUSTOM")
                break
        
        self.update_date_example()
    
    def on_date_custom_format_changed(self, text):
        """
        Appelé lorsque l'utilisateur modifie le format personnalisé.
        
        Args:
            text: Nouveau texte du format personnalisé
        """
        DATE_FORMATS["CUSTOM"]["format"] = text
        self.update_date_example()
    
    def update_date_example(self):
        """Met à jour l'exemple de format de date."""
        now = datetime.now()
        
        for format_key, radio in self.date_radio_buttons.items():
            if radio.isChecked():
                if format_key == "CUSTOM":
                    try:
                        # Convertir le format personnalisé en format de date Python
                        user_format = self.date_custom_edit.text()
                        # Remplacer les tokens de format
                        py_format = user_format.replace("dd", "%d").replace("MM", "%m").replace("yyyy", "%Y")
                        py_format = py_format.replace("HH", "%H").replace("mm", "%M").replace("ss", "%S")
                        py_format = py_format.replace("yy", "%y")
                        
                        formatted_date = now.strftime(py_format)
                        self.date_example_label.setText(formatted_date)
                    except Exception:
                        self.date_example_label.setText("Format invalide")
                else:
                    # Utiliser le format prédéfini
                    date_format = DATE_FORMATS[format_key]["format"]
                    
                    # Convertir en format Python
                    py_format = date_format.replace("dd", "%d").replace("MM", "%m").replace("yyyy", "%Y")
                    py_format = py_format.replace("HH", "%H").replace("mm", "%M").replace("ss", "%S")
                    py_format = py_format.replace("yy", "%y")
                    
                    formatted_date = now.strftime(py_format)
                    self.date_example_label.setText(formatted_date)
                break
    
    def apply_date_format(self):
        """Applique le format de date sélectionné."""
        for format_key, radio in self.date_radio_buttons.items():
            if radio.isChecked():
                self.date_format = format_key
                
                if format_key == "CUSTOM":
                    # Enregistrer le format personnalisé
                    custom_format = self.date_custom_edit.text()
                    DATE_FORMATS["CUSTOM"]["format"] = custom_format
                    
                    # Générer un format Excel personnalisé
                    excel_format = custom_format
                    excel_format = excel_format.replace("dd", "dd").replace("MM", "mm").replace("yyyy", "yyyy")
                    excel_format = excel_format.replace("HH", "hh").replace("mm", "mm").replace("ss", "ss")
                    excel_format = excel_format.replace("yy", "yy")
                    
                    DATE_FORMATS["CUSTOM"]["excel_format"] = excel_format
                
                QMessageBox.information(
                    self,
                    self.translate("date_format"),
                    self.translate("settings_saved")
                )
                break
    
    def on_preview_file_changed(self, index):
        """
        Appelé lorsque l'utilisateur change de fichier dans le combobox de prévisualisation.
        
        Args:
            index: Index du fichier sélectionné
        """
        if index >= 0:
            self.refresh_preview()

    def refresh_preview(self):
        """Actualise la prévisualisation du fichier sélectionné avec cache et chargement asynchrone."""
        if not self.directory or self.preview_combo.count() == 0:
            return
        
        current_index = self.preview_combo.currentIndex()
        if current_index < 0:
            return
        
        selected_file = self.preview_combo.itemText(current_index)
        file_path = os.path.join(self.directory, selected_file)
        
        try:
            # Vérifier si le fichier existe
            if not os.path.exists(file_path):
                self.display_preview_tab([], [])
                return
            
            # Obtenir l'heure de modification du fichier
            file_mtime = os.path.getmtime(file_path)
            
            # Vérifier le cache en premier
            cached_data = self.preview_cache.get(file_path, file_mtime)
            if cached_data is not None:
                headers, data = cached_data
                self.display_preview_tab(headers, data)
                return
            
            # Afficher un indicateur de chargement
            self.display_preview_tab([["Chargement en cours..."]], [["⏳ Chargement du fichier..."]])
            
            # Annuler le worker précédent s'il existe
            if self.preview_worker and self.preview_worker.isRunning():
                self.preview_worker.cancel()
                self.preview_worker.wait(1000)  # Attendre 1 seconde max
            
            # Démarrer le chargement asynchrone
            header_start = self.spinbox_header_start.value()
            header_rows = self.spinbox_header.value()
            
            self.preview_worker = PreviewWorker(file_path, header_start, header_rows, self)
            self.preview_worker.preview_loaded.connect(self._on_preview_loaded)
            self.preview_worker.preview_error.connect(self._on_preview_error)
            self.preview_worker.start()
            
        except Exception as e:
            self.display_preview_tab([["Erreur"]], [[f"❌ Erreur: {str(e)}"]])
    
    def _on_preview_loaded(self, file_path, headers, data):
        """Handler appelé quand la prévisualisation est chargée avec succès"""
        try:
            # Ajouter au cache
            file_mtime = os.path.getmtime(file_path)
            self.preview_cache.put(file_path, (headers, data), file_mtime)
            
            # Afficher la prévisualisation
            self.display_preview_tab(headers, data)
            
        except Exception as e:
            self.display_preview_tab([["Erreur"]], [[f"❌ Erreur lors de l'affichage: {str(e)}"]])
    
    def _on_preview_error(self, file_path, error_message):
        """Handler appelé en cas d'erreur de chargement"""
        filename = os.path.basename(file_path)
        self.display_preview_tab(
            [["Erreur de chargement"]], 
            [[f"❌ Impossible de charger {filename}"], [f"Erreur: {error_message}"]]
        )
    
    def load_excel_preview_tab(self, file_path):
        """
        Charge la prévisualisation d'un fichier Excel dans l'onglet prévisualisation.
        
        Args:
            file_path: Chemin du fichier Excel à prévisualiser
        """
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            ws = wb.active
            
            # Obtenir les données (limité à 200 lignes)
            preview_data = []
            headers = []
            
            # Récupérer les en-têtes
            header_start = self.spinbox_header_start.value()
            header_rows = self.spinbox_header.value()
            
            for row in range(header_start, header_start + header_rows):
                header_row = []
                for cell in ws[row]:
                    header_row.append(cell.value)
                headers.append(header_row)
            
            # Récupérer les données
            row_count = 0
            for row in ws.iter_rows(min_row=header_start + header_rows):
                if row_count >= 200:
                    break
                
                row_data = [cell.value for cell in row]
                preview_data.append(row_data)
                row_count += 1
            
            # Afficher les données
            self.display_preview_tab(headers, preview_data)
            
            wb.close()
        except FileNotFoundError:
            raise FileNotFoundError(f"Fichier Excel non trouvé: {file_path}")
        except PermissionError:
            raise PermissionError(f"Impossible d'accéder au fichier Excel (fichier ouvert?): {file_path}")
        except openpyxl.utils.exceptions.InvalidFileException:
            raise ValueError(f"Fichier Excel corrompu ou format invalide: {file_path}")
        except (AttributeError, KeyError) as e:
            raise ValueError(f"Structure Excel invalide dans {file_path}: {str(e)}")
        except MemoryError:
            raise MemoryError(f"Fichier Excel trop volumineux: {file_path}")
        except Exception as e:
            # Gestion d'erreur de fallback avec plus de contexte
            raise RuntimeError(f"Erreur inattendue lors de la lecture Excel {file_path}: {type(e).__name__}: {str(e)}")
    
    def load_csv_preview_tab(self, file_path):
        """
        Charge la prévisualisation d'un fichier CSV dans l'onglet prévisualisation.
        
        Args:
            file_path: Chemin du fichier CSV à prévisualiser
        """
        try:
            # Détection de l'encodage du fichier
            encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']
            delimiter = ','
            
            for encoding in encodings:
                try:
                    with open(file_path, 'r', encoding=encoding) as f:
                        sample = f.read(4096)
                        sniffer = csv.Sniffer()
                        delimiter = sniffer.sniff(sample).delimiter
                        break
                except Exception:
                    continue
            
            # Lecture du CSV avec pandas
            df = pd.read_csv(file_path, delimiter=delimiter, header=None, encoding=encoding)
            
            # Limiter le nombre de lignes
            if len(df) > 200 + self.spinbox_header_start.value() + self.spinbox_header.value():
                df = df.iloc[:(200 + self.spinbox_header_start.value() + self.spinbox_header.value())]
            
            # Récupérer les en-têtes
            headers = []
            header_start = self.spinbox_header_start.value() - 1
            header_rows = self.spinbox_header.value()
            
            for row in range(header_start, header_start + header_rows):
                if row < len(df):
                    headers.append(df.iloc[row].tolist())
            
            # Récupérer les données
            preview_data = []
            for row in range(header_start + header_rows, len(df)):
                preview_data.append(df.iloc[row].tolist())
            
            # Afficher les données
            self.display_preview_tab(headers, preview_data)
            
        except FileNotFoundError:
            raise FileNotFoundError(f"Fichier CSV non trouvé: {file_path}")
        except PermissionError:
            raise PermissionError(f"Impossible d'accéder au fichier CSV: {file_path}")
        except UnicodeDecodeError as e:
            raise ValueError(f"Problème d'encodage dans {file_path}: {str(e)}")
        except pd.errors.EmptyDataError:
            raise ValueError(f"Fichier CSV vide: {file_path}")
        except pd.errors.ParserError as e:
            raise ValueError(f"Format CSV invalide dans {file_path}: {str(e)}")
        except MemoryError:
            raise MemoryError(f"Fichier CSV trop volumineux: {file_path}")
        except Exception as e:
            # Gestion d'erreur de fallback avec plus de contexte
            raise RuntimeError(f"Erreur inattendue lors de la lecture CSV {file_path}: {type(e).__name__}: {str(e)}")
    
    def display_preview_tab(self, headers, data):
        """
        Affiche les données dans le tableau de prévisualisation de l'onglet.
        
        Args:
            headers: Données d'en-tête
            data: Données à afficher
        """
        if not headers or not headers[0]:
            return
        
        # Configurer le tableau
        self.preview_table.clear()
        self.preview_table.setRowCount(len(data))
        
        # Combiner les en-têtes multi-lignes pour l'affichage
        combined_headers = self._combine_multi_headers(headers)
        self.preview_table.setColumnCount(len(combined_headers))
        
        # Définir les en-têtes des colonnes combinées
        self.preview_table.setHorizontalHeaderLabels([str(h) if h is not None else "" for h in combined_headers])
        
        # Ajouter les données
        for row_idx, row_data in enumerate(data):
            for col_idx, value in enumerate(row_data[:len(headers[-1])]):
                item = QTableWidgetItem(str(value) if value is not None else "")
                self.preview_table.setItem(row_idx, col_idx, item)
        
        # Ajuster la taille des colonnes
        self.preview_table.resizeColumnsToContents()
    
    def _combine_multi_headers(self, headers):
        """
        Combine les en-têtes multi-lignes en un seul en-tête pour l'affichage.
        
        Args:
            headers: Liste des lignes d'en-têtes
            
        Returns:
            Liste des en-têtes combinées
        """
        if not headers or len(headers) == 1:
            return headers[0] if headers else []
        
        # Déterminer le nombre de colonnes maximum
        max_cols = max(len(row) for row in headers if row) if headers else 0
        combined = []
        
        for col_idx in range(max_cols):
            combined_cell = []
            for row in headers:
                if row and col_idx < len(row) and row[col_idx] not in [None, '']:
                    cell_value = str(row[col_idx]).strip()
                    if cell_value and cell_value not in combined_cell:
                        combined_cell.append(cell_value)
            
            # Combiner les parties non vides avec un séparateur
            combined.append(' - '.join(combined_cell) if combined_cell else f'Col_{col_idx + 1}')
        
        return combined
        
        # Mettre à jour l'étiquette d'information
        if len(data) >= 200:
            self.preview_info_label.setText(self.translate("preview_limited", 200))
        else:
            self.preview_info_label.setText(f"{len(data)} lignes affichées")

    def cleanup_resources(self):
        """Nettoie proprement toutes les ressources."""
        print("🧹 Nettoyage des ressources...")
        
        # Arrêter le timer proprement
        if hasattr(self, 'datetime_timer') and self.datetime_timer:
            self.datetime_timer.stop()
            print("  ✅ Timer arrêté")
        
        # Arrêter le worker proprement (pas brutalement)
        if hasattr(self, 'compilation_worker') and self.compilation_worker:
            if self.compilation_worker.isRunning():
                print("  ⏳ Arrêt du worker en cours...")
                self.compilation_worker.quit()  # ← Plus doux que terminate()
                
                # Attendre max 3 secondes
                if not self.compilation_worker.wait(3000):
                    print("  ⚠️ Worker ne répond pas, arrêt forcé")
                    self.compilation_worker.terminate()
                else:
                    print("  ✅ Worker arrêté proprement")
        
        # Fermer les fichiers ouverts si nécessaire
        
        print("🧹 Nettoyage terminé")




    # MÉTHODES PRINCIPALES
    def choose_directory(self):
        """Ouvre une boîte de dialogue pour choisir le répertoire de travail avec validation de sécurité."""
        try:
            directory = QFileDialog.getExistingDirectory(
                self,
                self.translate("choose_directory"),
                self.directory if self.directory else os.path.expanduser("~")
            )
            
            if directory:
                # Validation de sécurité du répertoire
                security_validator = SecurityValidator()
                is_valid, error_message = security_validator.validate_directory_path(directory)
                
                if is_valid:
                    self.directory = directory
                    self.label_directory.setText(directory)
                    self.load_files()
                    
                    # Sauvegarder automatiquement la session après changement de répertoire
                    self.save_current_session()
                    
                    # Metrics de monitoring
                    if MONITORING_AVAILABLE:
                        log_event("directory_selected", {
                            "directory": directory,
                            "timestamp": datetime.now().isoformat()
                        })
                        increment_counter("directory_selections_total")
                    
                    # Logger structuré
                    if hasattr(self, 'structured_logger') and self.structured_logger:
                        self.structured_logger.log_structured("INFO", "directory_selected", 
                                                            directory=directory)
                    
                    logging.info(f"Répertoire sélectionné: {directory}")
                else:
                    # Afficher l'erreur de sécurité
                    QMessageBox.warning(
                        self,
                        "Erreur de Sécurité",
                        f"Le répertoire sélectionné n'est pas sécurisé:\n{error_message}"
                    )
                    
                    # Metrics de monitoring pour les erreurs
                    if MONITORING_AVAILABLE:
                        log_event("directory_selection_failed", {
                            "directory": directory,
                            "error": error_message,
                            "timestamp": datetime.now().isoformat()
                        })
                        increment_counter("directory_selection_failures_total")
                    
                    # Logger structuré
                    if hasattr(self, 'structured_logger') and self.structured_logger:
                        self.structured_logger.log_structured("WARNING", "directory_selection_failed", 
                                                            directory=directory,
                                                            error_message=error_message)
                    
                    logging.warning(f"Répertoire rejeté pour des raisons de sécurité: {directory} - {error_message}")
                    
        except Exception as e:
            # Utiliser le gestionnaire d'erreurs robuste
            if hasattr(self, 'error_handler') and self.error_handler:
                success, message, recovery_data = self.error_handler.handle_error(
                    e, {'operation': 'choose_directory', 'current_directory': self.directory}
                )
                
                if not success:
                    QMessageBox.critical(
                        self,
                        "Erreur",
                        f"Impossible de sélectionner le répertoire:\n{message}"
                    )
            else:
                # Fallback si le gestionnaire d'erreurs n'est pas disponible
                logging.error(f"Erreur lors de la sélection du répertoire: {e}")
                QMessageBox.critical(
                    self,
                    "Erreur",
                    f"Impossible de sélectionner le répertoire:\n{str(e)}"
                )
    
    def load_files(self):
        """Charge la liste des fichiers Excel et CSV du répertoire sélectionné avec validation de sécurité."""
        if not self.directory:
            return
        
        self.files = []
        self.list_files.clear()
        self.preview_combo.clear()
        
        # Nettoyer le cache de prévisualisation quand le répertoire change
        if hasattr(self, 'preview_cache'):
            self.preview_cache.clear()
        
        # Extensions supportées
        excel_extensions = ['.xlsx', '.xls', '.xlsm', '.xltx', '.xltm']
        csv_extensions = ['.csv', '.tsv', '.txt'] if self.checkbox_csv.isChecked() else []
        extensions = excel_extensions + csv_extensions
        
        # Initialiser le validateur de sécurité
        security_manager = SecurityManager()
        safe_files = []
        rejected_files = []
        total_size = 0
        
        try:
            for file in os.listdir(self.directory):
                if any(file.lower().endswith(ext) for ext in extensions):
                    try:
                        # Validation du nom de fichier
                        security_manager.validate_filename(file)
                        
                        # Construire le chemin complet
                        file_path = os.path.join(self.directory, file)
                        safe_path = security_manager.sanitize_file_path(file_path)
                        
                        # Vérifier la taille du fichier
                        file_size = os.path.getsize(safe_path)
                        if file_size > MAX_FILE_SIZE:
                            rejected_files.append(f"{file} (trop volumineux: {file_size/(1024*1024):.1f}MB)")
                            continue
                        
                        # Vérifier la taille totale de session
                        if total_size + file_size > MAX_SESSION_SIZE:
                            rejected_files.append(f"{file} (dépassement limite session)")
                            continue
                        
                        # Vérification d'intégrité basique
                        integrity_info = security_manager.verify_file_integrity(safe_path)
                        
                        # Fichier validé - l'ajouter à la liste
                        safe_files.append(file)
                        total_size += file_size
                        
                        item = QListWidgetItem(file)
                        # Ajouter une icône de sécurité pour indiquer que le fichier est sûr
                        item.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogApplyButton))
                        self.list_files.addItem(item)
                        self.preview_combo.addItem(file)
                        
                    except SecurityError as e:
                        rejected_files.append(f"{file} ({str(e)})")
                        logging.warning(f"Fichier rejeté pour sécurité: {file} - {e}")
                        continue
                    except Exception as e:
                        rejected_files.append(f"{file} (erreur: {str(e)})")
                        logging.error(f"Erreur lors de la validation de {file}: {e}")
                        continue

            self.files = safe_files
            
            # Réinitialiser le checkbox à décoché
            self.checkbox_all_files.setChecked(False) 
            
            # NOUVEAU : Mettre à jour le combo des sources préliminaires
            self.update_preliminary_source_combo()
            
            # Afficher les statistiques de sécurité
            if rejected_files:
                self._show_security_report(safe_files, rejected_files, total_size)
            
            # Metrics de monitoring pour le chargement des fichiers
            if MONITORING_AVAILABLE:
                log_event("files_loaded", {
                    "safe_files_count": len(safe_files),
                    "rejected_files_count": len(rejected_files),
                    "total_size_mb": total_size / (1024 * 1024),
                    "directory": self.directory,
                    "timestamp": datetime.now().isoformat()
                })
                set_gauge("current_files_count", len(safe_files))
                set_gauge("current_session_size_mb", total_size / (1024 * 1024))
                if rejected_files:
                    increment_counter("files_rejected_total", {"count": len(rejected_files)})
            
            # Logger structuré
            if hasattr(self, 'structured_logger') and self.structured_logger:
                self.structured_logger.log_structured("INFO", "files_loaded", 
                                                    safe_files_count=len(safe_files),
                                                    rejected_files_count=len(rejected_files),
                                                    total_size_mb=total_size / (1024 * 1024),
                                                    directory=self.directory)
            
            logging.info(f"Fichiers chargés: {len(safe_files)} fichiers sécurisés, {len(rejected_files)} rejetés")
            self.update_selection_count()
            
            if self.files:
                self.refresh_preview()
                
        except PermissionError:
            QMessageBox.warning(
                self,
                self.translate("error"),
                "Accès au répertoire refusé. Vérifiez vos permissions."
            )
        except Exception as e:
            QMessageBox.warning(
                self,
                self.translate("error"),
                f"Erreur lors du chargement des fichiers: {str(e)}"
            )
    
    def _show_security_report(self, safe_files: List[str], rejected_files: List[str], total_size: int):
        """Affiche un rapport de sécurité des fichiers traités"""
        msg = QMessageBox(self)
        msg.setWindowTitle("🔒 Rapport de Sécurité")
        msg.setIcon(QMessageBox.Icon.Information)
        
        safe_count = len(safe_files)
        rejected_count = len(rejected_files)
        total_count = safe_count + rejected_count
        
        # Calcul des pourcentages
        safe_percent = (safe_count / total_count * 100) if total_count > 0 else 0
        
        # Formatage de la taille
        total_size_mb = total_size / (1024 * 1024)
        max_size_gb = MAX_SESSION_SIZE / (1024 * 1024 * 1024)
        
        report = f"""<h3>📊 Analyse de Sécurité Terminée</h3>
        
<b>✅ Fichiers sécurisés :</b> {safe_count} ({safe_percent:.1f}%)
<b>❌ Fichiers rejetés :</b> {rejected_count}
<b>📦 Taille totale :</b> {total_size_mb:.1f} MB / {max_size_gb:.1f} GB

<h4>📋 Fichiers rejetés :</h4>
<ul>"""
        
        for rejected in rejected_files[:10]:  # Limiter à 10 pour l'affichage
            report += f"<li><code>{rejected}</code></li>"
        
        if len(rejected_files) > 10:
            report += f"<li><i>... et {len(rejected_files) - 10} autres</i></li>"
        
        report += """</ul>

<p><b>🔒 Seuls les fichiers sécurisés sont disponibles pour compilation.</b></p>"""
        
        msg.setText(report)
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()
    
    def compile_files(self):
        """Lance la compilation des fichiers sélectionnés avec validation de sécurité finale."""
        # Démarrer le monitoring de performance pour cette opération
        operation_id = None
        if hasattr(self, 'performance_monitor') and self.performance_monitor:
            operation_id = self.performance_monitor.start_operation("file_compilation")
            self._current_operation_id = operation_id  # Sauvegarder pour utilisation ultérieure
        
        # Logger structuré pour l'opération
        if hasattr(self, 'structured_logger') and self.structured_logger:
            self.structured_logger.log_structured("INFO", "compilation_started", 
                                                operation_id=operation_id)
        
        selected_files = [item.text() for item in self.list_files.selectedItems()]
        
        if not selected_files:
            # Logger structuré pour l'erreur
            if hasattr(self, 'structured_logger') and self.structured_logger:
                self.structured_logger.log_structured("WARNING", "compilation_aborted", 
                                                    reason="no_files_selected",
                                                    operation_id=operation_id)
            
            # Finaliser le monitoring en cas d'abandon
            if hasattr(self, 'performance_monitor') and self.performance_monitor and operation_id:
                self.performance_monitor.end_operation("file_compilation", operation_id)
            
            # Incrementer les métriques d'erreur
            if MONITORING_AVAILABLE:
                increment_counter("compilation_aborted_total", {"reason": "no_files_selected"})
            
            QMessageBox.warning(
                self,
                self.translate("warning"),
                self.translate("select_files_message")
            )
            return
        
        # Validation de sécurité finale avant compilation
        security_validator = SecurityValidator()
        is_valid, error_message, file_details = security_validator.validate_file_selection(
            selected_files, self.directory
        )
        
        if not is_valid:
            # Logger structuré pour l'erreur de sécurité
            if hasattr(self, 'structured_logger') and self.structured_logger:
                self.structured_logger.log_structured("ERROR", "compilation_aborted", 
                                                    reason="security_validation_failed",
                                                    error_message=error_message,
                                                    operation_id=operation_id)
            
            # Finaliser le monitoring en cas d'abandon
            if hasattr(self, 'performance_monitor') and self.performance_monitor and operation_id:
                self.performance_monitor.end_operation("file_compilation", operation_id)
            
            # Incrementer les métriques d'erreur
            if MONITORING_AVAILABLE:
                increment_counter("compilation_aborted_total", {"reason": "security_validation_failed"})
            
            QMessageBox.critical(
                self,
                "🔒 Erreur de Sécurité",
                f"Les fichiers sélectionnés ne peuvent pas être compilés pour des raisons de sécurité:\n\n{error_message}"
            )
            logging.error(f"Compilation bloquée par sécurité: {error_message}")
            return
        
        # Afficher un résumé de sécurité si demandé
        total_size_mb = sum(details['size'] for details in file_details) / (1024 * 1024)
        logging.info(f"Validation de sécurité réussie: {len(file_details)} fichiers, {total_size_mb:.1f}MB")
        
        # Enregistrer les métriques de la compilation
        if MONITORING_AVAILABLE:
            set_gauge("compilation_files_count", len(selected_files))
            set_gauge("compilation_total_size_mb", total_size_mb)
            log_event("compilation_validation_success", {
                "files_count": len(selected_files),
                "total_size_mb": total_size_mb
            })
        
        # NOUVELLE VÉRIFICATION : Cohérence des en-têtes multi-lignes
        if len(selected_files) > 1 and self.spinbox_header.value() > 1:
            try:
                is_consistent, warnings = self._verify_multiheader_consistency(selected_files)
                
                if not is_consistent:
                    # Afficher dialogue d'avertissement
                    warning_dialog = QMessageBox(self)
                    warning_dialog.setIcon(QMessageBox.Icon.Warning)
                    warning_dialog.setWindowTitle("⚠️ Incohérence des en-têtes multi-lignes")
                    warning_dialog.setText(
                        "Les fichiers sélectionnés ont des structures d'en-têtes incompatibles.\n"
                        "La compilation peut produire des résultats incohérents.\n\n"
                        "Voulez-vous continuer malgré tout ?"
                    )
                    warning_dialog.setDetailedText("\n".join(warnings))
                    warning_dialog.setStandardButtons(
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                    )
                    warning_dialog.setDefaultButton(QMessageBox.StandardButton.No)
                    
                    if warning_dialog.exec() == QMessageBox.StandardButton.No:
                        logging.info("Compilation annulée par l'utilisateur suite à incohérence des en-têtes")
                        return  # Annuler la compilation
                    else:
                        logging.warning(f"Compilation continuée malgré incohérence des en-têtes: {len(warnings)} avertissements")
                        
            except Exception as e:
                # Si la vérification échoue, logger mais continuer
                logging.warning(f"Erreur lors de la vérification des en-têtes multi-lignes: {e}")
        
        # Vérification préliminaire si activée
        if self.verification_enabled:
            compatible_files, incompatible_files = self.verify_files(selected_files)
            
            if incompatible_files:
                # Afficher le rapport de vérification
                dialog = VerificationReportDialog(self, compatible_files, incompatible_files)
                result = dialog.exec()
                
                if result == QDialog.DialogCode.Rejected:
                    return  # L'utilisateur a annulé
                
                if dialog.continue_with_compatible:
                    # Continuer avec seulement les fichiers compatibles
                    selected_files = compatible_files
                    if not selected_files:
                        QMessageBox.information(
                            self,
                            self.translate("info"),
                            "Aucun fichier compatible pour la compilation."
                        )
                        return
        
        # Préparer les paramètres de compilation
        header_start_row = self.spinbox_header_start.value()
        header_rows = self.spinbox_header.value()
        filename_option = self.combo_filename_option.currentData()
        sort_data = self.checkbox_sort_data.isChecked()
        sort_column = self.get_sort_column_index()
        repeat_headers = self.checkbox_repeat_header.isChecked()
        remove_empty_rows = self.checkbox_remove_empty_rows.isChecked()
        remove_duplicates = self.checkbox_remove_duplicates.isChecked()
        
        # NOUVEAU : Paramètres pour les informations préliminaires
        include_preliminary = self.checkbox_preliminary.isChecked()
        preliminary_source_file = self.preliminary_source_file if include_preliminary else ""
        
        # Désactiver le bouton et basculer en mode compact
        self.button_compile.setEnabled(False)
        self.enter_compact_mode()  # Gérer l'interface responsive
        self.progress_widget.start_operation(f"Compilation de {len(selected_files)} fichier(s)")
        
        # Sauvegarder la liste des fichiers pour le calcul de progression
        self.files = selected_files
        
        # Créer et démarrer le worker
        self.compilation_worker = CompilationWorker(
            selected_files,
            self.directory,
            header_start_row,
            header_rows,
            filename_option,
            sort_data,
            sort_column,
            repeat_headers,
            remove_empty_rows,
            remove_duplicates,
            self.date_format,
            include_preliminary, 
            preliminary_source_file
        )
        
        # Connecter les signaux de manière thread-safe
        self.compilation_worker.progress.connect(self.update_progress, Qt.ConnectionType.QueuedConnection)
        self.compilation_worker.error.connect(self.compilation_error, Qt.ConnectionType.QueuedConnection)
        self.compilation_worker.finished.connect(self.compilation_finished, Qt.ConnectionType.QueuedConnection)
        
        # Connecter les nouveaux signaux thread-safe
        if hasattr(self.compilation_worker, 'progress_detail'):
            self.compilation_worker.progress_detail.connect(self._handle_progress_detail, Qt.ConnectionType.QueuedConnection)
        if hasattr(self.compilation_worker, 'warning'):
            self.compilation_worker.warning.connect(self._handle_warning, Qt.ConnectionType.QueuedConnection)
        if hasattr(self.compilation_worker, 'info'):
            self.compilation_worker.info.connect(self._handle_info, Qt.ConnectionType.QueuedConnection)
        if hasattr(self.compilation_worker, 'status_update'):
            self.compilation_worker.status_update.connect(self._handle_status_update, Qt.ConnectionType.QueuedConnection)
        
        # Démarrer la compilation
        self.compilation_worker.start()
        
        # Démarrer le timer de timeout
        self.start_worker_timeout()
        
        logging.info(f"Début de la compilation de {len(selected_files)} fichiers")
    
    def verify_files(self, file_list):
        """
        Vérifie la compatibilité des fichiers pour la compilation.
        
        Args:
            file_list: Liste des noms de fichiers à vérifier
            
        Returns:
            Tuple[List[str], List[Tuple[str, str]]]: (fichiers_compatibles, fichiers_incompatibles_avec_raisons)
        """
        compatible_files = []
        incompatible_files = []
        
        header_start_row = self.spinbox_header_start.value()
        header_rows = self.spinbox_header.value()
        
        # Créer une barre de progression pour la vérification
        progress_dialog = QProgressDialog(
            "Vérification des fichiers en cours...",
            "Annuler",
            0,
            len(file_list),
            self
        )
        progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        progress_dialog.setMinimumDuration(0)
        
        for i, file_name in enumerate(file_list):
            if progress_dialog.wasCanceled():
                break
                
            progress_dialog.setValue(i)
            progress_dialog.setLabelText(f"Vérification: {file_name}")
            
            file_path = os.path.join(self.directory, file_name)
            
            # Vérifier le cache de validation en premier
            cached_result = self.validation_cache.get_validation_result(file_path)
            if cached_result is not None:
                is_valid, errors = cached_result
                if is_valid:
                    compatible_files.append(file_name)
                else:
                    incompatible_files.append((file_name, "; ".join(errors)))
                continue
            
            try:
                if file_name.lower().endswith(('.xlsx', '.xls', '.xlsm', '.xltx', '.xltm')):
                    is_compatible, reason = FileVerification.verify_excel_file(
                        file_path, header_start_row, header_rows
                    )
                elif file_name.lower().endswith('.csv'):
                    is_compatible, reason = FileVerification.verify_csv_file(
                        file_path, header_start_row, header_rows
                    )
                elif file_name.lower().endswith(('.tsv', '.txt')):
                    is_compatible, reason = FileVerification.verify_text_file(
                        file_path, header_start_row, header_rows
                    )
                else:
                    is_compatible = False
                    reason = "Format de fichier non supporté"
                
                # Mettre en cache le résultat de validation
                if is_compatible:
                    compatible_files.append(file_name)
                    self.validation_cache.set_validation_result(file_path, True, [])
                else:
                    incompatible_files.append((file_name, reason))
                    self.validation_cache.set_validation_result(file_path, False, [reason])
                    
            except Exception as e:
                error_msg = f"Erreur inattendue: {str(e)}"
                incompatible_files.append((file_name, error_msg))
                self.validation_cache.set_validation_result(file_path, False, [error_msg])
        
        progress_dialog.setValue(len(file_list))
        progress_dialog.close()
        
        return compatible_files, incompatible_files
    
    def _verify_multiheader_consistency(self, file_list):
        """
        Vérifie la cohérence des en-têtes multi-lignes entre fichiers
        
        Args:
            file_list: Liste des noms de fichiers à vérifier
            
        Returns:
            Tuple[bool, List[str]]: (is_consistent, warning_messages)
        """
        if len(file_list) <= 1 or self.spinbox_header.value() <= 1:
            return True, []
        
        reference_structure = None
        reference_file = None
        inconsistencies = []
        
        for file_name in file_list:
            try:
                file_path = os.path.join(self.directory, file_name)
                headers = self._extract_header_structure(file_path)
                
                if reference_structure is None:
                    reference_structure = headers
                    reference_file = file_name
                    continue
                
                # Comparer avec la structure de référence
                if not self._compare_header_structures(reference_structure, headers):
                    # Détails de l'incompatibilité pour un meilleur diagnostic
                    ref_cols = [len(row) for row in reference_structure]
                    curr_cols = [len(row) for row in headers]
                    inconsistencies.append(
                        f"❌ {file_name}: Structure d'en-têtes incompatible avec {reference_file} "
                        f"(colonnes référence: {ref_cols}, fichier actuel: {curr_cols})"
                    )
                    
            except Exception as e:
                inconsistencies.append(f"❌ {file_name}: Erreur lecture en-têtes ({str(e)})")
        
        return len(inconsistencies) == 0, inconsistencies
    
    def _extract_header_structure(self, file_path):
        """Extrait la structure des en-têtes (nombre colonnes + aperçu contenu)"""
        try:
            if file_path.lower().endswith(('.xlsx', '.xls', '.xlsm', '.xltx', '.xltm')):
                wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
                ws = wb.active
                
                headers = []
                for row in range(self.spinbox_header_start.value(), 
                               self.spinbox_header_start.value() + self.spinbox_header.value()):
                    try:
                        header_row = [cell.value for cell in ws[row]]
                        headers.append(header_row)
                    except IndexError:
                        # Ligne n'existe pas, remplir avec None
                        headers.append([])
                
                wb.close()
                return headers
            else:
                # Pour CSV/TSV/TXT
                encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']
                
                for encoding in encodings:
                    try:
                        with open(file_path, 'r', encoding=encoding) as f:
                            lines = f.readlines()
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    raise ValueError("Impossible de décoder le fichier")
                
                headers = []
                delimiter = self._detect_csv_delimiter(file_path, lines)
                
                for i in range(self.spinbox_header_start.value() - 1, 
                             self.spinbox_header_start.value() - 1 + self.spinbox_header.value()):
                    if i < len(lines):
                        headers.append(lines[i].strip().split(delimiter))
                    else:
                        headers.append([])
                
                return headers
                
        except Exception as e:
            raise Exception(f"Erreur lors de l'extraction des en-têtes: {str(e)}")
    
    def _detect_csv_delimiter(self, file_path, lines):
        """Détecte le délimiteur CSV"""
        if not lines:
            return ','
        
        # Tester les délimiteurs communs
        delimiters = [',', ';', '\t', '|']
        first_line = lines[0] if lines else ""
        
        for delimiter in delimiters:
            if delimiter in first_line:
                return delimiter
        
        return ','  # Par défaut
    
    def test_multiheader_consistency(self):
        """Méthode de test pour la vérification de cohérence des en-têtes (peut être supprimée plus tard)"""
        try:
            # Test avec des fichiers factices pour validation
            test_files = ["test1.xlsx", "test2.xlsx"]
            is_consistent, warnings = self._verify_multiheader_consistency(test_files)
            
            logging.info(f"Test de cohérence des en-têtes: {'✅ Cohérent' if is_consistent else '❌ Incohérent'}")
            if warnings:
                logging.info(f"Avertissements de test: {warnings}")
                
            return is_consistent, warnings
            
        except Exception as e:
            logging.error(f"Erreur lors du test de cohérence: {e}")
            return False, [f"Erreur de test: {str(e)}"]
    
    def _compare_header_structures(self, struct1, struct2):
        """Compare deux structures d'en-têtes"""
        # Vérifier le nombre de lignes d'en-têtes
        if len(struct1) != len(struct2):
            return False
        
        # Vérifier le nombre de colonnes pour chaque ligne
        for row1, row2 in zip(struct1, struct2):
            if len(row1) != len(row2):
                return False
        
        return True
    
    def show_preview(self):
        """Affiche la boîte de dialogue de prévisualisation des données."""
        selected_files = [item.text() for item in self.list_files.selectedItems()]
        
        if not selected_files:
            QMessageBox.information(
                self,
                self.translate("preview"),
                "Veuillez sélectionner au moins un fichier à prévisualiser."
            )
            return
        
        try:
            dialog = PreviewDialog(
                self,
                self.directory,
                selected_files,
                self.spinbox_header_start.value(),
                self.spinbox_header.value()
            )
            dialog.exec()
        except Exception as e:
            QMessageBox.warning(
                self,
                self.translate("error"),
                f"Erreur lors de la prévisualisation: {str(e)}"
            )
    
    def show_date_format_dialog(self):
        """Affiche la boîte de dialogue de choix du format de date."""
        dialog = DateFormatDialog(self, self.date_format)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.date_format = dialog.get_selected_format()
            
            # Mettre à jour l'onglet format de date
            if self.date_format in self.date_radio_buttons:
                self.date_radio_buttons[self.date_format].setChecked(True)
                if self.date_format == "CUSTOM":
                    self.date_custom_edit.setEnabled(True)
                    self.date_custom_edit.setText(DATE_FORMATS["CUSTOM"]["format"])
                else:
                    self.date_custom_edit.setEnabled(False)
            
            self.update_date_example()
    
    def show_language_dialog(self):
        """Affiche la boîte de dialogue de choix de langue."""
        dialog = LanguageDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_language = dialog.get_selected_language()
            TranslationManager().set_language(new_language)
    
    def show_about_dialog(self):
        """Affiche la boîte de dialogue À propos."""
        QMessageBox.about(
            self,
            self.translate("about"),
            f"""<h2>{self.translate("app_title")}</h2>
            <p>Version 3.0</p>
            <p>Développé par: <b>{self.translate("developer_name")}</b></p>
            <p>{self.translate("developer_title")}</p>
            <p>Email: zimkada@gmail.com</p>
            <p><i>{self.translate("copyright_notice")}</i></p>"""
        )
    
    def show_ip_warning(self):
        """Affiche l'avertissement de propriété intellectuelle au démarrage."""
        dialog = IPWarningDialog(self)
        if dialog.exec() == QDialog.DialogCode.Rejected:
            # L'utilisateur a choisi de quitter
            sys.exit(0)
    
    def check_and_show_tutorial(self):
        """Vérifie s'il faut afficher le tutoriel et l'affiche si nécessaire"""
        if TutorialManager.should_show_tutorial():
            self.show_tutorial()
    
    def show_tutorial(self):
        """Affiche l'assistant tutoriel"""
        try:
            TutorialManager.show_tutorial(self)
        except Exception as e:
            logging.error(f"Erreur lors de l'affichage du tutoriel: {e}")
            QMessageBox.warning(
                self, 
                "Erreur", 
                f"Impossible d'afficher le tutoriel:\n{str(e)}"
            )
    
    def show_user_guide(self):
        """Affiche le guide utilisateur complet"""
        try:
            guide_dialog = UserGuideDialog(self)
            guide_dialog.exec()
        except Exception as e:
            logging.error(f"Erreur lors de l'affichage du guide utilisateur: {e}")
            QMessageBox.warning(
                self, 
                "Erreur", 
                f"Impossible d'afficher le guide utilisateur:\n{str(e)}"
            )
    
    def save_settings(self):
        """Sauvegarde les paramètres actuels."""
        SettingsManager().save_settings(self)
        QMessageBox.information(
            self,
            self.translate("save_settings"),
            self.translate("settings_saved")
        )
    
    def load_settings(self):
        """Charge les paramètres sauvegardés."""
        if SettingsManager().load_settings(self):
            QMessageBox.information(
                self,
                self.translate("load_settings"),
                self.translate("settings_loaded")
            )
        else:
            QMessageBox.information(
                self,
                self.translate("load_settings"),
                "Aucun paramètre sauvegardé trouvé."
            )

    # MÉTHODES FINALES
    
    def get_sort_column_index(self):
        """
        Obtient l'index de la colonne de tri à partir du texte saisi.
        Supporte les formats : A,B,C (lettres) et 1,2,3 (nombres)
        
        Returns:
            int: Index de la colonne (0-based)
        """
        sort_text = self.lineedit_sort_column.text().strip().upper()
        
        if not sort_text:
            return 0
        
        # Format lettre (A, B, C, AA, AB, etc.)
        if sort_text.isalpha():
            try:
                # Conversion Excel-style vers index 0-based
                col_index = 0
                for char in sort_text:
                    col_index = col_index * 26 + (ord(char) - ord('A') + 1)
                return max(0, col_index - 1)  # Convertir en 0-based
            except:
                return 0
        
        # Format numérique (1=0, 2=1, etc.)
        try:
            return max(0, int(sort_text) - 1)
        except ValueError:
            return 0
    
    def update_progress(self, value):
        """
        Met à jour la barre de progression de manière thread-safe.
        
        Args:
            value: Valeur actuelle de progression
        """
        ThreadSafeGUIHelper.assert_main_thread("update_progress")
        
        total_files = len(self.files) if self.files else 1
        progress_percent = int((value / total_files) * 100)
        detail = f"Fichier {value}/{total_files}"
        self.progress_update_signal.emit(progress_percent, detail)
    
    # Slots thread-safe pour gestion des signaux
    def _handle_progress_update(self, progress, detail):
        """Slot thread-safe pour mise à jour progression."""
        self.progress_widget.update_progress(progress, detail)
    
    def _handle_compilation_error(self, error_message):
        """Slot thread-safe pour erreurs de compilation."""
        self.stop_worker_timeout()
        self.compilation_error(error_message)
    
    def _handle_compilation_success(self, message, successful, failed, output_path):
        """Slot thread-safe pour succès de compilation."""
        self.stop_worker_timeout()
        self.compilation_success(message, successful, failed, output_path)
    
    def _handle_progress_finish(self, success, message):
        """Slot thread-safe pour finalisation progression."""
        self.progress_widget.finish_operation(success, message)
    
    def _handle_progress_reset(self):
        """Slot thread-safe pour reset progression."""
        self.progress_widget.reset()
    
    def _handle_button_enable(self, enabled):
        """Slot thread-safe pour activer/désactiver bouton compilation."""
        self.button_compile.setEnabled(enabled)

    def compilation_error(self, error_message):
        """
        Gère les erreurs de compilation de manière thread-safe.
        
        Args:
            error_message: Message d'erreur
        """
        ThreadSafeGUIHelper.assert_main_thread("compilation_error")
        
        logging.error(f"Erreur de compilation: {error_message}")
        
        # Logger structuré de l'erreur
        if hasattr(self, 'structured_logger') and self.structured_logger:
            self.structured_logger.log_structured("ERROR", "compilation_failed", 
                                                error_message=error_message,
                                                operation_id=getattr(self, '_current_operation_id', None))
        
        # Finaliser le monitoring de performance en cas d'erreur
        if hasattr(self, 'performance_monitor') and self.performance_monitor:
            self.performance_monitor.end_operation("file_compilation", 
                getattr(self, '_current_operation_id', f"file_compilation_{time.time()}"))
        
        # Incrementer les métriques d'erreur
        if MONITORING_AVAILABLE:
            increment_counter("compilation_errors_total")
            log_event("compilation_error", {
                "error_message": error_message,
                "timestamp": datetime.now().isoformat(),
                "operation_id": getattr(self, '_current_operation_id', None)
            })
        
        # Mettre à jour le widget de progression via signal thread-safe
        if "annulée" in error_message.lower():
            self.progress_finish_signal.emit(False, "Compilation annulée")
        else:
            self.progress_finish_signal.emit(False, f"Erreur: {error_message}")
        
        # Réactiver le bouton et masquer la progression après un délai
        QTimer.singleShot(2000, self._reset_ui_after_operation)
    
    def compilation_success(self, message, successful, failed, output_path):
        """
        Gère les succès de compilation de manière thread-safe.
        
        Args:
            message: Message de succès
            successful: Nombre de fichiers compilés avec succès
            failed: Nombre de fichiers en échec
            output_path: Chemin du fichier de sortie
        """
        ThreadSafeGUIHelper.assert_main_thread("compilation_success")
        
        logging.info(f"Compilation terminée: {message}")
        
        # Logger structuré du succès
        if hasattr(self, 'structured_logger') and self.structured_logger:
            self.structured_logger.log_structured("INFO", "compilation_success", 
                                                message=message,
                                                successful_files=successful,
                                                failed_files=failed,
                                                output_path=output_path,
                                                operation_id=getattr(self, '_current_operation_id', None))
        
        # Finaliser le monitoring de performance en cas de succès
        if hasattr(self, 'performance_monitor') and self.performance_monitor:
            self.performance_monitor.end_operation("file_compilation", 
                getattr(self, '_current_operation_id', f"file_compilation_{time.time()}"))
        
        # Incrementer les métriques de succès
        if MONITORING_AVAILABLE:
            increment_counter("compilation_success_total", {"files_processed": successful + failed})
            set_gauge("last_compilation_files_successful", successful)
            set_gauge("last_compilation_files_failed", failed)
        
        # Mettre à jour le widget de progression via signal thread-safe
        if failed > 0:
            self.progress_finish_signal.emit(True, f"Compilation terminée avec {failed} fichier(s) en échec")
        else:
            self.progress_finish_signal.emit(True, f"Compilation réussie: {successful} fichier(s) compilé(s)")
        
        # Afficher le message de succès
        if successful > 0:
            success_message = f"Compilation terminée avec succès!\n\n"
            success_message += f"Fichiers traités: {successful}\n"
            if failed > 0:
                success_message += f"Fichiers en échec: {failed}\n"
            success_message += f"Fichier de sortie: {output_path}"
            
            # Afficher dans une boîte de dialogue
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Icon.Information)
            msg_box.setWindowTitle("Compilation terminée")
            msg_box.setText(success_message)
            msg_box.exec()
        
        # Réactiver le bouton et masquer la progression après un délai
        QTimer.singleShot(3000, self._reset_ui_after_operation)
    
    def cancel_compilation(self):
        """Annule la compilation en cours"""
        self.stop_worker_timeout()
        if self.compilation_worker and self.compilation_worker.isRunning():
            self.compilation_worker.cancel()
            
            # Metrics de monitoring pour l'annulation
            if MONITORING_AVAILABLE:
                increment_counter("compilation_cancellations_total")
                log_event("compilation_cancelled", {
                    "timestamp": datetime.now().isoformat(),
                    "operation_id": getattr(self, '_current_operation_id', None)
                })
            
            # Logger structuré
            if hasattr(self, 'structured_logger') and self.structured_logger:
                self.structured_logger.log_structured("INFO", "compilation_cancelled", 
                                                    operation_id=getattr(self, '_current_operation_id', None))
            
            logging.info("Demande d'annulation de compilation envoyée")
        else:
            # Si pas de worker en cours, réinitialiser l'interface
            self._reset_ui_after_operation()
    
    def _reset_ui_after_operation(self):
        """Remet l'interface à l'état initial après une opération"""
        ThreadSafeGUIHelper.assert_main_thread("_reset_ui_after_operation")
        
        self.button_enable_signal.emit(True)
        self.exit_compact_mode()  # Sortir du mode compact
        self.progress_reset_signal.emit()
    
    def _handle_progress_detail(self, value: int, detail: str):
        """Handler thread-safe pour les détails de progression"""
        ThreadSafeGUIHelper.assert_main_thread("_handle_progress_detail")
        self.progress_update_signal.emit(value, detail)
    
    def _handle_warning(self, message: str):
        """Handler thread-safe pour les avertissements"""
        ThreadSafeGUIHelper.assert_main_thread("_handle_warning")
        logging.warning(f"Avertissement compilation: {message}")
        # Optionnel: afficher dans l'interface
    
    def _handle_info(self, message: str):
        """Handler thread-safe pour les informations"""
        ThreadSafeGUIHelper.assert_main_thread("_handle_info")
        logging.info(f"Info compilation: {message}")
        # Optionnel: afficher dans l'interface
    
    def _handle_status_update(self, status: str):
        """Handler thread-safe pour les mises à jour de statut avec messages améliorés"""
        ThreadSafeGUIHelper.assert_main_thread("_handle_status_update")
        
        # Utiliser le widget de progression amélioré
        if hasattr(self, 'progress_widget'):
            self.progress_widget.set_status(status)
        
        # Fallback pour compatibilité
        if hasattr(self, 'status_label'):
            self.status_label.setText(status)
    
    def start_worker_timeout(self):
        """Démarre le timer de timeout pour le worker de compilation"""
        if self.timeout_timer:
            self.timeout_timer.stop()
            
        self.timeout_timer = QTimer()
        self.timeout_timer.setSingleShot(True)
        self.timeout_timer.timeout.connect(self.handle_worker_timeout)
        self.timeout_timer.start(self.worker_timeout * 1000)  # Convertir en ms
        
        logging.info(f"Timer de timeout démarré: {self.worker_timeout} secondes")
    
    def stop_worker_timeout(self):
        """Arrête le timer de timeout"""
        if self.timeout_timer:
            self.timeout_timer.stop()
            self.timeout_timer = None
    
    def handle_worker_timeout(self):
        """Gère le timeout du worker"""
        logging.warning("Timeout du worker de compilation détecté")
        
        # Afficher dialogue utilisateur
        reply = QMessageBox.question(
            self,
            self.translate("timeout_title"),
            self.translate("timeout_message", self.worker_timeout // 60),
            QMessageBox.StandardButton.Retry | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Retry
        )
        
        if reply == QMessageBox.StandardButton.Retry:
            # Prolonger le timeout
            self.worker_timeout += 300  # Ajouter 5 minutes
            self.start_worker_timeout()
            logging.info(f"Timeout prolongé de 5 minutes (nouveau timeout: {self.worker_timeout}s)")
        else:
            # Arrêter le worker
            self.cancel_compilation()
            logging.info("Compilation annulée suite au timeout")
    
    def setup_real_time_validation(self):
        """Configure la validation en temps réel pour les champs critiques"""
        try:
            # Validation du nom de fichier de sortie
            if hasattr(self, 'lineedit_output_name') and hasattr(self, 'output_indicator'):
                self.real_time_validator.add_field_validator(
                    "Nom de fichier de sortie", 
                    self.lineedit_output_name, 
                    validate_output_filename,
                    self.output_indicator
                )
            
            # Validation des lignes d'en-tête
            if hasattr(self, 'spinbox_header_start'):
                self.real_time_validator.add_field_validator(
                    "Ligne de début", 
                    self.spinbox_header_start, 
                    validate_header_start_row
                )
            
            if hasattr(self, 'spinbox_header'):
                self.real_time_validator.add_field_validator(
                    "Nombre d'en-têtes", 
                    self.spinbox_header, 
                    validate_header_rows
                )
            
            # Validation de la colonne de tri
            if hasattr(self, 'lineedit_sort_column'):
                self.real_time_validator.add_field_validator(
                    "Colonne de tri", 
                    self.lineedit_sort_column, 
                    validate_sort_column
                )
            
            logging.info("Validation en temps réel configurée")
            
        except Exception as e:
            logging.error(f"Erreur lors de la configuration de la validation: {e}")
    
    def on_validation_changed(self, field_name: str, is_valid: bool, message: str):
        """Gestionnaire appelé quand l'état de validation d'un champ change"""
        # Mettre à jour l'état du bouton de compilation
        self.update_compile_button_state()
        
        # Log des changements de validation (pour le debug)
        if not is_valid and message:
            logging.debug(f"Validation échouée pour {field_name}: {message}")
    
    def update_compile_button_state(self):
        """Met à jour l'état du bouton de compilation selon la validation"""
        if not hasattr(self, 'button_compile') or not hasattr(self, 'real_time_validator'):
            return
            
        # Vérifier si tous les champs sont valides
        all_valid = self.real_time_validator.is_all_valid()
        
        # Vérifier aussi si on a des fichiers sélectionnés et un répertoire
        has_files = (hasattr(self, 'list_files') and 
                    len([item for item in self.list_files.selectedItems()]) > 0)
        has_directory = bool(self.directory)
        
        # Le bouton est activé seulement si tout est valide
        can_compile = all_valid and has_files and has_directory
        
        self.button_compile.setEnabled(can_compile)
        
        # Mettre à jour le tooltip avec les erreurs éventuelles
        if not can_compile:
            issues = []
            if not all_valid:
                issues.append("Erreurs de validation dans les champs")
            if not has_directory:
                issues.append("Aucun répertoire sélectionné")
            if not has_files:
                issues.append("Aucun fichier sélectionné")
            
            tooltip = "Impossible de compiler:\n" + "\n".join(f"• {issue}" for issue in issues)
            if not all_valid:
                tooltip += "\n\nErreurs de validation:\n" + self.real_time_validator.get_validation_summary()
            
            self.button_compile.setToolTip(tooltip)
        else:
            self.button_compile.setToolTip("Lancer la compilation")
    
    def compilation_finished(self, result):
        """
        Méthode appelée à la fin de la compilation de manière thread-safe.
        
        Args:
            result: Tuple contenant les résultats de la compilation
        """
        ThreadSafeGUIHelper.assert_main_thread("compilation_finished")
        preliminary_info, headers, combined_data, merged_cells, successful_files, failed_files = result
        
        # Réactiver l'interface et sortir du mode compact
        self.button_enable_signal.emit(True)
        
        # Mettre à jour le widget de progression
        success_message = f"Compilation terminée: {len(successful_files)} fichier(s) traité(s)"
        if failed_files:
            success_message += f", {len(failed_files)} échec(s)"
        self.progress_finish_signal.emit(True, success_message)
        
        # Sortir du mode compact après un petit délai pour permettre à l'utilisateur de voir le message
        QTimer.singleShot(2000, self.exit_compact_mode)
        
        if not combined_data or not headers:
            # Afficher plus d'informations de debug
            debug_info = f"Headers: {len(headers) if headers else 0}, Data: {len(combined_data) if combined_data else 0}"
            debug_info += f", Successful files: {len(successful_files)}, Failed files: {len(failed_files)}"
            
            logging.warning(f"Aucune donnée compilée - Debug: {debug_info}")
            
            self.status_label.setText(self.translate("no_data"))
            QMessageBox.warning(
                self,
                self.translate("warning"),
                f"{self.translate('no_data_message')}\n\nDébug: {debug_info}"
            )
            return
        
        # Générer le nom du fichier de sortie
        output_filename = self.lineedit_output_name.text().strip()
        if not output_filename:
            output_filename = "compilation"
        
        if not output_filename.endswith('.xlsx'):
            output_filename += '.xlsx'
        
        output_path = os.path.join(self.directory, output_filename)
        
        try:
            # Écrire les données dans le fichier Excel
            self.write_excel_file(
                output_path,
                preliminary_info,
                headers,
                combined_data,
                merged_cells
            )
            
            # Afficher le rapport de compilation
            self.show_compilation_report(successful_files, failed_files, output_path)
            
            self.status_label.setText(self.translate("compilation_complete", output_path))
            logging.info(f"Compilation terminée: {output_path}")
            
            # Logger structuré de fin de compilation réussie
            if hasattr(self, 'structured_logger') and self.structured_logger:
                self.structured_logger.log_operation(
                    "file_compilation", 
                    "success",
                    successful_files=len(successful_files),
                    failed_files=len(failed_files),
                    output_file=output_path
                )
            
            # Finaliser le monitoring de performance
            if hasattr(self, 'performance_monitor') and self.performance_monitor:
                self.performance_monitor.end_operation("file_compilation", 
                    getattr(self, '_current_operation_id', f"file_compilation_{time.time()}"))
            
        except PermissionError:
            self.status_label.setText(self.translate("file_open_error"))
            QMessageBox.warning(
                self,
                self.translate("error"),
                self.translate("file_open_error_message")
            )
        except Exception as e:
            error_msg = str(e)
            self.status_label.setText(self.translate("compilation_failed", error_msg))
            QMessageBox.critical(
                self,
                self.translate("error"),
                self.translate("compilation_failed", error_msg)
            )
            logging.error(f"Erreur lors de l'écriture: {error_msg}")
    
    def write_excel_file(self, output_path, preliminary_info, headers, data, merged_cells):
        """
        Écrit les données compilées dans un fichier Excel.
        
        Args:
            output_path: Chemin du fichier de sortie
            preliminary_info: Informations préliminaires
            headers: En-têtes
            data: Données
            merged_cells: Cellules fusionnées
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Compilation"
        
        current_row = 1
        
        # Écrire les informations préliminaires
        if preliminary_info:
            current_row = ExcelFormatter.write_preliminary_info(ws, preliminary_info)
        
        # Écrire les en-têtes
        if headers:
            header_start_row = current_row
            current_row = ExcelFormatter.write_headers(ws, headers, current_row)
            
            # Appliquer les cellules fusionnées
            if merged_cells:
                ExcelFormatter.apply_merged_cells(
                    ws, merged_cells, 
                    self.spinbox_header_start.value(), 
                    header_start_row
                )
        
        # Écrire les données
        if data:
            ExcelFormatter.write_data(ws, data, current_row, self.date_format)
        
        # Appliquer le formatage
        if self.checkbox_auto_width.isChecked():
            ExcelFormatter.adjust_column_widths(ws)
        
        if self.checkbox_freeze_header.isChecked() and headers:
            freeze_row = header_start_row + len(headers) if preliminary_info else len(headers) + 1
            ExcelFormatter.freeze_header(ws, freeze_row)
        
        # Sauvegarder le fichier
        wb.save(output_path)
    
    def show_compilation_report(self, successful_files, failed_files, output_path):
        """
        Affiche le rapport de compilation.
        
        Args:
            successful_files: Liste des fichiers compilés avec succès
            failed_files: Liste des fichiers échoués
            output_path: Chemin du fichier de sorti
        """
        dialog = CompilationReportDialog(self, successful_files, failed_files, output_path)
        dialog.exec()
    
    def closeEvent(self, event):
        """
        Gère la fermeture de l'application.
        
        Args:
            event: Événement de fermeture
        """
        # Sauvegarder automatiquement les paramètres
        SettingsManager().save_settings(self)
        
        # Sauvegarder la session actuelle
        if hasattr(self, 'auto_save_manager') and self.auto_save_manager:
            self.save_current_session()
        
        # Sauvegarder la configuration
        if hasattr(self, 'config_manager') and self.config_manager:
            self.config_manager.save_configuration()
        
        # Arrêter la surveillance
        if hasattr(self, 'health_monitor') and self.health_monitor:
            self.health_monitor.stop_monitoring()
        
        # Logger la fermeture de l'application
        if hasattr(self, 'structured_logger') and self.structured_logger:
            self.structured_logger.log_structured("INFO", "application_shutdown")
        
        # Arrêter le monitoring global
        if MONITORING_AVAILABLE:
            log_event("application_shutdown", {"timestamp": datetime.now().isoformat()})
            stop_monitoring()
        
        # Nettoyer les ressources(arrêter le worker, fermer les fichiers, etc.)
        self.cleanup_resources()
        
        logging.info("Application fermée")
        event.accept()


# =====================================================
# TESTS UNITAIRES
# =====================================================

class TestExcelCompiler(unittest.TestCase):
    """Tests unitaires pour l'application."""
    
    def setUp(self):
        """Configuration des tests."""
        self.app = QApplication([])
        self.compiler = ModernExcelCompilerApp()
    
    def tearDown(self):
        """Nettoyage après les tests."""
        self.compiler.close()
        self.app.quit()
    
    def test_translation_manager(self):
        """Test du gestionnaire de traduction."""
        tm = TranslationManager()
        
        # Test changement de langue
        tm.set_language("en")
        self.assertEqual(tm.current_language, "en")
        self.assertEqual(tm.get_text("app_title"), "Professional Excel Compiler")
        
        tm.set_language("fr")
        self.assertEqual(tm.current_language, "fr")
        self.assertEqual(tm.get_text("app_title"), "Compilateur Excel Professionnel")
    
    def test_file_verification(self):
        """Test de la vérification des fichiers."""
        # Créer un fichier de test temporaire
        import tempfile
        
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_file:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws['A1'] = 'Test Header'
            ws['A2'] = 'Test Data'
            wb.save(temp_file.name)
            
            # Test de vérification
            is_compatible, reason = FileVerification.verify_excel_file(
                temp_file.name, 1, 1
            )
            self.assertTrue(is_compatible)
            
        # Nettoyer
        os.unlink(temp_file.name)
    
    def test_date_formats(self):
        """Test des formats de date."""
        self.assertIn("STANDARD", DATE_FORMATS)
        self.assertIn("FRENCH", DATE_FORMATS)
        self.assertIn("format", DATE_FORMATS["FRENCH"])  
        self.assertIn("excel_format", DATE_FORMATS["FRENCH"])
    
    def test_sort_column_conversion(self):
        """Test de la conversion des colonnes de tri."""
        self.compiler.lineedit_sort_column.setText("A")
        self.assertEqual(self.compiler.get_sort_column_index(), 0)
        
        self.compiler.lineedit_sort_column.setText("B")
        self.assertEqual(self.compiler.get_sort_column_index(), 1)
        
        self.compiler.lineedit_sort_column.setText("1")
        self.assertEqual(self.compiler.get_sort_column_index(), 0)


# =====================================================
# FONCTION PRINCIPALE
# =====================================================

def main():
    """Fonction principal pour lancer l'application."""
    setup_logging()
    try:
        # Configuration de l'application
        app = QApplication(sys.argv)
        app.setApplicationName("Excel Compiler")
        app.setApplicationVersion("3.0")
        app.setOrganizationName("GOUNOU N'GOBI")
        app.setOrganizationDomain("zimkada@gmail.com")
        
        # Définir l'icône de l'application si disponible
        if os.path.exists("icon.ico"):
            app.setWindowIcon(QIcon("icon.ico"))
        
        # Style de l'application
        app.setStyle('Fusion')
        
        # Palette de couleurs personnalisée
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor(240, 240, 240))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(0, 0, 0))
        palette.setColor(QPalette.ColorRole.Base, QColor(255, 255, 255))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(245, 245, 245))
        palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(255, 255, 220))
        palette.setColor(QPalette.ColorRole.ToolTipText, QColor(0, 0, 0))
        palette.setColor(QPalette.ColorRole.Text, QColor(0, 0, 0))
        palette.setColor(QPalette.ColorRole.Button, QColor(240, 240, 240))
        palette.setColor(QPalette.ColorRole.ButtonText, QColor(0, 0, 0))
        palette.setColor(QPalette.ColorRole.BrightText, QColor(255, 0, 0))
        palette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
        palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
        palette.setColor(QPalette.ColorRole.HighlightedText, QColor(0, 0, 0))
        app.setPalette(palette)
        
        # Créer et afficher la fenêtre principale
        compiler = ModernExcelCompilerApp()
        compiler.show()
        
        # Configuration du logging pour capturer les erreurs non gérées
        def handle_exception(exc_type, exc_value, exc_traceback):
            if issubclass(exc_type, KeyboardInterrupt):
                sys.__excepthook__(exc_type, exc_value, exc_traceback)
                return
            
            logging.critical("Exception non gérée", exc_info=(exc_type, exc_value, exc_traceback))
            QMessageBox.critical(
                None,
                "Erreur Critique",
                f"Une erreur inattendue s'est produite:\n{exc_type.__name__}: {exc_value}"
            )
        
        sys.excepthook = handle_exception
        
        # Démarrer la boucle d'événements
        sys.exit(app.exec())
        
    except Exception as e:
        logging.critical(f"Erreur critique au démarrage: {str(e)}")
        logging.critical(traceback.format_exc())
        
        # Créer une application minimale pour afficher l'erreur
        if 'app' not in locals():
            app = QApplication(sys.argv)
        
        QMessageBox.critical(
            None,
            "Erreur de Démarrage",
            f"Impossible de démarrer l'application:\n{str(e)}\n\nConsultez le fichier de log pour plus de détails."
        )
        sys.exit(1)


# ========================================
# TESTS UNITAIRES
# ========================================

class TestThreadSafeGUIHelper(unittest.TestCase):
    """Tests pour la classe ThreadSafeGUIHelper"""
    
    def setUp(self):
        """Configuration pour chaque test"""
        self.app = QApplication.instance()
        if self.app is None:
            self.app = QApplication([])
    
    def test_is_main_thread(self):
        """Test de détection du thread principal"""
        # Dans le thread principal, doit retourner True
        self.assertTrue(ThreadSafeGUIHelper.is_main_thread())
    
    def test_assert_main_thread_success(self):
        """Test assertion thread principal - succès"""
        try:
            ThreadSafeGUIHelper.assert_main_thread("test_method")
        except RuntimeError:
            self.fail("assert_main_thread() a échoué dans le thread principal")
    
    def test_invoke_in_main_thread(self):
        """Test d'invocation dans le thread principal"""
        # Créer un mock objet avec une méthode
        mock_obj = Mock()
        mock_obj.test_method = Mock(return_value="success")
        
        # Appel dans le thread principal
        result = ThreadSafeGUIHelper.invoke_in_main_thread(mock_obj, "test_method", "arg1", "arg2")
        
        # Vérifier l'appel direct
        mock_obj.test_method.assert_called_once_with("arg1", "arg2")


class TestMemoryManager(unittest.TestCase):
    """Tests pour la classe MemoryManager"""
    
    def setUp(self):
        """Configuration pour chaque test"""
        self.memory_manager = MemoryManager()
    
    def test_get_memory_usage(self):
        """Test de récupération de l'usage mémoire"""
        memory_usage = self.memory_manager.get_memory_usage()
        self.assertIsInstance(memory_usage, int)
        self.assertGreater(memory_usage, 0)
    
    def test_should_cleanup_memory(self):
        """Test de décision de nettoyage mémoire"""
        # Test avec usage mémoire artificiel faible
        with patch('psutil.Process') as mock_process:
            mock_process.return_value.memory_info.return_value.rss = 100 * 1024 * 1024  # 100MB
            result = self.memory_manager.should_cleanup_memory()
            self.assertFalse(result)
            
        # Test avec usage mémoire artificiel élevé
        with patch('psutil.Process') as mock_process:
            mock_process.return_value.memory_info.return_value.rss = 600 * 1024 * 1024  # 600MB
            result = self.memory_manager.should_cleanup_memory()
            self.assertTrue(result)
    
    def test_cleanup_memory(self):
        """Test de nettoyage mémoire"""
        # Le nettoyage ne doit pas lever d'exception
        try:
            self.memory_manager.cleanup_memory()
        except Exception as e:
            self.fail(f"cleanup_memory() a échoué: {e}")


class TestFileMetadataCache(unittest.TestCase):
    """Tests pour la classe FileMetadataCache"""
    
    def setUp(self):
        """Configuration pour chaque test"""
        self.cache = FileMetadataCache()
        self.test_file = tempfile.NamedTemporaryFile(delete=False, suffix='.txt')
        self.test_file.write(b"test content")
        self.test_file.close()
    
    def tearDown(self):
        """Nettoyage après chaque test"""
        try:
            os.unlink(self.test_file.name)
        except (OSError, FileNotFoundError):
            # Fichier déjà supprimé ou inaccessible
            pass
    
    def test_get_file_info_new_file(self):
        """Test récupération info fichier non mis en cache"""
        info = self.cache.get_file_info(self.test_file.name)
        
        self.assertIsNotNone(info)
        self.assertIn('size', info)
        self.assertIn('modified', info)
        self.assertIn('extension', info)
        self.assertEqual(info['extension'], '.txt')
    
    def test_get_file_info_cached_file(self):
        """Test récupération info fichier mis en cache"""
        # Premier appel - mise en cache
        info1 = self.cache.get_file_info(self.test_file.name)
        
        # Deuxième appel - récupération du cache
        info2 = self.cache.get_file_info(self.test_file.name)
        
        self.assertEqual(info1, info2)
    
    def test_get_file_info_nonexistent_file(self):
        """Test récupération info fichier inexistant"""
        info = self.cache.get_file_info("/nonexistent/file.txt")
        self.assertIsNone(info)


class TestCancellationToken(unittest.TestCase):
    """Tests pour la classe CancellationToken"""
    
    def setUp(self):
        """Configuration pour chaque test"""
        self.token = CancellationToken()
    
    def test_initial_state(self):
        """Test état initial du token"""
        self.assertFalse(self.token.is_cancelled())
    
    def test_cancel(self):
        """Test annulation du token"""
        self.token.cancel()
        self.assertTrue(self.token.is_cancelled())
    
    def test_reset(self):
        """Test remise à zéro du token"""
        self.token.cancel()
        self.assertTrue(self.token.is_cancelled())
        
        self.token.reset()
        self.assertFalse(self.token.is_cancelled())
    
    def test_throw_if_cancelled(self):
        """Test exception si annulé"""
        # Ne doit pas lever d'exception
        try:
            self.token.throw_if_cancelled()
        except Exception as e:
            self.fail(f"throw_if_cancelled() a échoué sans annulation: {e}")
        
        # Doit lever une exception après annulation
        self.token.cancel()
        with self.assertRaises(Exception):
            self.token.throw_if_cancelled()


class TestAdvancedFileDetector(unittest.TestCase):
    """Tests pour la classe AdvancedFileDetector"""
    
    def setUp(self):
        """Configuration pour chaque test"""
        self.detector = AdvancedFileDetector()
        self.temp_dir = tempfile.mkdtemp()
    
    def tearDown(self):
        """Nettoyage après chaque test"""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_detect_csv_encoding_utf8(self):
        """Test détection encodage CSV UTF-8"""
        test_file = os.path.join(self.temp_dir, "test_utf8.csv")
        with open(test_file, 'w', encoding='utf-8') as f:
            f.write("col1,col2,col3\nvalue1,value2,value3\n")
        
        encoding = self.detector.detect_csv_encoding(test_file)
        self.assertIn(encoding.lower(), ['utf-8', 'utf-8-sig'])
    
    def test_detect_csv_encoding_latin1(self):
        """Test détection encodage CSV Latin-1"""
        test_file = os.path.join(self.temp_dir, "test_latin1.csv")
        with open(test_file, 'w', encoding='latin-1') as f:
            f.write("col1,col2,col3\nvalué1,valué2,valué3\n")
        
        encoding = self.detector.detect_csv_encoding(test_file)
        self.assertIsInstance(encoding, str)
    
    def test_detect_csv_delimiter_comma(self):
        """Test détection délimiteur CSV virgule"""
        test_file = os.path.join(self.temp_dir, "test_comma.csv")
        with open(test_file, 'w', encoding='utf-8') as f:
            f.write("col1,col2,col3\nvalue1,value2,value3\n")
        
        delimiter = self.detector.detect_csv_delimiter(test_file)
        self.assertEqual(delimiter, ',')
    
    def test_detect_csv_delimiter_semicolon(self):
        """Test détection délimiteur CSV point-virgule"""
        test_file = os.path.join(self.temp_dir, "test_semicolon.csv")
        with open(test_file, 'w', encoding='utf-8') as f:
            f.write("col1;col2;col3\nvalue1;value2;value3\n")
        
        delimiter = self.detector.detect_csv_delimiter(test_file)
        self.assertEqual(delimiter, ';')
    
    def test_detect_csv_delimiter_tab(self):
        """Test détection délimiteur CSV tabulation"""
        test_file = os.path.join(self.temp_dir, "test_tab.tsv")
        with open(test_file, 'w', encoding='utf-8') as f:
            f.write("col1\tcol2\tcol3\nvalue1\tvalue2\tvalue3\n")
        
        delimiter = self.detector.detect_csv_delimiter(test_file)
        self.assertEqual(delimiter, '\t')


class TestValidationIndicator(unittest.TestCase):
    """Tests pour la classe ValidationIndicator"""
    
    def setUp(self):
        """Configuration pour chaque test"""
        self.app = QApplication.instance()
        if self.app is None:
            self.app = QApplication([])
        self.validator = ValidationIndicator()
    
    def test_initial_state(self):
        """Test état initial du validateur"""
        self.assertTrue(self.validator.is_valid)
        self.assertEqual(self.validator.message, "")
    
    def test_set_valid(self):
        """Test définition état valide"""
        self.validator.set_valid("test_field", True, "Valid input")
        self.assertTrue(self.validator.is_valid)
        self.assertEqual(self.validator.message, "Valid input")
    
    def test_set_invalid(self):
        """Test définition état invalide"""
        self.validator.set_valid("test_field", False, "Invalid input")
        self.assertFalse(self.validator.is_valid)
        self.assertEqual(self.validator.message, "Invalid input")


class TestResponsiveManager(unittest.TestCase):
    """Tests pour la classe ResponsiveManager"""
    
    def setUp(self):
        """Configuration pour chaque test"""
        self.app = QApplication.instance()
        if self.app is None:
            self.app = QApplication([])
        self.manager = ResponsiveManager()
    
    def test_initial_state(self):
        """Test état initial du gestionnaire responsive"""
        self.assertIsNone(self.manager.screen_size)
        self.assertEqual(self.manager.scale_factor, 1.0)
        self.assertEqual(self.manager.font_scale, 1.0)
        self.assertEqual(self.manager.current_breakpoint, 'medium')
    
    def test_get_breakpoint_small(self):
        """Test détection breakpoint petit écran"""
        breakpoint = self.manager.get_breakpoint(800)
        self.assertEqual(breakpoint, 'small')
    
    def test_get_breakpoint_medium(self):
        """Test détection breakpoint écran moyen"""
        breakpoint = self.manager.get_breakpoint(1200)
        self.assertEqual(breakpoint, 'medium')
    
    def test_get_breakpoint_large(self):
        """Test détection breakpoint grand écran"""
        breakpoint = self.manager.get_breakpoint(1600)
        self.assertEqual(breakpoint, 'large')
    
    def test_get_breakpoint_xlarge(self):
        """Test détection breakpoint très grand écran"""
        breakpoint = self.manager.get_breakpoint(3000)
        self.assertEqual(breakpoint, 'xlarge')


class TestCancellableProgressWidget(unittest.TestCase):
    """Tests pour la classe CancellableProgressWidget"""
    
    def setUp(self):
        """Configuration pour chaque test"""
        self.app = QApplication.instance()
        if self.app is None:
            self.app = QApplication([])
        self.widget = CancellableProgressWidget()
    
    def test_initial_state(self):
        """Test état initial du widget"""
        self.assertFalse(self.widget.is_cancelling)
        self.assertEqual(self.widget.last_progress, 0)
    
    def test_start_operation(self):
        """Test démarrage d'opération"""
        self.widget.start_operation("Test operation")
        self.assertFalse(self.widget.is_cancelling)
        self.assertIsNotNone(self.widget.start_time)
    
    def test_update_progress(self):
        """Test mise à jour progression"""
        self.widget.start_operation("Test operation")
        self.widget.update_progress(50, "Test detail")
        self.assertEqual(self.widget.last_progress, 50)
    
    def test_finish_operation_success(self):
        """Test fin d'opération avec succès"""
        self.widget.start_operation("Test operation")
        self.widget.finish_operation(True, "Success")
        # Vérifier que l'opération est terminée
        self.assertIsNotNone(self.widget.start_time)
    
    def test_finish_operation_failure(self):
        """Test fin d'opération avec échec"""
        self.widget.start_operation("Test operation")
        self.widget.finish_operation(False, "Failed")
        # Vérifier que l'opération est terminée
        self.assertIsNotNone(self.widget.start_time)
    
    def test_cancel_operation(self):
        """Test annulation d'opération"""
        self.widget.start_operation("Test operation")
        self.widget.cancel_operation()
        self.assertTrue(self.widget.is_cancelling)


class TestCompilationWorker(unittest.TestCase):
    """Tests pour la classe CompilationWorker (critique)"""
    
    def setUp(self):
        """Configuration pour chaque test"""
        self.app = QApplication.instance()
        if self.app is None:
            self.app = QApplication([])
        
        # Créer un dossier temporaire avec fichiers de test
        self.test_dir = tempfile.mkdtemp()
        self.test_files = []
        
        # Créer des fichiers CSV de test
        for i in range(3):
            file_path = os.path.join(self.test_dir, f"test_{i}.csv")
            with open(file_path, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(['Col1', 'Col2', 'Col3'])  # En-têtes
                for row in range(5):
                    writer.writerow([f'Data{i}_{row}_1', f'Data{i}_{row}_2', f'Data{i}_{row}_3'])
            self.test_files.append(f"test_{i}.csv")
    
    def tearDown(self):
        """Nettoyage après chaque test"""
        shutil.rmtree(self.test_dir, ignore_errors=True)
    
    def test_worker_initialization(self):
        """Test d'initialisation du worker"""
        worker = CompilationWorker(
            files=self.test_files,
            directory=self.test_dir,
            header_start_row=1,
            header_rows=1
        )
        
        self.assertEqual(worker.files, self.test_files)
        self.assertEqual(worker.directory, self.test_dir)
        self.assertEqual(worker.header_start_row, 1)
        self.assertEqual(worker.header_rows, 1)
    
    def test_worker_invalid_parameters(self):
        """Test validation des paramètres invalides"""
        # Test liste de fichiers vide
        with self.assertRaises(ValueError):
            CompilationWorker([], self.test_dir, 1, 1)
        
        # Test répertoire inexistant
        with self.assertRaises(ValueError):
            CompilationWorker(self.test_files, "/nonexistent", 1, 1)
        
        # Test ligne de début invalide
        with self.assertRaises(ValueError):
            CompilationWorker(self.test_files, self.test_dir, 0, 1)
    
    def test_worker_signals(self):
        """Test des signaux du worker"""
        worker = CompilationWorker(
            files=self.test_files,
            directory=self.test_dir,
            header_start_row=1,
            header_rows=1
        )
        
        # Vérifier que les signaux existent
        self.assertTrue(hasattr(worker, 'progress'))
        self.assertTrue(hasattr(worker, 'error'))
        self.assertTrue(hasattr(worker, 'finished'))
        self.assertTrue(hasattr(worker, 'progress_detail'))
        self.assertTrue(hasattr(worker, 'warning'))
    
    def test_worker_cancellation(self):
        """Test d'annulation du worker"""
        worker = CompilationWorker(
            files=self.test_files,
            directory=self.test_dir,
            header_start_row=1,
            header_rows=1
        )
        
        # Test méthode cancel
        worker.cancel()
        self.assertTrue(worker.cancellation_token.is_cancelled())


class TestErrorRecoveryManager(unittest.TestCase):
    """Tests pour la classe ErrorRecoveryManager"""
    
    def setUp(self):
        """Configuration pour chaque test"""
        self.recovery_manager = ErrorRecoveryManager()
    
    def test_add_error(self):
        """Test d'ajout d'erreur"""
        error_msg = "Test error message"
        self.recovery_manager.add_error("test_file.csv", error_msg)
        
        errors = self.recovery_manager.get_errors()
        self.assertEqual(len(errors), 1)
        self.assertIn("test_file.csv", errors)
        self.assertEqual(errors["test_file.csv"], error_msg)
    
    def test_should_continue_after_error(self):
        """Test de décision de continuation après erreur"""
        # Par défaut, doit continuer
        self.assertTrue(self.recovery_manager.should_continue_after_error("test.csv", "error"))
        
        # Après plusieurs erreurs sur le même fichier
        for i in range(5):
            self.recovery_manager.add_error("test.csv", f"error {i}")
        
        # Doit encore continuer (robuste)
        self.assertTrue(self.recovery_manager.should_continue_after_error("test.csv", "error"))
    
    def test_get_recovery_suggestions(self):
        """Test des suggestions de récupération"""
        suggestions = self.recovery_manager.get_recovery_suggestions("File not found")
        self.assertIsInstance(suggestions, list)
        self.assertGreater(len(suggestions), 0)
    
    def test_clear_errors(self):
        """Test de nettoyage des erreurs"""
        self.recovery_manager.add_error("test1.csv", "error1")
        self.recovery_manager.add_error("test2.csv", "error2")
        
        self.assertEqual(len(self.recovery_manager.get_errors()), 2)
        
        self.recovery_manager.clear_errors()
        self.assertEqual(len(self.recovery_manager.get_errors()), 0)


class TestChunkedFileReader(unittest.TestCase):
    """Tests pour la classe ChunkedFileReader"""
    
    def setUp(self):
        """Configuration pour chaque test"""
        self.test_dir = tempfile.mkdtemp()
        
        # Créer un gros fichier CSV de test
        self.large_file = os.path.join(self.test_dir, "large_test.csv")
        with open(self.large_file, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Col1', 'Col2', 'Col3'])  # En-têtes
            for i in range(1000):  # 1000 lignes de données
                writer.writerow([f'Data_{i}_1', f'Data_{i}_2', f'Data_{i}_3'])
    
    def tearDown(self):
        """Nettoyage après chaque test"""
        shutil.rmtree(self.test_dir, ignore_errors=True)
    
    def test_chunked_reading(self):
        """Test de lecture par chunks"""
        reader = ChunkedFileReader(chunk_size=100)
        
        total_rows = 0
        for chunk in reader.read_file(self.large_file):
            self.assertIsInstance(chunk, pd.DataFrame)
            self.assertLessEqual(len(chunk), 100)  # Taille de chunk respectée
            total_rows += len(chunk)
        
        # Vérifier que toutes les lignes ont été lues (1000 + 1 en-tête)
        self.assertGreaterEqual(total_rows, 1000)
    
    def test_empty_file(self):
        """Test avec fichier vide"""
        empty_file = os.path.join(self.test_dir, "empty.csv")
        with open(empty_file, 'w', encoding='utf-8') as f:
            pass  # Fichier vide
        
        reader = ChunkedFileReader()
        chunks = list(reader.read_file(empty_file))
        
        # Doit gérer gracieusement les fichiers vides
        self.assertIsInstance(chunks, list)


class TestInputValidator(unittest.TestCase):
    """Tests pour la classe InputValidator"""
    
    def setUp(self):
        """Configuration pour chaque test"""
        self.validator = InputValidator()
    
    def test_validate_positive_integer(self):
        """Test validation entier positif"""
        # Valeurs valides
        self.assertTrue(self.validator.is_valid_positive_integer("1"))
        self.assertTrue(self.validator.is_valid_positive_integer("100"))
        
        # Valeurs invalides
        self.assertFalse(self.validator.is_valid_positive_integer("0"))
        self.assertFalse(self.validator.is_valid_positive_integer("-1"))
        self.assertFalse(self.validator.is_valid_positive_integer("abc"))
        self.assertFalse(self.validator.is_valid_positive_integer(""))
    
    def test_validate_file_list(self):
        """Test validation liste de fichiers"""
        # Créer des fichiers temporaires
        temp_dir = tempfile.mkdtemp()
        try:
            file1 = os.path.join(temp_dir, "test1.csv")
            file2 = os.path.join(temp_dir, "test2.xlsx")
            
            # Créer les fichiers
            with open(file1, 'w') as f:
                f.write("test")
            with open(file2, 'w') as f:
                f.write("test")
            
            # Test validation
            valid, message = self.validator.validate_file_list([file1, file2], temp_dir)
            self.assertTrue(valid)
            
            # Test avec fichier inexistant
            invalid_file = os.path.join(temp_dir, "nonexistent.csv")
            valid, message = self.validator.validate_file_list([invalid_file], temp_dir)
            self.assertFalse(valid)
            
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
    
    def test_validate_column_reference(self):
        """Test validation référence de colonne"""
        # Références valides
        self.assertTrue(self.validator.is_valid_column_reference("A"))
        self.assertTrue(self.validator.is_valid_column_reference("Z"))
        self.assertTrue(self.validator.is_valid_column_reference("AA"))
        self.assertTrue(self.validator.is_valid_column_reference("1"))
        self.assertTrue(self.validator.is_valid_column_reference("10"))
        
        # Références invalides
        self.assertFalse(self.validator.is_valid_column_reference(""))
        self.assertFalse(self.validator.is_valid_column_reference("0"))
        self.assertFalse(self.validator.is_valid_column_reference("-1"))
        self.assertFalse(self.validator.is_valid_column_reference("@"))


class TestIntegrationEndToEnd(unittest.TestCase):
    """Tests d'intégration end-to-end"""
    
    def setUp(self):
        """Configuration pour tests d'intégration"""
        self.app = QApplication.instance()
        if self.app is None:
            self.app = QApplication([])
        
        self.test_dir = tempfile.mkdtemp()
        self.output_dir = tempfile.mkdtemp()
        
        # Créer plusieurs fichiers de test avec structures identiques
        self.test_files = []
        for i in range(3):
            file_path = os.path.join(self.test_dir, f"integration_test_{i}.csv")
            with open(file_path, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(['Name', 'Age', 'City'])  # En-têtes identiques
                for row in range(10):
                    writer.writerow([f'Person{i}_{row}', 20+row, f'City{i}'])
            self.test_files.append(f"integration_test_{i}.csv")
    
    def tearDown(self):
        """Nettoyage après tests d'intégration"""
        shutil.rmtree(self.test_dir, ignore_errors=True)
        shutil.rmtree(self.output_dir, ignore_errors=True)
    
    def test_simple_compilation_workflow(self):
        """Test workflow complet de compilation simple"""
        # Préparer les paramètres
        files = self.test_files
        directory = self.test_dir
        header_start_row = 1
        header_rows = 1
        
        # Créer et configurer le worker
        worker = CompilationWorker(
            files=files,
            directory=directory,
            header_start_row=header_start_row,
            header_rows=header_rows,
            filename_option="none",
            sort_data=False,
            sort_column=0,
            repeat_headers=False,
            remove_empty_rows=False,
            remove_duplicates=False,
            date_format="FRENCH"
        )
        
        # Variables pour capturer les résultats
        self.compilation_results = None
        self.compilation_error = None
        
        def on_finished(results):
            self.compilation_results = results
        
        def on_error(error):
            self.compilation_error = error
        
        # Connecter les signaux
        worker.finished.connect(on_finished)
        worker.error.connect(on_error)
        
        # Lancer la compilation de façon synchrone pour le test
        worker.run()
        
        # Vérifier les résultats
        self.assertIsNone(self.compilation_error, f"Erreur de compilation: {self.compilation_error}")
        self.assertIsNotNone(self.compilation_results, "Aucun résultat de compilation")
        
        # Décomposer les résultats
        preliminary_info, headers, combined_data, merged_cells, successful_files, failed_files = self.compilation_results
        
        # Vérifications
        self.assertIsInstance(headers, list, "Les en-têtes doivent être une liste")
        self.assertGreater(len(headers), 0, "Doit avoir des en-têtes")
        self.assertIsInstance(combined_data, list, "Les données doivent être une liste")
        self.assertGreater(len(combined_data), 0, "Doit avoir des données")
        self.assertEqual(len(successful_files), 3, "Tous les fichiers doivent être traités")
        self.assertEqual(len(failed_files), 0, "Aucun fichier ne doit échouer")
    
    def test_compilation_with_invalid_files(self):
        """Test compilation avec fichiers invalides"""
        # Ajouter un fichier invalide
        invalid_file = os.path.join(self.test_dir, "invalid.csv")
        with open(invalid_file, 'w', encoding='utf-8') as f:
            f.write("Invalid content without proper CSV structure")
        
        files = self.test_files + ["invalid.csv"]
        
        worker = CompilationWorker(
            files=files,
            directory=self.test_dir,
            header_start_row=1,
            header_rows=1
        )
        
        # Variables pour résultats
        self.compilation_results = None
        self.compilation_error = None
        
        def on_finished(results):
            self.compilation_results = results
        
        def on_error(error):
            self.compilation_error = error
        
        worker.finished.connect(on_finished)
        worker.error.connect(on_error)
        
        # Lancer la compilation
        worker.run()
        
        # La compilation doit continuer malgré le fichier invalide
        if self.compilation_results:
            preliminary_info, headers, combined_data, merged_cells, successful_files, failed_files = self.compilation_results
            
            # Doit avoir réussi pour les fichiers valides
            self.assertGreaterEqual(len(successful_files), 3)
            # Le fichier invalide peut être dans failed_files
            self.assertLessEqual(len(failed_files), 1)


class TestPerformanceBasic(unittest.TestCase):
    """Tests de performance de base"""
    
    def setUp(self):
        """Configuration pour tests de performance"""
        self.test_dir = tempfile.mkdtemp()
    
    def tearDown(self):
        """Nettoyage après tests de performance"""
        shutil.rmtree(self.test_dir, ignore_errors=True)
    
    def test_memory_manager_performance(self):
        """Test performance du gestionnaire de mémoire"""
        memory_manager = MemoryManager()
        
        start_time = time.time()
        
        # Test multiple appels de vérification mémoire
        for i in range(100):
            memory_usage = memory_manager.get_memory_usage()
            should_cleanup = memory_manager.should_cleanup_memory()
        
        end_time = time.time()
        elapsed = end_time - start_time
        
        # Doit être rapide (moins de 1 seconde pour 100 appels)
        self.assertLess(elapsed, 1.0, "Vérifications mémoire trop lentes")
    
    def test_file_cache_performance(self):
        """Test performance du cache de fichiers"""
        cache = FileMetadataCache()
        
        # Créer plusieurs fichiers de test
        test_files = []
        for i in range(10):
            file_path = os.path.join(self.test_dir, f"perf_test_{i}.csv")
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write("col1,col2,col3\ndata1,data2,data3\n")
            test_files.append(file_path)
        
        # Premier accès (mise en cache)
        start_time = time.time()
        for file_path in test_files:
            info = cache.get_file_info(file_path)
        first_access_time = time.time() - start_time
        
        # Deuxième accès (depuis le cache)
        start_time = time.time()
        for file_path in test_files:
            info = cache.get_file_info(file_path)
        cached_access_time = time.time() - start_time
        
        # L'accès en cache doit être plus rapide
        self.assertLess(cached_access_time, first_access_time, 
                       "L'accès en cache doit être plus rapide que le premier accès")


def run_tests():
    """Exécute tous les tests unitaires"""
    print("Exécution des tests unitaires pour ExcelCompiler...")
    
    # Créer une suite de tests
    test_suite = unittest.TestSuite()
    
    # Ajouter les classes de tests
    test_classes = [
        TestThreadSafeGUIHelper,
        TestMemoryManager,
        TestFileMetadataCache,
        TestCancellationToken,
        TestAdvancedFileDetector,
        TestValidationIndicator,
        TestResponsiveManager,
        TestCancellableProgressWidget,
        # Nouveaux tests critiques
        TestCompilationWorker,
        TestErrorRecoveryManager,
        TestChunkedFileReader,
        TestInputValidator,
        TestIntegrationEndToEnd,
        TestPerformanceBasic
    ]
    
    for test_class in test_classes:
        tests = unittest.TestLoader().loadTestsFromTestCase(test_class)
        test_suite.addTests(tests)
    
    # Exécuter les tests
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(test_suite)
    
    # Afficher le résumé
    print(f"\n{'='*50}")
    print(f"Tests exécutés: {result.testsRun}")
    print(f"Succès: {result.testsRun - len(result.failures) - len(result.errors)}")
    print(f"Échecs: {len(result.failures)}")
    print(f"Erreurs: {len(result.errors)}")
    print(f"{'='*50}")
    
    return result.wasSuccessful()


class TestFileFormatSupport(unittest.TestCase):
    """Tests de support des différents formats de fichiers"""
    
    def setUp(self):
        """Configuration pour tests de formats"""
        self.test_dir = tempfile.mkdtemp()
        self.detector = AdvancedFileDetector()
    
    def tearDown(self):
        """Nettoyage après tests de formats"""
        shutil.rmtree(self.test_dir, ignore_errors=True)
    
    def test_csv_comma_delimiter(self):
        """Test fichier CSV avec délimiteur virgule"""
        csv_file = os.path.join(self.test_dir, "test_comma.csv")
        with open(csv_file, 'w', encoding='utf-8', newline='') as f:
            f.write("Name,Age,City\nJohn,25,Paris\nJane,30,London\n")
        
        delimiter = self.detector.detect_csv_delimiter(csv_file)
        self.assertEqual(delimiter, ',')
    
    def test_csv_semicolon_delimiter(self):
        """Test fichier CSV avec délimiteur point-virgule"""
        csv_file = os.path.join(self.test_dir, "test_semicolon.csv")
        with open(csv_file, 'w', encoding='utf-8', newline='') as f:
            f.write("Name;Age;City\nJohn;25;Paris\nJane;30;London\n")
        
        delimiter = self.detector.detect_csv_delimiter(csv_file)
        self.assertEqual(delimiter, ';')
    
    def test_tsv_tab_delimiter(self):
        """Test fichier TSV avec délimiteur tabulation"""
        tsv_file = os.path.join(self.test_dir, "test_tab.tsv")
        with open(tsv_file, 'w', encoding='utf-8', newline='') as f:
            f.write("Name\tAge\tCity\nJohn\t25\tParis\nJane\t30\tLondon\n")
        
        delimiter = self.detector.detect_csv_delimiter(tsv_file)
        self.assertEqual(delimiter, '\t')
    
    def test_encoding_utf8(self):
        """Test détection encodage UTF-8"""
        utf8_file = os.path.join(self.test_dir, "test_utf8.csv")
        with open(utf8_file, 'w', encoding='utf-8') as f:
            f.write("Nom,Âge,Ville\nJean,25,Paris\nMarie,30,Lyon\n")
        
        encoding = self.detector.detect_csv_encoding(utf8_file)
        self.assertIn(encoding.lower(), ['utf-8', 'utf-8-sig'])
    
    def test_encoding_latin1(self):
        """Test détection encodage Latin-1"""
        latin1_file = os.path.join(self.test_dir, "test_latin1.csv")
        with open(latin1_file, 'w', encoding='latin-1') as f:
            f.write("Nom,Age,Ville\nJean,25,Montréal\nMarie,30,Québec\n")
        
        encoding = self.detector.detect_csv_encoding(latin1_file)
        self.assertIsInstance(encoding, str)


class TestAdvancedLogger(unittest.TestCase):
    """Tests pour la classe AdvancedLogger"""
    
    def setUp(self):
        """Configuration pour tests de logging"""
        self.test_dir = tempfile.mkdtemp()
        self.log_file = os.path.join(self.test_dir, "test.log")
        logging = AdvancedLogger("test_logger", self.log_file, max_size=1024*1024)
    
    def tearDown(self):
        """Nettoyage après tests de logging"""
        # Fermer tous les handlers pour éviter les verrous de fichiers
        for handler in logging.logger.handlers[:]:
            handler.close()
            logging.logger.removeHandler(handler)
        shutil.rmtree(self.test_dir, ignore_errors=True)
    
    def test_logger_creation(self):
        """Test création du logger"""
        self.assertIsNotNone(logging.logger)
        self.assertEqual(logging.logger.name, "test_logger")
    
    def test_log_levels(self):
        """Test des différents niveaux de log"""
        logging.info("Test info message")
        logging.warning("Test warning message")
        logging.error("Test error message")
        
        # Vérifier que le fichier de log est créé
        self.assertTrue(os.path.exists(self.log_file))
    
    def test_structured_logging(self):
        """Test du logging structuré"""
        context = {
            "operation": "test_operation",
            "file_count": 5,
            "duration": 1.23
        }
        
        logging.log_operation("Test operation completed", context)
        
        # Vérifier que le log contient les informations structurées
        if os.path.exists(self.log_file):
            with open(self.log_file, 'r', encoding='utf-8') as f:
                log_content = f.read()
                self.assertIn("test_operation", log_content)


class TestRealTimeValidator(unittest.TestCase):
    """Tests pour la classe RealTimeValidator"""
    
    def setUp(self):
        """Configuration pour tests de validation"""
        self.app = QApplication.instance()
        if self.app is None:
            self.app = QApplication([])
        
        # Créer un widget parent factice
        self.parent_widget = QWidget()
        self.validator = RealTimeValidator(self.parent_widget)
    
    def test_validator_initialization(self):
        """Test initialisation du validateur"""
        self.assertIsNotNone(self.validator)
        self.assertEqual(len(self.validator.validation_rules), 0)
    
    def test_add_validation_rule(self):
        """Test ajout de règle de validation"""
        def test_rule(value):
            return len(value) > 0, "La valeur ne peut pas être vide"
        
        field_name = "test_field"
        self.validator.add_validation_rule(field_name, test_rule)
        
        self.assertIn(field_name, self.validator.validation_rules)
    
    def test_validate_field(self):
        """Test validation d'un champ"""
        def test_rule(value):
            return len(value) > 0, "La valeur ne peut pas être vide"
        
        field_name = "test_field"
        self.validator.add_validation_rule(field_name, test_rule)
        
        # Test valeur valide
        is_valid, message = self.validator.validate_field(field_name, "test_value")
        self.assertTrue(is_valid)
        
        # Test valeur invalide
        is_valid, message = self.validator.validate_field(field_name, "")
        self.assertFalse(is_valid)
        self.assertEqual(message, "La valeur ne peut pas être vide")


class TestCriticalErrorHandling(unittest.TestCase):
    """Tests pour la gestion d'erreurs critiques"""
    
    def setUp(self):
        """Configuration pour tests d'erreurs critiques"""
        self.error_manager = ErrorRecoveryManager()
    
    def test_file_permission_error(self):
        """Test gestion erreur de permission de fichier"""
        error_msg = "Permission denied: /protected/file.csv"
        suggestions = self.error_manager.get_recovery_suggestions(error_msg)
        
        self.assertIsInstance(suggestions, list)
        self.assertGreater(len(suggestions), 0)
        
        # Vérifier qu'il y a des suggestions pertinentes
        suggestions_text = " ".join(suggestions).lower()
        self.assertTrue(any(word in suggestions_text for word in ["permission", "access", "admin"]))
    
    def test_memory_error_handling(self):
        """Test gestion erreur de mémoire"""
        error_msg = "MemoryError: Unable to allocate array"
        suggestions = self.error_manager.get_recovery_suggestions(error_msg)
        
        self.assertIsInstance(suggestions, list)
        suggestions_text = " ".join(suggestions).lower()
        self.assertTrue(any(word in suggestions_text for word in ["memory", "size", "chunk"]))
    
    def test_encoding_error_handling(self):
        """Test gestion erreur d'encodage"""
        error_msg = "UnicodeDecodeError: 'utf-8' codec can't decode"
        suggestions = self.error_manager.get_recovery_suggestions(error_msg)
        
        self.assertIsInstance(suggestions, list)
        suggestions_text = " ".join(suggestions).lower()
        self.assertTrue(any(word in suggestions_text for word in ["encoding", "utf", "character"]))


class TestSecurity(unittest.TestCase):
    """Tests de sécurité pour les nouvelles protections"""
    
    def setUp(self):
        """Configuration pour tests de sécurité"""
        self.temp_dir = tempfile.mkdtemp()
        self.security_manager = SecurityManager()
        self.security_validator = SecurityValidator()
    
    def tearDown(self):
        """Nettoyage après tests de sécurité"""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_path_traversal_protection(self):
        """Test protection contre path traversal"""
        dangerous_paths = [
            "../../../etc/passwd",
            "..\\..\\windows\\system32",
            "/etc/shadow",
            "C:\\Windows\\System32\\config\\SAM",
            "../../../../root/.ssh/id_rsa"
        ]
        
        for dangerous_path in dangerous_paths:
            with self.assertRaises(SecurityError):
                self.security_manager.sanitize_file_path(dangerous_path)
    
    def test_filename_validation(self):
        """Test validation des noms de fichiers"""
        # Noms valides
        valid_names = ["document.csv", "data_2024.xlsx", "report-final.txt"]
        for name in valid_names:
            self.assertTrue(self.security_manager.validate_filename(name))
        
        # Noms invalides
        invalid_names = ["<script>.csv", "file|pipe.xlsx", "con.txt", "prn.csv", "test?.xlsx"]
        for name in invalid_names:
            with self.assertRaises(SecurityError):
                self.security_manager.validate_filename(name)
    
    def test_file_size_limits(self):
        """Test des limites de taille de fichier"""
        # Créer un petit fichier (valide)
        small_file = os.path.join(self.temp_dir, "small.csv")
        with open(small_file, 'w') as f:
            f.write("col1,col2\ndata1,data2\n")
        
        # Doit passer la validation
        self.assertTrue(self.security_manager.check_file_size(small_file))
        
        # Simuler un gros fichier en modifiant la constante temporairement
        original_max = globals()['MAX_FILE_SIZE']
        globals()['MAX_FILE_SIZE'] = 100  # 100 bytes seulement
        
        try:
            with self.assertRaises(SecurityError):
                self.security_manager.check_file_size(small_file)
        finally:
            globals()['MAX_FILE_SIZE'] = original_max
    
    def test_forbidden_extensions(self):
        """Test des extensions interdites"""
        dangerous_files = [
            "virus.exe", "script.bat", "malware.vbs", 
            "trojan.com", "backdoor.scr", "keylogger.pif"
        ]
        
        for filename in dangerous_files:
            with self.assertRaises(SecurityError):
                self.security_manager.validate_filename(filename)
    
    def test_session_size_tracking(self):
        """Test du suivi de la taille de session"""
        # Créer plusieurs fichiers
        files = []
        for i in range(3):
            file_path = os.path.join(self.temp_dir, f"test_{i}.csv")
            with open(file_path, 'w') as f:
                f.write("col1,col2\n" + "data,data\n" * 100)  # ~1KB chacun
            files.append(file_path)
        
        # Ajouter les fichiers à la session
        for file_path in files:
            file_size = os.path.getsize(file_path)
            self.security_manager.add_processed_file(file_path, file_size)
        
        # Vérifier les statistiques
        stats = self.security_manager.get_session_stats()
        self.assertEqual(stats['files_count'], 3)
        self.assertGreater(stats['total_size'], 0)
    
    def test_integrity_verification(self):
        """Test de vérification d'intégrité"""
        # Créer un fichier CSV normal
        csv_file = os.path.join(self.temp_dir, "test.csv")
        with open(csv_file, 'w', encoding='utf-8') as f:
            f.write("Name,Age,City\nJohn,25,Paris\nJane,30,London\n")
        
        # Vérifier l'intégrité
        integrity_info = self.security_manager.verify_file_integrity(csv_file)
        
        self.assertIn('sha256', integrity_info)
        self.assertIn('mime_type', integrity_info)
        self.assertTrue(integrity_info['is_safe'])
        self.assertEqual(len(integrity_info['sha256']), 64)  # SHA-256 = 64 hex chars
    
    def test_suspicious_content_detection(self):
        """Test détection de contenu suspect"""
        # Créer un fichier avec contenu suspect
        malicious_file = os.path.join(self.temp_dir, "malicious.csv")
        with open(malicious_file, 'w', encoding='utf-8') as f:
            f.write("Name,Age,Script\nJohn,25,<script>alert('xss')</script>\n")
        
        # Doit détecter le contenu suspect
        with self.assertRaises(SecurityError):
            self.security_manager.verify_file_integrity(malicious_file)
    
    def test_directory_validation(self):
        """Test validation des répertoires"""
        # Répertoire valide (temp)
        is_valid, message = self.security_validator.validate_directory_path(self.temp_dir)
        self.assertTrue(is_valid)
        
        # Répertoire inexistant
        fake_dir = os.path.join(self.temp_dir, "nonexistent")
        is_valid, message = self.security_validator.validate_directory_path(fake_dir)
        self.assertFalse(is_valid)
    
    def test_file_selection_validation(self):
        """Test validation d'une sélection de fichiers"""
        # Créer des fichiers de test
        test_files = []
        for i in range(2):
            filename = f"test_{i}.csv"
            file_path = os.path.join(self.temp_dir, filename)
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(f"col1,col2\ndata{i},value{i}\n")
            test_files.append(filename)
        
        # Valider la sélection
        is_valid, error_message, file_details = self.security_validator.validate_file_selection(
            test_files, self.temp_dir
        )
        
        self.assertTrue(is_valid)
        self.assertEqual(len(file_details), 2)
        self.assertTrue(all(details['is_safe'] for details in file_details))


class TestMonitoringSystems(unittest.TestCase):
    """Tests pour les systèmes de surveillance et monitoring"""
    
    def setUp(self):
        """Configuration pour tests de surveillance"""
        self.temp_dir = tempfile.mkdtemp()
        
    def tearDown(self):
        """Nettoyage après tests de surveillance"""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_structured_logger_initialization(self):
        """Test initialisation du logger structuré"""
        logger = StructuredLogger("TestLogger", self.temp_dir)
        
        self.assertEqual(logger.name, "TestLogger")
        self.assertTrue(logger.log_dir.exists())
        self.assertIsNotNone(logger.session_id)
    
    def test_structured_logger_events(self):
        """Test logging d'événements structurés"""
        logger = StructuredLogger("TestLogger", self.temp_dir)
        
        # Test logging de différents niveaux
        logger.log_structured("INFO", "test_event", key1="value1", key2=42)
        logger.log_operation("test_operation", "success", duration=1.5)
        
        # Vérifier que les fichiers de log sont créés
        log_file = logger.log_dir / "testlogger.log"
        self.assertTrue(log_file.exists())
    
    def test_performance_monitor_metrics(self):
        """Test enregistrement de métriques de performance"""
        monitor = PerformanceMonitor()
        
        # Enregistrer quelques métriques
        monitor.record_metric("test_metric", 100.0, "units")
        monitor.record_metric("memory_usage", 75.5, "percent")
        
        self.assertEqual(len(monitor.metrics_buffer), 2)
        self.assertEqual(monitor.metrics_buffer[0].metric_name, "test_metric")
        self.assertEqual(monitor.metrics_buffer[0].value, 100.0)
    
    def test_performance_monitor_operations(self):
        """Test chronométrage d'opérations"""
        monitor = PerformanceMonitor()
        
        # Démarrer et terminer une opération
        operation_id = monitor.start_operation("test_operation")
        time.sleep(0.1)  # Simuler du travail
        monitor.end_operation("test_operation", operation_id)
        
        self.assertEqual(monitor.operation_counts["test_operation"], 1)
        self.assertEqual(len(monitor.operation_durations["test_operation"]), 1)
        self.assertGreater(monitor.operation_durations["test_operation"][0], 0.05)
    
    def test_health_monitor_registration(self):
        """Test enregistrement de health checks"""
        monitor = HealthMonitor()
        
        def dummy_check():
            return HealthCheckResult(
                timestamp=datetime.now().isoformat(),
                component="test",
                status="healthy",
                message="Test OK"
            )
        
        monitor.register_health_check("test_check", dummy_check, interval=10)
        
        self.assertIn("test_check", monitor.health_checks)
        self.assertEqual(monitor.health_checks["test_check"]["interval"], 10)
        self.assertEqual(monitor.health_checks["test_check"]["function"], dummy_check)
    
    def test_health_monitor_status(self):
        """Test récupération du statut de santé"""
        monitor = HealthMonitor()
        
        def healthy_check():
            return HealthCheckResult(
                timestamp=datetime.now().isoformat(),
                component="test",
                status="healthy",
                message="Test OK"
            )
        
        def warning_check():
            return HealthCheckResult(
                timestamp=datetime.now().isoformat(),
                component="test_warning",
                status="warning",
                message="Test Warning"
            )
        
        monitor.register_health_check("healthy", healthy_check)
        monitor.register_health_check("warning", warning_check)
        
        # Exécuter manuellement les checks
        monitor.health_checks["healthy"]["last_result"] = healthy_check()
        monitor.health_checks["warning"]["last_result"] = warning_check()
        
        status = monitor.get_health_status()
        
        self.assertEqual(status["overall_status"], "warning")  # Warning à cause du warning_check
        self.assertIn("healthy", status["checks"])
        self.assertIn("warning", status["checks"])
    
    def test_system_metrics_collection(self):
        """Test collecte des métriques système"""
        monitor = PerformanceMonitor()
        
        metrics = monitor.get_system_metrics()
        
        # Vérifier que les métriques principales sont présentes
        self.assertIn("memory", metrics)
        self.assertIn("cpu", metrics)
        self.assertIn("disk", metrics)
        
        # Vérifier les structures des métriques
        if metrics:  # Si psutil fonctionne
            self.assertIn("percent", metrics["memory"])
            self.assertIn("total", metrics["memory"])
            self.assertIsInstance(metrics["cpu"]["percent"], (int, float))
    
    def test_alert_generation(self):
        """Test génération d'alertes"""
        monitor = HealthMonitor()
        
        def critical_check():
            return HealthCheckResult(
                timestamp=datetime.now().isoformat(),
                component="critical_test",
                status="critical",
                message="Critical issue detected"
            )
        
        monitor.register_health_check("critical", critical_check, interval=1)
        
        # Simuler l'exécution du check
        result = critical_check()
        
        # Simuler l'ajout d'alerte (comme le ferait _monitoring_loop)
        alert = Alert(
            timestamp=result.timestamp,
            level=AlertLevel.CRITICAL,
            component=result.component,
            message=result.message
        )
        monitor.alert_queue.put(alert)
        
        # Vérifier la récupération des alertes
        alerts = monitor.get_recent_alerts(limit=1)
        self.assertEqual(len(alerts), 1)
        self.assertEqual(alerts[0].level, AlertLevel.CRITICAL)
        self.assertEqual(alerts[0].component, "critical_test")


class TestRobustnessSystems(unittest.TestCase):
    """Tests pour les systèmes de robustesse et récupération"""
    
    def setUp(self):
        """Configuration pour tests de robustesse"""
        self.temp_dir = tempfile.mkdtemp()
        
    def tearDown(self):
        """Nettoyage après tests de robustesse"""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_configuration_manager_creation(self):
        """Test création et chargement de configuration"""
        config_file = os.path.join(self.temp_dir, "test_config.json")
        config_manager = ConfigurationManager(config_file)
        
        # Charger la configuration par défaut
        config = config_manager.load_configuration()
        
        self.assertIsInstance(config, ConfigurationSchema)
        self.assertTrue(config.auto_save_enabled)
        self.assertEqual(config.language, "fr")
        self.assertTrue(os.path.exists(config_file))
    
    def test_configuration_validation(self):
        """Test validation de configuration"""
        config_file = os.path.join(self.temp_dir, "test_config.json")
        config_manager = ConfigurationManager(config_file)
        
        # Créer une configuration avec des valeurs invalides
        invalid_config = {
            "language": "invalid_lang",  # Invalide
            "security_level": "unknown", # Invalide
            "max_threads": -5,           # Invalide
            "auto_save_enabled": "yes"   # Mauvais type mais corrigible
        }
        
        # Valider et corriger
        validated = config_manager._validate_configuration(invalid_config)
        
        self.assertEqual(validated["language"], "fr")  # Corrigé
        self.assertEqual(validated["security_level"], "high")  # Corrigé
        self.assertEqual(validated["max_threads"], 1)  # Corrigé au minimum
        self.assertTrue(validated["auto_save_enabled"])  # Converti en bool
    
    def test_configuration_save_load_cycle(self):
        """Test cycle sauvegarde-chargement de configuration"""
        config_file = os.path.join(self.temp_dir, "test_config.json")
        config_manager = ConfigurationManager(config_file)
        
        # Modifier la configuration
        config_manager.set_setting("language", "en")
        config_manager.set_setting("max_threads", 8)
        config_manager.set_setting("auto_save_enabled", False)
        
        # Sauvegarder
        success = config_manager.save_configuration()
        self.assertTrue(success)
        
        # Créer un nouveau gestionnaire et charger
        new_config_manager = ConfigurationManager(config_file)
        loaded_config = new_config_manager.load_configuration()
        
        self.assertEqual(loaded_config.language, "en")
        self.assertEqual(loaded_config.max_threads, 8)
        self.assertFalse(loaded_config.auto_save_enabled)
    
    def test_error_handler_classification(self):
        """Test classification des types d'erreurs"""
        error_handler = AdvancedErrorHandler()
        
        # Test différents types d'erreurs
        file_error = FileNotFoundError("File not found")
        memory_error = MemoryError("Out of memory")
        permission_error = PermissionError("Access denied")
        
        self.assertEqual(error_handler._classify_error(file_error), ErrorType.FILE_ACCESS)
        self.assertEqual(error_handler._classify_error(memory_error), ErrorType.MEMORY_ERROR)
        self.assertEqual(error_handler._classify_error(permission_error), ErrorType.PERMISSION_ERROR)
    
    def test_error_handler_retry_strategy(self):
        """Test stratégie de retry automatique"""
        error_handler = AdvancedErrorHandler()
        
        # Test retry avec contexte
        context = {'max_retries': 3, 'current_retry': 0}
        success, message, updated_context = error_handler._strategy_retry(Exception("Test"), context)
        
        self.assertTrue(success)
        self.assertEqual(updated_context['current_retry'], 1)
        self.assertIn("Tentative 1/3", message)
        
        # Test retry au maximum
        context['current_retry'] = 3
        success, message, _ = error_handler._strategy_retry(Exception("Test"), context)
        
        self.assertFalse(success)
        self.assertIn("maximum", message)
    
    def test_error_handler_auto_fix_strategy(self):
        """Test stratégie de correction automatique"""
        error_handler = AdvancedErrorHandler()
        
        # Test correction d'erreur mémoire
        success, message, data = error_handler._strategy_auto_fix(
            MemoryError("Out of memory"), ErrorType.MEMORY_ERROR, {}
        )
        
        self.assertTrue(success)
        self.assertIn("mémoire", message.lower())
        self.assertTrue(data.get("memory_cleaned", False))
        
        # Test correction d'erreur de validation
        success, message, data = error_handler._strategy_auto_fix(
            Exception("Empty data found"), ErrorType.VALIDATION_ERROR, {}
        )
        
        self.assertTrue(success)
        self.assertTrue(data.get("skip_empty", False))
    
    def test_auto_save_manager_session_saving(self):
        """Test sauvegarde de session"""
        auto_save = AutoSaveManager(self.temp_dir)
        
        # Données de session de test
        session_data = {
            'directory': '/test/path',
            'settings': {'option1': True, 'option2': 'value'},
            'window_state': {'width': 800, 'height': 600}
        }
        
        # Sauvegarder la session
        success = auto_save.save_session_state(session_data)
        self.assertTrue(success)
        
        # Vérifier que le fichier a été créé
        backup_files = list(auto_save.backup_dir.glob("session_backup_*.json"))
        self.assertEqual(len(backup_files), 1)
        
        # Charger et vérifier
        loaded_backup = auto_save.load_latest_backup()
        self.assertIsNotNone(loaded_backup)
        self.assertEqual(loaded_backup['application_state'], session_data)
    
    def test_auto_save_manager_backup_cleanup(self):
        """Test nettoyage automatique des sauvegardes"""
        auto_save = AutoSaveManager(self.temp_dir)
        auto_save.max_backup_files = 3
        
        # Créer plusieurs sauvegardes
        for i in range(5):
            session_data = {'test': f'data_{i}'}
            auto_save.save_session_state(session_data)
            time.sleep(0.1)  # Assurer des timestamps différents
        
        # Vérifier que seulement 3 fichiers sont conservés
        backup_files = list(auto_save.backup_dir.glob("session_backup_*.json"))
        self.assertEqual(len(backup_files), 3)
    
    def test_auto_save_manager_manual_backup(self):
        """Test création de sauvegarde manuelle"""
        auto_save = AutoSaveManager(self.temp_dir)
        
        session_data = {'manual_test': True}
        success = auto_save.create_manual_backup(session_data, "test_backup")
        
        self.assertTrue(success)
        
        # Vérifier que le fichier manuel existe
        manual_files = list(auto_save.backup_dir.glob("manual_backup_test_backup_*.json"))
        self.assertEqual(len(manual_files), 1)
        
        # Vérifier le contenu
        with open(manual_files[0], 'r', encoding='utf-8') as f:
            backup_data = json.load(f)
        
        self.assertEqual(backup_data['type'], 'manual')
        self.assertEqual(backup_data['name'], 'test_backup')
        self.assertEqual(backup_data['application_state'], session_data)
    
    def test_auto_save_manager_interval_control(self):
        """Test contrôle de l'intervalle de sauvegarde"""
        auto_save = AutoSaveManager(self.temp_dir)
        auto_save.backup_interval = 1  # 1 seconde
        
        # Première sauvegarde
        success1 = auto_save.save_session_state({'test': 1})
        self.assertTrue(success1)
        
        # Tentative immédiate (doit être ignorée)
        success2 = auto_save.save_session_state({'test': 2})
        self.assertFalse(success2)  # Trop tôt
        
        # Attendre l'intervalle
        time.sleep(1.1)
        success3 = auto_save.save_session_state({'test': 3})
        self.assertTrue(success3)  # Maintenant ça marche


def run_tests():
    """Exécute tous les tests unitaires avec couverture étendue"""
    print("=" * 60)
    print("🧪 SUITE DE TESTS UNITAIRES ÉTENDUE - ExcelCompiler v3.1")
    print("=" * 60)
    
    # Créer une suite de tests
    test_suite = unittest.TestSuite()
    
    # Ajouter toutes les classes de tests (ordre par priorité)
    critical_tests = [
        TestCompilationWorker,           # Tests critiques du worker principal
        TestIntegrationEndToEnd,         # Tests d'intégration end-to-end
        TestCriticalErrorHandling,       # Tests gestion erreurs critiques
    ]
    
    core_tests = [
        TestThreadSafeGUIHelper,         # Tests thread safety
        TestMemoryManager,               # Tests gestion mémoire
        TestFileMetadataCache,           # Tests cache fichiers
        TestCancellationToken,           # Tests annulation
        TestErrorRecoveryManager,        # Tests récupération erreurs
        TestInputValidator,              # Tests validation entrées
    ]
    
    feature_tests = [
        TestAdvancedFileDetector,        # Tests détection formats
        TestFileFormatSupport,           # Tests support formats
        TestChunkedFileReader,           # Tests lecture par chunks
        TestAdvancedLogger,              # Tests logging avancé
        TestRealTimeValidator,           # Tests validation temps réel
    ]
    
    ui_tests = [
        TestValidationIndicator,         # Tests indicateurs validation
        TestResponsiveManager,           # Tests responsive design
        TestCancellableProgressWidget,   # Tests widget progression
    ]
    
    performance_tests = [
        TestPerformanceBasic,            # Tests performance de base
    ]
    
    security_tests = [
        TestSecurity,                    # Tests de sécurité
    ]
    
    monitoring_tests = [
        TestMonitoringSystems,           # Tests systèmes de surveillance
    ]
    
    robustness_tests = [
        TestRobustnessSystems,           # Tests systèmes de robustesse
    ]
    
    # Combiner toutes les classes de tests
    all_test_classes = critical_tests + core_tests + feature_tests + ui_tests + performance_tests + security_tests + monitoring_tests + robustness_tests
    
    print(f"📊 Chargement de {len(all_test_classes)} classes de tests...")
    print(f"   • Tests critiques: {len(critical_tests)}")
    print(f"   • Tests noyau: {len(core_tests)}")
    print(f"   • Tests fonctionnalités: {len(feature_tests)}")
    print(f"   • Tests interface: {len(ui_tests)}")
    print(f"   • Tests performance: {len(performance_tests)}")
    print(f"   • Tests sécurité: {len(security_tests)}")
    print(f"   • Tests surveillance: {len(monitoring_tests)}")
    print(f"   • Tests robustesse: {len(robustness_tests)}")
    print("-" * 60)
    
    for test_class in all_test_classes:
        tests = unittest.TestLoader().loadTestsFromTestCase(test_class)
        test_suite.addTests(tests)
    
    # Exécuter les tests avec verbosité
    runner = unittest.TextTestRunner(
        verbosity=2,
        stream=sys.stdout,
        descriptions=True,
        failfast=False
    )
    
    print(f"🚀 Démarrage de l'exécution des tests...")
    print("-" * 60)
    
    start_time = time.time()
    result = runner.run(test_suite)
    end_time = time.time()
    
    # Afficher le résumé détaillé
    print("\n" + "=" * 60)
    print("📋 RÉSUMÉ DES TESTS UNITAIRES")
    print("=" * 60)
    
    total_tests = result.testsRun
    successes = total_tests - len(result.failures) - len(result.errors)
    failure_count = len(result.failures)
    error_count = len(result.errors)
    
    success_rate = (successes / total_tests * 100) if total_tests > 0 else 0
    execution_time = end_time - start_time
    
    print(f"✅ Tests exécutés: {total_tests}")
    print(f"🎯 Succès: {successes} ({success_rate:.1f}%)")
    print(f"❌ Échecs: {failure_count}")
    print(f"💥 Erreurs: {error_count}")
    print(f"⏱️  Temps d'exécution: {execution_time:.2f}s")
    
    # Évaluation de la qualité
    if success_rate >= 95:
        quality_status = "🟢 EXCELLENT"
    elif success_rate >= 85:
        quality_status = "🟡 BON"
    elif success_rate >= 70:
        quality_status = "🟠 ACCEPTABLE"
    else:
        quality_status = "🔴 INSUFFISANT"
    
    print(f"🏆 Statut qualité: {quality_status}")
    
    # Recommandations
    if failure_count > 0 or error_count > 0:
        print("\n⚠️  ATTENTION:")
        if failure_count > 0:
            print(f"   • {failure_count} test(s) ont échoué - vérifiez la logique métier")
        if error_count > 0:
            print(f"   • {error_count} erreur(s) détectée(s) - vérifiez les dépendances")
        print("   • Consultez les détails ci-dessus pour corriger les problèmes")
    else:
        print("\n🎉 FÉLICITATIONS ! Tous les tests sont passés avec succès.")
        print("   L'application est prête pour la production.")
    
    print("=" * 60)
    
    return result.wasSuccessful()


if __name__ == "__main__":
    import sys
    
    # Si "--test" est passé en argument, exécuter les tests
    if len(sys.argv) > 1 and sys.argv[1] == "--test":
        run_tests()
    else:
        main()

    


