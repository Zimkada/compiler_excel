"""
Fenêtre principale de l'application ExcelCompiler v3.2
Auteur: GOUNOU N'GOBI Chabi Zimé
"""

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QMessageBox, QTabWidget, QTextBrowser,
    QLabel
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QIcon
from pathlib import Path

from ui.widgets import (
    FileSelectorWidget,
    OptionsWidget,
    ProgressWidget,
    ResultsWidget
)
from ui.workers import CompilationWorker
from ui.styles import EXCEL_GREEN, WHITE
from utils import logger


class MainWindow(QMainWindow):
    """
    Fenêtre principale de ExcelCompiler v3.2

    Architecture simplifiée avec détection automatique par défaut
    """

    def __init__(self):
        super().__init__()
        self.compilation_worker = None
        self.setup_ui()
        self.connect_signals()

    def setup_ui(self):
        """Configure l'interface utilisateur"""
        self.setWindowTitle("ExcelCompiler v3.2 - Compilateur Excel Intelligent")
        self.setMinimumSize(900, 700)

        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)

        # === HEADER ===
        header = self.create_header()
        main_layout.addWidget(header)

        # === TABS ===
        tabs = QTabWidget()
        tabs.setFont(QFont("Segoe UI", 9))
        tabs.setStyleSheet(f"""
            QTabWidget::pane {{
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                background-color: white;
            }}
            QTabBar::tab {{
                background-color: #f0f0f0;
                color: #333;
                padding: 10px 20px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }}
            QTabBar::tab:selected {{
                background-color: {EXCEL_GREEN};
                color: white;
                font-weight: bold;
            }}
            QTabBar::tab:hover {{
                background-color: #e0e0e0;
            }}
        """)

        # ONGLET 1: COMPILATION
        compilation_tab = self.create_compilation_tab()
        tabs.addTab(compilation_tab, "📝 Compilation")

        # ONGLET 2: RÉSULTATS
        results_tab = self.create_results_tab()
        tabs.addTab(results_tab, "📊 Résultats")

        # ONGLET 3: À PROPOS
        about_tab = self.create_about_tab()
        tabs.addTab(about_tab, "ℹ️  À propos")

        main_layout.addWidget(tabs)

        # === FOOTER (Status bar) ===
        self.status_label = QLabel("Prêt")
        self.status_label.setStyleSheet("color: #666; padding: 5px;")
        main_layout.addWidget(self.status_label)

    def create_header(self) -> QWidget:
        """Crée le header de l'application"""
        header = QWidget()
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(0, 0, 0, 0)

        title = QLabel("ExcelCompiler v3.2")
        title.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        title.setStyleSheet(f"color: {EXCEL_GREEN}; padding: 5px;")

        subtitle = QLabel("Compilateur Excel Intelligent avec Détection Automatique")
        subtitle.setFont(QFont("Segoe UI", 9))
        subtitle.setStyleSheet("color: #666; font-style: italic; padding-bottom: 10px;")

        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)

        return header

    def create_compilation_tab(self) -> QWidget:
        """Crée l'onglet de compilation principal"""
        from PyQt6.QtWidgets import QScrollArea

        tab = QWidget()
        tab_layout = QVBoxLayout(tab)
        tab_layout.setContentsMargins(0, 0, 0, 0)
        tab_layout.setSpacing(0)

        # Créer une zone scrollable pour le contenu
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QScrollArea.Shape.NoFrame)

        # Widget de contenu scrollable
        scroll_content = QWidget()
        layout = QVBoxLayout(scroll_content)
        layout.setSpacing(10)
        layout.setContentsMargins(5, 5, 5, 5)

        # Widget sélection fichiers
        self.file_selector = FileSelectorWidget()
        layout.addWidget(self.file_selector)

        # Widget options
        self.options_widget = OptionsWidget()
        layout.addWidget(self.options_widget)

        # Widget progression (caché par défaut)
        self.progress_widget = ProgressWidget()
        layout.addWidget(self.progress_widget)

        # Bouton compiler (fixe en bas, pas dans le scroll)
        layout.addStretch()

        # Ajouter le contenu à la zone scrollable
        scroll_area.setWidget(scroll_content)
        tab_layout.addWidget(scroll_area)

        # Zone fixe en bas pour le bouton compiler
        button_container = QWidget()
        button_container.setStyleSheet("background-color: white;")
        button_layout = QHBoxLayout(button_container)
        button_layout.setContentsMargins(10, 10, 10, 10)
        button_layout.addStretch()

        self.button_compile = QPushButton("▶️  COMPILER")
        self.button_compile.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        self.button_compile.setMinimumSize(200, 50)
        self.button_compile.setStyleSheet(f"""
            QPushButton {{
                background-color: {EXCEL_GREEN};
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px 20px;
            }}
            QPushButton:hover {{
                background-color: #1a5c37;
            }}
            QPushButton:pressed {{
                background-color: #14462a;
            }}
            QPushButton:disabled {{
                background-color: #cccccc;
                color: #666666;
            }}
        """)
        self.button_compile.clicked.connect(self.start_compilation)

        button_layout.addWidget(self.button_compile)
        button_layout.addStretch()

        tab_layout.addWidget(button_container)

        return tab

    def create_results_tab(self) -> QWidget:
        """Crée l'onglet des résultats"""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        self.results_widget = ResultsWidget()
        layout.addWidget(self.results_widget)

        return tab

    def create_about_tab(self) -> QWidget:
        """Crée l'onglet À propos"""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        about_text = QTextBrowser()
        about_text.setOpenExternalLinks(True)
        about_text.setHtml("""
        <h2 style='color: #217346;'>ExcelCompiler v3.2</h2>
        <p><b>Compilateur Excel Intelligent avec Détection Automatique</b></p>

        <h3>Auteur</h3>
        <p>GOUNOU N'GOBI Chabi Zimé<br>
        Data Manager & Data Analyst</p>

        <h3>Nouveautés v3.2</h3>
        <ul>
            <li>✨ <b>Détection par fichier de référence</b> - Indiquez l'en-tête du premier fichier, les autres sont alignés automatiquement</li>
            <li>🎯 <b>Architecture modulaire</b> - Code propre et maintenable</li>
            <li>⚡ <b>Interface simplifiée</b> - Workflow optimisé en quelques clics</li>
            <li>📊 <b>Statistiques détaillées</b> - Résultats complets avec métriques</li>
        </ul>

        <h3>Fonctionnalités</h3>
        <ul>
            <li>Compilation de fichiers Excel (.xlsx, .xls, .xlsm)</li>
            <li>Support CSV et TSV</li>
            <li>Détection par fichier de référence ou configuration manuelle</li>
            <li>Gestion des en-têtes multi-lignes</li>
            <li>Export en Excel, CSV ou TSV</li>
            <li>Ajout automatique du nom de fichier source</li>
        </ul>

        <h3>Utilisation</h3>
        <ol>
            <li>Sélectionnez un dossier contenant vos fichiers Excel</li>
            <li>Configurez les options (détection auto activée par défaut)</li>
            <li>Cliquez sur COMPILER</li>
            <li>Consultez les résultats dans l'onglet Résultats</li>
        </ol>

        <hr>
        <p style='text-align: center; color: #666;'>
        © 2025 GOUNOU N'GOBI Chabi Zimé - Tous droits réservés
        </p>
        """)

        layout.addWidget(about_text)

        return tab

    def connect_signals(self):
        """Connecte les signaux et slots"""
        # Bouton annuler compilation
        self.progress_widget.cancel_requested.connect(self.cancel_compilation)

        # Mise à jour status quand sélection fichiers change
        self.file_selector.files_selected.connect(self.update_status)

    def update_status(self, selected_files):
        """Met à jour la barre de status"""
        count = len(selected_files)
        if count == 0:
            self.status_label.setText("Aucun fichier sélectionné")
        elif count == 1:
            self.status_label.setText("1 fichier sélectionné")
        else:
            self.status_label.setText(f"{count} fichiers sélectionnés")

    def start_compilation(self):
        """Démarre la compilation"""
        # Vérifier qu'il y a des fichiers sélectionnés
        selected_files = self.file_selector.get_selected_files()
        if not selected_files:
            QMessageBox.warning(
                self,
                "Aucun fichier",
                "Veuillez sélectionner au moins un fichier à compiler."
            )
            return

        # Vérifier le nom du fichier de sortie
        output_filename = self.options_widget.get_output_file()
        if not output_filename:
            QMessageBox.warning(
                self,
                "Nom de fichier manquant",
                "Veuillez spécifier un nom pour le fichier de sortie."
            )
            return

        # Créer le chemin complet: même dossier que les fichiers sources
        source_directory = Path(selected_files[0]).parent
        output_file = str(source_directory / output_filename)

        # Récupérer les options
        options = self.options_widget.get_compilation_options()
        output_format = self.options_widget.get_output_format()

        logger.info(f"Démarrage compilation: {len(selected_files)} fichiers")
        logger.info(f"Fichier de sortie: {output_file}")

        # Désactiver le bouton compiler
        self.button_compile.setEnabled(False)

        # Afficher la barre de progression
        self.progress_widget.start(f"Compilation de {len(selected_files)} fichiers...")

        # Créer et démarrer le worker
        self.compilation_worker = CompilationWorker(
            selected_files,
            output_file,
            options,
            output_format
        )

        # Connecter les signaux du worker
        self.compilation_worker.progress_update.connect(self.on_progress_update)
        self.compilation_worker.compilation_finished.connect(self.on_compilation_finished)
        self.compilation_worker.error_occurred.connect(self.on_compilation_error)
        self.compilation_worker.compilation_cancelled.connect(self.on_compilation_cancelled)

        # Démarrer le worker
        self.compilation_worker.start()

    def on_progress_update(self, progress: int, message: str):
        """Mise à jour de la progression"""
        self.progress_widget.update_progress(progress, message)
        self.status_label.setText(message)

    def on_compilation_finished(self, result):
        """Appelé quand la compilation est terminée"""
        try:
            logger.info("Compilation terminée avec succès")

            # Mettre à jour la barre de progression
            self.progress_widget.finish_success(f"Compilation réussie! {result.total_rows} lignes")

            # Afficher les résultats
            self.results_widget.display_result(result)

            # Réactiver le bouton
            self.button_compile.setEnabled(True)

            # Message de succès
            output_filename = Path(result.output_file).name if result.output_file else "compilation.xlsx"
            QMessageBox.information(
                self,
                "Compilation réussie",
                f"La compilation est terminée!\n\n"
                f"• Fichiers: {result.successful_files}/{result.total_files}\n"
                f"• Lignes: {result.total_rows}\n"
                f"• Temps: {result.total_processing_time:.2f}s\n\n"
                f"Fichier créé: {output_filename}"
            )

            # Nettoyer le worker
            self.compilation_worker = None

        except Exception as e:
            logger.error(f"Erreur affichage résultats: {e}", exc_info=True)
            # Au moins réactiver le bouton
            self.button_compile.setEnabled(True)
            self.compilation_worker = None

    def on_compilation_error(self, error_message: str):
        """Appelé en cas d'erreur"""
        logger.error(f"Erreur compilation: {error_message}")

        # Mettre à jour la barre de progression
        self.progress_widget.finish_error(error_message)

        # Afficher l'erreur
        self.results_widget.display_error(error_message)

        # Réactiver le bouton
        self.button_compile.setEnabled(True)

        # Message d'erreur
        QMessageBox.critical(
            self,
            "Erreur de compilation",
            f"Une erreur s'est produite:\n\n{error_message}"
        )

        # Nettoyer le worker
        self.compilation_worker = None

    def cancel_compilation(self):
        """Demande l'annulation de la compilation en cours.

        On ne bloque PAS l'UI avec wait(): le worker détectera la demande
        avant le prochain fichier et émettra compilation_cancelled, qui
        finalisera l'état via on_compilation_cancelled().
        """
        if self.compilation_worker and self.compilation_worker.isRunning():
            logger.info("Annulation compilation demandée")
            self.status_label.setText("Annulation en cours...")
            self.compilation_worker.cancel()

    def on_compilation_cancelled(self):
        """Appelé quand le worker confirme l'annulation (pas une erreur)."""
        logger.info("Compilation annulée (confirmée par le worker)")
        self.progress_widget.finish_error("Compilation annulée")
        self.button_compile.setEnabled(True)
        self.status_label.setText("Compilation annulée")
        self.compilation_worker = None

    def closeEvent(self, event):
        """Appelé à la fermeture de la fenêtre"""
        # Annuler compilation en cours si nécessaire
        if self.compilation_worker and self.compilation_worker.isRunning():
            reply = QMessageBox.question(
                self,
                "Compilation en cours",
                "Une compilation est en cours. Voulez-vous vraiment quitter?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )

            if reply == QMessageBox.StandardButton.Yes:
                # Éviter que des slots se déclenchent sur une fenêtre en
                # cours de fermeture: on coupe les signaux avant d'attendre.
                try:
                    self.compilation_worker.blockSignals(True)
                except Exception:
                    pass
                self.compilation_worker.cancel()
                self.compilation_worker.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()
