"""
Fenêtre principale de l'application ExcelCompiler v3.2
Auteur: GOUNOU N'GOBI Chabi Zimé
"""

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QMessageBox, QTextBrowser, QLabel,
    QStackedWidget, QFrame, QScrollArea, QButtonGroup,
    QGraphicsDropShadowEffect
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QIcon, QColor
from pathlib import Path

from ui.widgets import (
    FileSelectorWidget,
    OptionsWidget,
    ProgressWidget,
    ResultsWidget
)
from ui.workers import CompilationWorker
from ui.styles import EXCEL_GREEN, WHITE
from ui.styles import theme as T
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

    @staticmethod
    def _apply_soft_shadow(widget, blur=24, dy=4, alpha=28):
        """Ombre portée douce pour donner du relief aux cartes (effet premium)."""
        shadow = QGraphicsDropShadowEffect(widget)
        shadow.setBlurRadius(blur)
        shadow.setXOffset(0)
        shadow.setYOffset(dy)
        shadow.setColor(QColor(15, 23, 42, alpha))  # slate translucide
        widget.setGraphicsEffect(shadow)

    def setup_ui(self):
        """Configure l'interface utilisateur (header héro + sidebar + pages)."""
        self.setWindowTitle("ExcelCompiler — Compilateur Excel Intelligent")
        self.setMinimumSize(1080, 720)
        self.resize(1200, 800)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        root = QVBoxLayout(central_widget)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # === HEADER HÉRO ===
        root.addWidget(self.create_header())

        # === CORPS : sidebar + pages ===
        body = QWidget()
        body_layout = QHBoxLayout(body)
        body_layout.setContentsMargins(T.SPACE_LG, T.SPACE_LG, T.SPACE_LG, T.SPACE_SM)
        body_layout.setSpacing(T.SPACE_LG)

        self.sidebar = self.create_sidebar()
        body_layout.addWidget(self.sidebar)

        # Pages empilées
        self.pages = QStackedWidget()
        self.pages.addWidget(self.create_compilation_tab())   # index 0
        self.pages.addWidget(self.create_results_tab())       # index 1
        self.pages.addWidget(self.create_about_tab())         # index 2
        body_layout.addWidget(self.pages, stretch=1)

        root.addWidget(body, stretch=1)

        # === FOOTER (barre de statut discrète) ===
        footer = QWidget()
        footer.setStyleSheet(f"background-color: {T.BG_SURFACE}; "
                             f"border-top: 1px solid {T.BORDER};")
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(T.SPACE_LG, T.SPACE_SM, T.SPACE_LG, T.SPACE_SM)
        self.status_label = QLabel("Prêt")
        self.status_label.setStyleSheet(f"color: {T.TEXT_MUTED}; font-size: {T.FONT_SIZE_SM}pt;")
        footer_layout.addWidget(self.status_label)
        footer_layout.addStretch()
        root.addWidget(footer)

    def create_header(self) -> QWidget:
        """Header héro : bandeau accent avec logo, titre et accroche."""
        header = QWidget()
        header.setFixedHeight(84)
        header.setStyleSheet(f"background-color: {T.ACCENT};")
        layout = QHBoxLayout(header)
        layout.setContentsMargins(T.SPACE_LG, 0, T.SPACE_LG, 0)
        layout.setSpacing(T.SPACE_MD)

        # Pastille logo
        logo = QLabel("📊")
        logo.setFixedSize(48, 48)
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo.setStyleSheet(
            f"background-color: rgba(255,255,255,40); border-radius: {T.RADIUS_MD}px; "
            f"font-size: 22pt;"
        )
        layout.addWidget(logo)

        # Titre + accroche
        text_col = QVBoxLayout()
        text_col.setSpacing(0)
        title = QLabel("ExcelCompiler")
        title.setStyleSheet(
            f"color: {T.TEXT_ON_ACCENT}; font-size: 18pt; font-weight: 800;"
        )
        subtitle = QLabel("Compilez vos fichiers Excel en un clic, sans effort.")
        subtitle.setStyleSheet("color: rgba(255,255,255,210); font-size: 10pt;")
        text_col.addWidget(title)
        text_col.addWidget(subtitle)
        layout.addLayout(text_col)

        layout.addStretch()

        # Badge version
        badge = QLabel("v3.2")
        badge.setStyleSheet(
            f"color: {T.TEXT_ON_ACCENT}; background-color: rgba(255,255,255,38); "
            f"border-radius: {T.RADIUS_SM}px; padding: 5px 12px; font-weight: 700;"
        )
        layout.addWidget(badge)

        return header

    def create_sidebar(self) -> QWidget:
        """Navigation latérale (remplace les onglets) pilotant le QStackedWidget."""
        sidebar = QFrame()
        sidebar.setFixedWidth(208)
        sidebar.setStyleSheet(
            f"QFrame {{ background-color: {T.BG_SURFACE}; "
            f"border: 1px solid {T.BORDER}; border-radius: {T.RADIUS_LG}px; }}"
        )
        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(T.SPACE_SM, T.SPACE_MD, T.SPACE_SM, T.SPACE_MD)
        layout.setSpacing(T.SPACE_XS)

        self.nav_group = QButtonGroup(self)
        self.nav_group.setExclusive(True)
        nav_items = [
            ("📝", "Compilation", 0),
            ("📊", "Résultats", 1),
            ("ℹ️", "À propos", 2),
        ]
        for icon, label, index in nav_items:
            btn = self._make_nav_button(icon, label)
            btn.clicked.connect(lambda _checked, i=index: self.navigate_to(i))
            self.nav_group.addButton(btn, index)
            layout.addWidget(btn)

        layout.addStretch()

        # Pied de sidebar : signature discrète
        sign = QLabel("© 2025\nGOUNOU N'GOBI C. Z.")
        sign.setStyleSheet(f"color: {T.TEXT_MUTED}; font-size: 8pt; padding: 8px;")
        layout.addWidget(sign)

        # Activer le premier item
        self.nav_group.button(0).setChecked(True)
        self._apply_soft_shadow(sidebar, blur=28, dy=6, alpha=22)
        return sidebar

    def _page_header(self, title: str, description: str) -> QWidget:
        """En-tête de page : titre fort + description discrète."""
        header = QWidget()
        col = QVBoxLayout(header)
        col.setContentsMargins(2, 0, 0, 0)
        col.setSpacing(2)
        t = QLabel(title)
        t.setStyleSheet(
            f"color: {T.TEXT_PRIMARY}; font-size: 16pt; font-weight: 800;"
        )
        d = QLabel(description)
        d.setStyleSheet(f"color: {T.TEXT_SECONDARY}; font-size: 10pt;")
        d.setWordWrap(True)
        col.addWidget(t)
        col.addWidget(d)
        return header

    def _make_nav_button(self, icon: str, label: str) -> QPushButton:
        """Bouton de navigation latérale, checkable, au style premium."""
        btn = QPushButton(f"  {icon}   {label}")
        btn.setCheckable(True)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setMinimumHeight(42)
        btn.setStyleSheet(f"""
            QPushButton {{
                text-align: left;
                padding: 8px 14px;
                border: none;
                border-radius: {T.RADIUS_SM}px;
                background-color: transparent;
                color: {T.TEXT_SECONDARY};
                font-size: 10pt;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background-color: {T.BG_SUBTLE};
                color: {T.TEXT_PRIMARY};
            }}
            QPushButton:checked {{
                background-color: {T.ACCENT_SOFT};
                color: {T.ACCENT};
                font-weight: 700;
            }}
        """)
        return btn

    def navigate_to(self, index: int):
        """Change de page et synchronise la sidebar."""
        self.pages.setCurrentIndex(index)
        btn = self.nav_group.button(index)
        if btn and not btn.isChecked():
            btn.setChecked(True)

    def create_compilation_tab(self) -> QWidget:
        """Page de compilation : contenu scrollable + barre d'actions fixe."""
        tab = QWidget()
        tab_layout = QVBoxLayout(tab)
        tab_layout.setContentsMargins(0, 0, 0, 0)
        tab_layout.setSpacing(0)

        # Zone scrollable (pas de défilement horizontal : le contenu s'adapte)
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QScrollArea.Shape.NoFrame)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll_area.setStyleSheet("QScrollArea { background: transparent; }")

        scroll_content = QWidget()
        layout = QVBoxLayout(scroll_content)
        layout.setSpacing(T.SPACE_MD)
        layout.setContentsMargins(0, 0, T.SPACE_SM, 0)

        layout.addWidget(self._page_header(
            "Compiler des fichiers",
            "Sélectionnez vos fichiers, choisissez le mode de détection, puis lancez la compilation."
        ))

        self.file_selector = FileSelectorWidget()
        layout.addWidget(self.file_selector)

        self.options_widget = OptionsWidget()
        layout.addWidget(self.options_widget)

        self.progress_widget = ProgressWidget()
        layout.addWidget(self.progress_widget)

        layout.addStretch()
        scroll_area.setWidget(scroll_content)
        tab_layout.addWidget(scroll_area, stretch=1)

        # Barre d'actions fixe (carte détachée par une bordure haute)
        button_container = QFrame()
        button_container.setStyleSheet(
            f"QFrame {{ background-color: {T.BG_SURFACE}; "
            f"border-top: 1px solid {T.BORDER}; border-radius: 0px; }}"
        )
        button_layout = QHBoxLayout(button_container)
        button_layout.setContentsMargins(T.SPACE_MD, T.SPACE_MD, T.SPACE_MD, T.SPACE_MD)
        button_layout.setSpacing(T.SPACE_MD)
        button_layout.addStretch()

        # Bouton aperçu (variante accent-outline, style hérité du QSS global)
        self.button_preview = QPushButton("👁   Aperçu")
        self.button_preview.setProperty("variant", "accent")
        self.button_preview.setMinimumSize(140, 48)
        self.button_preview.setCursor(Qt.CursorShape.PointingHandCursor)
        self.button_preview.clicked.connect(self.show_detection_preview)
        button_layout.addWidget(self.button_preview)

        # Bouton compiler (variante primaire)
        self.button_compile = QPushButton("▶   COMPILER")
        self.button_compile.setProperty("variant", "primary")
        self.button_compile.setFont(QFont(T.FONT_FAMILY, 12, QFont.Weight.Bold))
        self.button_compile.setMinimumSize(210, 48)
        self.button_compile.setCursor(Qt.CursorShape.PointingHandCursor)
        self.button_compile.clicked.connect(self.start_compilation)
        button_layout.addWidget(self.button_compile)

        tab_layout.addWidget(button_container)
        return tab

    def create_results_tab(self) -> QWidget:
        """Page des résultats de compilation."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(T.SPACE_MD)

        layout.addWidget(self._page_header(
            "Résultats",
            "Statistiques et détails de la dernière compilation."
        ))

        self.results_widget = ResultsWidget()
        layout.addWidget(self.results_widget)

        return tab

    def create_about_tab(self) -> QWidget:
        """Page À propos."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(T.SPACE_MD)

        layout.addWidget(self._page_header(
            "À propos",
            "ExcelCompiler — compilateur Excel intelligent."
        ))

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

    def show_detection_preview(self):
        """Affiche un aperçu de la détection des fichiers sélectionnés.

        Utilise la même logique de détection que la compilation, selon le
        mode actif (automatique / référence / manuel).
        """
        from PyQt6.QtWidgets import QApplication
        from PyQt6.QtGui import QCursor
        from core.compilation import ExcelCompiler
        from ui.widgets.preview_dialog import PreviewDialog

        selected_files = self.file_selector.get_selected_files()
        if not selected_files:
            QMessageBox.warning(
                self, "Aucun fichier",
                "Veuillez sélectionner au moins un fichier à prévisualiser."
            )
            return

        options = self.options_widget.get_compilation_options()
        self.status_label.setText("Analyse de la détection en cours...")
        QApplication.setOverrideCursor(QCursor(Qt.CursorShape.WaitCursor))
        try:
            previews = ExcelCompiler(options).preview_detection(selected_files)
        except Exception as e:
            logger.error(f"Erreur aperçu détection: {e}", exc_info=True)
            QMessageBox.critical(
                self, "Erreur d'aperçu",
                f"Impossible de générer l'aperçu:\n\n{e}"
            )
            return
        finally:
            QApplication.restoreOverrideCursor()
            self.status_label.setText("Prêt")

        PreviewDialog(previews, parent=self).exec()

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
