"""
Fenêtre principale de l'application ExcelCompiler v3.2
Auteur: GOUNOU N'GOBI Chabi Zimé
"""

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QMessageBox, QTextBrowser, QLabel,
    QStackedWidget, QFrame, QScrollArea, QButtonGroup,
    QGraphicsDropShadowEffect, QScroller
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QFont, QIcon, QColor, QPixmap
from pathlib import Path

from ui.widgets import (
    FileSelectorWidget,
    OptionsWidget,
    ProgressWidget,
    ResultsWidget
)
from ui.workers import CompilationWorker
from ui.styles import theme as T
from utils import logger, resource_path


class MainWindow(QMainWindow):
    """
    Fenêtre principale de ExcelCompiler v3.2

    Architecture modulaire ; détection par fichier de référence par défaut.
    """

    def __init__(self):
        super().__init__()
        self.compilation_worker = None
        # Corrections manuelles d'en-tête par fichier, choisies dans l'aperçu
        # et appliquées à la compilation. {chemin: (header_start_row, header_rows)}
        self._manual_overrides = {}
        # Rattachements manuels de colonnes (étape 6), choisis dans l'aperçu.
        # {libellé_source: libellé_cible_du_schéma}
        self._column_aliases = {}
        self._update_worker = None
        self.setup_ui()
        self.connect_signals()
        T.get_manager().theme_changed.connect(self.apply_theme)
        self._start_update_check()

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
        self.setWindowTitle("ExcelCompiler - Compilateur Excel Intelligent")
        self.setMinimumSize(1080, 720)
        self.resize(1200, 800)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        root = QVBoxLayout(central_widget)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # === HEADER HÉRO ===
        root.addWidget(self.create_header())

        # === BANNIÈRE DE MISE À JOUR (cachée par défaut) ===
        root.addWidget(self.create_update_banner())

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
        self.footer = QWidget()
        footer_layout = QHBoxLayout(self.footer)
        footer_layout.setContentsMargins(T.SPACE_LG, T.SPACE_SM, T.SPACE_LG, T.SPACE_SM)
        self.status_label = QLabel("Prêt")
        footer_layout.addWidget(self.status_label)
        footer_layout.addStretch()
        root.addWidget(self.footer)

        # Appliquer les styles inline dépendant du thème (footer, sidebar, etc.)
        self.apply_theme()

    def create_header(self) -> QWidget:
        """Header héro : bandeau accent avec logo, titre et accroche."""
        header = QWidget()
        self.header = header
        header.setFixedHeight(84)
        header.setStyleSheet(f"background-color: {T.ACCENT};")
        layout = QHBoxLayout(header)
        layout.setContentsMargins(T.SPACE_LG, 0, T.SPACE_LG, 0)
        layout.setSpacing(T.SPACE_MD)

        # Pastille logo : vrai logo de l'app (icon.ico) avec repli emoji
        logo = QLabel()
        logo.setFixedSize(48, 48)
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo_path = resource_path("icon.ico")
        logo_pixmap = QPixmap(str(logo_path)) if logo_path.exists() else QPixmap()
        if not logo_pixmap.isNull():
            logo.setPixmap(logo_pixmap.scaled(
                40, 40,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            ))
            logo.setStyleSheet(
                f"background-color: rgba(255,255,255,40); border-radius: {T.RADIUS_MD}px;"
            )
        else:
            logo.setText("📊")
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

        # Bouton bascule clair / sombre
        self.theme_toggle = QPushButton()
        self.theme_toggle.setFixedSize(40, 40)
        self.theme_toggle.setCursor(Qt.CursorShape.PointingHandCursor)
        self.theme_toggle.clicked.connect(self.toggle_theme)
        self._style_theme_toggle()
        layout.addWidget(self.theme_toggle)

        # Badge version
        badge = QLabel("v3.2")
        badge.setStyleSheet(
            f"color: {T.TEXT_ON_ACCENT}; background-color: rgba(255,255,255,38); "
            f"border-radius: {T.RADIUS_SM}px; padding: 5px 12px; font-weight: 700;"
        )
        layout.addWidget(badge)

        return header

    def create_update_banner(self) -> QWidget:
        """Bannière discrète signalant une mise à jour disponible.

        Cachée par défaut ; affichée seulement si le worker de vérification
        trouve une version plus récente. Non-bloquante : l'utilisateur peut
        l'ignorer et continuer à travailler.
        """
        banner = QFrame()
        self._update_banner = banner
        banner.setVisible(False)
        layout = QHBoxLayout(banner)
        layout.setContentsMargins(T.SPACE_LG, T.SPACE_SM, T.SPACE_LG, T.SPACE_SM)
        layout.setSpacing(T.SPACE_MD)

        self._update_label = QLabel("")
        layout.addWidget(self._update_label)
        layout.addStretch()

        self._update_download_btn = QPushButton("Télécharger")
        self._update_download_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        layout.addWidget(self._update_download_btn)

        dismiss = QPushButton("✕")
        dismiss.setFixedSize(28, 28)
        dismiss.setCursor(Qt.CursorShape.PointingHandCursor)
        dismiss.setToolTip("Masquer")
        dismiss.clicked.connect(lambda: banner.setVisible(False))
        layout.addWidget(dismiss)

        banner.setStyleSheet(
            f"QFrame {{ background-color: {T.ACCENT_SOFT}; "
            f"border-bottom: 1px solid {T.BORDER}; }}"
        )
        self._update_label.setStyleSheet(
            f"color: {T.ACCENT}; font-weight: 600;")
        return banner

    def _start_update_check(self):
        """Lance en arrière-plan la vérification de mise à jour (non-bloquant,
        silencieux en cas d'échec ou hors-ligne)."""
        try:
            from ui.workers import UpdateCheckWorker
            from config.constants import APP_VERSION
            self._update_worker = UpdateCheckWorker(APP_VERSION)
            self._update_worker.update_found.connect(self._on_update_found)
            self._update_worker.start()
        except Exception:
            logger.info("Vérification de mise à jour non démarrée", exc_info=True)

    def _on_update_found(self, info):
        """Affiche la bannière quand une mise à jour est disponible."""
        self._update_label.setText(
            f"🔔 Une nouvelle version ({info.latest_version}) est disponible."
        )
        try:
            self._update_download_btn.clicked.disconnect()
        except TypeError:
            pass
        url = info.download_url
        self._update_download_btn.clicked.connect(lambda: self._open_url(url))
        self._update_banner.setVisible(True)

    def _open_url(self, url: str):
        """Ouvre une URL web dans le navigateur par défaut (page de release)."""
        from PyQt6.QtGui import QDesktopServices
        from PyQt6.QtCore import QUrl
        try:
            QDesktopServices.openUrl(QUrl(url))
        except Exception as e:
            logger.warning(f"Impossible d'ouvrir l'URL {url}: {e}")

    def _style_theme_toggle(self):
        """Met à jour l'icône et le style du bouton de bascule de thème.

        Sur le bandeau accent (header), le bouton reste translucide blanc dans
        les deux thèmes ; seule l'icône vectorielle (lune / soleil) change.
        """
        is_dark = T.get_manager().is_dark()
        # Icône = action proposée : en clair on propose le sombre (lune), etc.
        icon = self._make_sun_icon() if is_dark else self._make_moon_icon()
        self.theme_toggle.setIcon(icon)
        self.theme_toggle.setIconSize(QSize(20, 20))
        self.theme_toggle.setText("")
        self.theme_toggle.setToolTip(
            "Passer en mode clair" if is_dark else "Passer en mode sombre"
        )
        self.theme_toggle.setStyleSheet(f"""
            QPushButton {{
                background-color: rgba(255,255,255,38);
                border: none;
                border-radius: {T.RADIUS_MD}px;
            }}
            QPushButton:hover {{
                background-color: rgba(255,255,255,70);
            }}
            QPushButton:pressed {{
                background-color: rgba(255,255,255,28);
            }}
        """)

    @staticmethod
    def _make_moon_icon(size: int = 20, color: str = "#FFFFFF") -> QIcon:
        """Croissant de lune monochrome (obtenu par soustraction de 2 cercles)."""
        from PyQt6.QtGui import QPainter, QPainterPath
        pm = QPixmap(size, size)
        pm.fill(Qt.GlobalColor.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        full = QPainterPath()
        full.addEllipse(3.0, 2.5, size - 6.0, size - 6.0)          # disque plein
        cut = QPainterPath()
        cut.addEllipse(7.5, 1.0, size - 6.0, size - 6.0)           # disque décalé (à soustraire)
        crescent = full.subtracted(cut)
        p.fillPath(crescent, QColor(color))
        p.end()
        return QIcon(pm)

    @staticmethod
    def _make_sun_icon(size: int = 20, color: str = "#FFFFFF") -> QIcon:
        """Soleil monochrome : disque central + 8 rayons."""
        import math
        from PyQt6.QtGui import QPainter, QPen
        pm = QPixmap(size, size)
        pm.fill(Qt.GlobalColor.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        c = size / 2.0
        # Disque central
        p.setBrush(QColor(color))
        p.setPen(Qt.PenStyle.NoPen)
        r = size * 0.22
        p.drawEllipse(int(c - r), int(c - r), int(2 * r), int(2 * r))
        # Rayons
        pen = QPen(QColor(color))
        pen.setWidthF(1.6)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        p.setPen(pen)
        r_in = size * 0.34
        r_out = size * 0.46
        for i in range(8):
            ang = math.pi * i / 4.0
            dx, dy = math.cos(ang), math.sin(ang)
            p.drawLine(
                int(c + dx * r_in), int(c + dy * r_in),
                int(c + dx * r_out), int(c + dy * r_out),
            )
        p.end()
        return QIcon(pm)

    def toggle_theme(self):
        """Bascule entre thème clair et sombre et persiste le choix."""
        mode = T.get_manager().toggle()
        try:
            from PyQt6.QtCore import QSettings
            QSettings("GOUNOU N'GOBI Chabi Zimé", "ExcelCompiler").setValue("theme", mode)
        except Exception:
            logger.warning("Impossible de mémoriser le thème", exc_info=True)

    def create_sidebar(self) -> QWidget:
        """Navigation latérale (remplace les onglets) pilotant le QStackedWidget."""
        sidebar = QFrame()
        self._sidebar_frame = sidebar
        self._nav_buttons = []
        sidebar.setFixedWidth(208)
        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(T.SPACE_SM, T.SPACE_MD, T.SPACE_SM, T.SPACE_MD)
        layout.setSpacing(T.SPACE_XS)

        self.nav_group = QButtonGroup(self)
        self.nav_group.setExclusive(True)
        nav_items = [
            ("📝", "Compilation", 0),
            ("✅", "Résultats", 1),
            ("ℹ️", "À propos", 2),
        ]
        for icon, label, index in nav_items:
            btn = self._make_nav_button(icon, label)
            btn.clicked.connect(lambda _checked, i=index: self.navigate_to(i))
            self.nav_group.addButton(btn, index)
            self._nav_buttons.append(btn)
            layout.addWidget(btn)

        layout.addStretch()

        # Pied de sidebar : signature discrète
        self._sidebar_sign = QLabel("© 2026\nGOUNOU N'GOBI C. Z.")
        layout.addWidget(self._sidebar_sign)

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
        d = QLabel(description)
        d.setWordWrap(True)
        col.addWidget(t)
        col.addWidget(d)
        # Mémoriser pour re-thématisation
        if not hasattr(self, "_page_header_titles"):
            self._page_header_titles = []
            self._page_header_descs = []
        self._page_header_titles.append(t)
        self._page_header_descs.append(d)
        self._style_page_header(t, d)
        return header

    @staticmethod
    def _style_page_header(title_label: QLabel, desc_label: QLabel):
        title_label.setStyleSheet(
            f"color: {T.TEXT_PRIMARY}; font-size: 16pt; font-weight: 800;"
        )
        desc_label.setStyleSheet(f"color: {T.TEXT_SECONDARY}; font-size: 10pt;")

    def _make_nav_button(self, icon: str, label: str) -> QPushButton:
        """Bouton de navigation latérale, checkable, au style premium."""
        btn = QPushButton(f"  {icon}   {label}")
        btn.setCheckable(True)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setMinimumHeight(42)
        self._style_nav_button(btn)
        return btn

    @staticmethod
    def _style_nav_button(btn: QPushButton):
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
        scroll_area.setStyleSheet(f"""
            QScrollArea {{ background: transparent; }}
            QScrollBar:vertical {{
                background: transparent;
                width: 10px;
                margin: 2px;
            }}
            QScrollBar::handle:vertical {{
                background: {T.BORDER_STRONG};
                border-radius: 5px;
                min-height: 30px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: {T.TEXT_MUTED};
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0;
            }}
        """)
        # Active le défilement tactile (glisser-déposer au doigt) sur le viewport
        QScroller.grabGesture(
            scroll_area.viewport(), QScroller.ScrollerGestureType.TouchGesture
        )

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
        self._action_bar = button_container
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
            "ExcelCompiler - compilateur Excel intelligent."
        ))

        self.about_text = QTextBrowser()
        self.about_text.setOpenExternalLinks(True)
        self._render_about()

        layout.addWidget(self.about_text)

        return tab

    def _render_about(self):
        """Rend le contenu HTML de l'onglet « À propos » aux couleurs du thème."""
        self.about_text.setHtml(f"""
        <h2 style='color: {T.ACCENT};'>ExcelCompiler v3.2</h2>
        <p><b>Compilateur Excel Intelligent avec Détection Automatique</b></p>

        <h3>Auteur</h3>
        <p>GOUNOU N'GOBI Chabi Zimé<br>
        Data Manager & Data Analyst</p>

        <h3>Nouveautés v3.2</h3>
        <ul>
            <li>✨ <b>Détection par fichier de référence</b> - Indiquez l'en-tête du premier fichier, les autres sont alignés automatiquement</li>
            <li>🧩 <b>Alignement des colonnes par libellé</b> - Fichiers aux colonnes désordonnées empilés correctement, jamais de corruption silencieuse</li>
            <li>🔗 <b>Cellules fusionnées propagées</b> et <b>en-têtes multi-lignes</b> aplatis proprement</li>
            <li>⚡ <b>Écriture Excel accélérée</b> - Compilations de gros volumes nettement plus rapides</li>
            <li>🛡️ <b>Protections</b> - Séparateur CSV détecté (;), classeurs multi-feuilles signalés, écrasement confirmé</li>
            <li>💾 <b>Options mémorisées</b> d'une session à l'autre, et <b>vérification de mise à jour</b> au démarrage</li>
        </ul>

        <h3>Fonctionnalités</h3>
        <ul>
            <li>Compilation de fichiers Excel (.xlsx, .xlsm), CSV et TSV</li>
            <li>Détection par fichier de référence, automatique ou manuelle</li>
            <li>Séparateur CSV détecté automatiquement (virgule, point-virgule, tabulation)</li>
            <li>Lignes de total / sous-total détectées (exclues ou marquées)</li>
            <li>Aperçu de détection avec correction par fichier avant compilation</li>
            <li>Export en Excel mis en forme, CSV ou TSV ; ajout du nom de fichier source</li>
        </ul>

        <h3>Utilisation</h3>
        <ol>
            <li>Sélectionnez un dossier ou glissez-déposez vos fichiers</li>
            <li>Choisissez le mode de détection (fichier de référence par défaut)</li>
            <li>Vérifiez avec « Aperçu », puis cliquez sur COMPILER</li>
            <li>Consultez les résultats et ouvrez directement le fichier produit</li>
        </ol>

        <hr>
        <p style='text-align: center; color: {T.TEXT_MUTED};'>
        © 2026 GOUNOU N'GOBI Chabi Zimé - Tous droits réservés
        </p>
        """)

    def apply_theme(self, *_):
        """Ré-applique le thème actif à toute la fenêtre.

        Reconstruit la feuille de style globale (QApplication) puis ré-applique
        les styles inline des éléments qui figent leurs couleurs à la
        construction (header, sidebar, footer, barre d'actions, contenus HTML).
        Connectée à ``ThemeManager.theme_changed``.
        """
        from PyQt6.QtWidgets import QApplication
        app = QApplication.instance()
        if app is not None:
            app.setStyleSheet(T.build_stylesheet())

        # Header (bandeau accent) + bouton de bascule
        self.header.setStyleSheet(f"background-color: {T.ACCENT};")
        self._style_theme_toggle()

        # Footer
        self.footer.setStyleSheet(
            f"background-color: {T.BG_SURFACE}; border-top: 1px solid {T.BORDER};")
        self.status_label.setStyleSheet(
            f"color: {T.TEXT_MUTED}; font-size: {T.FONT_SIZE_SM}pt;")

        # Sidebar
        self._sidebar_frame.setStyleSheet(
            f"QFrame {{ background-color: {T.BG_SURFACE}; "
            f"border: 1px solid {T.BORDER}; border-radius: {T.RADIUS_LG}px; }}"
        )
        self._sidebar_sign.setStyleSheet(
            f"color: {T.TEXT_MUTED}; font-size: 8pt; padding: 8px;")
        for btn in self._nav_buttons:
            self._style_nav_button(btn)

        # En-têtes de page
        for t, d in zip(self._page_header_titles, self._page_header_descs):
            self._style_page_header(t, d)

        # Barre d'actions de la compilation
        self._action_bar.setStyleSheet(
            f"QFrame {{ background-color: {T.BG_SURFACE}; "
            f"border-top: 1px solid {T.BORDER}; border-radius: 0px; }}"
        )

        # Contenu HTML « À propos »
        self._render_about()

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

        # Purger les overrides de fichiers qui ne sont plus sélectionnés.
        self._manual_overrides = {
            f: v for f, v in self._manual_overrides.items() if f in selected_files
        }

        options = self.options_widget.get_compilation_options()
        # Appliquer les corrections manuelles déjà choisies (priorité absolue).
        options.manual_overrides = dict(self._manual_overrides)
        options.column_aliases = dict(self._column_aliases)

        self.status_label.setText("Analyse de la détection en cours...")
        QApplication.setOverrideCursor(QCursor(Qt.CursorShape.WaitCursor))
        try:
            compiler = ExcelCompiler(options)
            previews = compiler.preview_detection(selected_files)
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

        # Le dialogue écrit les corrections dans options.manual_overrides ;
        # on les récupère ensuite pour la compilation.
        PreviewDialog(
            previews, parent=self,
            compiler=compiler, overrides=options.manual_overrides,
        ).exec()
        self._manual_overrides = dict(options.manual_overrides)
        # Récupérer les rattachements de colonnes éventuellement choisis.
        self._column_aliases = dict(compiler.options.column_aliases)

    def _open_path(self, path: str):
        """Ouvre un fichier ou dossier avec l'application par défaut de l'OS.

        Utilisé par les boutons « Ouvrir le fichier / le dossier » après une
        compilation réussie (G3). Échoue proprement (message) si l'ouverture
        n'est pas possible, sans crasher l'application.
        """
        from PyQt6.QtGui import QDesktopServices
        from PyQt6.QtCore import QUrl
        try:
            if not QDesktopServices.openUrl(QUrl.fromLocalFile(path)):
                raise RuntimeError("ouverture refusée par le système")
        except Exception as e:
            logger.warning(f"Impossible d'ouvrir {path}: {e}")
            QMessageBox.warning(
                self, "Ouverture impossible",
                f"Impossible d'ouvrir :\n{path}"
            )

    @staticmethod
    def _ensure_output_extension(filename: str, output_format) -> str:
        """Garantit que le nom de sortie porte l'extension du format choisi.

        Ex. format CSV mais nom « compilation.xlsx » -> « compilation.csv ».
        Évite d'écrire un contenu CSV/TSV dans un fichier .xlsx (illisible par
        Excel). Une extension déjà correcte est laissée telle quelle.
        """
        from core.compilation import OutputFormat
        ext_by_format = {
            OutputFormat.XLSX: '.xlsx',
            OutputFormat.CSV: '.csv',
            OutputFormat.TSV: '.tsv',
        }
        target_ext = ext_by_format.get(output_format, '.xlsx')
        stem = Path(filename).stem
        # Nom vide ou réduit à une extension (« .csv ») -> stem par défaut.
        if not stem or stem.startswith('.'):
            stem = 'compilation'
        return f"{stem}{target_ext}"

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

        # G4 — refuser une colonne de tri invalide (ex. « 1A ») plutôt que de
        # trier silencieusement sur la colonne A.
        if not self.options_widget.sort_column_is_valid():
            QMessageBox.warning(
                self, "Colonne de tri invalide",
                "La colonne de tri est invalide.\n\n"
                "Indiquez une lettre (A, B, …) ou un numéro (1, 2, …)."
            )
            return

        # Récupérer les options et le format
        options = self.options_widget.get_compilation_options()
        output_format = self.options_widget.get_output_format()

        # F1 — synchroniser l'extension du fichier de sortie avec le format
        # choisi : sinon on écrirait un contenu CSV dans un fichier .xlsx
        # (qu'Excel refuserait d'ouvrir). On corrige l'extension au besoin.
        output_filename = self._ensure_output_extension(output_filename, output_format)

        # Créer le chemin complet: même dossier que les fichiers sources
        source_directory = Path(selected_files[0]).parent
        output_file = str(source_directory / output_filename)

        # F3 — ne jamais compiler le fichier de sortie lui-même : s'il figure
        # dans la sélection (dossier déjà compilé auparavant), le retirer.
        # Sinon la sortie précédente serait réinjectée, doublant les données.
        output_resolved = Path(output_file).resolve()
        filtered = [f for f in selected_files
                    if Path(f).resolve() != output_resolved]
        if len(filtered) != len(selected_files):
            logger.info("Fichier de sortie retiré de la sélection (auto-inclusion évitée)")
            selected_files = filtered
        if not selected_files:
            QMessageBox.warning(
                self, "Aucun fichier",
                "Le seul fichier sélectionné est le fichier de sortie. "
                "Sélectionnez d'autres fichiers à compiler."
            )
            self.button_compile.setEnabled(True)
            return

        # F2 — confirmer l'écrasement si le fichier de sortie existe déjà.
        if Path(output_file).exists():
            reply = QMessageBox.question(
                self, "Fichier existant",
                f"Le fichier « {output_filename} » existe déjà dans ce dossier.\n\n"
                "Voulez-vous le remplacer ?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                self.button_compile.setEnabled(True)
                return

        # Appliquer les corrections manuelles d'en-tête choisies dans l'aperçu
        # (priorité absolue sur la détection), en ne gardant que les fichiers
        # toujours sélectionnés.
        options.manual_overrides = {
            f: v for f, v in self._manual_overrides.items() if f in selected_files
        }
        # Appliquer les rattachements manuels de colonnes (étape 6).
        options.column_aliases = dict(self._column_aliases)

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

            # Message de succès + accès direct au résultat (G3).
            output_file = result.output_file or ""
            output_filename = Path(output_file).name if output_file else "compilation.xlsx"
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Information)
            box.setWindowTitle("Compilation réussie")
            box.setText(
                f"La compilation est terminée !\n\n"
                f"• Fichiers : {result.successful_files}/{result.total_files}\n"
                f"• Lignes : {result.total_rows}\n"
                f"• Temps : {result.total_processing_time:.2f}s\n\n"
                f"Fichier créé : {output_filename}"
            )
            open_file_btn = None
            open_dir_btn = None
            if output_file and Path(output_file).exists():
                open_file_btn = box.addButton("Ouvrir le fichier",
                                              QMessageBox.ButtonRole.AcceptRole)
                open_dir_btn = box.addButton("Ouvrir le dossier",
                                             QMessageBox.ButtonRole.ActionRole)
            box.addButton("Fermer", QMessageBox.ButtonRole.RejectRole)
            box.exec()
            clicked = box.clickedButton()
            if clicked is open_file_btn and open_file_btn is not None:
                self._open_path(output_file)
            elif clicked is open_dir_btn and open_dir_btn is not None:
                self._open_path(str(Path(output_file).parent))

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
        # G1 — mémoriser les options pour le prochain lancement.
        try:
            self.options_widget.save_settings()
        except Exception:
            logger.warning("Impossible de mémoriser les options", exc_info=True)

        # Arrêter proprement la vérification de mise à jour si elle tourne encore
        # (appel réseau borné par timeout ; on n'attend que brièvement).
        if self._update_worker and self._update_worker.isRunning():
            try:
                self._update_worker.blockSignals(True)
                self._update_worker.wait(1500)
            except Exception:
                pass

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
                # Attendre la fin proprement, AVEC un délai borné: ne jamais
                # figer indéfiniment la fermeture si le worker est bloqué dans
                # une lecture longue (l'annulation est vérifiée entre étapes,
                # pas pendant une lecture I/O d'openpyxl). On ne tue PAS le
                # thread (terminate() corromprait un fichier en cours d'écriture)
                # — au pire le worker se terminera seul en arrière-plan.
                if not self.compilation_worker.wait(5000):
                    logger.warning(
                        "Worker toujours actif après 5s; fermeture sans attendre "
                        "davantage (le thread se terminera de lui-même)."
                    )
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()
