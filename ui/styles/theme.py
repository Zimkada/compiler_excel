"""
Design system premium — ExcelCompiler.

Source unique de vérité pour la direction artistique : palette, échelle
d'espacement, rayons, typographie et feuille de style globale (QSS) appliquée
à toute l'application. Style « premium clair épuré » (inspiration Notion /
Linear / Microsoft 365) avec le vert Excel comme couleur d'accent.

Réutilise et étend la palette historique de excel_theme.py (compatibilité).
"""

# ── Couleur d'accent (signature Excel) ──────────────────────────────────────
ACCENT = "#217346"           # vert Excel
ACCENT_HOVER = "#1B5E3A"     # accent plus profond (survol)
ACCENT_PRESSED = "#14462A"   # accent enfoncé
ACCENT_SOFT = "#E7F3EC"      # fond teinté accent (très clair)
ACCENT_BORDER = "#BFE0CB"    # bordure teintée accent

# ── Neutres (gris froids, lisibles) ─────────────────────────────────────────
BG_APP = "#F7F8FA"           # fond général de l'application
BG_SURFACE = "#FFFFFF"       # surfaces / cartes
BG_SUBTLE = "#F1F3F5"        # zones légèrement en retrait (hover discret)
BORDER = "#E6E8EB"           # bordures de cartes / séparateurs
BORDER_STRONG = "#D5D9DE"    # bordures de champs interactifs

TEXT_PRIMARY = "#1A1D21"     # titres / texte principal
TEXT_SECONDARY = "#5B636B"   # texte secondaire
TEXT_MUTED = "#8A929B"       # texte tertiaire / placeholder
TEXT_ON_ACCENT = "#FFFFFF"

# ── Sémantique ──────────────────────────────────────────────────────────────
SUCCESS = "#1E7E45"
SUCCESS_SOFT = "#E7F3EC"
WARNING = "#B7791F"
WARNING_SOFT = "#FEF4E2"
DANGER = "#C0392B"
DANGER_SOFT = "#FBEAE8"
INFO = "#2563EB"
INFO_SOFT = "#E8EFFD"

# ── Échelle d'espacement (pas de 4px) ───────────────────────────────────────
SPACE_XS = 4
SPACE_SM = 8
SPACE_MD = 16
SPACE_LG = 24
SPACE_XL = 32

# ── Rayons ──────────────────────────────────────────────────────────────────
RADIUS_SM = 6
RADIUS_MD = 10
RADIUS_LG = 14

# ── Typographie ─────────────────────────────────────────────────────────────
FONT_FAMILY = "Segoe UI"
FONT_SIZE_BASE = 10          # pt
FONT_SIZE_SM = 9
FONT_SIZE_LG = 12
FONT_SIZE_TITLE = 22
FONT_SIZE_SUBTITLE = 11


def build_stylesheet() -> str:
    """Construit la feuille de style globale appliquée à QApplication."""
    return f"""
    /* ===== Base ===== */
    QWidget {{
        background-color: {BG_APP};
        color: {TEXT_PRIMARY};
        font-family: "{FONT_FAMILY}";
        font-size: {FONT_SIZE_BASE}pt;
    }}

    QToolTip {{
        background-color: {TEXT_PRIMARY};
        color: {BG_SURFACE};
        border: none;
        padding: 6px 10px;
        border-radius: {RADIUS_SM}px;
        font-size: {FONT_SIZE_SM}pt;
    }}

    /* ===== Cartes / GroupBox ===== */
    QGroupBox {{
        background-color: {BG_SURFACE};
        border: 1px solid {BORDER};
        border-radius: {RADIUS_MD}px;
        margin-top: 16px;
        padding: 18px 16px 16px 16px;
        font-weight: 600;
    }}
    QGroupBox::title {{
        subcontrol-origin: margin;
        subcontrol-position: top left;
        left: 14px;
        top: 2px;
        padding: 0 4px;
        color: {TEXT_PRIMARY};
        font-size: {FONT_SIZE_BASE}pt;
        font-weight: 700;
    }}

    /* ===== Champs de saisie ===== */
    QLineEdit, QSpinBox, QComboBox {{
        background-color: {BG_SURFACE};
        border: 1px solid {BORDER_STRONG};
        border-radius: {RADIUS_SM}px;
        padding: 7px 10px;
        selection-background-color: {ACCENT_SOFT};
        selection-color: {TEXT_PRIMARY};
    }}
    QLineEdit:focus, QSpinBox:focus, QComboBox:focus {{
        border: 1px solid {ACCENT};
    }}
    QLineEdit:hover, QSpinBox:hover, QComboBox:hover {{
        border: 1px solid {ACCENT_BORDER};
    }}
    QLineEdit:disabled, QSpinBox:disabled, QComboBox:disabled {{
        background-color: {BG_SUBTLE};
        color: {TEXT_MUTED};
    }}
    QComboBox::drop-down {{
        border: none;
        width: 24px;
    }}
    QComboBox QAbstractItemView {{
        background-color: {BG_SURFACE};
        border: 1px solid {BORDER};
        border-radius: {RADIUS_SM}px;
        selection-background-color: {ACCENT_SOFT};
        selection-color: {TEXT_PRIMARY};
        outline: none;
        padding: 4px;
    }}
    QSpinBox::up-button, QSpinBox::down-button {{
        width: 18px;
        border: none;
        background: transparent;
    }}

    /* ===== Cases à cocher ===== */
    QCheckBox {{
        spacing: 8px;
        color: {TEXT_PRIMARY};
    }}
    QCheckBox::indicator {{
        width: 18px;
        height: 18px;
        border: 1px solid {BORDER_STRONG};
        border-radius: {RADIUS_SM}px;
        background-color: {BG_SURFACE};
    }}
    QCheckBox::indicator:hover {{
        border: 1px solid {ACCENT};
    }}
    QCheckBox::indicator:checked {{
        background-color: {ACCENT};
        border: 1px solid {ACCENT};
        image: none;
    }}

    /* ===== Boutons (par défaut: secondaire/neutre) ===== */
    QPushButton {{
        background-color: {BG_SURFACE};
        color: {TEXT_PRIMARY};
        border: 1px solid {BORDER_STRONG};
        border-radius: {RADIUS_SM}px;
        padding: 8px 16px;
        font-weight: 600;
    }}
    QPushButton:hover {{
        background-color: {BG_SUBTLE};
        border: 1px solid {ACCENT_BORDER};
    }}
    QPushButton:pressed {{
        background-color: {BORDER};
    }}
    QPushButton:disabled {{
        background-color: {BG_SUBTLE};
        color: {TEXT_MUTED};
        border: 1px solid {BORDER};
    }}

    /* Bouton primaire (propriété dynamique variant="primary") */
    QPushButton[variant="primary"] {{
        background-color: {ACCENT};
        color: {TEXT_ON_ACCENT};
        border: none;
        padding: 10px 22px;
    }}
    QPushButton[variant="primary"]:hover {{
        background-color: {ACCENT_HOVER};
    }}
    QPushButton[variant="primary"]:pressed {{
        background-color: {ACCENT_PRESSED};
    }}
    QPushButton[variant="primary"]:disabled {{
        background-color: {BORDER_STRONG};
        color: {BG_SURFACE};
    }}

    /* Bouton accent-outline (variant="accent") */
    QPushButton[variant="accent"] {{
        background-color: {BG_SURFACE};
        color: {ACCENT};
        border: 1.5px solid {ACCENT};
    }}
    QPushButton[variant="accent"]:hover {{
        background-color: {ACCENT_SOFT};
    }}

    /* ===== Listes ===== */
    QListWidget {{
        background-color: {BG_SURFACE};
        border: 1px solid {BORDER};
        border-radius: {RADIUS_MD}px;
        padding: 6px;
        outline: none;
    }}
    QListWidget::item {{
        padding: 7px 10px;
        border-radius: {RADIUS_SM}px;
        color: {TEXT_PRIMARY};
    }}
    QListWidget::item:hover {{
        background-color: {BG_SUBTLE};
    }}
    QListWidget::item:selected {{
        background-color: {ACCENT_SOFT};
        color: {TEXT_PRIMARY};
    }}

    /* ===== Barres de défilement (fines, discrètes) ===== */
    QScrollBar:vertical {{
        background: transparent;
        width: 10px;
        margin: 2px;
    }}
    QScrollBar::handle:vertical {{
        background: {BORDER_STRONG};
        border-radius: 5px;
        min-height: 30px;
    }}
    QScrollBar::handle:vertical:hover {{
        background: {TEXT_MUTED};
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0;
    }}
    QScrollBar:horizontal {{
        background: transparent;
        height: 10px;
        margin: 2px;
    }}
    QScrollBar::handle:horizontal {{
        background: {BORDER_STRONG};
        border-radius: 5px;
        min-width: 30px;
    }}
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
        width: 0;
    }}

    /* ===== Progress bar ===== */
    QProgressBar {{
        background-color: {BG_SUBTLE};
        border: none;
        border-radius: {RADIUS_SM}px;
        height: 10px;
        text-align: center;
        color: {TEXT_SECONDARY};
    }}
    QProgressBar::chunk {{
        background-color: {ACCENT};
        border-radius: {RADIUS_SM}px;
    }}

    /* ===== Text browsers (À propos, résultats) ===== */
    QTextBrowser {{
        background-color: {BG_SURFACE};
        border: 1px solid {BORDER};
        border-radius: {RADIUS_MD}px;
        padding: 12px;
    }}
    """
