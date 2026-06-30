"""
Design system premium — ExcelCompiler.

Source unique de vérité pour la direction artistique : palette, échelle
d'espacement, rayons, typographie et feuille de style globale (QSS) appliquée
à toute l'application. Style « premium » (inspiration Notion / Linear /
Microsoft 365) avec le vert Excel comme couleur d'accent.

Deux thèmes sont fournis : « light » (clair épuré) et « dark » (sombre moderne
et cohérent). Le thème actif est piloté par :class:`ThemeManager` ; les tokens
de couleur exposés au niveau module (``ACCENT``, ``BG_APP``, ``TEXT_PRIMARY``…)
sont rebindés à chaque changement de thème, de sorte que tout code lisant
``theme.TEXT_PRIMARY`` obtient toujours la valeur du thème courant.

Réutilise et étend la palette historique de excel_theme.py (compatibilité).
"""

from PyQt6.QtCore import QObject, pyqtSignal


# ── Tokens non chromatiques (identiques quel que soit le thème) ──────────────

# Échelle d'espacement (pas de 4px)
SPACE_XS = 4
SPACE_SM = 8
SPACE_MD = 16
SPACE_LG = 24
SPACE_XL = 32

# Rayons
RADIUS_SM = 6
RADIUS_MD = 10
RADIUS_LG = 14

# Typographie
FONT_FAMILY = "Segoe UI"
FONT_SIZE_BASE = 10          # pt
FONT_SIZE_SM = 9
FONT_SIZE_LG = 12
FONT_SIZE_TITLE = 22
FONT_SIZE_SUBTITLE = 11


# ── Palettes ─────────────────────────────────────────────────────────────────
# Chaque palette définit le même jeu de clés. Les tokens de couleur du module
# sont injectés depuis la palette active par ``set_theme``.

LIGHT_PALETTE = {
    # Accent (signature Excel)
    "ACCENT": "#217346",          # vert Excel
    "ACCENT_HOVER": "#1B5E3A",    # accent plus profond (survol)
    "ACCENT_PRESSED": "#14462A",  # accent enfoncé
    "ACCENT_SOFT": "#E7F3EC",     # fond teinté accent (très clair)
    "ACCENT_BORDER": "#BFE0CB",   # bordure teintée accent

    # Neutres (gris froids, lisibles)
    "BG_APP": "#F7F8FA",          # fond général de l'application
    "BG_SURFACE": "#FFFFFF",      # surfaces / cartes
    "BG_SUBTLE": "#F1F3F5",       # zones légèrement en retrait (hover discret)
    "BORDER": "#E6E8EB",          # bordures de cartes / séparateurs
    "BORDER_STRONG": "#D5D9DE",   # bordures de champs interactifs

    "TEXT_PRIMARY": "#1A1D21",    # titres / texte principal
    "TEXT_SECONDARY": "#5B636B",  # texte secondaire
    "TEXT_MUTED": "#8A929B",      # texte tertiaire / placeholder
    "TEXT_ON_ACCENT": "#FFFFFF",

    # Sémantique
    "SUCCESS": "#1E7E45",
    "SUCCESS_SOFT": "#E7F3EC",
    "WARNING": "#B7791F",
    "WARNING_SOFT": "#FEF4E2",
    "DANGER": "#C0392B",
    "DANGER_SOFT": "#FBEAE8",
    "INFO": "#2563EB",
    "INFO_SOFT": "#E8EFFD",
}

DARK_PALETTE = {
    # Accent : vert Excel un peu plus lumineux pour rester vif sur fond sombre
    "ACCENT": "#2EA065",          # vert Excel éclairci (contraste sur sombre)
    "ACCENT_HOVER": "#37B373",
    "ACCENT_PRESSED": "#268A57",
    "ACCENT_SOFT": "#16302450",   # voile vert translucide (sur surfaces sombres)
    "ACCENT_BORDER": "#2C5C42",   # bordure teintée accent, discrète

    # Neutres sombres (bleu-ardoise profond, pas de noir pur — plus premium)
    "BG_APP": "#0F1216",          # fond général
    "BG_SURFACE": "#181C22",      # surfaces / cartes
    "BG_SUBTLE": "#22272F",       # zones en retrait (hover discret)
    "BORDER": "#2A313A",          # bordures de cartes / séparateurs
    "BORDER_STRONG": "#3A434F",   # bordures de champs interactifs

    "TEXT_PRIMARY": "#ECEFF3",    # titres / texte principal
    "TEXT_SECONDARY": "#A9B2BD",  # texte secondaire
    "TEXT_MUTED": "#6B7480",      # texte tertiaire / placeholder
    "TEXT_ON_ACCENT": "#FFFFFF",

    # Sémantique (teintes vives lisibles sur fond sombre)
    "SUCCESS": "#3CCB7F",
    "SUCCESS_SOFT": "#16302450",
    "WARNING": "#E0A93B",
    "WARNING_SOFT": "#3A2E1450",
    "DANGER": "#F1675C",
    "DANGER_SOFT": "#3A1E1C50",
    "INFO": "#5B9BFF",
    "INFO_SOFT": "#1B2A4550",
}

PALETTES = {"light": LIGHT_PALETTE, "dark": DARK_PALETTE}

# Clés de couleur exposées comme tokens au niveau module.
_COLOR_KEYS = tuple(LIGHT_PALETTE.keys())

# Thème actif courant (nom). Rebindé par ``set_theme``.
current_theme = "light"


def set_theme(mode: str) -> None:
    """Active le thème ``mode`` ("light" ou "dark").

    Injecte les couleurs de la palette correspondante dans les tokens du
    module (``ACCENT``, ``BG_APP``…). Tout code lisant ``theme.TEXT_PRIMARY``
    après cet appel obtient la valeur du thème actif.
    """
    global current_theme
    if mode not in PALETTES:
        mode = "light"
    current_theme = mode
    palette = PALETTES[mode]
    globals().update(palette)


# Initialiser les tokens au thème par défaut dès l'import.
set_theme(current_theme)


class ThemeManager(QObject):
    """Hub central du thème : conserve le mode actif et notifie l'UI.

    - ``theme_changed`` est émis avec le nom du thème après chaque bascule.
    - Les widgets qui appliquent des styles inline se connectent à ce signal
      et ré-appliquent leurs styles via leur méthode ``apply_theme``.
    """

    theme_changed = pyqtSignal(str)

    def __init__(self):
        super().__init__()

    @property
    def mode(self) -> str:
        return current_theme

    def is_dark(self) -> bool:
        return current_theme == "dark"

    def set_mode(self, mode: str) -> None:
        """Change le thème et notifie les abonnés (idempotent)."""
        if mode == current_theme:
            return
        set_theme(mode)
        self.theme_changed.emit(current_theme)

    def toggle(self) -> str:
        """Bascule clair ↔ sombre et retourne le nouveau mode."""
        self.set_mode("light" if current_theme == "dark" else "dark")
        return current_theme


# Instance partagée par toute l'application.
manager = ThemeManager()


def build_stylesheet() -> str:
    """Construit la feuille de style globale appliquée à QApplication.

    Lit les tokens du module : appeler après ``set_theme`` (ou après une
    bascule via :class:`ThemeManager`) produit le QSS du thème actif.
    """
    return f"""
    /* ===== Base ===== */
    QWidget {{
        background-color: {BG_APP};
        color: {TEXT_PRIMARY};
        font-family: "{FONT_FAMILY}";
        font-size: {FONT_SIZE_BASE}pt;
    }}

    QToolTip {{
        background-color: {BG_SURFACE};
        color: {TEXT_PRIMARY};
        border: 1px solid {BORDER};
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
        color: {TEXT_PRIMARY};
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
        color: {TEXT_PRIMARY};
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

    /* ===== Text browsers / éditeurs (À propos, résultats, logs) ===== */
    QTextBrowser, QTextEdit {{
        background-color: {BG_SURFACE};
        color: {TEXT_PRIMARY};
        border: 1px solid {BORDER};
        border-radius: {RADIUS_MD}px;
        padding: 12px;
    }}

    /* ===== Dialogues ===== */
    QDialog {{
        background-color: {BG_APP};
    }}
    """
