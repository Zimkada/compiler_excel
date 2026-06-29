"""
Styles et thèmes pour l'UI.

- excel_theme : palette historique (conservée pour compatibilité).
- theme : design system premium (palette raffinée, tokens, QSS global).
"""

from .excel_theme import *  # noqa: F401,F403  (compat: EXCEL_GREEN, WHITE, ...)
from . import theme  # noqa: F401
from .theme import build_stylesheet  # noqa: F401

__all__ = [
    'EXCEL_GREEN', 'EXCEL_GREEN_HOVER', 'EXCEL_GREEN_PRESSED',
    'EXCEL_GREEN_LIGHT', 'OFFICE_BLUE', 'OFFICE_ORANGE',
    'OFFICE_YELLOW', 'NEUTRAL_DARK', 'WHITE',
    'theme', 'build_stylesheet',
]
