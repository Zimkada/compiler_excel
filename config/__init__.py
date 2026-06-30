"""
Module de configuration pour ExcelCompiler v3.2

Expose les constantes globales (config/constants.py). L'ancien gestionnaire
AppConfig a été retiré : il exposait des capacités (auto-save, backup,
crash_reporting, security_level…) qu'aucun module n'implémentait ni ne
consommait. Les seuls paramètres réellement appliqués vivent dans constants.py
et dans CompilationOptions.
"""

from .constants import *
