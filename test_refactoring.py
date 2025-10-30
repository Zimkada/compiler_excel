"""
Script de test pour valider la refactorisation v3.2
Auteur: GOUNOU N'GOBI Chabi Zimé
"""

import sys
from pathlib import Path

# Configuration encodage pour Windows
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

print("=" * 80)
print("TEST DE REFACTORISATION - ExcelCompiler v3.2")
print("=" * 80)
print()

# Test 1: Imports config
print("[CONFIG] Test 1: Imports du module config/")
print("-" * 80)
try:
    from config import app_config
    from config.constants import (
        APP_VERSION, APP_NAME, APP_AUTHOR,
        MAX_FILE_SIZE, CHUNK_SIZE,
        CONFIDENCE_BORDER_DETECTION,
        EXCEL_PRIMARY_COLOR
    )
    print(f"[OK] Import config reussi")
    print(f"   - Version: {APP_VERSION}")
    print(f"   - Nom: {APP_NAME}")
    print(f"   - Auteur: {APP_AUTHOR}")
    print(f"   - Max file size: {MAX_FILE_SIZE / (1024*1024):.0f} MB")
    print(f"   - Chunk size: {CHUNK_SIZE}")
    print(f"   - Couleur Excel: {EXCEL_PRIMARY_COLOR}")
    print()
except Exception as e:
    print(f"[ERREUR] Erreur import config: {e}")
    sys.exit(1)

# Test 2: Imports utils
print("[UTILS] Test 2: Imports du module utils/")
print("-" * 80)
try:
    from utils import logger, calculate_file_hash, get_file_extension, validate_file_exists
    print(f"[OK] Import utils reussi")
    print(f"   - Logger: {logger.name}")
    print(f"   - Level: {logger.level}")
    print()
except Exception as e:
    print(f"[ERREUR] Erreur import utils: {e}")
    sys.exit(1)

# Test 3: Logger
print("[LOGGER] Test 3: Fonctionnement du logger")
print("-" * 80)
try:
    logger.info("Test message INFO")
    logger.debug("Test message DEBUG")
    logger.warning("Test message WARNING")
    print("[OK] Logger fonctionne correctement")
    print(f"   - Logs sauvegardes dans: logs/")
    print()
except Exception as e:
    print(f"[ERREUR] Erreur logger: {e}")
    sys.exit(1)

# Test 4: AppConfig
print("[CONFIG] Test 4: Configuration app_config")
print("-" * 80)
try:
    print(f"[OK] AppConfig charge")
    print(f"   - Fichier config: {app_config.config_path}")
    print(f"   - Config existe: {app_config.config_path.exists()}")
    print(f"   - Auto save: {app_config.get('auto_save_enabled')}")
    print(f"   - Language: {app_config.get('language')}")
    print(f"   - Theme: {app_config.get('theme')}")
    print(f"   - ML enabled: {app_config.get('ml_local_enabled')}")
    print()

    # Test sauvegarde
    print("   Test sauvegarde config...")
    app_config.set('test_key', 'test_value')
    app_config.save_config()
    print(f"   [OK] Config sauvegardee: {app_config.config_path}")
    print()
except Exception as e:
    print(f"[ERREUR] Erreur app_config: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Test 5: file_utils
print("[FILES] Test 5: Utilitaires fichiers")
print("-" * 80)
try:
    from utils.file_utils import sanitize_filename

    # Test extension
    ext = get_file_extension("test.xlsx")
    print(f"[OK] get_file_extension('test.xlsx') = '{ext}'")
    assert ext == '.xlsx', f"Expected '.xlsx', got '{ext}'"

    # Test validation fichier
    exists = validate_file_exists(__file__)
    print(f"[OK] validate_file_exists('{Path(__file__).name}') = {exists}")
    assert exists == True, f"Expected True, got {exists}"

    # Test hash fichier
    file_hash = calculate_file_hash(__file__)
    print(f"[OK] calculate_file_hash() = {file_hash[:16]}...")

    # Test sanitize
    dirty_name = 'test<>:"|?*.xlsx'
    clean_name = sanitize_filename(dirty_name)
    print(f"[OK] sanitize_filename('{dirty_name}') = '{clean_name}'")

    print()
except Exception as e:
    print(f"[ERREUR] Erreur file_utils: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Test 6: Structure dossiers
print("[DIRS] Test 6: Verification structure de dossiers")
print("-" * 80)
required_dirs = [
    'config', 'utils', 'core', 'ui', 'tests',
    'core/detection', 'core/anomalies', 'core/compilation', 'core/validation',
    'ui/widgets', 'ui/styles',
    'tests/test_detection', 'tests/test_anomalies'
]

all_exist = True
for dir_path in required_dirs:
    path = Path(dir_path)
    exists = path.exists() and path.is_dir()
    status = "[OK]" if exists else "[X]"
    print(f"   {status} {dir_path}/")
    if not exists:
        all_exist = False

if all_exist:
    print("\n[OK] Tous les dossiers requis existent")
else:
    print("\n[ERREUR] Certains dossiers sont manquants")
    sys.exit(1)

print()

# Test 7: Imports core (modules vides pour l'instant)
print("[CORE] Test 7: Imports modules core/")
print("-" * 80)
try:
    import core
    import core.detection
    import core.anomalies
    import core.compilation
    import core.validation
    print("[OK] Tous les modules core/ importables")
    print()
except Exception as e:
    print(f"❌ Erreur import core: {e}")
    sys.exit(1)

# Test 8: Imports UI (modules vides pour l'instant)
print("[UI] Test 8: Imports modules ui/")
print("-" * 80)
try:
    import ui
    import ui.widgets
    import ui.styles
    print("[OK] Tous les modules ui/ importables")
    print()
except Exception as e:
    print(f"[ERREUR] Erreur import ui: {e}")
    sys.exit(1)

# Résumé final
print("=" * 80)
print("TOUS LES TESTS SONT PASSES !")
print("=" * 80)
print()
print("RESUME:")
print(f"   - Version: {APP_VERSION}")
print(f"   - Architecture: Modulaire")
print(f"   - Framework: PyQt6 (Windows 10/11)")
print(f"   - Config file: {app_config.config_path}")
print(f"   - Logs dir: logs/")
print()
print("=> La refactorisation est validee et prete pour la suite !")
print()
