# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec - Application modulaire ExcelCompiler v3.2 (point d'entrée app.py)
#
# Remplace l'ancien ExcelCompiler_Win10Plus.spec qui empaquetait le monolithe
# legacy compiler_Win10Plus.py (archivé dans v3.1_legacy/). Ce spec construit la
# véritable application modulaire (core/, ui/, config/, utils/) — celle couverte
# par la suite de tests.
#
# Les DLL Qt6 et plugins sont collectés automatiquement par le hook PyQt6 de
# PyInstaller 6.x : nul besoin de les lister en dur (l'ancien spec pointait vers
# un venv absent du dépôt, build non reproductible).

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Ressources lues via utils.resource_path() -> extraites à la racine
        # du bundle (sys._MEIPASS). Garder ces noms synchronisés avec le code.
        ('icon.ico', '.'),
        ('icon.png', '.'),
        ('folder.ico', '.'),
    ],
    hiddenimports=[
        # Modules applicatifs (importés dynamiquement par endroits via importlib
        # implicite des chaînes de packages — on les déclare par sécurité).
        'core', 'core.compilation', 'core.detection',
        'ui', 'ui.widgets', 'ui.workers', 'ui.styles',
        'config', 'utils',
        # Dépendances tierces du moteur.
        'pandas', 'openpyxl', 'numpy', 'dateutil', 'dateutil.parser',
        'openpyxl.styles', 'openpyxl.utils', 'openpyxl.worksheet',
        # PyQt6.
        'PyQt6.QtCore', 'PyQt6.QtGui', 'PyQt6.QtWidgets', 'PyQt6.sip',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Exclure les libs lourdes et inutiles au modulaire pour alléger le bundle.
    excludes=['xlsxwriter', 'psutil', 'tkinter', 'matplotlib', 'PyQt5'],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz, a.scripts, [], exclude_binaries=True, name='ExcelCompiler',
    debug=False, bootloader_ignore_signals=False, strip=False, upx=False,
    console=False, disable_windowed_traceback=False, argv_emulation=False,
    target_arch=None, codesign_identity=None, entitlements_file=None,
    icon=['icon.ico'],
    # Métadonnées de version embarquées dans l'exe (clic droit > Propriétés).
    version='version_info.txt',
)

coll = COLLECT(
    exe, a.binaries, a.zipfiles, a.datas, strip=False, upx=False,
    upx_exclude=[], name='ExcelCompiler',
)
