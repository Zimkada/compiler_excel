# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file - Version Win8+ (PyQt5 stable sans émojis)

a = Analysis(
    ['compiler_Win8Plus.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('icon.ico', '.'), 
        ('folder.ico', '.'),
    ],
    hiddenimports=[
        # Core libraries
        'pandas', 'openpyxl', 'xlsxwriter', 'numpy', 'dateutil', 'dateutil.parser', 'dateutil.tz', 
        'six', 'pkg_resources', 'psutil',
        # PyQt5 imports UNIQUEMENT
        'PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtWidgets', 'PyQt5.QtPrintSupport', 'PyQt5.sip',
        # Standard library
        'unittest.mock', 'concurrent.futures', 'logging.handlers', 'urllib.parse', 'dataclasses', 'contextlib',
        'threading', 'queue', 'platform', 'enum', 'configparser', 'pathlib', 'hashlib', 'mimetypes',
        'weakref', 'signal', 'tempfile', 'shutil', 'traceback', 'gc',
        # Sous-modules
        'openpyxl.styles', 'openpyxl.utils', 'openpyxl.workbook', 'openpyxl.worksheet',
        'pandas.plotting', 'pandas.io.common', 'pandas.io.excel', 'numpy.core', 'numpy.lib'
    ],
    hookspath=[], hooksconfig={}, runtime_hooks=[], 
    excludes=['PyQt6', 'PyQt6.QtCore', 'PyQt6.QtGui', 'PyQt6.QtWidgets'], # Exclure PyQt6
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz, a.scripts, [], exclude_binaries=True, name='ExcelCompiler_Win8Plus',
    debug=False, bootloader_ignore_signals=False, strip=False, upx=False, console=False,
    disable_windowed_traceback=False, argv_emulation=False, target_arch=None,
    codesign_identity=None, entitlements_file=None, icon=['icon.ico'],
)

coll = COLLECT(
    exe, a.binaries, a.zipfiles, a.datas, strip=False, upx=False, upx_exclude=[],
    name='ExcelCompiler_Win8Plus'
)