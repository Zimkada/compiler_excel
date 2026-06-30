# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file - Version Win10+ (PyQt6 moderne avec émojis)

a = Analysis(
    ['compiler_Win10Plus.py'],
    pathex=[],
    binaries=[
        # DLL Qt6 complètes pour Windows 10+
        ('compilexlsx_env_Win8Plus_Win10Plus/Lib/site-packages/PyQt6/Qt6/bin/Qt6Core.dll', '.'),
        ('compilexlsx_env_Win8Plus_Win10Plus/Lib/site-packages/PyQt6/Qt6/bin/Qt6Gui.dll', '.'),
        ('compilexlsx_env_Win8Plus_Win10Plus/Lib/site-packages/PyQt6/Qt6/bin/Qt6Widgets.dll', '.'),
        ('compilexlsx_env_Win8Plus_Win10Plus/Lib/site-packages/PyQt6/Qt6/bin/Qt6Network.dll', '.'),
        ('compilexlsx_env_Win8Plus_Win10Plus/Lib/site-packages/PyQt6/Qt6/bin/Qt6PrintSupport.dll', '.'),
        # DLL graphiques pour interface moderne
        ('compilexlsx_env_Win8Plus_Win10Plus/Lib/site-packages/PyQt6/Qt6/bin/d3dcompiler_47.dll', '.'),
        ('compilexlsx_env_Win8Plus_Win10Plus/Lib/site-packages/PyQt6/Qt6/bin/opengl32sw.dll', '.'),
        # DLL Visual C++
        ('compilexlsx_env_Win8Plus_Win10Plus/Lib/site-packages/PyQt6/Qt6/bin/msvcp140.dll', '.'),
        ('compilexlsx_env_Win8Plus_Win10Plus/Lib/site-packages/PyQt6/Qt6/bin/msvcp140_1.dll', '.'),
        ('compilexlsx_env_Win8Plus_Win10Plus/Lib/site-packages/PyQt6/Qt6/bin/msvcp140_2.dll', '.'),
        ('compilexlsx_env_Win8Plus_Win10Plus/Lib/site-packages/PyQt6/Qt6/bin/vcruntime140.dll', '.'),
        ('compilexlsx_env_Win8Plus_Win10Plus/Lib/site-packages/PyQt6/Qt6/bin/vcruntime140_1.dll', '.'),
    ],
    datas=[
        ('icon.ico', '.'), 
        ('folder.ico', '.'),
        # Plugins Qt complets pour interface riche
        ('compilexlsx_env_Win8Plus_Win10Plus/Lib/site-packages/PyQt6/Qt6/plugins/platforms/', 'platforms/'),
        ('compilexlsx_env_Win8Plus_Win10Plus/Lib/site-packages/PyQt6/Qt6/plugins/imageformats/', 'imageformats/'),
    ],
    hiddenimports=[
        # Core libraries
        'pandas', 'openpyxl', 'xlsxwriter', 'numpy', 'dateutil', 'dateutil.parser', 'dateutil.tz', 
        'six', 'pkg_resources', 'psutil',
        # PyQt6 imports
        'PyQt6.QtCore', 'PyQt6.QtGui', 'PyQt6.QtWidgets', 'PyQt6.QtPrintSupport', 'PyQt6.QtNetwork', 'PyQt6.sip',
        # Standard library
        'unittest.mock', 'concurrent.futures', 'logging.handlers', 'urllib.parse', 'dataclasses', 'contextlib',
        'threading', 'queue', 'platform', 'enum', 'configparser', 'pathlib', 'hashlib', 'mimetypes',
        'weakref', 'signal', 'tempfile', 'shutil', 'traceback', 'gc',
        # Sous-modules
        'openpyxl.styles', 'openpyxl.utils', 'openpyxl.workbook', 'openpyxl.worksheet',
        'pandas.plotting', 'pandas.io.common', 'pandas.io.excel', 'numpy.core', 'numpy.lib'
    ],
    hookspath=[], hooksconfig={}, runtime_hooks=[], excludes=[], noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz, a.scripts, [], exclude_binaries=True, name='ExcelCompiler_Win10Plus',
    debug=False, bootloader_ignore_signals=False, strip=False, upx=False, console=False,
    disable_windowed_traceback=False, argv_emulation=False, target_arch=None,
    codesign_identity=None, entitlements_file=None, icon=['icon.ico'],
)

coll = COLLECT(
    exe, a.binaries, a.zipfiles, a.datas, strip=False, upx=False, upx_exclude=[],
    name='ExcelCompiler_Win10Plus'
)