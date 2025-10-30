"""
Constantes globales de l'application ExcelCompiler v3.2
"""

# Versions
APP_VERSION = "3.2"
APP_NAME = "Excel Compiler Pro"
APP_AUTHOR = "GOUNOU N'GOBI Chabi Zimé"

# Limites fichiers
MAX_FILE_SIZE = 104857600  # 100 MB
MAX_MEMORY_USAGE = 536870912  # 512 MB
CHUNK_SIZE = 10000

# Formats supportés
SUPPORTED_INPUT_FORMATS = ['.xlsx', '.xls', '.csv', '.tsv', '.txt']
SUPPORTED_OUTPUT_FORMATS = ['xlsx', 'csv', 'parquet', 'json']

# Sécurité
SECURITY_LEVEL = "high"
VALIDATE_FILE_INTEGRITY = True

# Détection structure - Seuils de confiance
CONFIDENCE_BORDER_DETECTION = 0.90
CONFIDENCE_DENSITY_DETECTION = 0.80
CONFIDENCE_PATTERN_DETECTION = 0.70
CONFIDENCE_FALLBACK = 0.60

# Détection anomalies - Seuils
OUTLIER_IQR_MULTIPLIER = 3.0
MISSING_VALUES_THRESHOLD = 0.20  # 20%
DUPLICATE_CHECK_ENABLED = True

# ML local
ML_LOCAL_MIN_TRAINING_SIZE = 10
ML_LOCAL_ENABLED_BY_DEFAULT = False

# UI - Palette Excel Native
EXCEL_PRIMARY_COLOR = "#217346"  # Vert Excel
EXCEL_HOVER_COLOR = "#1a5c37"
EXCEL_DISABLED_COLOR = "#c7c7c7"

# Performance
MAX_THREADS = 4
MAX_PREVIEW_ROWS = 200

# Logging
LOG_ROTATION_SIZE = 10 * 1024 * 1024  # 10MB
LOG_BACKUP_COUNT = 5
