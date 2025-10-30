"""
Configuration centralisée du logging
"""

import logging
import sys
from pathlib import Path
from datetime import datetime
from logging.handlers import RotatingFileHandler

def setup_logger(name: str = "ExcelCompiler", level: str = "INFO") -> logging.Logger:
    """
    Configure et retourne un logger

    Args:
        name: Nom du logger
        level: Niveau de log (DEBUG, INFO, WARNING, ERROR, CRITICAL)

    Returns:
        Logger configuré
    """
    logger = logging.getLogger(name)

    # Éviter duplication si déjà configuré
    if logger.handlers:
        return logger

    logger.setLevel(getattr(logging, level.upper()))

    # Format
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # Handler console
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # Handler fichier avec rotation
    log_dir = Path.cwd() / "logs"
    log_dir.mkdir(exist_ok=True)

    log_file = log_dir / f"excelcompiler_{datetime.now().strftime('%Y%m%d')}.log"
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=10*1024*1024,  # 10MB
        backupCount=5,
        encoding='utf-8'
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    return logger

# Logger global
logger = setup_logger()
