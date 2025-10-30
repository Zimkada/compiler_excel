"""
Configuration de l'application chargée depuis excelcompiler_config.json
"""

import json
from pathlib import Path
from typing import Dict, Any
from .constants import *

class AppConfig:
    """Gestionnaire de configuration application"""

    CONFIG_FILE = "excelcompiler_config.json"

    def __init__(self):
        self.config_path = Path.cwd() / self.CONFIG_FILE
        self.config = self._load_config()

    def _load_config(self) -> Dict[str, Any]:
        """Charge la configuration depuis JSON"""
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Erreur chargement config: {e}. Utilisation config par défaut.")
                return self._get_default_config()
        else:
            return self._get_default_config()

    def _get_default_config(self) -> Dict[str, Any]:
        """Configuration par défaut"""
        return {
            "auto_save_enabled": True,
            "auto_save_interval": 300,
            "backup_enabled": True,
            "max_backup_files": 10,
            "max_memory_usage": MAX_MEMORY_USAGE,
            "max_threads": MAX_THREADS,
            "chunk_size": CHUNK_SIZE,
            "security_level": SECURITY_LEVEL,
            "validate_file_integrity": VALIDATE_FILE_INTEGRITY,
            "max_file_size": MAX_FILE_SIZE,
            "language": "fr",
            "theme": "excel_native",
            "show_preview": True,
            "enable_validation": True,
            "health_checks_enabled": True,
            "logging_level": "INFO",
            "metrics_retention_days": 7,
            "auto_recovery_enabled": True,
            "crash_reporting": True,
            "recovery_timeout": 30,
            "ml_local_enabled": ML_LOCAL_ENABLED_BY_DEFAULT
        }

    def save_config(self):
        """Sauvegarde la configuration"""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Erreur sauvegarde config: {e}")

    def get(self, key: str, default: Any = None) -> Any:
        """Récupère une valeur de configuration"""
        return self.config.get(key, default)

    def set(self, key: str, value: Any):
        """Définit une valeur de configuration"""
        self.config[key] = value

# Instance globale
app_config = AppConfig()
