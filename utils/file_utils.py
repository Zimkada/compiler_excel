"""
Utilitaires pour manipulation de fichiers
"""

import hashlib
from pathlib import Path
from typing import Optional

def calculate_file_hash(file_path: str, algorithm: str = 'sha256') -> str:
    """
    Calcule le hash d'un fichier

    Args:
        file_path: Chemin du fichier
        algorithm: Algorithme (md5, sha1, sha256)

    Returns:
        Hash hexadécimal
    """
    hash_func = hashlib.new(algorithm)

    with open(file_path, 'rb') as f:
        for chunk in iter(lambda: f.read(4096), b''):
            hash_func.update(chunk)

    return hash_func.hexdigest()

def get_file_extension(file_path: str) -> str:
    """Retourne l'extension du fichier (avec le point)"""
    return Path(file_path).suffix.lower()

def validate_file_exists(file_path: str) -> bool:
    """Vérifie qu'un fichier existe"""
    return Path(file_path).is_file()

def get_file_size_mb(file_path: str) -> float:
    """Retourne la taille du fichier en MB"""
    return Path(file_path).stat().st_size / (1024 * 1024)

def sanitize_filename(filename: str, max_length: int = 255) -> str:
    """
    Nettoie un nom de fichier pour éviter les caractères interdits

    Args:
        filename: Nom de fichier à nettoyer
        max_length: Longueur maximale

    Returns:
        Nom de fichier nettoyé
    """
    # Caractères interdits dans les noms de fichiers Windows
    forbidden_chars = ['<', '>', ':', '"', '|', '?', '*', '\\', '/', '\0']

    cleaned = filename
    for char in forbidden_chars:
        cleaned = cleaned.replace(char, '_')

    # Limiter la longueur
    if len(cleaned) > max_length:
        name, ext = Path(cleaned).stem, Path(cleaned).suffix
        max_name_length = max_length - len(ext)
        cleaned = name[:max_name_length] + ext

    return cleaned
