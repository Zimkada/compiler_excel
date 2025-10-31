"""
Classe abstraite et structures de données pour la détection intelligente de structure Excel
Auteur: GOUNOU N'GOBI Chabi Zimé
Version: 3.2
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any
from pathlib import Path
import pandas as pd


@dataclass
class DetectionResult:
    """
    Résultat de détection de structure pour un fichier Excel

    Contient toutes les informations détectées sur la structure d'un tableau
    dans un fichier (en-têtes, données, infos avant/après)
    """

    # Identification du fichier
    file_path: str

    # ===== Informations AVANT tableau (titre, contexte) =====
    has_leading_content: bool = False
    leading_start_row: int = 1  # Toujours ligne 1 en Excel
    leading_end_row: Optional[int] = None  # Dernière ligne avant en-têtes
    leading_preview: List[str] = field(default_factory=list)  # Aperçu des lignes

    # ===== Structure du TABLEAU (en-têtes) =====
    header_start_row: int = 1  # Ligne où commencent les en-têtes (base 1, Excel)
    header_rows: int = 1  # Nombre de lignes d'en-tête
    detected_headers: List[str] = field(default_factory=list)  # En-têtes détectés

    # ===== DONNÉES =====
    data_start_row: int = 2  # Ligne où commencent les données (calculé auto)
    data_end_row: int = 0  # Dernière ligne de données

    # ===== Informations APRÈS tableau (signature, date) =====
    has_trailing_content: bool = False
    trailing_start_row: Optional[int] = None
    trailing_end_row: Optional[int] = None
    trailing_row_count: int = 0
    trailing_preview: List[str] = field(default_factory=list)
    trailing_detection_reason: Optional[str] = None  # "density_drop", "empty_lines", etc.

    # ===== Méta-données de détection =====
    confidence: float = 0.0  # Score de confiance (0.0 à 1.0)
    detection_method: str = "unknown"  # "border", "density", "pattern", "hybrid"
    total_rows: int = 0  # Nombre total de lignes du fichier

    # ===== Validation croisée (comparaison avec autres fichiers) =====
    cross_validation_score: Optional[float] = None  # Similarité avec autres fichiers
    warning: Optional[str] = None  # Avertissement si détection suspecte

    # ===== Debug et traçabilité =====
    debug_info: Dict[str, Any] = field(default_factory=dict)

    def __post_init__(self):
        """Calculs automatiques après initialisation"""
        # Calcul automatique de data_start_row
        if self.data_start_row == 2:  # Valeur par défaut
            self.data_start_row = self.header_start_row + self.header_rows

        # Calcul des lignes leading
        if self.header_start_row > 1:
            self.has_leading_content = True
            self.leading_end_row = self.header_start_row - 1

    def to_dict(self) -> Dict[str, Any]:
        """Convertit en dictionnaire pour sérialisation"""
        return {
            'file_path': self.file_path,
            'header_start_row': self.header_start_row,
            'header_rows': self.header_rows,
            'data_start_row': self.data_start_row,
            'data_end_row': self.data_end_row,
            'confidence': self.confidence,
            'detection_method': self.detection_method,
            'has_leading_content': self.has_leading_content,
            'has_trailing_content': self.has_trailing_content,
            'detected_headers': self.detected_headers,
            'cross_validation_score': self.cross_validation_score,
            'warning': self.warning
        }

    def is_reliable(self, threshold: float = 0.75) -> bool:
        """Vérifie si la détection est fiable"""
        return self.confidence >= threshold

    def get_summary(self) -> str:
        """Retourne un résumé lisible de la détection"""
        summary = [
            f"Fichier: {Path(self.file_path).name}",
            f"En-têtes: lignes {self.header_start_row}-{self.header_start_row + self.header_rows - 1}",
            f"Données: lignes {self.data_start_row}-{self.data_end_row}",
            f"Confiance: {self.confidence:.0%}",
            f"Méthode: {self.detection_method}"
        ]

        if self.has_leading_content:
            summary.append(f"Infos avant: lignes 1-{self.leading_end_row}")

        if self.has_trailing_content:
            summary.append(f"Infos après: lignes {self.trailing_start_row}+ ({self.trailing_row_count} lignes)")

        if self.warning:
            summary.append(f"⚠️ {self.warning}")

        return "\n".join(summary)


class BaseDetector(ABC):
    """
    Classe abstraite pour tous les détecteurs de structure

    Tous les détecteurs (BorderDetector, DensityDetector, etc.)
    héritent de cette classe et implémentent la méthode detect()
    """

    def __init__(self, confidence_threshold: float = 0.5):
        """
        Initialise le détecteur

        Args:
            confidence_threshold: Seuil de confiance minimum (0.0 à 1.0)
        """
        self.confidence_threshold = confidence_threshold
        self.name = self.__class__.__name__

    @abstractmethod
    def detect(self, file_path: str, df: Optional[pd.DataFrame] = None) -> DetectionResult:
        """
        Détecte la structure d'un fichier Excel

        Args:
            file_path: Chemin du fichier à analyser
            df: DataFrame pandas optionnel (si déjà chargé, pour performance)

        Returns:
            DetectionResult avec toutes les informations détectées
        """
        pass

    def load_file(self, file_path: str, nrows: Optional[int] = None) -> pd.DataFrame:
        """
        Charge un fichier Excel/CSV en DataFrame

        Args:
            file_path: Chemin du fichier
            nrows: Nombre maximum de lignes à charger (None = tout)

        Returns:
            DataFrame pandas
        """
        ext = Path(file_path).suffix.lower()

        try:
            if ext in ['.xlsx', '.xls', '.xlsm']:
                df = pd.read_excel(file_path, header=None, nrows=nrows)
            elif ext == '.csv':
                df = pd.read_csv(file_path, header=None, nrows=nrows, encoding='utf-8-sig')
            elif ext in ['.tsv', '.txt']:
                df = pd.read_csv(file_path, header=None, nrows=nrows,
                                sep='\t', encoding='utf-8-sig')
            else:
                raise ValueError(f"Format non supporté: {ext}")

            return df

        except Exception as e:
            raise RuntimeError(f"Erreur chargement {file_path}: {e}")

    def extract_headers(self, df: pd.DataFrame, header_start_row: int,
                       header_rows: int) -> List[str]:
        """
        Extrait les en-têtes d'un DataFrame

        Args:
            df: DataFrame
            header_start_row: Ligne de début (base 1, Excel)
            header_rows: Nombre de lignes d'en-tête

        Returns:
            Liste des en-têtes (combinés si multi-lignes)
        """
        # Conversion Excel (base 1) vers pandas (base 0)
        start_idx = header_start_row - 1
        end_idx = start_idx + header_rows

        headers = []
        for i in range(start_idx, min(end_idx, len(df))):
            row = df.iloc[i]
            headers.append([str(v) if pd.notna(v) else "" for v in row])

        # Si multi-lignes, combiner les en-têtes
        if len(headers) > 1:
            combined = []
            for col_idx in range(len(headers[0])):
                parts = [headers[row][col_idx] for row in range(len(headers))]
                combined_header = " - ".join([p for p in parts if p]).strip()
                combined.append(combined_header if combined_header else f"Col{col_idx+1}")
            return combined
        elif len(headers) == 1:
            return [h if h else f"Col{i+1}" for i, h in enumerate(headers[0])]
        else:
            return []

    def calculate_row_density(self, df: pd.DataFrame, row_idx: int) -> float:
        """
        Calcule la densité d'une ligne (proportion de cellules non vides)

        Args:
            df: DataFrame
            row_idx: Index de la ligne (base 0, pandas)

        Returns:
            Densité entre 0.0 et 1.0
        """
        if row_idx >= len(df):
            return 0.0

        row = df.iloc[row_idx]
        non_empty = row.notna().sum()
        total = len(row)

        return non_empty / total if total > 0 else 0.0

    def is_likely_header_row(self, df: pd.DataFrame, row_idx: int) -> bool:
        """
        Détermine si une ligne ressemble à un en-tête

        Critères:
        - Densité élevée (>50%)
        - Plus de texte que de chiffres
        - Pas de ligne complètement vide

        Args:
            df: DataFrame
            row_idx: Index de la ligne

        Returns:
            True si c'est probablement un en-tête
        """
        if row_idx >= len(df):
            return False

        row = df.iloc[row_idx]

        # Critère 1: Densité
        density = self.calculate_row_density(df, row_idx)
        if density < 0.5:
            return False

        # Critère 2: Plus de texte que de chiffres
        text_count = 0
        number_count = 0

        for val in row:
            if pd.notna(val):
                try:
                    float(val)
                    number_count += 1
                except (ValueError, TypeError):
                    text_count += 1

        return text_count >= number_count

    def __repr__(self) -> str:
        return f"{self.name}(threshold={self.confidence_threshold})"
