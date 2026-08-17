"""
Détecteur basé sur un fichier de référence (mode semi-automatique)
Auteur: GOUNOU N'GOBI Chabi Zimé
Version: 3.2.1

Ce détecteur utilise les en-têtes d'un fichier de référence pour
détecter automatiquement les en-têtes dans les autres fichiers.
"""

from typing import Optional, List, Dict, Tuple
from pathlib import Path
import pandas as pd
import numpy as np
import re

from .base_detector import BaseDetector, DetectionResult


class ReferenceDetector(BaseDetector):
    """
    Détecteur qui utilise un fichier de référence pour détecter les autres

    Mode d'utilisation:
    1. L'utilisateur spécifie la ligne d'en-tête du premier fichier
    2. Le détecteur extrait les en-têtes de référence
    3. Pour les autres fichiers, il cherche la ligne la plus similaire

    Avantages:
    - Simple et intuitif
    - Fiable (l'utilisateur donne la bonne réponse)
    - Gère les cas mono et multi-lignes d'en-têtes
    """

    def __init__(self,
                 reference_file: str,
                 reference_header_row: int,
                 reference_header_lines: int = 1,
                 similarity_threshold: float = 0.95,
                 min_similarity: float = 0.70,
                 max_search_rows: int = 50,
                 empty_block_threshold: int = 3):
        """
        Initialise le détecteur avec référence

        Args:
            reference_file: Chemin du fichier de référence
            reference_header_row: Ligne d'en-tête (base 1, Excel)
            reference_header_lines: Nombre de lignes d'en-tête
            similarity_threshold: Seuil pour accepter automatiquement (0.95 = 95%)
            min_similarity: Seuil minimum pour accepter avec validation (0.70 = 70%)
            max_search_rows: Nombre max de lignes à chercher (50 par défaut)
            empty_block_threshold: Nombre de lignes vides consécutives qui
                marquent la fin du tableau (3 par défaut). Une ligne vide isolée
                ne termine pas les données.
        """
        super().__init__(confidence_threshold=0.5)

        self.reference_file = reference_file
        self.reference_header_row = reference_header_row
        self.reference_header_lines = reference_header_lines
        self.similarity_threshold = similarity_threshold
        self.min_similarity = min_similarity
        self.max_search_rows = max_search_rows
        self.empty_block_threshold = max(1, empty_block_threshold)

        # Extraire les en-têtes de référence
        self.reference_headers = self._extract_reference_headers()
        self.reference_words = self._extract_words(self.reference_headers)
        # Avertissement si la ligne indiquée ne ressemble pas à un en-tête.
        self.reference_warning = self._check_reference_plausibility()

    def _extract_reference_headers(self) -> List[str]:
        """
        Extrait les en-têtes du fichier de référence

        Returns:
            Liste des en-têtes de référence
        """
        try:
            df = self.load_file(self.reference_file)
            headers = self.extract_headers(
                df,
                self.reference_header_row,
                self.reference_header_lines
            )
            return headers
        except Exception as e:
            raise RuntimeError(
                f"Impossible d'extraire les en-têtes de référence: {e}"
            )

    def _check_reference_plausibility(self) -> Optional[str]:
        """Avertit si la ligne indiquée ne ressemble pas à une ligne d'en-tête.

        Un tableau a plusieurs colonnes, donc plusieurs libellés. Quand la ligne
        choisie n'en porte qu'un ou deux alors que le tableau est plus large,
        c'est presque toujours le TITRE du document (« SUIVI DES PLANTS DES
        CAMPAGNES… », souvent fusionné sur toute la largeur) et non l'en-tête.

        Ce cas est particulièrement traître : tous les fichiers du lot portent
        le même titre, la similarité affiche donc 100 % et rien ne signale
        l'erreur. La compilation « réussit » en produisant un tableau dont les
        vrais en-têtes sont devenus des lignes de données.

        Renvoie le message d'avertissement, ou None si la ligne est plausible.
        """
        # extract_headers remplace une colonne SANS libellé par « ColN ». Ces
        # placeholders sont précisément le signal recherché : ils ne comptent
        # donc pas comme de vrais libellés.
        placeholder = re.compile(r"^col\d+$", re.IGNORECASE)
        labels = [
            str(h).strip() for h in (self.reference_headers or [])
            if h is not None and str(h).strip()
            and str(h).strip().lower() != 'nan'
            and not placeholder.match(str(h).strip())
        ]
        n_labels = len(labels)
        n_cols = len(self.reference_headers or [])

        # Un tableau étroit (2-3 colonnes) peut légitimement n'avoir qu'un ou
        # deux libellés : on ne se prononce qu'à partir de 4 colonnes.
        if n_cols < 4 or n_labels == 0 or n_labels > 2:
            return None

        apercu = " / ".join(labels[:2])
        return (
            f"La ligne {self.reference_header_row} ne porte que {n_labels} "
            f"libellé(s) « {apercu} » pour {n_cols} colonnes : il s'agit "
            f"probablement du titre du document, pas de la ligne d'en-tête. "
            f"Vérifiez le numéro de ligne (les en-têtes sont souvent plus bas), "
            f"sinon vos en-têtes réels seront compilés comme des données."
        )

    def detect(self, file_path: str, df: Optional[pd.DataFrame] = None) -> DetectionResult:
        """
        Détecte la structure en cherchant une ligne similaire à la référence

        Args:
            file_path: Chemin du fichier à analyser
            df: DataFrame optionnel

        Returns:
            DetectionResult avec validation si similarité < 95%
        """
        try:
            # Si c'est le fichier de référence, retourner directement
            if Path(file_path).resolve() == Path(self.reference_file).resolve():
                return self._create_reference_result(file_path, df)

            # Charger le fichier
            if df is None:
                df = self.load_file(file_path)

            # Chercher la meilleure correspondance
            best_match = self._find_best_match(df)

            if best_match is None:
                return self._create_failed_result(
                    file_path,
                    f"Aucune ligne similaire trouvée (similarité max < {self.min_similarity:.0%})"
                )

            row_num, similarity, headers = best_match

            # Créer le résultat selon la similarité
            result = self._create_result_from_match(
                file_path, df, row_num, similarity, headers
            )

            return result

        except Exception as e:
            return self._create_failed_result(
                file_path,
                f"Erreur détection par référence: {e}"
            )

    def _create_reference_result(self, file_path: str,
                                 df: Optional[pd.DataFrame] = None) -> DetectionResult:
        """
        Crée le résultat pour le fichier de référence lui-même

        Args:
            file_path: Chemin du fichier de référence
            df: DataFrame optionnel

        Returns:
            DetectionResult avec confiance 1.0
        """
        if df is None:
            df = self.load_file(file_path)

        # Trouver la fin des données
        data_start = self.reference_header_row + self.reference_header_lines
        data_end = self._find_data_end(df, data_start - 1)  # base 0

        return DetectionResult(
            file_path=file_path,
            header_start_row=self.reference_header_row,
            header_rows=self.reference_header_lines,
            detected_headers=self.reference_headers,
            data_start_row=data_start,
            data_end_row=data_end,
            confidence=1.0,
            detection_method="reference",
            total_rows=len(df),
            debug_info={
                'is_reference_file': True,
                'detection_reason': 'reference_file'
            }
        )

    def _find_best_match(self, df: pd.DataFrame) -> Optional[Tuple[int, float, List[str]]]:
        """
        Trouve la ligne la plus similaire aux en-têtes de référence

        Args:
            df: DataFrame

        Returns:
            Tuple (row_num, similarity, headers) ou None si aucune ligne acceptable
        """
        best_row = None
        best_similarity = 0.0
        best_headers = []

        # Chercher dans les N premières lignes
        max_rows = min(self.max_search_rows, len(df))

        for row_idx in range(max_rows):
            row_num = row_idx + 1  # Base 1 (Excel)

            # Extraire les en-têtes de cette ligne
            headers = self.extract_headers(df, row_num, self.reference_header_lines)

            # Calculer la similarité
            similarity = self._calculate_similarity(headers)

            # Si similarité ≥ threshold, accepter immédiatement (optimisation)
            if similarity >= self.similarity_threshold:
                return (row_num, similarity, headers)

            # Sinon, garder le meilleur candidat
            if similarity > best_similarity:
                best_similarity = similarity
                best_row = row_num
                best_headers = headers

        # Vérifier si le meilleur candidat est acceptable
        if best_similarity >= self.min_similarity:
            return (best_row, best_similarity, best_headers)

        return None

    def _calculate_similarity(self, headers: List[str]) -> float:
        """
        Calcule la similarité entre des en-têtes et la référence

        Utilise l'indice de Jaccard sur les mots normalisés

        Args:
            headers: Liste d'en-têtes à comparer

        Returns:
            Score de similarité (0.0 à 1.0)
        """
        if not headers:
            return 0.0

        # Extraire les mots des en-têtes
        current_words = self._extract_words(headers)

        if not current_words:
            return 0.0

        # Indice de Jaccard
        intersection = len(self.reference_words & current_words)
        union = len(self.reference_words | current_words)

        if union == 0:
            return 0.0

        return intersection / union

    def _extract_words(self, headers: List[str]) -> set:
        """
        Extrait les mots significatifs des en-têtes

        Args:
            headers: Liste d'en-têtes

        Returns:
            Set de mots normalisés
        """
        words = set()

        # Mots à ignorer (stop words)
        stop_words = {
            'de', 'du', 'des', 'le', 'la', 'les', 'un', 'une',
            'et', 'ou', 'à', 'au', 'aux', 'en', 'pour', 'par',
            'dans', 'sur', 'avec', 'sans'
        }

        for header in headers:
            # Nettoyer
            header_clean = header.lower().strip()

            # Séparer par espaces, tirets, underscores, parenthèses
            tokens = re.split(r'[\s\-_()\/]+', header_clean)

            for token in tokens:
                # Retirer ponctuation
                token_clean = re.sub(r'[^\w]', '', token)

                # Garder mots significatifs (3+ caractères, pas stop words)
                if len(token_clean) >= 3 and token_clean not in stop_words:
                    words.add(token_clean)

        return words

    def _create_result_from_match(self, file_path: str, df: pd.DataFrame,
                                   row_num: int, similarity: float,
                                   headers: List[str]) -> DetectionResult:
        """
        Crée un DetectionResult à partir d'un match

        Args:
            file_path: Chemin du fichier
            df: DataFrame
            row_num: Numéro de ligne détecté (base 1)
            similarity: Score de similarité
            headers: En-têtes détectés

        Returns:
            DetectionResult
        """
        # Calculer data_start et data_end
        data_start = row_num + self.reference_header_lines
        data_end = self._find_data_end(df, data_start - 1)  # base 0

        # Déterminer le niveau de confiance
        if similarity >= self.similarity_threshold:
            confidence = 0.95
            warning = None
        elif similarity >= self.min_similarity:
            confidence = 0.80
            warning = (
                f"⚠️ Similarité moyenne ({similarity:.0%}). "
                f"Vérifiez que la ligne {row_num} est bien l'en-tête."
            )
        else:
            confidence = 0.50
            warning = (
                f"❌ Similarité faible ({similarity:.0%}). "
                f"La détection est probablement incorrecte."
            )

        result = DetectionResult(
            file_path=file_path,
            header_start_row=row_num,
            header_rows=self.reference_header_lines,
            detected_headers=headers,
            data_start_row=data_start,
            data_end_row=data_end,
            confidence=confidence,
            detection_method="reference",
            total_rows=len(df),
            warning=warning,
            debug_info={
                'reference_file': self.reference_file,
                'reference_header_row': self.reference_header_row,
                'similarity_score': similarity,
                'reference_headers': self.reference_headers,
                'detected_headers': headers
            }
        )

        return result

    def _find_data_end(self, df: pd.DataFrame, data_start_idx: int) -> int:
        """
        Trouve la fin des données (bloc de lignes vides consécutives)

        On ne coupe le tableau que lorsqu'on rencontre au moins
        ``empty_block_threshold`` lignes presque vides d'affilée. Une ligne
        vide isolée (séparateur, sous-total) ne termine PAS les données : elle
        sera filtrée en aval si ``remove_empty_rows`` est activé. Cela évite de
        tronquer silencieusement toutes les lignes situées après un trou unique.

        Args:
            df: DataFrame
            data_start_idx: Index de début des données (base 0, pandas)

        Returns:
            Index de fin des données, base 0, EXCLUSIF (= index de la première
            ligne du premier bloc vide qualifiant, ou len(df) si aucun).
        """
        n = len(df)
        run_start = None  # début du bloc vide courant (base 0)
        run_len = 0

        for row_idx in range(data_start_idx, n):
            density = self.calculate_row_density(df, row_idx)
            if density < 0.1:  # Ligne presque vide
                if run_start is None:
                    run_start = row_idx
                run_len += 1
                if run_len >= self.empty_block_threshold:
                    # Bloc vide qualifiant: les données s'arrêtent à son début
                    return run_start
            else:
                run_start = None
                run_len = 0

        return n  # Aucun bloc vide qualifiant: les données vont jusqu'au bout

    def _create_failed_result(self, file_path: str, reason: str) -> DetectionResult:
        """
        Crée un résultat de détection échouée

        Args:
            file_path: Chemin du fichier
            reason: Raison de l'échec

        Returns:
            DetectionResult avec confiance 0.0
        """
        return DetectionResult(
            file_path=file_path,
            confidence=0.0,
            detection_method="reference",
            warning=reason
        )
