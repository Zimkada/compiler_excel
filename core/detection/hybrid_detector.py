"""
Détecteur hybride avec cascade et validation croisée
Auteur: GOUNOU N'GOBI Chabi Zimé
Version: 3.2

Orchestrateur intelligent qui combine plusieurs détecteurs
"""

from typing import Optional, List, Dict, Tuple
from pathlib import Path
import pandas as pd
import numpy as np

from .base_detector import BaseDetector, DetectionResult
from .border_detector import BorderDetector
from .density_detector import DensityDetector
from .pattern_detector import PatternDetector


class HybridDetector(BaseDetector):
    """
    Détecteur hybride qui combine plusieurs méthodes de détection

    Stratégie en cascade:
    1. BorderDetector (si bordures présentes) - Poids 40%
    2. DensityDetector (toujours) - Poids 30%
    3. PatternDetector (toujours) - Poids 30%

    Score final = moyenne pondérée des détecteurs réussis
    + Bonus de validation croisée (similarité des en-têtes)

    La validation croisée compare les en-têtes détectés avec ceux
    des autres fichiers pour détecter les incohérences.
    """

    def __init__(self, confidence_threshold: float = 0.65,
                 enable_cross_validation: bool = True):
        """
        Initialise le détecteur hybride

        Args:
            confidence_threshold: Seuil de confiance minimum (0.65 par défaut)
            enable_cross_validation: Activer la validation croisée
        """
        super().__init__(confidence_threshold)
        self.enable_cross_validation = enable_cross_validation

        # Initialiser les détecteurs
        self.border_detector = BorderDetector(confidence_threshold=0.5)
        self.density_detector = DensityDetector(confidence_threshold=0.5)
        self.pattern_detector = PatternDetector(confidence_threshold=0.5)

        # Cache pour la validation croisée
        self._detection_cache: List[DetectionResult] = []

    def detect(self, file_path: str, df: Optional[pd.DataFrame] = None) -> DetectionResult:
        """
        Détecte la structure en combinant plusieurs détecteurs

        Args:
            file_path: Chemin du fichier
            df: DataFrame optionnel

        Returns:
            DetectionResult avec meilleur score de confiance
        """
        try:
            # Charger le fichier une seule fois
            if df is None:
                df = self.load_file(file_path)

            # Exécuter tous les détecteurs
            results = self._run_all_detectors(file_path, df)

            # Sélectionner le meilleur résultat
            best_result = self._select_best_result(results)

            if best_result is None:
                return self._create_failed_result(
                    file_path,
                    "Aucun détecteur n'a réussi à détecter la structure"
                )

            # Ajouter au cache pour validation croisée future
            self._detection_cache.append(best_result)

            # Effectuer la validation croisée si activée et si on a d'autres fichiers
            if self.enable_cross_validation and len(self._detection_cache) > 1:
                cross_validation_score = self._perform_cross_validation(best_result)
                best_result.cross_validation_score = cross_validation_score

                # Ajuster la confiance basée sur la validation croisée
                best_result = self._adjust_confidence_with_cross_validation(
                    best_result,
                    cross_validation_score
                )

            return best_result

        except Exception as e:
            return self._create_failed_result(
                file_path,
                f"Erreur détection hybride: {e}"
            )

    def detect_batch(self, file_paths: List[str]) -> List[DetectionResult]:
        """
        Détecte la structure de plusieurs fichiers avec validation croisée

        Cette méthode est optimale car elle permet de comparer tous les fichiers
        entre eux pour la validation croisée.

        Args:
            file_paths: Liste de chemins de fichiers

        Returns:
            Liste de DetectionResult
        """
        results = []

        # Première passe: détecter tous les fichiers
        for file_path in file_paths:
            result = self.detect(file_path)
            results.append(result)

        # Deuxième passe: validation croisée globale
        if self.enable_cross_validation and len(results) > 1:
            results = self._perform_global_cross_validation(results)

        return results

    def _run_all_detectors(self, file_path: str,
                          df: pd.DataFrame) -> List[DetectionResult]:
        """
        Exécute tous les détecteurs disponibles

        Args:
            file_path: Chemin du fichier
            df: DataFrame

        Returns:
            Liste de DetectionResult
        """
        results = []

        # 1. BorderDetector (seulement pour fichiers Excel)
        ext = Path(file_path).suffix.lower()
        if ext in ['.xlsx', '.xls', '.xlsm']:
            try:
                border_result = self.border_detector.detect(file_path, df)
                if border_result.confidence > 0:
                    results.append(border_result)
            except Exception:
                pass  # BorderDetector a échoué, continuer avec les autres

        # 2. DensityDetector (toujours)
        try:
            density_result = self.density_detector.detect(file_path, df)
            if density_result.confidence > 0:
                results.append(density_result)
        except Exception:
            pass

        # 3. PatternDetector (toujours)
        try:
            pattern_result = self.pattern_detector.detect(file_path, df)
            if pattern_result.confidence > 0:
                results.append(pattern_result)
        except Exception:
            pass

        return results

    def _select_best_result(self, results: List[DetectionResult]) -> Optional[DetectionResult]:
        """
        Sélectionne le meilleur résultat parmi les détecteurs

        Stratégie:
        - Si BorderDetector a un bon score (>0.8), le privilégier
        - Sinon, faire une moyenne pondérée des détecteurs

        Args:
            results: Liste de DetectionResult

        Returns:
            Meilleur DetectionResult ou None
        """
        if not results:
            return None

        # Pondérations par méthode
        weights = {
            'border': 0.40,
            'density': 0.30,
            'pattern': 0.30
        }

        # Chercher BorderDetector avec score élevé
        for result in results:
            if result.detection_method == 'border' and result.confidence > 0.80:
                # BorderDetector très confiant, le retourner directement
                result.debug_info['selection_reason'] = 'border_high_confidence'
                return result

        # Sinon, créer un résultat combiné
        return self._create_combined_result(results, weights)

    def _create_combined_result(self, results: List[DetectionResult],
                               weights: Dict[str, float]) -> DetectionResult:
        """
        Crée un résultat combiné à partir de plusieurs détecteurs

        Args:
            results: Liste de DetectionResult
            weights: Pondérations par méthode

        Returns:
            DetectionResult combiné
        """
        # Calculer le score pondéré
        total_weight = 0
        weighted_score = 0

        for result in results:
            method = result.detection_method
            weight = weights.get(method, 0)
            weighted_score += result.confidence * weight
            total_weight += weight

        combined_confidence = weighted_score / total_weight if total_weight > 0 else 0

        # Voter pour les valeurs détectées
        header_starts = [r.header_start_row for r in results]
        header_rows_counts = [r.header_rows for r in results]
        data_ends = [r.data_end_row for r in results]

        # Utiliser la médiane pour être robuste aux outliers
        voted_header_start = int(np.median(header_starts))
        voted_header_rows = int(np.median(header_rows_counts))
        voted_data_end = int(np.median(data_ends))

        # Choisir le résultat avec le header_start le plus proche du vote
        closest_result = min(
            results,
            key=lambda r: abs(r.header_start_row - voted_header_start)
        )

        # Créer le résultat combiné basé sur le résultat le plus proche
        combined = DetectionResult(
            file_path=closest_result.file_path,
            header_start_row=voted_header_start,
            header_rows=voted_header_rows,
            detected_headers=closest_result.detected_headers,
            data_start_row=voted_header_start + voted_header_rows,
            data_end_row=voted_data_end,
            has_trailing_content=closest_result.has_trailing_content,
            trailing_start_row=closest_result.trailing_start_row,
            trailing_end_row=closest_result.trailing_end_row,
            trailing_row_count=closest_result.trailing_row_count,
            trailing_detection_reason="hybrid",
            confidence=combined_confidence,
            detection_method="hybrid",
            total_rows=closest_result.total_rows,
            debug_info={
                'individual_results': [
                    {
                        'method': r.detection_method,
                        'confidence': r.confidence,
                        'header_start': r.header_start_row,
                        'data_end': r.data_end_row
                    }
                    for r in results
                ],
                'voted_values': {
                    'header_start': voted_header_start,
                    'header_rows': voted_header_rows,
                    'data_end': voted_data_end
                },
                'selection_reason': 'weighted_vote'
            }
        )

        return combined

    def _perform_cross_validation(self, result: DetectionResult) -> float:
        """
        Effectue la validation croisée avec les fichiers déjà détectés

        Compare les en-têtes du fichier actuel avec ceux des fichiers
        précédents pour détecter les incohérences.

        Args:
            result: DetectionResult à valider

        Returns:
            Score de validation croisée (0.0 à 1.0)
        """
        # Exclure le résultat actuel du cache
        other_results = [
            r for r in self._detection_cache
            if r.file_path != result.file_path
        ]

        if not other_results:
            return 1.0  # Aucune comparaison possible, retourner neutre

        # Calculer la similarité avec chaque autre fichier
        similarities = []

        for other in other_results:
            similarity = self._calculate_header_similarity(
                result.detected_headers,
                other.detected_headers
            )
            similarities.append(similarity)

        # Score = similarité moyenne
        avg_similarity = np.mean(similarities)

        return avg_similarity

    def _calculate_header_similarity(self, headers1: List[str],
                                    headers2: List[str]) -> float:
        """
        Calcule la similarité entre deux listes d'en-têtes

        Utilise l'indice de Jaccard sur les mots (pas les colonnes exactes)
        car les tableaux peuvent avoir des colonnes différentes mais
        des noms similaires.

        Args:
            headers1: Liste d'en-têtes 1
            headers2: Liste d'en-têtes 2

        Returns:
            Score de similarité (0.0 à 1.0)
        """
        if not headers1 or not headers2:
            return 0.0

        # Extraire tous les mots des en-têtes (normalisés)
        words1 = self._extract_header_words(headers1)
        words2 = self._extract_header_words(headers2)

        if not words1 or not words2:
            return 0.0

        # Indice de Jaccard
        intersection = len(words1 & words2)
        union = len(words1 | words2)

        similarity = intersection / union if union > 0 else 0.0

        return similarity

    def _extract_header_words(self, headers: List[str]) -> set:
        """
        Extrait les mots significatifs des en-têtes

        Args:
            headers: Liste d'en-têtes

        Returns:
            Set de mots normalisés
        """
        words = set()

        # Mots à ignorer (stop words basiques)
        stop_words = {'de', 'du', 'des', 'le', 'la', 'les', 'un', 'une',
                     'et', 'ou', 'à', 'au', 'aux', 'en', 'pour', 'par'}

        for header in headers:
            # Nettoyer et séparer
            header_clean = header.lower().strip()

            # Séparer par espaces, tirets, underscores
            import re
            tokens = re.split(r'[\s\-_]+', header_clean)

            for token in tokens:
                # Garder seulement les mots significatifs (3+ caractères)
                if len(token) >= 3 and token not in stop_words:
                    words.add(token)

        return words

    def _adjust_confidence_with_cross_validation(self, result: DetectionResult,
                                                 cross_val_score: float) -> DetectionResult:
        """
        Ajuste la confiance basée sur le score de validation croisée

        Args:
            result: DetectionResult original
            cross_val_score: Score de validation croisée

        Returns:
            DetectionResult avec confiance ajustée
        """
        # Si la similarité est très faible (<0.3), diminuer la confiance
        if cross_val_score < 0.3:
            result.confidence = result.confidence * 0.8
            result.warning = (
                f"Détection suspecte: en-têtes très différents des autres fichiers "
                f"(similarité: {cross_val_score:.0%})"
            )
        # Si la similarité est faible (<0.5), avertir
        elif cross_val_score < 0.5:
            result.warning = (
                f"En-têtes partiellement différents des autres fichiers "
                f"(similarité: {cross_val_score:.0%})"
            )
        # Si la similarité est élevée (>0.7), bonus de confiance
        elif cross_val_score > 0.7:
            result.confidence = min(1.0, result.confidence * 1.1)

        return result

    def _perform_global_cross_validation(self, results: List[DetectionResult]) -> List[DetectionResult]:
        """
        Effectue une validation croisée globale sur tous les résultats

        Compare tous les fichiers entre eux pour détecter les outliers.

        Args:
            results: Liste de DetectionResult

        Returns:
            Liste de DetectionResult avec scores ajustés
        """
        if len(results) < 2:
            return results

        # Calculer la similarité de chaque fichier avec tous les autres
        for i, result in enumerate(results):
            other_results = [r for j, r in enumerate(results) if j != i]

            similarities = [
                self._calculate_header_similarity(
                    result.detected_headers,
                    other.detected_headers
                )
                for other in other_results
            ]

            avg_similarity = np.mean(similarities)
            result.cross_validation_score = avg_similarity

            # Ajuster la confiance
            result = self._adjust_confidence_with_cross_validation(
                result,
                avg_similarity
            )

        return results

    def clear_cache(self):
        """Vide le cache de validation croisée"""
        self._detection_cache.clear()

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
            detection_method="hybrid",
            warning=reason
        )
