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

from .base_detector import BaseDetector, DetectionResult, prune_phantom_columns
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

            # Élaguer les colonnes parasites (cellules vides formatées en fin de
            # tableau). Sans cela, la densité est diluée et la détection échoue.
            df, pruned_cols = prune_phantom_columns(df)

            # Exécuter tous les détecteurs
            results = self._run_all_detectors(file_path, df)

            # Sélectionner le meilleur résultat
            best_result = self._select_best_result(results)

            if best_result is None:
                return self._create_failed_result(
                    file_path,
                    "Aucun détecteur n'a réussi à détecter la structure"
                )

            # Signaler l'élagage (sans dégrader la confiance) : l'utilisateur
            # doit en être informé dans l'aperçu pour décider en connaissance.
            if pruned_cols > 0:
                best_result.debug_info['pruned_phantom_columns'] = pruned_cols
                kept = df.shape[1]
                notice = (
                    f"{pruned_cols} colonne(s) vide(s) parasite(s) ignorée(s) "
                    f"(le tableau utile fait {kept} colonne(s)). Vérifiez que vos "
                    f"données tiennent bien dans ces {kept} premières colonnes."
                )
                best_result.warning = (
                    f"{best_result.warning} · {notice}"
                    if best_result.warning else notice
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
        if ext in ['.xlsx', '.xlsm']:
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
        # Voter pour la ligne d'en-tête
        header_starts = [r.header_start_row for r in results]

        # Accord inter-détecteurs : quand >= 2 détecteurs indépendants pointent
        # la MÊME ligne d'en-tête, cet accord est une preuve en soi. La moyenne
        # pondérée, elle, DILUE : deux détecteurs corrects à 0.54 et 0.73
        # donnaient 0.64 combiné — sous le seuil de 0.65, et la détection
        # correcte était rejetée au profit du fallback manuel. En cas d'accord,
        # la confiance = meilleure confiance individuelle + bonus, plafonnée.
        agreed_start = None
        for start in set(header_starts):
            if header_starts.count(start) >= 2:
                agreed_start = start
                break

        if agreed_start is not None:
            agreeing = [r for r in results if r.header_start_row == agreed_start]
            voted_header_start = agreed_start
            combined_confidence = min(
                0.95, max(r.confidence for r in agreeing) + 0.15
            )
            # Les votes suivants (header_rows, data_end) ne considèrent que les
            # détecteurs d'accord sur l'en-tête : un détecteur qui s'est trompé
            # de ligne n'a pas voix au chapitre sur le reste de la structure.
            voters = agreeing
        else:
            # Désaccord : moyenne pondérée historique + médiane.
            total_weight = 0
            weighted_score = 0
            for result in results:
                weight = weights.get(result.detection_method, 0)
                weighted_score += result.confidence * weight
                total_weight += weight
            combined_confidence = (
                weighted_score / total_weight if total_weight > 0 else 0
            )
            voted_header_start = int(np.median(header_starts))
            # Même règle qu'en cas d'accord : seuls les détecteurs qui ont
            # trouvé la ligne d'en-tête RETENUE votent sur le reste de la
            # structure. Un détecteur qui s'est trompé de ligne décrit un autre
            # tableau que celui qu'on garde ; son header_rows ne veut rien dire
            # ici. Sans ce filtre, sur un fichier où density disait (start=3,
            # rows=2) et pattern (start=4, rows=1), la médiane de [2, 1] donnait
            # rows=1 : l'en-tête sur deux lignes n'était plus aplati, ses
            # libellés ne correspondaient plus à ceux des autres fichiers du lot
            # et l'aligneur créait des colonnes en double.
            voters = [r for r in results if r.header_start_row == voted_header_start]
            if not voters:
                # La médiane peut tomber entre deux lignes proposées : personne
                # ne l'a votée. On retombe alors sur l'ensemble des détecteurs.
                voters = results

        # Nombre de lignes d'en-tête : écarter les valeurs aberrantes (> 5,
        # symptôme d'un détecteur ayant absorbé des données dans l'en-tête,
        # cf. bug DensityDetector header_rows=21) avant de prendre la médiane.
        # Si tout est aberrant, défaut sûr : 1 ligne.
        sane_header_rows = [r.header_rows for r in voters if 1 <= r.header_rows <= 5]
        voted_header_rows = int(np.median(sane_header_rows)) if sane_header_rows else 1

        # Fin des données : ne JAMAIS tronquer silencieusement. Si un détecteur
        # dit « jusqu'au bout » (0) on garde 0 ; sinon on prend le MAX des fins
        # candidates — garder trop de lignes est bénin (les lignes vides et
        # totaux sont filtrés en aval), en perdre est destructeur. L'ancienne
        # médiane pouvait couper la moitié des données sur simple désaccord.
        # NB : on considère ici TOUS les détecteurs, pas seulement `voters` :
        # restreindre le vote abaisserait ce maximum et tronquerait des lignes.
        data_ends = [r.data_end_row for r in results]
        voted_data_end = 0 if any(e == 0 for e in data_ends) else int(max(data_ends))

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
                'selection_reason': (
                    'detector_agreement' if agreed_start is not None
                    else 'weighted_vote'
                )
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
        Ajuste la confiance à partir du score de validation croisée.

        Principe : la validation croisée ne sert qu'à RENFORCER la confiance
        quand plusieurs fichiers se ressemblent fortement. Elle ne doit jamais
        dégrader une détection ni produire d'avertissement.

        En effet, une faible similarité des en-têtes entre fichiers n'indique
        PAS une erreur de détection : elle signifie simplement que les fichiers
        portent sur des sujets différents (ex. listes par établissement, dont
        seuls quelques libellés de colonnes coïncident). Pénaliser ou avertir
        dans ce cas produisait des faux positifs systématiques, y compris sur
        des détections parfaitement correctes.

        Le score est consigné en debug pour diagnostic, sans warning visible.

        Args:
            result: DetectionResult original
            cross_val_score: Score de validation croisée (0.0 à 1.0)

        Returns:
            DetectionResult avec confiance éventuellement renforcée
        """
        result.debug_info['cross_validation_score'] = cross_val_score

        # Bonus de confiance uniquement quand les fichiers se ressemblent fortement
        if cross_val_score > 0.7:
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
