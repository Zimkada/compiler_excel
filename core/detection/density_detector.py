"""
Détecteur de structure basé sur l'analyse de densité des lignes
Auteur: GOUNOU N'GOBI Chabi Zimé
Version: 3.2

Précision attendue: 85% pour tous types de fichiers
"""

from typing import Optional, Dict, List, Tuple
from pathlib import Path
import pandas as pd
import numpy as np

from .base_detector import BaseDetector, DetectionResult


class DensityDetector(BaseDetector):
    """
    Détecte la structure d'un tableau en analysant la densité des lignes

    Principe:
    - Les lignes d'en-tête ont une densité élevée (>70%)
    - Les lignes de données ont une densité constante (>50%)
    - Les lignes vides ou d'information ont une densité faible (<30%)
    - Les changements brusques de densité indiquent des transitions

    Score de confiance basé sur:
    - Cohérence de la densité dans la zone de données
    - Écart de densité entre en-têtes et infos avant/après
    - Continuité des données
    """

    def __init__(self, confidence_threshold: float = 0.5,
                 header_density_min: float = 0.50,
                 data_density_min: float = 0.30):
        """
        Initialise le détecteur de densité

        Args:
            confidence_threshold: Seuil de confiance minimum
            header_density_min: Densité minimale pour un en-tête (50%)
            data_density_min: Densité minimale pour une ligne de données (30%)
        """
        super().__init__(confidence_threshold)
        self.header_density_min = header_density_min
        self.data_density_min = data_density_min

    def detect(self, file_path: str, df: Optional[pd.DataFrame] = None) -> DetectionResult:
        """
        Détecte la structure en analysant la densité des lignes

        Args:
            file_path: Chemin du fichier
            df: DataFrame optionnel (chargé si None)

        Returns:
            DetectionResult avec informations détectées
        """
        try:
            # Charger le fichier si nécessaire
            if df is None:
                df = self.load_file(file_path)

            if len(df) == 0:
                return self._create_failed_result(file_path, "Fichier vide")

            # Analyser la densité de chaque ligne
            density_profile = self._analyze_density_profile(df)

            # Détecter les zones (leading, headers, data, trailing)
            zones = self._detect_zones(df, density_profile)

            if zones is None:
                return self._create_failed_result(
                    file_path,
                    "Impossible de détecter une structure cohérente"
                )

            # Extraire les en-têtes
            headers = self.extract_headers(
                df,
                zones['header_start'],
                zones['header_rows']
            )

            # Calculer le score de confiance
            confidence = self._calculate_confidence(density_profile, zones)

            # Créer le résultat
            result = DetectionResult(
                file_path=file_path,
                header_start_row=zones['header_start'],
                header_rows=zones['header_rows'],
                detected_headers=headers,
                data_start_row=zones['data_start'],
                data_end_row=zones['data_end'],
                has_trailing_content=zones['has_trailing'],
                trailing_start_row=zones.get('trailing_start'),
                trailing_end_row=zones.get('trailing_end'),
                trailing_row_count=zones.get('trailing_count', 0),
                trailing_detection_reason="density_drop",
                confidence=confidence,
                detection_method="density",
                total_rows=len(df),
                debug_info={
                    'density_profile': density_profile,
                    'zones': zones,
                    'header_density_min': self.header_density_min,
                    'data_density_min': self.data_density_min
                }
            )

            # Le __post_init__ calculera has_leading_content

            return result

        except Exception as e:
            return self._create_failed_result(
                file_path,
                f"Erreur analyse densité: {e}"
            )

    def _analyze_density_profile(self, df: pd.DataFrame) -> List[Dict]:
        """
        Analyse la densité de chaque ligne

        Args:
            df: DataFrame

        Returns:
            Liste de dicts avec stats par ligne
        """
        profile = []

        for row_idx in range(len(df)):
            row = df.iloc[row_idx]

            # Calculer densité
            density = self.calculate_row_density(df, row_idx)

            # Compter types de données
            text_count = 0
            number_count = 0
            empty_count = 0

            for val in row:
                if pd.isna(val) or val == '':
                    empty_count += 1
                else:
                    try:
                        float(val)
                        number_count += 1
                    except (ValueError, TypeError):
                        text_count += 1

            profile.append({
                'row_idx': row_idx,  # Base 0 (pandas)
                'row_num': row_idx + 1,  # Base 1 (Excel)
                'density': density,
                'text_count': text_count,
                'number_count': number_count,
                'empty_count': empty_count,
                'is_likely_header': self.is_likely_header_row(df, row_idx)
            })

        return profile

    def _detect_zones(self, df: pd.DataFrame,
                     density_profile: List[Dict]) -> Optional[Dict]:
        """
        Détecte les différentes zones du fichier

        Args:
            df: DataFrame
            density_profile: Résultat de _analyze_density_profile

        Returns:
            Dict avec zones détectées ou None
        """
        # Chercher la première ligne candidate pour en-tête
        header_start_idx = None

        for i, row_info in enumerate(density_profile):
            # Un en-tête doit avoir:
            # - Densité suffisante
            # - Plus de texte que de chiffres
            if (row_info['density'] >= self.header_density_min and
                row_info['is_likely_header']):
                header_start_idx = i
                break

        if header_start_idx is None:
            return None

        # Déterminer le nombre de lignes d'en-tête.
        #
        # Départage par le CONTENU (même recette que BorderDetector, juin 2026,
        # cf. test_border_header_rows.py) : une ligne de données majoritairement
        # textuelle (ex. « P0 | 0 | V0 », 2 textes / 1 nombre) satisfait
        # is_likely_header et se faisait absorber dans l'en-tête — jusqu'à 21
        # lignes d'en-tête détectées, corrompant data_start. On n'étend donc
        # l'en-tête que si la ligne ne contient AUCUNE cellule numérique
        # (les vrais en-têtes sont purement textuels), borné à 5 lignes.
        #
        # Garde-fou tableau 100% textuel : si aucune ligne suivante ne contient
        # de nombre, le critère numérique n'a pas de butoir et engloutirait les
        # données -> défaut sûr : 1 seule ligne d'en-tête.
        header_rows = 1
        following_has_numbers = any(
            density_profile[i]['number_count'] > 0
            for i in range(header_start_idx + 1, len(density_profile))
        )
        if following_has_numbers:
            next_idx = header_start_idx + 1
            max_header_idx = min(header_start_idx + 4, len(density_profile) - 1)
            while next_idx <= max_header_idx:
                next_row = density_profile[next_idx]
                if (next_row['density'] >= self.header_density_min and
                        next_row['is_likely_header'] and
                        next_row['number_count'] == 0):
                    header_rows += 1
                    next_idx += 1
                else:
                    break

        # Début des données
        data_start_idx = header_start_idx + header_rows

        # Trouver la fin des données
        data_end_idx = self._detect_data_end(
            density_profile,
            data_start_idx
        )

        # Vérifier s'il y a du contenu trailing
        trailing_info = self._detect_trailing_from_density(
            density_profile,
            data_end_idx
        )

        zones = {
            'header_start': header_start_idx + 1,  # Base 1 (Excel)
            'header_rows': header_rows,
            'data_start': data_start_idx + 1,  # Base 1 (Excel)
            'data_end': data_end_idx + 1,  # Base 1 (Excel)
            'has_trailing': trailing_info['has_trailing'],
        }

        if trailing_info['has_trailing']:
            zones['trailing_start'] = trailing_info['start_row']
            zones['trailing_end'] = trailing_info['end_row']
            zones['trailing_count'] = trailing_info['row_count']

        return zones

    def _detect_data_end(self, density_profile: List[Dict],
                        data_start_idx: int) -> int:
        """
        Détecte la fin des données

        Args:
            density_profile: Profil de densité
            data_start_idx: Index de début des données (base 0)

        Returns:
            Index de fin des données (base 0)
        """
        # Calculer la densité moyenne des premières lignes de données
        sample_size = min(5, len(density_profile) - data_start_idx)
        if sample_size <= 0:
            return data_start_idx

        sample_densities = [
            density_profile[data_start_idx + i]['density']
            for i in range(sample_size)
        ]
        avg_data_density = np.mean(sample_densities)

        # Seuil pour détecter une chute de densité
        threshold = max(self.data_density_min, avg_data_density * 0.5)

        # Parcourir les lignes et détecter où la densité chute
        data_end_idx = data_start_idx
        consecutive_low = 0

        for i in range(data_start_idx, len(density_profile)):
            density = density_profile[i]['density']

            if density < threshold:
                consecutive_low += 1
                # Si 2 lignes consécutives sous le seuil, les données sont terminées
                if consecutive_low >= 2:
                    break
            else:
                consecutive_low = 0
                data_end_idx = i

        return data_end_idx

    def _detect_trailing_from_density(self, density_profile: List[Dict],
                                     data_end_idx: int) -> Dict:
        """
        Détecte le contenu trailing basé sur la densité

        Args:
            density_profile: Profil de densité
            data_end_idx: Index de fin des données (base 0)

        Returns:
            Dict avec infos trailing
        """
        # Chercher des lignes non vides après data_end
        non_empty_rows = []

        for i in range(data_end_idx + 1, len(density_profile)):
            if density_profile[i]['density'] > 0.1:  # Au moins 10% rempli
                non_empty_rows.append(i)

        has_trailing = len(non_empty_rows) > 0

        if has_trailing:
            return {
                'has_trailing': True,
                'start_row': non_empty_rows[0] + 1,  # Base 1
                'end_row': non_empty_rows[-1] + 1,  # Base 1
                'row_count': len(non_empty_rows)
            }
        else:
            return {'has_trailing': False}

    def _calculate_confidence(self, density_profile: List[Dict],
                            zones: Dict) -> float:
        """
        Calcule le score de confiance de la détection

        Args:
            density_profile: Profil de densité
            zones: Zones détectées

        Returns:
            Score entre 0.0 et 1.0
        """
        header_start_idx = zones['header_start'] - 1  # Base 0
        data_start_idx = zones['data_start'] - 1  # Base 0
        data_end_idx = zones['data_end'] - 1  # Base 0

        # 1. Score de densité des en-têtes (30%)
        header_densities = [
            density_profile[i]['density']
            for i in range(header_start_idx, data_start_idx)
            if i < len(density_profile)
        ]
        header_score = np.mean(header_densities) if header_densities else 0

        # 2. Score de cohérence des données (40%)
        data_densities = [
            density_profile[i]['density']
            for i in range(data_start_idx, data_end_idx + 1)
            if i < len(density_profile)
        ]

        if data_densities:
            # Cohérence = 1 - écart-type normalisé
            data_std = np.std(data_densities)
            data_mean = np.mean(data_densities)
            data_coherence = 1 - min(1.0, data_std / (data_mean + 0.01))
        else:
            data_coherence = 0

        # 3. Score de séparation (30%)
        # Vérifier que la densité avant les en-têtes est plus faible
        separation_score = 0

        if header_start_idx > 0:
            leading_densities = [
                density_profile[i]['density']
                for i in range(0, header_start_idx)
            ]
            leading_mean = np.mean(leading_densities) if leading_densities else 0

            # Bonus si les en-têtes sont bien séparés du contenu précédent
            if header_score > leading_mean + 0.2:
                separation_score = 0.8
            elif header_score > leading_mean:
                separation_score = 0.5
        else:
            # Pas de contenu avant, séparation parfaite
            separation_score = 1.0

        # Score final pondéré
        confidence = (header_score * 0.3 +
                     data_coherence * 0.4 +
                     separation_score * 0.3)

        return min(1.0, max(0.0, confidence))

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
            detection_method="density",
            warning=reason
        )
