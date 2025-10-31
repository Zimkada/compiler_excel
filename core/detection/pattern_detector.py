"""
Détecteur de structure basé sur l'analyse de patterns dans les données
Auteur: GOUNOU N'GOBI Chabi Zimé
Version: 3.2

Précision attendue: 75% (complément aux autres détecteurs)
"""

from typing import Optional, Dict, List, Set
from pathlib import Path
import pandas as pd
import numpy as np
import re

from .base_detector import BaseDetector, DetectionResult


class PatternDetector(BaseDetector):
    """
    Détecte la structure en analysant les patterns de données

    Principe:
    - Les en-têtes contiennent des mots-clés typiques (nom, prénom, date, etc.)
    - Les données ont des patterns réguliers (colonnes de même type)
    - Les lignes de signature/info ont des patterns textuels spécifiques
    - Détection de changements de structure

    Score de confiance basé sur:
    - Présence de mots-clés typiques dans les en-têtes
    - Cohérence des types de données par colonne
    - Régularité des patterns
    """

    # Mots-clés génériques pour détecter les en-têtes (pas secteur-spécifique)
    HEADER_KEYWORDS = {
        # Identification
        'nom', 'prenom', 'prénom', 'name', 'firstname', 'lastname',
        'id', 'identifiant', 'numéro', 'numero', 'code', 'reference', 'référence',

        # Temporel
        'date', 'jour', 'mois', 'année', 'annee', 'year', 'month', 'day',
        'période', 'periode', 'heure', 'time',

        # Quantitatif
        'montant', 'quantité', 'quantite', 'nombre', 'total', 'somme',
        'prix', 'price', 'amount', 'quantity', 'count',

        # Descriptif
        'description', 'libellé', 'libelle', 'titre', 'title', 'type',
        'catégorie', 'categorie', 'category', 'statut', 'status',

        # Localisation
        'lieu', 'place', 'location', 'adresse', 'address', 'ville', 'city',
        'région', 'region', 'pays', 'country'
    }

    # Patterns pour détecter le contenu trailing (signature, date, etc.)
    TRAILING_PATTERNS = [
        r'signature',
        r'sign[ée]?\s+[àa]',
        r'fait\s+[àa]',
        r'le\s+\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}',
        r'date\s*:',
        r'directeur',
        r'responsable',
        r'chef',
        r'président',
        r'président',
    ]

    def __init__(self, confidence_threshold: float = 0.5):
        """
        Initialise le détecteur de patterns

        Args:
            confidence_threshold: Seuil de confiance minimum
        """
        super().__init__(confidence_threshold)

    def detect(self, file_path: str, df: Optional[pd.DataFrame] = None) -> DetectionResult:
        """
        Détecte la structure en analysant les patterns

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

            # Analyser les patterns
            pattern_analysis = self._analyze_patterns(df)

            # Détecter les zones
            zones = self._detect_zones_from_patterns(df, pattern_analysis)

            if zones is None:
                return self._create_failed_result(
                    file_path,
                    "Impossible de détecter une structure par patterns"
                )

            # Extraire les en-têtes
            headers = self.extract_headers(
                df,
                zones['header_start'],
                zones['header_rows']
            )

            # Calculer le score de confiance
            confidence = self._calculate_confidence(pattern_analysis, zones)

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
                trailing_detection_reason="pattern_match",
                confidence=confidence,
                detection_method="pattern",
                total_rows=len(df),
                debug_info={
                    'pattern_analysis': pattern_analysis,
                    'header_keywords_found': zones.get('header_keywords', [])
                }
            )

            return result

        except Exception as e:
            return self._create_failed_result(
                file_path,
                f"Erreur analyse patterns: {e}"
            )

    def _analyze_patterns(self, df: pd.DataFrame) -> List[Dict]:
        """
        Analyse les patterns de chaque ligne

        Args:
            df: DataFrame

        Returns:
            Liste de dicts avec patterns par ligne
        """
        analysis = []

        for row_idx in range(len(df)):
            row = df.iloc[row_idx]

            # Analyser le contenu textuel
            text_content = ' '.join([
                str(v).lower() for v in row if pd.notna(v)
            ])

            # Détecter les mots-clés d'en-tête
            header_keywords = self._find_header_keywords(text_content)

            # Détecter les patterns de signature/trailing
            trailing_patterns = self._find_trailing_patterns(text_content)

            # Analyser les types de données
            type_pattern = self._analyze_column_types(row)

            # Score de ressemblance à un en-tête
            header_score = len(header_keywords) / max(1, len(row.dropna()))

            analysis.append({
                'row_idx': row_idx,  # Base 0
                'row_num': row_idx + 1,  # Base 1
                'header_keywords': list(header_keywords),
                'header_keyword_count': len(header_keywords),
                'trailing_patterns': trailing_patterns,
                'trailing_pattern_count': len(trailing_patterns),
                'type_pattern': type_pattern,
                'header_score': header_score,
                'text_content': text_content
            })

        return analysis

    def _find_header_keywords(self, text: str) -> Set[str]:
        """
        Trouve les mots-clés d'en-tête dans un texte

        Args:
            text: Texte à analyser (déjà en minuscules)

        Returns:
            Set de mots-clés trouvés
        """
        found = set()

        for keyword in self.HEADER_KEYWORDS:
            # Recherche de mot entier (avec limites de mots)
            pattern = r'\b' + re.escape(keyword) + r'\b'
            if re.search(pattern, text):
                found.add(keyword)

        return found

    def _find_trailing_patterns(self, text: str) -> List[str]:
        """
        Trouve les patterns de signature/trailing dans un texte

        Args:
            text: Texte à analyser (déjà en minuscules)

        Returns:
            Liste de patterns trouvés
        """
        found = []

        for pattern in self.TRAILING_PATTERNS:
            if re.search(pattern, text):
                found.append(pattern)

        return found

    def _analyze_column_types(self, row: pd.Series) -> List[str]:
        """
        Analyse les types de données dans une ligne

        Args:
            row: Ligne à analyser

        Returns:
            Liste de types par colonne ('text', 'number', 'date', 'empty')
        """
        types = []

        for val in row:
            if pd.isna(val) or val == '':
                types.append('empty')
            elif isinstance(val, (int, float)):
                types.append('number')
            elif self._is_date_like(str(val)):
                types.append('date')
            else:
                types.append('text')

        return types

    def _is_date_like(self, text: str) -> bool:
        """
        Vérifie si un texte ressemble à une date

        Args:
            text: Texte à vérifier

        Returns:
            True si ça ressemble à une date
        """
        date_patterns = [
            r'\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}',  # 12/01/2024
            r'\d{4}[/\-]\d{1,2}[/\-]\d{1,2}',    # 2024-01-12
            r'\d{1,2}\s+[a-zA-Z]+\s+\d{4}',      # 12 janvier 2024
        ]

        for pattern in date_patterns:
            if re.search(pattern, text):
                return True

        return False

    def _detect_zones_from_patterns(self, df: pd.DataFrame,
                                    pattern_analysis: List[Dict]) -> Optional[Dict]:
        """
        Détecte les zones basées sur les patterns

        Args:
            df: DataFrame
            pattern_analysis: Résultat de _analyze_patterns

        Returns:
            Dict avec zones ou None
        """
        # Chercher la ligne avec le plus de mots-clés d'en-tête
        best_header_idx = None
        best_header_score = 0

        for i, row_info in enumerate(pattern_analysis):
            if row_info['header_score'] > best_header_score:
                best_header_score = row_info['header_score']
                best_header_idx = i

        # Si aucun en-tête détecté avec confiance
        if best_header_idx is None or best_header_score < 0.1:
            return None

        # Vérifier si la ligne suivante est aussi un en-tête
        header_rows = 1
        next_idx = best_header_idx + 1

        if next_idx < len(pattern_analysis):
            next_score = pattern_analysis[next_idx]['header_score']
            # Si la ligne suivante a aussi un bon score, l'inclure
            if next_score > 0.1 and next_score >= best_header_score * 0.5:
                header_rows = 2

        # Début des données
        data_start_idx = best_header_idx + header_rows

        # Trouver la fin des données en cherchant des patterns de trailing
        data_end_idx = self._detect_data_end_from_patterns(
            pattern_analysis,
            data_start_idx
        )

        # Détecter le contenu trailing
        trailing_info = self._detect_trailing_from_patterns(
            pattern_analysis,
            data_end_idx
        )

        zones = {
            'header_start': best_header_idx + 1,  # Base 1
            'header_rows': header_rows,
            'data_start': data_start_idx + 1,  # Base 1
            'data_end': data_end_idx + 1,  # Base 1
            'has_trailing': trailing_info['has_trailing'],
            'header_keywords': pattern_analysis[best_header_idx]['header_keywords']
        }

        if trailing_info['has_trailing']:
            zones['trailing_start'] = trailing_info['start_row']
            zones['trailing_end'] = trailing_info['end_row']
            zones['trailing_count'] = trailing_info['row_count']

        return zones

    def _detect_data_end_from_patterns(self, pattern_analysis: List[Dict],
                                      data_start_idx: int) -> int:
        """
        Détecte la fin des données en cherchant des patterns de trailing

        Args:
            pattern_analysis: Analyse des patterns
            data_start_idx: Index de début des données (base 0)

        Returns:
            Index de fin des données (base 0)
        """
        data_end_idx = data_start_idx

        # Parcourir les lignes après data_start
        for i in range(data_start_idx, len(pattern_analysis)):
            row_info = pattern_analysis[i]

            # Si on trouve des patterns de trailing, les données sont terminées
            if row_info['trailing_pattern_count'] > 0:
                break

            # Si on trouve à nouveau des mots-clés d'en-tête (improbable mais possible)
            if row_info['header_score'] > 0.3:
                break

            # Sinon, c'est encore des données
            data_end_idx = i

        return data_end_idx

    def _detect_trailing_from_patterns(self, pattern_analysis: List[Dict],
                                      data_end_idx: int) -> Dict:
        """
        Détecte le contenu trailing basé sur les patterns

        Args:
            pattern_analysis: Analyse des patterns
            data_end_idx: Index de fin des données (base 0)

        Returns:
            Dict avec infos trailing
        """
        trailing_rows = []

        # Chercher des lignes avec patterns de trailing après data_end
        for i in range(data_end_idx + 1, len(pattern_analysis)):
            row_info = pattern_analysis[i]

            # Si la ligne contient des patterns de trailing ou du texte
            if (row_info['trailing_pattern_count'] > 0 or
                len(row_info['text_content'].strip()) > 0):
                trailing_rows.append(i)

        has_trailing = len(trailing_rows) > 0

        if has_trailing:
            return {
                'has_trailing': True,
                'start_row': trailing_rows[0] + 1,  # Base 1
                'end_row': trailing_rows[-1] + 1,  # Base 1
                'row_count': len(trailing_rows)
            }
        else:
            return {'has_trailing': False}

    def _calculate_confidence(self, pattern_analysis: List[Dict],
                            zones: Dict) -> float:
        """
        Calcule le score de confiance

        Args:
            pattern_analysis: Analyse des patterns
            zones: Zones détectées

        Returns:
            Score entre 0.0 et 1.0
        """
        header_start_idx = zones['header_start'] - 1  # Base 0
        header_rows = zones['header_rows']

        # 1. Score basé sur les mots-clés d'en-tête (50%)
        header_keyword_counts = [
            pattern_analysis[header_start_idx + i]['header_keyword_count']
            for i in range(header_rows)
            if header_start_idx + i < len(pattern_analysis)
        ]

        avg_keywords = np.mean(header_keyword_counts) if header_keyword_counts else 0
        # Normaliser (3+ mots-clés = score maximal)
        keyword_score = min(1.0, avg_keywords / 3.0)

        # 2. Score basé sur la cohérence des types de colonnes (30%)
        data_start_idx = zones['data_start'] - 1
        data_end_idx = zones['data_end'] - 1

        type_consistency_score = self._calculate_type_consistency(
            pattern_analysis,
            data_start_idx,
            data_end_idx
        )

        # 3. Score de structure (20%)
        # Bonus si on a détecté du contenu trailing avec patterns spécifiques
        structure_score = 0.5  # Score de base

        if zones['has_trailing']:
            trailing_start_idx = zones['trailing_start'] - 1
            # Vérifier si le trailing contient des patterns
            trailing_has_patterns = any(
                pattern_analysis[i]['trailing_pattern_count'] > 0
                for i in range(trailing_start_idx, len(pattern_analysis))
                if i < len(pattern_analysis)
            )
            if trailing_has_patterns:
                structure_score = 1.0

        # Score final pondéré
        confidence = (keyword_score * 0.5 +
                     type_consistency_score * 0.3 +
                     structure_score * 0.2)

        return min(1.0, max(0.0, confidence))

    def _calculate_type_consistency(self, pattern_analysis: List[Dict],
                                   data_start_idx: int,
                                   data_end_idx: int) -> float:
        """
        Calcule la cohérence des types de données par colonne

        Args:
            pattern_analysis: Analyse des patterns
            data_start_idx: Index de début (base 0)
            data_end_idx: Index de fin (base 0)

        Returns:
            Score de cohérence entre 0.0 et 1.0
        """
        if data_end_idx < data_start_idx:
            return 0.0

        # Extraire les patterns de types pour les lignes de données
        type_patterns = [
            pattern_analysis[i]['type_pattern']
            for i in range(data_start_idx, data_end_idx + 1)
            if i < len(pattern_analysis)
        ]

        if not type_patterns:
            return 0.0

        # Calculer la cohérence par colonne
        num_cols = len(type_patterns[0]) if type_patterns else 0
        if num_cols == 0:
            return 0.0

        consistency_scores = []

        for col_idx in range(num_cols):
            # Types dans cette colonne
            col_types = [
                row_types[col_idx]
                for row_types in type_patterns
                if col_idx < len(row_types) and row_types[col_idx] != 'empty'
            ]

            if not col_types:
                continue

            # Type le plus fréquent
            from collections import Counter
            type_counts = Counter(col_types)
            most_common_type, count = type_counts.most_common(1)[0]

            # Score = proportion du type le plus fréquent
            consistency = count / len(col_types)
            consistency_scores.append(consistency)

        # Score moyen de cohérence
        return np.mean(consistency_scores) if consistency_scores else 0.0

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
            detection_method="pattern",
            warning=reason
        )
