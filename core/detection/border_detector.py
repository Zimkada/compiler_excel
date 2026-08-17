"""
Détecteur de structure basé sur l'analyse des bordures Excel
Auteur: GOUNOU N'GOBI Chabi Zimé
Version: 3.2

Précision attendue: 95% pour fichiers avec bordures
"""

from typing import Optional, Dict, List, Tuple
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border

from .base_detector import BaseDetector, DetectionResult


class BorderDetector(BaseDetector):
    """
    Détecte la structure d'un tableau en analysant les bordures des cellules

    Principe:
    - Les en-têtes ont généralement des bordures complètes
    - Les données ont des bordures sur les côtés
    - Les lignes vides n'ont pas de bordures
    - Les infos avant/après ont peu ou pas de bordures

    Score de confiance basé sur:
    - Densité des bordures dans la zone détectée
    - Continuité des bordures
    - Proportion de cellules avec bordures complètes
    """

    def __init__(self, confidence_threshold: float = 0.5,
                 border_density_threshold: float = 0.40):
        """
        Initialise le détecteur de bordures

        Args:
            confidence_threshold: Seuil de confiance minimum (0.0 à 1.0)
            border_density_threshold: Densité minimale de bordures pour détecter
                                     une zone de tableau (40% par défaut)
        """
        super().__init__(confidence_threshold)
        self.border_density_threshold = border_density_threshold

    def detect(self, file_path: str, df: Optional[pd.DataFrame] = None) -> DetectionResult:
        """
        Détecte la structure en analysant les bordures

        Args:
            file_path: Chemin du fichier Excel
            df: DataFrame optionnel (ignoré, on charge avec openpyxl)

        Returns:
            DetectionResult avec informations détectées
        """
        ext = Path(file_path).suffix.lower()

        # Vérifier que c'est un fichier Excel (pas CSV)
        if ext not in ['.xlsx', '.xlsm']:
            return self._create_failed_result(
                file_path,
                "BorderDetector nécessite un fichier Excel (.xlsx, .xls, .xlsm)"
            )

        try:
            # Charger le fichier avec openpyxl pour accéder aux bordures
            wb = load_workbook(file_path, data_only=True)
            ws = wb.active

            # Analyser les bordures de chaque ligne
            border_analysis = self._analyze_borders(ws)

            # Charger le DataFrame (nécessaire pour distinguer en-tête et données)
            if df is None:
                df = self.load_file(file_path)

            # Détecter la zone de tableau (le contenu départage en-tête/données
            # quand les bordures ne suffisent pas — cas du tableau quadrillé)
            table_zone = self._detect_table_zone(border_analysis, df, ws)

            if table_zone is None:
                return self._create_failed_result(
                    file_path,
                    "Aucune zone de tableau détectée avec bordures"
                )

            # Extraire les informations
            header_start, header_end, data_end = table_zone

            # Calculer le nombre de lignes d'en-tête
            header_rows = header_end - header_start + 1

            # Extraire les en-têtes
            headers = self.extract_headers(df, header_start, header_rows)

            # Calculer le score de confiance
            confidence = self._calculate_confidence(border_analysis, table_zone)

            # Détecter le contenu trailing
            trailing_info = self._detect_trailing_content(
                df, border_analysis, data_end
            )

            # Créer le résultat
            result = DetectionResult(
                file_path=file_path,
                header_start_row=header_start,
                header_rows=header_rows,
                detected_headers=headers,
                data_start_row=header_end + 1,
                data_end_row=data_end,
                has_trailing_content=trailing_info['has_trailing'],
                trailing_start_row=trailing_info.get('start_row'),
                trailing_end_row=trailing_info.get('end_row'),
                trailing_row_count=trailing_info.get('row_count', 0),
                trailing_detection_reason="border_analysis",
                confidence=confidence,
                detection_method="border",
                total_rows=len(df),
                debug_info={
                    'border_analysis': border_analysis,
                    'border_density_threshold': self.border_density_threshold
                }
            )

            # Le __post_init__ de DetectionResult calculera has_leading_content

            return result

        except Exception as e:
            return self._create_failed_result(
                file_path,
                f"Erreur analyse bordures: {e}"
            )

    def _analyze_borders(self, worksheet) -> Dict[int, Dict]:
        """
        Analyse les bordures de chaque ligne

        Args:
            worksheet: Feuille openpyxl

        Returns:
            Dictionnaire {row_num: {border_stats}}
            row_num est en base 1 (Excel)
        """
        analysis = {}
        max_row = worksheet.max_row
        max_col = worksheet.max_column

        for row_num in range(1, max_row + 1):
            # Statistiques pour cette ligne
            total_cells = 0
            cells_with_borders = 0
            cells_with_full_borders = 0
            border_types = {'top': 0, 'bottom': 0, 'left': 0, 'right': 0}

            for col_num in range(1, max_col + 1):
                cell = worksheet.cell(row_num, col_num)
                total_cells += 1

                # Vérifier si la cellule a une bordure
                border = cell.border
                has_any_border = False
                has_full_border = True

                # Analyser chaque côté
                for side in ['top', 'bottom', 'left', 'right']:
                    side_border = getattr(border, side)
                    if side_border and side_border.style:
                        border_types[side] += 1
                        has_any_border = True
                    else:
                        has_full_border = False

                if has_any_border:
                    cells_with_borders += 1

                if has_full_border:
                    cells_with_full_borders += 1

            # Calculer les densités
            border_density = cells_with_borders / total_cells if total_cells > 0 else 0
            full_border_density = cells_with_full_borders / total_cells if total_cells > 0 else 0

            analysis[row_num] = {
                'total_cells': total_cells,
                'cells_with_borders': cells_with_borders,
                'cells_with_full_borders': cells_with_full_borders,
                'border_density': border_density,
                'full_border_density': full_border_density,
                'border_types': border_types
            }

        return analysis

    def _header_end_from_vertical_merges(self, worksheet,
                                         header_start: int,
                                         max_end: int) -> Optional[int]:
        """Fin de l'en-tête déduite des cellules FUSIONNÉES VERTICALEMENT.

        Un en-tête sur deux niveaux se construit presque toujours ainsi : la
        colonne d'identité est fusionnée verticalement (``A3:A4``) pendant que
        les catégories sont fusionnées horizontalement au-dessus de leurs
        sous-colonnes (``B3:C3``). Cette fusion verticale est une information
        de STRUCTURE, indépendante du contenu : elle dit que les lignes 3 et 4
        forment un seul bloc d'en-tête.

        C'est le signal qui manque quand les données ne contiennent aucun
        nombre (« NEANT », texte libre, ligne vide) : le critère par contenu
        est alors aveugle et retombe à 1 seule ligne d'en-tête, coupant le
        tableau au mauvais endroit.

        Une fusion PROFONDE est rejetée, jamais tronquée : ``A3:A12`` décrit un
        libellé recopié sur des lignes de données (regroupement visuel), pas un
        en-tête de dix niveaux. La ramener à ``max_end`` engloutirait des lignes
        de données dans l'en-tête. Au-delà de ``max_end``, on préfère donc ne
        rien conclure et laisser le défaut sûr (1 ligne) s'appliquer.

        Renvoie la dernière ligne d'en-tête, ou None si aucune fusion verticale
        plausible ne démarre sur ``header_start``.
        """
        if worksheet is None:
            return None
        header_end = None
        for rng in getattr(worksheet, 'merged_cells', {}).ranges:
            # Fusion verticale démarrant sur la 1re ligne d'en-tête.
            if rng.min_row == header_start and rng.max_row > rng.min_row:
                if rng.max_row > max_end:
                    continue  # fusion trop profonde : non concluante
                if header_end is None or rng.max_row > header_end:
                    header_end = rng.max_row
        return header_end

    def _detect_table_zone(self, border_analysis: Dict[int, Dict],
                           df, worksheet=None) -> Optional[Tuple[int, int, int]]:
        """
        Détecte la zone du tableau (en-tête + données).

        Les bordures localisent le tableau (première ligne bordée = début).
        Mais dans un tableau entièrement quadrillé, en-tête et données ont les
        mêmes bordures : on ne peut pas les distinguer par les bordures seules.
        On départage donc par le CONTENU : les lignes d'en-tête sont textuelles,
        les lignes de données contiennent des valeurs numériques. Cela évite de
        gonfler le nombre de lignes d'en-tête avec des lignes de données.

        Args:
            border_analysis: Résultat de _analyze_borders (clés = lignes base 1)
            df: DataFrame du fichier (base 0)

        Returns:
            Tuple (header_start, header_end, data_end) ou None
            Numéros de lignes en base 1 (Excel)
        """
        rows = sorted(border_analysis.keys())

        # 1) Première ligne bordée = début du tableau (donc des en-têtes)
        header_start = None
        for row_num in rows:
            if border_analysis[row_num]['border_density'] >= self.border_density_threshold:
                header_start = row_num
                break

        if header_start is None:
            return None

        # 2) Fin du tableau = dernière ligne consécutivement bordée
        data_end = header_start
        for row_num in rows:
            if row_num < header_start:
                continue
            if border_analysis[row_num]['border_density'] < self.border_density_threshold * 0.3:
                break
            data_end = row_num

        # 3) Fin des en-têtes par le contenu. Dans ces tableaux, les lignes de
        #    données contiennent des valeurs numériques (n° d'ordre, matricule,
        #    montants...) tandis que les lignes d'en-tête sont purement
        #    textuelles. On étend l'en-tête tant que la ligne ne contient aucune
        #    cellule numérique, en s'arrêtant à la première ligne de données.
        #    Borné à 5 lignes pour rester robuste.
        #
        #    Garde-fou: ce critère ne fonctionne que si les DONNÉES contiennent
        #    des nombres. Pour un tableau 100% textuel (annuaire, etc.), aucune
        #    ligne ne servirait de butoir et l'en-tête engloutirait les données.
        #    Dans ce cas on retombe sur le défaut sûr : 1 seule ligne d'en-tête.
        header_end = header_start
        max_header_end = min(header_start + 4, data_end)
        data_zone_has_numbers = any(
            self._row_numeric_count(df, rn - 1) > 0
            for rn in range(header_start + 1, data_end + 1)
        )
        if data_zone_has_numbers:
            for row_num in range(header_start + 1, max_header_end + 1):
                if self._row_numeric_count(df, row_num - 1) == 0:
                    header_end = row_num
                else:
                    break
        else:
            # Aucun nombre dans la zone de données : le critère par contenu est
            # aveugle (données « NEANT », texte libre, ou fichier non rempli).
            # Les fusions verticales décrivent alors la structure de l'en-tête.
            # Sans ce repli, ces fichiers étaient détectés avec 1 seule ligne
            # d'en-tête ; leurs libellés n'étaient plus aplatis comme ceux du
            # reste du lot et l'aligneur créait des colonnes en double.
            merged_end = self._header_end_from_vertical_merges(
                worksheet, header_start, max_header_end
            )
            if merged_end is not None:
                header_end = merged_end

        return (header_start, header_end, data_end)

    def _row_numeric_count(self, df, row_idx: int) -> int:
        """
        Nombre de cellules numériquement interprétables dans une ligne (base 0).

        Sert à distinguer une ligne d'en-tête (libellés textuels) d'une ligne de
        données (qui contient des valeurs : nombres, dates, booléens...). Toute
        valeur convertible en nombre compte ; les libellés textuels ne comptent
        pas. Une ligne hors limites renvoie 0.
        """
        if row_idx >= len(df):
            return 0
        count = 0
        for val in df.iloc[row_idx]:
            if pd.notna(val):
                try:
                    float(val)
                    count += 1
                except (ValueError, TypeError):
                    pass
        return count

    def _calculate_confidence(self, border_analysis: Dict[int, Dict],
                            table_zone: Tuple[int, int, int]) -> float:
        """
        Calcule le score de confiance de la détection

        Args:
            border_analysis: Résultat de _analyze_borders
            table_zone: (header_start, header_end, data_end)

        Returns:
            Score entre 0.0 et 1.0
        """
        header_start, header_end, data_end = table_zone

        # Calculer la densité moyenne des bordures dans la zone détectée
        total_density = 0
        rows_in_zone = 0

        for row_num in range(header_start, data_end + 1):
            if row_num in border_analysis:
                total_density += border_analysis[row_num]['border_density']
                rows_in_zone += 1

        avg_density = total_density / rows_in_zone if rows_in_zone > 0 else 0

        # Calculer la densité des bordures complètes dans les en-têtes
        header_full_density = 0
        header_rows = 0

        for row_num in range(header_start, header_end + 1):
            if row_num in border_analysis:
                header_full_density += border_analysis[row_num]['full_border_density']
                header_rows += 1

        avg_header_density = header_full_density / header_rows if header_rows > 0 else 0

        # Score basé sur:
        # - 60% densité moyenne dans la zone
        # - 40% densité des bordures complètes dans les en-têtes
        confidence = (avg_density * 0.6) + (avg_header_density * 0.4)

        # Bonus si les bordures sont très cohérentes
        if avg_density > 0.7:
            confidence = min(1.0, confidence + 0.1)

        return min(1.0, max(0.0, confidence))

    def _detect_trailing_content(self, df: pd.DataFrame,
                                 border_analysis: Dict[int, Dict],
                                 data_end: int) -> Dict:
        """
        Détecte le contenu après le tableau

        Args:
            df: DataFrame
            border_analysis: Résultat de _analyze_borders
            data_end: Dernière ligne de données (base 1)

        Returns:
            Dict avec infos trailing
        """
        total_rows = len(df)

        # S'il n'y a pas de lignes après data_end
        if data_end >= total_rows:
            return {'has_trailing': False}

        # Analyser les lignes après data_end
        trailing_start = data_end + 1
        trailing_end = total_rows

        # Compter les lignes non vides après data_end
        non_empty_rows = 0
        for row_idx in range(data_end, total_rows):  # pandas base 0
            density = self.calculate_row_density(df, row_idx)
            if density > 0.1:  # Au moins 10% de cellules remplies
                non_empty_rows += 1

        has_trailing = non_empty_rows > 0

        return {
            'has_trailing': has_trailing,
            'start_row': trailing_start if has_trailing else None,
            'end_row': trailing_end if has_trailing else None,
            'row_count': non_empty_rows
        }

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
            detection_method="border",
            warning=reason
        )
