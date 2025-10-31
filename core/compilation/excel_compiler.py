"""
Moteur principal de compilation Excel avec détection automatique
Auteur: GOUNOU N'GOBI Chabi Zimé
Version: 3.2
"""

import time
import logging
from pathlib import Path
from typing import List, Optional, Dict, Any, Tuple
import pandas as pd
import numpy as np

from .compilation_models import (
    CompilationOptions,
    CompilationResult,
    FileCompilationResult,
    FilenameOption,
    DateFormat,
    OutputFormat
)
from ..detection import HybridDetector, DetectionResult
from utils import logger


class ExcelCompiler:
    """
    Moteur de compilation principal avec détection automatique

    Compile plusieurs fichiers Excel/CSV en un seul fichier,
    en utilisant la détection automatique de structure (v3.2)
    ou les paramètres manuels (fallback).

    Fonctionnalités:
    - Détection automatique par fichier
    - Validation croisée entre fichiers
    - Traitement par chunks pour gros fichiers
    - Gestion mémoire optimisée
    - Support multiple formats (Excel, CSV, TSV)
    """

    def __init__(self, options: Optional[CompilationOptions] = None):
        """
        Initialise le compilateur

        Args:
            options: Options de compilation (utilise les défauts si None)
        """
        self.options = options or CompilationOptions()
        self.logger = logger

        # Détecteur de structure (si détection auto activée)
        self.detector: Optional[HybridDetector] = None
        if self.options.auto_detect_structure:
            self.detector = HybridDetector(
                confidence_threshold=self.options.detection_confidence_threshold,
                enable_cross_validation=self.options.enable_cross_validation
            )

    def compile_files(self, file_paths: List[str],
                     output_file: str,
                     output_format: OutputFormat = OutputFormat.XLSX) -> CompilationResult:
        """
        Compile plusieurs fichiers en un seul

        Args:
            file_paths: Liste des chemins de fichiers à compiler
            output_file: Chemin du fichier de sortie
            output_format: Format de sortie

        Returns:
            CompilationResult avec statistiques et détails
        """
        start_time = time.time()

        # Initialiser le résultat
        result = CompilationResult(
            total_files=len(file_paths),
            output_file=output_file
        )

        self.logger.info(f"Début compilation de {len(file_paths)} fichiers")

        try:
            # Étape 1: Détection automatique si activée
            detection_results = self._detect_structures(file_paths, result)

            # Étape 2: Charger et compiler les données
            combined_data, headers = self._load_and_combine_files(
                file_paths,
                detection_results,
                result
            )

            if not combined_data or not headers:
                result.success = False
                result.warnings.append("Aucune donnée compilée")
                return result

            # Étape 3: Post-traitement (tri, doublons, etc.)
            combined_data = self._post_process_data(combined_data, headers, result)

            # Étape 4: Écrire le fichier de sortie
            self._write_output_file(
                combined_data,
                headers,
                output_file,
                output_format,
                result
            )

            # Finaliser les statistiques
            result.final_headers = [str(h) for h in headers[-1]] if headers else []
            result.total_processing_time = time.time() - start_time
            result.calculate_statistics()

            self.logger.info(f"Compilation terminée: {result.successful_files}/{result.total_files} fichiers")

            return result

        except Exception as e:
            self.logger.error(f"Erreur compilation: {e}", exc_info=True)
            result.success = False
            result.warnings.append(f"Erreur fatale: {e}")
            result.total_processing_time = time.time() - start_time
            return result

    def _detect_structures(self, file_paths: List[str],
                          result: CompilationResult) -> Dict[str, DetectionResult]:
        """
        Détecte la structure de chaque fichier

        Args:
            file_paths: Liste des fichiers
            result: Résultat de compilation (pour warnings)

        Returns:
            Dict {file_path: DetectionResult}
        """
        detection_results = {}

        if not self.options.auto_detect_structure or self.detector is None:
            self.logger.info("Détection automatique désactivée, utilisation paramètres manuels")
            return detection_results

        self.logger.info(f"Détection automatique de structure sur {len(file_paths)} fichiers")

        try:
            # Détection batch pour validation croisée optimale
            batch_results = self.detector.detect_batch(file_paths)

            for detection in batch_results:
                detection_results[detection.file_path] = detection

                # Avertir si confiance faible
                if detection.confidence < self.options.detection_confidence_threshold:
                    result.warnings.append(
                        f"{Path(detection.file_path).name}: "
                        f"Détection faible confiance ({detection.confidence:.0%})"
                    )

                # Avertir si cross-validation suspecte
                if detection.warning:
                    result.warnings.append(
                        f"{Path(detection.file_path).name}: {detection.warning}"
                    )

        except Exception as e:
            self.logger.warning(f"Erreur détection automatique: {e}")
            result.warnings.append(f"Détection automatique échouée, utilisation paramètres manuels")

        return detection_results

    def _load_and_combine_files(self, file_paths: List[str],
                                detection_results: Dict[str, DetectionResult],
                                result: CompilationResult) -> Tuple[List[List], List[List]]:
        """
        Charge et combine les données de tous les fichiers

        Args:
            file_paths: Liste des fichiers
            detection_results: Résultats de détection par fichier
            result: Résultat de compilation

        Returns:
            Tuple (combined_data, headers)
        """
        combined_data = []
        global_headers = None
        preliminary_info = []

        # Charger les informations préliminaires si demandées
        if self.options.include_preliminary and self.options.preliminary_source_file:
            preliminary_info = self._load_preliminary_info(
                self.options.preliminary_source_file
            )

        # Traiter chaque fichier
        for i, file_path in enumerate(file_paths):
            file_start_time = time.time()

            try:
                self.logger.info(f"Traitement fichier {i+1}/{len(file_paths)}: {Path(file_path).name}")

                # Charger le fichier avec détection ou paramètres manuels
                file_data, file_headers, detection_info = self._load_single_file(
                    file_path,
                    detection_results.get(file_path),
                    i == 0 and preliminary_info
                )

                if file_data is None:
                    # Erreur de chargement
                    file_result = FileCompilationResult(
                        file_path=file_path,
                        success=False,
                        error_message="Erreur chargement fichier"
                    )
                    result.add_file_result(file_result)
                    continue

                # Premier fichier: initialiser les en-têtes globaux
                if global_headers is None:
                    global_headers = file_headers

                    # Si on ajoute le nom de fichier, ajouter la colonne aux en-têtes
                    if self.options.filename_option != FilenameOption.NONE:
                        global_headers = self._add_filename_header(global_headers)

                # Vérifier compatibilité des en-têtes
                if not self._are_headers_compatible(global_headers, file_headers):
                    result.warnings.append(
                        f"{Path(file_path).name}: En-têtes incompatibles, "
                        f"ajustement automatique"
                    )
                    file_data = self._adjust_columns(file_data, file_headers, global_headers)

                # Ajouter le nom du fichier si demandé
                if self.options.filename_option != FilenameOption.NONE:
                    file_data = self._add_filename_column(
                        file_data,
                        file_path,
                        self.options.filename_option
                    )

                # Ajouter les en-têtes répétés si demandé
                if self.options.repeat_headers and i > 0:
                    # Utiliser les en-têtes globaux (qui incluent "Fichier source" si nécessaire)
                    combined_data.extend(global_headers)

                # Ajouter les données
                combined_data.extend(file_data)

                # Enregistrer le résultat du fichier
                file_result = FileCompilationResult(
                    file_path=file_path,
                    success=True,
                    rows_added=len(file_data),
                    processing_time=time.time() - file_start_time,
                    **detection_info
                )
                result.add_file_result(file_result)

            except Exception as e:
                self.logger.error(f"Erreur traitement {file_path}: {e}")
                file_result = FileCompilationResult(
                    file_path=file_path,
                    success=False,
                    error_message=str(e),
                    processing_time=time.time() - file_start_time
                )
                result.add_file_result(file_result)

        # Ajouter les informations préliminaires au début si demandées
        if preliminary_info and combined_data:
            combined_data = preliminary_info + combined_data

        return combined_data, global_headers if global_headers else []

    def _load_single_file(self, file_path: str,
                         detection: Optional[DetectionResult],
                         include_preliminary: bool = False) -> Tuple[Optional[List], List, Dict]:
        """
        Charge un fichier individuel

        Args:
            file_path: Chemin du fichier
            detection: Résultat de détection (None si manuel)
            include_preliminary: Inclure infos préliminaires de ce fichier

        Returns:
            Tuple (data, headers, detection_info_dict)
        """
        ext = Path(file_path).suffix.lower()

        # Charger selon le format
        if ext in ['.xlsx', '.xls', '.xlsm']:
            return self._load_excel_file(file_path, detection, include_preliminary)
        elif ext == '.csv':
            return self._load_csv_file(file_path, detection, include_preliminary)
        elif ext in ['.tsv', '.txt']:
            return self._load_tsv_file(file_path, detection, include_preliminary)
        else:
            self.logger.warning(f"Format non supporté: {ext}")
            return None, [], {}

    def _load_excel_file(self, file_path: str,
                        detection: Optional[DetectionResult],
                        include_preliminary: bool) -> Tuple[Optional[List], List, Dict]:
        """
        Charge un fichier Excel

        Args:
            file_path: Chemin du fichier
            detection: Résultat de détection
            include_preliminary: Inclure infos préliminaires

        Returns:
            Tuple (data, headers, detection_info)
        """
        # Déterminer les paramètres de chargement
        if detection and detection.confidence >= self.options.detection_confidence_threshold:
            # Utiliser la détection
            header_start_row = detection.header_start_row
            header_rows = detection.header_rows
            data_start_row = detection.data_start_row
            data_end_row = detection.data_end_row if detection.data_end_row > 0 else None

            detection_info = {
                'detection_used': True,
                'header_start_row': header_start_row,
                'header_rows': header_rows,
                'data_start_row': data_start_row,
                'data_end_row': data_end_row or 0,
                'detection_confidence': detection.confidence,
                'detection_method': detection.detection_method
            }
        else:
            # Utiliser les paramètres manuels
            header_start_row = self.options.manual_header_start_row
            header_rows = self.options.manual_header_rows
            data_start_row = header_start_row + header_rows
            data_end_row = None

            detection_info = {
                'detection_used': False,
                'header_start_row': header_start_row,
                'header_rows': header_rows,
                'data_start_row': data_start_row,
                'data_end_row': 0,
                'detection_confidence': 0.0,
                'detection_method': 'manual'
            }

        # Charger avec pandas
        df = pd.read_excel(file_path, header=None)

        # Extraire les en-têtes
        headers = []
        for row_idx in range(header_start_row - 1, header_start_row - 1 + header_rows):
            if row_idx < len(df):
                header_row = df.iloc[row_idx].tolist()
                headers.append(header_row)

        # Extraire les données
        data = []
        start_idx = data_start_row - 1
        end_idx = (data_end_row if data_end_row else len(df))

        for row_idx in range(start_idx, end_idx):
            if row_idx < len(df):
                row = df.iloc[row_idx].tolist()

                # Supprimer lignes vides si demandé
                if self.options.remove_empty_rows and self._is_row_empty(row):
                    continue

                data.append(row)

        return data, headers, detection_info

    def _load_csv_file(self, file_path: str,
                      detection: Optional[DetectionResult],
                      include_preliminary: bool) -> Tuple[Optional[List], List, Dict]:
        """Charge un fichier CSV (similaire à Excel mais avec encodage)"""
        # Essayer différents encodages
        encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']

        for encoding in encodings:
            try:
                df = pd.read_csv(file_path, header=None, encoding=encoding)
                # Utiliser la même logique que Excel
                return self._extract_data_from_dataframe(df, detection, include_preliminary)
            except UnicodeDecodeError:
                continue

        raise ValueError(f"Impossible de décoder {file_path}")

    def _load_tsv_file(self, file_path: str,
                      detection: Optional[DetectionResult],
                      include_preliminary: bool) -> Tuple[Optional[List], List, Dict]:
        """Charge un fichier TSV"""
        df = pd.read_csv(file_path, header=None, sep='\t', encoding='utf-8-sig')
        return self._extract_data_from_dataframe(df, detection, include_preliminary)

    def _extract_data_from_dataframe(self, df: pd.DataFrame,
                                    detection: Optional[DetectionResult],
                                    include_preliminary: bool) -> Tuple[List, List, Dict]:
        """Extrait données et en-têtes d'un DataFrame (logique commune CSV/TSV/Excel)"""
        # Déterminer les paramètres de chargement
        if detection and detection.confidence >= self.options.detection_confidence_threshold:
            # Utiliser la détection
            header_start_row = detection.header_start_row
            header_rows = detection.header_rows
            data_start_row = detection.data_start_row
            data_end_row = detection.data_end_row if detection.data_end_row > 0 else None

            detection_info = {
                'detection_used': True,
                'header_start_row': header_start_row,
                'header_rows': header_rows,
                'data_start_row': data_start_row,
                'data_end_row': data_end_row or 0,
                'detection_confidence': detection.confidence,
                'detection_method': detection.detection_method
            }
        else:
            # Utiliser les paramètres manuels
            header_start_row = self.options.manual_header_start_row
            header_rows = self.options.manual_header_rows
            data_start_row = header_start_row + header_rows
            data_end_row = None

            detection_info = {
                'detection_used': False,
                'header_start_row': header_start_row,
                'header_rows': header_rows,
                'data_start_row': data_start_row,
                'data_end_row': 0,
                'detection_confidence': 0.0,
                'detection_method': 'manual'
            }

        # Extraire les en-têtes
        headers = []
        for row_idx in range(header_start_row - 1, header_start_row - 1 + header_rows):
            if row_idx < len(df):
                header_row = df.iloc[row_idx].tolist()
                headers.append(header_row)

        # Extraire les données
        data = []
        start_idx = data_start_row - 1
        end_idx = (data_end_row if data_end_row else len(df))

        for row_idx in range(start_idx, end_idx):
            if row_idx < len(df):
                row = df.iloc[row_idx].tolist()

                # Supprimer lignes vides si demandé
                if self.options.remove_empty_rows and self._is_row_empty(row):
                    continue

                data.append(row)

        return data, headers, detection_info

    def _is_row_empty(self, row: List) -> bool:
        """Vérifie si une ligne est vide"""
        return all(pd.isna(val) or val == '' for val in row)

    def _are_headers_compatible(self, headers1: List[List], headers2: List[List]) -> bool:
        """Vérifie si deux en-têtes sont compatibles"""
        if not headers1 or not headers2:
            return False

        # Comparer les dernières lignes d'en-têtes (les plus importantes)
        h1 = headers1[-1] if headers1 else []
        h2 = headers2[-1] if headers2 else []

        return len(h1) == len(h2)

    def _adjust_columns(self, data: List[List], source_headers: List[List],
                       target_headers: List[List]) -> List[List]:
        """Ajuste les colonnes pour correspondre aux en-têtes cibles"""
        # Ajouter des colonnes vides si nécessaire
        target_col_count = len(target_headers[-1]) if target_headers else 0
        source_col_count = len(source_headers[-1]) if source_headers else 0

        if source_col_count >= target_col_count:
            return data  # Aucun ajustement nécessaire

        adjusted_data = []
        for row in data:
            adjusted_row = row + [None] * (target_col_count - len(row))
            adjusted_data.append(adjusted_row)

        return adjusted_data

    def _add_filename_header(self, headers: List[List]) -> List[List]:
        """
        Ajoute 'Fichier source' aux en-têtes

        Args:
            headers: En-têtes existants

        Returns:
            En-têtes avec colonne ajoutée
        """
        new_headers = []
        for i, header_row in enumerate(headers):
            if i == len(headers) - 1:
                # Dernière ligne d'en-têtes: ajouter "Fichier source"
                new_headers.append(header_row + ["Fichier source"])
            else:
                # Autres lignes: ajouter cellule vide
                new_headers.append(header_row + [""])
        return new_headers

    def _add_filename_column(self, data: List[List], file_path: str,
                           option: FilenameOption) -> List[List]:
        """Ajoute une colonne avec le nom du fichier"""
        filename = Path(file_path).name

        if option == FilenameOption.WITHOUT_EXTENSION:
            filename = Path(file_path).stem

        # Ajouter la colonne à chaque ligne
        return [row + [filename] for row in data]

    def _load_preliminary_info(self, source_file: str) -> List[List]:
        """
        Charge les informations préliminaires d'un fichier

        Les informations préliminaires sont les lignes AVANT les en-têtes
        (titre, contexte, date, etc.)

        Args:
            source_file: Chemin du fichier source

        Returns:
            Liste de lignes préliminaires
        """
        if not Path(source_file).exists():
            self.logger.warning(f"Fichier préliminaire introuvable: {source_file}")
            return []

        try:
            ext = Path(source_file).suffix.lower()

            # Charger le fichier
            if ext in ['.xlsx', '.xls', '.xlsm']:
                df = pd.read_excel(source_file, header=None)
            elif ext == '.csv':
                # Essayer différents encodages
                df = None
                for encoding in ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']:
                    try:
                        df = pd.read_csv(source_file, header=None, encoding=encoding)
                        break
                    except UnicodeDecodeError:
                        continue
                if df is None:
                    raise ValueError("Impossible de décoder le fichier CSV")
            else:
                self.logger.warning(f"Format non supporté pour preliminary: {ext}")
                return []

            # Déterminer où se termine le preliminary
            # (avant le header_start_row)
            header_start_row = self.options.manual_header_start_row

            if header_start_row <= 1:
                # Pas de lignes avant les en-têtes
                return []

            # Extraire les lignes avant les en-têtes
            preliminary = []
            for row_idx in range(0, header_start_row - 1):
                if row_idx < len(df):
                    row = df.iloc[row_idx].tolist()
                    preliminary.append(row)

            self.logger.info(f"Chargé {len(preliminary)} lignes préliminaires de {Path(source_file).name}")
            return preliminary

        except Exception as e:
            self.logger.error(f"Erreur chargement preliminary: {e}")
            return []

    def _post_process_data(self, data: List[List], headers: List[List],
                          result: CompilationResult) -> List[List]:
        """
        Post-traitement des données (tri, doublons, etc.)

        Args:
            data: Données combinées
            headers: En-têtes
            result: Résultat (pour warnings)

        Returns:
            Données traitées
        """
        # Supprimer les doublons
        if self.options.remove_duplicates:
            self.logger.info("Suppression des doublons")
            data = self._remove_duplicates(data, headers, result)

        # Trier les données
        if self.options.sort_data:
            self.logger.info(f"Tri des données par colonne {self.options.sort_column}")
            data = self._sort_data(data, self.options.sort_column)

        return data

    def _remove_duplicates(self, data: List[List], headers: List[List],
                          result: CompilationResult) -> List[List]:
        """Supprime les lignes dupliquées"""
        original_count = len(data)

        # Convertir en tuples pour utiliser set
        unique_rows = []
        seen = set()

        for row in data:
            row_tuple = tuple(str(v) for v in row)
            if row_tuple not in seen:
                seen.add(row_tuple)
                unique_rows.append(row)

        removed_count = original_count - len(unique_rows)
        if removed_count > 0:
            result.warnings.append(f"{removed_count} doublons supprimés")

        return unique_rows

    def _sort_data(self, data: List[List], sort_column: int) -> List[List]:
        """Trie les données par colonne"""
        try:
            return sorted(data, key=lambda row: row[sort_column] if sort_column < len(row) else '')
        except Exception as e:
            self.logger.warning(f"Erreur tri: {e}")
            return data

    def _write_output_file(self, data: List[List], headers: List[List],
                          output_file: str, output_format: OutputFormat,
                          result: CompilationResult):
        """
        Écrit le fichier de sortie

        Args:
            data: Données à écrire
            headers: En-têtes
            output_file: Chemin de sortie
            output_format: Format de sortie
            result: Résultat (pour statistiques)
        """
        self.logger.info(f"Écriture fichier de sortie: {output_file}")

        # Combiner en-têtes et données
        all_data = headers + data

        # Convertir en DataFrame
        df = pd.DataFrame(all_data)

        # Écrire selon le format
        if output_format == OutputFormat.XLSX:
            df.to_excel(output_file, index=False, header=False)
        elif output_format == OutputFormat.CSV:
            df.to_csv(output_file, index=False, header=False, encoding='utf-8-sig')
        elif output_format == OutputFormat.TSV:
            df.to_csv(output_file, index=False, header=False, sep='\t', encoding='utf-8-sig')

        result.output_file = output_file
        result.success = True
