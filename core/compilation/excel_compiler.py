"""
Moteur principal de compilation Excel avec détection automatique
Auteur: GOUNOU N'GOBI Chabi Zimé
Version: 3.2
"""

import time
import logging
from pathlib import Path
from typing import List, Optional, Dict, Any, Tuple, Callable
import pandas as pd
import numpy as np
import openpyxl

from .compilation_models import (
    CompilationOptions,
    CompilationResult,
    FileCompilationResult,
    FilePreview,
    FilenameOption,
    DateFormat,
    OutputFormat
)
from ..detection import HybridDetector, ReferenceDetector, DetectionResult
from ..detection.base_detector import prune_phantom_columns
from .excel_formatter import ExcelFormatter
from utils import logger


class CompilationCancelled(Exception):
    """Levée en interne quand l'utilisateur annule la compilation."""
    pass


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

        # Valeur effective de repeat_headers pour la compilation courante
        # (peut être neutralisée par compile_files si tri/dédup actifs, sans
        # muter self.options). Initialisée ici pour robustesse.
        self._effective_repeat_headers = self.options.repeat_headers

        # Callbacks de progression / annulation (définis par compile_files)
        self._progress_callback: Optional[Callable[[int, str], None]] = None
        self._cancel_check: Optional[Callable[[], bool]] = None

        # Détecteur de structure selon le mode choisi
        self.detector: Optional[HybridDetector] = None
        self.reference_detector: Optional[ReferenceDetector] = None

        if self.options.use_reference_mode:
            # Mode référence: pas de détecteur pour l'instant
            # On le créera dans _detect_structures avec le premier fichier
            self.logger.info("Mode référence activé")
        elif self.options.auto_detect_structure:
            # Mode automatique hybride
            self.detector = HybridDetector(
                confidence_threshold=self.options.detection_confidence_threshold,
                enable_cross_validation=self.options.enable_cross_validation
            )
            self.logger.info("Mode détection automatique hybride activé")

    def compile_files(self, file_paths: List[str],
                     output_file: str,
                     output_format: OutputFormat = OutputFormat.XLSX,
                     progress_callback: Optional[Callable[[int, str], None]] = None,
                     cancel_check: Optional[Callable[[], bool]] = None) -> CompilationResult:
        """
        Compile plusieurs fichiers en un seul

        Args:
            file_paths: Liste des chemins de fichiers à compiler
            output_file: Chemin du fichier de sortie
            output_format: Format de sortie
            progress_callback: Fonction optionnelle appelée avec (pourcentage,
                message) à chaque étape pour suivre l'avancement réel.
            cancel_check: Fonction optionnelle renvoyant True si l'utilisateur a
                demandé l'annulation. Consultée avant chaque fichier; la
                compilation s'arrête alors proprement.

        Returns:
            CompilationResult avec statistiques et détails
        """
        start_time = time.time()
        self._progress_callback = progress_callback
        self._cancel_check = cancel_check

        # Initialiser le résultat
        result = CompilationResult(
            total_files=len(file_paths),
            output_file=output_file
        )

        self.logger.info(f"Début compilation de {len(file_paths)} fichiers")

        # Garde-fou: répéter les en-têtes entre fichiers est incompatible avec
        # le tri ou la déduplication (les lignes d'en-tête réinjectées seraient
        # traitées comme des données et triées/dédupliquées au milieu du tableau).
        # Les options de transformation priment; on neutralise la répétition pour
        # CETTE compilation uniquement, sans muter l'objet options de l'appelant.
        self._effective_repeat_headers = self.options.repeat_headers
        if self.options.repeat_headers and (self.options.sort_data
                                            or self.options.remove_duplicates):
            self._effective_repeat_headers = False
            warn = ("Option 'répéter les en-têtes' désactivée car incompatible "
                    "avec le tri / la suppression des doublons")
            self.logger.warning(warn)
            result.warnings.append(warn)

        try:
            # Étape 1: Détection automatique si activée
            self._emit_progress(5, "Analyse de la structure des fichiers...")
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
            self._emit_progress(82, "Post-traitement (tri, doublons)...")
            combined_data = self._post_process_data(combined_data, headers, result)

            # Étape 4: Écrire le fichier de sortie
            self._emit_progress(90, "Écriture du fichier de sortie...")
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

            self._emit_progress(100, "Compilation terminée")
            self.logger.info(f"Compilation terminée: {result.successful_files}/{result.total_files} fichiers")

            return result

        except CompilationCancelled:
            self.logger.info("Compilation annulée par l'utilisateur")
            result.success = False
            result.cancelled = True
            result.warnings.append("Compilation annulée par l'utilisateur")
            result.total_processing_time = time.time() - start_time
            return result

        except Exception as e:
            self.logger.error(f"Erreur compilation: {e}", exc_info=True)
            result.success = False
            result.warnings.append(f"Erreur fatale: {e}")
            result.total_processing_time = time.time() - start_time
            return result

    def _emit_progress(self, percent: int, message: str):
        """Notifie la progression si un callback est défini (jamais bloquant)."""
        if self._progress_callback:
            try:
                self._progress_callback(percent, message)
            except Exception as e:
                self.logger.warning(f"Callback de progression a échoué: {e}")

    def _check_cancel(self):
        """Lève CompilationCancelled si l'utilisateur a demandé l'annulation."""
        if self._cancel_check and self._cancel_check():
            raise CompilationCancelled()

    def preview_detection(self, file_paths: List[str]) -> List[FilePreview]:
        """
        Aperçu de la détection pour chaque fichier, SANS compiler.

        Utilise exactement la même logique de détection que la compilation
        (mode automatique, référence ou manuel selon les options), de sorte
        que l'aperçu reflète fidèlement ce qui sera produit. Chaque fichier
        est traité indépendamment : une erreur sur l'un n'interrompt pas les
        autres.

        Returns:
            Liste de FilePreview, dans l'ordre des fichiers fournis.
        """
        # Réutilise la détection réelle (un CompilationResult sert de réceptacle
        # aux éventuels warnings, qu'on n'expose pas ici).
        scratch = CompilationResult(total_files=len(file_paths))
        detections = self._detect_structures(file_paths, scratch)

        previews: List[FilePreview] = []
        for file_path in file_paths:
            previews.append(self._build_preview(file_path, detections.get(file_path)))
        return previews

    def _build_preview(self, file_path: str,
                       detection: Optional[DetectionResult]) -> FilePreview:
        """Construit l'aperçu d'un fichier à partir de sa détection (ou du
        mode manuel si aucune détection n'est disponible)."""
        try:
            # Déterminer les paramètres effectifs via la MÊME logique que la
            # compilation (override > détection > manuel) — zéro divergence.
            structure = self._resolve_structure(file_path, detection)
            header_start = structure['header_start_row']
            header_rows = structure['header_rows']
            data_start = structure['data_start_row']
            data_end = structure['data_end_row']
            confidence = structure['detection_confidence']
            method = structure['detection_method']
            # Warning : celui de la détection seulement si elle est réellement
            # utilisée ; sinon (override / manuel) message adapté.
            if method == 'override':
                warning = "Correction manuelle appliquée par l'utilisateur."
            elif structure['detection_used']:
                warning = detection.warning
            else:
                warning = None

            # Lire le fichier pour extraire en-têtes et compter les données.
            # On élague les colonnes parasites comme la compilation réelle, pour
            # que l'aperçu reflète fidèlement ce qui sera produit.
            ext = Path(file_path).suffix.lower()
            if ext in ['.xlsx', '.xls', '.xlsm']:
                df = pd.read_excel(file_path, header=None)
                df, _ = prune_phantom_columns(df)
            elif ext == '.csv':
                df = None
                for enc in ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']:
                    try:
                        df = pd.read_csv(file_path, header=None, encoding=enc)
                        break
                    except UnicodeDecodeError:
                        continue
                if df is None:
                    raise ValueError("Impossible de décoder le fichier CSV")
            elif ext in ['.tsv', '.txt']:
                df = pd.read_csv(file_path, header=None, sep='\t', encoding='utf-8-sig')
            else:
                return FilePreview(file_path=file_path, success=False,
                                   error=f"Format non supporté: {ext}")

            if ext != '.xlsx' and ext != '.xls' and ext != '.xlsm':
                df, _ = prune_phantom_columns(df)

            n = len(df)
            # En-têtes (fusion multi-lignes via le helper du détecteur de base)
            headers: List[str] = []
            for idx in range(header_start - 1, min(header_start - 1 + header_rows, n)):
                headers.append([
                    str(v) if pd.notna(v) else "" for v in df.iloc[idx].tolist()
                ])
            flat_headers = self._flatten_headers(headers)

            # Étendue des données. On compte les lignes EXACTEMENT comme le
            # chargement réel : dans la plage [data_start, end[, en excluant les
            # lignes vides si remove_empty_rows est actif (sinon l'aperçu
            # surestimerait le nombre de lignes par rapport à la compilation).
            end = data_end if (data_end and data_end > 0) else n
            end = min(end, n)
            data_row_count = 0
            for idx in range(data_start - 1, end):
                if idx < 0 or idx >= n:
                    continue
                row = df.iloc[idx].tolist()
                if self.options.remove_empty_rows and self._is_row_empty(row):
                    continue
                data_row_count += 1

            return FilePreview(
                file_path=file_path,
                success=True,
                header_start_row=header_start,
                header_rows=header_rows,
                data_start_row=data_start,
                data_end_row=end,
                detected_headers=flat_headers,
                confidence=confidence,
                detection_method=method,
                data_row_count=data_row_count,
                warning=warning,
            )

        except Exception as e:
            self.logger.warning(f"Aperçu impossible pour {Path(file_path).name}: {e}")
            return FilePreview(file_path=file_path, success=False, error=str(e))

    def _flatten_headers(self, header_rows: List[List]) -> List[str]:
        """Fusionne d'éventuelles lignes d'en-tête multiples en une liste
        de libellés (séparateur ' - '), pour affichage."""
        if not header_rows:
            return []
        if len(header_rows) == 1:
            return [str(h) for h in header_rows[0]]
        col_count = max(len(r) for r in header_rows)
        flat = []
        for col in range(col_count):
            parts = []
            for row in header_rows:
                if col < len(row) and row[col]:
                    parts.append(str(row[col]))
            flat.append(" - ".join(parts))
        return flat

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

        # MODE 1: Fichier de référence
        if self.options.use_reference_mode:
            self.logger.info(f"Mode référence: détection basée sur le fichier de référence")
            return self._detect_with_reference(file_paths, result)

        # MODE 2: Détection automatique hybride
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
        total = len(file_paths)
        for i, file_path in enumerate(file_paths):
            file_start_time = time.time()

            # Permettre l'annulation avant de démarrer chaque fichier
            self._check_cancel()

            # Progression: répartir la plage 15%->80% sur les fichiers
            if total > 0:
                pct = 15 + int((i / total) * 65)
                self._emit_progress(
                    pct,
                    f"Traitement {i+1}/{total} : {Path(file_path).name}"
                )

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

                # Ajouter le nom du fichier si demandé
                if self.options.filename_option != FilenameOption.NONE:
                    file_data = self._add_filename_column(
                        file_data,
                        file_path,
                        self.options.filename_option
                    )
                    # Ajouter aussi la colonne aux en-têtes du fichier pour la comparaison
                    file_headers = self._add_filename_header(file_headers)

                # Vérifier compatibilité des en-têtes
                if not self._are_headers_compatible(global_headers, file_headers):
                    result.warnings.append(
                        f"{Path(file_path).name}: En-têtes incompatibles, "
                        f"ajustement automatique"
                    )
                    file_data = self._adjust_columns(file_data, file_headers, global_headers)

                # Ajouter les en-têtes répétés si demandé
                if self._effective_repeat_headers and i > 0:
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

    def _resolve_structure(self, file_path: str,
                           detection: Optional[DetectionResult]) -> Dict[str, Any]:
        """Détermine la structure effective d'un fichier selon l'ordre de
        priorité, identique pour l'aperçu ET la compilation (zéro divergence) :

            1. Correction manuelle par fichier (manual_overrides) — priorité absolue
            2. Détection automatique/référence si confiance >= seuil
            3. Paramètres manuels globaux (fallback)

        Returns:
            Dict avec les clés de FileCompilationResult :
            detection_used, header_start_row, header_rows, data_start_row,
            data_end_row, detection_confidence, detection_method.
        """
        override = self.options.manual_overrides.get(file_path)
        if override is not None:
            header_start_row, header_rows = override
            return {
                'detection_used': False,
                'header_start_row': header_start_row,
                'header_rows': header_rows,
                'data_start_row': header_start_row + header_rows,
                'data_end_row': 0,
                'detection_confidence': 1.0,
                'detection_method': 'override',
            }

        if detection and detection.confidence >= self.options.detection_confidence_threshold:
            return {
                'detection_used': True,
                'header_start_row': detection.header_start_row,
                'header_rows': detection.header_rows,
                'data_start_row': detection.data_start_row,
                'data_end_row': detection.data_end_row if detection.data_end_row > 0 else 0,
                'detection_confidence': detection.confidence,
                'detection_method': detection.detection_method,
            }

        header_start_row = self.options.manual_header_start_row
        header_rows = self.options.manual_header_rows
        return {
            'detection_used': False,
            'header_start_row': header_start_row,
            'header_rows': header_rows,
            'data_start_row': header_start_row + header_rows,
            'data_end_row': 0,
            'detection_confidence': 0.0,
            'detection_method': 'manual',
        }

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
        # Déterminer les paramètres de chargement (override > détection > manuel)
        detection_info = self._resolve_structure(file_path, detection)
        header_start_row = detection_info['header_start_row']
        header_rows = detection_info['header_rows']
        data_start_row = detection_info['data_start_row']
        data_end_row = detection_info['data_end_row'] or None

        # Charger avec pandas (et élaguer les colonnes parasites pour rester
        # cohérent avec la détection — sinon l'export traînerait des milliers
        # de colonnes vides).
        df = pd.read_excel(file_path, header=None)
        df, _ = prune_phantom_columns(df)

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
                return self._extract_data_from_dataframe(df, file_path, detection, include_preliminary)
            except UnicodeDecodeError:
                continue

        raise ValueError(f"Impossible de décoder {file_path}")

    def _load_tsv_file(self, file_path: str,
                      detection: Optional[DetectionResult],
                      include_preliminary: bool) -> Tuple[Optional[List], List, Dict]:
        """Charge un fichier TSV"""
        df = pd.read_csv(file_path, header=None, sep='\t', encoding='utf-8-sig')
        return self._extract_data_from_dataframe(df, file_path, detection, include_preliminary)

    def _extract_data_from_dataframe(self, df: pd.DataFrame,
                                    file_path: str,
                                    detection: Optional[DetectionResult],
                                    include_preliminary: bool) -> Tuple[List, List, Dict]:
        """Extrait données et en-têtes d'un DataFrame (logique commune CSV/TSV/Excel)"""
        # Cohérence avec la détection : élaguer d'éventuelles colonnes parasites.
        df, _ = prune_phantom_columns(df)

        # Déterminer les paramètres de chargement (override > détection > manuel)
        detection_info = self._resolve_structure(file_path, detection)
        header_start_row = detection_info['header_start_row']
        header_rows = detection_info['header_rows']
        data_start_row = detection_info['data_start_row']
        data_end_row = detection_info['data_end_row'] or None

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
            data = self._sort_data(data, self.options.sort_column, headers, result)

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

    def _sort_data(self, data: List[List], sort_column: int,
                   headers: List[List], result: CompilationResult) -> List[List]:
        """
        Trie les données par colonne, de façon robuste aux types mixtes.

        - Si la colonne demandée n'existe pas, on avertit l'utilisateur
          (warning visible) plutôt que d'ignorer silencieusement le tri.
        - La clé de tri est normalisée pour ne jamais comparer directement
          des types incompatibles (str vs nombre), ce qui levait un TypeError
          avalé silencieusement et laissait les données non triées.
        """
        if not data:
            return data

        # Vérifier que la colonne existe (au moins sur les en-têtes ou la 1re ligne)
        col_count = len(headers[-1]) if headers else len(data[0])
        if sort_column < 0 or sort_column >= col_count:
            msg = (f"Tri ignoré: colonne {sort_column + 1} hors limites "
                   f"(le tableau a {col_count} colonnes)")
            self.logger.warning(msg)
            result.warnings.append(msg)
            return data

        def sort_key(row):
            value = row[sort_column] if sort_column < len(row) else None
            # Cellule vide -> en fin de tri
            if value is None or (isinstance(value, float) and pd.isna(value)) or value == '':
                return (2, '')
            # Nombres triés numériquement, dans un groupe distinct des chaînes
            if isinstance(value, bool):
                return (1, str(value))
            if isinstance(value, (int, float)):
                return (0, float(value))
            # Tout le reste comparé en chaîne (casse-insensible)
            return (1, str(value).lower())

        try:
            return sorted(data, key=sort_key)
        except Exception as e:
            msg = f"Tri échoué, données laissées dans l'ordre d'origine: {e}"
            self.logger.warning(msg)
            result.warnings.append(msg)
            return data

    def _write_output_file(self, data: List[List], headers: List[List],
                          output_file: str, output_format: OutputFormat,
                          result: CompilationResult):
        """
        Écrit le fichier de sortie avec formatage professionnel

        Args:
            data: Données à écrire
            headers: En-têtes
            output_file: Chemin de sortie
            output_format: Format de sortie
            result: Résultat (pour statistiques)
        """
        self.logger.info(f"Écriture fichier de sortie: {output_file}")

        if output_format == OutputFormat.XLSX:
            # Utiliser openpyxl pour un formatage professionnel
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Compilation"

            # Écrire les en-têtes avec style
            current_row = ExcelFormatter.write_headers(ws, headers, start_row=1)

            # Écrire les données avec bordures
            date_format_str = self.options.date_format.value.upper()
            ExcelFormatter.write_data(ws, data, start_row=current_row,
                                     date_format=date_format_str)

            # Appliquer les options de formatage
            ExcelFormatter.adjust_column_widths(ws, min_width=10, max_width=50)

            # Figer les en-têtes (ligne après les en-têtes)
            freeze_row = len(headers) + 1
            ExcelFormatter.freeze_header(ws, freeze_row)

            # Sauvegarder
            wb.save(output_file)
            self.logger.info("Formatage Excel appliqué avec succès")

        elif output_format == OutputFormat.CSV:
            # Pour CSV, utiliser pandas (pas de formatage)
            all_data = headers + data
            df = pd.DataFrame(all_data)
            df.to_csv(output_file, index=False, header=False, encoding='utf-8-sig')

        elif output_format == OutputFormat.TSV:
            # Pour TSV, utiliser pandas (pas de formatage)
            all_data = headers + data
            df = pd.DataFrame(all_data)
            df.to_csv(output_file, index=False, header=False, sep='\t', encoding='utf-8-sig')

        result.output_file = output_file
        result.success = True

    def _detect_with_reference(self, file_paths: List[str],
                                result: CompilationResult) -> Dict[str, DetectionResult]:
        """
        Détecte la structure en utilisant un fichier de référence

        Args:
            file_paths: Liste des fichiers
            result: Résultat de compilation (pour warnings)

        Returns:
            Dict {file_path: DetectionResult}
        """
        detection_results = {}

        if not file_paths:
            return detection_results

        # Le premier fichier est la référence
        reference_file = file_paths[0]
        self.logger.info(f"Fichier de référence: {Path(reference_file).name}")

        try:
            # Créer le détecteur de référence
            self.reference_detector = ReferenceDetector(
                reference_file=reference_file,
                reference_header_row=self.options.reference_header_row,
                reference_header_lines=self.options.reference_header_lines,
                similarity_threshold=0.95,  # 95% pour acceptation automatique
                min_similarity=0.70,  # 70% minimum pour validation
                max_search_rows=50  # Chercher jusqu'à 50 lignes
            )

            # Détection sur tous les fichiers (y compris la référence)
            for file_path in file_paths:
                try:
                    detection = self.reference_detector.detect(file_path)
                    detection_results[file_path] = detection

                    # Log si avertissement
                    if detection.warning:
                        self.logger.warning(f"{Path(file_path).name}: {detection.warning}")
                        result.warnings.append(f"{Path(file_path).name}: {detection.warning}")

                    # Log info détection
                    if detection.debug_info and 'similarity_score' in detection.debug_info:
                        similarity = detection.debug_info['similarity_score']
                        self.logger.info(
                            f"{Path(file_path).name}: ligne {detection.header_start_row}, "
                            f"similarité {similarity:.0%}"
                        )

                except Exception as e:
                    self.logger.error(f"Erreur détection {Path(file_path).name}: {e}")
                    result.warnings.append(f"Erreur détection {Path(file_path).name}: {e}")

            # Marquer qu'on a utilisé la détection auto (mode référence)
            result.auto_detection_used = True
            result.detection_success_rate = len(detection_results) / len(file_paths) if file_paths else 0

        except Exception as e:
            self.logger.error(f"Erreur création détecteur référence: {e}", exc_info=True)
            result.warnings.append(f"Erreur mode référence: {e}")

        return detection_results
