"""
Structures de données pour la compilation
Auteur: GOUNOU N'GOBI Chabi Zimé
Version: 3.2
"""

from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any, Tuple
from pathlib import Path
from enum import Enum


class FilenameOption(Enum):
    """Options pour l'ajout du nom de fichier source"""
    NONE = "none"
    WITH_EXTENSION = "with_extension"
    WITHOUT_EXTENSION = "without_extension"


class DateFormat(Enum):
    """Formats de date supportés"""
    FRENCH = "french"  # JJ/MM/AAAA
    AMERICAN = "american"  # MM/DD/YYYY
    ISO = "iso"  # YYYY-MM-DD
    DATETIME_FRENCH = "datetime_french"  # JJ/MM/AAAA HH:MM:SS


class OutputFormat(Enum):
    """Formats de sortie supportés"""
    XLSX = "xlsx"
    CSV = "csv"
    TSV = "tsv"


@dataclass
class CompilationOptions:
    """
    Options de compilation pour personnaliser le traitement

    Contient toutes les options qui contrôlent le comportement
    de la compilation (tri, doublons, en-têtes, etc.)
    """

    # Options de base
    filename_option: FilenameOption = FilenameOption.NONE
    sort_data: bool = False
    sort_column: int = 0
    repeat_headers: bool = False
    remove_empty_rows: bool = True
    remove_duplicates: bool = False
    date_format: DateFormat = DateFormat.FRENCH

    # Options préliminaires (informations avant tableau)
    include_preliminary: bool = False
    preliminary_source_file: Optional[str] = None

    # Options de détection (v3.2)
    auto_detect_structure: bool = True
    enable_cross_validation: bool = True
    detection_confidence_threshold: float = 0.65

    # Options mode référence (v3.2.1 - nouveau)
    use_reference_mode: bool = False
    reference_header_row: int = 1
    reference_header_lines: int = 1

    # Options manuelles (fallback si détection échoue)
    manual_header_start_row: int = 1
    manual_header_rows: int = 1

    # Corrections manuelles par fichier (v3.2) — priorité absolue sur la
    # détection et sur le manuel global. Permet à l'utilisateur de fixer
    # lui-même l'en-tête d'un fichier mal identifié, fichier par fichier.
    # {chemin_fichier: (header_start_row, header_rows)} (lignes 1-based)
    manual_overrides: Dict[str, Tuple[int, int]] = field(default_factory=dict)

    # Dé-fusion des cellules fusionnées (v3.2) — propage la valeur d'une
    # cellule fusionnée sur toute sa plage (ex. un département fusionné
    # verticalement est recopié sur chaque ligne). Traitement déterministe ;
    # désactivable si l'utilisateur préfère le comportement brut de pandas.
    unmerge_cells: bool = True

    # Aplatissement des en-têtes multi-lignes (v3.2) — quand l'en-tête s'étend
    # sur plusieurs lignes (catégorie + sous-catégorie, ex. « NOMBRE DE CAS »
    # au-dessus de « 1er cycle » / « 2nd cycle »), les fusionne en un seul
    # libellé par colonne (« NOMBRE DE CAS - 1er cycle »). Produit une unique
    # ligne d'en-tête propre en sortie. Désactivable pour garder les lignes
    # d'en-tête brutes telles quelles.
    flatten_multiindex_headers: bool = True

    # Lignes de sous-total / total (v3.2) — les tableaux intercalent souvent des
    # lignes d'agrégat (« ENSEMBLE COMMUNE », « ENSEMBLE DEPARTEMENT »…). Par
    # défaut elles sont EXCLUES (évite le double comptage) et leur nombre est
    # signalé. Si drop_subtotal_rows=False, elles sont conservées et, si
    # mark_subtotal_rows est actif, une colonne « Type de ligne » est ajoutée
    # (détail / sous-total / total) pour permettre un filtrage propre.
    drop_subtotal_rows: bool = True
    mark_subtotal_rows: bool = True
    subtotal_row_label: str = "Type de ligne"
    # Mots-clés configurables (None = listes par défaut du subtotal_detector).
    subtotal_keywords: Optional[List[str]] = None
    total_keywords: Optional[List[str]] = None

    # Blocs de signature en pied de tableau (v3.2) — un formulaire rempli se
    # termine souvent par « Fait à X, le … » puis la qualité et le nom du
    # signataire. Ces lignes sont SOUS le tableau : compilées, elles produisent
    # des lignes sans identité dont le texte tombe dans des colonnes de
    # chiffres. Exclues par défaut ; la détection exige cumulativement une
    # identité vide, aucune valeur numérique et un marqueur de signature
    # (voir subtotal_detector.is_signature_row).
    drop_signature_rows: bool = True

    # Alignement des colonnes par libellé (v3.2) — quand les fichiers n'ont pas
    # exactement les mêmes colonnes (ordre différent, colonne en plus / en
    # moins), aligne chaque colonne sur son LIBELLÉ normalisé plutôt que sur sa
    # position. Les colonnes communes sont empilées correctement ; les colonnes
    # absentes d'un fichier deviennent des cellules vides (jamais inventées) ;
    # les colonnes inconnues sont ajoutées au schéma global et signalées. Évite
    # la corruption silencieuse de l'empilement positionnel. Désactivable pour
    # retrouver le comportement positionnel (empilement par position + padding).
    align_columns_by_label: bool = True

    # Alias de colonnes (v3.2, étape 6) — mapping manuel résolvant les cas que
    # l'alignement par libellé n'ose pas deviner. Quand une colonne d'un fichier
    # porte un libellé différent de celui du schéma (ex. « Sexe (M/F) » vs
    # « Sexe »), l'utilisateur la rattache explicitement : la clé est le libellé
    # SOURCE, la valeur le libellé CIBLE du schéma global. Comparaison sans
    # accents ni casse (comme l'alignement). Vide par défaut = aucun alias.
    column_aliases: Dict[str, str] = field(default_factory=dict)

    # Garde-fous de sécurité (v3.2) — réellement appliqués au chargement.
    # max_file_size_mb : un fichier plus volumineux est rejeté proprement
    #   (le fichier est marqué en échec, les autres continuent). 0 = sans limite.
    # max_rows_per_file : borne le nombre de lignes lues d'un fichier pour
    #   neutraliser les .xlsx aux dimensions gonflées (anti-explosion mémoire).
    #   0 = sans limite. Voir aussi le plafond colonnes du merge_handler.
    max_file_size_mb: float = 100.0
    max_rows_per_file: int = 1_000_000

    def to_dict(self) -> Dict[str, Any]:
        """Convertit en dictionnaire"""
        return {
            'filename_option': self.filename_option.value,
            'sort_data': self.sort_data,
            'sort_column': self.sort_column,
            'repeat_headers': self.repeat_headers,
            'remove_empty_rows': self.remove_empty_rows,
            'remove_duplicates': self.remove_duplicates,
            'date_format': self.date_format.value,
            'include_preliminary': self.include_preliminary,
            'preliminary_source_file': self.preliminary_source_file,
            'auto_detect_structure': self.auto_detect_structure,
            'enable_cross_validation': self.enable_cross_validation,
            'detection_confidence_threshold': self.detection_confidence_threshold,
            'use_reference_mode': self.use_reference_mode,
            'reference_header_row': self.reference_header_row,
            'reference_header_lines': self.reference_header_lines,
            'manual_header_start_row': self.manual_header_start_row,
            'manual_header_rows': self.manual_header_rows,
            'manual_overrides': {k: list(v) for k, v in self.manual_overrides.items()},
            'unmerge_cells': self.unmerge_cells,
            'flatten_multiindex_headers': self.flatten_multiindex_headers,
            'drop_subtotal_rows': self.drop_subtotal_rows,
            'mark_subtotal_rows': self.mark_subtotal_rows,
            'subtotal_row_label': self.subtotal_row_label,
            'subtotal_keywords': self.subtotal_keywords,
            'total_keywords': self.total_keywords,
            'drop_signature_rows': self.drop_signature_rows,
            'align_columns_by_label': self.align_columns_by_label,
            'column_aliases': dict(self.column_aliases),
            'max_file_size_mb': self.max_file_size_mb,
            'max_rows_per_file': self.max_rows_per_file
        }


@dataclass
class FilePreview:
    """
    Aperçu de la détection pour un fichier, avant compilation.

    Permet à l'utilisateur de vérifier ce que le système a détecté
    (ligne d'en-tête, en-têtes, étendue des données) sans lancer la
    compilation complète.
    """

    file_path: str
    success: bool = True
    header_start_row: int = 1
    header_rows: int = 1
    data_start_row: int = 2
    data_end_row: int = 0
    detected_headers: List[str] = field(default_factory=list)
    confidence: float = 0.0
    detection_method: str = "manual"
    data_row_count: int = 0
    warning: Optional[str] = None
    error: Optional[str] = None

    @property
    def filename(self) -> str:
        return Path(self.file_path).name


@dataclass
class FileCompilationResult:
    """
    Résultat de compilation pour un fichier individuel

    Contient les informations sur le traitement d'un fichier
    et les détails de détection utilisés
    """

    file_path: str
    success: bool = False
    rows_added: int = 0
    error_message: Optional[str] = None

    # Informations de détection
    detection_used: bool = False
    header_start_row: int = 1
    header_rows: int = 1
    data_start_row: int = 2
    data_end_row: int = 0
    detection_confidence: float = 0.0
    detection_method: str = "manual"

    # Nombre de lignes de total / sous-total détectées dans ce fichier
    # (exclues si drop_subtotal_rows, sinon conservées et éventuellement marquées).
    subtotal_rows: int = 0

    # Nombre de lignes du bloc de signature en pied de tableau écartées
    # (« Fait à …, le … », « Le Directeur », nom du signataire).
    signature_rows: int = 0

    # Confiance d'une détection ÉCARTÉE (sous le seuil) quand on est retombé
    # sur les paramètres manuels. 0.0 = aucune détection écartée. Sert à
    # signaler ce rejet à l'utilisateur au lieu d'un basculement silencieux.
    rejected_confidence: float = 0.0

    # Métriques
    processing_time: float = 0.0
    memory_used: int = 0

    def __str__(self) -> str:
        if self.success:
            return (f"{Path(self.file_path).name}: {self.rows_added} lignes "
                   f"(détection: {self.detection_method}, confiance: {self.detection_confidence:.0%})")
        else:
            return f"{Path(self.file_path).name}: ERREUR - {self.error_message}"


@dataclass
class CompilationResult:
    """
    Résultat global de la compilation

    Contient toutes les informations sur la compilation complète,
    y compris les statistiques et les fichiers traités
    """

    # Statut
    success: bool = False
    cancelled: bool = False
    output_file: Optional[str] = None

    # Statistiques globales
    total_files: int = 0
    successful_files: int = 0
    failed_files: int = 0
    total_rows: int = 0
    total_processing_time: float = 0.0

    # Détails par fichier
    file_results: List[FileCompilationResult] = field(default_factory=list)
    failed_file_details: List[tuple] = field(default_factory=list)  # [(filename, error)]

    # Informations de détection (v3.2)
    auto_detection_used: bool = False
    detection_success_rate: float = 0.0
    average_detection_confidence: float = 0.0

    # En-têtes finaux
    final_headers: List[str] = field(default_factory=list)

    # Métriques mémoire
    peak_memory_mb: float = 0.0

    # Avertissements
    warnings: List[str] = field(default_factory=list)

    def add_file_result(self, file_result: FileCompilationResult):
        """Ajoute le résultat d'un fichier"""
        self.file_results.append(file_result)

        if file_result.success:
            self.successful_files += 1
            self.total_rows += file_result.rows_added
        else:
            self.failed_files += 1
            self.failed_file_details.append(
                (file_result.file_path, file_result.error_message)
            )

    def calculate_statistics(self):
        """Calcule les statistiques finales"""
        self.total_files = len(self.file_results)

        if self.file_results:
            self.total_processing_time = sum(
                r.processing_time for r in self.file_results
            )

            # Statistiques de détection
            detection_results = [
                r for r in self.file_results
                if r.detection_used and r.success
            ]

            if detection_results:
                self.auto_detection_used = True
                self.detection_success_rate = len(detection_results) / self.total_files
                self.average_detection_confidence = sum(
                    r.detection_confidence for r in detection_results
                ) / len(detection_results)

        # Déterminer le succès global
        self.success = self.successful_files > 0

    def get_summary(self) -> str:
        """Retourne un résumé lisible"""
        summary = [
            "=" * 60,
            "RESULTAT DE COMPILATION",
            "=" * 60,
            f"Statut: {'SUCCES' if self.success else 'ECHEC'}",
            f"Fichier de sortie: {self.output_file or 'N/A'}",
            "",
            "Statistiques:",
            f"  - Fichiers traités: {self.successful_files}/{self.total_files}",
            f"  - Lignes compilées: {self.total_rows}",
            f"  - Temps total: {self.total_processing_time:.2f}s",
            f"  - Mémoire pic: {self.peak_memory_mb:.1f} MB",
        ]

        if self.auto_detection_used:
            summary.extend([
                "",
                "Détection automatique:",
                f"  - Taux de réussite: {self.detection_success_rate:.0%}",
                f"  - Confiance moyenne: {self.average_detection_confidence:.0%}",
            ])

        if self.failed_files > 0:
            summary.extend([
                "",
                f"Fichiers échoués ({self.failed_files}):"
            ])
            for filename, error in self.failed_file_details[:5]:  # Max 5
                summary.append(f"  - {Path(filename).name}: {error}")
            if len(self.failed_file_details) > 5:
                summary.append(f"  ... et {len(self.failed_file_details) - 5} autres")

        if self.warnings:
            summary.extend([
                "",
                f"Avertissements ({len(self.warnings)}):"
            ])
            for warning in self.warnings[:3]:  # Max 3
                summary.append(f"  - {warning}")
            if len(self.warnings) > 3:
                summary.append(f"  ... et {len(self.warnings) - 3} autres")

        summary.append("=" * 60)

        return "\n".join(summary)

    def to_dict(self) -> Dict[str, Any]:
        """Convertit en dictionnaire pour sérialisation"""
        return {
            'success': self.success,
            'output_file': self.output_file,
            'total_files': self.total_files,
            'successful_files': self.successful_files,
            'failed_files': self.failed_files,
            'total_rows': self.total_rows,
            'total_processing_time': self.total_processing_time,
            'auto_detection_used': self.auto_detection_used,
            'detection_success_rate': self.detection_success_rate,
            'average_detection_confidence': self.average_detection_confidence,
            'final_headers': self.final_headers,
            'peak_memory_mb': self.peak_memory_mb,
            'warnings': self.warnings,
            'failed_files': [
                {'file': f, 'error': e} for f, e in self.failed_file_details
            ]
        }
