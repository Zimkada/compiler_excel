"""
Tests des corrections manuelles par fichier (CompilationOptions.manual_overrides).

L'override d'un fichier doit avoir la PRIORITÉ ABSOLUE sur la détection
automatique et sur le manuel global, à la fois dans l'aperçu et dans la
compilation réelle (zéro divergence).
"""

import openpyxl

from core.compilation import ExcelCompiler, CompilationOptions, OutputFormat
from core.compilation.compilation_models import FilenameOption


def _make_xlsx(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for ri, row in enumerate(rows, 1):
        for ci, val in enumerate(row, 1):
            ws.cell(ri, ci, val)
    wb.save(path)
    return str(path)


def _file_with_title_rows(tmp_path, name="titre.xlsx"):
    """Fichier dont le vrai en-tête est en ligne 3 (2 lignes de titre avant)."""
    return _make_xlsx(tmp_path / name, [
        ["RAPPORT MENSUEL", None, None],   # ligne 1 : titre
        ["Service Grossesses", None, None],  # ligne 2 : sous-titre
        ["Nom", "Age", "Ville"],            # ligne 3 : VRAI en-tête
        ["Alice", 30, "Paris"],
        ["Bob", 25, "Lyon"],
        ["Carol", 40, "Nice"],
    ])


class TestManualOverridesPreview:
    def test_override_takes_priority_in_preview(self, tmp_path):
        f = _file_with_title_rows(tmp_path)
        opts = CompilationOptions(
            auto_detect_structure=True, use_reference_mode=False,
            manual_overrides={f: (3, 1)},  # en-tête forcé en ligne 3
        )
        p = ExcelCompiler(opts).preview_detection([f])[0]
        assert p.success
        assert p.detection_method == "override"
        assert p.header_start_row == 3
        assert p.detected_headers[:3] == ["Nom", "Age", "Ville"]
        assert p.data_row_count == 3  # Alice, Bob, Carol

    def test_override_beats_low_confidence_and_manual(self, tmp_path):
        f = _file_with_title_rows(tmp_path)
        # Manuel global pointe (à tort) sur la ligne 1 ; l'override doit gagner.
        opts = CompilationOptions(
            auto_detect_structure=False, use_reference_mode=False,
            manual_header_start_row=1, manual_header_rows=1,
            manual_overrides={f: (3, 1)},
        )
        p = ExcelCompiler(opts).preview_detection([f])[0]
        assert p.header_start_row == 3
        assert p.detected_headers[:3] == ["Nom", "Age", "Ville"]

    def test_override_multiline_header(self, tmp_path):
        f = _file_with_title_rows(tmp_path)
        # En-tête sur 2 lignes (lignes 2 et 3 fusionnées en libellés)
        opts = CompilationOptions(
            auto_detect_structure=True, use_reference_mode=False,
            manual_overrides={f: (2, 2)},
        )
        p = ExcelCompiler(opts).preview_detection([f])[0]
        assert p.header_start_row == 2
        assert p.header_rows == 2
        assert p.data_start_row == 4  # données commencent après les 2 lignes
        assert p.data_row_count == 3


class TestManualOverridesCompilation:
    def test_override_applied_in_compilation(self, tmp_path):
        f = _file_with_title_rows(tmp_path)
        opts = CompilationOptions(
            auto_detect_structure=True, use_reference_mode=False,
            filename_option=FilenameOption.NONE,
            manual_overrides={f: (3, 1)},
        )
        out = tmp_path / "out.xlsx"
        result = ExcelCompiler(opts).compile_files([f], str(out), OutputFormat.XLSX)
        assert result.success
        assert result.total_rows == 3  # 3 lignes de données, pas les titres
        fr = result.file_results[0]
        assert fr.header_start_row == 3
        assert fr.detection_method == "override"

    def test_preview_matches_compilation_with_override(self, tmp_path):
        """Fidélité : l'aperçu annonce le même nombre de lignes que la
        compilation réelle, override appliqué."""
        f = _file_with_title_rows(tmp_path)
        opts = CompilationOptions(
            auto_detect_structure=True, use_reference_mode=False,
            filename_option=FilenameOption.NONE, remove_empty_rows=True,
            manual_overrides={f: (3, 1)},
        )
        preview = ExcelCompiler(opts).preview_detection([f])[0]
        result = ExcelCompiler(opts).compile_files(
            [f], str(tmp_path / "out.xlsx"), OutputFormat.XLSX
        )
        assert preview.data_row_count == result.total_rows

    def test_override_only_affects_targeted_file(self, tmp_path):
        """Un override sur un fichier ne doit pas affecter les autres."""
        f1 = _file_with_title_rows(tmp_path, "avec_titre.xlsx")
        f2 = _make_xlsx(tmp_path / "propre.xlsx", [
            ["Nom", "Age", "Ville"],
            ["Dan", 22, "Brest"],
            ["Eve", 28, "Tours"],
        ])
        opts = CompilationOptions(
            auto_detect_structure=True, use_reference_mode=False,
            filename_option=FilenameOption.NONE,
            manual_overrides={f1: (3, 1)},  # override sur f1 seulement
        )
        previews = ExcelCompiler(opts).preview_detection([f1, f2])
        p1 = next(p for p in previews if p.filename == "avec_titre.xlsx")
        p2 = next(p for p in previews if p.filename == "propre.xlsx")
        assert p1.detection_method == "override"
        assert p1.header_start_row == 3
        assert p2.detection_method != "override"  # f2 garde la détection auto
