"""
Tests de l'aperçu de détection (ExcelCompiler.preview_detection).

L'aperçu doit refléter fidèlement ce que produira la compilation, pour les
trois modes (automatique / référence / manuel), traiter chaque fichier
indépendamment et signaler proprement les erreurs.
"""

import openpyxl
from openpyxl.styles import Border, Side

import pytest

from core.compilation import ExcelCompiler, CompilationOptions, FilePreview
from core.compilation.compilation_models import FilenameOption


_THIN = Side(style="thin")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _make_bordered_xlsx(path, rows, border_from_row=1):
    wb = openpyxl.Workbook()
    ws = wb.active
    for ri, row in enumerate(rows, 1):
        for ci, val in enumerate(row, 1):
            cell = ws.cell(ri, ci, val)
            if ri >= border_from_row:
                cell.border = _BORDER
    wb.save(path)
    return str(path)


def _simple_file(tmp_path, name="f.xlsx"):
    return _make_bordered_xlsx(tmp_path / name, [
        ["Nom", "Age", "Ville"],
        ["Alice", 30, "Paris"],
        ["Bob", 25, "Lyon"],
        ["Carol", 40, "Nice"],
    ])


class TestPreviewDetection:
    def test_returns_one_preview_per_file(self, tmp_path):
        f1 = _simple_file(tmp_path, "a.xlsx")
        f2 = _simple_file(tmp_path, "b.xlsx")
        opts = CompilationOptions(auto_detect_structure=True, use_reference_mode=False)
        previews = ExcelCompiler(opts).preview_detection([f1, f2])
        assert len(previews) == 2
        assert all(isinstance(p, FilePreview) for p in previews)

    def test_auto_mode_detects_header_and_data(self, tmp_path):
        f = _simple_file(tmp_path)
        opts = CompilationOptions(auto_detect_structure=True, use_reference_mode=False)
        p = ExcelCompiler(opts).preview_detection([f])[0]
        assert p.success
        assert p.header_start_row == 1
        assert p.detected_headers[:3] == ["Nom", "Age", "Ville"]
        assert p.data_row_count == 3

    def test_reference_mode_preview(self, tmp_path):
        f = _simple_file(tmp_path)
        opts = CompilationOptions(
            use_reference_mode=True, auto_detect_structure=False,
            reference_header_row=1, reference_header_lines=1,
            filename_option=FilenameOption.NONE,
        )
        p = ExcelCompiler(opts).preview_detection([f])[0]
        assert p.success
        assert p.detection_method == "reference"
        assert p.header_start_row == 1

    def test_manual_mode_uses_manual_params(self, tmp_path):
        f = _simple_file(tmp_path)
        # en-tête forcé en ligne 1 manuellement
        opts = CompilationOptions(
            use_reference_mode=False, auto_detect_structure=False,
            manual_header_start_row=1, manual_header_rows=1,
        )
        p = ExcelCompiler(opts).preview_detection([f])[0]
        assert p.success
        assert p.detection_method == "manual"
        assert p.header_start_row == 1
        assert p.detected_headers[:3] == ["Nom", "Age", "Ville"]

    def test_missing_file_is_reported_not_raised(self, tmp_path):
        good = _simple_file(tmp_path)
        opts = CompilationOptions(auto_detect_structure=True, use_reference_mode=False)
        previews = ExcelCompiler(opts).preview_detection([good, str(tmp_path / "nope.xlsx")])
        assert previews[0].success
        assert previews[1].success is False
        assert previews[1].error

    def test_preview_matches_compilation_data_count(self, tmp_path):
        """Le nombre de lignes annoncé par l'aperçu doit correspondre à la
        compilation réelle (fidélité)."""
        f = _simple_file(tmp_path)
        opts = CompilationOptions(
            auto_detect_structure=True, use_reference_mode=False,
            filename_option=FilenameOption.NONE, remove_empty_rows=True,
        )
        comp = ExcelCompiler(opts)
        preview = comp.preview_detection([f])[0]

        from core.compilation import OutputFormat
        out = tmp_path / "out.xlsx"
        result = ExcelCompiler(opts).compile_files([f], str(out), OutputFormat.XLSX)
        assert preview.data_row_count == result.total_rows

    def test_filename_property(self, tmp_path):
        f = _simple_file(tmp_path, "monfichier.xlsx")
        opts = CompilationOptions(auto_detect_structure=True, use_reference_mode=False)
        p = ExcelCompiler(opts).preview_detection([f])[0]
        assert p.filename == "monfichier.xlsx"
