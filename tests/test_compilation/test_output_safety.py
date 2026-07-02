"""
Tests des garde-fous protégeant les données utilisateur (Vague 3, F1/F4).

- F1 : l'extension du fichier de sortie est synchronisée avec le format choisi
  (pas de contenu CSV dans un .xlsx).
- F4 : un classeur multi-feuilles est SIGNALÉ (seule la première est compilée),
  sinon les données des autres feuilles seraient oubliées silencieusement.

F2 (confirmation d'écrasement) et F3 (exclusion du fichier de sortie de la
sélection) vivent dans l'orchestration UI (QMessageBox, sélection de fichiers)
et sont vérifiés manuellement ; la logique testable de F3 (comparaison de
chemins) est triviale et couverte par l'usage.
"""

import openpyxl
import pytest

from core.compilation import ExcelCompiler, CompilationOptions, OutputFormat
from core.compilation.compilation_models import FilenameOption
from ui.main_window import MainWindow


def _comp():
    return ExcelCompiler(CompilationOptions(
        use_reference_mode=True, reference_header_row=1, reference_header_lines=1,
        filename_option=FilenameOption.NONE,
    ))


# --------------------------------------------------------------------------- #
# F1 — extension synchronisée au format
# --------------------------------------------------------------------------- #
class TestOutputExtension:
    @pytest.mark.parametrize("name,fmt,expected", [
        ("compilation.xlsx", OutputFormat.CSV, "compilation.csv"),
        ("sortie.csv", OutputFormat.XLSX, "sortie.xlsx"),
        ("data", OutputFormat.TSV, "data.tsv"),
        ("compilation.xlsx", OutputFormat.XLSX, "compilation.xlsx"),
        ("rapport.TSV", OutputFormat.CSV, "rapport.csv"),
    ])
    def test_extension_matches_format(self, name, fmt, expected):
        assert MainWindow._ensure_output_extension(name, fmt) == expected

    def test_empty_name_gets_default_stem(self):
        # Un nom réduit à une extension -> stem par défaut.
        assert MainWindow._ensure_output_extension(".csv", OutputFormat.XLSX) \
            == "compilation.xlsx"


# --------------------------------------------------------------------------- #
# F4 — avertissement multi-feuilles
# --------------------------------------------------------------------------- #
class TestMultiSheetWarning:
    def _make_multisheet(self, path, sheet_names):
        wb = openpyxl.Workbook()
        wb.active.title = sheet_names[0]
        wb.active.append(['Nom', 'Val'])
        wb.active.append(['a', 1])
        for name in sheet_names[1:]:
            ws = wb.create_sheet(name)
            ws.append(['Nom', 'Val'])
            ws.append(['b', 2])
        wb.save(path)
        return str(path)

    def test_preview_warns_multisheet(self, tmp_path):
        p = self._make_multisheet(tmp_path / 'm.xlsx', ['Janvier', 'Février'])
        preview = _comp().preview_detection([p])[0]
        assert preview.warning is not None
        assert 'feuilles' in preview.warning
        assert 'Janvier' in preview.warning  # nom de la feuille compilée

    def test_compilation_warns_multisheet(self, tmp_path):
        p = self._make_multisheet(tmp_path / 'm.xlsx', ['S1', 'S2', 'S3'])
        out = tmp_path / 'out.xlsx'
        result = _comp().compile_files([p], str(out), OutputFormat.XLSX)
        assert any('feuilles' in w for w in result.warnings)
        assert any('3 feuilles' in w for w in result.warnings)

    def test_single_sheet_no_warning(self, tmp_path):
        p = self._make_multisheet(tmp_path / 's.xlsx', ['Seule'])
        preview = _comp().preview_detection([p])[0]
        assert not (preview.warning and 'feuilles' in preview.warning)

    def test_sheet_names_helper(self, tmp_path):
        p = self._make_multisheet(tmp_path / 'm.xlsx', ['A', 'B'])
        names = ExcelCompiler._sheet_names(p)
        assert names == ['A', 'B']

    def test_sheet_names_on_bad_file_returns_empty(self, tmp_path):
        bad = tmp_path / 'bad.xlsx'
        bad.write_bytes(b'not a real xlsx')
        assert ExcelCompiler._sheet_names(str(bad)) == []
