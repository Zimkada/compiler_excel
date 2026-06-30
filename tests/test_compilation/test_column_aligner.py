"""
Tests adverses de l'alignement des colonnes par libellé (défaut D5).

Quand les fichiers n'ont pas les mêmes colonnes (ordre différent, colonne en
plus / en moins), l'empilement positionnel corrompt silencieusement les données.
L'aligneur projette chaque fichier sur un schéma global identifié par LIBELLÉ :
colonnes communes au bon endroit, colonnes absentes -> vides (jamais inventées),
colonnes inconnues -> ajoutées à droite et signalées.
"""

import openpyxl
import pandas as pd

from core.compilation import ExcelCompiler, CompilationOptions
from core.compilation.column_aligner import (
    ColumnAligner, normalize_label, _column_keys,
)


def _make_xlsx(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for ri, row in enumerate(rows, 1):
        for ci, val in enumerate(row, 1):
            if val is not None:
                ws.cell(ri, ci, val)
    wb.save(path)
    return str(path)


class TestNormalizeAndKeys:
    def test_normalize_accents_case_spaces(self):
        assert normalize_label("  Âge  ") == "age"
        assert normalize_label("RÉGION") == "region"
        assert normalize_label(None) == ""

    def test_duplicate_labels_disambiguated(self):
        # Deux colonnes au même libellé -> appariées par occurrence, pas fusionnées.
        assert _column_keys(["Effectif", "Effectif"]) == ["effectif#0", "effectif#1"]

    def test_empty_labels_not_merged(self):
        # Deux colonnes sans libellé restent distinctes (identité positionnelle).
        keys = _column_keys(["A", None, ""])
        assert keys[0] == "a#0"
        assert keys[1] != keys[2]


class TestAlignerUnit:
    def test_reorders_by_label(self):
        a = ColumnAligner()
        a.set_reference(["Region", "Sexe", "Cas"])
        rows, new = a.align(["Sexe", "Region"], [["M", "Nord"], ["F", "Sud"]])
        assert new == []
        # Remis dans l'ordre du schéma ; « Cas » absent du fichier -> None.
        assert rows == [["Nord", "M", None], ["Sud", "F", None]]

    def test_unknown_column_appended_and_reported(self):
        a = ColumnAligner()
        a.set_reference(["Region", "Cas"])
        rows, new = a.align(["Region", "Cas", "Annee"], [["Est", 10, 2024]])
        assert new == ["Annee"]
        assert a.schema_labels() == ["Region", "Cas", "Annee"]
        assert rows == [["Est", 10, 2024]]

    def test_accent_case_insensitive_match(self):
        a = ColumnAligner()
        a.set_reference(["Région", "Cas"])
        rows, new = a.align(["region", "CAS"], [["Nord", 5]])
        assert new == []          # « region » == « Région »
        assert rows == [["Nord", 5]]

    def test_row_wider_than_header_extra_cell_ignored(self):
        # Une ligne plus large que son en-tête : la cellule sans colonne
        # identifiable est ignorée (pas de colonne fantôme inventée).
        a = ColumnAligner()
        a.set_reference(["A", "B"])
        rows, _ = a.align(["A", "B"], [[1, 2, 3]])
        assert rows == [[1, 2]]

    def test_never_invents_value(self):
        a = ColumnAligner()
        a.set_reference(["A", "B", "C"])
        rows, _ = a.align(["A"], [["x"]])
        assert rows == [["x", None, None]]   # B et C absents -> vides


class TestEndToEnd:
    def _compile(self, tmp_path, files, **opt):
        out = str(tmp_path / "out.xlsx")
        opts = CompilationOptions(
            auto_detect_structure=False,
            manual_header_start_row=1, manual_header_rows=1, **opt,
        )
        res = ExcelCompiler(opts).compile_files(files, out)
        return pd.read_excel(out, header=None), res

    def test_reordered_columns_aligned(self, tmp_path):
        f1 = _make_xlsx(tmp_path / "f1.xlsx", [["Region", "Sexe", "Cas"], ["Nord", "M", 10]])
        f2 = _make_xlsx(tmp_path / "f2.xlsx", [["Sexe", "Cas", "Region"], ["F", 20, "Sud"]])
        df, res = self._compile(tmp_path, [f1, f2])
        assert list(df.iloc[0]) == ["Region", "Sexe", "Cas"]
        assert list(df.iloc[1]) == ["Nord", "M", 10]
        # La 2e ligne est remise dans l'ordre du schéma, PAS empilée telle quelle.
        assert list(df.iloc[2]) == ["Sud", "F", 20]

    def test_positional_mode_preserved_when_disabled(self, tmp_path):
        f1 = _make_xlsx(tmp_path / "f1.xlsx", [["Region", "Sexe", "Cas"], ["Nord", "M", 10]])
        f2 = _make_xlsx(tmp_path / "f2.xlsx", [["Sexe", "Cas", "Region"], ["F", 20, "Sud"]])
        df, res = self._compile(tmp_path, [f1, f2], align_columns_by_label=False)
        # Comportement historique : empilement positionnel (sans réordonnancement).
        assert list(df.iloc[2]) == ["F", 20, "Sud"]

    def test_new_column_in_later_file_padded_and_warned(self, tmp_path):
        f1 = _make_xlsx(tmp_path / "f1.xlsx", [["Region", "Cas"], ["Nord", 10]])
        f2 = _make_xlsx(tmp_path / "f2.xlsx", [["Region", "Cas", "Annee"], ["Sud", 20, 2024]])
        df, res = self._compile(tmp_path, [f1, f2])
        assert list(df.iloc[0]) == ["Region", "Cas", "Annee"]
        # F1 (vu avant l'apparition d'« Annee ») est rembourré à droite.
        assert list(df.iloc[1]) == ["Nord", 10, None] or pd.isna(df.iloc[1, 2])
        assert list(df.iloc[2])[:2] == ["Sud", 20]
        assert any("Annee" in w for w in res.warnings)

    def test_missing_column_becomes_empty(self, tmp_path):
        f1 = _make_xlsx(tmp_path / "f1.xlsx", [["Region", "Sexe", "Cas"], ["Nord", "M", 10]])
        f2 = _make_xlsx(tmp_path / "f2.xlsx", [["Region", "Cas"], ["Sud", 20]])
        df, res = self._compile(tmp_path, [f1, f2])
        # « Sexe » manque dans F2 -> cellule vide, jamais inventée.
        assert df.iloc[2, 0] == "Sud"
        assert pd.isna(df.iloc[2, 1])
        assert df.iloc[2, 2] == 20

    def test_row_count_unchanged_by_alignment(self, tmp_path):
        # L'alignement ne change QUE les colonnes : le nombre de lignes par
        # fichier (donc l'invariant aperçu = sortie) reste intact.
        f1 = _make_xlsx(tmp_path / "f1.xlsx", [["A", "B"], [1, 2], [3, 4]])
        f2 = _make_xlsx(tmp_path / "f2.xlsx", [["B", "A"], [5, 6]])
        comp = ExcelCompiler(CompilationOptions(
            auto_detect_structure=False, manual_header_start_row=1, manual_header_rows=1))
        previews = comp.preview_detection([f1, f2])
        out = str(tmp_path / "out.xlsx")
        res = comp.compile_files([f1, f2], out)
        for prev, fr in zip(previews, res.file_results):
            assert prev.data_row_count == fr.rows_added
