"""
Tests du rattachement manuel des colonnes (éditeur 2D — étape 6).

L'alignement par libellé (étape 5) n'ose pas deviner qu'une colonne
« Sexe (M/F) » est la colonne « Sexe » du schéma. Un alias manuel
({source: cible}) le permet : la colonne source fusionne alors avec la cible.
Sans alias, le comportement de l'étape 5 est strictement inchangé.
"""

import openpyxl
import pandas as pd

from core.compilation import ExcelCompiler, CompilationOptions
from core.compilation.column_aligner import ColumnAligner


def _make_xlsx(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for ri, row in enumerate(rows, 1):
        for ci, val in enumerate(row, 1):
            if val is not None:
                ws.cell(ri, ci, val)
    wb.save(path)
    return str(path)


class TestAliasUnit:
    def test_alias_merges_into_target(self):
        a = ColumnAligner({"Sexe (M/F)": "Sexe"})
        a.set_reference(["Region", "Sexe"])
        rows, new = a.align(["Region", "Sexe (M/F)"], [["Nord", "M"]])
        assert new == []                      # aucune colonne nouvelle
        assert a.schema_labels() == ["Region", "Sexe"]
        assert rows == [["Nord", "M"]]

    def test_no_alias_keeps_step5_behaviour(self):
        a = ColumnAligner()
        a.set_reference(["Region", "Sexe"])
        _, new = a.align(["Region", "Sexe (M/F)"], [["Nord", "M"]])
        assert new == ["Sexe (M/F)"]          # colonne inconnue, comme étape 5
        assert a.schema_labels() == ["Region", "Sexe", "Sexe (M/F)"]

    def test_alias_case_accent_insensitive(self):
        a = ColumnAligner({"SEXE (m/f)": "sexe"})
        a.set_reference(["Region", "Sexe"])
        _, new = a.align(["Region", "Sexe (M/F)"], [["N", "M"]])
        assert new == []

    def test_alias_never_applied_to_reference(self):
        # L'alias définit des cibles : il ne doit pas renommer la référence.
        a = ColumnAligner({"Region": "Zone"})
        a.set_reference(["Region", "Sexe"])
        assert a.schema_labels() == ["Region", "Sexe"]


class TestAliasEndToEnd:
    def _compile(self, tmp_path, files, **opt):
        out = str(tmp_path / "out.xlsx")
        opts = CompilationOptions(
            auto_detect_structure=False,
            manual_header_start_row=1, manual_header_rows=1, **opt,
        )
        res = ExcelCompiler(opts).compile_files(files, out)
        return pd.read_excel(out, header=None), res

    def test_alias_aligns_renamed_column(self, tmp_path):
        f1 = _make_xlsx(tmp_path / "f1.xlsx", [["Region", "Sexe"], ["Nord", "M"]])
        f2 = _make_xlsx(tmp_path / "f2.xlsx", [["Region", "Sexe (M/F)"], ["Sud", "F"]])
        df, res = self._compile(tmp_path, [f1, f2],
                                column_aliases={"Sexe (M/F)": "Sexe"})
        # Deux colonnes seulement : la colonne renommée a fusionné dans « Sexe ».
        assert list(df.iloc[0]) == ["Region", "Sexe"]
        assert list(df.iloc[1]) == ["Nord", "M"]
        assert list(df.iloc[2]) == ["Sud", "F"]

    def test_without_alias_creates_separate_column(self, tmp_path):
        f1 = _make_xlsx(tmp_path / "f1.xlsx", [["Region", "Sexe"], ["Nord", "M"]])
        f2 = _make_xlsx(tmp_path / "f2.xlsx", [["Region", "Sexe (M/F)"], ["Sud", "F"]])
        df, res = self._compile(tmp_path, [f1, f2])
        # Sans alias : 3 colonnes, F dans la nouvelle « Sexe (M/F) », signalée.
        assert list(df.iloc[0]) == ["Region", "Sexe", "Sexe (M/F)"]
        assert any("Sexe (M/F)" in w for w in res.warnings)
