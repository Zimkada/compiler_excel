"""
Tests de non-régression du BorderDetector sur le nombre de lignes d'en-tête.

Bug corrigé : dans un tableau entièrement quadrillé, le détecteur étendait
les en-têtes sur toutes les lignes de données (header_rows = 5, 13, voire 29),
car il s'appuyait sur la seule densité de bordures. data_start_row tombait
alors après la fin du fichier => 0 ligne de données extraite.

Correctif : la fin des en-têtes est déterminée par le contenu (les lignes de
données contiennent des valeurs numériques, les en-têtes sont textuels), avec
un garde-fou pour les tableaux 100 % textuels (en-tête = 1 ligne par défaut).
"""

import openpyxl
from openpyxl.styles import Border, Side

from core.detection.border_detector import BorderDetector


_THIN = Side(style="thin")
_FULL_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _make_bordered_xlsx(path, rows, border_from_row=1):
    wb = openpyxl.Workbook()
    ws = wb.active
    for ri, row in enumerate(rows, 1):
        for ci, val in enumerate(row, 1):
            cell = ws.cell(ri, ci, val)
            if ri >= border_from_row:
                cell.border = _FULL_BORDER
    wb.save(path)
    return str(path)


class TestBorderHeaderRows:
    def test_single_header_row_in_fully_gridded_table(self, tmp_path):
        """Tableau quadrillé, en-tête textuel + données numériques => 1 ligne."""
        p = _make_bordered_xlsx(tmp_path / "t.xlsx", [
            ["N° ordre", "Matricule", "Nom"],
            [1, 79974, "OLODO"],
            [2, 35304, "MIGAN"],
            [3, 116541, "YAROU"],
        ])
        r = BorderDetector().detect(p)
        assert r.header_start_row == 1
        assert r.header_rows == 1
        assert r.data_start_row == 2

    def test_leading_blank_rows_then_header(self, tmp_path):
        """En-tête précédé de lignes vides + titres (cas des fichiers réels)."""
        p = _make_bordered_xlsx(tmp_path / "t.xlsx", [
            [None, None, None],
            [None, None, None],
            ["Nom", "Age", "Ville"],
            ["Alice", 30, "Paris"],
            ["Bob", 25, "Lyon"],
        ], border_from_row=3)
        r = BorderDetector().detect(p)
        assert r.header_start_row == 3
        assert r.header_rows == 1
        assert r.data_start_row == 4

    def test_multiline_header(self, tmp_path):
        """En-tête réel sur 2 lignes textuelles avant des données numériques."""
        p = _make_bordered_xlsx(tmp_path / "t.xlsx", [
            ["Info", "Info"],
            ["Nom", "Age"],
            ["Alice", 30],
            ["Bob", 25],
        ])
        r = BorderDetector().detect(p)
        assert r.header_start_row == 1
        assert r.header_rows == 2

    def test_all_text_table_defaults_to_one_header_row(self, tmp_path):
        """Tableau 100 % texte : pas de signal numérique => en-tête = 1 ligne,
        les données ne doivent pas être avalées."""
        p = _make_bordered_xlsx(tmp_path / "t.xlsx", [
            ["Nom", "Ville"],
            ["Alice", "Paris"],
            ["Bob", "Lyon"],
            ["Carol", "Nice"],
        ])
        r = BorderDetector().detect(p)
        assert r.header_rows == 1
        assert r.data_start_row == 2

    def test_numeric_cell_in_header_row_is_safe(self, tmp_path):
        """Si la ligne d'en-tête contient un nombre, on garde 1 ligne (sûr)."""
        p = _make_bordered_xlsx(tmp_path / "t.xlsx", [
            ["Nom", 2024],
            ["Alice", 12],
            ["Bob", 15],
        ])
        r = BorderDetector().detect(p)
        assert r.header_rows == 1
        assert r.data_start_row == 2
