"""
Tests du FileSelectorWidget : ajout de fichiers (glisser-déposer), filtrage
des extensions, déduplication et état vide.
"""

import os

import openpyxl
import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

pytest.importorskip("PyQt6")
from PyQt6.QtWidgets import QApplication  # noqa: E402

from ui.widgets.file_selector_widget import FileSelectorWidget  # noqa: E402


@pytest.fixture(scope="module")
def qapp():
    yield QApplication.instance() or QApplication([])


@pytest.fixture
def widget(qapp):
    w = FileSelectorWidget()
    yield w


def _xlsx(path):
    wb = openpyxl.Workbook()
    wb.active.append(["a", "b"])
    wb.save(path)
    return str(path)


class TestAddFiles:
    def test_add_supported_files(self, widget, tmp_path):
        f1 = _xlsx(tmp_path / "a.xlsx")
        f2 = _xlsx(tmp_path / "b.xlsx")
        widget.add_files([f1, f2])
        assert len(widget.all_files) == 2
        assert len(widget.get_selected_files()) == 2  # auto-sélection

    def test_duplicates_ignored(self, widget, tmp_path):
        f1 = _xlsx(tmp_path / "a.xlsx")
        widget.add_files([f1])
        widget.add_files([f1])
        assert len(widget.all_files) == 1

    def test_unsupported_extension_filtered(self, widget, tmp_path):
        f1 = _xlsx(tmp_path / "a.xlsx")
        bad = tmp_path / "note.docx"
        bad.write_text("x")
        widget.add_files([f1, str(bad)])
        assert len(widget.all_files) == 1

    def test_nonexistent_path_ignored(self, widget, tmp_path):
        widget.add_files([str(tmp_path / "ghost.xlsx")])
        assert len(widget.all_files) == 0

    def test_files_sorted_by_name(self, widget, tmp_path):
        fb = _xlsx(tmp_path / "zebra.xlsx")
        fa = _xlsx(tmp_path / "alpha.xlsx")
        widget.add_files([fb, fa])
        names = [widget.list_files.item(i).text() for i in range(widget.list_files.count())]
        assert names == ["alpha.xlsx", "zebra.xlsx"]


class TestSelectAllInteraction:
    """La case « Tout sélectionner » est à la fois une ACTION (clic) et un
    REFLET de l'état de la liste. Confondre les deux faisait qu'en
    désélectionnant UN fichier, la case se décochait et effaçait TOUTE la
    sélection (l'utilisateur perdait ses 46 autres fichiers).
    """

    def _load(self, widget, tmp_path, n=5):
        files = [_xlsx(tmp_path / f"f{i}.xlsx") for i in range(n)]
        widget.add_files(files)
        return files

    def _deselect_one(self, widget, row):
        from PyQt6.QtCore import QItemSelectionModel

        item = widget.list_files.item(row)
        widget.list_files.selectionModel().select(
            widget.list_files.indexFromItem(item),
            QItemSelectionModel.SelectionFlag.Deselect,
        )

    def test_deselecting_one_file_keeps_the_others(self, widget, tmp_path):
        self._load(widget, tmp_path, n=5)
        assert len(widget.get_selected_files()) == 5

        self._deselect_one(widget, 2)

        selected = widget.get_selected_files()
        assert len(selected) == 4
        assert not any(f.endswith("f2.xlsx") for f in selected)

    def test_checkbox_unchecks_on_partial_selection(self, widget, tmp_path):
        self._load(widget, tmp_path, n=3)
        assert widget.checkbox_select_all.isChecked()

        self._deselect_one(widget, 0)
        assert not widget.checkbox_select_all.isChecked()

    def test_checkbox_rechecks_when_all_selected_again(self, widget, tmp_path):
        self._load(widget, tmp_path, n=3)
        self._deselect_one(widget, 0)
        assert not widget.checkbox_select_all.isChecked()

        widget.list_files.selectAll()
        assert widget.checkbox_select_all.isChecked()
        assert len(widget.get_selected_files()) == 3

    def test_checkbox_still_works_as_an_action(self, widget, tmp_path):
        """Non-régression : un vrai clic doit toujours tout (dé)sélectionner."""
        self._load(widget, tmp_path, n=4)

        widget.checkbox_select_all.setChecked(False)  # clic simulé
        assert widget.get_selected_files() == []

        widget.checkbox_select_all.setChecked(True)
        assert len(widget.get_selected_files()) == 4

    def test_count_label_reflects_partial_selection(self, widget, tmp_path):
        self._load(widget, tmp_path, n=5)
        self._deselect_one(widget, 1)
        assert widget.label_file_count.text() == "4/5 fichiers sélectionnés"
