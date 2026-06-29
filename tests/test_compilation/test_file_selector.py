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
