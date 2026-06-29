"""
Tests du widget d'options : exposition des 3 modes de détection
(automatique / fichier de référence / manuel) et exclusivité mutuelle.

Ces tests instancient le vrai widget PyQt6 en mode offscreen. Ils sont
ignorés si PyQt6 n'est pas disponible.
"""

import os

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

pytest.importorskip("PyQt6")
from PyQt6.QtWidgets import QApplication  # noqa: E402

from ui.widgets.options_widget import OptionsWidget  # noqa: E402


@pytest.fixture(scope="module")
def qapp():
    app = QApplication.instance() or QApplication([])
    yield app


@pytest.fixture
def widget(qapp):
    w = OptionsWidget()
    w.show()  # nécessaire pour que isVisible() reflète l'état en offscreen
    yield w
    w.close()


class TestDetectionModes:
    def test_default_is_reference_mode(self, widget):
        o = widget.get_compilation_options()
        assert o.use_reference_mode is True
        assert o.auto_detect_structure is False

    def test_auto_mode_sets_auto_detect(self, widget):
        widget.checkbox_auto_mode.setChecked(True)
        o = widget.get_compilation_options()
        assert o.auto_detect_structure is True
        assert o.use_reference_mode is False

    def test_manual_mode(self, widget):
        widget.checkbox_manual_mode.setChecked(True)
        o = widget.get_compilation_options()
        assert o.auto_detect_structure is False
        assert o.use_reference_mode is False

    def test_modes_are_mutually_exclusive(self, widget):
        widget.checkbox_auto_mode.setChecked(True)
        assert not widget.checkbox_reference_mode.isChecked()
        assert not widget.checkbox_manual_mode.isChecked()

        widget.checkbox_manual_mode.setChecked(True)
        assert not widget.checkbox_auto_mode.isChecked()
        assert not widget.checkbox_reference_mode.isChecked()

    def test_unchecking_last_mode_falls_back_to_reference(self, widget):
        # référence est coché par défaut ; le décocher doit le re-cocher
        widget.checkbox_reference_mode.setChecked(False)
        assert widget.checkbox_reference_mode.isChecked()

    def test_panels_visibility_follows_mode(self, widget):
        widget.checkbox_manual_mode.setChecked(True)
        assert widget.manual_group.isVisibleTo(widget)
        assert not widget.reference_group.isVisibleTo(widget)

        widget.checkbox_reference_mode.setChecked(True)
        assert widget.reference_group.isVisibleTo(widget)
        assert not widget.manual_group.isVisibleTo(widget)

        widget.checkbox_auto_mode.setChecked(True)
        assert not widget.reference_group.isVisibleTo(widget)
        assert not widget.manual_group.isVisibleTo(widget)

    def test_common_options_passed_in_auto_mode(self, widget):
        widget.checkbox_auto_mode.setChecked(True)
        widget.checkbox_remove_duplicates.setChecked(True)
        o = widget.get_compilation_options()
        assert o.auto_detect_structure is True
        assert o.remove_duplicates is True

    def test_is_auto_mode_enabled_helper(self, widget):
        widget.checkbox_auto_mode.setChecked(True)
        assert widget.is_auto_mode_enabled() is True
