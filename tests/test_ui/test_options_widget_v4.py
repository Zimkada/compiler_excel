"""
Tests des améliorations de confort de l'OptionsWidget (Vague 4).

- G4 : validation de la colonne de tri (parse_sort_column, sort_column_is_valid).
- G2 : exclusivité native des 3 modes (QRadioButton + QButtonGroup) et
  visibilité des panneaux de configuration.
- G1 : persistance des options via QSettings (sauvegarde / restauration).

Nécessite une QApplication (widgets Qt) ; ignoré si PyQt6 absent.
"""

import pytest

pytest.importorskip("PyQt6.QtWidgets")

from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import QSettings

from ui.widgets.options_widget import OptionsWidget


@pytest.fixture(scope="module")
def qapp():
    app = QApplication.instance() or QApplication([])
    yield app


@pytest.fixture
def widget(qapp):
    # Repartir d'un état de settings propre pour chaque test.
    QSettings(OptionsWidget._SETTINGS_ORG,
              OptionsWidget._SETTINGS_APP).remove("options")
    w = OptionsWidget()
    w.show()  # sinon isVisible() renvoie toujours False
    yield w
    w.close()
    QSettings(OptionsWidget._SETTINGS_ORG,
              OptionsWidget._SETTINGS_APP).remove("options")


# --------------------------------------------------------------------------- #
# G4 — validation de la colonne de tri
# --------------------------------------------------------------------------- #
class TestSortColumnParsing:
    @pytest.mark.parametrize("text,expected", [
        ("A", 0), ("B", 1), ("Z", 25), ("AA", 26), ("aa", 26),
        ("1", 0), ("3", 2), ("10", 9),
    ])
    def test_valid(self, text, expected):
        assert OptionsWidget.parse_sort_column(text) == expected

    @pytest.mark.parametrize("text", ["1A", "A1", "", "  ", "!", "0", "-1", "A-"])
    def test_invalid_returns_none(self, text):
        assert OptionsWidget.parse_sort_column(text) is None

    def test_valid_when_sort_disabled(self, widget):
        widget.checkbox_sort_data.setChecked(False)
        widget.lineedit_sort_column.setText("1A")  # invalide mais tri éteint
        assert widget.sort_column_is_valid() is True

    def test_invalid_when_sort_enabled(self, widget):
        widget.checkbox_sort_data.setChecked(True)
        widget.lineedit_sort_column.setText("1A")
        assert widget.sort_column_is_valid() is False

    def test_valid_when_sort_enabled_good_column(self, widget):
        widget.checkbox_sort_data.setChecked(True)
        widget.lineedit_sort_column.setText("C")
        assert widget.sort_column_is_valid() is True


# --------------------------------------------------------------------------- #
# G2 — exclusivité et visibilité des modes
# --------------------------------------------------------------------------- #
class TestModeExclusivity:
    def test_reference_is_default(self, widget):
        assert widget.checkbox_reference_mode.isChecked()
        assert not widget.checkbox_auto_mode.isChecked()
        assert not widget.checkbox_manual_mode.isChecked()

    def test_selecting_auto_deselects_others(self, widget):
        widget.checkbox_auto_mode.setChecked(True)
        assert widget.checkbox_auto_mode.isChecked()
        assert not widget.checkbox_reference_mode.isChecked()
        assert not widget.checkbox_manual_mode.isChecked()

    def test_manual_panel_shown_in_manual_mode(self, widget):
        widget.checkbox_manual_mode.setChecked(True)
        assert not widget.manual_group.isHidden()
        assert widget.reference_group.isHidden()

    def test_reference_panel_shown_in_reference_mode(self, widget):
        widget.checkbox_auto_mode.setChecked(True)
        widget.checkbox_reference_mode.setChecked(True)
        assert not widget.reference_group.isHidden()
        assert widget.manual_group.isHidden()


# --------------------------------------------------------------------------- #
# G1 — persistance
# --------------------------------------------------------------------------- #
class TestPersistence:
    def test_options_restored_after_save(self, qapp):
        QSettings(OptionsWidget._SETTINGS_ORG,
                  OptionsWidget._SETTINGS_APP).remove("options")
        w1 = OptionsWidget()
        w1.checkbox_auto_mode.setChecked(True)
        w1.spinbox_ref_header.setValue(7)
        w1.checkbox_remove_duplicates.setChecked(True)
        w1.combo_format.setCurrentIndex(1)
        w1.save_settings()

        w2 = OptionsWidget()  # restore_settings dans __init__
        assert w2.checkbox_auto_mode.isChecked()
        assert w2.spinbox_ref_header.value() == 7
        assert w2.checkbox_remove_duplicates.isChecked()
        assert w2.combo_format.currentIndex() == 1

        QSettings(OptionsWidget._SETTINGS_ORG,
                  OptionsWidget._SETTINGS_APP).remove("options")

    def test_defaults_when_no_saved_settings(self, qapp):
        QSettings(OptionsWidget._SETTINGS_ORG,
                  OptionsWidget._SETTINGS_APP).remove("options")
        w = OptionsWidget()
        # Valeurs par défaut de l'UI préservées.
        assert w.checkbox_reference_mode.isChecked()
        assert w.spinbox_ref_header.value() == 1

    def test_restored_options_produce_identical_compilation_options(self, qapp):
        """INVARIANT critique : après restauration, get_compilation_options()
        doit être IDENTIQUE à ce qui a été sauvé. Sinon l'utilisateur qui
        redémarre obtient une compilation différente de ce qu'il a configuré.
        Régression trouvée en certification : sort_column n'était pas persistée
        (le tri retombait silencieusement sur la colonne A)."""
        QSettings(OptionsWidget._SETTINGS_ORG,
                  OptionsWidget._SETTINGS_APP).remove("options")
        w1 = OptionsWidget()
        w1.checkbox_manual_mode.setChecked(True)
        w1.spinbox_header_start.setValue(4)
        w1.spinbox_header_rows.setValue(2)
        w1.combo_format.setCurrentIndex(2)  # TSV
        w1.combo_date_format.setCurrentIndex(3)
        w1.checkbox_remove_empty.setChecked(False)
        w1.checkbox_drop_subtotals.setChecked(False)
        w1.checkbox_sort_data.setChecked(True)
        w1.lineedit_sort_column.setText("C")
        opt1 = w1.get_compilation_options()
        w1.save_settings()

        w2 = OptionsWidget()
        opt2 = w2.get_compilation_options()

        for f in ['use_reference_mode', 'auto_detect_structure',
                  'manual_header_start_row', 'manual_header_rows',
                  'date_format', 'remove_empty_rows', 'drop_subtotal_rows',
                  'sort_data', 'sort_column']:
            assert getattr(opt1, f) == getattr(opt2, f), f"champ {f} divergent"
        assert w1.get_output_format() == w2.get_output_format()
        assert opt2.sort_column == 2  # colonne C bien restaurée

        QSettings(OptionsWidget._SETTINGS_ORG,
                  OptionsWidget._SETTINGS_APP).remove("options")

    def test_output_name_persisted(self, qapp):
        QSettings(OptionsWidget._SETTINGS_ORG,
                  OptionsWidget._SETTINGS_APP).remove("options")
        w1 = OptionsWidget()
        w1.lineedit_output.setText("rapport_special.xlsx")
        w1.save_settings()
        w2 = OptionsWidget()
        assert w2.get_output_file() == "rapport_special.xlsx"
        QSettings(OptionsWidget._SETTINGS_ORG,
                  OptionsWidget._SETTINGS_APP).remove("options")
