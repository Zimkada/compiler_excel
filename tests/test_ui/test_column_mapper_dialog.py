"""
Tests du dialogue de rattachement des colonnes (étape 6, offscreen).

Vérifie le calcul schéma/inconnues à partir des aperçus, et l'écriture des
alias dans le dict partagé selon les choix de combos.
"""

import os
import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6.QtWidgets import QApplication

from core.compilation import FilePreview
from ui.widgets.column_mapper_dialog import ColumnMapperDialog, _KEEP


@pytest.fixture(scope="module")
def app():
    return QApplication.instance() or QApplication([])


def _preview(path, headers):
    return FilePreview(file_path=path, success=True, detected_headers=headers)


def test_schema_and_unknowns_computed(app):
    previews = [
        _preview("f1.xlsx", ["Region", "Sexe", "Cas"]),
        _preview("f2.xlsx", ["Region", "Sexe (M/F)", "Cas"]),
    ]
    dlg = ColumnMapperDialog(previews, aliases={})
    assert dlg._schema_labels == ["Region", "Sexe", "Cas"]
    # Seule « Sexe (M/F) » est inconnue (Region/Cas correspondent).
    assert list(dlg._combos.keys()) == ["Sexe (M/F)"]


def test_apply_writes_alias(app):
    previews = [
        _preview("f1.xlsx", ["Region", "Sexe"]),
        _preview("f2.xlsx", ["Region", "Sexe (M/F)"]),
    ]
    aliases = {}
    dlg = ColumnMapperDialog(previews, aliases=aliases)
    dlg._combos["Sexe (M/F)"].setCurrentText("Sexe")
    dlg._on_apply()
    assert aliases == {"Sexe (M/F)": "Sexe"}


def test_keep_choice_removes_previous_alias(app):
    previews = [
        _preview("f1.xlsx", ["Region", "Sexe"]),
        _preview("f2.xlsx", ["Region", "Sexe (M/F)"]),
    ]
    aliases = {"Sexe (M/F)": "Sexe"}          # alias préexistant
    dlg = ColumnMapperDialog(previews, aliases=aliases)
    # Le combo doit être pré-rempli sur « Sexe ».
    assert dlg._combos["Sexe (M/F)"].currentText() == "Sexe"
    # Repasser sur « garder » doit retirer l'alias.
    dlg._combos["Sexe (M/F)"].setCurrentText(_KEEP)
    dlg._on_apply()
    assert aliases == {}


def test_no_unknowns_no_combos(app):
    previews = [
        _preview("f1.xlsx", ["Region", "Sexe"]),
        _preview("f2.xlsx", ["Region", "Sexe"]),
    ]
    dlg = ColumnMapperDialog(previews, aliases={})
    assert dlg._combos == {}
