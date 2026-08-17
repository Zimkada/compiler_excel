"""
Tests du message d'échec de compilation présenté à l'utilisateur.

Bug corrigé : le worker composait « Compilation échouée : … » en concaténant
les TROIS PREMIERS avertissements, quelle que soit leur nature. Or la plupart
des avertissements sont informatifs (« N colonne(s) vide(s) parasite(s)
ignorée(s) », « colonne inconnue ajoutée », « lignes de total exclues ») et
accompagnent aussi les compilations réussies. L'utilisateur voyait donc une
erreur bloquante lui demandant de vérifier ses colonnes, alors que la cause
réelle de l'échec était ailleurs.
"""

import os

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

pytest.importorskip("PyQt6")

from ui.workers.compilation_worker import CompilationWorker  # noqa: E402


class _Result:
    """Double minimal de CompilationResult pour le composeur de message."""

    def __init__(self, warnings=None, file_results=None):
        self.warnings = warnings or []
        self.file_results = file_results or []


class _FileResult:
    def __init__(self, file_path, success, error_message=None):
        self.file_path = file_path
        self.success = success
        self.error_message = error_message


class TestFailureMessage:
    def test_informative_warnings_are_not_presented_as_the_cause(self):
        """Le cas du bug : que des avertissements informatifs."""
        result = _Result(warnings=[
            "LTA BNK.xlsx: 1 colonne(s) vide(s) parasite(s) ignorée(s) "
            "(le tableau utile fait 4 colonne(s)).",
            "GOUMORI.xlsx: 1 colonne(s) inconnue(s) ajoutée(s) : observations",
        ])
        msg = CompilationWorker._build_failure_message(result)

        assert "parasite" not in msg
        assert "observations" not in msg
        assert "aucune donnée exploitable" in msg.lower()

    def test_blocking_warning_is_reported(self):
        result = _Result(warnings=[
            "fichier.xlsx: 2 colonne(s) vide(s) parasite(s) ignorée(s).",
            "Aucune donnée compilée",
        ])
        msg = CompilationWorker._build_failure_message(result)

        assert "Aucune donnée compilée" in msg
        assert "parasite" not in msg

    def test_fatal_error_is_reported(self):
        result = _Result(warnings=["Erreur fatale: fichier corrompu"])
        msg = CompilationWorker._build_failure_message(result)
        assert "Erreur fatale: fichier corrompu" in msg

    def test_failed_files_are_named(self):
        """À défaut d'avertissement bloquant, on nomme les fichiers en échec."""
        result = _Result(
            warnings=["x.xlsx: 1 colonne(s) vide(s) parasite(s) ignorée(s)."],
            file_results=[
                _FileResult("C:/data/a.xlsx", False, "Erreur chargement fichier"),
                _FileResult("C:/data/b.xlsx", True),
            ],
        )
        msg = CompilationWorker._build_failure_message(result)

        assert "a.xlsx" in msg
        assert "Erreur chargement fichier" in msg
        assert "b.xlsx" not in msg

    def test_message_always_starts_with_the_verdict(self):
        for result in (
            _Result(),
            _Result(warnings=["Aucune donnée compilée"]),
            _Result(file_results=[_FileResult("f.xlsx", False, "boom")]),
        ):
            assert CompilationWorker._build_failure_message(result).startswith(
                "Compilation échouée"
            )

    def test_no_warnings_at_all_gives_actionable_guidance(self):
        """Sans le moindre indice, le message doit orienter vers la bonne
        question plutôt que rester muet."""
        msg = CompilationWorker._build_failure_message(_Result())
        assert "en-tête" in msg.lower()
        assert "sélectionnés" in msg.lower()
