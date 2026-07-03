"""
Tests de la vérification de mise à jour (H2).

Couvre la logique PURE (comparaison de versions, parsing de la réponse GitHub)
et le comportement silencieux/non-bloquant du worker en cas d'échec réseau.
La vérification ne doit JAMAIS crasher ni bloquer l'application.
"""

import json

import pytest

from utils.update_checker import (
    parse_version, is_newer, evaluate_release, UpdateInfo, RELEASES_PAGE_URL,
)


# --------------------------------------------------------------------------- #
# parse_version
# --------------------------------------------------------------------------- #
class TestParseVersion:
    @pytest.mark.parametrize("text,expected", [
        ("3.2", (3, 2)),
        ("v3.2", (3, 2)),
        ("3.2.1", (3, 2, 1)),
        ("version 3.2.0", (3, 2, 0)),
        ("3.2.1-beta", (3, 2, 1)),
        ("v10.0", (10, 0)),
    ])
    def test_valid(self, text, expected):
        assert parse_version(text) == expected

    @pytest.mark.parametrize("text", ["", "abc", None, "vX.Y"])
    def test_invalid(self, text):
        assert parse_version(text) is None


# --------------------------------------------------------------------------- #
# is_newer
# --------------------------------------------------------------------------- #
class TestIsNewer:
    @pytest.mark.parametrize("latest,current,expected", [
        ("v3.3", "3.2", True),
        ("3.2.1", "3.2", True),
        ("v3.10", "v3.9", True),        # comparaison numérique, pas lexicale
        ("3.2", "3.2", False),
        ("3.2", "3.2.1", False),
        ("3.2", "v3.3", False),
        ("", "3.2", False),             # version illisible -> jamais de MAJ
        ("abc", "3.2", False),
    ])
    def test(self, latest, current, expected):
        assert is_newer(latest, current) is expected


# --------------------------------------------------------------------------- #
# evaluate_release
# --------------------------------------------------------------------------- #
class TestEvaluateRelease:
    def test_update_available(self):
        body = json.dumps({
            "tag_name": "v3.3",
            "html_url": "https://github.com/x/y/releases/tag/v3.3",
        })
        info = evaluate_release(body, "3.2")
        assert info.update_available is True
        assert info.latest_version == "v3.3"
        assert "releases" in info.download_url

    def test_no_update_same_version(self):
        body = json.dumps({"tag_name": "v3.2"})
        info = evaluate_release(body, "3.2")
        assert info.update_available is False

    def test_no_update_older_release(self):
        body = json.dumps({"tag_name": "v3.1"})
        assert evaluate_release(body, "3.2").update_available is False

    def test_malformed_json_is_silent(self):
        info = evaluate_release("pas du json", "3.2")
        assert info.update_available is False
        assert isinstance(info, UpdateInfo)

    def test_empty_json_is_silent(self):
        assert evaluate_release("{}", "3.2").update_available is False

    def test_missing_tag_falls_back_to_releases_page(self):
        body = json.dumps({"name": "v3.5"})  # 'name' toléré si 'tag_name' absent
        info = evaluate_release(body, "3.2")
        assert info.update_available is True

    def test_download_url_defaults_when_absent(self):
        body = json.dumps({"tag_name": "v3.9"})  # pas de html_url
        info = evaluate_release(body, "3.2")
        assert info.download_url == RELEASES_PAGE_URL


# --------------------------------------------------------------------------- #
# Worker : silencieux et non-bloquant hors-ligne
# --------------------------------------------------------------------------- #
class TestWorkerOffline:
    def test_offline_emits_nothing_and_finishes(self):
        pytest.importorskip("PyQt6.QtWidgets")
        from PyQt6.QtWidgets import QApplication
        import utils.update_checker as uc
        from ui.workers.update_worker import UpdateCheckWorker

        QApplication.instance() or QApplication([])
        original = uc.LATEST_RELEASE_URL
        try:
            # Port fermé -> échec réseau immédiat.
            uc.LATEST_RELEASE_URL = "http://127.0.0.1:9/nope"
            worker = UpdateCheckWorker("3.2", timeout=2.0)
            emitted = []
            worker.update_found.connect(lambda info: emitted.append(info))
            worker.start()
            worker.wait(8000)
            assert worker.isFinished()
            assert emitted == []  # silencieux : aucune notification
        finally:
            uc.LATEST_RELEASE_URL = original
