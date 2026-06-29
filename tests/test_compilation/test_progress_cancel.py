"""
Tests des callbacks de progression et d'annulation de compile_files.

Vérifie:
- B.1: la progression émise est croissante, atteint 100%, et expose des
  paliers par fichier (et non un saut 0 -> 100).
- B.2: l'annulation arrête réellement la compilation, sans écrire de sortie.
- La robustesse (callback défaillant non fatal) et l'absence de fuite de
  callbacks entre deux compilations sur la même instance.
"""

import openpyxl
import pandas as pd

from core.compilation import ExcelCompiler, CompilationOptions, OutputFormat
from core.compilation.compilation_models import FilenameOption


def _make_xlsx(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    wb.save(path)


def _options():
    return CompilationOptions(
        use_reference_mode=True, auto_detect_structure=False,
        reference_header_row=1, reference_header_lines=1,
        filename_option=FilenameOption.NONE, remove_empty_rows=True,
    )


def _make_files(tmp_path, count):
    paths = []
    for i in range(count):
        p = tmp_path / f"f{i}.xlsx"
        _make_xlsx(p, [["Nom", "Age"], [f"A{i}", i], [f"B{i}", i + 1]])
        paths.append(str(p))
    return paths


class TestProgress:
    def test_progress_is_monotonic_and_reaches_100(self, tmp_path):
        paths = _make_files(tmp_path, 5)
        events = []
        ExcelCompiler(_options()).compile_files(
            paths, str(tmp_path / "out.xlsx"), OutputFormat.XLSX,
            progress_callback=lambda pct, msg: events.append((pct, msg)),
        )
        pcts = [e[0] for e in events]
        assert len(events) > 2
        assert pcts == sorted(pcts)        # monotone croissante
        assert pcts[-1] == 100

    def test_progress_has_per_file_steps(self, tmp_path):
        paths = _make_files(tmp_path, 5)
        events = []
        ExcelCompiler(_options()).compile_files(
            paths, str(tmp_path / "out.xlsx"), OutputFormat.XLSX,
            progress_callback=lambda pct, msg: events.append((pct, msg)),
        )
        mid = {p for p, _ in events if 15 <= p < 80}
        assert len(mid) >= 3  # plusieurs paliers intermédiaires distincts
        assert any("Traitement" in m for _, m in events)

    def test_failing_progress_callback_is_not_fatal(self, tmp_path):
        paths = _make_files(tmp_path, 3)

        def bad(pct, msg):
            raise RuntimeError("boom UI")

        out = tmp_path / "out.xlsx"
        result = ExcelCompiler(_options()).compile_files(
            paths, str(out), OutputFormat.XLSX, progress_callback=bad,
        )
        assert result.success
        assert out.exists()


class TestCancellation:
    def test_cancel_immediately_writes_nothing(self, tmp_path):
        paths = _make_files(tmp_path, 3)
        out = tmp_path / "out.xlsx"
        result = ExcelCompiler(_options()).compile_files(
            paths, str(out), OutputFormat.XLSX, cancel_check=lambda: True,
        )
        assert result.cancelled is True
        assert result.success is False
        assert not out.exists()

    def test_cancel_midway_stops(self, tmp_path):
        paths = _make_files(tmp_path, 5)
        state = {"n": 0}

        def cancel():
            state["n"] += 1
            return state["n"] > 2

        out = tmp_path / "out.xlsx"
        result = ExcelCompiler(_options()).compile_files(
            paths, str(out), OutputFormat.XLSX, cancel_check=cancel,
        )
        assert result.cancelled is True
        assert not out.exists()


class TestCallbackHygiene:
    def test_callbacks_do_not_leak_between_runs(self, tmp_path):
        paths = _make_files(tmp_path, 3)
        comp = ExcelCompiler(_options())

        ev1 = []
        comp.compile_files(paths, str(tmp_path / "a.xlsx"), OutputFormat.XLSX,
                           progress_callback=lambda p, m: ev1.append(p))
        count_after_first = len(ev1)
        assert count_after_first > 0

        # Un appel SANS callback ne doit pas réutiliser l'ancien
        comp.compile_files(paths, str(tmp_path / "b.xlsx"), OutputFormat.XLSX)
        assert len(ev1) == count_after_first

    def test_no_callbacks_still_works(self, tmp_path):
        paths = _make_files(tmp_path, 3)
        out = tmp_path / "out.xlsx"
        result = ExcelCompiler(_options()).compile_files(
            paths, str(out), OutputFormat.XLSX,
        )
        assert result.success
        rows = pd.read_excel(out, header=None).values.tolist()
        names = [r[0] for r in rows if r[0] not in (None, "Nom")]
        assert len(names) == 6
