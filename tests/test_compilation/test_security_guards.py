"""
Tests des garde-fous de sécurité réellement appliqués au chargement (Phase 1).

Vérifie que les limites historiquement déclarées mais jamais vérifiées sont
désormais effectives :

- max_file_size_mb : un fichier trop volumineux est rejeté PROPREMENT (marqué
  en échec) sans interrompre la compilation des autres fichiers.
- max_rows_per_file : un fichier aux dimensions démesurées est tronqué (borne
  anti-explosion mémoire), avec avertissement, jamais d'échec silencieux.
- Désactivation (valeur 0) : aucune limite appliquée.
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


def _options(**overrides):
    base = dict(
        use_reference_mode=True,
        auto_detect_structure=False,
        reference_header_row=1,
        reference_header_lines=1,
        filename_option=FilenameOption.NONE,
        remove_empty_rows=True,
    )
    base.update(overrides)
    return CompilationOptions(**base)


# --------------------------------------------------------------------------- #
# max_file_size_mb
# --------------------------------------------------------------------------- #
def test_oversized_file_rejected_others_continue(tmp_path):
    """Un fichier au-dessus de la limite est marqué en échec ; les autres
    fichiers sont compilés normalement."""
    big = tmp_path / "big.xlsx"
    small = tmp_path / "small.xlsx"
    header = ["Nom", "Valeur"]
    _make_xlsx(big, [header] + [[f"x{i}", i] for i in range(50)])
    _make_xlsx(small, [header, ["a", 1], ["b", 2]])

    # Limite minuscule : le 1er (référence) passe car la référence est lue par
    # le détecteur, mais le contrôle de taille s'applique au CHARGEMENT. On met
    # une limite qui rejette 'big' (plus gros) et garde 'small'.
    big_mb = big.stat().st_size / (1024 * 1024)
    small_mb = small.stat().st_size / (1024 * 1024)
    limit = (big_mb + small_mb) / 2  # entre les deux

    out = tmp_path / "out.xlsx"
    # 'small' en premier pour servir de référence valide.
    result = ExcelCompiler(_options(max_file_size_mb=limit)).compile_files(
        [str(small), str(big)], str(out), OutputFormat.XLSX
    )

    # La compilation globale aboutit, 'big' est en échec, 'small' réussit.
    assert result.successful_files == 1
    assert result.failed_files == 1
    # Un message explicite de taille figure dans les résultats du fichier échoué.
    failed = [r for r in result.file_results if not r.success]
    assert failed and "volumineux" in (failed[0].error_message or "").lower()


def test_no_size_limit_when_zero(tmp_path):
    """max_file_size_mb=0 désactive complètement la limite."""
    header = ["Nom", "Valeur"]
    p = tmp_path / "f.xlsx"
    _make_xlsx(p, [header, ["a", 1], ["b", 2]])
    out = tmp_path / "out.xlsx"
    result = ExcelCompiler(_options(max_file_size_mb=0)).compile_files(
        [str(p)], str(out), OutputFormat.XLSX
    )
    assert result.success
    assert result.successful_files == 1


# --------------------------------------------------------------------------- #
# max_rows_per_file
# --------------------------------------------------------------------------- #
def test_row_cap_truncates_with_warning(tmp_path):
    """Au-delà de max_rows_per_file, les lignes sont tronquées et signalées."""
    header = ["Nom", "Valeur"]
    rows = [header] + [[f"n{i}", i] for i in range(20)]  # 1 en-tête + 20 données
    p = tmp_path / "many.xlsx"
    _make_xlsx(p, rows)

    out = tmp_path / "out.xlsx"
    # Cap à 6 lignes BRUTES lues (en-tête incluse) -> 5 lignes de données max.
    result = ExcelCompiler(_options(max_rows_per_file=6)).compile_files(
        [str(p)], str(out), OutputFormat.XLSX
    )
    assert result.success
    produced = pd.read_excel(out, header=None).values.tolist()
    # Comportement observable : la troncature est effective (en-tête + <= 5
    # lignes de données au lieu des 20 d'origine).
    assert len(produced) <= 6
    assert len(produced) < 21
    # Transparence : un avertissement de troncature est remonté à l'utilisateur
    # (pas seulement dans le log).
    assert any("tronqu" in w.lower() for w in result.warnings)


def test_no_row_cap_when_zero(tmp_path):
    """max_rows_per_file=0 lit toutes les lignes."""
    header = ["Nom", "Valeur"]
    rows = [header] + [[f"n{i}", i] for i in range(20)]
    p = tmp_path / "many.xlsx"
    _make_xlsx(p, rows)
    out = tmp_path / "out.xlsx"
    result = ExcelCompiler(_options(max_rows_per_file=0)).compile_files(
        [str(p)], str(out), OutputFormat.XLSX
    )
    assert result.success
    produced = pd.read_excel(out, header=None).values.tolist()
    assert len(produced) == 21  # en-tête + 20 données, rien tronqué
