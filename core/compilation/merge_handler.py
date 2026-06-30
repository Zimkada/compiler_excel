"""
Dé-fusion des cellules fusionnées d'un fichier Excel.
Auteur: GOUNOU N'GOBI Chabi Zimé
Version: 3.2

pandas.read_excel ne place la valeur d'une cellule fusionnée que dans le coin
haut-gauche de la fusion et laisse le reste à NaN. Pour les tableaux réels
(ex. un département « ALIBORI » fusionné verticalement sur 26 lignes de CEG),
cela vide visuellement des colonnes pourtant porteuses d'information.

Ce module lit les vraies plages fusionnées et propage la valeur du coin
haut-gauche sur TOUTE la plage :
    - fusion verticale  -> recopie vers le bas  (fill-down)
    - fusion horizontale -> recopie vers la droite (fill-across)

C'est un traitement strictement déterministe (aucune heuristique) : on ne fait
que matérialiser une information déjà présente dans le fichier source.

Performance : on lit les valeurs en mode ``read_only`` (rapide, ne parse pas
tout le classeur) et les plages de fusion directement dans le XML de la feuille
(les feuilles Excel réelles contiennent des milliers de colonnes fantômes que
le mode normal d'openpyxl parserait intégralement, coûtant plusieurs secondes
par fichier). En cas de difficulté de lecture XML, on retombe sur openpyxl
normal puis sur pandas — jamais de crash, jamais de perte de données.
"""

import zipfile
from typing import List, Tuple, Optional

import pandas as pd
import openpyxl
from openpyxl.utils import range_boundaries

# Au-delà de cette colonne/ligne, une plage « fusionnée » relève d'un artefact
# de fichier (styles appliqués sur des colonnes lointaines) et non d'une vraie
# fusion de tableau. On l'ignore, dans le même esprit que prune_phantom_columns.
_MAX_REAL_COL = 200
_MAX_REAL_ROW = 100_000

# Espaces de noms OOXML pour le parsing des relations et de la feuille.
_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def load_with_unmerge(file_path: str, sheet_name=0) -> pd.DataFrame:
    """Charge une feuille Excel en propageant les cellules fusionnées.

    Args:
        file_path: Chemin du fichier .xlsx/.xlsm.
        sheet_name: Feuille à lire (index ou nom), défaut première feuille.

    Returns:
        DataFrame sans en-tête (header=None), colonnes 0..N-1, fusions
        propagées. Équivaut à ``pd.read_excel(file_path, header=None)`` si le
        fichier n'a aucune fusion réelle. Retombe sur pandas (sans dé-fusion)
        en cas d'échec de lecture, pour ne jamais bloquer la compilation.
    """
    # 1) Lire les valeurs (rapide) en mode read_only.
    grid = _read_values_fast(file_path, sheet_name)
    if grid is None:
        # Lecture rapide impossible : fallback openpyxl complet (lent mais sûr).
        return _load_with_unmerge_openpyxl(file_path, sheet_name)

    n_rows = len(grid)
    n_cols = len(grid[0]) if n_rows else 0
    if n_rows == 0 or n_cols == 0 or all(v is None for row in grid for v in row):
        return pd.DataFrame()

    # 2) Lire les plages de fusion (rapide) depuis le XML.
    merges = _read_merges_fast(file_path, sheet_name)
    if merges is None:
        # Impossible de lire les fusions par le XML : fallback openpyxl complet
        # (les valeurs lues plus haut sont correctes, mais on a besoin des
        # fusions ; on relit tout proprement pour ne pas en perdre).
        return _load_with_unmerge_openpyxl(file_path, sheet_name)

    _propagate_merges(grid, merges, n_rows, n_cols)
    return pd.DataFrame(grid)


def _read_values_fast(file_path: str, sheet_name) -> Optional[List[list]]:
    """Lit les valeurs (formules en cache incluses) en mode read_only.

    Retourne une grille (liste de listes, 0-based) bornée à _MAX_REAL_COL
    colonnes, ou None si la lecture échoue (l'appelant choisit un fallback).
    """
    wb = None
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
        ws = wb.worksheets[sheet_name] if isinstance(sheet_name, int) else wb[sheet_name]
        grid: List[list] = []
        for row in ws.iter_rows(values_only=True):
            grid.append(list(row[:_MAX_REAL_COL]))
        # Uniformiser la largeur des lignes (iter_rows peut renvoyer des lignes
        # de longueurs variables selon le contenu).
        width = max((len(r) for r in grid), default=0)
        width = min(width, _MAX_REAL_COL)
        for r in grid:
            if len(r) < width:
                r.extend([None] * (width - len(r)))
            elif len(r) > width:
                del r[width:]
        return grid
    except Exception:
        return None
    finally:
        if wb is not None:
            wb.close()


def _read_merges_fast(file_path: str, sheet_name) -> Optional[List[Tuple[int, int, int, int]]]:
    """Lit les plages fusionnées directement dans le XML de la feuille.

    Retourne une liste de tuples (min_row, min_col, max_row, max_col) 1-based,
    ou None si le mapping/parsing échoue (l'appelant choisit un fallback).
    """
    try:
        import xml.etree.ElementTree as ET

        with zipfile.ZipFile(file_path) as z:
            sheet_xml_path = _resolve_sheet_path(z, sheet_name)
            if sheet_xml_path is None:
                return None
            xml = z.read(sheet_xml_path)

        root = ET.fromstring(xml)
        merges: List[Tuple[int, int, int, int]] = []
        # <mergeCells><mergeCell ref="A1:B2"/></mergeCells>
        for mc in root.iter(f"{{{_NS_MAIN}}}mergeCell"):
            ref = mc.get("ref")
            if not ref or ":" not in ref:
                continue
            min_col, min_row, max_col, max_row = range_boundaries(ref)
            merges.append((min_row, min_col, max_row, max_col))
        return merges
    except Exception:
        return None


def _resolve_sheet_path(z: zipfile.ZipFile, sheet_name) -> Optional[str]:
    """Résout le chemin XML interne de la feuille demandée via les relations
    officielles du classeur (workbook.xml + .rels), pour gérer correctement les
    classeurs à feuilles multiples ou réordonnées.
    """
    import xml.etree.ElementTree as ET

    try:
        wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    except KeyError:
        return None

    # Ordre des feuilles + leur r:id.
    sheets = []  # (name, r_id)
    for sh in wb_xml.iter(f"{{{_NS_MAIN}}}sheet"):
        r_id = sh.get(f"{{{_NS_REL_DOC}}}id")
        sheets.append((sh.get("name"), r_id))
    if not sheets:
        return None

    # r:id -> target (chemin du sheetN.xml).
    rid_to_target = {}
    for rel in rels_xml.iter(f"{{{_NS_PKG_REL}}}Relationship"):
        rid_to_target[rel.get("Id")] = rel.get("Target")

    # Sélectionner la feuille demandée.
    if isinstance(sheet_name, int):
        if sheet_name < 0 or sheet_name >= len(sheets):
            return None
        _, r_id = sheets[sheet_name]
    else:
        r_id = next((rid for nm, rid in sheets if nm == sheet_name), None)
    if r_id is None:
        return None

    target = rid_to_target.get(r_id)
    if not target:
        return None
    # Normaliser le chemin (les targets sont relatifs à xl/).
    target = target.lstrip("/")
    if not target.startswith("xl/"):
        target = "xl/" + target
    return target


def _propagate_merges(grid: List[list], merges: List[Tuple[int, int, int, int]],
                      n_rows: int, n_cols: int) -> None:
    """Propage en place la valeur du coin haut-gauche de chaque fusion réelle
    sur toute sa plage (bornée à la taille de la grille)."""
    for min_row, min_col, max_row, max_col in merges:
        if min_row == max_row and min_col == max_col:
            continue  # fusion dégénérée
        if min_col > _MAX_REAL_COL or min_row > _MAX_REAL_ROW:
            continue  # artefact hors zone réelle
        r0, c0 = min_row - 1, min_col - 1
        if r0 < 0 or c0 < 0 or r0 >= n_rows or c0 >= n_cols:
            continue
        value = grid[r0][c0]
        if value is None or value == "":
            continue  # rien à propager
        r1 = min(max_row - 1, n_rows - 1)
        c1 = min(max_col - 1, n_cols - 1)
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                grid[r][c] = value


def _load_with_unmerge_openpyxl(file_path: str, sheet_name) -> pd.DataFrame:
    """Chemin de repli robuste : lecture openpyxl complète (data_only) avec
    rattrapage des formules non cachées par leur texte, puis propagation des
    fusions. Plus lent (parse tout le classeur) mais sans dépendance au XML.
    Utilisé si la voie rapide échoue (format inhabituel, XML illisible).
    """
    try:
        wb_val = openpyxl.load_workbook(file_path, data_only=True)
    except Exception:
        # Dernier recours : pandas. Si lui aussi échoue (fichier réellement
        # illisible / non-Excel), relever une erreur claire plutôt qu'un
        # message technique opaque — le moteur l'affichera telle quelle.
        try:
            return pd.read_excel(file_path, header=None, sheet_name=sheet_name)
        except Exception as e:
            raise ValueError(
                f"Fichier illisible ou format Excel non reconnu : {e}"
            ) from e

    wb_fml = None
    try:
        try:
            ws = (wb_val.worksheets[sheet_name] if isinstance(sheet_name, int)
                  else wb_val[sheet_name])
        except (KeyError, IndexError):
            return pd.read_excel(file_path, header=None, sheet_name=sheet_name)

        n_rows = ws.max_row or 0
        n_cols = min(ws.max_column or 0, _MAX_REAL_COL)
        if n_rows == 0 or n_cols == 0:
            return pd.DataFrame()

        grid = [[ws.cell(r, c).value for c in range(1, n_cols + 1)]
                for r in range(1, n_rows + 1)]
        if all(v is None for row in grid for v in row):
            return pd.DataFrame()

        # Rattraper les formules sans valeur cachée par leur texte.
        ws_fml = None
        for r in range(n_rows):
            for c in range(n_cols):
                if grid[r][c] is None:
                    if ws_fml is None:
                        wb_fml = openpyxl.load_workbook(file_path, data_only=False)
                        ws_fml = (wb_fml.worksheets[sheet_name]
                                  if isinstance(sheet_name, int)
                                  else wb_fml[sheet_name])
                    raw = ws_fml.cell(r + 1, c + 1).value
                    if isinstance(raw, str) and raw.startswith("="):
                        grid[r][c] = raw

        merges = [(rng.min_row, rng.min_col, rng.max_row, rng.max_col)
                  for rng in ws.merged_cells.ranges]
        _propagate_merges(grid, merges, n_rows, n_cols)
        return pd.DataFrame(grid)
    finally:
        wb_val.close()
        if wb_fml is not None:
            wb_fml.close()
