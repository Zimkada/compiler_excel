"""
Détection et classification des lignes de total intercalées dans les données.
Auteur: GOUNOU N'GOBI Chabi Zimé
Version: 3.2

Les tableaux administratifs intercalent souvent des lignes d'agrégat
(« ENSEMBLE COMMUNE », « ENSEMBLE DEPARTEMENT », « TOTAL GENERAL ») au milieu
des données. Les compiler sans précaution fausserait tout calcul ultérieur
(double comptage : la ligne de total additionne déjà les lignes au-dessus).

La détection est strictement déterministe : une ligne est classée selon qu'une
de ses cellules contient l'un des mots-clés configurés en tant que mot(s)
entier(s) (insensible à la casse et aux accents). Deux niveaux sont distingués :
    - « total »     : agrégat de niveau supérieur (ENSEMBLE DEPARTEMENT, TOTAL GENERAL)
    - « sous-total » : agrégat intermédiaire (ENSEMBLE COMMUNE)
    - « détail »    : ligne de donnée normale

Aucune décision n'est silencieuse côté métier : le moteur exclut par défaut mais
signale le compte ; sinon il garde les lignes et peut les marquer (colonne
« Type de ligne »). L'utilisateur peut ajuster les mots-clés.

Limite assumée : le tiret/slash comptant comme frontière de mot (\\b), un libellé
de donnée tel que « total-ville » serait classé comme total. Choix délibéré : il
faut conserver « SOUS-TOTAL » et « S/TOTAL » (fréquents) ; un établissement
réellement nommé « total-… » est invraisemblable, et reste rattrapable à la main.
"""

import re
import unicodedata
from functools import lru_cache
from typing import List, Sequence

# Niveaux de ligne (valeurs de la colonne marqueur « Type de ligne »).
ROW_KIND_DETAIL = "détail"
ROW_KIND_SUBTOTAL = "sous-total"
ROW_KIND_TOTAL = "total"

# Mots-clés par niveau, observés sur les tableaux réels. Comparaison sans accents
# ni casse, espaces compactés. Les mots-clés de « total » (niveau supérieur) sont
# testés EN PREMIER : « ENSEMBLE DEPARTEMENT » doit l'emporter sur « ENSEMBLE ».
DEFAULT_TOTAL_KEYWORDS: List[str] = [
    "ENSEMBLE DEPARTEMENT",
    "TOTAL GENERAL",
    "TOTAL GENERALE",
]
DEFAULT_SUBTOTAL_KEYWORDS: List[str] = [
    "ENSEMBLE COMMUNE",
    "ENSEMBLE",
    "SOUS-TOTAL",
    "SOUS TOTAL",
    "TOTAL",
    "TOTAUX",
]


def _normalize(text: str) -> str:
    """Minuscule, sans accents, espaces compactés — pour une comparaison robuste."""
    nfkd = unicodedata.normalize("NFKD", text)
    no_accents = "".join(c for c in nfkd if not unicodedata.combining(c))
    return " ".join(no_accents.lower().split())


@lru_cache(maxsize=256)
def _keyword_pattern(keyword: str):
    """Compile un motif qui matche le mot-clé comme séquence de MOTS ENTIERS.

    Évite les faux positifs en sous-chaîne : « TOTAL » ne doit pas matcher
    « TOTALITE », ni « ENSEMBLE » matcher « RASSEMBLEMENT ». Les frontières \\b
    encadrent le mot-clé normalisé.
    """
    norm = _normalize(keyword)
    return re.compile(r"\b" + re.escape(norm) + r"\b")


def _row_contains(row: Sequence, keywords: Sequence[str]) -> bool:
    """Vrai si une cellule texte de la ligne contient l'un des mots-clés en
    tant que mot(s) entier(s), après normalisation casse/accents/espaces."""
    patterns = [_keyword_pattern(k) for k in keywords if k]
    if not patterns:
        return False
    for val in row:
        if val is None or isinstance(val, (int, float)):
            continue
        text = _normalize(str(val))
        if not text:
            continue
        if any(p.search(text) for p in patterns):
            return True
    return False


def classify_row(row: Sequence,
                 subtotal_keywords: Sequence[str] = DEFAULT_SUBTOTAL_KEYWORDS,
                 total_keywords: Sequence[str] = DEFAULT_TOTAL_KEYWORDS) -> str:
    """Classe une ligne en ROW_KIND_TOTAL, ROW_KIND_SUBTOTAL ou ROW_KIND_DETAIL.

    Les mots-clés de « total » sont testés avant ceux de « sous-total » pour
    qu'un agrégat de niveau supérieur (« ENSEMBLE DEPARTEMENT ») ne soit pas
    classé à tort comme sous-total à cause du mot « ENSEMBLE ».
    """
    if _row_contains(row, total_keywords):
        return ROW_KIND_TOTAL
    if _row_contains(row, subtotal_keywords):
        return ROW_KIND_SUBTOTAL
    return ROW_KIND_DETAIL


def is_total_row(row: Sequence,
                 subtotal_keywords: Sequence[str] = DEFAULT_SUBTOTAL_KEYWORDS,
                 total_keywords: Sequence[str] = DEFAULT_TOTAL_KEYWORDS) -> bool:
    """Vrai si la ligne est un total OU un sous-total (à exclure si drop activé)."""
    return classify_row(row, subtotal_keywords, total_keywords) != ROW_KIND_DETAIL
