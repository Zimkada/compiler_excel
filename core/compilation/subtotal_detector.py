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


# Marqueurs d'un bloc de signature en pied de tableau. Observés tels quels sur
# les formulaires administratifs réels (lieu/date, qualité du signataire).
# NB : l'apostrophe n'est pas une frontière de mot pour \b, et _normalize (mise
# en commun avec la détection des totaux) ne la remplace pas. « CHEF
# D'ETABLISSEMENT » est donc listé sous ses deux graphies plutôt que de toucher
# à une normalisation partagée.
SIGNATURE_KEYWORDS: List[str] = [
    "FAIT A",
    "LE DIRECTEUR",
    "LA DIRECTRICE",
    "CHEF D ETABLISSEMENT",
    "CHEF D'ETABLISSEMENT",
    "CHEF ETABLISSEMENT",
    "LE CENSEUR",
    "LE PROVISEUR",
    "SIGNATURE",
    "CACHET",
]


def is_signature_row(row: Sequence,
                     identity_col: int = 0,
                     keywords: Sequence[str] = SIGNATURE_KEYWORDS) -> bool:
    """Vrai si la ligne appartient au bloc de signature sous le tableau.

    Un formulaire rempli se termine souvent par « Fait à X, le … » suivi de la
    qualité et du nom du signataire. Ces lignes sont sous le tableau, pas
    dedans : compilées, elles produisent des lignes sans établissement dont le
    texte atterrit dans des colonnes de chiffres.

    Trois conditions CUMULATIVES, pour ne jamais sacrifier une vraie donnée :
      1. la colonne d'identité (``identity_col``, la 1re par défaut) est vide —
         une ligne de données nomme toujours son établissement ;
      2. la ligne ne contient AUCUNE valeur numérique — une ligne de données
         porte des chiffres, un bloc de signature n'en a pas ;
      3. au moins une cellule correspond à un marqueur de signature.

    La condition 3 exige VRAIMENT un marqueur. Une version antérieure acceptait
    aussi « une seule cellule de texte » pour rattraper le nom du signataire
    seul sous un « Le Directeur ». Cette tolérance supprimait des DONNÉES : une
    ligne ``[None, 'NEANT', None]`` (établissement n'ayant pas rempli sa colonne
    d'identité) ou ``[None, None, 'Aucun élève inscrit']`` était comptée comme
    signature et effacée silencieusement. Un nom de signataire isolé est
    désormais conservé — une ligne en trop est anodine, une donnée perdue ne
    l'est pas.

    N'est PAS du ressort de cette fonction : les lignes de total/sous-total,
    qui ont leur propre classification (``classify_row``) et leur propre option
    utilisateur. Elles sont donc explicitement exclues ici, sans quoi un
    « TOTAL GENERAL » sans identité serait compté comme signature et supprimé
    même lorsque l'utilisateur a demandé de conserver les totaux.
    """
    values = list(row)
    if not values:
        return False

    # 1. Colonne d'identité renseignée -> c'est une ligne de données.
    if 0 <= identity_col < len(values):
        ident = values[identity_col]
        if ident is not None and str(ident).strip() and str(ident).strip().lower() != 'nan':
            return False

    # Une ligne d'agrégat relève de classify_row, pas de la détection de
    # signature : la laisser passer ici court-circuiterait drop_subtotal_rows.
    if classify_row(row) != ROW_KIND_DETAIL:
        return False

    texts: List[str] = []
    for val in values:
        if val is None:
            continue
        if isinstance(val, bool):
            continue
        # 2. Une valeur numérique => ligne de données, jamais une signature.
        if isinstance(val, (int, float)):
            if not (isinstance(val, float) and val != val):  # ignorer NaN
                return False
            continue
        text = str(val).strip()
        if not text or text.lower() == 'nan':
            continue
        # Un nombre saisi en texte compte aussi comme valeur numérique.
        if _looks_numeric(text):
            return False
        texts.append(text)

    if not texts:
        return False  # ligne vide : traitée par le filtre de lignes vides

    # 3. Un marqueur de signature est EXIGÉ (cf. docstring) : sans lui, on
    #    conserve la ligne, quitte à garder une mention de trop.
    return _row_contains(texts, keywords)


def _looks_numeric(text: str) -> bool:
    """Vrai si le texte représente un nombre (« 30 », « 1 234,5 », « 12.0 »)."""
    cleaned = text.replace(" ", "").replace(" ", "").replace(",", ".")
    if not cleaned:
        return False
    try:
        float(cleaned)
        return True
    except ValueError:
        return False
