"""
Alignement des colonnes par libellé entre fichiers de structures différentes.
Auteur: GOUNOU N'GOBI Chabi Zimé
Version: 3.2

Plusieurs fichiers à compiler n'ont pas toujours EXACTEMENT les mêmes colonnes :
ordre différent, une colonne en plus ici, une colonne en moins là. L'empilement
positionnel (ligne par ligne, colonne par colonne) corrompt alors les données
silencieusement : la colonne « Sexe » d'un fichier se retrouve empilée sous la
colonne « Âge » d'un autre simplement parce qu'elles occupent le même rang.

Cet aligneur projette chaque fichier sur un SCHÉMA GLOBAL de colonnes identifiées
par leur LIBELLÉ normalisé (sans accents ni casse, espaces compactés) :

    - colonne commune        -> empilée au bon endroit, quel que soit son rang ;
    - colonne absente du fichier -> cellule vide (None), jamais inventée ;
    - colonne inconnue (présente dans ce fichier seulement) -> ajoutée à droite
      du schéma global et signalée (transparence, jamais silencieux).

Le schéma global est INCRÉMENTAL : il s'étend à mesure que de nouveaux libellés
apparaissent. Les colonnes déjà vues gardent leur position ; les nouvelles sont
appendues. C'est volontairement « assisté » et non magique : on ne fusionne
jamais deux libellés différents, on ne devine aucune correspondance floue.

Cas des libellés DUPLIQUÉS dans un même fichier (ex. deux colonnes « Effectif ») :
les occurrences sont appariées dans l'ordre (1re→1re, 2e→2e) pour ne pas écraser
l'une avec l'autre.
"""

import unicodedata
from typing import List, Optional, Sequence, Tuple


def normalize_label(value) -> str:
    """Minuscule, sans accents, espaces compactés — clé d'identité d'une colonne.

    Les valeurs vides / None donnent une clé vide ; deux colonnes sans libellé
    ne sont donc PAS fusionnées par erreur (voir _column_keys)."""
    if value is None:
        return ""
    text = str(value)
    nfkd = unicodedata.normalize("NFKD", text)
    no_accents = "".join(c for c in nfkd if not unicodedata.combining(c))
    return " ".join(no_accents.lower().split())


def _column_keys(label_row: Sequence,
                 aliases: Optional[dict] = None) -> List[str]:
    """Transforme une ligne de libellés en clés d'identité, en désambiguïsant
    les doublons et les colonnes sans nom.

    - Un libellé non vide répété reçoit un suffixe d'occurrence (#0, #1, …) pour
      apparier les occurrences positionnellement sans les fusionner.
    - Une colonne sans libellé reçoit une clé positionnelle unique (jamais
      fusionnée avec une autre colonne vide).
    - Si `aliases` (libellé_source_normalisé -> libellé_cible_normalisé) est
      fourni, un libellé source aliasé prend l'identité de sa CIBLE : la colonne
      fusionne alors avec la colonne cible du schéma (mapping manuel, étape 6).
    """
    aliases = aliases or {}
    keys: List[str] = []
    seen: dict = {}
    for idx, raw in enumerate(label_row):
        norm = normalize_label(raw)
        if not norm:
            # Colonne sans libellé : identité strictement positionnelle.
            keys.append(f"\x00empty\x00{idx}")
            continue
        # Rattachement manuel : le libellé source devient son libellé cible.
        norm = aliases.get(norm, norm)
        occ = seen.get(norm, 0)
        seen[norm] = occ + 1
        keys.append(f"{norm}#{occ}")
    return keys


class ColumnAligner:
    """Aligne incrémentalement les fichiers sur un schéma global par libellé.

    Usage :
        aligner = ColumnAligner()
        aligner.set_reference(global_header_row)          # 1er fichier
        rows, new_labels = aligner.align(file_header_row, file_rows)  # suivants
        final_header_row = aligner.schema_labels()        # à la fin
    """

    def __init__(self, column_aliases: Optional[dict] = None):
        # Clés d'identité des colonnes du schéma global, dans l'ordre final.
        self._keys: List[str] = []
        # Libellé d'affichage associé à chaque clé (celui vu en premier).
        self._labels: List = []
        # Alias manuels (étape 6), normalisés source -> cible. Appliqués aux
        # fichiers projetés, jamais à la référence (qui définit les cibles).
        self._aliases = {
            normalize_label(k): normalize_label(v)
            for k, v in (column_aliases or {}).items()
            if normalize_label(k) and normalize_label(v)
        }

    def set_reference(self, header_row: Sequence) -> None:
        """Initialise le schéma global à partir des libellés du 1er fichier."""
        self._keys = _column_keys(header_row)
        self._labels = list(header_row)

    @property
    def initialized(self) -> bool:
        return bool(self._keys)

    def schema_labels(self) -> List:
        """Libellés du schéma global final (en-tête de sortie aligné)."""
        return list(self._labels)

    def width(self) -> int:
        return len(self._keys)

    def align(self, header_row: Sequence,
              rows: List[List]) -> Tuple[List[List], List]:
        """Projette les lignes d'un fichier sur le schéma global.

        Étend le schéma avec les colonnes inconnues de ce fichier (appendues à
        droite), puis réordonne chaque ligne pour que la colonne i de la sortie
        corresponde à la colonne i du schéma. Les colonnes absentes du fichier
        valent None.

        Renvoie (lignes_alignées, libellés_des_colonnes_nouvelles_ajoutées).
        """
        file_keys = _column_keys(header_row, self._aliases)

        # 1) Étendre le schéma avec les nouvelles colonnes (ordre d'apparition).
        new_labels: List = []
        key_to_schema_idx = {k: i for i, k in enumerate(self._keys)}
        for fk, label in zip(file_keys, header_row):
            if fk not in key_to_schema_idx:
                key_to_schema_idx[fk] = len(self._keys)
                self._keys.append(fk)
                self._labels.append(label)
                new_labels.append(label)

        # 2) Position de chaque colonne du fichier dans le schéma.
        #    file_col_idx -> schema_idx
        file_to_schema = [key_to_schema_idx[fk] for fk in file_keys]

        # 3) Réordonner chaque ligne sur la largeur du schéma.
        width = len(self._keys)
        aligned: List[List] = []
        for row in rows:
            out: List[Optional[object]] = [None] * width
            for file_col, value in enumerate(row):
                if file_col >= len(file_to_schema):
                    # Cellule au-delà des libellés connus (ligne plus large que
                    # l'en-tête) : ignorée, faute de colonne identifiable.
                    continue
                out[file_to_schema[file_col]] = value
            aligned.append(out)

        return aligned, new_labels
