"""
Dialogue de rattachement manuel des colonnes (éditeur 2D — étape 6).
Auteur: GOUNOU N'GOBI Chabi Zimé

Complément assisté de l'alignement par libellé (étape 5). L'alignement
automatique rattache les colonnes de même libellé ; il n'ose PAS deviner qu'une
colonne « Sexe (M/F) » est la colonne « Sexe » du schéma. Ce dialogue laisse
l'utilisateur faire ce rattachement à la main, fichier par fichier.

À partir des aperçus déjà calculés (FilePreview.detected_headers), il prend le
1er fichier comme schéma de référence et liste, pour les autres, les libellés
qui ne correspondent à aucune colonne du schéma. Pour chacun, l'utilisateur
choisit soit de le rattacher à une colonne du schéma, soit de le garder comme
nouvelle colonne. Le résultat alimente options.column_aliases
({libellé_source: libellé_cible}).
"""

from typing import Dict, List, Optional

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QScrollArea, QWidget, QComboBox, QGridLayout
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

from core.compilation import FilePreview
from core.compilation.column_aligner import normalize_label
from ui.styles import theme as T

# Choix « ne pas rattacher » dans les combos (garder comme colonne distincte).
_KEEP = "- Garder comme nouvelle colonne -"


class ColumnMapperDialog(QDialog):
    """Rattachement manuel des colonnes non reconnues vers le schéma de référence."""

    def __init__(self, previews: List[FilePreview], parent=None,
                 aliases: Optional[Dict[str, str]] = None):
        super().__init__(parent)
        self.previews = [p for p in previews if p.success]
        # Dict partagé d'alias (modifié en place ; relu par l'appelant).
        self.aliases = aliases if aliases is not None else {}
        # combos: (libellé_source -> QComboBox) pour relire les choix.
        self._combos: Dict[str, QComboBox] = {}
        self._schema_labels: List[str] = []
        self.setup_ui()

    def _compute_schema_and_unknowns(self):
        """Schéma = libellés du 1er fichier ; inconnues = libellés des autres
        fichiers absents du schéma (par libellé normalisé)."""
        if not self.previews:
            return [], []
        schema = [h for h in self.previews[0].detected_headers if h]
        schema_norm = {normalize_label(h) for h in schema}

        unknown_labels: List[str] = []
        seen_unknown = set()
        for p in self.previews[1:]:
            for h in p.detected_headers:
                if not h:
                    continue
                n = normalize_label(h)
                if n in schema_norm or n in seen_unknown:
                    continue
                seen_unknown.add(n)
                unknown_labels.append(h)
        return schema, unknown_labels

    def setup_ui(self):
        self.setWindowTitle("Rattacher les colonnes")
        self.setMinimumSize(620, 460)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        title = QLabel("🔗 Rattacher les colonnes non reconnues")
        title.setFont(QFont(T.FONT_FAMILY, 14, QFont.Weight.Bold))
        title.setStyleSheet(f"color: {T.TEXT_PRIMARY};")
        layout.addWidget(title)

        self._schema_labels, unknowns = self._compute_schema_and_unknowns()

        subtitle = QLabel(
            "Le premier fichier sert de référence. Les colonnes des autres "
            "fichiers dont le libellé ne correspond à aucune colonne de "
            "référence sont listées ci-dessous : rattachez-les si elles "
            "désignent la même donnée sous un autre nom."
        )
        subtitle.setStyleSheet(f"color: {T.TEXT_SECONDARY}; font-size: {T.FONT_SIZE_SM}pt;")
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        if not unknowns:
            empty = QLabel("✅ Toutes les colonnes correspondent déjà au schéma "
                           "de référence. Aucun rattachement nécessaire.")
            empty.setStyleSheet(f"color: {T.SUCCESS}; font-size: {T.FONT_SIZE_SM}pt;")
            empty.setWordWrap(True)
            layout.addWidget(empty)
        else:
            layout.addWidget(self._build_mapping_area(unknowns), stretch=1)

        # Boutons
        button_row = QHBoxLayout()
        button_row.addStretch()
        cancel_btn = QPushButton("Annuler")
        cancel_btn.setMinimumWidth(100)
        cancel_btn.clicked.connect(self.reject)
        ok_btn = QPushButton("✅ Appliquer")
        ok_btn.setMinimumWidth(120)
        ok_btn.setProperty("variant", "primary")
        ok_btn.clicked.connect(self._on_apply)
        button_row.addWidget(cancel_btn)
        button_row.addWidget(ok_btn)
        layout.addLayout(button_row)

    def _build_mapping_area(self, unknowns: List[str]) -> QWidget:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        content = QWidget()
        grid = QGridLayout(content)
        grid.setColumnStretch(1, 1)
        grid.setVerticalSpacing(8)

        # Pré-remplir les combos avec un alias déjà choisi (réouverture).
        reverse = {normalize_label(k): v for k, v in self.aliases.items()}

        for row, label in enumerate(unknowns):
            src = QLabel(label)
            src.setStyleSheet(f"color: {T.ACCENT}; font-weight: 600;")
            arrow = QLabel("→")
            arrow.setAlignment(Qt.AlignmentFlag.AlignCenter)

            combo = QComboBox()
            combo.addItem(_KEEP)
            for s in self._schema_labels:
                combo.addItem(s)
            # Restaurer un choix précédent s'il existe.
            prev_target = reverse.get(normalize_label(label))
            if prev_target and prev_target in self._schema_labels:
                combo.setCurrentText(prev_target)

            self._combos[label] = combo
            grid.addWidget(src, row, 0)
            grid.addWidget(arrow, row, 1)
            grid.addWidget(combo, row, 2)

        scroll.setWidget(content)
        return scroll

    def _on_apply(self):
        """Construit la table d'alias depuis les combos et la renvoie à
        l'appelant via le dict partagé (clé = libellé source, valeur = cible)."""
        for label, combo in self._combos.items():
            choice = combo.currentText()
            n = normalize_label(label)
            if choice == _KEEP:
                # Désélection : retirer un éventuel alias précédent (par libellé).
                for k in [k for k in self.aliases if normalize_label(k) == n]:
                    del self.aliases[k]
            else:
                self.aliases[label] = choice
        self.accept()
