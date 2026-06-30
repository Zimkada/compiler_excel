"""
Dialogue d'aperçu de la détection avant compilation.
Auteur: GOUNOU N'GOBI Chabi Zimé
"""

from typing import List, Optional, Dict, Tuple

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QScrollArea, QWidget, QFrame
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

from core.compilation import FilePreview
from ui.styles import theme as T


class PreviewDialog(QDialog):
    """
    Affiche, par fichier, ce que le système a détecté : ligne d'en-tête,
    confiance, libellés de colonnes et nombre de lignes de données.

    Permet à l'utilisateur de vérifier la détection ET de corriger
    manuellement, fichier par fichier, la ligne d'en-tête mal identifiée
    (bouton « Corriger l'en-tête » → sélection visuelle).

    Les corrections sont écrites dans le dict `overrides` fourni
    ({chemin: (header_start_row, header_rows)}), que l'appelant réutilise
    ensuite pour la compilation — garantissant que l'aperçu et la
    compilation produisent le même résultat.
    """

    def __init__(self, previews: List[FilePreview], parent=None,
                 compiler=None, overrides: Optional[Dict[str, Tuple[int, int]]] = None):
        super().__init__(parent)
        self.previews = list(previews)
        # Compilateur servant à recalculer l'aperçu d'un fichier après correction.
        self._compiler = compiler
        # Dict partagé d'overrides (modifié en place ; relu par l'appelant).
        self.overrides = overrides if overrides is not None else {}
        # Index des cartes par chemin, pour rafraîchissement ciblé.
        self._cards: Dict[str, QFrame] = {}
        self.setup_ui()

    def setup_ui(self):
        self.setWindowTitle("Aperçu de la détection")
        self.setMinimumSize(680, 500)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        # Titre + résumé
        self.title = QLabel("👁 Aperçu de la détection")
        self.title.setFont(QFont(T.FONT_FAMILY, 14, QFont.Weight.Bold))
        self.title.setStyleSheet(f"color: {T.TEXT_PRIMARY};")
        layout.addWidget(self.title)

        self.subtitle = QLabel()
        self.subtitle.setStyleSheet(
            f"color: {T.TEXT_SECONDARY}; font-size: {T.FONT_SIZE_SM}pt;")
        self.subtitle.setWordWrap(True)
        layout.addWidget(self.subtitle)
        self._update_subtitle()

        # Zone scrollable des cartes par fichier
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        content = QWidget()
        self._content_layout = QVBoxLayout(content)
        self._content_layout.setSpacing(8)

        for preview in self.previews:
            card = self._make_card(preview)
            self._cards[preview.file_path] = card
            self._content_layout.addWidget(card)
        self._content_layout.addStretch()

        scroll.setWidget(content)
        layout.addWidget(scroll, stretch=1)

        # Boutons : rattacher les colonnes (multi-fichiers) + fermer
        button_row = QHBoxLayout()
        # Le rattachement n'a de sens qu'avec ≥ 2 fichiers lisibles à aligner.
        ok_files = [p for p in self.previews if p.success]
        if self._compiler is not None and len(ok_files) >= 2:
            map_btn = QPushButton("🔗 Rattacher les colonnes")
            map_btn.setToolTip(
                "Rattacher manuellement une colonne portant un libellé différent "
                "à une colonne du schéma de référence (1er fichier)."
            )
            map_btn.clicked.connect(self._open_column_mapper)
            button_row.addWidget(map_btn)
        button_row.addStretch()
        close_btn = QPushButton("Fermer")
        close_btn.setMinimumWidth(100)
        close_btn.clicked.connect(self.accept)
        button_row.addWidget(close_btn)
        layout.addLayout(button_row)

    def _update_subtitle(self):
        ok_count = sum(1 for p in self.previews if p.success)
        self.subtitle.setText(
            f"{ok_count}/{len(self.previews)} fichier(s) analysé(s). "
            "Vérifiez la ligne d'en-tête ; corrigez-la si elle est erronée "
            "avant de compiler."
        )

    def _make_card(self, preview: FilePreview) -> QFrame:
        card = QFrame()
        card.setFrameShape(QFrame.Shape.StyledPanel)
        card.setStyleSheet(
            f"QFrame {{ background-color: {T.BG_SURFACE}; border: 1px solid {T.BORDER}; "
            f"border-radius: {T.RADIUS_MD}px; }}"
        )
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(12, 10, 12, 10)
        card_layout.setSpacing(4)

        # Ligne 1 : nom de fichier + indicateur + bouton corriger
        header_row = QHBoxLayout()
        name = QLabel(preview.filename)
        name.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        header_row.addWidget(name)
        header_row.addStretch()

        is_override = preview.detection_method == "override"
        if not preview.success:
            badge = QLabel("⛔ Erreur")
            badge.setStyleSheet(f"color: {T.DANGER}; font-weight: bold;")
        elif is_override:
            badge = QLabel("✏ Corrigé")
            badge.setStyleSheet(f"color: {T.ACCENT}; font-weight: 700;")
        elif preview.confidence >= 0.65 or preview.detection_method == "reference":
            badge = QLabel("● Détecté")
            badge.setStyleSheet(f"color: {T.SUCCESS}; font-weight: 700;")
        elif preview.detection_method == "manual":
            badge = QLabel("⚙ Manuel")
            badge.setStyleSheet(f"color: {T.TEXT_SECONDARY};")
        else:
            badge = QLabel("▲ Confiance faible")
            badge.setStyleSheet(f"color: {T.WARNING}; font-weight: 700;")
        header_row.addWidget(badge)

        # Bouton de correction manuelle (sur tous les fichiers lisibles)
        if preview.success and self._compiler is not None:
            fix_btn = QPushButton("Corriger l'en-tête")
            fix_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            fix_btn.setFont(QFont("Segoe UI", 8))
            fix_btn.clicked.connect(
                lambda _=False, fp=preview.file_path: self._open_picker(fp)
            )
            header_row.addWidget(fix_btn)
        card_layout.addLayout(header_row)

        if not preview.success:
            err = QLabel(preview.error or "Erreur inconnue")
            err.setStyleSheet(f"color: {T.DANGER}; font-size: {T.FONT_SIZE_SM}pt;")
            err.setWordWrap(True)
            card_layout.addWidget(err)
            return card

        # Ligne 2 : infos de détection
        info_parts = [f"En-tête : ligne {preview.header_start_row}"]
        if preview.header_rows > 1:
            info_parts[-1] += f" ({preview.header_rows} lignes)"
        info_parts.append(f"{preview.data_row_count} ligne(s) de données")
        if preview.detection_method not in ("manual", "override"):
            info_parts.append(f"confiance {preview.confidence:.0%}")
        info = QLabel(" · ".join(info_parts))
        info.setStyleSheet(f"color: {T.TEXT_SECONDARY}; font-size: {T.FONT_SIZE_SM}pt;")
        card_layout.addWidget(info)

        # Ligne 3 : en-têtes détectés
        headers_text = "  |  ".join(
            h for h in preview.detected_headers if h
        )[:300]
        if headers_text:
            headers = QLabel(f"Colonnes : {headers_text}")
            headers.setStyleSheet(f"color: {T.ACCENT}; font-size: {T.FONT_SIZE_SM}pt;")
            headers.setWordWrap(True)
            card_layout.addWidget(headers)

        # Avertissement éventuel
        if preview.warning:
            warn = QLabel(f"⚠ {preview.warning}")
            warn.setStyleSheet(f"color: {T.WARNING}; font-size: 8pt;")
            warn.setWordWrap(True)
            card_layout.addWidget(warn)

        return card

    def _open_column_mapper(self):
        """Ouvre le dialogue de rattachement des colonnes et enregistre les
        alias choisis dans les options du compilateur (relues par l'appelant)."""
        from ui.widgets.column_mapper_dialog import ColumnMapperDialog

        dlg = ColumnMapperDialog(
            self.previews, parent=self,
            aliases=self._compiler.options.column_aliases,
        )
        dlg.exec()

    def _open_picker(self, file_path: str):
        """Ouvre le sélecteur visuel d'en-tête pour un fichier, applique le
        choix comme override, puis rafraîchit la carte concernée."""
        from ui.widgets.header_picker_dialog import HeaderPickerDialog

        # Pré-remplir avec la valeur actuellement affichée pour ce fichier.
        current = next((p for p in self.previews if p.file_path == file_path), None)
        start = current.header_start_row if current else 1
        rows = current.header_rows if current else 1

        picker = HeaderPickerDialog(
            file_path, initial_start_row=start, initial_header_rows=rows, parent=self
        )
        if picker.exec() != QDialog.DialogCode.Accepted:
            return
        chosen = picker.get_result()
        if chosen is None:
            return

        # Enregistrer l'override et recalculer l'aperçu de CE fichier.
        self.overrides[file_path] = chosen
        self._compiler.options.manual_overrides[file_path] = chosen
        new_preview = self._compiler.preview_detection([file_path])[0]

        # Remplacer le preview mémorisé et reconstruire la carte.
        for i, p in enumerate(self.previews):
            if p.file_path == file_path:
                self.previews[i] = new_preview
                break
        self._refresh_card(file_path)
        self._update_subtitle()

    def _refresh_card(self, file_path: str):
        """Remplace la carte d'un fichier par une carte reconstruite."""
        old = self._cards.get(file_path)
        if old is None:
            return
        preview = next((p for p in self.previews if p.file_path == file_path), None)
        if preview is None:
            return
        new_card = self._make_card(preview)
        idx = self._content_layout.indexOf(old)
        self._content_layout.insertWidget(idx, new_card)
        old.setParent(None)
        old.deleteLater()
        self._cards[file_path] = new_card
