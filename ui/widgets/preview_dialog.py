"""
Dialogue d'aperçu de la détection avant compilation.
Auteur: GOUNOU N'GOBI Chabi Zimé
"""

from typing import List

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QScrollArea, QWidget, QFrame
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

from core.compilation import FilePreview
from ui.styles import EXCEL_GREEN, OFFICE_ORANGE
from ui.styles import theme as T


class PreviewDialog(QDialog):
    """
    Affiche, par fichier, ce que le système a détecté : ligne d'en-tête,
    confiance, libellés de colonnes et nombre de lignes de données.

    Permet à l'utilisateur de vérifier la détection avant de compiler.
    """

    def __init__(self, previews: List[FilePreview], parent=None):
        super().__init__(parent)
        self.previews = previews
        self.setup_ui()

    def setup_ui(self):
        self.setWindowTitle("Aperçu de la détection")
        self.setMinimumSize(640, 480)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        # Titre + résumé
        ok_count = sum(1 for p in self.previews if p.success)
        title = QLabel("👁 Aperçu de la détection")
        title.setFont(QFont(T.FONT_FAMILY, 14, QFont.Weight.Bold))
        title.setStyleSheet(f"color: {T.TEXT_PRIMARY};")
        layout.addWidget(title)

        subtitle = QLabel(
            f"{ok_count}/{len(self.previews)} fichier(s) analysé(s). "
            "Vérifiez la ligne d'en-tête détectée avant de compiler."
        )
        subtitle.setStyleSheet(f"color: {T.TEXT_SECONDARY}; font-size: {T.FONT_SIZE_SM}pt;")
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        # Zone scrollable des cartes par fichier
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setSpacing(8)

        for preview in self.previews:
            content_layout.addWidget(self._make_card(preview))
        content_layout.addStretch()

        scroll.setWidget(content)
        layout.addWidget(scroll, stretch=1)

        # Bouton fermer
        button_row = QHBoxLayout()
        button_row.addStretch()
        close_btn = QPushButton("Fermer")
        close_btn.setMinimumWidth(100)
        close_btn.clicked.connect(self.accept)
        button_row.addWidget(close_btn)
        layout.addLayout(button_row)

    def _make_card(self, preview: FilePreview) -> QWidget:
        card = QFrame()
        card.setFrameShape(QFrame.Shape.StyledPanel)
        card.setStyleSheet(
            f"QFrame {{ background-color: {T.BG_SURFACE}; border: 1px solid {T.BORDER}; "
            f"border-radius: {T.RADIUS_MD}px; }}"
        )
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(12, 10, 12, 10)
        card_layout.setSpacing(4)

        # Ligne 1 : nom de fichier + indicateur
        header_row = QHBoxLayout()
        name = QLabel(preview.filename)
        name.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        header_row.addWidget(name)
        header_row.addStretch()

        if not preview.success:
            badge = QLabel("⛔ Erreur")
            badge.setStyleSheet(f"color: {T.DANGER}; font-weight: bold;")
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
        if preview.detection_method not in ("manual",):
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
