"""
Dialogue de sélection visuelle de la ligne d'en-tête d'un fichier.
Auteur: GOUNOU N'GOBI Chabi Zimé

Affiche les premières lignes d'un fichier dans une grille. L'utilisateur
clique la ligne qui contient le vrai en-tête (et, si besoin, indique combien
de lignes l'en-tête occupe). Retourne (header_start_row, header_rows) en
numérotation 1-based, ou None si annulé.
"""

from pathlib import Path
from typing import Optional, Tuple

import pandas as pd
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QSpinBox, QHeaderView, QAbstractItemView
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QColor

from core.detection.base_detector import prune_phantom_columns
from ui.styles import theme as T

# Nombre de lignes affichées dans l'aperçu (large mais borné pour la perf)
_MAX_PREVIEW_ROWS = 40
# Nombre de colonnes affichées (au-delà, on tronque pour rester lisible)
_MAX_PREVIEW_COLS = 20


class HeaderPickerDialog(QDialog):
    """Sélecteur visuel de la ligne d'en-tête d'un seul fichier."""

    def __init__(self, file_path: str,
                 initial_start_row: int = 1,
                 initial_header_rows: int = 1,
                 parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self._result: Optional[Tuple[int, int]] = None
        self._initial_start_row = max(1, initial_start_row)
        self._initial_header_rows = max(1, initial_header_rows)
        self._df = None
        self._load_error: Optional[str] = None
        self._load_dataframe()
        self.setup_ui()

    def _load_dataframe(self):
        """Charge les premières lignes du fichier (tous formats supportés)."""
        try:
            ext = Path(self.file_path).suffix.lower()
            if ext in ('.xlsx', '.xls', '.xlsm'):
                df = pd.read_excel(self.file_path, header=None, nrows=_MAX_PREVIEW_ROWS)
            elif ext == '.csv':
                df = None
                for enc in ('utf-8-sig', 'utf-8', 'latin-1', 'cp1252'):
                    try:
                        df = pd.read_csv(self.file_path, header=None,
                                         nrows=_MAX_PREVIEW_ROWS, encoding=enc)
                        break
                    except UnicodeDecodeError:
                        continue
                if df is None:
                    raise ValueError("Impossible de décoder le fichier CSV")
            elif ext in ('.tsv', '.txt'):
                df = pd.read_csv(self.file_path, header=None, sep='\t',
                                 nrows=_MAX_PREVIEW_ROWS, encoding='utf-8-sig')
            else:
                raise ValueError(f"Format non supporté : {ext}")

            df, _ = prune_phantom_columns(df)
            self._df = df
        except Exception as e:  # noqa: BLE001 — on reporte l'erreur à l'UI
            self._load_error = str(e)

    def setup_ui(self):
        self.setWindowTitle(f"Choisir l'en-tête — {Path(self.file_path).name}")
        self.setMinimumSize(720, 520)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        title = QLabel("🖱 Cliquez sur la ligne d'en-tête")
        title.setFont(QFont(T.FONT_FAMILY, 14, QFont.Weight.Bold))
        title.setStyleSheet(f"color: {T.TEXT_PRIMARY};")
        layout.addWidget(title)

        subtitle = QLabel(
            "Sélectionnez la ligne qui contient les titres de colonnes. "
            "Les lignes au-dessus seront ignorées ; les lignes en dessous "
            "deviendront les données."
        )
        subtitle.setStyleSheet(f"color: {T.TEXT_SECONDARY}; font-size: {T.FONT_SIZE_SM}pt;")
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        if self._load_error is not None:
            err = QLabel(f"⛔ Lecture impossible : {self._load_error}")
            err.setStyleSheet(f"color: {T.DANGER};")
            err.setWordWrap(True)
            layout.addWidget(err)
            close = QPushButton("Fermer")
            close.clicked.connect(self.reject)
            layout.addWidget(close, alignment=Qt.AlignmentFlag.AlignRight)
            return

        # Grille des lignes
        self.table = self._build_table()
        layout.addWidget(self.table, stretch=1)

        # Contrôle « nombre de lignes d'en-tête »
        rows_row = QHBoxLayout()
        rows_label = QLabel("Nombre de lignes d'en-tête :")
        rows_label.setStyleSheet(f"color: {T.TEXT_PRIMARY};")
        self.spin_header_rows = QSpinBox()
        self.spin_header_rows.setMinimum(1)
        self.spin_header_rows.setMaximum(10)
        self.spin_header_rows.setValue(self._initial_header_rows)
        self.spin_header_rows.setToolTip(
            "Si l'en-tête tient sur plusieurs lignes (titres fusionnés), "
            "indiquez combien de lignes il occupe."
        )
        self.spin_header_rows.valueChanged.connect(self._update_highlight)
        rows_row.addWidget(rows_label)
        rows_row.addWidget(self.spin_header_rows)
        rows_row.addStretch()

        self.selection_label = QLabel("")
        self.selection_label.setStyleSheet(
            f"color: {T.ACCENT}; font-weight: 700;")
        rows_row.addWidget(self.selection_label)
        layout.addLayout(rows_row)

        # Boutons
        button_row = QHBoxLayout()
        button_row.addStretch()
        cancel_btn = QPushButton("Annuler")
        cancel_btn.setMinimumWidth(100)
        cancel_btn.clicked.connect(self.reject)
        validate_btn = QPushButton("✅ Valider l'en-tête")
        validate_btn.setMinimumWidth(140)
        validate_btn.setProperty("variant", "primary")
        validate_btn.clicked.connect(self._on_validate)
        button_row.addWidget(cancel_btn)
        button_row.addWidget(validate_btn)
        layout.addLayout(button_row)

        # Sélection initiale
        self.table.selectRow(self._initial_start_row - 1)
        self._update_highlight()

    def _build_table(self) -> QTableWidget:
        df = self._df
        n_rows = min(len(df), _MAX_PREVIEW_ROWS)
        n_cols = min(df.shape[1], _MAX_PREVIEW_COLS)

        table = QTableWidget(n_rows, n_cols)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        # En-tête vertical = numéro de ligne réel (1-based) du fichier
        table.setVerticalHeaderLabels([str(i + 1) for i in range(n_rows)])
        table.setHorizontalHeaderLabels(
            [self._col_letter(c) for c in range(n_cols)]
        )

        for r in range(n_rows):
            for c in range(n_cols):
                val = df.iat[r, c]
                text = "" if pd.isna(val) else str(val)
                item = QTableWidgetItem(text)
                table.setItem(r, c, item)

        table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Interactive)
        table.resizeColumnsToContents()
        table.itemSelectionChanged.connect(self._update_highlight)
        return table

    @staticmethod
    def _col_letter(idx: int) -> str:
        """0 -> A, 25 -> Z, 26 -> AA (notation tableur)."""
        result = ""
        idx += 1
        while idx > 0:
            idx, rem = divmod(idx - 1, 26)
            result = chr(ord('A') + rem) + result
        return result

    def _selected_start_row(self) -> int:
        """Ligne d'en-tête sélectionnée (1-based), défaut 1 si rien."""
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return 1
        return rows[0].row() + 1

    def _update_highlight(self):
        """Surligne visuellement la plage d'en-tête (start .. start+rows-1)."""
        start = self._selected_start_row()
        header_rows = self.spin_header_rows.value()
        highlight = QColor(T.ACCENT)
        highlight.setAlpha(45)
        n_rows = self.table.rowCount()
        n_cols = self.table.columnCount()

        for r in range(n_rows):
            in_header = start - 1 <= r < start - 1 + header_rows
            for c in range(n_cols):
                item = self.table.item(r, c)
                if item is None:
                    continue
                item.setBackground(highlight if in_header else QColor(0, 0, 0, 0))

        end_data = start + header_rows
        self.selection_label.setText(
            f"En-tête : ligne {start}"
            + (f"–{start + header_rows - 1}" if header_rows > 1 else "")
            + f"  ·  données dès la ligne {end_data}"
        )

    def _on_validate(self):
        self._result = (self._selected_start_row(), self.spin_header_rows.value())
        self.accept()

    def get_result(self) -> Optional[Tuple[int, int]]:
        """Retourne (header_start_row, header_rows) ou None si annulé."""
        return self._result
