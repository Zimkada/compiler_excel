"""
Formatage Excel pour la sortie de compilation
Auteur: GOUNOU N'GOBI Chabi Zimé
Version: 3.2
"""

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
from typing import List, Optional


# Couleurs Excel
EXCEL_COLORS = {
    "PRIMARY": "217346",  # Vert Excel
    "LIGHT_TEXT": "FFFFFF",  # Blanc
}

# Formats de date Excel
DATE_FORMATS_EXCEL = {
    "FRENCH": "dd/mm/yyyy",
    "AMERICAN": "mm/dd/yyyy",
    "ISO": "yyyy-mm-dd",
    "DATETIME_FRENCH": "dd/mm/yyyy hh:mm:ss",
}


class ExcelFormatter:
    """
    Classe responsable du formatage du fichier Excel de sortie.
    Applique des styles professionnels aux en-têtes et données.
    """

    # Styles prédéfinis
    HEADER_STYLE = {
        'font': Font(bold=True, color=EXCEL_COLORS["LIGHT_TEXT"], size=11),
        'fill': PatternFill(
            start_color=EXCEL_COLORS["PRIMARY"],
            end_color=EXCEL_COLORS["PRIMARY"],
            fill_type='solid'
        ),
        'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True)
    }

    DATA_BORDER = Border(
        left=Side(border_style='thin', color='D0D0D0'),
        right=Side(border_style='thin', color='D0D0D0'),
        top=Side(border_style='thin', color='D0D0D0'),
        bottom=Side(border_style='thin', color='D0D0D0')
    )

    @staticmethod
    def write_headers(worksheet, headers: List[List], start_row: int) -> int:
        """
        Écrit les en-têtes avec style professionnel

        Args:
            worksheet: Feuille de calcul openpyxl
            headers: Liste de listes représentant les en-têtes
            start_row: Ligne de début

        Returns:
            Numéro de la ligne suivante
        """
        current_row = start_row

        for header_row in headers:
            for col_idx, value in enumerate(header_row, 1):
                cell = worksheet.cell(row=current_row, column=col_idx, value=value)
                cell.font = ExcelFormatter.HEADER_STYLE['font']
                cell.fill = ExcelFormatter.HEADER_STYLE['fill']
                cell.alignment = ExcelFormatter.HEADER_STYLE['alignment']
                cell.border = ExcelFormatter.DATA_BORDER

            current_row += 1

        return current_row

    @staticmethod
    def write_data(worksheet, data: List[List], start_row: int,
                   date_format: str = "FRENCH") -> int:
        """
        Écrit les données avec bordures et format de date

        Args:
            worksheet: Feuille de calcul openpyxl
            data: Liste de listes représentant les données
            start_row: Ligne de début
            date_format: Format de date (FRENCH, AMERICAN, ISO)

        Returns:
            Numéro de la ligne suivante
        """
        current_row = start_row
        excel_date_format = DATE_FORMATS_EXCEL.get(date_format, "dd/mm/yyyy")

        for row_data in data:
            for col_idx, value in enumerate(row_data, 1):
                cell = worksheet.cell(row=current_row, column=col_idx, value=value)
                cell.border = ExcelFormatter.DATA_BORDER

                # Appliquer format de date si nécessaire
                if isinstance(value, datetime):
                    cell.number_format = excel_date_format

            current_row += 1

        return current_row

    @staticmethod
    def adjust_column_widths(worksheet, min_width: int = 10, max_width: int = 50):
        """
        Ajuste automatiquement la largeur des colonnes

        Args:
            worksheet: Feuille de calcul openpyxl
            min_width: Largeur minimale
            max_width: Largeur maximale
        """
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)

            for cell in column:
                try:
                    if cell.value is not None:
                        cell_length = len(str(cell.value))
                        if cell_length > max_length:
                            max_length = cell_length
                except (TypeError, AttributeError, ValueError):
                    # Ignorer les erreurs de conversion
                    pass

            # Ajuster la largeur avec min/max
            adjusted_width = max(min_width, min(max_length + 2, max_width))
            worksheet.column_dimensions[column_letter].width = adjusted_width

    @staticmethod
    def freeze_header(worksheet, freeze_row: int):
        """
        Fige les volets à la ligne spécifiée

        Args:
            worksheet: Feuille de calcul openpyxl
            freeze_row: Ligne où figer (les lignes au-dessus seront figées)
        """
        if freeze_row > 1:
            worksheet.freeze_panes = worksheet.cell(row=freeze_row, column=1)

    @staticmethod
    def apply_alternating_rows(worksheet, start_row: int, end_row: int,
                               color: str = "F2F2F2"):
        """
        Applique une couleur alternée aux lignes de données

        Args:
            worksheet: Feuille de calcul openpyxl
            start_row: Ligne de début
            end_row: Ligne de fin
            color: Couleur hexadécimale (sans #)
        """
        fill = PatternFill(start_color=color, end_color=color, fill_type='solid')

        for row_idx in range(start_row, end_row + 1):
            if row_idx % 2 == 0:  # Lignes paires
                for cell in worksheet[row_idx]:
                    if cell.fill.fill_type is None:
                        cell.fill = fill
