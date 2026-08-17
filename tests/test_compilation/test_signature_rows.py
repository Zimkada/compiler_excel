"""
Tests des lignes de signature en pied de tableau.

Un formulaire administratif rempli se termine souvent par « Fait à X, le … »
puis la qualité et le nom du signataire. Ces lignes sont SOUS le tableau, pas
dedans : compilées, elles produisent des lignes sans établissement dont le
texte atterrit dans des colonnes de chiffres (constaté sur une compilation
réelle de 47 fichiers : 27 lignes parasites pour 47 lignes de données).

La détection exige CUMULATIVEMENT une identité vide, aucune valeur numérique et
un marqueur de signature. On vérifie surtout l'absence de FAUX POSITIFS : une
ligne de données ne doit JAMAIS être écartée.
"""

import openpyxl

from core.compilation import ExcelCompiler, CompilationOptions
from core.compilation.compilation_models import FilenameOption
from core.compilation.subtotal_detector import is_signature_row


def _make_xlsx(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for ri, row in enumerate(rows, 1):
        for ci, val in enumerate(row, 1):
            if val is not None:
                ws.cell(ri, ci, val)
    wb.save(path)
    return str(path)


class TestSignatureUnit:
    """Cas relevés tels quels sur les fichiers réels."""

    def test_fait_a_le(self):
        assert is_signature_row([None, None, "Fait à Founougo, le 13/07/2026", None])

    def test_le_directeur(self):
        assert is_signature_row([None, None, None, "Le Directeur,"])

    def test_chef_etablissement(self):
        assert is_signature_row([None, None, None, "le chef d'établissement"])

    def test_signataire_name_alone_is_kept(self):
        """Le nom du signataire SEUL est conservé : aucun marqueur ne permet de
        le distinguer d'une donnée.

        Une version antérieure écartait toute ligne sans identité ne portant
        qu'une seule cellule de texte. Cette tolérance supprimait de vraies
        données (« NEANT » d'un établissement, « Aucun élève inscrit »…). Garder
        une mention de trop est anodin ; perdre une donnée ne l'est pas.
        """
        assert not is_signature_row([None, None, None, "Aoudou NAMATA", None])

    def test_same_mention_repeated_across_columns(self):
        assert is_signature_row(
            [None, "Fait à Guéné le 14/07/2026", "Fait à Guéné le 14/07/2026", None]
        )


class TestNoFalsePositive:
    """Le garde-fou : une vraie ligne de données n'est jamais écartée."""

    def test_neant_without_identity_is_data(self):
        """Régression (revue) : un établissement ayant saisi « NEANT » sans
        remplir sa colonne d'identité voyait sa ligne supprimée."""
        assert not is_signature_row([None, "NEANT", None, None])

    def test_free_text_without_identity_is_data(self):
        assert not is_signature_row([None, None, "Aucun élève inscrit", None])

    def test_total_row_is_not_a_signature(self):
        """Régression (revue) : un « TOTAL GENERAL » sans identité était compté
        comme signature. Il relève de classify_row et de drop_subtotal_rows —
        sinon l'option « conserver les totaux » était contournée."""
        assert not is_signature_row([None, "TOTAL GENERAL", None])

    def test_subtotal_row_is_not_a_signature(self):
        assert not is_signature_row([None, "ENSEMBLE COMMUNE", None])

    def test_data_row_with_identity(self):
        assert not is_signature_row(["CEG ARBONGA", 189.0, 189.0, 72.0, 68.0])

    def test_row_without_identity_but_with_numbers(self):
        """Identité vide MAIS chiffres présents -> ligne de données (continuation).
        C'est la condition qui protège les tableaux à identité fusionnée."""
        assert not is_signature_row([None, 12, 8, 5, 3])

    def test_number_stored_as_text(self):
        """Un nombre saisi en texte compte comme valeur numérique."""
        assert not is_signature_row([None, None, None, "45", None])

    def test_french_decimal_as_text(self):
        assert not is_signature_row([None, None, "1 234,5", None])

    def test_identity_present_without_numbers(self):
        """Établissement renseigné mais aucune donnée chiffrée : ligne de
        données incomplète, pas une signature."""
        assert not is_signature_row(["CEG X", None, None, None])

    def test_empty_row_is_not_signature(self):
        """Une ligne vide relève du filtre de lignes vides, pas d'ici."""
        assert not is_signature_row([None, None, None, None])

    def test_establishment_named_like_a_signature_kept(self):
        """Faux positif redouté : l'identité est renseignée, donc conservée
        quoi que contiennent les autres cellules."""
        assert not is_signature_row(["CEG LE DIRECTEUR", 10, 5, 3, 2])


class TestSignatureRowsInCompilation:
    """Bout en bout : le bloc de signature n'atteint jamais la sortie."""

    def _compile(self, tmp_path, drop_signature_rows=True):
        src = _make_xlsx(tmp_path / "ceg.xlsx", [
            ["Etablissement", "Total", "Survivants"],
            ["CEG ARBONGA", 189, 189],
            ["CEG BAGOU", 72, 68],
            [None, None, None],
            [None, None, "Fait à Kandi, le 13/07/2026"],
            [None, None, "Le Directeur"],
            [None, None, "Zenabou BANI SEIDOU"],
        ])
        out = tmp_path / "out.xlsx"
        opts = CompilationOptions(
            use_reference_mode=False,
            auto_detect_structure=False,
            manual_header_start_row=1,
            manual_header_rows=1,
            filename_option=FilenameOption.NONE,
            drop_signature_rows=drop_signature_rows,
        )
        result = ExcelCompiler(opts).compile_files([src], str(out))
        assert result.successful_files == 1
        wb = openpyxl.load_workbook(out)
        ws = wb.active
        rows = [
            [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            for r in range(1, ws.max_row + 1)
        ]
        wb.close()
        return rows, result

    def test_signature_keyword_inside_the_table_is_kept(self, tmp_path):
        """Régression (revue) : le mot-clé seul ne suffit pas — il faut être
        APRÈS une rupture. Dans une liste de personnel, « Le Censeur » est une
        donnée, pas une signature, et la ligne doit survivre."""
        src = _make_xlsx(tmp_path / "personnel.xlsx", [
            ["Etablissement", "Fonction", "Adjoint"],
            ["CEG A", "Le Directeur", "Paul"],
            [None, "Le Censeur", "Marie"],
            ["CEG B", "Le Proviseur", "Jean"],
            [None, None, None],
            [None, None, "Fait à Kandi, le 13/07/2026"],
        ])
        out = tmp_path / "out.xlsx"
        opts = CompilationOptions(
            use_reference_mode=False, auto_detect_structure=False,
            manual_header_start_row=1, manual_header_rows=1,
            filename_option=FilenameOption.NONE, drop_signature_rows=True,
        )
        result = ExcelCompiler(opts).compile_files([src], str(out))
        assert result.successful_files == 1

        wb = openpyxl.load_workbook(out)
        ws = wb.active
        rows = [
            [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            for r in range(1, ws.max_row + 1)
        ]
        wb.close()

        flat = [str(v) for r in rows for v in r if v is not None]
        # Les lignes DANS le tableau sont conservées…
        assert "Le Censeur" in flat
        assert "Marie" in flat
        assert "Le Proviseur" in flat
        # …et le pied de page, lui, est bien écarté.
        assert not any("Fait à" in v for v in flat)

    def test_signature_block_excluded(self, tmp_path):
        """Les lignes PORTANT un marqueur sont écartées.

        Le nom du signataire seul (« Zenabou BANI SEIDOU »), lui, est conservé :
        rien ne le distingue d'une donnée. Compromis assumé — garder une mention
        de trop est anodin, supprimer une donnée ne l'est pas (cf.
        test_signataire_name_alone_is_kept).
        """
        rows, _ = self._compile(tmp_path)
        assert rows[0] == ["Etablissement", "Total", "Survivants"]
        flat = [str(v) for r in rows for v in r if v is not None]
        assert not any("Directeur" in v or "Fait à" in v for v in flat)
        assert "CEG ARBONGA" in flat and "CEG BAGOU" in flat

    def test_exclusion_is_reported(self, tmp_path):
        """Jamais silencieux : le nombre de lignes écartées est signalé."""
        _, result = self._compile(tmp_path)
        assert any("signature" in w for w in result.warnings)

    def test_option_disabled_keeps_rows(self, tmp_path):
        """drop_signature_rows=False conserve le comportement historique."""
        rows, _ = self._compile(tmp_path, drop_signature_rows=False)
        flat = [str(v) for r in rows for v in r if v is not None]
        assert any("Directeur" in v for v in flat)


class TestShiftedFileCompilation:
    """Bout en bout : un fichier saisi décalé s'empile au bon endroit."""

    def test_shifted_file_aligns_with_reference(self, tmp_path):
        ref = _make_xlsx(tmp_path / "a_ref.xlsx", [
            ["Etablissement", "Total", "Survivants"],
            ["CEG ARBONGA", 189, 189],
        ])
        # Même formulaire, saisi à partir de la colonne C.
        shifted = _make_xlsx(tmp_path / "b_shifted.xlsx", [
            [None, None, "Etablissement", "Total", "Survivants"],
            [None, None, "CEG MADINA", 30, 20],
        ])
        out = tmp_path / "out.xlsx"
        opts = CompilationOptions(
            use_reference_mode=False,
            auto_detect_structure=False,
            manual_header_start_row=1,
            manual_header_rows=1,
            filename_option=FilenameOption.NONE,
        )
        result = ExcelCompiler(opts).compile_files([ref, shifted], str(out))
        assert result.successful_files == 2

        wb = openpyxl.load_workbook(out)
        ws = wb.active
        rows = [
            [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            for r in range(1, ws.max_row + 1)
        ]
        wb.close()

        # Aucune colonne fantôme ajoutée à droite, et MADINA bien aligné.
        assert ws.max_column == 3
        assert rows[0] == ["Etablissement", "Total", "Survivants"]
        assert ["CEG MADINA", 30, 20] in rows
