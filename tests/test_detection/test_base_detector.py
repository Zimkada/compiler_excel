"""
Tests unitaires pour BaseDetector et DetectionResult
"""

import pytest
from pathlib import Path
from core.detection import BaseDetector, DetectionResult


class TestDetectionResult:
    """Tests pour la classe DetectionResult"""

    def test_creation_basic(self):
        """Test création basique d'un DetectionResult"""
        result = DetectionResult(
            file_path="test.xlsx",
            header_start_row=1,
            header_rows=1,
            data_start_row=2,
            data_end_row=10
        )

        assert result.file_path == "test.xlsx"
        assert result.header_start_row == 1
        assert result.header_rows == 1
        assert result.data_start_row == 2
        assert result.data_end_row == 10

    def test_auto_calculation_leading_content(self):
        """Test calcul automatique du contenu leading"""
        result = DetectionResult(
            file_path="test.xlsx",
            header_start_row=5,  # En-têtes commencent ligne 5
            header_rows=1
        )

        # __post_init__ devrait détecter leading content
        assert result.has_leading_content == True
        assert result.leading_end_row == 4

    def test_auto_calculation_data_start(self):
        """Test calcul automatique de data_start_row"""
        result = DetectionResult(
            file_path="test.xlsx",
            header_start_row=3,
            header_rows=2
        )

        # data_start_row devrait être calculé automatiquement
        assert result.data_start_row == 5  # 3 + 2

    def test_is_reliable(self):
        """Test méthode is_reliable"""
        result = DetectionResult(
            file_path="test.xlsx",
            confidence=0.80
        )

        assert result.is_reliable(threshold=0.75) == True
        assert result.is_reliable(threshold=0.85) == False

    def test_to_dict(self):
        """Test conversion en dictionnaire"""
        result = DetectionResult(
            file_path="test.xlsx",
            header_start_row=1,
            confidence=0.95,
            detection_method="border"
        )

        data = result.to_dict()

        assert isinstance(data, dict)
        assert data['file_path'] == "test.xlsx"
        assert data['confidence'] == 0.95
        assert data['detection_method'] == "border"

    def test_get_summary(self):
        """Test génération de résumé"""
        result = DetectionResult(
            file_path="test.xlsx",
            header_start_row=2,
            header_rows=1,
            data_end_row=20,
            confidence=0.85,
            detection_method="hybrid"
        )

        summary = result.get_summary()

        assert "test.xlsx" in summary
        assert "85%" in summary
        assert "hybrid" in summary


class TestBaseDetector:
    """Tests pour la classe BaseDetector"""

    def test_cannot_instantiate_abstract(self):
        """Test qu'on ne peut pas instancier BaseDetector directement"""
        with pytest.raises(TypeError):
            BaseDetector()

    def test_calculate_row_density(self):
        """Test calcul de densité de ligne"""
        import pandas as pd

        # Créer un détecteur concret pour tester
        from core.detection import DensityDetector

        detector = DensityDetector()

        # DataFrame de test
        df = pd.DataFrame([
            ['A', 'B', 'C', 'D', 'E'],  # 100% rempli
            ['A', None, 'C', None, 'E'],  # 60% rempli
            [None, None, None, None, None],  # 0% rempli
        ])

        assert detector.calculate_row_density(df, 0) == 1.0
        assert detector.calculate_row_density(df, 1) == 0.6
        assert detector.calculate_row_density(df, 2) == 0.0

    def test_is_likely_header_row(self):
        """Test détection de ligne d'en-tête"""
        import pandas as pd
        from core.detection import DensityDetector

        detector = DensityDetector()

        # DataFrame de test
        df = pd.DataFrame([
            ['Nom', 'Prénom', 'Age', 'Ville'],  # En-tête (texte)
            ['Jean', 'Dupont', 25, 'Paris'],  # Données (mixte)
            [1, 2, 3, 4],  # Données (nombres)
        ])

        assert detector.is_likely_header_row(df, 0) == True  # En-tête
        assert detector.is_likely_header_row(df, 2) == False  # Nombres

    def test_extract_headers_single_line(self):
        """Test extraction d'en-têtes sur une ligne"""
        import pandas as pd
        from core.detection import DensityDetector

        detector = DensityDetector()

        df = pd.DataFrame([
            ['Nom', 'Prénom', 'Age'],
            ['Jean', 'Dupont', 25]
        ])

        headers = detector.extract_headers(df, header_start_row=1, header_rows=1)

        assert headers == ['Nom', 'Prénom', 'Age']

    def test_extract_headers_multi_line(self):
        """Test extraction d'en-têtes multi-lignes"""
        import pandas as pd
        from core.detection import DensityDetector

        detector = DensityDetector()

        df = pd.DataFrame([
            ['Informations', 'Informations', 'Détails'],
            ['Nom', 'Prénom', 'Age'],
            ['Jean', 'Dupont', 25]
        ])

        headers = detector.extract_headers(df, header_start_row=1, header_rows=2)

        assert len(headers) == 3
        assert 'Informations - Nom' in headers[0]
        assert 'Informations - Prénom' in headers[1]
        assert 'Détails - Age' in headers[2]
