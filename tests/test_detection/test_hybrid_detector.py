"""
Tests unitaires pour HybridDetector
"""

import pytest
from pathlib import Path
from core.detection import HybridDetector


class TestHybridDetector:
    """Tests pour HybridDetector"""

    def test_initialization(self):
        """Test initialisation du détecteur"""
        detector = HybridDetector()

        assert detector.confidence_threshold == 0.65
        assert detector.enable_cross_validation == True
        assert detector.border_detector is not None
        assert detector.density_detector is not None
        assert detector.pattern_detector is not None

    def test_initialization_custom_params(self):
        """Test initialisation avec paramètres personnalisés"""
        detector = HybridDetector(
            confidence_threshold=0.75,
            enable_cross_validation=False
        )

        assert detector.confidence_threshold == 0.75
        assert detector.enable_cross_validation == False

    def test_detect_single_file(self):
        """Test détection sur un fichier réel"""
        detector = HybridDetector()
        sample_dir = Path("tests/test_data/sample_files")

        # Trouver un fichier de test
        excel_files = list(sample_dir.glob("*.xlsx"))
        if not excel_files:
            pytest.skip("Aucun fichier Excel de test disponible")

        test_file = excel_files[0]
        result = detector.detect(str(test_file))

        # Vérifications basiques
        assert result is not None
        assert result.file_path == str(test_file)
        assert result.confidence >= 0
        assert result.detection_method in ['border', 'density', 'pattern', 'hybrid']
        assert result.header_start_row >= 1
        assert result.header_rows >= 1

    def test_detect_batch_with_cross_validation(self):
        """Test détection batch avec validation croisée"""
        detector = HybridDetector(enable_cross_validation=True)
        sample_dir = Path("tests/test_data/sample_files")

        excel_files = list(sample_dir.glob("*.xlsx"))[:3]  # Limiter à 3 fichiers
        if len(excel_files) < 2:
            pytest.skip("Pas assez de fichiers pour test batch")

        file_paths = [str(f) for f in excel_files]
        results = detector.detect_batch(file_paths)

        # Vérifications
        assert len(results) == len(file_paths)

        # Tous les résultats doivent avoir un score de validation croisée
        for result in results:
            assert result.cross_validation_score is not None
            assert 0.0 <= result.cross_validation_score <= 1.0

    def test_cache_cleared(self):
        """Test vidage du cache"""
        detector = HybridDetector()
        sample_dir = Path("tests/test_data/sample_files")

        excel_files = list(sample_dir.glob("*.xlsx"))[:2]
        if len(excel_files) < 2:
            pytest.skip("Pas assez de fichiers")

        # Détecter quelques fichiers
        for f in excel_files:
            detector.detect(str(f))

        assert len(detector._detection_cache) > 0

        # Vider le cache
        detector.clear_cache()
        assert len(detector._detection_cache) == 0

    def test_header_similarity_calculation(self):
        """Test calcul de similarité des en-têtes"""
        detector = HybridDetector()

        headers1 = ['Nom', 'Prénom', 'Age', 'Ville']
        headers2 = ['Nom', 'Prénom', 'Téléphone', 'Email']

        # Extraire les mots
        words1 = detector._extract_header_words(headers1)
        words2 = detector._extract_header_words(headers2)

        # Il devrait y avoir des mots en commun (nom, prenom)
        assert len(words1 & words2) > 0

        # Calculer similarité
        similarity = detector._calculate_header_similarity(headers1, headers2)
        assert 0.0 < similarity < 1.0  # Similarité partielle

    def test_header_similarity_identical(self):
        """Test similarité avec en-têtes identiques"""
        detector = HybridDetector()

        headers = ['Nom', 'Prénom', 'Age']

        similarity = detector._calculate_header_similarity(headers, headers)
        assert similarity == 1.0  # Identiques

    def test_header_similarity_different(self):
        """Test similarité avec en-têtes totalement différents"""
        detector = HybridDetector()

        headers1 = ['Nom', 'Prénom', 'Age']
        headers2 = ['Montant', 'Date', 'Référence']

        similarity = detector._calculate_header_similarity(headers1, headers2)
        # Devrait être faible (peut-être > 0 si mots courts ignorés)
        assert similarity < 0.5

    def test_extract_header_words(self):
        """Test extraction de mots des en-têtes"""
        detector = HybridDetector()

        headers = [
            'Nom de famille',
            'Prénom',
            'Age de la personne',
            'Code postal'
        ]

        words = detector._extract_header_words(headers)

        # Vérifier que les mots significatifs sont extraits
        assert 'nom' in words
        assert 'prenom' in words or 'prénom' in words
        assert 'age' in words
        assert 'code' in words
        assert 'postal' in words

        # Les stop words ne devraient pas être inclus
        assert 'de' not in words
        assert 'la' not in words
