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


class TestCrossValidationAdjustment:
    """La validation croisée ne doit jamais dégrader ni alarmer, seulement
    renforcer la confiance quand les fichiers se ressemblent fortement."""

    def _result(self, confidence=0.8):
        from core.detection import DetectionResult
        return DetectionResult(file_path="f.xlsx", confidence=confidence)

    def test_low_similarity_does_not_degrade_confidence(self):
        detector = HybridDetector()
        r = self._result(confidence=0.80)
        adjusted = detector._adjust_confidence_with_cross_validation(r, 0.10)
        assert adjusted.confidence == 0.80  # inchangée

    def test_low_similarity_does_not_set_warning(self):
        detector = HybridDetector()
        r = self._result()
        adjusted = detector._adjust_confidence_with_cross_validation(r, 0.10)
        assert adjusted.warning is None

    def test_medium_similarity_no_warning(self):
        detector = HybridDetector()
        r = self._result()
        adjusted = detector._adjust_confidence_with_cross_validation(r, 0.45)
        assert adjusted.warning is None
        assert adjusted.confidence == 0.80

    def test_high_similarity_gives_bonus(self):
        detector = HybridDetector()
        r = self._result(confidence=0.80)
        adjusted = detector._adjust_confidence_with_cross_validation(r, 0.90)
        assert adjusted.confidence > 0.80
        assert adjusted.confidence <= 1.0

    def test_score_recorded_in_debug(self):
        detector = HybridDetector()
        r = self._result()
        adjusted = detector._adjust_confidence_with_cross_validation(r, 0.33)
        assert adjusted.debug_info.get('cross_validation_score') == 0.33

    def test_batch_on_dissimilar_files_emits_no_warning(self):
        """Régression: des fichiers réels dissemblables mais correctement
        détectés ne doivent produire aucun avertissement *de validation
        croisée*. Le warning d'élagage de colonnes parasites est légitime
        (il alerte l'utilisateur) et donc toléré ici."""
        detector = HybridDetector(enable_cross_validation=True)
        sample_dir = Path("tests/test_data/sample_files")
        files = [str(f) for f in sample_dir.glob("CEG*.xlsx")]
        if len(files) < 2:
            pytest.skip("Pas assez de fichiers réels")
        results = detector.detect_batch(files)
        for r in results:
            # Un éventuel warning ne doit provenir QUE de l'élagage de colonnes
            # parasites, jamais de la validation croisée.
            if r.warning is not None:
                assert r.debug_info.get('pruned_phantom_columns'), (
                    f"Warning inattendu (non lié à l'élagage): {r.warning!r}"
                )


class TestVoterRestrictionOnDisagreement:
    """En cas de DÉSACCORD sur la ligne d'en-tête, seuls les détecteurs ayant
    trouvé la ligne retenue votent sur le reste de la structure.

    Bug corrigé : la règle « un détecteur qui s'est trompé de ligne n'a pas voix
    au chapitre » n'était appliquée que dans la branche « accord ». En cas de
    désaccord, TOUS les détecteurs votaient sur header_rows. Sur un fichier réel
    où density disait (start=3, rows=2) et pattern (start=4, rows=1), la médiane
    de [2, 1] donnait rows=1 : l'en-tête à deux niveaux n'était plus aplati, ses
    libellés divergeaient du reste du lot et l'aligneur créait des colonnes en
    double après « Fichier source ».
    """

    def _result(self, method, start, rows, confidence, data_end=0):
        from core.detection.base_detector import DetectionResult
        return DetectionResult(
            file_path="f.xlsx", header_start_row=start, header_rows=rows,
            data_start_row=start + rows, data_end_row=data_end,
            confidence=confidence, detection_method=method,
        )

    def test_disagreeing_detector_does_not_vote_on_header_rows(self):
        detector = HybridDetector()
        results = [
            self._result("density", start=3, rows=2, confidence=0.85),
            self._result("pattern", start=4, rows=1, confidence=0.83),
        ]
        combined = detector._create_combined_result(
            results, {"border": 0.40, "density": 0.30, "pattern": 0.30}
        )
        # La médiane des starts retient 3 ; seul density l'a proposée.
        assert combined.header_start_row == 3
        assert combined.header_rows == 2
        assert combined.data_start_row == 5

    def test_data_end_still_uses_all_detectors(self):
        """Non-régression : data_end prend le MAX de TOUS les détecteurs.
        Le restreindre aux voteurs tronquerait des lignes de données."""
        detector = HybridDetector()
        results = [
            self._result("density", start=3, rows=2, confidence=0.85, data_end=5),
            self._result("pattern", start=4, rows=1, confidence=0.83, data_end=40),
        ]
        combined = detector._create_combined_result(
            results, {"border": 0.40, "density": 0.30, "pattern": 0.30}
        )
        assert combined.data_end_row == 40

    def test_agreement_branch_unchanged(self):
        """Non-régression : quand deux détecteurs s'accordent sur la ligne,
        le comportement existant (bonus de confiance) est préservé."""
        detector = HybridDetector()
        results = [
            self._result("density", start=3, rows=2, confidence=0.54),
            self._result("pattern", start=3, rows=2, confidence=0.73),
            self._result("border", start=7, rows=1, confidence=0.40),
        ]
        combined = detector._create_combined_result(
            results, {"border": 0.40, "density": 0.30, "pattern": 0.30}
        )
        assert combined.header_start_row == 3
        assert combined.header_rows == 2
        assert combined.confidence > 0.73  # bonus d'accord appliqué
