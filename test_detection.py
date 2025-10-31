"""
Script de test pour valider les détecteurs de structure
Auteur: GOUNOU N'GOBI Chabi Zimé
"""

import sys
from pathlib import Path

# Configuration encodage pour Windows
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

print("=" * 80)
print("TEST DES DETECTEURS - ExcelCompiler v3.2")
print("=" * 80)
print()

# Import des détecteurs
try:
    from core.detection import (
        HybridDetector,
        BorderDetector,
        DensityDetector,
        PatternDetector
    )
    print("[OK] Import des detecteurs reussi")
    print()
except Exception as e:
    print(f"[ERREUR] Erreur import detecteurs: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Trouver les fichiers de test
sample_dir = Path("tests/test_data/sample_files")

if not sample_dir.exists():
    print(f"[ERREUR] Dossier {sample_dir} n'existe pas")
    sys.exit(1)

# Lister les fichiers Excel
excel_files = list(sample_dir.glob("*.xlsx")) + list(sample_dir.glob("*.xls"))

if not excel_files:
    print(f"[ERREUR] Aucun fichier Excel trouve dans {sample_dir}")
    sys.exit(1)

print(f"Fichiers trouves: {len(excel_files)}")
for f in excel_files:
    print(f"  - {f.name}")
print()

# Test 1: Tester un fichier avec tous les détecteurs
print("=" * 80)
print("TEST 1: Detection individuelle sur 1 fichier")
print("=" * 80)
print()

test_file = excel_files[0]
print(f"Fichier de test: {test_file.name}")
print("-" * 80)

# BorderDetector
print("\n[BORDER] BorderDetector")
try:
    border_detector = BorderDetector()
    border_result = border_detector.detect(str(test_file))
    print(border_result.get_summary())
    print()
except Exception as e:
    print(f"[ERREUR] {e}")
    import traceback
    traceback.print_exc()

# DensityDetector
print("\n[DENSITY] DensityDetector")
try:
    density_detector = DensityDetector()
    density_result = density_detector.detect(str(test_file))
    print(density_result.get_summary())
    print()
except Exception as e:
    print(f"[ERREUR] {e}")
    import traceback
    traceback.print_exc()

# PatternDetector
print("\n[PATTERN] PatternDetector")
try:
    pattern_detector = PatternDetector()
    pattern_result = pattern_detector.detect(str(test_file))
    print(pattern_result.get_summary())
    print()
except Exception as e:
    print(f"[ERREUR] {e}")
    import traceback
    traceback.print_exc()

# HybridDetector
print("\n[HYBRID] HybridDetector")
try:
    hybrid_detector = HybridDetector()
    hybrid_result = hybrid_detector.detect(str(test_file))
    print(hybrid_result.get_summary())
    print()
except Exception as e:
    print(f"[ERREUR] {e}")
    import traceback
    traceback.print_exc()

# Test 2: Detection batch avec validation croisée
print("=" * 80)
print("TEST 2: Detection batch avec validation croisee")
print("=" * 80)
print()

# Limiter à 5 fichiers pour le test
test_files = excel_files[:5]
print(f"Fichiers a tester: {len(test_files)}")
for f in test_files:
    print(f"  - {f.name}")
print()

try:
    hybrid_detector = HybridDetector(enable_cross_validation=True)
    batch_results = hybrid_detector.detect_batch([str(f) for f in test_files])

    print("Resultats:")
    print("-" * 80)

    for result in batch_results:
        print(f"\nFichier: {Path(result.file_path).name}")
        print(f"  Methode: {result.detection_method}")
        print(f"  Confiance: {result.confidence:.0%}")
        print(f"  En-tetes: ligne {result.header_start_row} ({result.header_rows} lignes)")
        print(f"  Donnees: lignes {result.data_start_row}-{result.data_end_row}")

        if result.cross_validation_score is not None:
            print(f"  Validation croisee: {result.cross_validation_score:.0%}")

        if result.warning:
            print(f"  [ATTENTION] {result.warning}")

        # Afficher les en-têtes détectés
        print(f"  En-tetes detectes ({len(result.detected_headers)}):")
        for i, header in enumerate(result.detected_headers[:10], 1):  # Max 10
            print(f"    {i}. {header}")
        if len(result.detected_headers) > 10:
            print(f"    ... et {len(result.detected_headers) - 10} autres")

    print()

except Exception as e:
    print(f"[ERREUR] {e}")
    import traceback
    traceback.print_exc()

# Test 3: Statistiques globales
print("=" * 80)
print("TEST 3: Statistiques globales")
print("=" * 80)
print()

try:
    # Détecter tous les fichiers
    print(f"Detection de {len(excel_files)} fichiers...")
    hybrid_detector = HybridDetector(enable_cross_validation=True)
    all_results = hybrid_detector.detect_batch([str(f) for f in excel_files])

    # Statistiques
    successful = [r for r in all_results if r.confidence > 0.5]
    low_confidence = [r for r in all_results if 0.3 < r.confidence <= 0.5]
    failed = [r for r in all_results if r.confidence <= 0.3]

    print(f"\nResultats:")
    print(f"  [OK] Reussis (confiance > 50%): {len(successful)}")
    print(f"  [ATTENTION] Confiance faible (30-50%): {len(low_confidence)}")
    print(f"  [ERREUR] Echecs (confiance < 30%): {len(failed)}")
    print()

    # Confiance moyenne
    if all_results:
        avg_confidence = sum(r.confidence for r in all_results) / len(all_results)
        print(f"  Confiance moyenne: {avg_confidence:.0%}")

    # Methodes utilisées
    methods = {}
    for r in all_results:
        method = r.detection_method
        methods[method] = methods.get(method, 0) + 1

    print(f"\n  Methodes utilisees:")
    for method, count in methods.items():
        print(f"    - {method}: {count} fichiers")

    # Fichiers avec warnings
    with_warnings = [r for r in all_results if r.warning]
    if with_warnings:
        print(f"\n  [ATTENTION] {len(with_warnings)} fichiers avec avertissements:")
        for r in with_warnings:
            print(f"    - {Path(r.file_path).name}: {r.warning}")

    print()

except Exception as e:
    print(f"[ERREUR] {e}")
    import traceback
    traceback.print_exc()

# Résumé final
print("=" * 80)
print("FIN DES TESTS")
print("=" * 80)
print()
