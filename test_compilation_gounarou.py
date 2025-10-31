"""
Test de compilation sur fichiers CEG-GOUNAROU (même canevas, sans bordures)
Auteur: GOUNOU N'GOBI Chabi Zimé
"""

import sys
from pathlib import Path

# Configuration encodage
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

print("=" * 80)
print("TEST COMPILATION - CEG-GOUNAROU SEM 2 (Meme canevas)")
print("=" * 80)
print()

from core.compilation import (
    ExcelCompiler,
    CompilationOptions,
    FilenameOption,
    OutputFormat
)

# Tous les fichiers CEG-GOUNAROU
sample_dir = Path("tests/test_data/sample_files")
gounarou_files = sorted(sample_dir.glob("*CEG-GOUNAROU*.xlsx"))

print(f"Fichiers trouves: {len(gounarou_files)}")
for f in gounarou_files:
    print(f"  - {f.name}")
print()

# Test 1: Voir un fichier exemple
print("[TEST 1] Structure d'un fichier exemple")
print("-" * 80)
if gounarou_files:
    import pandas as pd
    test_file = gounarou_files[0]
    df = pd.read_excel(test_file, header=None)

    print(f"Fichier: {test_file.name}")
    print(f"Taille: {len(df)} lignes, {len(df.columns)} colonnes")
    print()
    print("Premieres 10 lignes:")
    for i in range(min(10, len(df))):
        row_content = str(list(df.iloc[i]))[:100]
        print(f"  Ligne {i+1}: {row_content}")
    print()

# Test 2: Détection automatique
print("[TEST 2] Detection automatique sur 3 premiers fichiers")
print("-" * 80)

from core.detection import HybridDetector

detector = HybridDetector(enable_cross_validation=True)
test_files_paths = [str(f) for f in gounarou_files[:3]]

results = detector.detect_batch(test_files_paths)

for result in results:
    print(f"\n{Path(result.file_path).name}:")
    print(f"  Method: {result.detection_method}")
    print(f"  Confidence: {result.confidence:.0%}")
    print(f"  Headers: ligne {result.header_start_row} ({result.header_rows} lignes)")
    print(f"  Data: lignes {result.data_start_row} a {result.data_end_row}")
    print(f"  => {result.data_end_row - result.data_start_row + 1} lignes de donnees")
    if result.cross_validation_score:
        print(f"  Cross-validation: {result.cross_validation_score:.0%}")

print()

# Test 3: Compilation COMPLETE de tous les fichiers
print("[TEST 3] Compilation de TOUS les fichiers CEG-GOUNAROU")
print("-" * 80)

options = CompilationOptions(
    auto_detect_structure=True,
    enable_cross_validation=True,
    remove_empty_rows=True,
    filename_option=FilenameOption.WITHOUT_EXTENSION
)

compiler = ExcelCompiler(options)

output_file = "compilation_gounarou_complet.xlsx"

file_paths = [str(f) for f in gounarou_files]

print(f"Compilation de {len(file_paths)} fichiers...")
print()

result = compiler.compile_files(
    file_paths,
    output_file,
    OutputFormat.XLSX
)

# Afficher resultat
print(result.get_summary())
print()

if result.success and Path(output_file).exists():
    # Verifier le fichier de sortie
    import pandas as pd
    df_output = pd.read_excel(output_file, header=None)

    print("[VERIFICATION FICHIER DE SORTIE]")
    print("-" * 80)
    print(f"Taille: {len(df_output)} lignes, {len(df_output.columns)} colonnes")
    print()
    print("Premieres 5 lignes:")
    for i in range(min(5, len(df_output))):
        print(f"  Ligne {i+1}: {list(df_output.iloc[i])[:5]}...")
    print()
    print("Dernieres 5 lignes:")
    for i in range(max(0, len(df_output)-5), len(df_output)):
        print(f"  Ligne {i+1}: {list(df_output.iloc[i])[:5]}...")

print()
print("=" * 80)
print("FIN DU TEST")
print("=" * 80)
