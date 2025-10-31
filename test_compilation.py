"""
Script de test pour valider la compilation avec détection automatique
Auteur: GOUNOU N'GOBI Chabi Zimé
"""

import sys
from pathlib import Path

# Configuration encodage pour Windows
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

print("=" * 80)
print("TEST DE COMPILATION - ExcelCompiler v3.2")
print("=" * 80)
print()

# Test 1: Import des modules
print("[TEST 1] Import des modules")
print("-" * 80)
try:
    from core.compilation import (
        ExcelCompiler,
        CompilationOptions,
        FilenameOption,
        DateFormat,
        OutputFormat
    )
    print("[OK] Tous les modules importes avec succes")
    print()
except Exception as e:
    print(f"[ERREUR] Erreur import: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Test 2: Création d'options par défaut
print("[TEST 2] Creation d'options par defaut")
print("-" * 80)
try:
    options = CompilationOptions()
    print(f"[OK] Options creees")
    print(f"   - Auto detection: {options.auto_detect_structure}")
    print(f"   - Cross validation: {options.enable_cross_validation}")
    print(f"   - Confidence threshold: {options.detection_confidence_threshold}")
    print(f"   - Remove empty rows: {options.remove_empty_rows}")
    print()
except Exception as e:
    print(f"[ERREUR] {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Test 3: Création du compilateur
print("[TEST 3] Creation du compilateur")
print("-" * 80)
try:
    compiler = ExcelCompiler(options)
    print(f"[OK] Compilateur cree")
    print(f"   - Detecteur: {compiler.detector is not None}")
    print()
except Exception as e:
    print(f"[ERREUR] {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Test 4: Compilation réelle sur fichiers exemples
print("[TEST 4] Compilation de fichiers exemples")
print("-" * 80)

sample_dir = Path("tests/test_data/sample_files")
if not sample_dir.exists():
    print(f"[ATTENTION] Dossier {sample_dir} n'existe pas, test saute")
    print()
else:
    # Trouver 3 fichiers de test
    excel_files = list(sample_dir.glob("*.xlsx"))[:3]

    if len(excel_files) < 2:
        print(f"[ATTENTION] Pas assez de fichiers pour test, test saute")
        print()
    else:
        try:
            print(f"Compilation de {len(excel_files)} fichiers...")
            for f in excel_files:
                print(f"   - {f.name}")
            print()

            # Fichier de sortie
            output_file = "test_compilation_output.xlsx"

            # Options de test
            test_options = CompilationOptions(
                auto_detect_structure=True,
                enable_cross_validation=True,
                remove_empty_rows=True,
                filename_option=FilenameOption.WITHOUT_EXTENSION
            )

            # Créer compilateur
            test_compiler = ExcelCompiler(test_options)

            # Compiler
            file_paths = [str(f) for f in excel_files]
            result = test_compiler.compile_files(
                file_paths,
                output_file,
                OutputFormat.XLSX
            )

            # Afficher résultat
            print()
            print(result.get_summary())
            print()

            if result.success:
                print(f"[OK] Compilation reussie!")
                print(f"   Fichier de sortie: {result.output_file}")
                print(f"   Lignes compilees: {result.total_rows}")
                print()

                # Vérifier que le fichier existe
                if Path(output_file).exists():
                    file_size = Path(output_file).stat().st_size
                    print(f"   [OK] Fichier cree ({file_size} octets)")
                else:
                    print(f"   [ERREUR] Fichier non cree")

            else:
                print(f"[ERREUR] Compilation echouee")
                if result.warnings:
                    print("Warnings:")
                    for w in result.warnings:
                        print(f"   - {w}")

            print()

        except Exception as e:
            print(f"[ERREUR] Exception: {e}")
            import traceback
            traceback.print_exc()

# Test 5: Test sans détection automatique (mode manuel)
print("[TEST 5] Compilation en mode manuel (sans detection)")
print("-" * 80)

if sample_dir.exists() and len(excel_files) >= 2:
    try:
        # Options manuelles
        manual_options = CompilationOptions(
            auto_detect_structure=False,
            manual_header_start_row=1,
            manual_header_rows=1,
            remove_empty_rows=True
        )

        manual_compiler = ExcelCompiler(manual_options)

        output_file_manual = "test_compilation_manual.xlsx"

        result_manual = manual_compiler.compile_files(
            file_paths[:2],  # Seulement 2 fichiers
            output_file_manual,
            OutputFormat.XLSX
        )

        if result_manual.success:
            print(f"[OK] Compilation manuelle reussie")
            print(f"   Fichiers: {result_manual.successful_files}/{result_manual.total_files}")
            print(f"   Lignes: {result_manual.total_rows}")
        else:
            print(f"[ERREUR] Compilation manuelle echouee")

        print()

    except Exception as e:
        print(f"[ERREUR] {e}")
        import traceback
        traceback.print_exc()
else:
    print(f"[ATTENTION] Pas de fichiers pour test, test saute")
    print()

# Résumé final
print("=" * 80)
print("FIN DES TESTS")
print("=" * 80)
print()
