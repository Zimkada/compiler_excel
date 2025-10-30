# ExcelCompiler v3.2 - Architecture Modulaire

## 🎯 Objectifs de la v3.2

Cette version apporte une refonte majeure avec :
- ✅ Architecture modulaire maintenable
- ✅ Détection intelligente de structure (85-95% précision)
- ✅ ML local optionnel (apprentissage personnalisé)
- ✅ Détection d'anomalies
- ✅ UI moderne (Accordion + Palette Excel Native)
- ✅ PyQt6 uniquement (Windows 10/11)

## 📂 Structure du projet

```
compiler_excel/
├── v3.1_legacy/                    # Version 3.1 PyQt5 (archivée)
│   ├── compiler.py                 # Version principale PyQt5
│   ├── compiler_Win8Plus.py        # Version Win8+ PyQt5
│   └── ExcelCompiler_Win8Plus.spec # Spec PyInstaller PyQt5
│
├── config/                         # Configuration application
│   ├── __init__.py
│   ├── constants.py                # Constantes globales
│   └── app_config.py               # Gestionnaire configuration JSON
│
├── core/                           # Logique métier
│   ├── detection/                  # Détection intelligente structure
│   │   ├── base_detector.py        # Classe abstraite
│   │   ├── border_detector.py      # Détection par bordures (95%)
│   │   ├── density_detector.py     # Détection par densité (85%)
│   │   ├── pattern_detector.py     # Détection par patterns (75%)
│   │   ├── hybrid_detector.py      # Orchestrateur cascade
│   │   └── ml_local.py             # ML local optionnel
│   │
│   ├── anomalies/                  # Détection anomalies
│   │   ├── anomaly_detector.py     # Détecteur anomalies
│   │   └── anomaly_report.py       # Générateur rapport
│   │
│   ├── compilation/                # Moteur compilation
│   │   ├── compiler.py             # Compilateur principal
│   │   ├── chunked_processor.py    # Traitement par chunks
│   │   └── format_handlers.py      # Handlers formats
│   │
│   └── validation/                 # Validation données
│       ├── security_validator.py   # Validation sécurité
│       └── data_validator.py       # Validation données
│
├── ui/                             # Interface PyQt6
│   ├── main_window.py              # Fenêtre principale
│   ├── widgets/                    # Widgets personnalisés
│   │   ├── accordion_widget.py     # Widget accordion progressif
│   │   ├── file_selector.py        # Sélecteur fichiers
│   │   ├── preview_panel.py        # Panneau aperçu
│   │   └── progress_widget.py      # Indicateur progression
│   │
│   └── styles/                     # Styles et thèmes
│       ├── excel_theme.py          # Palette Excel Native (#217346)
│       └── stylesheet.py           # QSS styles
│
├── utils/                          # Utilitaires
│   ├── __init__.py
│   ├── logger.py                   # Configuration logging
│   └── file_utils.py               # Utilitaires fichiers
│
├── tests/                          # Tests unitaires
│   ├── test_detection/             # Tests détection
│   ├── test_anomalies/             # Tests anomalies
│   └── test_data/                  # Données de test
│       └── sample_files/
│
├── compiler_Win10Plus.py           # Point d'entrée v3.2 (PyQt6)
├── ExcelCompiler_Win10Plus.spec    # Spec PyInstaller PyQt6
├── requirements_win10_pyqt6.txt    # Dépendances PyQt6
└── README_V3.2.md                  # Ce fichier

```

## 🚀 Développement

### Prérequis

- Python 3.8+
- Windows 10/11
- PyQt6

### Installation environnement

```bash
# Créer environnement virtuel
python -m venv venv_v3.2

# Activer l'environnement
venv_v3.2\Scripts\activate

# Installer dépendances
pip install -r requirements_win10_pyqt6.txt
```

### Lancer l'application

```bash
python compiler_Win10Plus.py
```

### Tests

```bash
# Installer pytest
pip install pytest pytest-cov

# Lancer tous les tests
pytest tests/ -v

# Avec couverture
pytest tests/ --cov=core --cov=ui --cov-report=html
```

## 📋 Planning d'implémentation

### ✅ Phase 0 - Préparation (Complétée)
- [x] Création branche Git feature/v3.2-pyqt6
- [x] Archivage v3.1 PyQt5 dans v3.1_legacy/
- [x] Création structure modulaire
- [x] Modules config/ et utils/ de base

### 🔄 Phase 1 - Détection intelligente (En cours)
- [ ] BaseDetector et DetectionResult
- [ ] BorderDetector (détection par bordures)
- [ ] DensityDetector (détection par densité)
- [ ] PatternDetector (détection par patterns)
- [ ] HybridDetector (orchestrateur)
- [ ] Tests unitaires détection

### ⏳ Phase 2 - ML local optionnel
- [ ] PersonalMLModel
- [ ] Intégration dans HybridDetector
- [ ] UI activation ML
- [ ] Tests ML

### ⏳ Phase 3 - Détection anomalies
- [ ] AnomalyDetector
- [ ] AnomalyReport
- [ ] UI bouton "Vérifier anomalies"
- [ ] Tests anomalies

### ⏳ Phase 4 - UI moderne
- [ ] Palette Excel Native (excel_theme.py)
- [ ] Stylesheet global
- [ ] AccordionWidget
- [ ] Intégration UI

### ⏳ Phase 5 - Finalisation
- [ ] Tests complets
- [ ] Documentation
- [ ] Build PyInstaller
- [ ] Installateur Inno Setup

## 🔧 Technologies utilisées

- **Interface** : PyQt6 (Windows 10/11 uniquement)
- **Backend** : pandas, openpyxl, numpy
- **ML** : scikit-learn (optionnel)
- **Tests** : pytest
- **Build** : PyInstaller
- **Installateur** : Inno Setup

## 📝 Différences v3.1 vs v3.2

| Fonctionnalité | v3.1 | v3.2 |
|----------------|------|------|
| **Architecture** | Monolithique (12k lignes) | Modulaire |
| **Interface** | PyQt5/PyQt6 | PyQt6 uniquement |
| **Détection structure** | Manuelle | Automatique (85-95%) |
| **ML** | ❌ | ✅ Optionnel local |
| **Anomalies** | ❌ | ✅ Manuel |
| **UI** | Grid 3x3 | Accordion progressif |
| **Palette** | Standard | Excel Native (#217346) |
| **Windows** | 8.1+ | 10/11 |

## 📞 Support

**Auteur** : GOUNOU N'GOBI Chabi Zimé
**Email** : zimkada@gmail.com
**Version** : 3.2 (en développement)

---

© 2024-2025 GOUNOU N'GOBI Chabi Zimé - Tous droits réservés
