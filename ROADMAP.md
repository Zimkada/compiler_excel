# Roadmap — ExcelCompiler

Ce document est la **source de vérité** des fonctionnalités prévues mais non
encore implémentées. Toute fonctionnalité décrite ici n'existe **pas** dans la
version courante : ne pas la présenter comme disponible dans le marketing,
le README ou l'interface tant qu'elle n'est pas livrée et testée.

> Convention : une fonctionnalité ne quitte cette roadmap pour le README /
> l'interface qu'une fois **implémentée ET couverte par des tests**.

---

## État actuel (livré et testé)

- Compilation multi-fichiers Excel (.xlsx/.xls/.xlsm), CSV, TSV.
- Détection de structure par **fichier de référence** (mode par défaut de
  l'interface) : l'utilisateur indique l'en-tête du 1er fichier, les autres
  sont alignés par similarité (Jaccard).
- Détection de fin de tableau robuste (coupe sur bloc de N lignes vides
  consécutives, pas sur un trou isolé).
- Options : tri (type-safe), suppression doublons / lignes vides, répétition
  d'en-têtes, colonne « fichier source », formats de date (dont date+heure).
- Export XLSX (formaté), CSV, TSV.
- Progression réelle et annulation fonctionnelle.

---

## À venir

### 1. Mode de détection automatique hybride (reconnexion)

**Statut : détection réparée, pas encore exposée dans l'UI.**

Le `HybridDetector` (BorderDetector + DensityDetector + PatternDetector + vote
pondéré + validation croisée) est implémenté mais reste **inaccessible depuis
l'interface** (seuls les modes « référence » et « manuel » sont proposés).

Avancement :
- ✅ Bug critique corrigé : le BorderDetector gonflait `header_rows` (5/13/29
  lignes) dans les tableaux quadrillés → 0 donnée extraite. Désormais 6/6
  fichiers réels correctement détectés, données complètes. Couvert par tests.
- ✅ Validation croisée assainie : elle ne dégrade plus la confiance et ne pose
  plus d'avertissements alarmants ; elle ne fait que renforcer la confiance
  quand les fichiers se ressemblent fortement. Couvert par tests.

Reste à faire avant exposition :
- ⏳ **Exposer** une case « Détection automatique » dans `OptionsWidget`, idéalement
  avec un aperçu de la détection (ligne d'en-tête trouvée) avant compilation.
- ⏳ Harmoniser la convention `data_end_row` entre détecteurs (le ReferenceDetector
  renvoie base 0 exclusive, le BorderDetector base 1) — sans impact aujourd'hui
  (l'usage est borné à `len(df)`) mais fragile.
- Fichiers concernés : `ui/widgets/options_widget.py`,
  `core/compilation/excel_compiler.py`.

### 2. ML local optionnel (apprentissage personnalisé)

**Statut : non implémenté** (`core/detection/ml_local.py` n'existe pas).

- Objectif : apprendre des corrections de l'utilisateur pour améliorer la
  détection sur ses fichiers récurrents.
- Paramètres prévus (à recâbler depuis `config/constants.py` quand implémenté) :
  - taille minimale d'entraînement : 10 exemples,
  - désactivé par défaut.
- Dépendance : scikit-learn (optionnelle, ne pas l'imposer au build de base).
- Intégration prévue : comme détecteur supplémentaire dans le HybridDetector.

### 3. Détection d'anomalies

**Statut : non implémenté** (`core/anomalies/` est vide).

- Objectif : signaler doublons, valeurs manquantes, valeurs aberrantes,
  incohérences de format, avant ou après compilation.
- Seuils prévus (à recâbler depuis `config/constants.py` quand implémenté) :
  - multiplicateur IQR pour les outliers : 3.0,
  - seuil de valeurs manquantes : 20 %.
- UI prévue : bouton « Vérifier les anomalies » + rapport.

### 4. Traitement par chunks pour gros fichiers

**Statut : options présentes mais non câblées.** `CompilationOptions` expose
`enable_chunked_processing`, `chunk_size`, `max_memory_percent`, mais le moteur
charge tout en mémoire (`pd.read_excel(...).tolist()`).

- Objectif : compiler de gros volumes sans saturer la mémoire.
- Piste : rester en DataFrame de bout en bout (`pd.concat`) plutôt que de
  manipuler des `List[List]` Python, et traiter par lots.
- Prévoir un vrai plafond mémoire effectif (aujourd'hui `max_memory_percent`
  n'a aucun effet).

### 5. Formats d'export supplémentaires (Parquet, JSON)

**Statut : annoncés dans les constantes, non implémentés.** Seuls XLSX/CSV/TSV
existent (`OutputFormat`).

- Ajouter les variantes à `OutputFormat` et au point d'écriture
  (`_write_output_file`) + à l'`OptionsWidget`.

---

## Hors périmètre fonctionnel (chantier commercialisation)

À traiter séparément de la roadmap fonctionnelle :

- Renommage du produit (« Excel » est une marque Microsoft — risque juridique).
- Système de licence / période d'essai.
- Signature de code de l'exécutable (sinon SmartScreen « éditeur inconnu »).
- Mises à jour automatiques (version.json distant + notification).
- Télémétrie d'erreurs pour le support.
