# 🚀 ExcelCompiler v3.2 — Notes de version

## ✨ Nouveautés de la v3.2

### 🔧 Compilation
- Compilation de fichiers **Excel (.xlsx, .xlsm), CSV et TSV** en un seul tableau.
- **Détection par fichier de référence** : indiquez l'en-tête du premier fichier,
  les autres sont alignés automatiquement (similarité de libellés).
- **Alignement des colonnes par libellé** : les fichiers aux colonnes dans un
  ordre différent (ou avec une colonne en plus/en moins) sont empilés
  correctement, sans corruption silencieuse.
- **En-têtes multi-lignes** aplatis proprement en un libellé par colonne.
- **Dé-fusion des cellules fusionnées** : une valeur fusionnée (ex. un
  département) est propagée sur toute sa plage.
- **Lignes de sous-total / total** détectées : exclues par défaut (évite le
  double comptage) ou conservées et marquées, au choix.
- **Aperçu de détection** avant compilation + correction manuelle par fichier.
- Tri par colonne et suppression des doublons.

### ⚡ Performance
- **Écriture Excel accélérée** (environ ×2 sur les gros volumes) : écriture par
  lots, bordures adaptatives au-delà d'un seuil, largeurs de colonnes estimées
  sur échantillon — l'intégrité des données est préservée.

### 🛡️ Sécurité et robustesse
- **Garde-fou de taille de fichier** réellement appliqué : un fichier trop
  volumineux est rejeté proprement sans interrompre les autres.
- **Borne anti-explosion mémoire** sur le nombre de lignes lues (avec
  avertissement transparent).
- **Annulation réactive** de la compilation, y compris pendant le traitement
  d'un gros fichier.
- **Séparateur CSV détecté** automatiquement (virgule, point-virgule,
  tabulation) — fini les CSV français lus en une seule colonne.
- Lecture tolérante aux encodages (UTF-8, Latin-1, CP1252) pour CSV/TSV.

### 🔒 Protection des données à l'export
- L'**extension du fichier de sortie** est synchronisée avec le format choisi.
- **Confirmation** avant d'écraser un fichier de sortie existant.
- Le fichier de sortie est **exclu des sources** s'il se trouve dans le dossier
  (évite de recompiler la sortie précédente).
- Les **classeurs multi-feuilles** sont signalés (seule la 1ʳᵉ est compilée).

### 💡 Confort
- **Options mémorisées** d'une session à l'autre (mode, cases, format, nom de
  sortie, colonne de tri).
- Après compilation, boutons **« Ouvrir le fichier / le dossier »**.
- Colonne de tri **validée** (une saisie invalide est refusée, pas de tri
  silencieux sur la mauvaise colonne).
- **Vérification de mise à jour** au démarrage (bannière discrète, silencieuse
  si hors-ligne).

### 🎯 Prêt à l'emploi
- Détection par référence activée par défaut.
- Format de date français par défaut.
- Sortie Excel mise en forme (en-têtes stylés, volets figés, largeurs ajustées).

### 🖥️ Compatibilité
- Windows 10 / 11 (64 bits), interface PyQt6.

---
© 2026 GOUNOU N'GOBI Chabi Zimé — Tous droits réservés
