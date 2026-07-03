# ExcelCompiler v3.2 — Guide d'utilisation

**Compilez plusieurs fichiers Excel/CSV en un seul tableau, automatiquement.**

ExcelCompiler regroupe des fichiers qui ont la **même structure de colonnes**
(par exemple un même formulaire rempli par plusieurs établissements) en un
unique fichier consolidé, en alignant correctement les colonnes même si elles
ne sont pas exactement identiques d'un fichier à l'autre.

---

## 1. Installation

1. Lancez `ExcelCompiler_Setup.exe` et suivez l'assistant.
2. L'application nécessite **Windows 10 ou 11** (64 bits).
3. Au premier lancement, le composant Microsoft Visual C++ est installé si besoin.

---

## 2. Démarrage rapide

1. **Sélectionnez vos fichiers** — bouton « 📂 Choisir un dossier » (charge tous
   les fichiers Excel/CSV du dossier) ou **glissez-déposez** vos fichiers
   directement dans la zone prévue.
2. **Choisissez le mode de détection** des en-têtes (voir §3).
3. (Optionnel) Cliquez sur **« 👁 Aperçu »** pour vérifier ce qui sera détecté,
   et corriger un en-tête fichier par fichier si besoin.
4. Cliquez sur **« ▶ COMPILER »**.
5. Le fichier consolidé est créé **dans le dossier des fichiers sources**, et
   les statistiques s'affichent dans l'onglet **Résultats**.

---

## 3. Modes de détection des en-têtes

Le tableau de chaque fichier ne commence pas toujours à la première ligne (titre,
logo, lignes vides…). ExcelCompiler doit savoir **où est la ligne d'en-tête**.

- **Par fichier de référence (recommandé)** — vous indiquez le numéro de ligne
  de l'en-tête du **premier** fichier ; les autres fichiers sont alignés
  automatiquement en cherchant la ligne la plus ressemblante.
- **Manuel** — vous fixez vous-même la ligne d'en-tête, appliquée à tous les
  fichiers.

Dans l'**aperçu**, un fichier mal détecté peut être corrigé individuellement
(ligne d'en-tête forcée), et une colonne au libellé inhabituel peut être
rattachée manuellement à la bonne colonne du tableau final.

---

## 4. Formats pris en charge

| Entrée | Sortie |
|--------|--------|
| `.xlsx`, `.xlsm` (Excel moderne) | `.xlsx` (mis en forme) |
| `.csv`, `.tsv`, `.txt` (texte délimité) | `.csv`, `.tsv` |

> Les anciens fichiers `.xls` (Excel 97-2003) ne sont pas pris en charge :
> ouvrez-les dans Excel et enregistrez-les en `.xlsx` au préalable.

Pour les fichiers texte, le **séparateur est détecté automatiquement**
(virgule, point-virgule — courant en français —, ou tabulation), ainsi que
l'encodage (UTF-8, Latin-1, Windows-1252).

---

## 5. Options utiles

- **Ajouter le nom du fichier source** — ajoute une colonne indiquant de quel
  fichier provient chaque ligne (avec ou sans extension).
- **En-têtes multi-lignes** — fusionnés proprement en un libellé par colonne
  (ex. « NOMBRE DE CAS - 1er cycle »).
- **Cellules fusionnées** — la valeur d'une cellule fusionnée (ex. un nom de
  département) est recopiée sur toutes les lignes concernées.
- **Lignes de total / sous-total** — détectées automatiquement ; **exclues par
  défaut** (pour éviter le double comptage) ou conservées et marquées, au choix.
- **Tri** et **suppression des doublons** sur le tableau final.

Vos réglages (mode, options, format, nom de sortie) sont **mémorisés** et
restaurés au prochain lancement. Après une compilation réussie, un bouton
permet d'**ouvrir directement** le fichier produit ou son dossier.

---

## 6. Limites et bonnes pratiques

- Taille maximale par fichier : **100 Mo** (un fichier plus gros est ignoré, les
  autres sont compilés normalement).
- Pour de très gros fichiers, le nombre de lignes lues est borné par sécurité
  (un avertissement le signale).
- Les fichiers doivent partager une **structure de colonnes proche** ; des
  colonnes manquantes deviennent des cellules vides (jamais inventées).
- Dans un classeur à **plusieurs feuilles**, seule la première est compilée —
  l'application le signale ; placez les données à compiler sur la 1ʳᵉ feuille.
- Le fichier de sortie est automatiquement **exclu des sources** s'il se trouve
  déjà dans le dossier (pas de double comptage), et son **écrasement est
  confirmé** avant remplacement.

L'application **vérifie au démarrage** si une version plus récente est
disponible (via GitHub) et l'annonce par une bannière discrète — sans jamais
gêner votre travail, ni rien envoyer d'autre qu'une simple requête de version.

---

## 7. Où sont les fichiers ?

- **Sortie** : dans le dossier des fichiers sources.
- **Journaux** (en cas de problème) : `%LOCALAPPDATA%\ExcelCompiler\logs\`.

---

## Auteur & licence

© 2026 GOUNOU N'GOBI Chabi Zimé — Data Manager & Data Analyst.
Distribué sous licence BSD 3-Clause (voir `LICENSE`).
Voir aussi `privacy_policy.html` et `terms_of_service.html`.
