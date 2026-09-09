# GPXBuilder

Convertisseur de fichiers **Excel (.xlsx/.xlsm) / CSV / TXT / GPX** vers un
fichier **GPX compatible GSAK** (Geocaching Swiss Army Knife).

🌐 **Application en ligne : [gpxbuilder.streamlit.app](https://gpxbuilder.streamlit.app/)**

Aucune installation n'est nécessaire : ouvrez le lien ci-dessus dans votre
navigateur.

---

## Utilisation

1. **Uploader des fichiers**
   Sur la page, cliquez dans la zone d'upload (ou glissez-déposez) pour
   ajouter un ou plusieurs fichiers `.xlsx`, `.xlsm`, `.csv`, `.txt` ou
   `.gpx`. Le mode de traitement est détecté automatiquement selon le
   contenu de la liste :

   | Fichiers ajoutés | Mode |
   |---|---|
   | Un seul `.xlsx` / `.csv` / `.txt` | Conversion simple |
   | Plusieurs `.xlsx` / `.csv` / `.txt` | Conversion de chaque fichier puis fusion en un seul GPX |
   | Plusieurs `.gpx` | Fusion telle quelle (aucun reformatage) |
   | Mélange `.gpx` + autres | Mode mixte : les `.gpx` sont fusionnés tels quels, les autres convertis — le tout combiné dans un seul GPX |
   | Un seul `.gpx`, seul dans la liste | Rien à faire (ce fichier existe déjà tel quel) |

2. **Démarrer**
   Cliquez sur le bouton **🚀 Démarrer**. Le journal de traitement s'affiche
   (fichier par fichier, avec les succès, lignes ignorées et erreurs), puis
   un bouton **⬇️ Télécharger le fichier GPX** apparaît avec le résultat,
   nommé automatiquement `geocaches_JJMMAA_HHMMSS.gpx`.

   Le journal peut aussi être enregistré via **📋 Enregistrer le journal
   (.txt)**.

3. **Clear**
   Le bouton **🧹 Clear** réinitialise entièrement l'écran (fichiers
   uploadés, journal, résultat) pour repartir sur un nouveau traitement.

### Formats reconnus en entrée

- **XLSX / CSV avec en-tête** : les colonnes sont reconnues automatiquement
  par leur nom (français ou anglais, avec ou sans accents, insensible à la
  casse) — `Code`, `Nom`, `Latitude`, `Longitude`, `Type`, `Taille`,
  `Difficulté`/`Terrain`, `Propriétaire`, `Description`, `Indice`, `Note`,
  etc. Seules les colonnes Code/Nom/Latitude/Longitude sont utiles, le reste
  est optionnel.
- **XLSX / CSV sans en-tête** : si aucune colonne n'est reconnue par son nom,
  le programme devine le rôle de chaque colonne par son contenu (code GC,
  coordonnées...).
- **TXT (texte libre)** : une cache par ligne, au format
  `GCxxxxx Nom N50°.. E010°.. texte libre` ou
  `Nom (GCxxxxx) N50°.. E010°.. texte libre`. Le texte libre après les
  coordonnées est repris dans le champ Note.
- **GPX** : fusionné tel quel, sans reformatage.

### Garde-fous automatiques

- Coordonnées invalides (hors plage ou `(0, 0)`) : ligne ignorée, détail
  dans le journal.
- Codes GC en double (même fichier ou entre plusieurs fichiers) : les deux
  caches sont tout de même incluses, avec un avertissement dans le journal.
- Fichier illisible ou mal formé dans un lot : sauté, le traitement continue
  avec les fichiers suivants.
- Encodage CSV : repli automatique UTF-8 → Windows-1252 → Latin-1.

### Où sont placées les notes personnelles (GSAK)

Le contenu d'une colonne `UserNote`/`Note` (ou le texte libre après les
coordonnées dans un fichier TXT) est placé dans le **vrai champ Note de
GSAK** (celui avec l'icône), et non dans un champ personnalisé — exactement
le format utilisé par GSAK lui-même lors de ses propres exports.

---

## Développement / auto-hébergement

Ce dépôt contient tout le nécessaire pour relancer l'application ailleurs
que sur `gpxbuilder.streamlit.app` :

- `app.py` — l'interface web (upload de fichiers, traitement, téléchargement
  du GPX)
- `gpxbuilder_core.py` — toute la logique métier (parsing XLSX/CSV/TXT,
  détection auto des colonnes, fusion GPX, garde-fous, génération du GPX
  final compatible GSAK)
- `requirements.txt` — dépendances (`streamlit`, `openpyxl`)

### Lancer en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

### Déployer sur Streamlit Community Cloud

1. Poussez ces 3 fichiers sur un dépôt GitHub (à la racine, ou dans un
   sous-dossier — indiquez alors le chemin de `app.py` lors du déploiement).
2. Allez sur [share.streamlit.io](https://share.streamlit.io), connectez
   votre compte GitHub.
3. « New app » → sélectionnez le dépôt, la branche, et `app.py` comme
   fichier principal.
4. Déployez : Streamlit installera automatiquement `requirements.txt`.

---

## Différences avec la version bureau (Tkinter)

- **Sélection de fichiers** : upload via le navigateur au lieu de
  « Ajouter des fichiers... » / « Ajouter un dossier... ». L'ajout d'un
  dossier entier n'a pas de sens côté web ; sélectionnez plusieurs fichiers
  à la fois dans le sélecteur.
- **Dossier de destination** : n'existe plus (une app web n'a pas accès au
  disque de l'utilisateur). Le fichier GPX généré est proposé directement
  en téléchargement.
- **Journal** : affiché après traitement plutôt qu'en direct ligne par
  ligne.
- Toute la logique de conversion (détection des colonnes, formats acceptés,
  garde-fous, placement des notes dans le champ Note de GSAK, etc.) reste
  **strictement identique** à la version bureau : elle est reprise telle
  quelle dans `gpxbuilder_core.py`.
mémoriser.
