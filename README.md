# GPXBuilder — version Streamlit

Version web de GPXBuilder, compatible avec [streamlit.io](https://streamlit.io)
(Streamlit Community Cloud) ou tout autre hébergeur Streamlit.

## Contenu

- `app.py` — l'interface web (upload de fichiers, traitement, téléchargement du GPX)
- `gpxbuilder_core.py` — toute la logique métier du script original, **inchangée**
  (parsing XLSX/CSV/TXT, détection auto des colonnes, fusion GPX, garde-fous, etc.)
- `requirements.txt` — dépendances (`streamlit`, `openpyxl`)

## Lancer en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Déployer sur streamlit.io (Community Cloud)

1. Créez un dépôt GitHub contenant ces 3 fichiers (à la racine, ou dans un
   sous-dossier — indiquez alors le chemin de `app.py` lors du déploiement).
2. Allez sur https://share.streamlit.io, connectez votre compte GitHub.
3. « New app » → sélectionnez le dépôt, la branche, et `app.py` comme fichier
   principal.
4. Déployez : Streamlit installera automatiquement `requirements.txt`.

## Ce qui change par rapport à la version bureau (Tkinter)

- **Sélection de fichiers** : upload via glisser-déposer ou sélecteur (au lieu
  de « Ajouter des fichiers... » / « Ajouter un dossier... »). L'ajout d'un
  dossier entier n'a pas de sens côté web ; sélectionnez plusieurs fichiers à
  la fois dans le sélecteur du navigateur.
- **Dossier de destination** : n'existe plus (une app web n'a pas accès au
  disque de l'utilisateur). Le fichier GPX généré est proposé directement en
  téléchargement, et vous pouvez modifier son nom avant de lancer le
  traitement.
- **Journal** : affiché après traitement (pas en direct ligne par ligne,
  Streamlit ré-exécute le script à chaque interaction), avec les mêmes
  boutons « copier » (zone de texte sélectionnable) et « enregistrer le
  journal ».
- Toute la logique de conversion (détection des colonnes, formats acceptés,
  garde-fous, placement des notes dans le champ Note de GSAK, etc.) reste
  **strictement identique** à celle décrite dans le guide d'utilisation
  fourni : elle a été reprise telle quelle dans `gpxbuilder_core.py`, sans
  aucune modification du code de traitement.

## Limitation notable

`cacheconverter_config.json` (mémorisation du dernier dossier de destination)
n'est plus utilisé, puisqu'il n'y a plus de dossier de destination local à
mémoriser.
