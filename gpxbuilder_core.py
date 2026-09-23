#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GPXBuilder (anciennement CacheConverter / XLSX_CSV_TXT_to_GPX)
==================================================================================
Convertisseur XLSX / CSV / TXT / GPX -> GPX compatible GSAK (format Groundspeak cache 1/0/1)

Interface graphique (Tkinter) permettant de :
  - choisir un fichier d'entrée (.xlsx, .csv ou .txt)
  - choisir le dossier de destination (le fichier GPX est nommé automatiquement)
  - lancer la conversion
  - suivre en direct un journal (log) indiquant le nombre de caches ajoutées
    et si tout s'est bien déroulé (erreurs / lignes ignorées)

Dépendances : openpyxl (uniquement nécessaire pour lire les fichiers .xlsx)
    pip install openpyxl

-----------------------------------------------------------------------------
COLONNES ATTENDUES DANS LE FICHIER D'ENTREE
-----------------------------------------------------------------------------
Le script reconnaît automatiquement plusieurs variantes de noms de colonnes
(en français ou en anglais, avec ou sans accents, insensible à la casse).
Seules les colonnes "code", "nom", "latitude" et "longitude" sont obligatoires,
toutes les autres sont optionnelles.

Champ           | Exemples d'en-têtes acceptés
----------------|--------------------------------------------------------------
code            | Code, GC Code, Waypoint, GCCode, ID
nom             | Nom, Name, Cache Name, Titre
latitude        | Latitude, Lat, Y
longitude       | Longitude, Lon, Lng, Long, X
coordonnees     | Coordonnees, Coordinates, Coord, Coords, Position, GPS
                   (utile si latitude/longitude ne sont pas séparées, ex :
                    "N 48° 51.402 E 002° 17.567")
type            | Type, Cache Type, Type de cache
taille          | Taille, Container, Size
difficulte      | Difficulte, Difficulty, D
terrain         | Terrain, T
proprietaire    | Proprietaire, Owner, Placed By, Placé par
pays            | Pays, Country
region          | Region, State, Département
description     | Description, Description courte, Short Description
description_longue | Description longue, Long Description, Details
indice          | Indice, Hint, Hints
date            | Date, Date placee, Placed Date, Hidden
url             | URL, Lien, Link
disponible      | Disponible, Available, Active
archivee        | Archivee, Archived
-----------------------------------------------------------------------------
Le script détecte aussi automatiquement les fichiers XLSX/CSV SANS ligne d'en-tête :
si aucune colonne "code"/"latitude"/"longitude" n'est reconnue par nom, il analyse
le contenu des colonnes pour deviner leur rôle (code GC, coordonnées, etc.).

Fichiers .TXT (texte libre, une cache par ligne, format variable) :
Chaque ligne peut contenir un ou plusieurs codes GC, sous deux formes :
  - "GCxxxxx Nom de la cache N50°.. E010°.. texte_libre"      (code en tête)
  - "Nom de la cache (GCxxxxx) N50°.. E010°.. texte_libre"    (code entre parenthèses)
Le texte libre après les coordonnées (indice de résolution, mot de passe, etc.)
est repris tel quel dans le champ Note. Toute portion de ligne sans coordonnées
valides est ignorée (avec message dans le journal), ce qui gère aussi le cas de
plusieurs caches collées sur une même ligne sans séparateur.
-----------------------------------------------------------------------------
"""

import io
import os
import re
import csv
import sys
import json
import queue
import threading
import subprocess
import unicodedata
import xml.etree.ElementTree as ET
import xml.dom.minidom as minidom
from datetime import datetime, timezone

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
except ImportError:
    tk = None  # Permet d'importer ce module (ex: tests) sans Tkinter installé

APP_NAME = "GPXBuilder"
APP_VERSION = "1.0.0"

HELP_TEXT = f"""{APP_NAME} v{APP_VERSION} — Aide rapide

FORMATS ACCEPTÉS
  .xlsx / .csv : colonnes reconnues par leur nom (Code, Nom, Latitude,
  Longitude, Type, Difficulté, Terrain, Note...), en français ou anglais.
  Si aucune colonne n'est reconnue, le programme suppose qu'il n'y a pas
  d'en-tête et devine les colonnes par leur contenu.
  .txt : texte libre, une cache par ligne, au format
  "GCxxxxx Nom N50°.. E010°.." ou "Nom (GCxxxxx) N50°.. E010°..".
  .gpx : fusionné tel quel, sans reformatage.

MODE DE TRAITEMENT (automatique selon la liste de fichiers)
  - Un seul fichier .xlsx/.csv/.txt  -> conversion simple
  - Plusieurs fichiers convertibles  -> conversion + fusion en un seul GPX
  - Plusieurs fichiers .gpx          -> fusion telle quelle
  - Mélange des deux                 -> mode mixte (chacun selon son type)
  - Un seul fichier .gpx seul        -> rien à faire (il existe déjà)

UTILISATION
  1. Ajoutez un ou plusieurs fichiers (bouton, dossier, ou glisser dans
     la liste n'est pas supporté : utilisez "Ajouter des fichiers...").
  2. Choisissez le dossier de destination (mémorisé pour la prochaine fois).
  3. Cliquez sur "Démarrer". Le nom du GPX est généré automatiquement
     (geocaches_JJMMAA_HHMMSS.gpx).

NOTES PERSONNELLES ET GSAK
  Une colonne "UserNote"/"Note" (ou le texte libre après les coordonnées
  dans un .txt) est placée dans le vrai champ Note de GSAK (celui avec
  l'icône), pas dans un champ personnalisé.

GARDE-FOUS AUTOMATIQUES
  - Coordonnées hors plage ou (0,0) : ligne ignorée (détail dans le journal)
  - Codes GC en double (même fichier ou entre plusieurs fichiers) :
    les deux caches sont incluses, avec un avertissement dans le journal
  - Fichier illisible dans un lot : sauté, les autres continuent
  - Encodage CSV : repli automatique UTF-8 / Windows-1252 / Latin-1

EN CAS DE PROBLÈME
  "Module openpyxl introuvable" : réinstallez-le précisément pour
  l'interpréteur Python indiqué dans le journal (voir le guide complet).
  "Colonnes de coordonnées introuvables" : le journal détaille les
  colonnes détectées et reconnues pour vous aider à corriger le fichier.

Pour le guide complet et détaillé, voir le fichier PDF fourni avec ce
programme (bouton ci-dessous s'il est présent dans le même dossier).
"""


# =============================================================================
# 1. NORMALISATION DES EN-TETES DE COLONNES
# =============================================================================

COLUMN_ALIASES = {
    "code": ["code", "gc", "gc code", "gccode", "waypoint", "id"],
    "nom": ["nom", "name", "cache name", "titre"],
    "latitude": ["latitude", "lat", "y"],
    "longitude": ["longitude", "lon", "lng", "long", "x"],
    "coordonnees": ["coordonnees", "coordinates", "coord", "coords", "position", "gps"],
    "type": ["type", "cache type", "type de cache"],
    "taille": ["taille", "container", "size"],
    "difficulte": ["difficulte", "difficulty", "d"],
    "terrain": ["terrain", "t"],
    "proprietaire": ["proprietaire", "owner", "placed by", "place par"],
    "pays": ["pays", "country"],
    "region": ["region", "state", "departement"],
    "description": ["description", "description courte", "short description"],
    "description_longue": ["description longue", "long description", "details"],
    "indice": ["indice", "hint", "hints"],
    "notes": ["usernote", "user note", "note", "notes", "personal note", "personalnote", "note personnelle"],
    "date": ["date", "date placee", "placed date", "hidden"],
    "url": ["url", "lien", "link"],
    "disponible": ["disponible", "available", "active"],
    "archivee": ["archivee", "archived"],
}

# Table de correspondance types abrégés -> types officiels Groundspeak
TYPE_ALIASES = {
    "tradi": "Traditional Cache",
    "traditional": "Traditional Cache",
    "multi": "Multi-cache",
    "multi-cache": "Multi-cache",
    "mystery": "Unknown Cache",
    "unknown": "Unknown Cache",
    "enigme": "Unknown Cache",
    "letterbox": "Letterbox Hybrid",
    "wherigo": "Wherigo Cache",
    "earthcache": "Earthcache",
    "virtual": "Virtual Cache",
    "virtuelle": "Virtual Cache",
    "event": "Event Cache",
    "evenement": "Event Cache",
    "webcam": "Webcam Cache",
    "night": "Night Cache",
    "cito": "Cache In Trash Out Event",
}

CONTAINER_ALIASES = {
    "micro": "Micro",
    "nano": "Micro",
    "small": "Small",
    "petite": "Small",
    "regular": "Regular",
    "normale": "Regular",
    "large": "Large",
    "grande": "Large",
    "very large": "Very large",
    "not chosen": "Not chosen",
    "other": "Other",
    "autre": "Other",
}


def _strip_accents(text):
    """Normalise un texte pour faciliter la comparaison : retire les accents,
    la ponctuation courante (. _ - :), les espaces superflus et le BOM éventuel,
    puis met en minuscules."""
    text = str(text).replace("\ufeff", "").strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = text.lower()
    for ch in (".", "_", "-", ":", "'"):
        text = text.replace(ch, " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def build_header_map(fieldnames):
    """
    Associe à chaque champ canonique (ex: 'latitude') le nom de colonne réel
    trouvé dans le fichier (ex: 'Lat'). Retourne un dict {champ_canonique: nom_colonne_reel}.
    """
    normalized = {_strip_accents(f): f for f in fieldnames if f is not None}
    header_map = {}
    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in normalized:
                header_map[canonical] = normalized[alias]
                break
    return header_map


def get_field(row, header_map, canonical, default=""):
    col = header_map.get(canonical)
    if col is None:
        return default
    value = row.get(col, default)
    if value is None:
        return default
    return str(value).strip()


# =============================================================================
# 2. LECTURE DES FICHIERS D'ENTREE (XLSX / CSV)
# =============================================================================

CSV_ENCODINGS_TO_TRY = ("utf-8-sig", "cp1252", "latin-1")
CSV_DELIMITER_CANDIDATES = [",", ";", "\t", "|", ":"]


def _decode_file_with_fallback(path, encodings=CSV_ENCODINGS_TO_TRY):
    """Essaie plusieurs encodages courants (utf-8, puis Windows/Latin) et retourne
    (contenu_texte, encodage_utilisé). 'latin-1' ne peut jamais échouer (il mappe
    tous les octets), il sert donc de dernier repli garanti."""
    last_error = None
    for i, enc in enumerate(encodings):
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                return f.read(), enc
        except (UnicodeDecodeError, LookupError) as exc:
            last_error = exc
            if i == len(encodings) - 1:
                raise
            continue
    raise last_error  # pragma: no cover


def _detect_csv_delimiter(sample):
    """Détecte le séparateur le plus probable dans un échantillon de texte CSV."""
    delimiter = None
    try:
        dialect = csv.Sniffer().sniff(sample[:8192], delimiters="".join(CSV_DELIMITER_CANDIDATES))
        delimiter = dialect.delimiter
    except csv.Error:
        pass

    if delimiter is None:
        # Repli fiable : on compte l'occurrence de chaque séparateur candidat
        # sur les premières lignes, et on choisit celui qui donne un nombre
        # de colonnes cohérent (>1) et constant.
        lines = [l for l in sample.splitlines() if l.strip()][:5]
        best_delim, best_score = ",", -1
        for d in CSV_DELIMITER_CANDIDATES:
            counts = [line.count(d) for line in lines] or [0]
            if counts[0] > 0 and len(set(counts)) == 1:
                score = counts[0]
                if score > best_score:
                    best_delim, best_score = d, score
        if best_score <= 0:
            overall_counts = {d: sample.count(d) for d in CSV_DELIMITER_CANDIDATES}
            best_delim = max(overall_counts, key=overall_counts.get)
            if overall_counts[best_delim] == 0:
                best_delim = ","
        delimiter = best_delim

    return delimiter


def read_csv_rows(path):
    """Lit un fichier CSV/TSV et retourne (fieldnames, liste_de_dicts).
    Détecte automatiquement le séparateur (virgule, point-virgule, tabulation,
    pipe, deux-points) ainsi que l'encodage (UTF-8, puis Windows-1252/Latin-1
    en repli pour les vieux exports Excel)."""
    content, _encoding = _decode_file_with_fallback(path)
    delimiter = _detect_csv_delimiter(content)
    reader = csv.DictReader(io.StringIO(content), delimiter=delimiter)
    rows = [row for row in reader]
    fieldnames = reader.fieldnames or []
    return fieldnames, rows


def read_csv_raw(path):
    """Lit un fichier CSV/TSV SANS traiter la première ligne comme un en-tête.
    Retourne une liste de listes de chaînes (une par ligne non vide)."""
    content, _encoding = _decode_file_with_fallback(path)
    delimiter = _detect_csv_delimiter(content)
    reader = csv.reader(io.StringIO(content), delimiter=delimiter)
    return [row for row in reader if any(cell.strip() for cell in row)]


def read_xlsx_rows(path):
    """Lit un fichier XLSX et retourne (fieldnames, liste_de_dicts)."""
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    try:
        header_row = next(rows_iter)
    except StopIteration:
        return [], []

    fieldnames = [str(h) if h is not None else f"col_{i}" for i, h in enumerate(header_row)]
    rows = []
    for raw_row in rows_iter:
        if raw_row is None or all(v is None for v in raw_row):
            continue
        row = {}
        for i, value in enumerate(raw_row):
            key = fieldnames[i] if i < len(fieldnames) else f"col_{i}"
            row[key] = "" if value is None else value
        rows.append(row)
    return fieldnames, rows


def read_xlsx_raw(path):
    """Lit un fichier XLSX SANS traiter la première ligne comme un en-tête.
    Retourne une liste de listes de chaînes (une par ligne non vide)."""
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = []
    for raw_row in ws.iter_rows(values_only=True):
        if raw_row is None or all(v is None for v in raw_row):
            continue
        rows.append(["" if v is None else str(v) for v in raw_row])
    return rows


def read_input_file(path):
    """Détecte l'extension et lit le fichier en conséquence (mode normal, avec en-tête)."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return read_csv_rows(path)
    elif ext in (".xlsx", ".xlsm"):
        try:
            return read_xlsx_rows(path)
        except ImportError as exc:
            raise ImportError(
                f"Le module 'openpyxl' n'a pas pu être importé ({exc}).\n\n"
                "Cela arrive généralement quand 'pip install openpyxl' installe le module "
                "pour une autre version/installation de Python que celle utilisée pour "
                "lancer ce programme (cas fréquent avec plusieurs Python installés sous Windows).\n\n"
                f"Interpréteur Python actuellement utilisé par ce programme :\n    {sys.executable}\n\n"
                "Essayez d'installer openpyxl précisément pour CET interpréteur, en tapant "
                "dans une invite de commande (cmd) :\n\n"
                f'    "{sys.executable}" -m pip install openpyxl\n\n'
                "Puis relancez ce programme."
            ) from exc
    else:
        raise ValueError(f"Format de fichier non supporté : {ext} (attendu .xlsx ou .csv)")


def read_input_file_headerless(path):
    """Détecte l'extension et lit le fichier SANS traiter la première ligne comme
    un en-tête (utilisé en repli quand aucune colonne n'est reconnue par son nom)."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return read_csv_raw(path)
    elif ext in (".xlsx", ".xlsm"):
        return read_xlsx_raw(path)
    else:
        raise ValueError(f"Format de fichier non supporté : {ext} (attendu .xlsx ou .csv)")


# =============================================================================
# 3. CONVERSION DES COORDONNEES
# =============================================================================

DMM_REGEX = re.compile(
    r"""([NnSs])\s*0*(\d{1,3})[°\s]+(\d{1,2}(?:[.,]\d+)?)['\u2019\u2032\s]*"""
    r""".*?"""
    r"""([EeWwOo])\s*0*(\d{1,3})[°\s]+(\d{1,2}(?:[.,]\d+)?)['\u2019\u2032\s]*""",
    re.VERBOSE | re.DOTALL,
)


def _dmm_to_decimal(hemisphere, degrees, minutes):
    value = float(degrees) + float(minutes.replace(",", ".")) / 60.0
    if hemisphere.upper() in ("S", "W", "O"):
        value = -value
    return value


def parse_coordinates(row, header_map):
    """
    Retourne (lat, lon) en degrés décimaux (float), ou (None, None) si non trouvés.
    Supporte :
      - colonnes latitude/longitude en degrés décimaux (ex: 48.856614)
      - une colonne unique "coordonnees" au format geocaching (ex: N 48° 51.402 E 002° 17.567)
    """
    lat_raw = get_field(row, header_map, "latitude")
    lon_raw = get_field(row, header_map, "longitude")

    if lat_raw and lon_raw:
        try:
            return float(lat_raw.replace(",", ".")), float(lon_raw.replace(",", "."))
        except ValueError:
            pass  # peut-être au format DMM, on tente la regex ci-dessous

    combined = get_field(row, header_map, "coordonnees")
    candidates = [combined, f"{lat_raw} {lon_raw}".strip()]
    for candidate in candidates:
        if not candidate:
            continue
        match = DMM_REGEX.search(candidate)
        if match:
            ns, lat_deg, lat_min, ew, lon_deg, lon_min = match.groups()
            lat = _dmm_to_decimal(ns, lat_deg, lat_min)
            lon = _dmm_to_decimal(ew, lon_deg, lon_min)
            return lat, lon

    return None, None


# =============================================================================
# 3bis. DETECTION AUTOMATIQUE DES COLONNES (fichiers SANS ligne d'en-tête)
# =============================================================================
# Utilisée uniquement en repli, quand aucune colonne "code"/"latitude"/"longitude"
# n'est reconnue par son nom dans la ligne d'en-tête supposée. On analyse alors le
# CONTENU des colonnes pour deviner leur rôle, avec cette logique (validée avec
# l'utilisateur) :
#   - Code    : colonne où la majorité des valeurs ressemblent à un code GC
#   - Nom     : colonne immédiatement à DROITE du Code (peut être vide)
#   - Latitude / Longitude : détectées par leur CONTENU (DMM avec lettre
#     d'hémisphère, ou décimal avec gestion virgule/point, dans une plage valide)
#   - Note    : colonne immédiatement à DROITE de la Longitude (position seule,
#     le contenu peut être vide, textuel ou purement numérique)

GC_CODE_REGEX = re.compile(r"^GC[0-9A-Z]{2,8}$", re.IGNORECASE)

# Variante non ancrée (recherche n'importe où dans une ligne de texte libre),
# utilisée par le parseur de fichiers .txt, avec limites de mot pour éviter
# de matcher accidentellement une sous-chaîne d'un autre mot.
GC_CODE_SCAN_REGEX = re.compile(r"\bGC[0-9A-Z]{2,8}\b", re.IGNORECASE)

# Une seule coordonnée avec lettre d'hémisphère, ex: "N50°22.325" ou "E 004° 21.123"
SINGLE_DMM_REGEX = re.compile(
    r"([NnSsEeWwOo])\s*0*(\d{1,3})[°\s]+(\d{1,2}(?:[.,]\d+)?)"
)


def _looks_like_gc_code(value):
    return bool(GC_CODE_REGEX.match(value.strip()))


def _column_gc_score(values):
    non_empty = [v for v in values if v.strip()]
    if not non_empty:
        return 0.0
    matches = sum(1 for v in non_empty if _looks_like_gc_code(v))
    return matches / len(non_empty)


def _parse_single_dmm(value):
    """Retourne (hemisphere_majuscule, valeur_decimale) ou (None, None) si aucun
    motif DMM avec lettre d'hémisphère n'est trouvé dans la valeur."""
    match = SINGLE_DMM_REGEX.search(value)
    if not match:
        return None, None
    hemi, deg, minu = match.groups()
    return hemi.upper(), _dmm_to_decimal(hemi, deg, minu)


def classify_columns_by_content(raw_rows):
    """
    Analyse une liste de lignes brutes (liste de listes de chaînes, SANS en-tête,
    la première ligne étant déjà une ligne de données) et tente de deviner
    l'index de chaque colonne utile.

    Retourne un dict {"code": idx, "nom": idx_ou_None, "latitude": idx,
    "longitude": idx, "notes": idx_ou_None} en cas de succès, ou None si aucune
    colonne ne ressemble suffisamment à des codes GC (détection impossible).
    """
    if not raw_rows:
        return None

    num_cols = max(len(r) for r in raw_rows)
    columns = [[(row[i] if i < len(row) else "") for row in raw_rows] for i in range(num_cols)]

    # --- 1. Colonne Code : la colonne avec le meilleur taux de codes GC ---
    gc_scores = [_column_gc_score(col) for col in columns]
    code_idx = max(range(num_cols), key=lambda i: gc_scores[i])
    if gc_scores[code_idx] < 0.5:
        return None  # Aucune colonne ne ressemble à des codes GC : on ne devine rien

    # --- 2. Colonne Nom : juste à droite du Code ---
    name_idx = code_idx + 1 if code_idx + 1 < num_cols else None

    used = {code_idx, name_idx} if name_idx is not None else {code_idx}
    candidate_idxs = [i for i in range(num_cols) if i not in used]

    # --- 3. Latitude / Longitude : d'abord par motif DMM (lettre d'hémisphère) ---
    lat_idx, lon_idx = None, None
    for i in candidate_idxs:
        non_empty = [v for v in columns[i] if v.strip()][:20]  # échantillon
        if not non_empty:
            continue
        hemispheres = [h for h, _ in (_parse_single_dmm(v) for v in non_empty) if h]
        if not hemispheres:
            continue
        n_count = sum(1 for h in hemispheres if h in "NS")
        e_count = sum(1 for h in hemispheres if h in "EWO")
        if n_count >= e_count and n_count >= len(hemispheres) * 0.5 and lat_idx is None:
            lat_idx = i
        elif e_count > n_count and e_count >= len(hemispheres) * 0.5 and lon_idx is None:
            lon_idx = i

    # --- 4. Repli : colonnes décimales pures (gère virgule ou point décimal) ---
    if lat_idx is None or lon_idx is None:
        decimal_cols = []
        for i in candidate_idxs:
            if i in (lat_idx, lon_idx):
                continue
            non_empty = [v for v in columns[i] if v.strip()][:20]
            if not non_empty:
                continue
            values = []
            ok = True
            for v in non_empty:
                try:
                    values.append(float(v.strip().replace(",", ".")))
                except ValueError:
                    ok = False
                    break
            if ok and values:
                decimal_cols.append((i, values))

        # Une colonne dont une valeur dépasse ±90 ne peut être QUE une longitude
        if lon_idx is None:
            for i, values in decimal_cols:
                if any(abs(v) > 90 for v in values):
                    lon_idx = i
                    break

        remaining = [(i, v) for i, v in decimal_cols if i not in (lat_idx, lon_idx)]
        if lat_idx is None and remaining:
            lat_idx = remaining[0][0]
            remaining = remaining[1:]
        if lon_idx is None and remaining:
            lon_idx = remaining[0][0]

    if lat_idx is None or lon_idx is None:
        return None  # Impossible de localiser les coordonnées : on ne devine rien

    # --- 5. Colonne Note : juste à droite de la Longitude (position seule) ---
    notes_idx = None
    after_lon = lon_idx + 1
    if after_lon < num_cols and after_lon not in (code_idx, name_idx, lat_idx):
        notes_idx = after_lon

    return {
        "code": code_idx,
        "nom": name_idx,
        "latitude": lat_idx,
        "longitude": lon_idx,
        "notes": notes_idx,
    }


def build_headerless_rows_and_mapping(raw_rows):
    """
    À partir de lignes brutes (sans en-tête) et du résultat de
    classify_columns_by_content(), construit :
      - une liste de dicts {colonne_synthetique: valeur} exploitable par le
        reste du programme (comme s'il s'agissait de rows normales)
      - un header_map {champ_canonique: colonne_synthetique}
      - la liste des noms de colonnes synthétiques (pour les messages de log)
    Retourne (rows, header_map, synthetic_fieldnames) ou (None, None, None)
    si la détection automatique a échoué.
    """
    mapping_idx = classify_columns_by_content(raw_rows)
    if mapping_idx is None:
        return None, None, None

    num_cols = max(len(r) for r in raw_rows)
    col_keys = [f"Colonne{i + 1}" for i in range(num_cols)]

    rows = []
    for r in raw_rows:
        rows.append({col_keys[i]: (r[i] if i < len(r) else "") for i in range(num_cols)})

    header_map = {}
    for canonical, idx in mapping_idx.items():
        if idx is not None:
            header_map[canonical] = col_keys[idx]

    return rows, header_map, col_keys


def parse_txt_freeform(path, log):
    """
    Parse un fichier .txt en texte libre (une ou plusieurs caches par ligne,
    format variable, voir la documentation en tête de fichier) et retourne
    (rows, fieldnames, header_map) directement exploitables par generate_gpx().

    Algorithme (voir l'analyse détaillée convenue avec l'utilisateur) :
      - On repère TOUS les codes GC présents dans une ligne, peu importe leur
        position, et on découpe la ligne en autant de segments.
      - Pour chaque segment, si le code est précédé d'une parenthèse ouvrante
        ("(" juste avant, ex: "Nom (GCxxxxx)") -> le nom est le texte avant
        cette parenthèse. Sinon (code en tête) -> le nom est le texte après
        le code, avant les coordonnées.
      - Les coordonnées sont recherchées juste après le code (et son éventuelle
        parenthèse fermante), dans la portion de ligne qui précède le début du
        segment suivant. Si aucune coordonnée valide n'est trouvée, le segment
        est ignoré (couvre notamment le cas de deux caches collées sans
        séparateur, où la première n'a pas de coordonnées).
      - Tout le texte restant après les coordonnées (jusqu'à la fin du segment)
        est repris tel quel comme "note" (indice de résolution, mot de passe...).
    """
    content, _encoding = _decode_file_with_fallback(path)
    lines = content.splitlines()

    fieldnames = ["code", "nom", "latitude", "longitude", "notes"]
    header_map = {"code": "code", "nom": "nom", "latitude": "latitude", "longitude": "longitude", "notes": "notes"}
    rows = []

    for line_no, raw_line in enumerate(lines, start=1):
        line = raw_line.strip()
        if not line:
            continue

        matches = list(GC_CODE_SCAN_REGEX.finditer(line))
        if not matches:
            continue  # Ligne sans aucun code GC : pas une cache, ignorée silencieusement

        pos = 0
        for idx, m in enumerate(matches):
            code = m.group(0).upper()
            start, end = m.span()

            label = f"Ligne {line_no}" + (f", segment {idx + 1}/{len(matches)}" if len(matches) > 1 else "")

            # --- Détection du format : code entre parenthèses, ou code en tête ---
            before = line[pos:start]
            stripped_before = before.rstrip()
            is_format_paren = stripped_before.endswith("(")
            name_before = stripped_before[:-1].strip(" -") if is_format_paren else None

            # --- Position de départ pour chercher les coordonnées (après le code,
            #     et après la parenthèse fermante ")" le cas échéant) ---
            after_code = line[end:]
            search_offset = end
            if is_format_paren:
                lstripped = after_code.lstrip()
                if lstripped.startswith(")"):
                    consumed = len(after_code) - len(lstripped)
                    search_offset = end + consumed + 1

            window_end = matches[idx + 1].start() if idx + 1 < len(matches) else len(line)
            window = line[search_offset:window_end]

            coord_match = DMM_REGEX.search(window)
            if coord_match is None:
                log(f"  [IGNORÉE] {label} : code {code} trouvé mais coordonnées manquantes ou invalides.")
                # Ce segment ne "consomme" que son propre code (+ parenthèse
                # fermante éventuelle) : le reste du texte (ex: "Nom (" d'une
                # cache suivante au format parenthèses) reste disponible pour
                # le segment suivant.
                pos = search_offset
                continue

            # Coordonnées trouvées : ce segment consomme tout le texte jusqu'au
            # segment suivant (nom éventuel + coordonnées + note).
            pos = window_end

            ns, lat_deg, lat_min, ew, lon_deg, lon_min = coord_match.groups()
            lat = _dmm_to_decimal(ns, lat_deg, lat_min)
            lon = _dmm_to_decimal(ew, lon_deg, lon_min)

            if name_before is not None:
                name = name_before
            else:
                name = window[: coord_match.start()].strip(" -")

            note = window[coord_match.end():].strip(" -:")

            rows.append(
                {
                    "code": code,
                    "nom": name,
                    "latitude": str(lat),
                    "longitude": str(lon),
                    "notes": note,
                }
            )

    return rows, fieldnames, header_map


def load_rows_and_header_map(path, log):
    """
    Fonction d'orchestration principale de la lecture :
      0. Si le fichier est un .txt, on utilise le parseur de texte libre dédié
         (voir parse_txt_freeform), qui a sa propre logique de détection.
      1. Sinon (XLSX/CSV), on lit normalement (1ère ligne = en-tête) et on tente
         de reconnaître les colonnes par leur nom.
      2. Si AUCUNE colonne Code/Latitude/Longitude/Coordonnées n'est reconnue par
         son nom, on suppose que le fichier n'a pas d'en-tête du tout : on relit
         le fichier en entier (y compris la 1ère ligne, traitée comme une donnée)
         et on devine le rôle de chaque colonne par son contenu.
      3. Si cette détection automatique échoue aussi, on repart avec les données
         du mode normal : l'erreur usuelle ("colonnes introuvables") sera levée
         plus loin par generate_gpx(), avec le détail des colonnes détectées.

    Retourne (rows, fieldnames, header_map_override) où header_map_override est
    None en mode normal XLSX/CSV (generate_gpx fera lui-même la reconnaissance),
    ou un dict déjà prêt à l'emploi en mode "sans en-tête" ou pour un fichier .txt.
    """
    ext = os.path.splitext(path)[1].lower()

    if ext == ".txt":
        rows, fieldnames, header_map = parse_txt_freeform(path, log)
        log(f"{len(rows)} cache(s) valide(s) détectée(s) dans le fichier texte libre.")
        return rows, fieldnames, header_map

    fieldnames, rows = read_input_file(path)
    header_map = build_header_map(fieldnames)

    has_code = "code" in header_map
    has_coords = ("latitude" in header_map and "longitude" in header_map) or "coordonnees" in header_map

    if has_code or has_coords:
        return rows, fieldnames, None  # Mode normal : rien à faire de spécial

    log(
        "Aucune colonne Code/Latitude/Longitude reconnue dans la ligne d'en-tête : "
        "tentative de détection automatique en supposant que le fichier n'a pas "
        "d'en-tête (analyse du contenu des colonnes)..."
    )
    raw_rows = read_input_file_headerless(path)
    headerless_rows, headerless_map, synthetic_fieldnames = build_headerless_rows_and_mapping(raw_rows)

    if headerless_map is None:
        log("Détection automatique sans en-tête infructueuse (aucun code GC reconnu).")
        return rows, fieldnames, None  # On repart avec le mode normal (qui échouera avec un message clair)

    details = ", ".join(
        f"{k} -> {v}" for k, v in headerless_map.items() if v is not None
    )
    log(f"Détection automatique réussie (fichier sans en-tête) : {details}")
    log(f"({len(headerless_rows)} ligne(s) de données détectée(s), 1ère ligne incluse)")
    return headerless_rows, synthetic_fieldnames, headerless_map


# =============================================================================
# 4. GENERATION DU GPX (format Groundspeak / GSAK)
# =============================================================================

GPX_NS = "http://www.topografix.com/GPX/1/0"
GS_NS = "http://www.groundspeak.com/cache/1/0/1"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"

ET.register_namespace("", GPX_NS)
ET.register_namespace("groundspeak", GS_NS)
ET.register_namespace("xsi", XSI_NS)


def normalize_type(raw_type):
    if not raw_type:
        return "Unknown Cache"
    key = _strip_accents(raw_type)
    return TYPE_ALIASES.get(key, raw_type)


def normalize_container(raw_container):
    if not raw_container:
        return "Not chosen"
    key = _strip_accents(raw_container)
    return CONTAINER_ALIASES.get(key, raw_container)


def normalize_bool(raw_value, default=True):
    if raw_value == "" or raw_value is None:
        return default
    key = _strip_accents(raw_value)
    return key in ("1", "true", "vrai", "yes", "oui", "x", "available", "disponible")


def build_wpt_element(row, header_map, row_number, log):
    """Construit un élément <wpt> GPX/Groundspeak à partir d'une ligne de données.
    Retourne l'élément XML, ou None si la ligne doit être ignorée (avec message de log)."""

    code = get_field(row, header_map, "code")
    name = get_field(row, header_map, "nom")
    lat, lon = parse_coordinates(row, header_map)

    if lat is None or lon is None:
        log(f"  [IGNORÉE] Ligne {row_number} : coordonnées manquantes ou invalides.")
        return None
    if not (-90.0 <= lat <= 90.0) or not (-180.0 <= lon <= 180.0):
        log(
            f"  [IGNORÉE] Ligne {row_number} : coordonnées hors plage valide "
            f"(lat={lat}, lon={lon} — attendu lat entre -90/90 et lon entre -180/180)."
        )
        return None
    if abs(lat) < 1e-9 and abs(lon) < 1e-9:
        log(f"  [IGNORÉE] Ligne {row_number} : coordonnées (0, 0) suspectes, probablement une erreur d'export.")
        return None
    if not code and not name:
        log(f"  [IGNORÉE] Ligne {row_number} : ni code ni nom renseigné.")
        return None
    if not code:
        code = f"WP{row_number:04d}"
    if not name:
        # Repli : si aucune colonne 'Nom' dédiée n'existe (ou qu'elle est vide),
        # on tente d'abord la colonne 'Description', sinon on met un nom fictif
        # 'XXX' pour que la cache soit tout de même importée dans GSAK.
        name = get_field(row, header_map, "description") or "XXX"

    cache_type = normalize_type(get_field(row, header_map, "type"))
    container = normalize_container(get_field(row, header_map, "taille"))
    difficulty = get_field(row, header_map, "difficulte", "1")
    terrain = get_field(row, header_map, "terrain", "1")
    owner = get_field(row, header_map, "proprietaire", "Inconnu")
    country = get_field(row, header_map, "pays")
    region = get_field(row, header_map, "region")
    short_desc = get_field(row, header_map, "description")
    long_desc = get_field(row, header_map, "description_longue")
    hint = get_field(row, header_map, "indice")
    url = get_field(row, header_map, "url", f"https://coord.info/{code}")
    date_raw = get_field(row, header_map, "date")
    available = normalize_bool(get_field(row, header_map, "disponible"), default=True)
    archived = normalize_bool(get_field(row, header_map, "archivee"), default=False)

    try:
        difficulty_f = f"{float(str(difficulty).replace(',', '.')):.1f}"
    except ValueError:
        difficulty_f = "1.0"
    try:
        terrain_f = f"{float(str(terrain).replace(',', '.')):.1f}"
    except ValueError:
        terrain_f = "1.0"

    # Date au format ISO 8601 attendu par GPX
    iso_date = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    if date_raw:
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
            try:
                iso_date = datetime.strptime(date_raw, fmt).strftime("%Y-%m-%dT00:00:00Z")
                break
            except ValueError:
                continue

    wpt = ET.Element(f"{{{GPX_NS}}}wpt", {"lat": f"{lat:.6f}", "lon": f"{lon:.6f}"})

    ET.SubElement(wpt, f"{{{GPX_NS}}}time").text = iso_date
    ET.SubElement(wpt, f"{{{GPX_NS}}}name").text = code
    ET.SubElement(wpt, f"{{{GPX_NS}}}desc").text = f"{name} par {owner}, {cache_type} ({difficulty_f}/{terrain_f})"
    ET.SubElement(wpt, f"{{{GPX_NS}}}url").text = url
    ET.SubElement(wpt, f"{{{GPX_NS}}}urlname").text = name
    ET.SubElement(wpt, f"{{{GPX_NS}}}sym").text = "Geocache Found" if False else "Geocache"
    ET.SubElement(wpt, f"{{{GPX_NS}}}type").text = "Geocache|" + cache_type

    cache = ET.SubElement(
        wpt,
        f"{{{GS_NS}}}cache",
        {
            "id": code.lstrip("GC") or "0",
            "available": "True" if available else "False",
            "archived": "True" if archived else "False",
        },
    )
    ET.SubElement(cache, f"{{{GS_NS}}}name").text = name
    ET.SubElement(cache, f"{{{GS_NS}}}placed_by").text = owner
    ET.SubElement(cache, f"{{{GS_NS}}}owner", {"id": "0"}).text = owner
    ET.SubElement(cache, f"{{{GS_NS}}}type").text = cache_type
    ET.SubElement(cache, f"{{{GS_NS}}}container").text = container
    ET.SubElement(cache, f"{{{GS_NS}}}difficulty").text = difficulty_f
    ET.SubElement(cache, f"{{{GS_NS}}}terrain").text = terrain_f
    ET.SubElement(cache, f"{{{GS_NS}}}country").text = country
    ET.SubElement(cache, f"{{{GS_NS}}}state").text = region
    ET.SubElement(cache, f"{{{GS_NS}}}short_description", {"html": "False"}).text = short_desc
    ET.SubElement(cache, f"{{{GS_NS}}}long_description", {"html": "False"}).text = long_desc
    ET.SubElement(cache, f"{{{GS_NS}}}encoded_hints").text = hint

    # Notes personnelles (colonne UserNote/Note/...) : confirmé par un export réel
    # de GSAK, le champ "Note" (avec l'icône) est en fait stocké comme une entrée
    # de journal spéciale à l'intérieur de <groundspeak:cache>, PAS comme une
    # extension gsak:. GSAK reconnaît cette entrée grâce à son id réservé "-2",
    # son type "Write note" et son auteur "GSAK".
    notes = get_field(row, header_map, "notes")
    if notes:
        logs_el = ET.SubElement(cache, f"{{{GS_NS}}}logs")
        log_el = ET.SubElement(logs_el, f"{{{GS_NS}}}log", {"id": "-2"})
        ET.SubElement(log_el, f"{{{GS_NS}}}date").text = iso_date
        ET.SubElement(log_el, f"{{{GS_NS}}}type").text = "Write note"
        ET.SubElement(log_el, f"{{{GS_NS}}}finder", {"id": "0"}).text = "GSAK"
        ET.SubElement(log_el, f"{{{GS_NS}}}text", {"encoded": "False"}).text = notes

    return wpt


def make_empty_gpx_root(creator=None):
    """Crée un élément racine <gpx> vide, prêt à recevoir des <wpt>."""
    if creator is None:
        creator = APP_NAME
    return ET.Element(
        f"{{{GPX_NS}}}gpx",
        {
            "version": "1.0",
            "creator": creator,
            f"{{{XSI_NS}}}schemaLocation": (
                f"{GPX_NS} {GPX_NS}/gpx.xsd {GS_NS} {GS_NS}/cache.xsd"
            ),
        },
    )


def write_gpx_file(gpx_root, output_path):
    """Sérialise et écrit un élément racine <gpx> dans un fichier."""
    rough_string = ET.tostring(gpx_root, encoding="utf-8")
    pretty = minidom.parseString(rough_string).toprettyxml(indent="  ", encoding="UTF-8")
    with open(output_path, "wb") as f:
        f.write(pretty)


def validate_header_map_has_coords(header_map, fieldnames):
    """Vérifie que le header_map contient de quoi localiser des coordonnées ;
    lève une ValueError explicite (avec le détail des colonnes) sinon."""
    missing_required = [f for f in ("latitude", "longitude") if f not in header_map]
    if "coordonnees" not in header_map and missing_required:
        detected = ", ".join(repr(f) for f in fieldnames) if fieldnames else "(aucune colonne détectée)"
        recognized = ", ".join(f"{k} -> {v!r}" for k, v in header_map.items()) or "(aucune)"
        raise ValueError(
            "Colonnes de coordonnées introuvables. Le fichier doit contenir une colonne "
            "'Latitude' ET une colonne 'Longitude' (ou une colonne 'Coordonnées').\n\n"
            f"Colonnes détectées dans le fichier ({len(fieldnames)}) : {detected}\n"
            f"Colonnes reconnues par le programme : {recognized}\n\n"
            "Si vos colonnes Latitude/Longitude apparaissent ci-dessus dans la liste détectée "
            "mais pas dans la liste reconnue, vérifiez qu'elles ne contiennent pas de caractères "
            "inhabituels, et que le fichier utilise bien un séparateur standard "
            "(virgule, point-virgule ou tabulation) si c'est un CSV."
        )


def build_wpts_from_rows(rows, header_map, log, seen_codes, source_label=""):
    """
    Construit la liste des éléments <wpt> à partir de 'rows', en mettant à jour
    les compteurs et le journal au fil de l'eau.
    `seen_codes` est un dict {code_gc: origine} partagé entre plusieurs appels
    (utilisé pour détecter les doublons ENTRE plusieurs fichiers en mode lot/fusion).
    `source_label` est ajouté aux messages de log (ex: " [fichier.csv]").
    Retourne (liste_wpt, nb_succes, nb_ignorees, nb_notes).
    """
    wpts = []
    success_count = 0
    skipped_count = 0
    notes_count = 0

    for i, row in enumerate(rows, start=1):
        wpt = build_wpt_element(row, header_map, i, log)
        if wpt is None:
            skipped_count += 1
            continue
        wpts.append(wpt)
        success_count += 1
        code_for_log = get_field(row, header_map, "code") or get_field(row, header_map, "nom")
        log(f"  [OK] Ligne {i}{source_label} : {code_for_log} ajoutée.")
        if get_field(row, header_map, "notes"):
            notes_count += 1

        code_val = get_field(row, header_map, "code")
        if code_val:
            if code_val in seen_codes:
                log(
                    f"  [ATTENTION] Ligne {i}{source_label} : le code '{code_val}' apparaît aussi "
                    f"dans {seen_codes[code_val]}. Les deux caches sont incluses dans le GPX."
                )
            else:
                seen_codes[code_val] = source_label.strip(" []") or f"ligne {i}"

    return wpts, success_count, skipped_count, notes_count


def generate_gpx(rows, fieldnames, output_path, log, header_map_override=None):
    """
    Génère le fichier GPX à partir des lignes lues (un seul fichier source).
    `log` est une fonction appelée avec chaque message de progression.
    `header_map_override`, si fourni, est utilisé tel quel à la place de la
    détection par nom de colonne (utilisé pour le mode "sans en-tête").
    Retourne (nb_succes, nb_ignorees, nb_total).
    """
    header_map = header_map_override if header_map_override is not None else build_header_map(fieldnames)
    validate_header_map_has_coords(header_map, fieldnames)

    gpx = make_empty_gpx_root()
    seen_codes = {}
    wpts, success_count, skipped_count, notes_count = build_wpts_from_rows(rows, header_map, log, seen_codes)
    for wpt in wpts:
        gpx.append(wpt)
    total = len(rows)

    write_gpx_file(gpx, output_path)

    if notes_count:
        log(f"{notes_count} note(s) personnelle(s) incluse(s) directement dans le GPX (champ Note de GSAK).")

    return success_count, skipped_count, total


# =============================================================================
# 4bis. TRAITEMENT COMBINÉ : CSV/XLSX/TXT (à convertir) + GPX (à fusionner)
# =============================================================================
# Chaque fichier de la liste est traité selon son propre type : les fichiers
# CSV/XLSX/TXT sont convertis normalement, les fichiers GPX sont fusionnés tels
# quels (sans reformatage) — le tout est combiné dans un seul GPX de sortie.
# Cela couvre naturellement les trois cas : lot de conversion pur, fusion GPX
# pure, ou mélange des deux.

CONVERTIBLE_EXTENSIONS = {".csv", ".xlsx", ".xlsm", ".txt"}
GPX_EXTENSION = ".gpx"


def _strip_whitespace_only_text(elem):
    """Retire récursivement le texte/tail purement blanc (espaces, retours à la
    ligne) d'un élément XML et de ses enfants. Utilisé avant de ré-insérer des
    éléments venant d'un GPX déjà indenté, pour éviter que son indentation
    d'origine ne se cumule avec celle appliquée à la sortie finale."""
    if elem.text is not None and elem.text.strip() == "":
        elem.text = None
    if elem.tail is not None and elem.tail.strip() == "":
        elem.tail = None
    for child in elem:
        _strip_whitespace_only_text(child)


def read_gpx_wpts(path):
    """Lit un fichier GPX existant et retourne la liste de ses éléments <wpt>,
    TELS QUELS sur le fond (aucune modification de leur contenu ni de leurs
    balises), peu importe leur origine (ce script, GSAK, Pocket Query officiel
    geocaching.com...). Seul l'espacement/indentation de présentation est
    nettoyé, pour un ré-assemblage propre dans le fichier fusionné."""
    tree = ET.parse(path)
    root = tree.getroot()
    wpts = []
    for elem in root:
        localname = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        if localname == "wpt":
            _strip_whitespace_only_text(elem)
            wpts.append(elem)
    return wpts


def _get_wpt_gc_code(wpt_elem):
    """Extrait le code GC d'un élément <wpt> (balise <name> ou <n> selon
    l'outil source), utilisé uniquement pour la détection de doublons."""
    for child in wpt_elem:
        localname = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if localname in ("name", "n"):
            return (child.text or "").strip()
    return None


def _wpt_is_full_cache(wpt_elem):
    """Indique si un élément <wpt> est une VRAIE cache Geocaching (il porte
    l'extension <groundspeak:cache> avec ses détails : nom, description,
    indice, attributs...), par opposition à un simple waypoint additionnel
    (stage virtuel, coordonnée corrigée, parking...) exporté par GSAK sous
    forme de <wpt> "enfant" rattaché à une cache via gsak:wptExtension/Parent,
    mais qui n'est pas lui-même une cache."""
    for child in wpt_elem.iter():
        localname = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if localname == "cache":
            return True
    return False


def run_combined_processing(
    file_paths, output_path, log, on_file_progress=None, only_full_caches=False
):
    """
    Traite une liste de fichiers, chacun selon son propre type :
      - .gpx                    -> fusionné tel quel (aucun reformatage)
      - .csv / .xlsx / .txt     -> converti avec toute la logique habituelle
    Le tout est combiné dans un seul GPX de sortie.

    `only_full_caches` (par défaut False, comportement historique inchangé) :
    si True, pour les fichiers .gpx source, seuls les <wpt> qui sont de
    VRAIES caches (avec l'extension <groundspeak:cache>) sont conservés.
    Les waypoints additionnels sans cette extension (stages virtuels,
    coordonnées corrigées, parkings... exportés par GSAK comme <wpt>
    "enfants" d'une cache) sont exclus du GPX final, pour éviter qu'un
    logiciel qui ne reconnaît pas l'extension GSAK gsak:wptExtension/Parent
    ne les affiche comme de fausses caches sans détails. Un waypoint exclu
    n'est retiré que du fichier final ; le fichier source n'est jamais
    modifié, et le nombre exclu est indiqué dans le journal.

    Un fichier qui échoue (colonnes non reconnues, fichier corrompu/illisible,
    etc.) est sauté : le traitement continue avec les fichiers suivants, avec
    un message clair dans le journal.

    Les doublons de code GC ENTRE fichiers sont détectés (mais inclus quand
    même), peu importe le type d'origine des fichiers concernés.

    `on_file_progress(index, total)`, si fourni, est appelé après chaque
    fichier traité (pour une barre de progression).

    Retourne (liste_de_resumes_par_fichier, nb_succes_total, nb_lignes_total).
    Chaque résumé a un champ "kind" ("gpx" ou "convert") en plus du reste.
    """
    gpx = make_empty_gpx_root(creator=f"{APP_NAME} {APP_VERSION}")
    seen_codes = {}
    summaries = []
    total_success = 0
    total_rows = 0
    total_files = len(file_paths)

    for idx, path in enumerate(file_paths, start=1):
        fname = os.path.basename(path)
        ext = os.path.splitext(path)[1].lower()
        log(f"--- Fichier : {fname} ---")
        try:
            if ext == GPX_EXTENSION:
                wpts = read_gpx_wpts(path)
                count = 0
                excluded = 0
                for wpt in wpts:
                    if only_full_caches and not _wpt_is_full_cache(wpt):
                        excluded += 1
                        continue
                    code = _get_wpt_gc_code(wpt)
                    if code:
                        if code in seen_codes:
                            log(
                                f"  [ATTENTION] {fname} : le code '{code}' apparaît aussi dans "
                                f"{seen_codes[code]}. Les deux caches sont incluses dans le GPX."
                            )
                        else:
                            seen_codes[code] = fname
                    gpx.append(wpt)
                    count += 1

                if excluded:
                    log(
                        f"  {excluded} waypoint(s) additionnel(s) (stage/parking/correction, "
                        f"sans détails de cache) exclu(s) du GPX final."
                    )
                log(f"{fname} : OK — {count} cache(s) fusionnée(s) (GPX existant, non modifié).")
                summaries.append(
                    {"file": fname, "status": "OK", "kind": "gpx", "count": count, "excluded": excluded}
                )
                total_success += count
                total_rows += count
            else:
                rows, fieldnames, header_map_override = load_rows_and_header_map(path, log)
                header_map = (
                    header_map_override if header_map_override is not None else build_header_map(fieldnames)
                )
                validate_header_map_has_coords(header_map, fieldnames)

                wpts, success, skipped, notes_count = build_wpts_from_rows(
                    rows, header_map, log, seen_codes, source_label=f" [{fname}]"
                )
                for wpt in wpts:
                    gpx.append(wpt)
                total = len(rows)

                log(f"{fname} : OK — {total} cache(s) attendue(s), {success} générée(s), {skipped} ignorée(s).")
                summaries.append(
                    {"file": fname, "status": "OK", "kind": "convert", "total": total, "success": success, "skipped": skipped}
                )
                total_success += success
                total_rows += total
        except Exception as exc:  # noqa: BLE001
            log(f"{fname} : ÉCHEC — {exc}")
            summaries.append({"file": fname, "status": "ÉCHEC", "error": str(exc)})
        finally:
            if on_file_progress:
                on_file_progress(idx, total_files)

    write_gpx_file(gpx, output_path)
    return summaries, total_success, total_rows


def classify_file_list(paths):
    """
    Détermine s'il y a quelque chose à faire avec la liste de fichiers :
      - "noop"    : un seul fichier .gpx seul dans la liste -> rien à faire,
                    ce fichier existe déjà tel quel.
      - "process" : tout le reste (1 fichier convertible seul, plusieurs
                    fichiers de n'importe quel type/mélange) -> traitement.
      - None + message d'erreur : liste vide.
    """
    if not paths:
        return None, "Veuillez ajouter au moins un fichier à traiter."

    if len(paths) == 1 and os.path.splitext(paths[0])[1].lower() == GPX_EXTENSION:
        return "noop", None

    return "process", None
