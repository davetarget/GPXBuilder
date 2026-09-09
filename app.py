#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GPXBuilder — version Streamlit
================================
Interface web (Streamlit) pour le convertisseur GPXBuilder, réutilisant
telle quelle toute la logique métier du script original (gpxbuilder_core.py) :
parsing XLSX/CSV/TXT, détection automatique des colonnes, fusion de GPX,
génération du GPX final compatible GSAK.

Lancement local :
    pip install -r requirements.txt
    streamlit run app.py

Déploiement : compatible tel quel avec streamlit.io (Streamlit Community Cloud) —
il suffit de pousser ce dossier (app.py + gpxbuilder_core.py + requirements.txt)
sur un dépôt GitHub puis de le connecter sur share.streamlit.io.
"""

import os
import shutil
import tempfile
from datetime import datetime

import streamlit as st

import gpxbuilder_core as core

st.set_page_config(
    page_title=f"{core.APP_NAME} v{core.APP_VERSION}",
    page_icon="🧭",
    layout="centered",
)

# -----------------------------------------------------------------------
# Etat de session
# -----------------------------------------------------------------------
if "log_lines" not in st.session_state:
    st.session_state.log_lines = []
if "result_bytes" not in st.session_state:
    st.session_state.result_bytes = None
if "result_name" not in st.session_state:
    st.session_state.result_name = None
if "summaries" not in st.session_state:
    st.session_state.summaries = None

def log(message):
    st.session_state.log_lines.append(str(message))


# -----------------------------------------------------------------------
# En-tête
# -----------------------------------------------------------------------
st.title(f"🧭 {core.APP_NAME}")
st.caption(
    f"v{core.APP_VERSION} — Convertisseur de fichiers Excel / CSV / TXT / GPX "
    "vers GPX compatible GSAK"
)

with st.expander("ℹ️ Aide rapide"):
    st.text(core.HELP_TEXT)

st.divider()

# -----------------------------------------------------------------------
# 1. Fichiers à traiter
# -----------------------------------------------------------------------
st.subheader("1. Fichiers à traiter")
uploaded_files = st.file_uploader(
    "Ajoutez un ou plusieurs fichiers (.xlsx, .xlsm, .csv, .txt, .gpx)",
    type=["xlsx", "xlsm", "csv", "txt", "gpx"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.write(f"**{len(uploaded_files)} fichier(s)** ajouté(s) :")
    for f in uploaded_files:
        st.write(f"- {f.name} ({f.size} octets)")

    paths_only = [f.name for f in uploaded_files]
    mode, error = core.classify_file_list(paths_only)
    if error:
        st.error(error)
    elif mode == "noop":
        st.info(
            "Ce fichier .gpx existe déjà tel quel : il n'y a rien à convertir ni à "
            "fusionner. Ajoutez au moins un autre fichier si vous voulez le combiner "
            "avec d'autres données."
        )
    else:
        exts = sorted({os.path.splitext(f.name)[1].lower() for f in uploaded_files})
        if len(uploaded_files) == 1:
            mode_desc = "Conversion simple d'un seul fichier."
        elif exts == [".gpx"]:
            mode_desc = "Fusion de plusieurs fichiers GPX existants (aucun reformatage)."
        elif ".gpx" not in exts:
            mode_desc = "Conversion de plusieurs fichiers puis fusion en un seul GPX."
        else:
            mode_desc = "Mode mixte : les .gpx sont fusionnés tels quels, les autres sont convertis — le tout combiné dans un seul GPX."
        st.caption(f"Mode détecté automatiquement : {mode_desc}")
else:
    mode = None

st.divider()

# -----------------------------------------------------------------------
# 2. Nom du fichier de sortie
# -----------------------------------------------------------------------
st.subheader("2. Fichier GPX généré")
now = datetime.now()
default_name = f"geocaches_{now.strftime('%d%m%y_%H%M%S')}.gpx"
st.caption(f"Nom automatique proposé : `{default_name}`")
custom_name = st.text_input(
    "Vous pouvez modifier le nom du fichier de sortie si besoin",
    value=default_name,
)
if not custom_name.lower().endswith(".gpx"):
    custom_name += ".gpx"

st.divider()

# -----------------------------------------------------------------------
# 3. Traitement
# -----------------------------------------------------------------------
st.subheader("3. Traitement")

start = st.button(
    "🚀 Démarrer",
    type="primary",
    disabled=not uploaded_files or mode not in ("process",),
)

if start:
    st.session_state.log_lines = []
    st.session_state.result_bytes = None
    st.session_state.result_name = None
    st.session_state.summaries = None

    tmp_dir = tempfile.mkdtemp(prefix="gpxbuilder_")
    try:
        # Ecrit les fichiers uploadés sur disque (chemins requis par la logique
        # existante, inchangée par rapport au script original) en conservant
        # leur nom d'origine.
        saved_paths = []
        for f in uploaded_files:
            dest = os.path.join(tmp_dir, f.name)
            with open(dest, "wb") as out:
                out.write(f.getbuffer())
            saved_paths.append(dest)

        out_path = os.path.join(tmp_dir, custom_name)

        with st.spinner("Traitement en cours..."):
            log(f"Traitement de {len(saved_paths)} fichier(s)...")
            summaries, total_success, total_rows = core.run_combined_processing(
                saved_paths, out_path, log
            )

            log("-" * 60)
            log("Résumé par fichier :")
            for s in summaries:
                if s["status"] == "OK":
                    if s["kind"] == "gpx":
                        log(f"  - {s['file']} : OK ({s['count']} cache(s) fusionnée(s))")
                    else:
                        log(
                            f"  - {s['file']} : OK ({s['success']}/{s['total']} cache(s), "
                            f"{s['skipped']} ignorée(s))"
                        )
                else:
                    log(f"  - {s['file']} : ÉCHEC ({s['error']})")

            nb_echecs = sum(1 for s in summaries if s["status"] == "ÉCHEC")
            log("-" * 60)
            log(
                f"Terminé : {total_success} cache(s) au total dans le GPX "
                f"({total_rows} ligne(s)/entrée(s) traitée(s))."
            )
            if nb_echecs:
                log(f"{nb_echecs} fichier(s) en échec (voir détails ci-dessus).")

            with open(out_path, "rb") as f:
                st.session_state.result_bytes = f.read()
            st.session_state.result_name = custom_name
            st.session_state.summaries = summaries

        if total_success == 0:
            st.warning(
                f"Traitement terminé, mais 0 cache dans le GPX final "
                f"({total_rows} ligne(s) traitée(s)). Voir le journal ci-dessous."
            )
        else:
            st.success(
                f"✅ {total_success} cache(s) au total dans le GPX "
                f"({total_rows} ligne(s)/entrée(s) traitée(s))."
            )
    except Exception as exc:  # noqa: BLE001
        log(f"[ERREUR] {exc}")
        st.error(f"Échec du traitement : {exc}")
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

# -----------------------------------------------------------------------
# Téléchargement du résultat
# -----------------------------------------------------------------------
if st.session_state.result_bytes:
    st.download_button(
        "⬇️ Télécharger le fichier GPX",
        data=st.session_state.result_bytes,
        file_name=st.session_state.result_name,
        mime="application/gpx+xml",
        type="primary",
    )

# -----------------------------------------------------------------------
# Journal
# -----------------------------------------------------------------------
if st.session_state.log_lines:
    st.subheader("Journal")
    log_text = "\n".join(st.session_state.log_lines)
    st.text_area("Détail du traitement", value=log_text, height=300)
    st.download_button(
        "📋 Enregistrer le journal (.txt)",
        data=log_text,
        file_name="journal_gpxbuilder.txt",
        mime="text/plain",
    )

st.divider()
st.caption(
    "Note : contrairement à la version bureau, cette version web ne mémorise pas "
    "de dossier de destination — chaque traitement se termine par un téléchargement "
    "direct du fichier GPX généré."
)
