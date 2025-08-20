import streamlit as st
import pandas as pd
import os
import time
import urllib.parse





# === Listes de mots-clés pour le mapping ===
video_keywords = [
    "In-Stream-Classic", "In-Stream-Trueview", "Video-Dynamic-Creative", "Connected-TV",
    "Synchro-TV", "In-Stream-trueview-for-reach", "In-Stream-trueview-for-action",
    "In-Stream-trueview-for-lead", "In-Stream-non-skippable", "In-Stream-bumper",
    "In-Stream-catch-up", "Out-Stream-inread", "Out-Stream-pave-video", "Video-Ad-Serving-Template"
]

display_keywords = [
    "Display-IAB", "Display-Native", "Redirect", "Display-Social", "Display-Dynamic-Creative",
    "Digital-Out-of-Home", "Emailing", "Display-panorama", "Display-scratch", "Display-slider",
    "Display-inread", "Display-swipe-to-site", "Display-habillage", "Display-habillage-video",
    "Display-habillage-sliding", "Display-habillage-swapping", "Pixel+Click-Command"
]

def map_tracking_type(value):
    val = str(value).strip()
    if val in video_keywords:
        return "VIDEO"
    elif val == "Stream-Audio":
        return "AUDIO"
    elif val in display_keywords:
        return "DISPLAY"
    else:
        return "OTHER"
    
    
    
# --- Sanitize des valeurs UTM uniquement ---


_SPECIALS = set(list('!"#$%&\'()*+,/:;<=?>@[\\]^`{|}~'))  # caractères à remplacer
_UTM_KEYS = {"utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content"}

def _sanitize_text_for_utm(s: str) -> str:
    if s is None:
        return s
    s = str(s)
    # convertir é/è -> e (et majuscules)
    s = (s.replace("é", "e").replace("è", "e")
           .replace("É", "E").replace("È", "E"))
    # remplacer les caractères spéciaux listés par "_"
    s = "".join("_" if ch in _SPECIALS else ch for ch in s)
    return s

def sanitize_url_utm_values(url: str) -> str:
    """
    Ne modifie que les valeurs des paramètres UTM.
    Conserve schéma, hôte, chemin et autres paramètres intacts.
    Reconstruction manuelle de la query pour ne pas ré-encoder nos underscores.
    """
    if not url or not isinstance(url, str):
        return url
    try:
        parts = urllib.parse.urlsplit(url)
        query_pairs = urllib.parse.parse_qsl(parts.query, keep_blank_values=True)

        # Nettoyer uniquement les valeurs des UTM
        new_pairs = []
        for k, v in query_pairs:
            if k.lower() in _UTM_KEYS:
                v = _sanitize_text_for_utm(v)
            # ⚠️ k et v sont déjà “propres” pour notre besoin : on reconstruit sans urlencode
            # mais on protège quand même le signe & dans la valeur pour ne pas casser le split
            v = str(v).replace("&", "%26")
            new_pairs.append(f"{k}={v}")

        new_query = "&".join(new_pairs)
        sanitized = urllib.parse.urlunsplit((parts.scheme, parts.netloc, parts.path, new_query, parts.fragment))
        return sanitized
    except Exception:
        # En cas d'URL non parseable, on retourne tel quel pour ne pas la casser
        return url

def generate_files(df, output_folder="exports_cm"):
    os.makedirs(output_folder, exist_ok=True)
    grouped = df.groupby(["Nom de la campagne*", "Début  JJ/MM/AAAA", "Fin  JJ/MM/AAAA"])
    filepaths = []

    # ✅ Liste des sites à insérer dans le 2e onglet
    site_list = [
        "M6_FRA", "TF1_FRA", "Socialyse_FRA", "Canal +_FRA", "Youtube_FRA", "BidManager_DfaSite_563002",
        "Spotify_FRA", "nrj.fr", "Lagardere_FRA", "DV360 - Jellyfish - Local",
        "DART Search : Google : 506200", "DV360 - Jellyfish - Traiteur", "Altice Media FR",
        "DV360 - Jellyfish - Mini Sites", "Yummipets", "DV360 - Jellyfish - Branding",
        "191 Media", "DV360 - Jellyfish - Voyage", "Criteo", "Snapchat FR", "Pinterest FR",
        "DV360 - Jellyfish - Marque", "FR_NRJGlobal", "DV360 - Jellyfish - Cartes Cadeaux",
        "Azerion FR", "Carrefour Energies", "DV360 - CRF Marketing - Agence 79", "Mappy",
        "SoundCast", "Groupe Sncf", "Seedtag_FR", "BonjourRATP", "366", "Viewpay", "Ogury",
        "LeBonCoin", "TheBradery"
    ]

    for (campagne, start, end), group in grouped:
        if pd.isna(campagne) or pd.isna(start) or pd.isna(end):
            continue

        header_data = {
            "Consultant_Email": "",
            "Campaign_Name": campagne,
            "Advertiser Name": "NEW Carrefour One - Havas",
            "Landing page": "https://www.carrefour.fr/",
            "Start_Date": start.strftime('%d/%m/%Y'),
            "End_Date": end.strftime('%d/%m/%Y')
        }

        rows = []
        for _, row in group.iterrows():
            if pd.isna(row["Format size (CM only)"]):
                continue

            site_name = "DV360 - CRF Marketing - Agence 79" if str(row["Régie"]).strip() == "Open" else str(row["Régie"]).strip()
            tracking_type = map_tracking_type(row["Tracking Type"])
            
            # ✅ URL nettoyée (UTM uniquement)
            original_url = row.get("URL de redirection trackée")
            safe_url = sanitize_url_utm_values(original_url)


            rows.append({
                "Site Name": site_name,
                "Type de Tracking": tracking_type,
                "Format": row["Format size (CM only)"],
                "Placement_Name": row["Placement CM = creative DV360"],
                "Creative_Name": row["Creative Name"],
                "Add_ClickTroughUrl": safe_url,
                "IMPRESSION_JAVASCRIPT_EVENT_TAG" : row["XPLN script CM"]
            })

        if not rows:
            continue

        df_table = pd.DataFrame(rows)
        df_sites_sheet = pd.DataFrame(site_list, columns=["Site Name"])

        safe_name = campagne.replace(" ", "_").replace("/", "_")
        filename = f"CM_{safe_name}_carrefour.xlsx"
        filepath = os.path.join(output_folder, filename)

        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            # Onglet principal (données de campagne)
            for i, (key, val) in enumerate(header_data.items()):
                pd.DataFrame([[key, val]]).to_excel(writer, index=False, header=False, startrow=i, startcol=0)
            df_table.to_excel(writer, index=False, startrow=9)

            # ➕ Ajout de l'onglet Sites
            df_sites_sheet.to_excel(writer, index=False, sheet_name="Sites")

        filepaths.append(filepath)

    return filepaths

# === Interface Streamlit ===

st.title("📊 Générateur de fichiers Media Carrefour")

uploaded_file = st.file_uploader("Dépose ici le fichier Excel de l'URL Builder", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="URL BUILDER", skiprows=1)
        df = df.dropna(how='all')
        df["Début  JJ/MM/AAAA"] = pd.to_datetime(df["Début  JJ/MM/AAAA"], errors='coerce')
        df["Fin  JJ/MM/AAAA"] = pd.to_datetime(df["Fin  JJ/MM/AAAA"], errors='coerce')
        df = df.dropna(subset=["Nom de la campagne*", "Début  JJ/MM/AAAA", "Fin  JJ/MM/AAAA", "Format size (CM only)"])

        st.success("✅ Fichier chargé avec succès.")

        if st.button("🚀 Générer les fichiers"):
            with st.spinner("⏳ Génération des fichiers en cours..."):
                start = time.time()
                paths = generate_files(df)
                duration = round(time.time() - start, 2)

            st.success(f"✅ {len(paths)} fichiers générés en {duration} secondes.")
            st.info("ℹ️ N’oubliez pas d’introduire votre adresse mail en colonne B (ligne Consultant_Email). Merci 🙏")

            for p in paths:
                with open(p, "rb") as f:
                    data = f.read()
                st.download_button(label=f"Télécharger {os.path.basename(p)}", data=data, file_name=os.path.basename(p))

    except Exception as e:
        st.error(f"❌ Erreur lors du traitement du fichier : {e}")

