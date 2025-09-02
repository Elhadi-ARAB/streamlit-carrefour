
import streamlit as st
import pandas as pd
import os
import re
import time
import urllib.parse
import unicodedata






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


ALLOWED_PLATFORMS = ["DV360", "Ogury", "Seedtag_FR", "Spotify_FRA", "M6+"]


def map_tracking_type(value):
    val = str(value).strip()
    if val in video_keywords:
        return "VIDEO"
    elif val == "Stream-Audio":
        return "AUDIO"
    elif val in display_keywords:
        return "DISPLAY"
    else:
        print(f"⚠️ Format Type non reconnu: {val}. Contactez l'équipe Data.")
        return "OTHER"

    
    
    

_SPECIALS = set(list('!"#$%&\'()*+,/:;<=?>@[\\]^{|}~%'))  # caractères à remplacer par "_"
_UTM_KEYS = {"utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content"}

def _strip_accents(s: str) -> str:
    """Supprime les accents de manière générique (é→e, ñ→n, etc.)."""
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in nfkd if not unicodedata.combining(ch))

def _sanitize_text_for_utm(s: str) -> str:
    if s is None:
        return s
    s = str(s)
    s = _strip_accents(s)
    # 🔁 remplace espaces ET caractères spéciaux par "_"
    s = "".join("_" if (ch in _SPECIALS or ch.isspace()) else ch for ch in s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s

def _idna_hostname(host: str) -> str:
    """Normalise le nom d’hôte (lowercase + IDNA)."""
    if not host:
        return host
    host = host.strip().lower()
    try:
        return host.encode("idna").decode("ascii")
    except Exception:
        return host

def _normalize_path(path: str) -> str:
    """
    Normalise le chemin :
    - évite le double-encodage des % déjà présents
    - encode les caractères non sûrs selon RFC 3986
    """
    if path is None or path == "":
        return "/"
    # on “décode” une fois puis on ré-encode proprement
    unquoted = urllib.parse.unquote(path)
    # safe: on garde / % et les caractères réservés usuels en path
    return urllib.parse.quote(unquoted, safe="/%:@-._~!$&'()*+,;=")

def _encode_q_value(v: str) -> str:
    """Encode standard pour les valeurs de query (hors UTM)."""
    if v is None:
        return ""
    # on évite de ré-encoder les % des séquences %XX déjà valides
    return urllib.parse.quote_plus(v, safe="%:@-._~!$'()*,")

def normalize_url_preserving_utm(url: str) -> str:
    """
    1) Normalise schéma/hôte/chemin.
    2) Re-encode proprement TOUTES les paires query, SAUF les valeurs UTM.
       - UTM : on ne modifie pas le texte, on échappe seulement & et = pour ne pas casser la query.
       - Non-UTM : encodage standard (quote_plus).
    """
    if not url or not isinstance(url, str):
        return url

    u = url.strip()
    if u.startswith("www."):
        u = "https://" + u  # ajout d’un schéma par défaut si nécessaire

    parts = urllib.parse.urlsplit(u)

    scheme = (parts.scheme or "https").lower()
    netloc = _idna_hostname(parts.netloc)
    path = _normalize_path(parts.path)

    # Parse query, en conservant l’ordre et les valeurs vides
    pairs = urllib.parse.parse_qsl(parts.query, keep_blank_values=True)

    new_pairs = []
    for k, v in pairs:
        k_enc = urllib.parse.quote(k, safe="%:@-._~")  # clés encodées proprement

        if k.lower() in _UTM_KEYS:
            v_safe = (v or "")
            v_safe = v_safe.replace("&", "%26").replace("=", "%3D")
            if v_safe == "":
                new_pairs.append(f"{k_enc}")            # ❗ pas de "=" si vide
            else:
                new_pairs.append(f"{k_enc}={v_safe}")
        else:
            v_enc = _encode_q_value(v or "")
            if v_enc == "":
                new_pairs.append(f"{k_enc}")            # ❗ pas de "=" si vide
            else:
                new_pairs.append(f"{k_enc}={v_enc}")


    new_query = "&".join(new_pairs)

    # Fragment : on l’encode proprement sans double-encodage
    fragment = urllib.parse.quote(urllib.parse.unquote(parts.fragment or ""), safe="%:@-._~!$&'()*+,;=/")

    return urllib.parse.urlunsplit((scheme, netloc, path, new_query, fragment))

def sanitize_url_utm_values(url: str) -> str:
    """
    2) Ne modifie QUE les valeurs des paramètres UTM.
       Conserve schéma, hôte, chemin et autres paramètres intacts.
    """
    if not url or not isinstance(url, str):
        return url
    try:
        parts = urllib.parse.urlsplit(url)
        pairs = urllib.parse.parse_qsl(parts.query, keep_blank_values=True)

        new_pairs = []
        for k, v in pairs:
            k_enc = urllib.parse.quote(k, safe="%:@-._~")
            if k.lower() in _UTM_KEYS:
                raw = _sanitize_text_for_utm(v)
                raw = (raw or "").replace("&", "%26").replace("=", "%3D")
                if raw == "":
                    new_pairs.append(f"{k_enc}")            # ❗ pas de "=" si vide
                else:
                    new_pairs.append(f"{k_enc}={raw}")
            else:
                if v == "":
                    new_pairs.append(f"{k_enc}")            # ❗ pas de "=" si vide
                else:
                    new_pairs.append(f"{k_enc}={v}")

        new_query = "&".join(new_pairs)
        return urllib.parse.urlunsplit((parts.scheme, parts.netloc, parts.path, new_query, parts.fragment))
    except Exception:
        return url


def generate_files(df, output_folder="exports_cm"):
    os.makedirs(output_folder, exist_ok=True)
    grouped = df.groupby(["Nom de la campagne*", "Début  JJ/MM/AAAA", "Fin  JJ/MM/AAAA"])
    filepaths = []
    ignored_platforms = []   # plateformes non autorisées
    unknown_tracking = []    # tracking type non reconnus

    for (campagne, start, end), group in grouped:
        if pd.isna(campagne) or pd.isna(start) or pd.isna(end):
            continue

        header_data = {
            "Consultant_Email": "",
            "Campaign_Name": campagne,
            "Advertiser Name": "NEW Carrefour One - Havas",
            "Landing page": "https://www.carrefour.fr/",
            "Start_Date": start,
            "End_Date": end
        }

        rows = []
        for _, row in group.iterrows():
            if pd.isna(row["Format size (CM only)"]):
                continue

            # ➜ contrôle plateforme
            platform = str(row["Platform"]).strip()
            if platform not in ALLOWED_PLATFORMS:
                ignored_platforms.append((campagne, platform))
                continue

            # ➜ mapping spécifique DV360
            if platform == "DV360":
                site_name = "DV360 - CRF Marketing - Agence 79"
            elif platform == "TF1":
                site_name = "TF1_FRA"
            elif platform == "SNCF":
                site_name = "Groupe Sncf"
            else:
                site_name = platform


            # ➜ contrôle Tracking Type
            raw_tracking = row.get("Tracking Type", "")
            tracking_type = map_tracking_type(raw_tracking)
            if tracking_type == "OTHER":
                unknown_tracking.append((campagne, str(raw_tracking).strip()))
                continue  # on n’inclut pas cette ligne dans l’export

            original_url = row.get("URL de redirection trackée")
            formatted_url = normalize_url_preserving_utm(original_url)
            safe_url = sanitize_url_utm_values(formatted_url)

            rows.append({
                "Site Name": site_name,
                "Type de Tracking": tracking_type,
                "Format": row["Format size (CM only)"],
                "Placement_Name": row["Placement CM = creative DV360"],
                "Creative_Name": row["Creative Name"],
                "Add_ClickTroughUrl": safe_url,
                "IMPRESSION_JAVASCRIPT_EVENT_TAG": row["XPLN script CM"]
            })

        if not rows:
            continue

        df_table = pd.DataFrame(rows)
        safe_name = str(campagne).replace(" ", "_").replace("/", "_")
        filename = f"CM_{safe_name}_carrefour.xlsx"
        filepath = os.path.join(output_folder, filename)

        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            for i, (key, val) in enumerate(header_data.items()):
                pd.DataFrame([[key, val]]).to_excel(writer, index=False, header=False, startrow=i, startcol=0)
            df_table.to_excel(writer, index=False, startrow=9)

        filepaths.append(filepath)

    # ➜ on renvoie les deux rapports d’erreurs
    return filepaths, ignored_platforms, unknown_tracking


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
                paths, ignored_platforms, unknown_tracking = generate_files(df)
                duration = round(time.time() - start, 2)

            has_error = False

            # ⚠️ Plateformes non autorisées
            if ignored_platforms:
                has_error = True
                st.error("❌ Des lignes ont été ignorées car la plateforme n’est pas autorisée :")
                for campagne, plat in ignored_platforms:
                    st.write(f"- Campagne **{campagne}** → Plateforme non reconnue : **{plat}**")

            # ⚠️ Tracking Type non reconnu
            if unknown_tracking:
                has_error = True
                st.error("❌ Des lignes ont été ignorées car le 'Tracking Type' est non reconnu :")
                for campagne, trk in unknown_tracking:
                    st.write(f"- Campagne **{campagne}** → Tracking Type non reconnu : **{trk}**")
                st.info("ℹ️ Valeurs acceptées : vérifiez dans l’onglet Input A79 de l’URL Builder si le Tracking Type y est bien renseigné")

            if has_error:
                st.warning("🚫 Aucun fichier n’a été généré. Corrigez le fichier source puis réessayez.")
            else:
                st.success(f"✅ {len(paths)} fichiers générés en {duration} secondes.")
                st.info("ℹ️ N’oubliez pas d’introduire votre adresse mail en colonne B (ligne Consultant_Email). Merci 🙏")
                for p in paths:
                    with open(p, "rb") as f:
                        data = f.read()
                    st.download_button(label=f"Télécharger {os.path.basename(p)}", data=data, file_name=os.path.basename(p))


    except Exception as e:
        st.error(f"❌ Erreur lors du traitement du fichier : {e}")

