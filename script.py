import streamlit as st
import pandas as pd
import os
import time




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

def generate_files(df, output_folder="exports_cm"):
    os.makedirs(output_folder, exist_ok=True)
    grouped = df.groupby(["Nom de la campagne*", "Début  JJ/MM/AAAA", "Fin  JJ/MM/AAAA"])
    filepaths = []

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

            rows.append({
                "Site Name": site_name,
                "Type de Tracking": tracking_type,
                "Format": row["Format size (CM only)"],
                "Placement_Name": row["Placement CM = creative DV360"],
                "Creative_Name": row["Creative Name"],
                "Add_ClickTroughUrl": row["URL de redirection trackée"]
            })

        if not rows:
            continue

        df_table = pd.DataFrame(rows)
        safe_name = campagne.replace(" ", "_").replace("/", "_")
        filename = f"CM_{safe_name}_carrefour.xlsx"
        filepath = os.path.join(output_folder, filename)

        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            for i, (key, val) in enumerate(header_data.items()):
                pd.DataFrame([[key, val]]).to_excel(writer, index=False, header=False, startrow=i, startcol=0)
            df_table.to_excel(writer, index=False, startrow=9)

        filepaths.append(filepath)

    return filepaths

# === Interface Streamlit ===

st.title("📊 Générateur de fichiers Media Carrefour")

uploaded_file = st.file_uploader("Dépose ici le fichier Excel du URL Builder", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="A79 - Display URL BUILDER", skiprows=1)
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

            for p in paths:
                with open(p, "rb") as f:
                    data = f.read()
                st.download_button(label=f"Télécharger {os.path.basename(p)}", data=data, file_name=os.path.basename(p))

    except Exception as e:
        st.error(f"❌ Erreur lors du traitement du fichier : {e}")
