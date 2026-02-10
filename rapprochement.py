import streamlit as st
import pandas as pd
import yaml
import unicodedata
import re
from rapidfuzz import fuzz
from io import BytesIO

# -----------------------------
# CHARGEMENT CONFIG PRIVÉE
# -----------------------------
@st.cache_data
def load_config():
    with open("config.yaml", "r") as f:
        return yaml.safe_load(f)

CONFIG = load_config()

# -----------------------------
# FONCTIONS UTILITAIRES
# -----------------------------
def normalize_text(text):
    if pd.isna(text):
        return ""
    text = str(text).upper()
    text = unicodedata.normalize("NFD", text)
    text = text.encode("ascii", "ignore").decode("utf-8")
    text = re.sub(r"[^A-Z0-9 ]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text

def fuzzy_compare(a, b, threshold):
    if not a or not b:
        return False, 0
    score = fuzz.token_sort_ratio(a, b)
    return score >= threshold, score

# -----------------------------
# FONCTION PRINCIPALE
# -----------------------------
def process_files(file1, file2):
    config = CONFIG

    key = config["key"]
    strict_columns = config["strict_columns"]
    fuzzy_columns = config["fuzzy_columns"]
    threshold = config["fuzzy_threshold"]

    # Lecture fichiers Excel
    df1 = pd.read_excel(file1, dtype=str)
    df2 = pd.read_excel(file2, dtype=str)

    # Vérification colonnes
    required_cols = [key] + strict_columns + fuzzy_columns
    missing_1 = [c for c in required_cols if c not in df1.columns]
    missing_2 = [c for c in required_cols if c not in df2.columns]

    if missing_1 or missing_2:
        error_msg = ""
        if missing_1:
            error_msg += f"❌ Colonnes manquantes dans le fichier 1 : {missing_1}\n"
        if missing_2:
            error_msg += f"❌ Colonnes manquantes dans le fichier 2 : {missing_2}\n"
        raise ValueError(error_msg)

    # Normalisation
    for col in required_cols:
        df1[col + "_norm"] = df1[col].apply(normalize_text)
        df2[col + "_norm"] = df2[col].apply(normalize_text)

    # Merge
    df = df1.merge(
        df2,
        on=key,
        how="outer",
        suffixes=("_1", "_2"),
        indicator=True
    )

    # Comparaisons
    for col in strict_columns:
        df[f"{col}_identique"] = df[col + "_norm_1"] == df[col + "_norm_2"]

    for col in fuzzy_columns:
        identiques = []
        scores = []
        for _, row in df.iterrows():
            ok, score = fuzzy_compare(
                row.get(col + "_norm_1", ""),
                row.get(col + "_norm_2", ""),
                threshold
            )
            identiques.append(ok)
            scores.append(score)

        df[f"{col}_identique"] = identiques
        df[f"{col}_score"] = scores

    # Statut général
    def ligne_statut(row):
        if row["_merge"] != "both":
            return "Manquant"
        checks = [row[f"{c}_identique"] for c in strict_columns + fuzzy_columns]
        return "OK" if all(checks) else "Différence"

    df["statut"] = df.apply(ligne_statut, axis=1)

    # Export Excel en mémoire
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df[df["statut"] == "OK"].to_excel(writer, "Correspondances_OK", index=False)
        df[df["statut"] == "Différences"].to_excel(writer, "Différences", index=False)
        df[df["statut"] == "Manquants"].to_excel(writer, "Manquants", index=False)

    output.seek(0)
    return output

# -----------------------------
# INTERFACE STREAMLIT
# -----------------------------
st.set_page_config(page_title="Rapprochement automatique", layout="centered")

st.title("🧩 Rapprochement automatique Excel")
st.write(
    "Téléversez **deux fichiers Excel**. "
    "Les règles de rapprochement sont gérées automatiquement."
)

file1 = st.file_uploader("📄 Fichier Excel 1", type=["xlsx"])
file2 = st.file_uploader("📄 Fichier Excel 2", type=["xlsx"])

if file1 and file2 and config_file:
    with st.spinner("⏳ Patientez... La personne tourne en rond et n'aime pas attendre 😅"):
        st.image("https://media.giphy.com/media/3o7TKtnuHOHHUjR38Y/giphy.gif", width=120)
        try:
            output_file = process_files(file1, file2)
            st.success("✅ Rapprochement terminé avec succès")
            st.download_button(
                "📥 Télécharger le fichier résultat",
                data=output_file,
                file_name="rapprochement_resultat.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(str(e))


