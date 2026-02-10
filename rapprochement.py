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
    for col in strict_columns + fuzzy_columns + [key]:
        df1[col + "_norm"] = df1[col].astype(str).apply(normalize_text)
        df2[col + "_norm"] = df2[col].astype(str).apply(normalize_text)

    # Merge
    df = df1.merge(
        df2,
        on=key,
        how="outer",
        suffixes=("_1", "_2"),
        indicator=True
    )

    # Comparaisons
    df = df1.merge(df2, on=key, how="outer", suffixes=("_1", "_2"), indicator=True)

    # Comparaisons EXACTEMENT COMME LE SCRIPT ORIGINAL
    comparisons = {}
    scores = {}

    # Colonnes strictes
    for col in strict_columns:
        df[f"{col}_identique"] = df[col + "_norm_1"] == df[col + "_norm_2"]
        comparisons[col] = df[col + "_norm_1"] == df[col + "_norm_2"]

    # Colonnes fuzzy
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
        comp_list = []
        score_list = []
        for idx, row in df.iterrows():
            val_1 = row.get(col + "_norm_1", "")
            val_2 = row.get(col + "_norm_2", "")
            comp, score = fuzzy_compare(val_1, val_2, threshold)
            comp_list.append(comp)
            score_list.append(score)
        comparisons[col] = pd.Series(comp_list)
        scores[col] = pd.Series(score_list)

    # Ajout résultats
    for col in strict_columns + fuzzy_columns:
        df[f"{col}_identique"] = comparisons[col]

        df[f"{col}_identique"] = identiques
        df[f"{col}_score"] = scores
    for col in fuzzy_columns:
        df[f"{col}_score"] = scores[col]

    # Statut général
    def ligne_statut(row):
        if row["_merge"] != "both":
            return "Manquant"
        checks = [row[f"{c}_identique"] for c in strict_columns + fuzzy_columns]
        return "OK" if all(checks) else "Différence"
        cols_check = [f"{c}_identique" for c in strict_columns + fuzzy_columns]
        if all(row[col] for col in cols_check):
            return "OK"
        else:
            return "Différence"

    df["statut"] = df.apply(ligne_statut, axis=1)

    # Export Excel en mémoire
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df[df["statut"] == "OK"].to_excel(writer, "Correspondances_OK", index=False)
        df[df["statut"] == "Différences"].to_excel(writer, "Différences", index=False)
        df[df["statut"] == "Manquants"].to_excel(writer, "Manquants", index=False)

        df[df["statut"] == "OK"].to_excel(writer, sheet_name="Correspondances_OK", index=False)
        df[df["statut"] == "Différence"].to_excel(writer, sheet_name="Différences", index=False)
        df[df["statut"] == "Manquant"].to_excel(writer, sheet_name="Manquants", index=False)
    output.seek(0)
    return output

@@ -141,4 +133,4 @@
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(str(e))
            st.error(f"❌ Erreur : {e}")

