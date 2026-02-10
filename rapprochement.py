import pandas as pd
import unicodedata
import re
from rapidfuzz import fuzz
import yaml

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

def fuzzy_compare(a, b, threshold=85):
    if not a or not b:
        return False, 0
    score = fuzz.token_sort_ratio(a, b)
    return score >= threshold, score

# -----------------------------
# LECTURE DU FICHIER DE CONFIG
# -----------------------------
with open("config.yaml", "r") as f:
    config = yaml.safe_load(f)

file_1 = config["file_1"]
file_2 = config["file_2"]
key = config["key"]
strict_columns = config["strict_columns"]
fuzzy_columns = config["fuzzy_columns"]
threshold = config["fuzzy_threshold"]
output_file = config["output_file"]

# -----------------------------
# CHARGEMENT DES FICHIERS
# -----------------------------
df1 = pd.read_excel(file_1, dtype=str)
df2 = pd.read_excel(file_2, dtype=str)

# -----------------------------
# NORMALISATION
# -----------------------------
for col in strict_columns + fuzzy_columns + [key]:
    df1[col + "_norm"] = df1[col].astype(str).apply(normalize_text)
    df2[col + "_norm"] = df2[col].astype(str).apply(normalize_text)

# -----------------------------
# MERGE
# -----------------------------
df = df1.merge(df2, on=key, how="outer", suffixes=("_1", "_2"), indicator=True)

# -----------------------------
# COMPARAISON
# -----------------------------
comparisons = {}
scores = {}

# Colonnes strictes
for col in strict_columns:
    comparisons[col] = df[col + "_norm_1"] == df[col + "_norm_2"]

# Colonnes fuzzy
for col in fuzzy_columns:
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

# -----------------------------
# AJOUT RESULTATS
# -----------------------------
for col in strict_columns + fuzzy_columns:
    df[f"{col}_identique"] = comparisons[col]

for col in fuzzy_columns:
    df[f"{col}_score"] = scores[col]

# -----------------------------
# STATUT GENERAL
# -----------------------------
def ligne_statut(row):
    if row["_merge"] != "both":
        return "Manquant"
    cols_check = [f"{c}_identique" for c in strict_columns + fuzzy_columns]
    if all(row[col] for col in cols_check):
        return "OK"
    else:
        return "Différence"

df["statut"] = df.apply(ligne_statut, axis=1)

# -----------------------------
# EXPORT EXCEL
# -----------------------------
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df[df["statut"] == "OK"].to_excel(writer, sheet_name="Correspondances_OK", index=False)
    df[df["statut"] == "Différence"].to_excel(writer, sheet_name="Différences", index=False)
    df[df["statut"] == "Manquant"].to_excel(writer, sheet_name="Manquants", index=False)

print(f" Rapprochement terminé. Tu peux récuperer le résultat dans {output_file}")
