"""Contrôle Voyageurs vs GMAO - application Streamlit (fichier unique, 2 pages : Détection / Suivi détecté).

Installation : pip install -r requirements.txt
Lancement    : streamlit run app.py
"""
import html
import io
import re
import unicodedata

import altair as alt
import ftfy
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill
from rapidfuzz import fuzz
import streamlit as st


# =========================
# OUTILS COMMUNS
# =========================
def _fix_num(df):
    """NUM_CMDE_AN : remet les décimaux entiers (1.0) en entiers (1) pour que les clés correspondent."""
    df["NUM_CMDE_AN"] = pd.Series(
        [int(x) if isinstance(x, float) and x.is_integer() else x for x in df["NUM_CMDE_AN"]],
        index=df.index, dtype=object,
    )
    return df


def _roundtrip(df):
    """Écrit un DataFrame dans un « fichier Excel » en mémoire et renvoie ses octets."""
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


# =========================
# NORMALISATION (cellule 2 du notebook)
# =========================
corrections = {
    "POàŠLE": "POELE",
    "CRàŠPES": "CREPES",
    "à€": "A",
    "FRAàŠCHE": "FRAICHE",
    "JUSàŠ": "JUS",
}


def normalize_text(text):
    if pd.isna(text):
        return ""
    text = str(text)
    text = ftfy.fix_text(text)
    for wrong, correct in corrections.items():
        text = text.replace(wrong, correct)
    text = unicodedata.normalize('NFKD', text)
    text = ''.join(c for c in text if not unicodedata.combining(c))
    text = ''.join(c if c.isalnum() or c.isspace() else ' ' for c in text)
    text = ' '.join(text.upper().split())
    return text


# =========================
# NETTOYAGE (cellules 5 et 6 du notebook)
# =========================
def safe_col(df, col):
    return df[col] if col in df.columns else pd.NA


def clean_text(text):
    if pd.isna(text):
        return ""
    text = str(text)
    text = ftfy.fix_text(text)
    text = unicodedata.normalize('NFKD', text)
    text = text.encode('ascii', 'ignore').decode('utf-8')
    text = text.lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return text.strip()


def format_titre(text):
    if pd.isna(text):
        return text

    text = str(text).strip()

    match = re.match(r"^(\d{8})\s+", text)
    if match:
        return f"{match.group(1)} - {text[match.end():]}"

    return text


# =========================
# ÉTAPE 1 : préparation des deux sources (cellules 1, 2 et 3 du notebook)
# =========================
def preparer_sources(tb_bytes, cmd_bytes, log):
    # ---- Cellule 1 : consolidation des feuilles ----
    xls = pd.ExcelFile(io.BytesIO(tb_bytes))

    log("📄 Feuilles trouvées : " + " | ".join(xls.sheet_names))

    # Sélection des feuilles (2 à 9)
    sheets_to_use = xls.sheet_names[1:10]

    columns_to_keep = [
        "NUM_CMDE_AN",
        "TITRE_CMDE",
        "TITRE",
        "DES_ETAT_CMDE",
        "DATE_CMDE",
        "NOM",
        "QTE_CMDEE",
        "PRIXHT_LIGNE_CMDE",
        "DES_MAGASIN"
    ]

    dfs = []
    total_rows = 0

    for sheet in sheets_to_use:
        df = pd.read_excel(xls, sheet_name=sheet)

        # Nettoyage colonnes
        df.columns = df.columns.str.strip()

        # Garder colonnes utiles
        df = df.reindex(columns=columns_to_keep)

        # Supprimer lignes totalement vides
        df = df.dropna(how="all")

        nb_rows = len(df)
        total_rows += nb_rows
        log(f"✅ {sheet} → {nb_rows} lignes récupérées")

        df["SOURCE_FEUILLE"] = sheet
        dfs.append(df)

    final_df = pd.concat(dfs, ignore_index=True)
    log(f"📊 TOTAL GLOBAL : {total_rows} lignes")

    # ---- Cellule 2 : normalisation de TITRE ----
    df = pd.read_excel(io.BytesIO(_roundtrip(final_df)))
    df['TITRE'] = df['TITRE'].apply(normalize_text)
    voy_bytes = _roundtrip(df)

    # ---- Cellule 3 : fichier GMAO ----
    wb = load_workbook(io.BytesIO(cmd_bytes))
    ws = wb.active

    def fix_cell(cell):
        if cell.value is not None:
            if isinstance(cell.value, str) and (cell.value.startswith("=") or cell.value.startswith("=-@")):
                cell.value = "" + cell.value.lstrip("=-@").strip()

    for row in ws.iter_rows():
        for cell in row:
            fix_cell(cell)

    buf = io.BytesIO()
    wb.save(buf)
    gmao_bytes = buf.getvalue()

    return voy_bytes, gmao_bytes


def charger(voy_bytes, gmao_bytes):
    """Charge des copies neuves (l'audit modifie les DataFrames en place)."""
    gmao = pd.read_excel(io.BytesIO(gmao_bytes))
    gmao = _fix_num(gmao)
    voy = pd.read_excel(io.BytesIO(voy_bytes))
    voy = _fix_num(voy)
    return gmao, voy


# =========================
# ÉTAPE 2 : audit -> resultat.xlsx (cellule 5 du notebook)
# =========================
def audit(gmao, voy, log):
    # =========================
    # CLEAN DATA
    # =========================
    for df in [gmao, voy]:
        for col in ["TITRE", "NOM", "DES_ETAT_CMDE"]:
            if col in df.columns:
                df[col] = df[col].apply(clean_text)

        if "DATE_CMDE" in df.columns:
            df["DATE_CMDE"] = pd.to_datetime(df["DATE_CMDE"], errors="coerce").dt.date

        if "QTE_CMDEE" in df.columns:
            df["QTE_CMDEE"] = pd.to_numeric(df["QTE_CMDEE"], errors="coerce")

        if "PRIXHT_LIGNE_CMDE" in df.columns:
            df["PRIXHT_LIGNE_CMDE"] = pd.to_numeric(df["PRIXHT_LIGNE_CMDE"], errors="coerce")

    # =========================
    # KEY
    # =========================
    gmao["KEY"] = gmao["NUM_CMDE_AN"].astype(str) + "|" + gmao["TITRE"].astype(str) + "|" + gmao["DATE_CMDE"].astype(str)
    voy["KEY"] = voy["NUM_CMDE_AN"].astype(str) + "|" + voy["TITRE"].astype(str) + "|" + voy["DATE_CMDE"].astype(str)

    # =========================
    # NOUVELLES LIGNES
    # =========================
    def is_new(row):
        subset = voy[voy["NUM_CMDE_AN"] == row["NUM_CMDE_AN"]]
        if subset.empty:
            return True
        for t in subset["TITRE"]:
            if fuzz.ratio(str(row["TITRE"]), str(t)) >= 90:
                return False
        return True

    nouvelles_lignes = gmao[gmao.apply(is_new, axis=1)].copy()
    nouvelles_lignes["Type"] = "Nouvelle ligne"

    # =========================
    # MERGE
    # =========================
    merged = gmao.merge(voy, on="KEY", how="inner", suffixes=("_gmao", "_voy"))
    merged["Type"] = ""

    # =========================
    # SCORE MATCH
    # =========================
    def score_match(row):
        score = 0

        if row.get("NOM_gmao") == row.get("NOM_voy"):
            score += 1

        if abs((row.get("QTE_CMDEE_gmao") or 0) - (row.get("QTE_CMDEE_voy") or 0)) < 0.01:
            score += 1

        if abs((row.get("PRIXHT_LIGNE_CMDE_gmao") or 0) - (row.get("PRIXHT_LIGNE_CMDE_voy") or 0)) < 0.01:
            score += 1

        return score

    merged["match_score"] = merged.apply(score_match, axis=1)

    merged = (
        merged.sort_values("match_score", ascending=False)
        .drop_duplicates(subset=["KEY", "PRIXHT_LIGNE_CMDE_gmao"], keep="first")
    )

    # =========================
    # DETECTION CHANGEMENTS
    # =========================
    tol = 0.01

    merged.loc[safe_col(merged, "DES_ETAT_CMDE_gmao") != safe_col(merged, "DES_ETAT_CMDE_voy"), "Type"] += "Changement d'état | "
    merged.loc[safe_col(merged, "NOM_gmao") != safe_col(merged, "NOM_voy"), "Type"] += "Changement Fournisseur | "
    merged.loc[(safe_col(merged, "QTE_CMDEE_gmao").fillna(0) - safe_col(merged, "QTE_CMDEE_voy").fillna(0)).abs() > tol, "Type"] += "Changement QTE | "
    merged.loc[(safe_col(merged, "PRIXHT_LIGNE_CMDE_gmao").fillna(0) - safe_col(merged, "PRIXHT_LIGNE_CMDE_voy").fillna(0)).abs() > tol, "Type"] += "Changement PRIX | "

    merged = merged[merged["Type"] != ""].copy()

    merged["Type"] = merged["Type"].str.rstrip(" |")
    merged["Type"] = merged["Type"].str.replace(" | ", ", ", regex=False)

    # =========================
    # STRUCTURE FINALE
    # =========================
    colonnes_finales = [
        "NUM_CMDE_AN", "TITRE_CMDE", "TITRE", "DES_ETAT_CMDE", "DATE_CMDE",
        "NOM", "QTE_CMDEE", "PRIXHT_LIGNE_CMDE", "PRIX_TOTAL_HT", "TOTALHT",
        "TOTALTVA", "TYPE_LIGNE_CMDE", "DES_FAMILLE_FOURN", "DES_MAGASIN",
        "ETAT_gmao", "ETAT_voy", "NOM_gmao", "NOM_voy",
        "QTE_gmao", "QTE_voy", "PRIX_gmao", "PRIX_voy", "Type"
    ]

    merged["NUM_CMDE_AN"] = safe_col(merged, "NUM_CMDE_AN_gmao")
    merged["TITRE_CMDE"] = safe_col(merged, "TITRE_CMDE_gmao")
    merged["TITRE"] = safe_col(merged, "TITRE_gmao")
    merged["DES_ETAT_CMDE"] = safe_col(merged, "DES_ETAT_CMDE_gmao")
    merged["DATE_CMDE"] = safe_col(merged, "DATE_CMDE_gmao")
    merged["NOM"] = safe_col(merged, "NOM_gmao")
    merged["QTE_CMDEE"] = safe_col(merged, "QTE_CMDEE_gmao")
    merged["PRIXHT_LIGNE_CMDE"] = safe_col(merged, "PRIXHT_LIGNE_CMDE_gmao")
    merged["DES_MAGASIN"] = safe_col(merged, "DES_MAGASIN_gmao")

    merged["ETAT_gmao"] = safe_col(merged, "DES_ETAT_CMDE_gmao")
    merged["ETAT_voy"] = safe_col(merged, "DES_ETAT_CMDE_voy")
    merged["QTE_gmao"] = safe_col(merged, "QTE_CMDEE_gmao")
    merged["QTE_voy"] = safe_col(merged, "QTE_CMDEE_voy")
    merged["PRIX_gmao"] = safe_col(merged, "PRIXHT_LIGNE_CMDE_gmao")
    merged["PRIX_voy"] = safe_col(merged, "PRIXHT_LIGNE_CMDE_voy")

    for col in colonnes_finales:
        if col not in merged.columns:
            merged[col] = pd.NA

    resultat_changements = merged[colonnes_finales].copy()

    # =========================
    # NOUVELLES LIGNES
    # =========================
    def prepare_new(df):
        df2 = pd.DataFrame()

        for col in colonnes_finales:
            df2[col] = pd.NA

        df2["NUM_CMDE_AN"] = safe_col(df, "NUM_CMDE_AN")
        df2["TITRE_CMDE"] = safe_col(df, "TITRE_CMDE")
        df2["TITRE"] = safe_col(df, "TITRE")
        df2["DES_ETAT_CMDE"] = safe_col(df, "DES_ETAT_CMDE")
        df2["DATE_CMDE"] = safe_col(df, "DATE_CMDE")
        df2["NOM"] = safe_col(df, "NOM")
        df2["QTE_CMDEE"] = safe_col(df, "QTE_CMDEE")
        df2["PRIXHT_LIGNE_CMDE"] = safe_col(df, "PRIXHT_LIGNE_CMDE")
        df2["PRIX_TOTAL_HT"] = safe_col(df, "PRIX_TOTAL_HT")
        df2["TOTALHT"] = safe_col(df, "TOTALHT")
        df2["TOTALTVA"] = safe_col(df, "TOTALTVA")
        df2["TYPE_LIGNE_CMDE"] = safe_col(df, "TYPE_LIGNE_CMDE")
        df2["DES_FAMILLE_FOURN"] = safe_col(df, "DES_FAMILLE_FOURN")
        df2["DES_MAGASIN"] = safe_col(df, "DES_MAGASIN")

        df2["ETAT_gmao"] = safe_col(df, "DES_ETAT_CMDE")
        df2["NOM_gmao"] = safe_col(df, "NOM")
        df2["QTE_gmao"] = safe_col(df, "QTE_CMDEE")
        df2["PRIX_gmao"] = safe_col(df, "PRIXHT_LIGNE_CMDE")

        df2["Type"] = "Nouvelle ligne"

        return df2[colonnes_finales]

    nouvelles_final = prepare_new(nouvelles_lignes).copy()

    # =========================
    # FORMAT TITRE (FINAL)
    # =========================
    resultat_changements["TITRE"] = resultat_changements["TITRE"].apply(format_titre)
    nouvelles_final["TITRE"] = nouvelles_final["TITRE"].apply(format_titre)

    # =========================
    # EXPORT (en mémoire)
    # =========================
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        resultat_changements.to_excel(writer, sheet_name="Audit complet", index=False)
        nouvelles_final.to_excel(writer, sheet_name="Nouvelles lignes", index=False)
    buf.seek(0)

    # =========================
    # COLORATION
    # =========================
    wb = load_workbook(buf)

    def fill(c):
        return PatternFill(start_color=c, end_color=c, fill_type="solid")

    colors = {
        "Changement d'état": fill("FFC000"),
        "Changement Fournisseur": fill("00B0F0"),
        "Changement QTE": fill("FFAEC9"),
        "Changement PRIX": fill("FF99CC"),
        "Nouvelle ligne": fill("C6EFCE"),
    }

    for sheet in ["Audit complet", "Nouvelles lignes"]:
        ws = wb[sheet]
        col_type = [c.value for c in ws[1]].index("Type") + 1

        for row in ws.iter_rows(min_row=2):
            val = row[col_type - 1].value
            if val:
                for k in colors:
                    if k in val:
                        for cell in row:
                            cell.fill = colors[k]

    out = io.BytesIO()
    wb.save(out)
    log(f"✔ Audit terminé : {len(resultat_changements)} changement(s), {len(nouvelles_final)} nouvelle(s) ligne(s)")
    return out.getvalue(), len(resultat_changements), len(nouvelles_final)


# =========================
# ÉTAPE 3 : rapprochement -> rapprochement_titres_fusionne.xlsx (cellule 6 du notebook)
# =========================
def appliquer_fusion_et_style(src):
    wb = load_workbook(src)
    ws = wb.active

    start_row = 2
    max_row = ws.max_row

    # Fusionner les NUM_CMDE_AN identiques
    while start_row <= max_row:
        current_val = ws.cell(row=start_row, column=1).value  # Col 1: NUM_CMDE_AN
        end_row = start_row

        while (end_row + 1 <= max_row and
               ws.cell(row=end_row + 1, column=1).value == current_val):
            end_row += 1

        if end_row > start_row and current_val is not None:
            ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
            ws.cell(row=start_row, column=1).alignment = Alignment(horizontal="center", vertical="center")

        start_row = end_row + 1

    # Ajustement largeur colonnes
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except Exception:
                pass
        ws.column_dimensions[column].width = min(max_length + 2, 50)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def rapprochement(gmao, voy, log):
    # Nettoyage rapide pour la comparaison
    gmao["TITRE_CLEAN"] = gmao["TITRE"].apply(clean_text)
    voy["TITRE_CLEAN"] = voy["TITRE"].apply(clean_text)

    # =========================
    # IDENTIFICATION NOUVELLES LIGNES
    # =========================
    def is_new(row, reference_df):
        subset = reference_df[reference_df["NUM_CMDE_AN"] == row["NUM_CMDE_AN"]]
        if subset.empty:
            return True
        for t in subset["TITRE_CLEAN"]:
            if fuzz.ratio(row["TITRE_CLEAN"], t) >= 90:
                return False
        return True

    nouvelles_lignes = gmao[gmao.apply(is_new, axis=1, reference_df=voy)].copy()

    # =========================
    # CONSTRUCTION DU COMPARATIF
    # =========================
    nums_concernes = nouvelles_lignes["NUM_CMDE_AN"].unique()

    df_gmao_view = nouvelles_lignes[["NUM_CMDE_AN", "TITRE_CMDE", "TITRE", "DES_ETAT_CMDE"]].copy()
    df_gmao_view.columns = ["NUM_CMDE_AN", "TITRE_CMDE_GMAO", "TITRE_GMAO", "DES_ETAT_CMDE_GMAO"]

    df_voy_view = voy[voy["NUM_CMDE_AN"].isin(nums_concernes)][["NUM_CMDE_AN", "TITRE"]].copy()
    df_voy_view.columns = ["NUM_CMDE_AN", "TITRE_CONSO_ROBUSTE"]

    comparatif_titres = pd.merge(df_gmao_view, df_voy_view, on="NUM_CMDE_AN", how="outer")
    comparatif_titres = comparatif_titres.sort_values("NUM_CMDE_AN")

    # =========================
    # EXPORT ET FINITIONS (en mémoire)
    # =========================
    buf = io.BytesIO()
    comparatif_titres.to_excel(buf, index=False)
    buf.seek(0)

    out = appliquer_fusion_et_style(buf)
    log(f"✔ Rapprochement terminé : {len(comparatif_titres)} ligne(s)")
    return out, len(comparatif_titres)


# =========================
# SYNTHÈSE DU TB PAR FEUILLE (lue directement dans le fichier TB de correspondance)
# =========================
def synthese_tb(tb_bytes):
    """Pour chaque feuille utilisée du TB (les mêmes que la consolidation : 2e à 10e) :
    nombre de NUM_CMDE_AN distincts et somme de PRIX_TOTAL_HT. Ne touche pas au traitement principal."""
    xls = pd.ExcelFile(io.BytesIO(tb_bytes))
    lignes = []
    for feuille in xls.sheet_names[1:10]:
        df = pd.read_excel(xls, sheet_name=feuille)
        df.columns = df.columns.astype(str).str.strip()
        df = df.reindex(columns=["NUM_CMDE_AN", "PRIX_TOTAL_HT"]).dropna(how="all")
        df = _fix_num(df)   # 1.0 et 1 comptent pour une seule commande
        montant = pd.to_numeric(df["PRIX_TOTAL_HT"], errors="coerce")
        lignes.append({
            "FEUILLE": feuille,
            "NB_COMMANDES": int(df["NUM_CMDE_AN"].nunique()),
            "TOTAL_HT": float(montant.sum()),
        })
    return pd.DataFrame(lignes, columns=["FEUILLE", "NB_COMMANDES", "TOTAL_HT"])


# =========================
# ANALYSE DES NOUVELLES LIGNES (section « Analyse » de la page Détection)
# =========================
def analyser_nouvelles(resultat_bytes, voy):
    """Analyse la feuille « Nouvelles lignes » de resultat.xlsx. N'altère aucun calcul de l'audit.

    - lignes        : les nouvelles lignes, avec la feuille du TB rattachée à leur commande (colonne FEUILLE) ;
    - total_ht      : somme de PRIX_TOTAL_HT sur toutes les nouvelles lignes ;
    - nb_commandes  : nombre de NUM_CMDE_AN distincts ;
    - par_feuille   : par feuille, nombre de commandes distinctes, nombre de lignes et somme de PRIX_TOTAL_HT.

    La GMAO n'a pas de notion de feuille : la feuille d'une nouvelle ligne est celle où le TB contient déjà ce
    NUM_CMDE_AN (cas d'un article ajouté à une commande connue). Une commande totalement nouvelle n'est dans
    aucune feuille du TB : elle n'a pas de feuille (FEUILLE vide) et n'entre pas dans le regroupement par feuille.
    """
    lignes = pd.read_excel(io.BytesIO(resultat_bytes), sheet_name="Nouvelles lignes")
    lignes = _fix_num(lignes)
    lignes["PRIX_TOTAL_HT"] = pd.to_numeric(lignes["PRIX_TOTAL_HT"], errors="coerce")

    feuille_par_cmde = voy.drop_duplicates("NUM_CMDE_AN").set_index("NUM_CMDE_AN")["SOURCE_FEUILLE"]
    lignes.insert(1, "FEUILLE", lignes["NUM_CMDE_AN"].map(feuille_par_cmde))   # vide si commande absente du TB

    par_feuille = (
        lignes.groupby("FEUILLE")
        .agg(NB_COMMANDES=("NUM_CMDE_AN", "nunique"), NB_LIGNES=("NUM_CMDE_AN", "size"),
             TOTAL_HT=("PRIX_TOTAL_HT", "sum"))
        .reset_index()
        .sort_values("TOTAL_HT", ascending=False)
        .reset_index(drop=True)
    )
    audit_df = pd.read_excel(io.BytesIO(resultat_bytes), sheet_name="Audit complet")
    types = audit_df["Type"].fillna("").astype(str) if "Type" in audit_df.columns else pd.Series([], dtype=str)
    changements_par_type = pd.DataFrame({
        "TYPE": ["État", "Fournisseur", "Quantité", "Prix"],
        "NB": [int(types.str.contains(t, regex=False).sum())
               for t in ("Changement d'état", "Changement Fournisseur", "Changement QTE", "Changement PRIX")],
    })
    return {
        "changements_par_type": changements_par_type,
        "lignes": lignes,
        "total_ht": float(lignes["PRIX_TOTAL_HT"].sum()),
        "nb_commandes": int(lignes["NUM_CMDE_AN"].nunique()),
        "nb_sans_montant": int(lignes["PRIX_TOTAL_HT"].isna().sum()),
        "nb_sans_feuille": int(lignes["FEUILLE"].isna().sum()),
        "montant_sans_feuille": float(lignes.loc[lignes["FEUILLE"].isna(), "PRIX_TOTAL_HT"].sum()),
        "par_feuille": par_feuille,
    }


# =========================
# POINT D'ENTRÉE
# =========================
def run(tb_bytes, cmd_bytes, log=print):
    voy_bytes, gmao_bytes = preparer_sources(tb_bytes, cmd_bytes, log)

    gmao, voy = charger(voy_bytes, gmao_bytes)
    resultat, nb_changements, nb_nouvelles = audit(gmao, voy, log)

    # L'audit modifie gmao et voy en place : on repart de copies neuves
    gmao, voy = charger(voy_bytes, gmao_bytes)
    rapproch, nb_rapproch = rapprochement(gmao, voy, log)
    analyse = analyser_nouvelles(resultat, voy)
    analyse["synthese_tb"] = synthese_tb(tb_bytes)

    return {
        "analyse": analyse,
        "resultat": resultat,
        "rapprochement": rapproch,
        "nb_changements": nb_changements,
        "nb_nouvelles": nb_nouvelles,
        "nb_rapprochement": nb_rapproch,
    }


# ==========================================================
# SUIVI DÉTECTÉ : TB de correspondance vs suivi budgétaire
# Moteur commun aux scripts Communication / SI / Maintenance.
# La logique reprend celle des notebooks ; seuls les réglages
# propres à chaque service sont dans PROFILS_SUIVI.
# ==========================================================
from dataclasses import dataclass
from datetime import datetime


class ErreurSuivi(Exception):
    """Erreur dont le message est destiné à l'utilisateur (il ne contient jamais de données des fichiers)."""


class Journal:
    """Collecte les messages d'un traitement (un Journal par traitement : rien n'est partagé entre sessions)."""

    def __init__(self):
        self.messages = []

    def _ajouter(self, niveau, msg, args):
        self.messages.append((niveau, (msg % args) if args else msg))

    def info(self, msg, *args):
        self._ajouter("info", msg, args)

    def warning(self, msg, *args):
        self._ajouter("warning", msg, args)


COLONNES_TB = [
    "NUM_CMDE_AN", "TITRE_CMDE", "TITRE", "DES_ETAT_CMDE", "DATE_CMDE",
    "NOM", "QTE_CMDEE", "PRIXHT_LIGNE_CMDE", "PRIX_TOTAL_HT", "DES_MAGASIN",
]


@dataclass(frozen=True)
class ProfilSuivi:
    """Réglages propres à un service."""
    prefixes_feuilles: tuple      # feuilles du TB utilisées (le nom commence par...)
    exclure_annulees: bool        # écarter les commandes à l'état « annulée »
    cle_code_article: bool        # True : clé = n° commande + code article ; False : n° commande + libellé
    tri: tuple                    # tri du TB consolidé (détermine quelle ligne est gardée en cas de doublon)
    format_export: bool           # arrondi propre + virgule décimale FR dans le fichier de résultat
    description: str


PROFILS_SUIVI = {
    "Communication": ProfilSuivi(
        prefixes_feuilles=("marketing",),
        exclure_annulees=True,
        cle_code_article=False,
        tri=("DATE_CMDE", "NUM_CMDE_AN"),
        format_export=True,
        description=(
            "- Feuille du TB : celle dont le nom commence par **marketing**\n"
            "- Commandes **annulées exclues**\n"
            "- Nouvelles lignes : numéro de commande absent du suivi\n"
            "- Changements : lignes appariées par **numéro de commande + libellé**"
        ),
    ),
    "SI": ProfilSuivi(
        prefixes_feuilles=("si",),
        exclure_annulees=False,
        cle_code_article=False,
        tri=("SOURCE_FEUILLE", "DATE_CMDE", "NUM_CMDE_AN"),
        format_export=False,
        description=(
            "- Feuille du TB : celle dont le nom commence par **SI**\n"
            "- Commandes annulées conservées\n"
            "- Nouvelles lignes : numéro de commande absent du suivi\n"
            "- Changements : lignes appariées par **numéro de commande + libellé**\n"
            "- Les nombres sont exportés tels quels (version SI du script)"
        ),
    ),
    "Maintenance": ProfilSuivi(
        prefixes_feuilles=("maint interne", "maint externe"),
        exclure_annulees=False,
        cle_code_article=True,
        tri=("SOURCE_FEUILLE", "DATE_CMDE", "NUM_CMDE_AN"),
        format_export=True,
        description=(
            "- Feuilles du TB : **maint interne** et **maint externe**\n"
            "- Commandes annulées conservées\n"
            "- Nouvelles lignes et changements : lignes appariées par "
            "**numéro de commande + code article** (nombre de 6 à 10 chiffres du libellé)"
        ),
    ),
}


@dataclass(frozen=True)
class ConfigSuivi:
    feuille_bdd: str = "BDD 2026"
    feuille_tb: str = "Sheet1"
    col_bdd: str = "Numéro de bon de compte"
    col_tb: str = "NUM_CMDE_AN"
    col_detail_bdd: str = "Détail"
    col_titre_tb: str = "TITRE"
    tolerance_numerique: float = 1e-6
    # (colonne TB, colonne BDD, type)
    champs_a_comparer: tuple = (
        ("DES_ETAT_CMDE", "Etat", "texte"),
        ("QTE_CMDEE", "Quantité", "numerique"),
        ("PRIX_TOTAL_HT", "MONTANT BC", "numerique"),
    )
    colonnes_entieres: tuple = ("Quantité", "QTE_CMDEE")
    decimales_montant: int = 2
    colonnes_numeriques_tb: tuple = ("QTE_CMDEE", "PRIXHT_LIGNE_CMDE", "PRIX_TOTAL_HT")
    annee: int = 2026


RE_ESPACES = re.compile(r"\s+", flags=re.UNICODE)
RE_SUFFIXE_FLOAT = re.compile(r"\.0+$")
RE_NON_ALPHANUM = re.compile(r"[^A-Z0-9]", flags=re.UNICODE)
RE_CODE_ARTICLE = re.compile(r"(\d{6,10})")


# ---------- Nettoyage ----------
def _retirer_accents(v):
    if not v:
        return v
    nfd = unicodedata.normalize("NFD", v)
    return "".join(c for c in nfd if not unicodedata.combining(c))


def nettoyer_cles(serie):
    s = serie.astype("string")
    s = s.fillna("")
    s = s.map(_retirer_accents)
    s = s.str.replace(RE_ESPACES, "", regex=True)
    s = s.str.replace(RE_SUFFIXE_FLOAT, "", regex=True)
    s = s.str.upper()
    return s


def nettoyer_libelle(serie):
    s = nettoyer_cles(serie)
    s = s.str.replace(RE_NON_ALPHANUM, "", regex=True)
    return s


def extraire_code_article(serie):
    s = serie.astype("string").fillna("")
    codes = s.str.extract(RE_CODE_ARTICLE, expand=False)
    codes = codes.fillna("")
    return codes


def convertir_numerique(serie, log=None):
    s = serie.astype("string").fillna("")
    s = s.str.replace(RE_ESPACES, "", regex=True)
    s = s.str.replace(",", ".", regex=False)
    s_num = pd.to_numeric(s, errors="coerce")

    non_convertibles = s_num.isna() & (s.str.len() > 0)
    if non_convertibles.any() and log is not None:
        log.warning(
            "%d valeur(s) numérique(s) illisibles ignorées (ex : %s).",
            non_convertibles.sum(),
            s[non_convertibles].iloc[0],
        )
    return s_num


def formater_numerique_export(serie, entier=False, decimales=2, log=None):
    v = convertir_numerique(serie, log)
    if entier:
        return v.round(0).astype("Int64")
    return v.round(decimales)


def formater_numerique_fr(serie, entier=False, decimales=2, log=None):
    v = formater_numerique_export(serie, entier=entier, decimales=decimales, log=log)
    if entier:
        return v.astype("string").fillna("")

    def _fmt(x):
        if pd.isna(x):
            return ""
        return f"{x:.{decimales}f}".replace(".", ",")

    return v.map(_fmt).astype("string")


def _normaliser_etat(text):
    if pd.isna(text):
        return ""
    text = str(text).strip().lower()
    text = unicodedata.normalize("NFD", text)
    return "".join(c for c in text if unicodedata.category(c) != "Mn")


# ---------- Étape 1 : consolidation du TB de correspondance ----------
def consolider_tb(tb_bytes, profil, cfg, log):
    """Équivalent de la cellule « TB consolidation » des notebooks. Renvoie le TB consolidé (octets .xlsx)."""
    try:
        xls = pd.ExcelFile(io.BytesIO(tb_bytes))
    except Exception as exc:
        raise ErreurSuivi("Le TB de correspondance n'a pas pu être lu (est-ce bien un fichier Excel valide ?).") from exc

    sheets_to_use = [s for s in xls.sheet_names if s.strip().lower().startswith(profil.prefixes_feuilles)]
    if not sheets_to_use:
        raise ErreurSuivi(
            f"Aucune feuille du TB ne commence par {' / '.join(profil.prefixes_feuilles)}. "
            f"Feuilles trouvées : {xls.sheet_names}"
        )
    log.info("Feuilles du TB utilisées : %s", ", ".join(sheets_to_use))

    dfs = []
    total_before = 0
    total_after = 0

    for sheet in sheets_to_use:
        try:
            df = pd.read_excel(xls, sheet_name=sheet)
            df.columns = df.columns.str.strip()

            missing = set(COLONNES_TB) - set(df.columns)
            if missing:
                log.warning("Feuille '%s' : colonnes manquantes : %s", sheet, ", ".join(sorted(missing)))

            df = df.reindex(columns=COLONNES_TB)
            df = df.replace(r"^\s*$", pd.NA, regex=True)
            df = df.dropna(how="all")

            lignes_avant = len(df)

            df["DATE_CMDE"] = pd.to_datetime(df["DATE_CMDE"], errors="coerce", dayfirst=True)
            df = df[df["DATE_CMDE"].dt.year == cfg.annee]

            if profil.exclure_annulees:
                df = df[df["DES_ETAT_CMDE"].apply(_normaliser_etat) != "annulee"]

            lignes_apres = len(df)
            total_before += lignes_avant
            total_after += lignes_apres

            df["SOURCE_FEUILLE"] = sheet
            dfs.append(df)
            log.info("Feuille '%s' : %d ligne(s) avant filtres, %d conservée(s)", sheet, lignes_avant, lignes_apres)
        except Exception as exc:
            log.warning("Erreur sur la feuille '%s' : %s", sheet, exc)

    if dfs:
        final_df = pd.concat(dfs, ignore_index=True)
        final_df = final_df.sort_values(by=list(profil.tri), ignore_index=True)
    else:
        final_df = pd.DataFrame(columns=COLONNES_TB + ["SOURCE_FEUILLE"])

    log.info("TB consolidé : %d ligne(s) sur %d (année %d)", len(final_df), total_before, cfg.annee)

    buf = io.BytesIO()
    final_df.to_excel(buf, index=False)   # feuille « Sheet1 », comme dans les notebooks
    return buf.getvalue()


# ---------- Chargement ----------
def charger_feuille(contenu, feuille, colonne_cle, libelle, log):
    try:
        df = pd.read_excel(io.BytesIO(contenu), sheet_name=feuille, engine="openpyxl", dtype=str)
    except ValueError as exc:
        try:
            feuilles_dispo = pd.ExcelFile(io.BytesIO(contenu), engine="openpyxl").sheet_names
        except Exception:
            feuilles_dispo = ["<indisponible>"]
        raise ErreurSuivi(
            f"Impossible de lire la feuille '{feuille}' du fichier {libelle}. "
            f"Feuilles disponibles : {feuilles_dispo}"
        ) from exc
    except Exception as exc:
        raise ErreurSuivi(f"Le fichier {libelle} n'a pas pu être lu (est-ce bien un fichier Excel valide ?).") from exc

    colonnes_brutes = list(df.columns)
    df.columns = df.columns.astype(str).str.strip()
    renommees = [(a, b) for a, b in zip(colonnes_brutes, df.columns) if a != b]
    if renommees:
        log.warning("En-têtes avec espaces parasites corrigés dans %s : %s", libelle, renommees)

    if colonne_cle not in df.columns:
        raise ErreurSuivi(
            f"Colonne '{colonne_cle}' absente du fichier {libelle} (feuille '{feuille}'). "
            f"Colonnes disponibles : {list(df.columns)}"
        )
    return df


# ---------- Logique métier ----------
def detecter_nouvelles_lignes(bdd, tb, cfg, profil, log):
    """Lignes du TB dont la clé n'existe pas dans le suivi.

    Clé = numéro de commande (Communication, SI) ou numéro + code article (Maintenance).
    """
    bdd = bdd.copy()
    tb = tb.copy()

    col_bdd, col_tb = cfg.col_bdd, cfg.col_tb
    bdd[col_bdd] = nettoyer_cles(bdd[col_bdd])
    tb[col_tb] = nettoyer_cles(tb[col_tb])

    if profil.cle_code_article:
        a_code_article = cfg.col_detail_bdd in bdd.columns and cfg.col_titre_tb in tb.columns
        if a_code_article:
            bdd["Code_article"] = extraire_code_article(bdd[cfg.col_detail_bdd])
            tb["Code_article"] = extraire_code_article(tb[cfg.col_titre_tb])
            bdd["_CLE"] = bdd[col_bdd] + "||" + bdd["Code_article"]
            tb["_CLE"] = tb[col_tb] + "||" + tb["Code_article"]
        else:
            log.warning(
                "Colonne '%s' (suivi) ou '%s' (TB) absente : les nouvelles lignes se basent "
                "uniquement sur le numéro de commande, sans code article.",
                cfg.col_detail_bdd, cfg.col_titre_tb,
            )
            bdd["_CLE"] = bdd[col_bdd]
            tb["_CLE"] = tb[col_tb]
    else:
        bdd["_CLE"] = bdd[col_bdd]
        tb["_CLE"] = tb[col_tb]

    vides_tb = (tb[col_tb] == "").sum()
    if vides_tb:
        log.warning(
            "%d ligne(s) du TB ont une clé vide après nettoyage et seront ignorées "
            "(elles ne peuvent pas être comparées).",
            vides_tb,
        )

    bdd = bdd[bdd[col_bdd] != ""]
    tb_valide = tb[tb[col_tb] != ""]

    doublons_bdd = bdd["_CLE"].duplicated().sum()
    if doublons_bdd:
        log.warning("%d clé(s) en doublon détectée(s) dans le suivi budgétaire.", doublons_bdd)

    doublons_tb = tb_valide["_CLE"].duplicated().sum()
    if doublons_tb:
        log.warning("%d clé(s) en doublon détectée(s) dans le TB.", doublons_tb)

    cles_bdd = set(bdd["_CLE"])
    nouvelles = tb_valide[~tb_valide["_CLE"].isin(cles_bdd)].copy()

    nouvelles["Nb_occurrences_cle"] = nouvelles.groupby(col_tb)[col_tb].transform("size")
    nouvelles = nouvelles.drop(columns=["_CLE"], errors="ignore")
    return nouvelles


def detecter_changements(bdd, tb, cfg, profil, log):
    """Changements (état, quantité, montant) sur les lignes présentes dans les deux fichiers."""
    colonnes_vides = [cfg.col_tb, cfg.col_titre_tb, "Champ", "Valeur_BDD", "Valeur_TB"]

    colonnes_bdd_requises = [cfg.col_bdd, cfg.col_detail_bdd] + [b for _, b, _ in cfg.champs_a_comparer]
    colonnes_tb_requises = [cfg.col_tb, cfg.col_titre_tb] + [t for t, _, _ in cfg.champs_a_comparer]

    manquantes_bdd = [c for c in colonnes_bdd_requises if c not in bdd.columns]
    manquantes_tb = [c for c in colonnes_tb_requises if c not in tb.columns]
    if manquantes_bdd or manquantes_tb:
        log.warning(
            "Détection de changements ignorée (colonnes manquantes) - suivi : %s | TB : %s",
            manquantes_bdd, manquantes_tb,
        )
        return pd.DataFrame(columns=colonnes_vides)

    bdd2 = bdd[colonnes_bdd_requises].copy()
    tb2 = tb[colonnes_tb_requises].copy()

    num_bdd = nettoyer_cles(bdd2[cfg.col_bdd])
    num_tb = nettoyer_cles(tb2[cfg.col_tb])

    if profil.cle_code_article:
        code_bdd = extraire_code_article(bdd2[cfg.col_detail_bdd])
        bdd2["_CODE_ARTICLE"] = code_bdd
        bdd2["_CLE"] = num_bdd + "||" + code_bdd
        code_tb = extraire_code_article(tb2[cfg.col_titre_tb])
        tb2["_CODE_ARTICLE"] = code_tb
        tb2["_CLE"] = num_tb + "||" + code_tb

        sans_code_bdd = (code_bdd == "").sum()
        if sans_code_bdd:
            log.warning("%d ligne(s) du suivi n'ont pas de code article détectable dans '%s'.",
                        sans_code_bdd, cfg.col_detail_bdd)
        sans_code_tb = (code_tb == "").sum()
        if sans_code_tb:
            log.warning("%d ligne(s) du TB n'ont pas de code article détectable dans '%s'.",
                        sans_code_tb, cfg.col_titre_tb)
        nom_cle = "Numéro+Code article"
    else:
        det_bdd = nettoyer_libelle(bdd2[cfg.col_detail_bdd])
        bdd2["_CLE"] = num_bdd + "||" + det_bdd
        det_tb = nettoyer_libelle(tb2[cfg.col_titre_tb])
        tb2["_CLE"] = num_tb + "||" + det_tb
        nom_cle = "Numéro+Libellé"

    bdd2 = bdd2[num_bdd != ""]
    tb2 = tb2[num_tb != ""]

    dup_bdd = bdd2["_CLE"].duplicated()
    if dup_bdd.any():
        log.warning(
            "%d clé(s) composite(s) (%s) en doublon dans le suivi : "
            "seule la première occurrence de chacune est comparée.", dup_bdd.sum(), nom_cle,
        )
        bdd2 = bdd2.drop_duplicates(subset="_CLE", keep="first")

    dup_tb = tb2["_CLE"].duplicated()
    if dup_tb.any():
        log.warning(
            "%d clé(s) composite(s) (%s) en doublon dans le TB : "
            "seule la première occurrence de chacune est comparée.", dup_tb.sum(), nom_cle,
        )
        tb2 = tb2.drop_duplicates(subset="_CLE", keep="first")

    fusion = tb2.merge(bdd2, on="_CLE", how="inner")
    log.info("%d ligne(s) appariée(s) entre le TB et le suivi pour la détection de changements.", len(fusion))

    resultats = []
    for col_tb_champ, col_bdd_champ, type_champ in cfg.champs_a_comparer:
        if type_champ == "numerique":
            v_tb = convertir_numerique(fusion[col_tb_champ], log).fillna(0.0)
            v_bdd = convertir_numerique(fusion[col_bdd_champ], log).fillna(0.0)
            diff = (v_tb - v_bdd).abs() > cfg.tolerance_numerique
        else:
            t_tb = nettoyer_cles(fusion[col_tb_champ])
            t_bdd = nettoyer_cles(fusion[col_bdd_champ])
            diff = t_tb != t_bdd

        if not diff.any():
            continue

        sous = fusion.loc[diff, [cfg.col_tb, cfg.col_titre_tb]].copy()
        sous.insert(2, "Champ", col_bdd_champ)

        if type_champ == "numerique" and profil.format_export:
            entier = col_bdd_champ in cfg.colonnes_entieres
            sous["Valeur_BDD"] = formater_numerique_export(
                fusion.loc[diff, col_bdd_champ], entier=entier, decimales=cfg.decimales_montant, log=log)
            sous["Valeur_TB"] = formater_numerique_export(
                fusion.loc[diff, col_tb_champ], entier=entier, decimales=cfg.decimales_montant, log=log)
        else:
            sous["Valeur_BDD"] = fusion.loc[diff, col_bdd_champ]
            sous["Valeur_TB"] = fusion.loc[diff, col_tb_champ]

        resultats.append(sous)

    if not resultats:
        return pd.DataFrame(columns=colonnes_vides)

    changements = pd.concat(resultats, ignore_index=True)
    changements = changements.sort_values([cfg.col_tb, "Champ"]).reset_index(drop=True)
    return changements


def preparer_export_nouvelles(nouvelles, cfg, log):
    nouvelles = nouvelles.copy()
    for col in cfg.colonnes_numeriques_tb:
        if col not in nouvelles.columns:
            continue
        entier = col in cfg.colonnes_entieres
        nouvelles[col] = formater_numerique_fr(nouvelles[col], entier=entier, decimales=cfg.decimales_montant, log=log)
    return nouvelles


def preparer_export_changements(changements):
    if changements.empty:
        return changements
    changements = changements.copy()
    for col in ("Valeur_BDD", "Valeur_TB"):
        est_numerique = pd.api.types.is_numeric_dtype(changements[col]) or pd.api.types.is_extension_array_dtype(
            changements[col]
        )
        if est_numerique:
            s = changements[col].astype("string").fillna("")
            changements[col] = s.str.replace(".", ",", regex=False)
    return changements


def calculer_kpi_suivi(nouvelles, changements, cfg, nb_lignes_bdd):
    """Indicateurs affichés sur la page Suivi, calculés sur les résultats bruts (avant mise en forme de l'export).

    Un écart = valeur du TB - valeur du suivi (positif : le TB est plus élevé que le suivi).
    """
    champ_bdd = {tb_col: bdd_col for tb_col, bdd_col, _ in cfg.champs_a_comparer}
    champ_montant, champ_qte = champ_bdd.get("PRIX_TOTAL_HT"), champ_bdd.get("QTE_CMDEE")

    montant_nouvelles = 0.0
    if len(nouvelles) and "PRIX_TOTAL_HT" in nouvelles.columns:
        montant_nouvelles = float(convertir_numerique(nouvelles["PRIX_TOTAL_HT"]).sum())

    kpi = {"montant_nouvelles": montant_nouvelles, "ecart_montant": 0.0, "ecart_qte": 0.0,
           "commandes_touchees": 0, "lignes_en_changement": 0, "taux_changement": 0.0, "plus_gros_ecart": None}
    if not len(changements):
        return kpi

    ecart = (convertir_numerique(changements["Valeur_TB"]).fillna(0.0)
             - convertir_numerique(changements["Valeur_BDD"]).fillna(0.0))
    est_montant = changements["Champ"] == champ_montant
    est_qte = changements["Champ"] == champ_qte
    kpi["ecart_montant"] = float(ecart[est_montant].sum())
    kpi["ecart_qte"] = float(ecart[est_qte].sum())
    kpi["commandes_touchees"] = int(changements[cfg.col_tb].nunique())
    kpi["lignes_en_changement"] = int(len(changements[[cfg.col_tb, cfg.col_titre_tb]].drop_duplicates()))
    if nb_lignes_bdd:
        kpi["taux_changement"] = kpi["lignes_en_changement"] / nb_lignes_bdd * 100
    if est_montant.any():
        i = ecart[est_montant].abs().idxmax()
        kpi["plus_gros_ecart"] = {"valeur": float(ecart[i]), "commande": str(changements.loc[i, cfg.col_tb])}
    return kpi


def executer_suivi(tb_bytes, bdd_bytes, profil, annee=2026, feuille_bdd="BDD 2026"):
    """Traitement complet. Renvoie le fichier de résultat (octets), les statistiques et le journal."""
    log = Journal()
    cfg = ConfigSuivi(feuille_bdd=feuille_bdd, annee=int(annee))

    tb_consolide = consolider_tb(tb_bytes, profil, cfg, log)
    bdd = charger_feuille(bdd_bytes, cfg.feuille_bdd, cfg.col_bdd, "de suivi budgétaire", log)
    tb = charger_feuille(tb_consolide, cfg.feuille_tb, cfg.col_tb, "TB consolidé", log)

    nouvelles_tb = detecter_nouvelles_lignes(bdd, tb, cfg, profil, log)
    nb_commandes_distinctes = nouvelles_tb[cfg.col_tb].nunique()

    changements = detecter_changements(bdd, tb, cfg, profil, log)

    if profil.format_export:
        nouvelles_export = preparer_export_nouvelles(nouvelles_tb, cfg, log)
        changements_export = preparer_export_changements(changements)
    else:
        nouvelles_export, changements_export = nouvelles_tb, changements

    sortie = io.BytesIO()
    with pd.ExcelWriter(sortie, engine="openpyxl") as writer:
        nouvelles_export.to_excel(writer, index=False, sheet_name="Nouvelles")
        changements_export.to_excel(writer, index=False, sheet_name="Changements")

    repartition = changements["Champ"].value_counts().to_dict() if len(changements) else {}
    stats = {
        "lignes_bdd": len(bdd),
        "lignes_tb": len(tb),
        "nouvelles": len(nouvelles_tb),
        "commandes_distinctes": int(nb_commandes_distinctes),
        "changements": len(changements),
        "repartition": repartition,
        "kpi": calculer_kpi_suivi(nouvelles_tb, changements, cfg, len(bdd)),
    }
    return {
        "xlsx": sortie.getvalue(),
        "stats": stats,
        "journal": log.messages,
        "nouvelles": nouvelles_export,
        "changements": changements_export,
    }


# ==========================================================
# SÉCURITÉ COMMUNE ET INTERFACE STREAMLIT (2 pages)
# ==========================================================
import hmac
import logging
import time


st.set_page_config(page_title="Contrôle Voyageurs vs GMAO", page_icon="📊", layout="wide")

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
ESPACES_DONNEES = ("det", "suivi")  # une zone de résultats par page


def delai_inactivite_minutes():
    """Délai avant effacement automatique (secret INACTIVITE_MINUTES, 5 minutes par défaut)."""
    try:
        v = float(st.secrets.get("INACTIVITE_MINUTES", 5))
    except Exception:
        v = 5.0
    return max(1.0, v)


def effacer_donnees():
    """Supprime résultats, fichiers déposés et mots de passe de sortie de la session (toutes les pages)."""
    for espace in ESPACES_DONNEES:
        for k in (f"res_{espace}", f"enc_{espace}", f"enc_cle_{espace}"):
            st.session_state.pop(k, None)
    st.session_state["reset_n"] = st.session_state.get("reset_n", 0) + 1  # vide les zones de dépôt


CSS_CONNEXION = """
<style>
[data-testid="stApp"] { background: #faf9f6; }
[data-testid="stHeader"] { background: transparent; }
[data-testid="stMainBlockContainer"] { max-width: 700px; padding-top: 3.5rem; }
.cx-entete { text-align: center; margin-bottom: 1.6rem; }
.cx-icone { width: 90px; height: 90px; margin: 0 auto 1.2rem; border-radius: 24px; font-size: 2.6rem;
  display: flex; align-items: center; justify-content: center; color: #fff;
  background: linear-gradient(135deg, #4aa8f0, #1a7fd6); box-shadow: 0 12px 32px rgba(26,140,220,.30); }
.cx-entete h1 { font-family: Georgia, "Palatino Linotype", serif; font-weight: 400; font-size: 2.4rem; white-space: nowrap;
  color: #1a1a1a; margin: 0; padding: 0; }
.cx-entete p { color: #6b6b6b; font-size: 1.05rem; margin: .6rem 0 0; }
[data-testid="stForm"] { background: #fff; border: 1px solid #ece9e2; border-radius: 24px;
  padding: 2rem 2.2rem; box-shadow: 0 10px 40px rgba(0,0,0,.06); }
[data-testid="stForm"] label p { color: #333; font-weight: 500; }
[data-testid="stForm"] [data-testid="stTextInputRootElement"] { background: #fff !important; border: 1px solid #d9d6cc !important; border-radius: 12px; }
[data-testid="stForm"] [data-testid="stElementContainer"]:has([data-testid="stFormSubmitButton"]) { width: 100% !important; }
[data-testid="stFormSubmitButton"], [data-testid="stFormSubmitButton"] > div { width: 100% !important; display: block; }
[data-testid="stFormSubmitButton"] button { width: 100% !important; height: 3.2rem; border: 0; border-radius: 14px;
  background: #1a8cdc; color: #fff; font-weight: 500; }
[data-testid="stFormSubmitButton"] button:hover { background: #1678bd; color: #fff; }
.cx-pied { text-align: center; color: #9a9a90; font-size: .9rem; margin-top: 1.6rem; }
</style>
"""

CSS_SESSION = """
<style>
.cx-carte { display: flex; align-items: center; gap: 12px; background: #fff; border: 1px solid #e6e9ef;
  border-radius: 18px; padding: 14px; box-shadow: 0 2px 8px rgba(0,0,0,.04); margin-bottom: .6rem; }
.cx-avatar { width: 52px; height: 52px; flex: none; border-radius: 50%; color: #fff; font-weight: 700;
  font-size: 1.3rem; display: flex; align-items: center; justify-content: center;
  background: linear-gradient(135deg, #4aa8f0, #1a7fd6); box-shadow: 0 4px 12px rgba(26,140,220,.35); }
.cx-statut { font-size: .7rem; letter-spacing: .08em; color: #6b7280; text-transform: uppercase; }
.cx-point { display: inline-block; width: 10px; height: 10px; border-radius: 50%; background: #10b981; margin-right: 6px; }
.cx-nom { font-weight: 700; font-size: 1.05rem; color: #111; word-break: break-word; }
.cx-nav { font-size: .8rem; color: #6b7280; margin: 1rem 0 .2rem; font-weight: 600; }
.st-key-btn_logout button { color: #d92d20; border: 1px solid #f4b4ae; background: #fff; }
.st-key-btn_logout button:hover { background: #fef3f2; color: #b42318; border-color: #f4b4ae; }
</style>
"""


def verifier_acces():
    """Bloque l'application tant que le bon nom d'utilisateur et le bon mot de passe n'ont pas été saisis.

    Le mot de passe (APP_PASSWORD) et le nom d'utilisateur (APP_USER, « controle de gestion » par défaut)
    viennent des secrets Streamlit, jamais du code.
    Si aucun mot de passe n'est configuré, l'application reste bloquée.
    """
    if st.session_state.get("auth_ok"):
        return

    try:
        attendu = str(st.secrets.get("APP_PASSWORD", ""))
        utilisateur_attendu = str(st.secrets.get("APP_USER", "controle de gestion"))
    except Exception:
        attendu = ""
        utilisateur_attendu = "controle de gestion"

    st.markdown(CSS_CONNEXION, unsafe_allow_html=True)
    st.markdown(
        '<div class="cx-entete"><div class="cx-icone">📊</div>'
        "<h1>Heureux de vous revoir</h1>"
        "<p>Connectez-vous pour accéder au contrôle Voyageurs / GMAO.</p></div>",
        unsafe_allow_html=True,
    )

    if st.session_state.pop("session_expiree", False):
        st.warning(
            f"Session fermée après {delai_inactivite_minutes():g} minute(s) d'inactivité : "
            "les données ont été effacées. Reconnecte-toi pour continuer."
        )

    if not attendu:
        st.error("Aucun mot de passe n'est configuré (secret APP_PASSWORD). L'application reste bloquée.")
        st.stop()

    if st.session_state.get("tentatives", 0) >= 5:
        st.error("Trop de tentatives. Recharge la page pour réessayer.")
        st.stop()

    with st.form("login"):
        utilisateur = st.text_input("Identifiant", placeholder="Votre identifiant")
        mdp = st.text_input("Mot de passe", type="password", placeholder="Votre mot de passe")
        ok = st.form_submit_button("Se connecter")
    st.markdown('<p class="cx-pied">Accès réservé aux utilisateurs autorisés</p>', unsafe_allow_html=True)

    if ok:
        # Les deux comparaisons sont toujours faites (pas de retour anticipé) ; l'utilisateur ignore casse et espaces autour.
        user_ok = hmac.compare_digest(utilisateur.strip().casefold().encode("utf-8"),
                                      utilisateur_attendu.strip().casefold().encode("utf-8"))
        mdp_ok = hmac.compare_digest(mdp.encode("utf-8"), attendu.encode("utf-8"))
        if user_ok and mdp_ok:
            st.session_state["auth_ok"] = True
            st.session_state["utilisateur"] = utilisateur_attendu.strip()
            st.session_state["last_activity"] = time.time()
            st.rerun()
        else:
            st.session_state["tentatives"] = st.session_state.get("tentatives", 0) + 1
            time.sleep(1)
            st.error("Nom d'utilisateur ou mot de passe incorrect.")
    st.stop()


def verifier_inactivite():
    """Efface les données et ferme la session si l'inactivité dépasse le délai."""
    dernier = st.session_state.get("last_activity", time.time())
    if time.time() - dernier > delai_inactivite_minutes() * 60:
        effacer_donnees()
        st.session_state["auth_ok"] = False
        st.session_state["session_expiree"] = True
        st.rerun()  # relance toute l'application -> retour à l'écran de connexion


@st.fragment(run_every="30s")
def surveillance_inactivite():
    """Minuteur : vérifie toutes les 30 s, même si l'utilisateur ne touche à rien.

    Il ne met PAS à jour l'heure d'activité : seules les actions de l'utilisateur le font.
    """
    verifier_inactivite()


def bloc_telechargement(espace, fichiers):
    """Boutons de téléchargement des fichiers de résultat.

    `fichiers` : liste de (libellé du bouton, nom du fichier, contenu).
    """
    st.subheader("Téléchargement")
    colonnes = st.columns(len(fichiers))
    for col, (libelle, nom, contenu) in zip(colonnes, fichiers):
        col.download_button(libelle, contenu, file_name=nom, mime=XLSX, width="stretch", key=f"dl_{espace}_{nom}")


def demarrer_traitement(espace):
    """À appeler quand un traitement démarre : efface l'ancien résultat de cette page."""
    for k in (f"res_{espace}", f"enc_{espace}", f"enc_cle_{espace}"):
        st.session_state.pop(k, None)


def terminer_traitement(espace, res):
    st.session_state[f"res_{espace}"] = res
    st.session_state[f"run_id_{espace}"] = st.session_state.get(f"run_id_{espace}", 0) + 1
    st.session_state["last_activity"] = time.time()   # un long traitement ne doit pas compter comme inactivité


# ==========================================================
# PAGE 1 : Détection TB et commandes (TB de correspondance vs GMAO)
# ==========================================================
COLS_TB_OBLIGATOIRES = ["NUM_CMDE_AN", "TITRE_CMDE", "TITRE", "DES_ETAT_CMDE", "DATE_CMDE"]
COLS_TB_OPTIONNELLES = ["NOM", "QTE_CMDEE", "PRIXHT_LIGNE_CMDE", "PRIX_TOTAL_HT", "DES_MAGASIN"]
COLS_GMAO_OBLIGATOIRES = ["NUM_CMDE_AN", "TITRE_CMDE", "TITRE", "DES_ETAT_CMDE", "DATE_CMDE"]
COLS_GMAO_OPTIONNELLES = ["NOM", "QTE_CMDEE", "PRIXHT_LIGNE_CMDE", "PRIX_TOTAL_HT", "TOTALHT", "TOTALTVA",
                          "TYPE_LIGNE_CMDE", "DES_FAMILLE_FOURN", "DES_MAGASIN"]


def verifier_colonnes(tb_bytes, cmd_bytes):
    """Contrôle les noms de colonnes AVANT le traitement (ne touche pas à la logique de calcul).

    Renvoie (erreurs, avertissements). Les erreurs bloquent le lancement ; les avertissements non.
    Seuls des noms de colonnes et de feuilles sont cités, jamais de données.
    """
    erreurs, avertissements = [], []

    # --- TB : feuilles 2 à 10, en-têtes nettoyés des espaces comme dans le traitement ---
    try:
        xls = pd.ExcelFile(io.BytesIO(tb_bytes))
    except Exception:
        return ["Le TB de correspondance n'est pas un fichier Excel lisible."], []
    feuilles = xls.sheet_names[1:10]
    if not feuilles:
        erreurs.append("TB de correspondance : il faut au moins 2 feuilles (la 1re est ignorée).")
    for feuille in feuilles:
        cols = {str(c).strip() for c in pd.read_excel(xls, sheet_name=feuille, nrows=0).columns}
        manque = [c for c in COLS_TB_OBLIGATOIRES if c not in cols]
        if manque:
            erreurs.append(f"TB, feuille « {feuille} » : colonne(s) obligatoire(s) absente(s) : {', '.join(manque)}")
        manque_opt = [c for c in COLS_TB_OPTIONNELLES if c not in cols]
        if manque_opt:
            avertissements.append(f"TB, feuille « {feuille} » : colonne(s) absente(s) : {', '.join(manque_opt)}")

    # --- GMAO : 1re feuille, noms exacts (le traitement ne retire pas les espaces ici) ---
    try:
        cols = [str(c) for c in pd.read_excel(io.BytesIO(cmd_bytes), nrows=0).columns]
    except Exception:
        return erreurs + ["Le fichier des commandes n'est pas un fichier Excel lisible."], avertissements
    for c in COLS_GMAO_OBLIGATOIRES:
        if c not in cols:
            indice = [x for x in cols if x.strip().upper() == c]
            suite = f" (trouvée avec des espaces : « {indice[0]} »)" if indice else ""
            erreurs.append(f"Fichier des commandes : colonne obligatoire absente : {c}{suite}")
    manque_opt = [c for c in COLS_GMAO_OPTIONNELLES if c not in cols]
    if manque_opt:
        avertissements.append(f"Fichier des commandes : colonne(s) absente(s) : {', '.join(manque_opt)}")
    return erreurs, avertissements


def fmt_montant(x):
    """1234567.8 -> « 1 234 567,80 » (format français)."""
    return f"{x:,.2f}".replace(",", " ").replace(".", ",")


def fmt_ecart(x):
    """Écart signé au format français : +1 234,50 / −300 / 0."""
    if abs(x) < 1e-9:
        return "0"
    texte = fmt_montant(abs(x)) if abs(x) % 1 > 1e-9 else f"{abs(x):,.0f}".replace(",", " ")
    return ("+" if x > 0 else "−") + texte


def section_indicateurs(a):
    """Indicateurs complémentaires : KPI et graphiques sur les changements et les nouvelles lignes."""
    lignes, synthese = a["lignes"], a["synthese_tb"]
    st.markdown("**Indicateurs complémentaires**")

    total_tb = float(synthese["TOTAL_HT"].sum())
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Montant total du TB", fmt_montant(total_tb))
    if not lignes.empty:
        part = (a["total_ht"] / total_tb * 100) if total_tb else 0.0
        k2.metric("Nouveaux montants / montant du TB", f"{part:.1f} %".replace(".", ","))
        k3.metric("Montant moyen par nouvelle ligne", fmt_montant(a["total_ht"] / len(lignes)))
        par_cmde = lignes.groupby("NUM_CMDE_AN")["PRIX_TOTAL_HT"].sum()
        k4.metric("Plus grosse nouvelle commande", fmt_montant(float(par_cmde.max())),
                  help=f"Commande {par_cmde.idxmax()}")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Changements détectés par type**")
        ch = a["changements_par_type"]
        if ch["NB"].sum() == 0:
            st.info("Aucun changement détecté.")
        else:
            st.altair_chart(
                alt.Chart(ch).mark_bar().encode(
                    x=alt.X("TYPE:N", title=None, sort=list(ch["TYPE"]), axis=alt.Axis(labelAngle=0)),
                    y=alt.Y("NB:Q", title="Lignes", axis=alt.Axis(tickMinStep=1)),
                    color=alt.Color("TYPE:N", legend=None, scale=alt.Scale(
                        domain=["État", "Fournisseur", "Quantité", "Prix"],
                        range=["#FFC000", "#00B0F0", "#FFAEC9", "#FF99CC"])),
                    tooltip=[alt.Tooltip("TYPE:N", title="Type"), alt.Tooltip("NB:Q", title="Lignes", format="d")],
                ).properties(height=300),
                width="stretch")
            st.caption("Une ligne qui cumule plusieurs changements est comptée dans chaque type.")
    with c2:
        st.markdown("**Nouvelles lignes par état de commande**")
        if lignes.empty or "DES_ETAT_CMDE" not in lignes.columns:
            st.info("Aucune nouvelle ligne.")
        else:
            etat = (lignes.assign(DES_ETAT_CMDE=lignes["DES_ETAT_CMDE"].fillna("(sans état)"))
                    .groupby("DES_ETAT_CMDE", as_index=False)
                    .agg(TOTAL_HT=("PRIX_TOTAL_HT", "sum"), NB_LIGNES=("NUM_CMDE_AN", "size")))
            st.altair_chart(
                alt.Chart(etat).mark_bar(color="#8b5cf6").encode(
                    y=alt.Y("DES_ETAT_CMDE:N", title=None, sort="-x", axis=alt.Axis(labelLimit=180)),
                    x=alt.X("TOTAL_HT:Q", title="Montant HT"),
                    tooltip=[alt.Tooltip("DES_ETAT_CMDE:N", title="État"),
                             alt.Tooltip("TOTAL_HT:Q", title="Montant HT", format=",.2f"),
                             alt.Tooltip("NB_LIGNES:Q", title="Lignes", format="d")],
                ).properties(height=300),
                width="stretch")

    if lignes.empty:
        return
    c3, c4 = st.columns(2)
    with c3:
        st.markdown("**Évolution mensuelle des nouvelles lignes**")
        dates = pd.to_datetime(lignes["DATE_CMDE"], errors="coerce") if "DATE_CMDE" in lignes.columns else pd.Series(dtype="datetime64[ns]")
        mois = (lignes.assign(MOIS=dates.dt.to_period("M").astype(str)).loc[dates.notna()]
                .groupby("MOIS", as_index=False)
                .agg(TOTAL_HT=("PRIX_TOTAL_HT", "sum"), NB_LIGNES=("NUM_CMDE_AN", "size")))
        if mois.empty:
            st.info("Aucune date de commande exploitable.")
        else:
            st.altair_chart(
                alt.Chart(mois).mark_bar(color="#10b981").encode(
                    x=alt.X("MOIS:O", title="Mois de commande", axis=alt.Axis(labelAngle=-30)),
                    y=alt.Y("TOTAL_HT:Q", title="Montant HT"),
                    tooltip=[alt.Tooltip("MOIS:O", title="Mois"),
                             alt.Tooltip("TOTAL_HT:Q", title="Montant HT", format=",.2f"),
                             alt.Tooltip("NB_LIGNES:Q", title="Lignes", format="d")],
                ).properties(height=300),
                width="stretch")
    with c4:
        st.markdown("**Top 10 des nouvelles commandes par montant**")
        top = (lignes.assign(CMDE=lignes["NUM_CMDE_AN"].astype(str))
               .groupby("CMDE", as_index=False)
               .agg(TOTAL_HT=("PRIX_TOTAL_HT", "sum"), NB_LIGNES=("NUM_CMDE_AN", "size"))
               .nlargest(10, "TOTAL_HT"))
        st.altair_chart(
            alt.Chart(top).mark_bar(color="#f97316").encode(
                y=alt.Y("CMDE:N", title="Commande", sort="-x"),
                x=alt.X("TOTAL_HT:Q", title="Montant HT"),
                tooltip=[alt.Tooltip("CMDE:N", title="Commande"),
                         alt.Tooltip("TOTAL_HT:Q", title="Montant HT", format=",.2f"),
                         alt.Tooltip("NB_LIGNES:Q", title="Lignes", format="d")],
            ).properties(height=300),
            width="stretch")


def section_analyse(a, run_id):
    """Section « Analyse des nouvelles lignes » : détail, total, et regroupement par feuille."""
    st.subheader("Analyse des nouvelles lignes")
    lignes, par_feuille = a["lignes"], a["par_feuille"]
    if lignes.empty:
        st.info("Aucune nouvelle ligne de commande détectée.")
    else:
        m1, m2, m3 = st.columns(3)
        m1.metric("Nouvelles lignes", len(lignes))
        m2.metric("Commandes distinctes (NUM_CMDE_AN)", a["nb_commandes"])
        m3.metric("Montant total PRIX_TOTAL_HT", fmt_montant(a["total_ht"]))
        if a["nb_sans_montant"]:
            st.warning(f"{a['nb_sans_montant']} nouvelle(s) ligne(s) sans PRIX_TOTAL_HT : elles ne sont pas dans les sommes.")

    st.markdown("**Par feuille du TB**")
    synthese = a["synthese_tb"]
    st.dataframe(
        synthese, width="stretch", hide_index=True,
        column_config={
            "FEUILLE": "Feuille",
            "NB_COMMANDES": st.column_config.NumberColumn("Commandes distinctes", format="%d"),
            "TOTAL_HT": st.column_config.NumberColumn("Somme PRIX_TOTAL_HT", format="%.2f"),
        },
    )
    graph = synthese
    g1, g2 = st.columns(2)
    with g1:
        st.markdown("**Répartition du montant PRIX_TOTAL_HT par feuille du TB**")
        donut = (
            alt.Chart(graph).mark_arc(innerRadius=70)
            .encode(
                theta=alt.Theta("TOTAL_HT:Q"),
                color=alt.Color("FEUILLE:N", title="Feuille",
                                legend=alt.Legend(orient="right", labelLimit=220, symbolType="circle")),
                tooltip=[alt.Tooltip("FEUILLE:N", title="Feuille"),
                         alt.Tooltip("TOTAL_HT:Q", title="Montant HT", format=",.2f")],
            )
            .properties(height=320)
        )
        st.altair_chart(donut, width="stretch")
    with g2:
        st.markdown("**Top 10 fournisseurs par montant HT**")
        rattachees = lignes[lignes["FEUILLE"].notna()] if not lignes.empty else lignes   # comme le reste de l'analyse par feuille : sans les commandes hors TB
        if not rattachees.empty and "NOM" in rattachees.columns and rattachees["NOM"].notna().any():
            top = (rattachees.assign(NOM=rattachees["NOM"].fillna("(sans fournisseur)"))
                   .groupby("NOM", as_index=False)
                   .agg(TOTAL_HT=("PRIX_TOTAL_HT", "sum"), NB_COMMANDES=("NUM_CMDE_AN", "nunique"))
                   .nlargest(10, "TOTAL_HT"))
            barres = (
                alt.Chart(top).mark_bar(color="#1a8cdc")
                .encode(
                    y=alt.Y("NOM:N", title=None, sort="-x", axis=alt.Axis(labelLimit=180)),
                    x=alt.X("TOTAL_HT:Q", title="Montant HT"),
                    tooltip=[alt.Tooltip("NOM:N", title="Fournisseur"),
                             alt.Tooltip("TOTAL_HT:Q", title="Montant HT", format=",.2f"),
                             alt.Tooltip("NB_COMMANDES:Q", title="Commandes", format="d")],
                )
                .properties(height=320)
            )
            st.altair_chart(barres, width="stretch")
        else:
            st.info("Aucune donnée fournisseur à afficher pour les lignes rattachées à une feuille du TB.")

    section_indicateurs(a)

    if lignes.empty:
        return
    st.markdown("**Détail des nouvelles lignes**")
    choix = st.multiselect("Filtrer par feuille", list(par_feuille["FEUILLE"]), key=f"flt_feuille_{run_id}")
    detail = lignes[lignes["FEUILLE"].isin(choix)] if choix else lignes
    detail = detail.dropna(axis=1, how="all").copy()
    if "FEUILLE" in detail.columns:
        detail["FEUILLE"] = detail["FEUILLE"].fillna("—")
    else:
        detail.insert(1, "FEUILLE", "—")
    detail["NUM_CMDE_AN"] = detail["NUM_CMDE_AN"].astype("string")   # évite les soucis d'affichage (types mixtes)
    st.dataframe(detail, width="stretch", hide_index=True,
                 column_config={"PRIX_TOTAL_HT": st.column_config.NumberColumn(format="%.2f")})


def page_detection():
    n = st.session_state.setdefault("reset_n", 0)

    st.title("🔎 Détection : TB et commandes (GMAO)")
    st.caption(
        "Dépose les deux fichiers d'entrée et lance le contrôle"
    )

    col1, col2 = st.columns(2)
    with col1:
        f_tb = st.file_uploader("1. TB de correspondance", type=["xlsx"], key=f"tb_{n}")
    with col2:
        f_cmd = st.file_uploader("2. Commandes GMAO", type=["xlsx"], key=f"cmd_{n}")

    lancer = st.button("Lancer le contrôle", type="primary", disabled=not (f_tb and f_cmd), key="go_det")

    if lancer:
        erreurs, avertissements = verifier_colonnes(f_tb.getvalue(), f_cmd.getvalue())
        if erreurs:
            st.error("Traitement non lancé : fichiers non conformes.\n\n" + "\n".join(f"- {e}" for e in erreurs))
            st.stop()
        if avertissements:
            st.warning("Colonnes facultatives absentes (le traitement continue, mais certains contrôles seront incomplets) :\n\n"
                       + "\n".join(f"- {a}" for a in avertissements))
        demarrer_traitement("det")
        with st.status("Traitement en cours…", expanded=True) as status:
            try:
                res = run(f_tb.getvalue(), f_cmd.getvalue(), log=st.write)
            except Exception as e:
                # Le journal serveur ne garde que le type d'erreur : jamais de contenu des fichiers.
                logging.getLogger("controle").error("Erreur de traitement (%s)", type(e).__name__)
                status.update(label="Erreur pendant le traitement", state="error")
                st.error("Le traitement a échoué. Vérifie que ce sont les bons fichiers (feuilles et colonnes attendues).")
                st.stop()
            status.update(label="Traitement terminé", state="complete", expanded=False)
        terminer_traitement("det", res)

    res = st.session_state.get("res_det")
    if res:
        m1, m2, m3 = st.columns(3)
        m1.metric("Lignes avec changement", res["nb_changements"])
        m2.metric("Nouvelles lignes", res["nb_nouvelles"])
        m3.metric("Lignes du rapprochement", res["nb_rapprochement"])

        section_analyse(res["analyse"], st.session_state.get("run_id_det"))

        st.subheader("Aperçu")
        feuilles = pd.read_excel(io.BytesIO(res["resultat"]), sheet_name=None)
        rapproch = pd.read_excel(io.BytesIO(res["rapprochement"]))
        tab1, tab2, tab3 = st.tabs(["Audit complet", "Nouvelles lignes", "Rapprochement des titres"])
        with tab1:
            st.dataframe(feuilles["Audit complet"], width="stretch")
        with tab2:
            st.dataframe(feuilles["Nouvelles lignes"], width="stretch")
        with tab3:
            st.dataframe(rapproch, width="stretch")


def _barres_h(df, y, x, titre_x, couleur, tooltips, tri="-x", hauteur=300):
    """Barres horizontales Altair (utilisé par les graphiques de la page Suivi)."""
    return st.altair_chart(
        alt.Chart(df).mark_bar(color=couleur).encode(
            y=alt.Y(f"{y}:N", title=None, sort=tri, axis=alt.Axis(labelLimit=220)),
            x=alt.X(f"{x}:Q", title=titre_x, axis=alt.Axis(tickMinStep=1) if titre_x in ("Lignes", "Nombre de lignes") else alt.Axis()),
            tooltip=tooltips,
        ).properties(height=hauteur),
        width="stretch")


def graphiques_suivi(res):
    """Graphiques de la page Suivi : changements, transitions d'état, écarts de montant, nouvelles lignes."""
    nouvelles, changements = res["nouvelles"].copy(), res["changements"].copy()
    st.markdown("**Graphiques**")

    # ---------- Ligne 1 : changements par champ | transitions d'état ----------
    g1, g2 = st.columns(2)
    with g1:
        st.markdown("**Changements par champ**")
        rep = res["stats"]["repartition"]
        if not rep:
            st.info("Aucun changement détecté.")
        else:
            libelles = {"Etat": "État", "Quantité": "Quantité", "MONTANT BC": "Montant"}
            df = pd.DataFrame({"CHAMP": [libelles.get(k, k) for k in rep], "NB": list(rep.values())})
            st.altair_chart(
                alt.Chart(df).mark_bar().encode(
                    x=alt.X("CHAMP:N", title=None, sort="-y", axis=alt.Axis(labelAngle=0)),
                    y=alt.Y("NB:Q", title="Lignes", axis=alt.Axis(tickMinStep=1)),
                    color=alt.Color("CHAMP:N", legend=None, scale=alt.Scale(
                        domain=["État", "Quantité", "Montant"], range=["#FFC000", "#FFAEC9", "#FF99CC"])),
                    tooltip=[alt.Tooltip("CHAMP:N", title="Champ"), alt.Tooltip("NB:Q", title="Lignes", format="d")],
                ).properties(height=300),
                width="stretch")
    with g2:
        st.markdown("**Transitions d'état (suivi → TB)**")
        etats = changements[changements["Champ"] == "Etat"] if len(changements) else changements
        if etats.empty:
            st.info("Aucun changement d'état détecté.")
        else:
            def propre(v):
                return v.astype("string").fillna("").str.strip().str.capitalize().replace("", "(vide)")
            trans = (propre(etats["Valeur_BDD"]) + " → " + propre(etats["Valeur_TB"])).value_counts().head(10)
            df = pd.DataFrame({"TRANSITION": trans.index, "NB": trans.values})
            _barres_h(df, "TRANSITION", "NB", "Lignes", "#8b5cf6",
                      [alt.Tooltip("TRANSITION:N", title="Transition"), alt.Tooltip("NB:Q", title="Lignes", format="d")])

    # ---------- Ligne 2 : plus gros écarts de montant | évolution mensuelle des nouvelles lignes ----------
    g3, g4 = st.columns(2)
    with g3:
        st.markdown("**Top 10 des écarts de montant**")
        mt = changements[changements["Champ"] == "MONTANT BC"].copy() if len(changements) else changements
        if mt.empty:
            st.info("Aucun changement de montant détecté.")
        else:
            mt["SUIVI"] = convertir_numerique(mt["Valeur_BDD"]).fillna(0.0)
            mt["TB"] = convertir_numerique(mt["Valeur_TB"]).fillna(0.0)
            mt["ECART"] = mt["TB"] - mt["SUIVI"]
            mt = mt.reindex(mt["ECART"].abs().sort_values(ascending=False).index).head(10)
            etiquette = mt["NUM_CMDE_AN"].astype(str) + " · " + mt["TITRE"].astype(str).str.slice(0, 28)
            mt["LIGNE"] = etiquette + mt.groupby(etiquette).cumcount().map(lambda i: "" if i == 0 else f" ({i + 1})")
            st.altair_chart(
                alt.Chart(mt).mark_bar().encode(
                    y=alt.Y("LIGNE:N", title=None, sort=list(mt["LIGNE"]), axis=alt.Axis(labelLimit=240)),
                    x=alt.X("ECART:Q", title="Écart (TB − suivi)"),
                    color=alt.condition(alt.datum.ECART > 0, alt.value("#10b981"), alt.value("#ef4444")),
                    tooltip=[alt.Tooltip("LIGNE:N", title="Ligne"),
                             alt.Tooltip("SUIVI:Q", title="Montant suivi", format=",.2f"),
                             alt.Tooltip("TB:Q", title="Montant TB", format=",.2f"),
                             alt.Tooltip("ECART:Q", title="Écart", format="+,.2f")],
                ).properties(height=300),
                width="stretch")
            st.caption("Vert : le TB est supérieur au suivi · Rouge : le TB est inférieur au suivi.")
    with g4:
        st.markdown("**Évolution mensuelle des nouvelles lignes**")
        if nouvelles.empty:
            st.info("Aucune nouvelle ligne.")
        else:
            nouvelles["MONTANT"] = convertir_numerique(nouvelles["PRIX_TOTAL_HT"]).fillna(0.0)
            dates = pd.to_datetime(nouvelles["DATE_CMDE"], errors="coerce")
            mois = (nouvelles.assign(MOIS=dates.dt.to_period("M").astype(str)).loc[dates.notna()]
                    .groupby("MOIS", as_index=False)
                    .agg(MONTANT=("MONTANT", "sum"), NB=("MONTANT", "size")))
            if mois.empty:
                st.info("Aucune date de commande exploitable.")
            else:
                st.altair_chart(
                    alt.Chart(mois).mark_bar(color="#10b981").encode(
                        x=alt.X("MOIS:O", title="Mois de commande", axis=alt.Axis(labelAngle=-30)),
                        y=alt.Y("MONTANT:Q", title="Montant HT"),
                        tooltip=[alt.Tooltip("MOIS:O", title="Mois"),
                                 alt.Tooltip("MONTANT:Q", title="Montant HT", format=",.2f"),
                                 alt.Tooltip("NB:Q", title="Lignes", format="d")],
                    ).properties(height=300),
                    width="stretch")

    # ---------- Ligne 3 : top 10 fournisseurs | nouvelles lignes par état ----------
    g5, g6 = st.columns(2)
    with g5:
        st.markdown("**Top 10 fournisseurs des nouvelles lignes**")
        if nouvelles.empty or "NOM" not in nouvelles.columns:
            st.info("Aucune nouvelle ligne.")
        else:
            nouvelles["MONTANT"] = convertir_numerique(nouvelles["PRIX_TOTAL_HT"]).fillna(0.0)
            top = (nouvelles.assign(NOM=nouvelles["NOM"].astype("string").fillna("(sans fournisseur)"))
                   .groupby("NOM", as_index=False).agg(MONTANT=("MONTANT", "sum"), NB=("MONTANT", "size"))
                   .nlargest(10, "MONTANT"))
            _barres_h(top, "NOM", "MONTANT", "Montant HT", "#1a8cdc",
                      [alt.Tooltip("NOM:N", title="Fournisseur"), alt.Tooltip("MONTANT:Q", title="Montant HT", format=",.2f"),
                       alt.Tooltip("NB:Q", title="Lignes", format="d")])
    with g6:
        st.markdown("**Nouvelles lignes par état de commande**")
        if nouvelles.empty or "DES_ETAT_CMDE" not in nouvelles.columns:
            st.info("Aucune nouvelle ligne.")
        else:
            nouvelles["MONTANT"] = convertir_numerique(nouvelles["PRIX_TOTAL_HT"]).fillna(0.0)
            etat = (nouvelles.assign(ETAT=nouvelles["DES_ETAT_CMDE"].astype("string").fillna("(sans état)")
                                     .str.strip().str.capitalize().replace("", "(sans état)"))
                    .groupby("ETAT", as_index=False).agg(MONTANT=("MONTANT", "sum"), NB=("MONTANT", "size")))
            _barres_h(etat, "ETAT", "MONTANT", "Montant HT", "#f97316",
                      [alt.Tooltip("ETAT:N", title="État"), alt.Tooltip("MONTANT:Q", title="Montant HT", format=",.2f"),
                       alt.Tooltip("NB:Q", title="Lignes", format="d")])


# ==========================================================
# PAGE 2 : Suivi détecté (TB de correspondance vs suivi budgétaire d'un service)
# ==========================================================
def page_suivi():
    n = st.session_state.setdefault("reset_n", 0)

    st.title("📋 Suivi détecté : TB et suivi budgétaire")
    st.caption(
        "Compare le TB de correspondance au fichier de suivi budgétaire d'un service : "
        "nouvelles lignes et changements (état, quantité, montant)."
    )

    service = st.selectbox("Service", list(PROFILS_SUIVI), key="suivi_service")
    profil = PROFILS_SUIVI[service]
    with st.expander("Règles appliquées à ce service"):
        st.markdown(profil.description)

    col1, col2 = st.columns(2)
    with col1:
        f_tb = st.file_uploader("1. TB de correspondance (.xlsx)", type=["xlsx"], key=f"s_tb_{n}")
    with col2:
        f_bdd = st.file_uploader(f"2. Suivi budgétaire {service} (.xlsm ou .xlsx)",
                                 type=["xlsm", "xlsx"], key=f"s_bdd_{n}")

    with st.expander("Paramètres"):
        annee = st.number_input("Année des commandes à contrôler", min_value=2000, max_value=2100,
                                value=2026, step=1, key="suivi_annee")
        feuille_bdd = st.text_input("Feuille du fichier de suivi budgétaire", value="BDD 2026", key="suivi_feuille")

    lancer = st.button("Lancer le suivi", type="primary", disabled=not (f_tb and f_bdd), key="go_suivi")

    if lancer:
        demarrer_traitement("suivi")
        with st.status("Traitement en cours…", expanded=False) as status:
            try:
                res = executer_suivi(f_tb.getvalue(), f_bdd.getvalue(), profil,
                                     annee=annee, feuille_bdd=feuille_bdd.strip())
            except ErreurSuivi as e:
                # Message explicite (feuille ou colonne manquante...) : il ne contient pas de données.
                status.update(label="Traitement impossible", state="error")
                st.error(str(e))
                st.stop()
            except Exception as e:
                logging.getLogger("controle").error("Erreur de suivi (%s)", type(e).__name__)
                status.update(label="Erreur pendant le traitement", state="error")
                st.error("Le traitement a échoué. Vérifie que ce sont les bons fichiers et le bon service.")
                st.stop()
            status.update(label="Traitement terminé", state="complete", expanded=False)
        res["service"] = service
        res["date"] = datetime.now().strftime("%Y%m%d_%H%M")
        terminer_traitement("suivi", res)

    res = st.session_state.get("res_suivi")
    if res:
        st.subheader(f"Résultats : {res['service']}")
        s = res["stats"]
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Lignes du suivi", s["lignes_bdd"])
        m2.metric("Lignes du TB", s["lignes_tb"])
        m3.metric("Nouvelles lignes", s["nouvelles"], help=f"{s['commandes_distinctes']} commande(s) distincte(s)")
        m4.metric("Changements", s["changements"])
        if s["repartition"]:
            st.caption("Changements par champ : " + " · ".join(f"{k} : {v}" for k, v in s["repartition"].items()))

        k = s.get("kpi")
        if k:
            st.markdown("**Indicateurs**")
            a1, a2, a3 = st.columns(3)
            a1.metric("Montant des nouvelles lignes", fmt_montant(k["montant_nouvelles"]),
                      help="Somme de PRIX_TOTAL_HT des lignes du TB absentes du suivi (à ajouter au suivi).")
            a2.metric("Écart net de montant", fmt_ecart(k["ecart_montant"]),
                      help="Somme de (montant TB - montant suivi) sur les lignes dont le montant diffère. "
                           "Positif : le suivi est inférieur au TB.")
            a3.metric("Écart net de quantité", fmt_ecart(k["ecart_qte"]),
                      help="Somme de (quantité TB - quantité suivi) sur les lignes dont la quantité diffère.")
            b1, b2, b3 = st.columns(3)
            b1.metric("Commandes touchées", k["commandes_touchees"],
                      help="Nombre de commandes distinctes ayant au moins un changement.")
            b2.metric("Taux de lignes en changement", f"{k['taux_changement']:.1f} %".replace(".", ","),
                      help=f"{k['lignes_en_changement']} ligne(s) distincte(s) avec au moins un changement, "
                           "divisées par le nombre de lignes du suivi.")
            gros = k["plus_gros_ecart"]
            b3.metric("Plus gros écart de montant", fmt_ecart(gros["valeur"]) if gros else "—",
                      help=f"Commande {gros['commande']}" if gros else "Aucun changement de montant.")

        graphiques_suivi(res)

        avertissements = [m for niveau, m in res["journal"] if niveau == "warning"]
        if avertissements:
            with st.expander(f"⚠️ {len(avertissements)} avertissement(s) à lire"):
                for m in avertissements:
                    st.write("• " + m)
        with st.expander("Détail du traitement"):
            for niveau, m in res["journal"]:
                st.write(("• " if niveau == "info" else "⚠️ ") + m)

        bloc_telechargement("suivi", [
            (f"⬇️ Télécharger Resultats_{res['service']}.xlsx",
             f"Resultats_{res['service']}_{res['date']}.xlsx", res["xlsx"]),
        ])

        st.subheader("Aperçu")
        # Aperçu uniquement : les valeurs (texte pour l'état, nombres pour quantité/montant) sont affichées
        # en texte pour que l'affichage soit homogène. Le fichier téléchargé n'est pas modifié.
        apercu_changements = res["changements"].copy()
        for col in ("Valeur_BDD", "Valeur_TB"):
            if col in apercu_changements.columns:
                apercu_changements[col] = apercu_changements[col].astype("string")
        tab1, tab2 = st.tabs(["Nouvelles", "Changements"])
        with tab1:
            st.dataframe(res["nouvelles"], width="stretch")
        with tab2:
            st.dataframe(apercu_changements, width="stretch")


# ==========================================================
# POINT D'ENTRÉE : accès, inactivité, barre latérale, navigation
# ==========================================================
verifier_acces()
verifier_inactivite()                              # avant de marquer l'activité : l'ordre compte
st.session_state["last_activity"] = time.time()    # cette action de l'utilisateur compte comme activité
surveillance_inactivite()

pages = [
    st.Page(page_detection, title="Détection TB et commandes", icon="🔎", url_path="detection", default=True),
    st.Page(page_suivi, title="Suivi détecté", icon="📋", url_path="suivi"),
]
navigation = st.navigation(pages, position="hidden")   # menu construit à la main dans la barre latérale

with st.sidebar:
    st.markdown(CSS_SESSION, unsafe_allow_html=True)
    nom = html.escape(st.session_state.get("utilisateur", "utilisateur"))
    st.markdown(
        f'<div class="cx-carte"><div class="cx-avatar">{nom[:1].upper()}</div>'
        f'<div><div class="cx-statut"><span class="cx-point"></span>Connecté</div>'
        f'<div class="cx-nom">{nom}</div></div></div>',
        unsafe_allow_html=True,
    )
    if st.button("Se déconnecter", icon=":material/logout:", width="stretch", key="btn_logout"):
        effacer_donnees()
        st.session_state["auth_ok"] = False
        st.session_state.pop("utilisateur", None)
        st.rerun()

    st.markdown('<div class="cx-nav">Navigation</div>', unsafe_allow_html=True)
    for p in pages:
        st.page_link(p)

    st.divider()
    if st.button("Effacer les données", icon=":material/delete:", width="stretch", key="btn_effacer"):
        effacer_donnees()
        st.rerun()
    st.caption(
        "Les fichiers sont traités en mémoire du serveur uniquement. "
        f"Effacement automatique et déconnexion après {delai_inactivite_minutes():g} min d'inactivité."
    )

navigation.run()
