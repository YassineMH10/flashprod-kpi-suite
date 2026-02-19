# app.py
# ============================================================
# ✅ PIPELINE UNIQUE : 6 CODES FUSIONNÉS (VERSION STREAMLIT)
# ✅ Upload requis : (1) fichier BRUT  +  (2) fichier COMPO
# ✅ Sorties : Excel final + Email subject + HTML
# ============================================================

import re
from io import BytesIO
from datetime import datetime, time, timedelta

import numpy as np
import pandas as pd
import streamlit as st

from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# ============================================================
# 🔧 HELPERS (robustes)
# ============================================================
def to_td(val):
    if val is None:
        return timedelta(0)
    if isinstance(val, str):
        v = val.strip()
        if v in ("-", "", "0", "00:00:00"):
            return timedelta(0)
    try:
        return pd.to_timedelta(val)
    except Exception:
        return timedelta(0)

def fmt_hms(td):
    total = int(td.total_seconds())
    sign = "-" if total < 0 else ""
    total = abs(total)
    h = total // 3600
    m = (total % 3600) // 60
    s = total % 60
    return f"{sign}{h:02d}:{m:02d}:{s:02d}"

def safe_int(x):
    try:
        return int(str(x).strip())
    except Exception:
        return 0

def to_seconds(val):
    if val in ("-", None, ""):
        return None
    if isinstance(val, time):
        return val.hour * 3600 + val.minute * 60 + val.second
    if isinstance(val, timedelta):
        return int(val.total_seconds())
    if isinstance(val, (int, float)):
        return int(val)
    if isinstance(val, str) and ":" in val:
        parts = val.split(":")
        if len(parts) == 3:
            h, m, s = parts
            return int(h) * 3600 + int(m) * 60 + int(s)
    return None

def time_to_seconds(val):
    if pd.isna(val):
        return None
    if isinstance(val, timedelta):
        return val.total_seconds()
    if isinstance(val, time):
        return val.hour * 3600 + val.minute * 60 + val.second
    if isinstance(val, str) and ":" in val:
        try:
            h, m, s = val.split(":")
            return int(h) * 3600 + int(m) * 60 + int(s)
        except Exception:
            return None
    return None

def sec_to_hms(sec):
    if sec is None or (isinstance(sec, float) and np.isnan(sec)):
        return "-"
    return str(timedelta(seconds=int(sec)))

def to_sec_series(s: pd.Series) -> pd.Series:
    s2 = s.replace("-", np.nan) if s.dtype == object else s
    return pd.to_timedelta(s2, errors="coerce").dt.total_seconds()

def ensure_unique_keep(cols):
    seen = set()
    out = []
    for c in cols:
        if c not in seen:
            out.append(c)
            seen.add(c)
    return out

def read_excel_any(uploaded_file) -> pd.DataFrame:
    """
    Lit .xlsx/.xls depuis un UploadedFile Streamlit.
    xlrd est requis pour .xls.
    """
    data = uploaded_file.read()
    bio = BytesIO(data)
    # pandas choisit l'engine selon extension si possible.
    return pd.read_excel(bio)

def workbook_to_bytes(wb) -> bytes:
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()

# ============================================================
# 🧭 UI
# ============================================================
st.set_page_config(page_title="Flash Prod Pipeline", layout="wide")
st.title("📊 Flash Prod Pipeline (BRUT + COMPO → Excel + Email)")

with st.sidebar:
    st.header("⚙️ Paramètres")
    projet = st.text_input("Nom du projet (ex: CNSS)", value="CNSS")
    date_input = st.text_input("Date Flash Prod (JJ/MM/AAAA)", value=datetime.now().strftime("%d/%m/%Y"))

    st.divider()
    st.header("📁 Uploads")
    raw_file = st.file_uploader("Fichier BRUT (source code 1)", type=["xls", "xlsx"])
    compo_file = st.file_uploader("Fichier COMPO (source code 3)", type=["xls", "xlsx"])

run_btn = st.button("🚀 Lancer le pipeline", type="primary", use_container_width=True)

# ============================================================
# ✅ PIPELINE
# ============================================================
if run_btn:
    # Checks
    if not raw_file or not compo_file:
        st.error("Il faut uploader les 2 fichiers : BRUT + COMPO.")
        st.stop()

    try:
        date_obj = datetime.strptime(date_input.strip(), "%d/%m/%Y")
    except Exception:
        st.error("Format date invalide. Utilise JJ/MM/AAAA (ex: 20/02/2026).")
        st.stop()

    projet_upper = (projet or "").strip().upper()
    if not projet_upper:
        st.error("Nom du projet vide.")
        st.stop()

    date_txt = date_obj.strftime("%d/%m/%Y")
    date_flash = date_obj.strftime("%d_%m_%Y")

    progress = st.progress(0, text="Démarrage...")
    logs = st.empty()

    def log(msg):
        logs.write(msg)

    try:
        # ============================================================
        # ✅ CODE 1 : Nettoyage brut -> rapport_nettoye.xlsx (en mémoire)
        # ============================================================
        progress.progress(5, text="Lecture BRUT + nettoyage (Code 1)...")
        df = read_excel_any(raw_file)

        df = df.dropna(how="all")
        df = df.dropna(axis=1, how="all")

        if "Unnamed: 0" in df.columns:
            df["Unnamed: 0"] = df["Unnamed: 0"].ffill()

        df = df.rename(columns={
            "Unnamed: 0": "Nom Agent",
            "Unnamed: 1": "Etat",
            "Unnamed: 4": "Occurances",
            "Unnamed: 6": "Temps total",
        })

        for col in ["Unnamed: 3", "Unnamed: 5"]:
            if col in df.columns:
                df = df.drop(columns=[col])

        cols_to_drop = []
        for col in df.columns:
            if isinstance(col, str) and col.startswith("Unnamed:"):
                try:
                    idx = int(col.split(":")[1])
                    if idx >= 7:
                        cols_to_drop.append(col)
                except Exception:
                    pass
        if cols_to_drop:
            df = df.drop(columns=cols_to_drop)

        if "Etat" in df.columns:
            df = df[df["Etat"].notna()]

        if "Nom Agent" in df.columns:
            df.insert(0, "Log Téléphonie1", df["Nom Agent"].astype(str).str.extract(r"Agent\s+(\d{4})")[0])
        else:
            df.insert(0, "Log Téléphonie1", None)

        def rename_second_pause(group):
            pauses = group[group["Etat"] == "Pause"].index
            if len(pauses) > 1:
                group.loc[pauses[1:], "Etat"] = "Pause générique"
            return group

        if "Etat" in df.columns and "Log Téléphonie1" in df.columns:
            df = df.groupby("Log Téléphonie1", group_keys=False).apply(rename_second_pause)

        def to_hms(val):
            if isinstance(val, str):
                match = re.match(r"(\d+)h(\d+)'(\d+)", val)
                if match:
                    h, m, s = match.groups()
                    return f"{int(h):02}:{int(m):02}:{int(s):02}"
            return val

        if "Temps total" in df.columns:
            df["Temps total"] = df["Temps total"].apply(to_hms)

        if "Etat" in df.columns:
            df["Etat"] = df["Etat"].replace({
                "Attente": "Attente global",
                "Pause": "Pause global",
                "Preview": "Histo Mailing",
            })
            df = df[df["Etat"].astype(str).str.lower() != "en attente"]

        log("✅ Code 1 OK")

        # ============================================================
        # ✅ CODE 2 : Pivot + KPI Agent -> rapport_final_formaté.xlsx (en mémoire)
        # ============================================================
        progress.progress(25, text="Pivot + KPI Agent (Code 2)...")
        df2 = df.copy()

        if "Log Téléphonie1" in df2.columns:
            df2.drop(columns=["Log Téléphonie1"], inplace=True)

        df2.insert(
            0,
            "Log Téléphonie1",
            df2["Nom Agent"].apply(
                lambda x: re.search(r"Agent\s+(\d{4})", str(x)).group(1)
                if re.search(r"Agent\s+(\d{4})", str(x)) else None
            )
        )

        unique_etats = df2["Etat"].dropna().unique()
        agent_base = df2[["Log Téléphonie1"]].drop_duplicates().reset_index(drop=True)

        for etat in unique_etats:
            agent_base[f"{etat} - Occurence"] = "-"
            agent_base[f"{etat} - Temps total"] = "-"

        for i, row in agent_base.iterrows():
            log_tel = row["Log Téléphonie1"]
            sub_df = df2[df2["Log Téléphonie1"] == log_tel]
            if sub_df.empty:
                continue
            for etat in unique_etats:
                bloc = sub_df[sub_df["Etat"] == etat]
                if bloc.empty:
                    continue
                occ_col = "Occurances" if "Occurances" in bloc.columns else ("Occurrences" if "Occurrences" in bloc.columns else None)
                occ_sum = safe_int(bloc[occ_col].fillna(0).sum()) if occ_col else 0
                tm_sum = bloc["Temps total"].apply(to_td).sum()
                agent_base.at[i, f"{etat} - Occurence"] = occ_sum if occ_sum != 0 else "-"
                agent_base.at[i, f"{etat} - Temps total"] = fmt_hms(tm_sum) if tm_sum != timedelta(0) else "-"

        etat_presence = ["Attente global", "Traitement", "Post-travail", "Pause global"]

        def total_presence_row(r):
            total = timedelta(0)
            for e in etat_presence:
                coln = f"{e} - Temps total"
                if coln in r:
                    total += to_td(r[coln])
            return fmt_hms(total)

        agent_base["Temps Total présence"] = agent_base.apply(total_presence_row, axis=1)

        def total_travail_row(r):
            return fmt_hms(to_td(r.get("Traitement - Temps total", "-")) + to_td(r.get("Post-travail - Temps total", "-")))

        agent_base.insert(
            agent_base.columns.get_loc("Temps Total présence") + 1,
            "Temps total Travail",
            agent_base.apply(total_travail_row, axis=1),
        )

        def taux_occupation_row(r):
            travail = to_td(r["Temps total Travail"]).total_seconds()
            presence = to_td(r["Temps Total présence"]).total_seconds()
            if presence <= 0:
                return "0.00%"
            return f"{(travail / presence) * 100:.2f}%"

        agent_base.insert(
            agent_base.columns.get_loc("Temps total Travail") + 1,
            "Taux d'occupation",
            agent_base.apply(taux_occupation_row, axis=1),
        )

        def productivite_row(r):
            occ = safe_int(r.get("Traitement - Occurence", 0))
            presence_h = to_td(r["Temps Total présence"]).total_seconds() / 3600
            if presence_h <= 0:
                return 0
            return round(occ / presence_h, 2)

        agent_base["Productivité"] = agent_base.apply(productivite_row, axis=1)

        def calc_dmc(r):
            occ = safe_int(r.get("Traitement - Occurence", 0))
            if occ <= 0:
                return "00:00:00"
            total = to_td(r.get("Traitement - Temps total", "00:00:00")) / occ
            return fmt_hms(total)

        agent_base["DMC"] = agent_base.apply(calc_dmc, axis=1)

        def calc_dmt(r):
            occ = safe_int(r.get("Traitement - Occurence", 0))
            if occ <= 0:
                return "00:00:00"
            total = (to_td(r.get("Traitement - Temps total", "00:00:00")) + to_td(r.get("Post-travail - Temps total", "00:00:00"))) / occ
            return fmt_hms(total)

        agent_base["DMT"] = agent_base.apply(calc_dmt, axis=1)

        wb2 = Workbook()
        ws2 = wb2.active
        ws2.title = "Résultat Final"

        for r_idx, row in enumerate(dataframe_to_rows(agent_base, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws2.cell(row=r_idx, column=c_idx, value=value)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if r_idx == 1:
                    cell.font = Font(bold=True)

        for col in ws2.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            ws2.column_dimensions[col_letter].width = max_length + 2

        rapport_final_bytes = workbook_to_bytes(wb2)
        log("✅ Code 2 OK")

        # ============================================================
        # ✅ CODE 3 : Merge COMPO + Rapport -> fichier_final_formaté.xlsx (en mémoire)
        # ============================================================
        progress.progress(45, text="Merge COMPO + rapport (Code 3)...")
        df_compo = read_excel_any(compo_file)
        df_rapport = pd.read_excel(BytesIO(rapport_final_bytes))

        df_merged = pd.merge(df_compo, df_rapport, on="Log Téléphonie1", how="left")

        df_merged = df_merged[df_merged["Nom Agent"].notna() & (df_merged["Nom Agent"] != "")]
        if "Matricule" in df_merged.columns:
            df_merged = df_merged.drop(columns=["Matricule"])

        colonnes_apres_ops = ["Temps Total présence", "Temps total Travail", "Taux d'occupation", "Productivité", "DMC", "DMT"]

        ops_index = df_merged.columns.get_loc("OPS") + 1 if "OPS" in df_merged.columns else 0
        colonnes_sans = [col for col in df_merged.columns if col not in colonnes_apres_ops]

        nouvel_ordre = (
            colonnes_sans[:ops_index]
            + [c for c in colonnes_apres_ops if c in df_merged.columns]
            + colonnes_sans[ops_index:]
        )

        df_final3 = df_merged[nouvel_ordre]

        bio3 = BytesIO()
        df_final3.to_excel(bio3, index=False)
        bio3.seek(0)

        wb3 = load_workbook(bio3)
        ws3 = wb3.active

        bold_font = Font(bold=True)
        center_align = Alignment(horizontal="center", vertical="center")
        thin_border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin"),
        )

        for row in ws3.iter_rows():
            for cell in row:
                cell.alignment = center_align
                cell.border = thin_border
                if cell.row == 1:
                    cell.font = bold_font

        table_range = f"A1:{get_column_letter(ws3.max_column)}{ws3.max_row}"
        table = Table(displayName="TableFinale", ref=table_range)
        style = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        table.tableStyleInfo = style
        ws3.add_table(table)

        for col in ws3.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws3.column_dimensions[col_letter].width = max_length + 2

        fichier_final_bytes = workbook_to_bytes(wb3)
        log("✅ Code 3 OK")

        # ============================================================
        # ✅ CODE 4 : Flash Prod Agent final + Taux Post + Moy Post + règle <10min
        # ============================================================
        progress.progress(60, text="Flash Prod Agent (Code 4)...")
        df4 = pd.read_excel(BytesIO(fichier_final_bytes))

        columns_to_keep = ensure_unique_keep([
            "Matricule RH","Log Téléphonie1","Nom Agent","File","Tls","OPS",
            "Temps Total présence","Temps total Travail","Taux d'occupation",
            "Productivité","DMC","DMT","Attente global - Temps total",
            "Appel entrant - Occurence","Appel entrant - Temps total",
            "Post-travail - Temps total","Break - Temps total",
            "BUG IT - Temps total","Meeting - Temps total",
            "Pause générique - Temps total","Training - Temps total",
            "Back Office - Temps total","Détachement - Temps total",
            "Call Back - Temps total","Mailing - Temps total","OJT - Temps total"
        ])
        df4 = df4[[c for c in columns_to_keep if c in df4.columns]]

        df4 = df4.fillna("-")
        df4 = df4.rename(columns={"Appel entrant - Occurence": "Appels entrants"})

        temp4 = BytesIO()
        df4.to_excel(temp4, index=False)
        temp4.seek(0)

        wb4 = load_workbook(temp4)
        ws4 = wb4.active
        ws4.title = "Flash Prod Agent"
        headers4 = {cell.value: cell.column for cell in ws4[1]}

        text_cols = ["Matricule RH","Log Téléphonie1","Nom Agent","File","Tls","OPS"]
        time_cols = [
            "Temps Total présence","Temps total Travail","DMC","DMT",
            "Attente global - Temps total","Appel entrant - Temps total",
            "Post-travail - Temps total","Break - Temps total",
            "BUG IT - Temps total","Meeting - Temps total",
            "Pause générique - Temps total","Training - Temps total",
            "Back Office - Temps total","Détachement - Temps total",
            "Call Back - Temps total","Mailing - Temps total","OJT - Temps total"
        ]

        for name in text_cols:
            if name in headers4:
                for r in range(2, ws4.max_row + 1):
                    ws4.cell(r, headers4[name]).number_format = "@"

        for name in time_cols:
            if name in headers4:
                for r in range(2, ws4.max_row + 1):
                    ws4.cell(r, headers4[name]).number_format = "hh:mm:ss"

        # Ajout colonne Taux Post-travail
        if "Post-travail - Temps total" in headers4 and "Temps total Travail" in headers4:
            col_post = headers4["Post-travail - Temps total"]
            ws4.insert_cols(col_post + 1)
            ws4.cell(1, col_post + 1).value = "Taux Post-travail"

        headers4 = {cell.value: cell.column for cell in ws4[1]}

        for r in range(2, ws4.max_row + 1):
            post = to_seconds(ws4.cell(r, headers4.get("Post-travail - Temps total")).value) if "Post-travail - Temps total" in headers4 else None
            work = to_seconds(ws4.cell(r, headers4.get("Temps total Travail")).value) if "Temps total Travail" in headers4 else None
            cell = ws4.cell(r, headers4["Taux Post-travail"])
            if post is None or work in (None, 0):
                cell.value = "-"
            else:
                cell.value = round(post / work, 4)
                cell.number_format = "0.00%"

        # Ajout colonne Moy Post-travail
        if "Taux Post-travail" in headers4:
            col_taux_post = headers4["Taux Post-travail"]
            ws4.insert_cols(col_taux_post + 1)
            ws4.cell(1, col_taux_post + 1).value = "Moy Post-travail"

        headers4 = {cell.value: cell.column for cell in ws4[1]}

        for r in range(2, ws4.max_row + 1):
            post_sec = to_seconds(ws4.cell(r, headers4.get("Post-travail - Temps total")).value) if "Post-travail - Temps total" in headers4 else None
            appels_val = ws4.cell(r, headers4.get("Appels entrants")).value if "Appels entrants" in headers4 else None

            if appels_val in ("-", None, ""):
                appels = None
            else:
                try:
                    appels = float(appels_val)
                except Exception:
                    appels = None

            cell = ws4.cell(r, headers4["Moy Post-travail"])
            if post_sec is None or appels in (None, 0):
                cell.value = "-"
            else:
                avg_sec = int(round(post_sec / appels))
                td = timedelta(seconds=avg_sec)
                cell.value = td
                cell.number_format = "hh:mm:ss"

        # Règle <10 min présence
        if "Temps Total présence" in headers4:
            col_presence = headers4["Temps Total présence"]
            for r in range(2, ws4.max_row + 1):
                sec = to_seconds(ws4.cell(r, col_presence).value)
                if sec is not None and sec < 600:
                    for c in range(col_presence, ws4.max_column + 1):
                        ws4.cell(r, c).value = "-"

        # Style final
        header_fill = PatternFill("solid", "D9E1F2")
        alt_fill = PatternFill("solid", "F7F7F7")
        thin = Side(style="thin")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for c in ws4[1]:
            c.fill = header_fill
            c.font = Font(bold=True)

        for r in range(2, ws4.max_row + 1):
            if r % 2 == 0:
                for c in range(1, ws4.max_column + 1):
                    ws4.cell(r, c).fill = alt_fill

        for row in ws4.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")

        ws4.freeze_panes = "A2"

        for col in ws4.columns:
            vals = [str(cell.value) for cell in col if cell.value not in (None, "")]
            max_len = max([len(v) for v in vals], default=10)
            ws4.column_dimensions[get_column_letter(col[0].column)].width = max_len + 2

        flash_agent_bytes = workbook_to_bytes(wb4)
        log("✅ Code 4 OK")

        # ============================================================
        # ✅ CODE 5 : Synthèse TL + Coaching + Excel final
        # ============================================================
        progress.progress(80, text="Synthèse TL + Excel final (Code 5)...")
        df5 = pd.read_excel(BytesIO(flash_agent_bytes), sheet_name="Flash Prod Agent")

        # conversions en secondes
        for c in ["Temps Total présence", "Temps total Travail", "Post-travail - Temps total", "DMC", "DMT",
                  "Meeting - Temps total", "Training - Temps total", "OJT - Temps total"]:
            if c in df5.columns:
                df5[c] = to_sec_series(df5[c])

        if "Productivité" in df5.columns:
            df5["Productivité"] = pd.to_numeric(df5["Productivité"], errors="coerce")

        for c in ["Meeting - Temps total", "Training - Temps total", "OJT - Temps total"]:
            if c not in df5.columns:
                df5[c] = 0.0

        df5["Temps Coaching (Mtg+Tr+OJT)"] = (
            df5["Meeting - Temps total"].fillna(0) +
            df5["Training - Temps total"].fillna(0) +
            df5["OJT - Temps total"].fillna(0)
        )

        present_mask = df5["Temps Total présence"].fillna(0) >= 600
        df5 = df5[df5["Temps Total présence"].fillna(0) > 0].copy()

        g = df5.groupby("Tls", dropna=False)

        sum_presence = g["Temps Total présence"].sum()
        sum_work = g["Temps total Travail"].sum()
        sum_post = g["Post-travail - Temps total"].sum()

        sum_meeting = g["Meeting - Temps total"].sum()
        sum_training = g["Training - Temps total"].sum()
        sum_ojt = g["OJT - Temps total"].sum()
        sum_coaching = g["Temps Coaching (Mtg+Tr+OJT)"].sum()

        agents_presents = df5[present_mask].groupby("Tls")["Log Téléphonie1"].nunique()
        prod_mask = present_mask & df5["Productivité"].notna() & (df5["Productivité"] > 0)
        agents_lt13 = df5[prod_mask & (df5["Productivité"] < 13)].groupby("Tls")["Log Téléphonie1"].nunique()

        df_tl = pd.DataFrame({
            "Tls": sum_presence.index,
            "Agents présents": agents_presents.reindex(sum_presence.index).fillna(0).astype(int),
            "Productivité": g["Productivité"].mean(),
            "Agents < 13": agents_lt13.reindex(sum_presence.index).fillna(0).astype(int),
            "Taux Post-travail": (sum_post / sum_work).replace([np.inf, -np.inf], np.nan),
            "Taux d'occupation": (sum_work / sum_presence).replace([np.inf, -np.inf], np.nan),
            "DMC": g["DMC"].mean(),
            "DMT": g["DMT"].mean(),
            "Heures Meeting": sum_meeting,
            "Heures Training": sum_training,
            "Heures OJT": sum_ojt,
            "Heures Coaching Total": sum_coaching,
            "% Coaching vs Connecté": (sum_coaching / sum_presence).replace([np.inf, -np.inf], np.nan),
        }).reset_index(drop=True)

        df_tl = df_tl[df_tl["Productivité"].fillna(0) > 0].copy()

        total_row = {
            "Tls": "TOTAL",
            "Agents présents": int(agents_presents.sum()),
            "Productivité": df_tl["Productivité"].mean(),
            "Agents < 13": int(agents_lt13.sum()),
            "Taux Post-travail": (sum_post.sum() / sum_work.sum()) if sum_work.sum() else np.nan,
            "Taux d'occupation": (sum_work.sum() / sum_presence.sum()) if sum_presence.sum() else np.nan,
            "DMC": df_tl["DMC"].mean(),
            "DMT": df_tl["DMT"].mean(),
            "Heures Meeting": sum_meeting.sum(),
            "Heures Training": sum_training.sum(),
            "Heures OJT": sum_ojt.sum(),
            "Heures Coaching Total": sum_coaching.sum(),
            "% Coaching vs Connecté": (sum_coaching.sum() / sum_presence.sum()) if sum_presence.sum() else np.nan,
        }
        df_tl = pd.concat([df_tl, pd.DataFrame([total_row])], ignore_index=True)

        for c in ["DMC", "DMT", "Heures Meeting", "Heures Training", "Heures OJT", "Heures Coaching Total"]:
            df_tl[c] = pd.to_timedelta(df_tl[c], unit="s", errors="coerce")

        wb5 = load_workbook(BytesIO(flash_agent_bytes))
        if "Flash Prod TL" in wb5.sheetnames:
            del wb5["Flash Prod TL"]

        ws5 = wb5.create_sheet("Flash Prod TL")
        ws5.append(df_tl.columns.tolist())
        for _, row in df_tl.iterrows():
            ws5.append(row.tolist())

        green = PatternFill("solid", "C6EFCE")
        red = PatternFill("solid", "F4CCCC")
        header = PatternFill("solid", "D9E1F2")
        total_fill = PatternFill("solid", "FFF2CC")
        thin = Side(style="thin")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        headers5 = {c.value: c.column for c in ws5[1]}
        formats = {
            "Agents présents": "0",
            "Productivité": "0.0",
            "Agents < 13": "0",
            "Taux Post-travail": "0.00%",
            "Taux d'occupation": "0.00%",
            "% Coaching vs Connecté": "0.00%",
            "DMC": "hh:mm:ss",
            "DMT": "hh:mm:ss",
            "Heures Meeting": "hh:mm:ss",
            "Heures Training": "hh:mm:ss",
            "Heures OJT": "hh:mm:ss",
            "Heures Coaching Total": "hh:mm:ss",
        }

        for k, fmt in formats.items():
            if k in headers5:
                col = headers5[k]
                for r in range(2, ws5.max_row + 1):
                    ws5.cell(r, col).number_format = fmt

        for c in ws5[1]:
            c.fill = header
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center")

        for r in range(2, ws5.max_row + 1):
            for c in range(1, ws5.max_column + 1):
                ws5.cell(r, c).border = border
                ws5.cell(r, c).alignment = Alignment(horizontal="center")

        last_row = ws5.max_row

        # Conditional formatting : Productivité
        if "Productivité" in headers5:
            colL = get_column_letter(headers5["Productivité"])
            ws5.conditional_formatting.add(
                f"{colL}2:{colL}{last_row-1}",
                FormulaRule(formula=[f"{colL}2>13"], fill=green)
            )
            ws5.conditional_formatting.add(
                f"{colL}2:{colL}{last_row-1}",
                FormulaRule(formula=[f"{colL}2<=13"], fill=red)
            )

        # Conditional formatting : Taux Post-travail
        if "Taux Post-travail" in headers5:
            colL = get_column_letter(headers5["Taux Post-travail"])
            ws5.conditional_formatting.add(
                f"{colL}2:{colL}{last_row-1}",
                FormulaRule(formula=[f"{colL}2<=0.08"], fill=green)
            )
            ws5.conditional_formatting.add(
                f"{colL}2:{colL}{last_row-1}",
                FormulaRule(formula=[f"{colL}2>0.08"], fill=red)
            )

        # DMC & DMT
        for col_name in ["DMC", "DMT"]:
            if col_name in headers5:
                col_letter = get_column_letter(headers5[col_name])
                ws5.conditional_formatting.add(
                    f"{col_letter}2:{col_letter}{last_row-1}",
                    FormulaRule(formula=[f"{col_letter}2<=TIME(0,3,0)"], fill=green)
                )
                ws5.conditional_formatting.add(
                    f"{col_letter}2:{col_letter}{last_row-1}",
                    FormulaRule(formula=[f"{col_letter}2>TIME(0,3,0)"], fill=red)
                )

        # Taux d'occupation
        if "Taux d'occupation" in headers5:
            colL = get_column_letter(headers5["Taux d'occupation"])
            ws5.conditional_formatting.add(
                f"{colL}2:{colL}{last_row-1}",
                FormulaRule(formula=[f"{colL}2>=0.7"], fill=green)
            )
            ws5.conditional_formatting.add(
                f"{colL}2:{colL}{last_row-1}",
                FormulaRule(formula=[f"{colL}2<0.7"], fill=red)
            )

        # Ligne TOTAL
        for c in range(1, ws5.max_column + 1):
            ws5.cell(ws5.max_row, c).fill = total_fill
            ws5.cell(ws5.max_row, c).font = Font(bold=True)

        for col in ws5.columns:
            ws5.column_dimensions[get_column_letter(col[0].column)].width = 18

        ws5.freeze_panes = "A2"

        excel_filename = f"Flash_Prod_{projet_upper}_{date_flash}.xlsx"
        excel_bytes = workbook_to_bytes(wb5)
        log("✅ Code 5 OK")

        # ============================================================
        # ✅ CODE 6 : Email (objet + HTML) + colonnes Agents + Coaching
        # ============================================================
        progress.progress(95, text="Email HTML (Code 6)...")
        df_tl_email = pd.read_excel(BytesIO(excel_bytes), sheet_name="Flash Prod TL")
        df_total = df_tl_email[df_tl_email["Tls"] == "TOTAL"].iloc[0]
        df_tl_email = df_tl_email[df_tl_email["Tls"] != "TOTAL"].copy()

        kpi_prod = round(float(df_total["Productivité"]), 1)
        kpi_occ = round(float(df_total["Taux d'occupation"]) * 100, 1)
        kpi_post = round(float(df_total["Taux Post-travail"]) * 100, 1)
        kpi_coach = round(float(df_total["% Coaching vs Connecté"]) * 100, 1) if pd.notna(df_total["% Coaching vs Connecté"]) else 0.0

        kpi_dmc_sec = time_to_seconds(df_total["DMC"])
        kpi_dmt_sec = time_to_seconds(df_total["DMT"])
        kpi_dmc_txt = sec_to_hms(kpi_dmc_sec)
        kpi_dmt_txt = sec_to_hms(kpi_dmt_sec)

        email_subject = f"{projet_upper} ==> Flash Prod AE de la journée du {date_txt} - Taux d'occupation {kpi_occ} %"

        def color_prod(v):
            if v > 13: return "#E8F5E9", "#1E8449"
            if v >= 12: return "#FEF9E7", "#7D6608"
            return "#FDEDEC", "#922B21"

        def color_post(v):
            if v <= 0.06: return "#E8F5E9", "#1E8449"
            if v <= 0.08: return "#FEF9E7", "#7D6608"
            return "#FDEDEC", "#922B21"

        def color_time(sec):
            if sec is None: return "#FFFFFF", "#111111"
            if sec <= 180: return "#E8F5E9", "#1E8449"
            if sec <= 210: return "#FEF9E7", "#7D6608"
            return "#FDEDEC", "#922B21"

        def color_occ(v):
            if v >= 0.7: return "#E8F5E9", "#1E8449"
            if v >= 0.65: return "#FEF9E7", "#7D6608"
            return "#FDEDEC", "#922B21"

        band_html = f"""
        <table style="width:100%;margin-bottom:25px;border-radius:10px;
        background:#F9FAFB;text-align:center;font-size:13px;">
        <tr style="font-weight:bold;color:#203864;">
        <td>📈 Productivité<br><span style="font-size:18px;">{kpi_prod}</span></td>
        <td>⏱️ DMC<br><span style="font-size:18px;">{kpi_dmc_txt}</span></td>
        <td>⏱️ DMT<br><span style="font-size:18px;">{kpi_dmt_txt}</span></td>
        <td>🧩 Taux Post-travail<br><span style="font-size:18px;">{kpi_post}%</span></td>
        <td>🎓 Coaching vs Connecté<br><span style="font-size:18px;">{kpi_coach}%</span></td>
        <td style="background:#E8F5E9;border-radius:8px;color:#1E8449;">
        🎯 Taux d’occupation<br>
        <span style="font-size:20px;font-weight:bold;">{kpi_occ}%</span>
        </td>
        </tr>
        </table>
        """

        table_html = """
        <table style="border-collapse:collapse;width:100%;font-size:13px;text-align:center;">
        <thead>
        <tr style="background:#203864;color:white;">
        <th>TL</th>
        <th>Agents présents</th>
        <th>Productivité</th>
        <th>Agents &lt; 13</th>
        <th>Taux Post-travail</th>
        <th>% Coaching</th>
        <th>DMC</th>
        <th>DMT</th>
        <th>Taux d’occupation</th>
        </tr>
        </thead><tbody>
        """

        for i, r in df_tl_email.iterrows():
            bg = "#F9FAFB" if i % 2 == 0 else "#FFFFFF"

            p_bg, p_cl = color_prod(float(r["Productivité"])) if pd.notna(r["Productivité"]) else ("#FFFFFF", "#111")
            pt_bg, pt_cl = color_post(float(r["Taux Post-travail"])) if pd.notna(r["Taux Post-travail"]) else ("#FFFFFF", "#111")

            dmc_sec = time_to_seconds(r["DMC"])
            dmt_sec = time_to_seconds(r["DMT"])

            dmc_bg, dmc_cl = color_time(dmc_sec)
            dmt_bg, dmt_cl = color_time(dmt_sec)

            occ_bg, occ_cl = color_occ(float(r["Taux d'occupation"])) if pd.notna(r["Taux d'occupation"]) else ("#FFFFFF", "#111")

            coach_pct = round(float(r["% Coaching vs Connecté"]) * 100, 1) if pd.notna(r["% Coaching vs Connecté"]) else 0.0

            table_html += f"""
            <tr style="background:{bg};">
              <td>{r['Tls']}</td>
              <td><b>{int(r['Agents présents']) if pd.notna(r['Agents présents']) else 0}</b></td>
              <td style="background:{p_bg};color:{p_cl};font-weight:bold;">{round(float(r['Productivité']),1) if pd.notna(r['Productivité']) else '-'}</td>
              <td><b>{int(r['Agents < 13']) if pd.notna(r['Agents < 13']) else 0}</b></td>
              <td style="background:{pt_bg};color:{pt_cl};font-weight:bold;">{round(float(r['Taux Post-travail'])*100,1) if pd.notna(r['Taux Post-travail']) else '-'}%</td>
              <td style="font-weight:bold;">{coach_pct}%</td>
              <td style="background:{dmc_bg};color:{dmc_cl};font-weight:bold;">{sec_to_hms(dmc_sec)}</td>
              <td style="background:{dmt_bg};color:{dmt_cl};font-weight:bold;">{sec_to_hms(dmt_sec)}</td>
              <td style="background:{occ_bg};color:{occ_cl};font-weight:bold;">{round(float(r["Taux d'occupation"])*100,1) if pd.notna(r["Taux d'occupation"]) else '-'}%</td>
            </tr>
            """

        table_html += "</tbody></table>"

        email_html = f"""
        <div style="background:#F3F6FB;padding:30px;font-family:Calibri,Arial;">
        <div style="max-width:950px;margin:auto;background:#FFFFFF;border-radius:14px;padding:26px;border:1px solid #E0E6ED;">

        <div style="border-left:5px solid #203864;padding-left:14px;margin-bottom:20px;">
        <div style="font-size:20px;font-weight:bold;color:#203864;">📊 Flash Production AE – {projet_upper}</div>
        <div style="font-size:13px;color:#6B7280;">Données du {date_txt}</div>
        </div>

        {band_html}

        <p>Bonjour,</p>
        <p>
        Vous trouverez ci-après les <b>réalisations KPI / Productivité par équipe AE</b>
        du <b>{date_txt}</b>, ainsi que le <b>détail par agent</b> en pièce jointe.
        </p>

        <p><b>➡️ Synthèse Productivité / équipe AE :</b></p>
        {table_html}

        <p style="margin-top:20px;">
        📎 <b>Pièce jointe :</b> Flash Prod AE – Détail Agents & Synthèse TL
        </p>

        <p style="margin-top:20px;">
        Cordialement,<br><br>
        <b style="color:#1F4E78;">Yassine MAHAMID</b><br>
        Analyste IDP – Workforce Management & Reporting
        </p>

        </div></div>
        """

        progress.progress(100, text="Terminé ✅")
        log("✅ Code 6 OK")

        # ============================================================
        # ✅ UI OUTPUTS
        # ============================================================
        st.success("Pipeline terminé ✅")

        col1, col2 = st.columns([1, 1])

        with col1:
            st.subheader("📌 Excel final")
            st.download_button(
                label="⬇️ Télécharger l'Excel",
                data=excel_bytes,
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with col2:
            st.subheader("✉️ Objet Email")
            st.code(email_subject)

        st.subheader("🧾 Corps Email (HTML)")
        st.markdown(email_html, unsafe_allow_html=True)

        st.divider()
        st.subheader("⬇️ Export Email")

        html_filename = f"Email_FlashProd_{projet_upper}_{date_flash}.html"
        st.download_button(
            label="Télécharger le HTML",
            data=email_html.encode("utf-8"),
            file_name=html_filename,
            mime="text/html",
            use_container_width=True,
        )

        # EML simple (ça marche souvent dans Outlook/Thunderbird, mais dépend des clients)
        eml_content = (
            f"Subject: {email_subject}\n"
            f"MIME-Version: 1.0\n"
            f"Content-Type: text/html; charset=utf-8\n\n"
            f"{email_html}"
        )
        eml_filename = f"Email_FlashProd_{projet_upper}_{date_flash}.eml"
        st.download_button(
            label="Télécharger en .eml",
            data=eml_content.encode("utf-8"),
            file_name=eml_filename,
            mime="message/rfc822",
            use_container_width=True,
        )

        # Option: aperçu synthèse TL
        with st.expander("📋 Aperçu table Flash Prod TL"):
            st.dataframe(df_tl_email, use_container_width=True)

    except Exception as e:
        progress.empty()
        st.error("Erreur pendant le pipeline.")
        st.exception(e)
