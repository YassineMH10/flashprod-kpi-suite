import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime, time, timedelta

from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


# =========================
# Streamlit config
# =========================
st.set_page_config(page_title="FlashProd KPI Suite", layout="wide")


# =========================
# Helpers (robustes)
# =========================
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
        h, m, s = val.split(":")
        return int(h) * 3600 + int(m) * 60 + int(s)
    return None

def sec_to_hms(sec):
    if sec is None:
        return "-"
    return str(timedelta(seconds=int(sec)))

def to_hms(val):
    if isinstance(val, str):
        match = re.match(r"(\d+)h(\d+)'(\d+)", val)
        if match:
            h, m, s = match.groups()
            return f"{int(h):02}:{int(m):02}:{int(s):02}"
    return val


# =========================
# Read Excel (xls/xlsx)
# =========================
def read_excel_any(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    bio = BytesIO(uploaded_file.getvalue())
    if name.endswith(".xls"):
        return pd.read_excel(bio, engine="xlrd")
    return pd.read_excel(bio, engine="openpyxl")


# =========================
# CODE 1: Nettoyage brut -> rapport_nettoye.xlsx (in-memory)
# =========================
def code1_clean_raw(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")

    if "Unnamed: 0" in df.columns:
        df["Unnamed: 0"] = df["Unnamed: 0"].ffill()

    df = df.rename(columns={
        "Unnamed: 0": "Nom Agent",
        "Unnamed: 1": "Etat",
        "Unnamed: 4": "Occurances",
        "Unnamed: 6": "Temps total"
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

    if "Temps total" in df.columns:
        df["Temps total"] = df["Temps total"].apply(to_hms)

    if "Etat" in df.columns:
        df["Etat"] = df["Etat"].replace({
            "Attente": "Attente global",
            "Pause": "Pause global",
            "Preview": "Histo Mailing"
        })
        df = df[df["Etat"].astype(str).str.lower() != "en attente"]

    return df


# =========================
# CODE 2: Pivot + KPI Agent -> df_rapport
# =========================
def code2_build_agent_kpis(df_clean: pd.DataFrame) -> pd.DataFrame:
    df2 = df_clean.copy()

    if "Log Téléphonie1" in df2.columns:
        df2 = df2.drop(columns=["Log Téléphonie1"])

    df2.insert(
        0,
        "Log Téléphonie1",
        df2["Nom Agent"].apply(lambda x: re.search(r"Agent\s+(\d{4})", str(x)).group(1)
                               if re.search(r"Agent\s+(\d{4})", str(x)) else None)
    )

    unique_etats = df2["Etat"].dropna().unique()
    agent_base = df2[["Log Téléphonie1"]].drop_duplicates().reset_index(drop=True)

    for etat in unique_etats:
        agent_base[f"{etat} - Occurence"] = "-"
        agent_base[f"{etat} - Temps total"] = "-"

    for i, row in agent_base.iterrows():
        log = row["Log Téléphonie1"]
        sub_df = df2[df2["Log Téléphonie1"] == log]
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
            col = f"{e} - Temps total"
            if col in r:
                total += to_td(r[col])
        return fmt_hms(total)

    agent_base["Temps Total présence"] = agent_base.apply(total_presence_row, axis=1)

    def total_travail_row(r):
        return fmt_hms(to_td(r.get("Traitement - Temps total", "-")) + to_td(r.get("Post-travail - Temps total", "-")))

    agent_base.insert(
        agent_base.columns.get_loc("Temps Total présence") + 1,
        "Temps total Travail",
        agent_base.apply(total_travail_row, axis=1)
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
        agent_base.apply(taux_occupation_row, axis=1)
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
        total = (to_td(r.get("Traitement - Temps total", "00:00:00")) +
                 to_td(r.get("Post-travail - Temps total", "00:00:00"))) / occ
        return fmt_hms(total)

    agent_base["DMT"] = agent_base.apply(calc_dmt, axis=1)

    return agent_base


# =========================
# CODE 3: Merge COMPO + Rapport -> df_final3
# =========================
def code3_merge_compo(df_compo: pd.DataFrame, df_rapport: pd.DataFrame) -> pd.DataFrame:
    df_merged = pd.merge(df_compo, df_rapport, on="Log Téléphonie1", how="left")
    df_merged = df_merged[df_merged["Nom Agent"].notna() & (df_merged["Nom Agent"] != "")]
    if "Matricule" in df_merged.columns:
        df_merged = df_merged.drop(columns=["Matricule"])

    colonnes_apres_ops = [
        "Temps Total présence", "Temps total Travail", "Taux d'occupation",
        "Productivité", "DMC", "DMT"
    ]

    ops_index = df_merged.columns.get_loc("OPS") + 1 if "OPS" in df_merged.columns else 0
    colonnes_sans = [c for c in df_merged.columns if c not in colonnes_apres_ops]

    nouvel_ordre = (
        colonnes_sans[:ops_index] +
        [c for c in colonnes_apres_ops if c in df_merged.columns] +
        colonnes_sans[ops_index:]
    )

    return df_merged[nouvel_ordre]


# =========================
# CODE 4: Flash Prod Agent final + Taux Post + Moy Post + règle <10min
# =========================
def code4_build_flash_agent_excel(df_final3: pd.DataFrame) -> bytes:
    df4 = df_final3.copy()

    columns_to_keep = [
        "Matricule RH","Log Téléphonie1","Nom Agent","File","Tls","OPS",
        "Temps Total présence","Temps total Travail","Taux d'occupation",
        "Productivité","DMC","DMT","Attente global - Temps total",
        "Appel entrant - Occurence","Appel entrant - Temps total",
        "Post-travail - Temps total","Break - Temps total",
        "BUG IT - Temps total","Meeting - Temps total",
        "Pause générique - Temps total","Training - Temps total",
        "Back Office - Temps total","Détachement - Temps total",
        "Call Back - Temps total","Mailing - Temps total","OJT - Temps total","BUG IT - Temps total"
    ]
    df4 = df4[[c for c in columns_to_keep if c in df4.columns]]
    df4 = df4.fillna("-")
    df4 = df4.rename(columns={"Appel entrant - Occurence": "Appels entrants"})

    # write temp workbook to memory
    temp_buf = BytesIO()
    df4.to_excel(temp_buf, index=False, engine="openpyxl")
    temp_buf.seek(0)

    wb4 = load_workbook(temp_buf)
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
        "Call Back - Temps total","Mailing - Temps total","OJT - Temps total","BUG IT - Temps total"
    ]

    for name in text_cols:
        if name in headers4:
            for r in range(2, ws4.max_row + 1):
                ws4.cell(r, headers4[name]).number_format = "@"

    for name in time_cols:
        if name in headers4:
            for r in range(2, ws4.max_row + 1):
                ws4.cell(r, headers4[name]).number_format = "hh:mm:ss"

    # Add "Taux Post-travail"
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

    # Add "Moy Post-travail"
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

    # Rule < 10 min presence: set "-" from presence col to end
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

    out = BytesIO()
    wb4.save(out)
    out.seek(0)
    return out.getvalue()


# =========================
# CODE 5: Synthèse TL + conditional formatting -> Excel final
# =========================
def code5_add_tl_sheet(excel_agent_bytes: bytes) -> bytes:
    wb5 = load_workbook(BytesIO(excel_agent_bytes))
    df5 = pd.read_excel(BytesIO(excel_agent_bytes), sheet_name="Flash Prod Agent", engine="openpyxl")

    for c in ["Temps Total présence", "Temps total Travail", "Post-travail - Temps total", "DMC", "DMT"]:
        if c in df5.columns:
            df5[c] = pd.to_timedelta(df5[c], errors="coerce").dt.total_seconds()

    if "Productivité" in df5.columns:
        df5["Productivité"] = pd.to_numeric(df5["Productivité"], errors="coerce")

    df5["Taux Post-travail"] = df5["Post-travail - Temps total"] / df5["Temps total Travail"]
    df5["Taux d'occupation"] = df5["Temps total Travail"] / df5["Temps Total présence"]

    df_tl = (
        df5.groupby("Tls")[["Productivité", "Taux Post-travail", "DMC", "DMT", "Taux d'occupation"]]
        .mean()
        .reset_index()
    )

    df_tl = df_tl[df_tl["Productivité"] > 0]

    total = pd.DataFrame([["TOTAL"] + df_tl.iloc[:, 1:].mean().tolist()], columns=df_tl.columns)
    df_tl = pd.concat([df_tl, total], ignore_index=True)

    df_tl["DMC"] = pd.to_timedelta(df_tl["DMC"], unit="s")
    df_tl["DMT"] = pd.to_timedelta(df_tl["DMT"], unit="s")

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
        "Productivité": "0.0",
        "Taux Post-travail": "0.00%",
        "Taux d'occupation": "0.00%",
        "DMC": "hh:mm:ss",
        "DMT": "hh:mm:ss"
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

    def add_rule(col_name, formula_good, formula_bad):
        if col_name not in headers5:
            return
        colL = get_column_letter(headers5[col_name])
        ws5.conditional_formatting.add(
            f"{colL}2:{colL}{last_row-1}",
            FormulaRule(formula=[formula_good.replace("{COL}", colL) + "2"], fill=green)
        )
        ws5.conditional_formatting.add(
            f"{colL}2:{colL}{last_row-1}",
            FormulaRule(formula=[formula_bad.replace("{COL}", colL) + "2"], fill=red)
        )

    # Productivité > 13 green else red
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

    # Taux Post <=0.08 green else red
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

    # DMC/DMT <= 03:00 green else red
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

    # Occupation >=0.7 green else red
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

    # Ligne TOTAL style
    for c in range(1, ws5.max_column + 1):
        ws5.cell(ws5.max_row, c).fill = total_fill
        ws5.cell(ws5.max_row, c).font = Font(bold=True)

    for col in ws5.columns:
        ws5.column_dimensions[get_column_letter(col[0].column)].width = 18
    ws5.freeze_panes = "A2"

    out = BytesIO()
    wb5.save(out)
    out.seek(0)
    return out.getvalue()


# =========================
# CODE 6: Email (objet + HTML)
# =========================
def code6_build_email(excel_final_bytes: bytes, projet_upper: str, date_txt: str):
    df_tl_email = pd.read_excel(BytesIO(excel_final_bytes), sheet_name="Flash Prod TL", engine="openpyxl")
    df_total = df_tl_email[df_tl_email["Tls"] == "TOTAL"].iloc[0]
    df_tl_email = df_tl_email[df_tl_email["Tls"] != "TOTAL"].copy()

    kpi_prod = round(float(df_total["Productivité"]), 1)
    kpi_occ = round(float(df_total["Taux d'occupation"]) * 100, 1)
    kpi_post = round(float(df_total["Taux Post-travail"]) * 100, 1)

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
        if sec is None:
            return "#F7F7F7", "#111827"
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
    <th>Productivité</th>
    <th>Taux Post-travail</th>
    <th>DMC</th>
    <th>DMT</th>
    <th>Taux d’occupation</th>
    </tr>
    </thead><tbody>
    """

    for i, r in df_tl_email.iterrows():
        bg = "#F9FAFB" if i % 2 == 0 else "#FFFFFF"

        p_bg, p_cl = color_prod(float(r["Productivité"]))
        pt_bg, pt_cl = color_post(float(r["Taux Post-travail"]))

        dmc_sec = time_to_seconds(r["DMC"])
        dmt_sec = time_to_seconds(r["DMT"])

        dmc_bg, dmc_cl = color_time(dmc_sec)
        dmt_bg, dmt_cl = color_time(dmt_sec)

        occ_bg, occ_cl = color_occ(float(r["Taux d'occupation"]))

        table_html += f"""
        <tr style="background:{bg};">
          <td>{r['Tls']}</td>
          <td style="background:{p_bg};color:{p_cl};font-weight:bold;">{round(float(r['Productivité']),1)}</td>
          <td style="background:{pt_bg};color:{pt_cl};font-weight:bold;">{round(float(r['Taux Post-travail'])*100,1)}%</td>
          <td style="background:{dmc_bg};color:{dmc_cl};font-weight:bold;">{sec_to_hms(dmc_sec)}</td>
          <td style="background:{dmt_bg};color:{dmt_cl};font-weight:bold;">{sec_to_hms(dmt_sec)}</td>
          <td style="background:{occ_bg};color:{occ_cl};font-weight:bold;">{round(float(r["Taux d'occupation"])*100,1)}%</td>
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
    📎 <b>Pièce jointe :</b> Flash Prod AE – Détail Agents &amp; Synthèse TL
    </p>

    <p style="margin-top:20px;">
    Cordialement,<br><br>
    <b style="color:#1F4E78;">MAHAMID Yassine</b><br>
    FlashProd KPI Suite<br>
    <span style="color:#6B7280;font-size:12px;">Developed by MAHAMID Yassine</span>
    </p>

    </div></div>
    """

    return email_subject, email_html


def copy_buttons(subject: str, html_body: str):
    component = f"""
    <div style="font-family:Calibri, Arial; margin-top:10px;">
      <div style="margin-bottom:6px; font-weight:bold;">📋 Copier vers Gmail :</div>

      <textarea id="sbj" rows="2" style="width:100%;">{subject}</textarea>
      <button style="margin-top:6px; padding:8px 14px; border:0; border-radius:6px; cursor:pointer; background:#28A745; color:white;"
              onclick="navigator.clipboard.writeText(document.getElementById('sbj').value)">
        ✅ Copier l'objet
      </button>

      <div style="height:10px;"></div>

      <div id="htmlBody" style="display:none;">{html_body}</div>
      <button style="padding:8px 14px; border:0; border-radius:6px; cursor:pointer; background:#0078D7; color:white;"
              onclick="navigator.clipboard.write([
                new ClipboardItem({{
                  'text/html': new Blob([document.getElementById('htmlBody').innerHTML], {{type:'text/html'}})
                }})
              ])">
        ✅ Copier le corps HTML (coller formaté dans Gmail)
      </button>
    </div>
    """
    st.components.v1.html(component, height=210, scrolling=False)


# =========================
# UI
# =========================
st.title("📊 FlashProd KPI Suite")
st.caption("Pipeline BRUT + COMPO → KPI Agents + Synthèse TL → Excel final + Email HTML")

with st.sidebar:
    st.header("Inputs")
    projet = st.text_input("Nom du projet (ex: CNSS)", value="CNSS")
    date_input = st.text_input("Date Flash Prod (JJ/MM/AAAA)", value=datetime.today().strftime("%d/%m/%Y"))

    st.markdown("---")
    raw_file = st.file_uploader("📁 Fichier BRUT (.xls / .xlsx)", type=["xls", "xlsx"])
    compo_file = st.file_uploader("📁 Fichier COMPO (.xls / .xlsx)", type=["xls", "xlsx"])

    st.markdown("---")
    run = st.button("🚀 Lancer le pipeline", type="primary")

if not run:
    st.info("Uploade BRUT + COMPO puis clique **Lancer le pipeline**.")
    st.markdown("<div style='text-align:center;color:#6c757d;font-size:12px;margin-top:16px;'>Developed by <b>MAHAMID Yassine</b></div>", unsafe_allow_html=True)
    st.stop()

if not raw_file or not compo_file:
    st.error("BRUT + COMPO sont obligatoires.")
    st.stop()

# Validate date
try:
    date_obj = datetime.strptime(date_input.strip(), "%d/%m/%Y")
except Exception:
    st.error("Format date invalide. Utilise JJ/MM/AAAA (ex: 09/02/2026).")
    st.stop()

projet_upper = projet.strip().upper()
date_txt = date_obj.strftime("%d/%m/%Y")
date_flash = date_obj.strftime("%d_%m_%Y")

# =========================
# Pipeline
# =========================
status_box = st.empty()
logs = []

def log_ok(msg):
    logs.append(msg)
    status_box.success("\n".join(logs))

try:
    # Read inputs
    df_raw = read_excel_any(raw_file)
    df_compo = read_excel_any(compo_file)

    # CODE 1
    df_clean = code1_clean_raw(df_raw)
    log_ok("✅ CODE 1 OK -> Nettoyage BRUT terminé")

    # CODE 2
    df_rapport = code2_build_agent_kpis(df_clean)
    log_ok("✅ CODE 2 OK -> KPI Agent / Pivot terminé")

    # CODE 3
    df_final3 = code3_merge_compo(df_compo, df_rapport)
    log_ok("✅ CODE 3 OK -> Merge COMPO + Rapport terminé")

    # CODE 4
    agent_excel_bytes = code4_build_flash_agent_excel(df_final3)
    log_ok("✅ CODE 4 OK -> Flash Prod Agent (Excel) généré")

    # CODE 5
    excel_final_bytes = code5_add_tl_sheet(agent_excel_bytes)
    excel_final_name = f"Flash_Prod_{projet_upper}_{date_flash}.xlsx"
    log_ok(f"✅ CODE 5 OK -> Excel final prêt ({excel_final_name})")

    # CODE 6
    email_subject, email_html = code6_build_email(excel_final_bytes, projet_upper, date_txt)
    log_ok("✅ CODE 6 OK -> Email (Objet + Corps HTML) généré")

except Exception as e:
    st.error(f"Erreur pipeline: {e}")
    st.stop()

# =========================
# Outputs
# =========================
col1, col2 = st.columns([1.1, 0.9], gap="large")

with col1:
    st.subheader("✉️ Email")
    st.text_input("Objet", value=email_subject)
    st.components.v1.html(email_html, height=560, scrolling=True)
    copy_buttons(email_subject, email_html)

with col2:
    st.subheader("⬇️ Excel final")
    st.download_button(
        label="Télécharger l'Excel final",
        data=excel_final_bytes,
        file_name=excel_final_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.subheader("👀 Aperçu (Top 50) – Flash Prod TL")
    try:
        df_tl_preview = pd.read_excel(BytesIO(excel_final_bytes), sheet_name="Flash Prod TL", engine="openpyxl")
        st.dataframe(df_tl_preview.head(50), use_container_width=True, height=420)
    except Exception as e:
        st.warning(f"Aperçu TL non disponible: {e}")

st.markdown("<div style='text-align:center;color:#6c757d;font-size:12px;margin-top:16px;'>Developed by <b>MAHAMID Yassine</b></div>", unsafe_allow_html=True)
