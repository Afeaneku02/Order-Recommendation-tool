# -*- coding: utf-8 -*-
"""
Spyder Editor

Author: Winfred Afeaneku

Started August 17, 2025.
"""


import re
from io import BytesIO
from datetime import datetime
from typing import Optional, Tuple, List, Dict

import pandas as pd
import streamlit as st
import openpyxl


# ==============================
# Links (MUST be strings)
# ==============================

DOWNLOAD_LINKS = {
    "Global Backorder": "https://app.powerbi.com/groups/me/reports/89f5947c-abf5-4cd4-95b4-25a158409d8d/ReportSection54962abf70e80976410c?ctid=39b03722-b836-496a-85ec-850f0957ca6b&experience=power-bi&bookmarkGuid=d4ceed12-f5a9-4ea3-a6f6-aac101929d19",
    "MG4 Tool": "https://app.powerbi.com/groups/me/apps/9c534bcc-8b34-4baf-a339-a346bf2d8ae8/reports/b428d9db-2c72-467d-b3e5-a66102dc42e7/ReportSection64b84a45b66b79870377?ctid=39b03722-b836-496a-85ec-850f0957ca6b&experience=power-bi",
    "SPM Search by Material": "https://app.powerbi.com/groups/me/apps/ba63bdfd-30f3-4175-9b04-2cd1054f2e90/reports/ca29414c-71df-4a11-8c39-1d8ed34f3404/ReportSection7a99ad4fc84437ca260d?ctid=39b03722-b836-496a-85ec-850f0957ca6b&experience=power-bi&bookmarkGuid=Bookmarkda7c9c30b8d03e9d7d23",
    "Forecast Export (spm_search_by_mtl_fct_prts)": "https://app.powerbi.com/groups/me/apps/ba63bdfd-30f3-4175-9b04-2cd1054f2e90/reports/ca29414c-71df-4a11-8c39-1d8ed34f3404/ReportSection87aebeb3cda703148f17?ctid=39b03722-b836-496a-85ec-850f0957ca6b&experience=power-bi&bookmarkGuid=Bookmarkda7c9c30b8d03e9d7d23",
}


# ==============================
# Streamlit cache helpers
# ==============================

@st.cache_data(show_spinner=False)
def read_table_cached(file, sheet_name=0, **kwargs) -> pd.DataFrame:
    """
    Reads either Excel or CSV from a Streamlit UploadedFile.
    - Excel: pd.read_excel (defaults to first sheet)
    - CSV: pd.read_csv
    """
    name = (getattr(file, "name", "") or "").lower()

    if name.endswith(".csv"):
        return pd.read_csv(file, **kwargs)

    # Excel: default FIRST sheet so we always return a DataFrame (not dict)
    return pd.read_excel(file, sheet_name=sheet_name)


# ==============================
# Column normalization / alias matching
# ==============================

@st.cache_data(show_spinner=False)
def forecast_avg_cached(fcst_export_df: pd.DataFrame) -> pd.DataFrame:
    return spm_search_by_mtl_fct_prts_Avg(fcst_export_df)


def _norm_col(c: str) -> str:
    c = str(c).strip().upper()
    c = re.sub(r"\s+", " ", c)
    c = re.sub(r"[^A-Z0-9]", "", c)
    return c


def find_col(df: pd.DataFrame, candidates: List[str], required: bool = True) -> Optional[str]:
    if df is None or df.empty:
        if required:
            raise ValueError("Input dataframe is empty.")
        return None

    norm_map = {_norm_col(c): c for c in df.columns}
    for cand in candidates:
        key = _norm_col(cand)
        if key in norm_map:
            return norm_map[key]

    if required:
        raise ValueError(f"Missing required column. Tried: {candidates}. Found: {list(df.columns)}")
    return None


def rename_to_canonical(df: pd.DataFrame, canonical_map: Dict[str, List[str]]) -> pd.DataFrame:
    out = df.copy()
    rename_dict = {}
    for canon, candidates in canonical_map.items():
        actual = find_col(out, candidates, required=True)
        rename_dict[actual] = canon
    out.rename(columns=rename_dict, inplace=True)
    return out


def normalize_keys(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c in out.columns:
            out[c] = out[c].astype(str).str.strip()
    return out


def validate_and_prepare_inputs(
    bo_df: pd.DataFrame,
    mg4_df: Optional[pd.DataFrame],
    spm_df: pd.DataFrame,
    fcst_export_df: pd.DataFrame,
    require_mg4: bool = True,
) -> Tuple[pd.DataFrame, Optional[pd.DataFrame], pd.DataFrame, pd.DataFrame]:
    # ---- BO canonical
    bo_df = rename_to_canonical(bo_df, {
        "PART NBR": ["PART NBR", "PART", "PARTNUMBER", "PART NUMBER"],
        "SHIP UNIT": ["SHIP UNIT", "SHIPUNIT", "DEPOT", "DIST CTR", "DIST_CTR"],
        "BO QTY": ["BO QTY", "BO", "BACKORDER", "BACK ORDER QTY", "BOQTY"],
        "ACCT UNIT": ["ACCT UNIT", "ACCOUNT UNIT", "ACCTUNIT", "ACCT"],
    })

    inv_actual = find_col(bo_df, ["INV CLASS", "INVCLASS", "INVENTORY CLASS"], required=False)
    if inv_actual and inv_actual != "INV CLASS":
        bo_df.rename(columns={inv_actual: "INV CLASS"}, inplace=True)

    bo_df = normalize_keys(bo_df, ["PART NBR", "SHIP UNIT"])

    # ---- MG4 canonical (ONLY if required)
    if require_mg4:
        if mg4_df is None or mg4_df.empty:
            raise ValueError("MG4 file is required for Factory Direct (R) or Vendor Direct (V).")

        mg4_df = rename_to_canonical(mg4_df, {
            "MATERIAL": ["MATERIAL", "material", "MATL", "MATL NUM", "MATL NUMBER"],
            "MG4_FLAG": [
                "PAG/ PDX MG 4", "PAG/PDX MG 4", "PAG PDX MG 4",
                "PAG/ PDX MG4", "PAG/PDX MG4",
                "MATERIAL GROUP 4", "Material group 4", "MG4", "MG 4"
            ],
        })
        mg4_df = normalize_keys(mg4_df, ["MATERIAL", "MG4_FLAG"])
        mg4_df["MG4_FLAG"] = mg4_df["MG4_FLAG"].astype(str).str.strip().str.upper()
    else:
        mg4_df = None  # explicitly not used

    # ---- SPM canonical
    spm_df = rename_to_canonical(spm_df, {
        "PRT NUM": ["PRT NUM", "PART", "PART NBR", "PART NUMBER"],
        "DIST_CTR": ["DIST_CTR", "DIST CTR", "DEPOT", "SHIP UNIT"],
    })
    for col in ["AVAIL", "INHOUSE", "WIP", "INTRANSIT", "ON_ORDER"]:
        actual = find_col(spm_df, [col, col.replace("_", " ")], required=False)
        if actual and actual != col:
            spm_df.rename(columns={actual: col}, inplace=True)

    spm_df = normalize_keys(spm_df, ["PRT NUM", "DIST_CTR"])

    # ---- Forecast canonical
    fcst_export_df = rename_to_canonical(fcst_export_df, {
        "Month": ["Month", "MONTH"],
        "PRT NUM": ["PRT NUM", "PART", "PART NBR", "PART NUMBER"],
        "Value": ["Value", "VALUE", "VAL"],
    })
    fcst_export_df = normalize_keys(fcst_export_df, ["PRT NUM", "Month"])

    return bo_df, mg4_df, spm_df, fcst_export_df


# ==============================
# Class-based recommendation factors
# ==============================

CLASS_COLUMN = "INV CLASS"
CLASS_REC_FACTORS = {"A": 0.90, "B": 0.80, "C": 0.75, "D": 0.70, "E": 0.70}
DEFAULT_REC_FACTOR = 0.70


def get_rec_factor_for_part(df: pd.DataFrame, part: str) -> float:
    if CLASS_COLUMN not in df.columns:
        return DEFAULT_REC_FACTOR
    classes = (
        df.loc[df["PART NBR"] == part, CLASS_COLUMN]
        .dropna()
        .astype(str)
        .str.upper()
        .str.strip()
    )
    if classes.empty:
        return DEFAULT_REC_FACTOR
    return CLASS_REC_FACTORS.get(classes.iloc[0], DEFAULT_REC_FACTOR)


# ==============================
# Utility helpers
# ==============================

def todays_date() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    if "Comments" not in df.columns:
        df.insert(min(8, len(df.columns)), "Comments", "")
    if "Last Week Comments" not in df.columns:
        idx = df.columns.get_loc("Comments") + 1
        df.insert(idx, "Last Week Comments", "")
    if "Recommendation" not in df.columns:
        insert_at = min(df.columns.get_loc("Comments") + 2, len(df.columns))
        df.insert(insert_at, "Recommendation", "")
    return df


# ==============================
# MG4 tagging
# ==============================

def mg4Copilot(db: pd.DataFrame, mg4_df: pd.DataFrame) -> pd.DataFrame:
    lookup = dict(zip(mg4_df["MATERIAL"], mg4_df["MG4_FLAG"]))
    db["MG4 Result"] = db["PART NBR"].map(lookup)
    return db


def partNumbers(db: pd.DataFrame, target_flag: str = "R") -> pd.DataFrame:
    target_flag = str(target_flag).strip().upper()
    data = db[db["MG4 Result"].astype(str).str.strip().str.upper() == target_flag]
    return data[["PART NBR", "SHIP UNIT", "ACCT UNIT"]].drop_duplicates()


# ==============================
# SPM & Forecast
# ==============================

def spm_inventory_data(part: str, dist_ctr: str, df_spm: pd.DataFrame):
    if str(dist_ctr).strip() == "3Q01":
        dist_ctr = "2003"

    result = df_spm[(df_spm["PRT NUM"] == str(part).strip()) & (df_spm["DIST_CTR"] == str(dist_ctr).strip())]
    if result.empty:
        return (0, 0, 0, 0, 0, True)

    row = result.iloc[0].copy()
    for inventory in ["AVAIL", "INHOUSE", "WIP", "INTRANSIT", "ON_ORDER"]:
        row[inventory] = pd.to_numeric(row.get(inventory, 0), errors="coerce")
    return (
        int(row.get("AVAIL", 0) or 0),
        int(row.get("INHOUSE", 0) or 0),
        int(row.get("WIP", 0) or 0),
        int(row.get("INTRANSIT", 0) or 0),
        int(row.get("ON_ORDER", 0) or 0),
        False,
    )


def spm_search_by_mtl_fct_prts_Avg(fcst_df_raw: pd.DataFrame) -> pd.DataFrame:
    df = fcst_df_raw.copy()
    df.columns = df.columns.astype(str).str.strip().str.replace("|", "", regex=False)
    if not all(c in df.columns for c in ["Month", "PRT NUM", "Value"]):
        raise ValueError("Forecast file must have columns: Month, PRT NUM, Value")

    filtered = df[df["Month"].isin(["F_M_1", "F_M_2", "F_M_3"])]
    avg = filtered.groupby("PRT NUM", as_index=False)["Value"].mean()
    avg.rename(columns={"Value": "Average_F_M_1_to_3"}, inplace=True)
    return avg


# ==============================
# Core calculations
# ==============================

def calc_total_bo_for_part_depot(global_bo_df: pd.DataFrame, part: str, dist_ctr: str) -> float:
    mask = (
        global_bo_df["PART NBR"].astype(str).str.strip() == str(part).strip()
    ) & (
        global_bo_df["SHIP UNIT"].astype(str).str.strip() == str(dist_ctr).strip()
    )
    return pd.to_numeric(global_bo_df.loc[mask, "BO QTY"], errors="coerce").fillna(0).sum()


def format_depot_entry(date_str, depot, bo, avail, inhouse, wip, intransit, onorder, note=None, note2=None):
    base = (
        f"{date_str}: {depot}: {bo} BO - AVAIL({avail}) - INHOUSE({inhouse}) - "
        f"WIP({wip}) - INTRANSIT({intransit}) - ON_ORDER({onorder})"
    )
    if note:
        base += f" - {note}"
    if note2:
        base += f" - {note2}"
    return base


# ==============================
# Recommendation note logic
# ==============================

def coverage_note(bo_total: float, avail: int, inhouse: int, wip: int, intransit: int) -> str:
    base_onhand = avail + inhouse + wip
    full_onhand = base_onhand + intransit

    if bo_total > base_onhand and bo_total <= full_onhand:
        return "In transit"
    if bo_total > avail and bo_total <= (avail + inhouse):
        return "INHOUSE"
    return "Covered"


def build_and_apply_comments_for_part(
    db_all: pd.DataFrame,
    fcst_avg: pd.DataFrame,
    part: str,
    depot_lines_collector: list,
    issues_collector: list,
    df_spm: pd.DataFrame,
) -> pd.DataFrame:
    today = todays_date()
    part_rows = db_all[db_all["PART NBR"] == part]
    if part_rows.empty:
        return db_all

    depots = sorted(part_rows["SHIP UNIT"].dropna().unique().astype(str).tolist())

    fcst_row = fcst_avg.loc[fcst_avg["PRT NUM"] == part, "Average_F_M_1_to_3"]
    fcst_val = float(fcst_row.iloc[0]) if not fcst_row.empty else 0.0

    rec_factor = get_rec_factor_for_part(db_all, part)

    depot_lines = []
    for depot in depots:
        avail, inhouse, wip, intransit, onorder, missing = spm_inventory_data(part, depot, df_spm)
        base_onhand = avail + inhouse + wip
        full_onhand = base_onhand + intransit

        bo_total = calc_total_bo_for_part_depot(db_all, part, depot)
        rec_pcs = int(max(0, (bo_total + fcst_val)) * rec_factor)

        note2 = None

        if missing:
            note = "Review SPM (missing row)"
            issues_collector.append({"PART NBR": part, "SHIP UNIT": depot, "Issue": "SPM missing row", "When": today})
        elif bo_total > full_onhand:
            gap = bo_total - full_onhand
            note = f"Recommend ship ~{max(gap, rec_pcs)} pcs"
        else:
            note = coverage_note(bo_total, avail, inhouse, wip, intransit)

        line = format_depot_entry(today, depot, int(bo_total), avail, inhouse, wip, intransit, onorder, note, note2)
        depot_lines.append(line)

        depot_lines_collector.append({
            "Date": today,
            "PART NBR": part,
            "SHIP UNIT": depot,
            "BO": int(bo_total),
            "AVAIL": avail,
            "INHOUSE": inhouse,
            "WIP": wip,
            "INTRANSIT": intransit,
            "ON_ORDER": onorder,
            "Forecast_Avg_F_M_1_to_3": fcst_val,
            "Recommendation_Note": note,
        })

        mask = (db_all["PART NBR"] == part) & (db_all["SHIP UNIT"].astype(str).str.strip() == str(depot).strip())
        db_all.loc[mask, "Recommendation"] = note

    combined = "; ".join(depot_lines)
    last_week_any = part_rows["Last Week Comments"].dropna().astype(str)
    last_week_blob = last_week_any.iloc[0] if not last_week_any.empty else ""
    final_comment = combined if not last_week_blob else f"{combined}; {last_week_blob}"

    db_all.loc[db_all["PART NBR"] == part, "Comments"] = final_comment
    return db_all


# ==============================
# End check/backfill
# ==============================

def backfill_missing_recommendations(
    db: pd.DataFrame,
    fcst_avg: pd.DataFrame,
    spm_df: pd.DataFrame,
    issues_collector: list,
    target_flag: str,
) -> pd.DataFrame:
    today = todays_date()
    tf = str(target_flag).strip().upper()

    mask_missing = db["Recommendation"].fillna("").astype(str).str.strip() == ""

    # If BOTH, backfill for ALL rows with missing recommendation.
    if tf == "BOTH":
        to_fix = db[mask_missing][["PART NBR", "SHIP UNIT"]].drop_duplicates()
    else:
        # Otherwise, backfill only the selected MG4 type (R or V)
        if "MG4 Result" not in db.columns:
            raise ValueError("MG4 Result missing. Ensure MG4 file was processed for R/V mode.")
        mask_target = db["MG4 Result"].astype(str).str.upper().str.strip() == tf
        to_fix = db[mask_target & mask_missing][["PART NBR", "SHIP UNIT"]].drop_duplicates()

    for _, r in to_fix.iterrows():
        part = str(r["PART NBR"]).strip()
        depot = str(r["SHIP UNIT"]).strip()

        avail, inhouse, wip, intransit, onorder, missing = spm_inventory_data(part, depot, spm_df)
        base_onhand = avail + inhouse + wip
        full_onhand = base_onhand + intransit

        bo_total = calc_total_bo_for_part_depot(db, part, depot)

        fcst_row = fcst_avg.loc[fcst_avg["PRT NUM"] == part, "Average_F_M_1_to_3"]
        fcst_val = float(fcst_row.iloc[0]) if not fcst_row.empty else 0.0

        rec_factor = get_rec_factor_for_part(db, part)
        rec_pcs = int(max(0, (bo_total + fcst_val)) * rec_factor)

        if missing:
            note = "Review SPM (missing row)"
        elif bo_total > full_onhand:
            gap = bo_total - full_onhand
            note = f"Recommend ship ~{max(gap, rec_pcs)} pcs"
        else:
            note = coverage_note(bo_total, avail, inhouse, wip, intransit)

        db.loc[
            (db["PART NBR"] == part) & (db["SHIP UNIT"].astype(str).str.strip() == depot),
            "Recommendation"
        ] = note

        issues_collector.append({
            "PART NBR": part,
            "SHIP UNIT": depot,
            "Issue": "Recommendation was blank; backfilled",
            "When": today
        })

    return db


# ==============================
# Reporting builders
# ==============================

def build_part_summary(depot_df: pd.DataFrame, full_db: pd.DataFrame) -> pd.DataFrame:
    if depot_df.empty:
        return pd.DataFrame(columns=[
            "PART NBR", "BO", "AVAIL", "INHOUSE", "WIP", "INTRANSIT", "ON_ORDER",
            "Total_OnHand", "Forecast_Avg_F_M_1_to_3", "Rec_Pieces_Est", "Status"
        ])

    agg = depot_df.groupby("PART NBR", as_index=False).agg({
        "BO": "sum",
        "AVAIL": "sum",
        "INHOUSE": "sum",
        "WIP": "sum",
        "INTRANSIT": "sum",
        "ON_ORDER": "sum",
        "Forecast_Avg_F_M_1_to_3": "max",
    })

    agg["Total_OnHand"] = agg["AVAIL"] + agg["INHOUSE"] + agg["WIP"]
    agg["Rec_Factor"] = agg["PART NBR"].apply(lambda p: get_rec_factor_for_part(full_db, p))
    agg["Rec_Pieces_Est"] = (agg["BO"] + agg["Forecast_Avg_F_M_1_to_3"]).clip(lower=0) * agg["Rec_Factor"]

    agg["Status"] = agg.apply(
        lambda r: "Ship more" if r["BO"] > (r["Total_OnHand"] + r["INTRANSIT"])
        else ("Healthy" if r["Total_OnHand"] >= r["Rec_Pieces_Est"] else "Covered"),
        axis=1
    )

    return agg[[
        "PART NBR", "BO", "AVAIL", "INHOUSE", "WIP", "INTRANSIT", "ON_ORDER",
        "Total_OnHand", "Forecast_Avg_F_M_1_to_3", "Rec_Pieces_Est", "Status"
    ]]


def build_planner_view_all(db_data: pd.DataFrame, depot_lines_df: pd.DataFrame) -> pd.DataFrame:
    if depot_lines_df.empty:
        out = db_data.copy()
        for c in ["BO", "AVAIL", "INHOUSE", "WIP", "INTRANSIT", "ON_ORDER",
                  "Total_OnHand", "Forecast_Avg_F_M_1_to_3", "Rec_Pieces_Est", "Status"]:
            out[c] = 0 if c != "Status" else ""
        return out

    temp = depot_lines_df.copy()
    temp["Total_OnHand"] = temp[["AVAIL", "INHOUSE", "WIP"]].sum(axis=1)

    rec_factor_map = {part: get_rec_factor_for_part(db_data, part) for part in temp["PART NBR"].unique()}
    temp["Rec_Factor"] = temp["PART NBR"].map(rec_factor_map)
    temp["Rec_Pieces_Est"] = (temp["BO"] + temp["Forecast_Avg_F_M_1_to_3"]).clip(lower=0) * temp["Rec_Factor"]

    temp["Status"] = temp.apply(
        lambda r: "Ship more" if r["BO"] > (r["Total_OnHand"] + r["INTRANSIT"])
        else ("Healthy" if r["Total_OnHand"] >= r["Rec_Pieces_Est"] else "Covered"),
        axis=1
    )

    depot_latest = (
        temp.sort_values(["PART NBR", "SHIP UNIT", "Date"], ascending=[True, True, False])
        .drop_duplicates(subset=["PART NBR", "SHIP UNIT"], keep="first")
    )

    merged = db_data.merge(
        depot_latest[["PART NBR", "SHIP UNIT", "BO", "AVAIL", "INHOUSE", "WIP", "INTRANSIT", "ON_ORDER",
                     "Total_OnHand", "Forecast_Avg_F_M_1_to_3", "Rec_Pieces_Est", "Status"]],
        on=["PART NBR", "SHIP UNIT"],
        how="left"
    )

    for c in ["BO", "AVAIL", "INHOUSE", "WIP", "INTRANSIT", "ON_ORDER",
              "Total_OnHand", "Forecast_Avg_F_M_1_to_3", "Rec_Pieces_Est"]:
        merged[c] = pd.to_numeric(merged[c], errors="coerce").fillna(0).astype(float)
    merged["Status"] = merged["Status"].fillna("")
    return merged


# ==============================
# Excel writer -> memory bytes
# ==============================

def _col_letter(n: int) -> str:
    import string
    letters = ""
    while n:
        n, rem = divmod(n - 1, 26)
        letters = string.ascii_uppercase[rem] + letters
    return letters


def build_outputs_workbook(db: pd.DataFrame, depot_lines_collector: list, issues_collector: list) -> bytes:
    depot_df = pd.DataFrame(depot_lines_collector) if depot_lines_collector else pd.DataFrame(
        columns=["Date", "PART NBR", "SHIP UNIT", "BO", "AVAIL", "INHOUSE", "WIP",
                 "INTRANSIT", "ON_ORDER", "Forecast_Avg_F_M_1_to_3", "Recommendation_Note"]
    )
    part_summary = build_part_summary(depot_df, db)
    issues_df = pd.DataFrame(issues_collector) if issues_collector else pd.DataFrame(
        columns=["PART NBR", "SHIP UNIT", "Issue", "When"]
    )
    planner_all = build_planner_view_all(db, depot_df)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        db.to_excel(writer, sheet_name="Data", index=False)
        depot_df.to_excel(writer, sheet_name="Depot_Lines", index=False)
        part_summary.to_excel(writer, sheet_name="Part_Summary", index=False)
        planner_all.to_excel(writer, sheet_name="Planner_View_All", index=False)
        issues_df.to_excel(writer, sheet_name="Issues", index=False)

        wb = writer.book
        ws = writer.sheets["Planner_View_All"]

        headers = {col: idx for idx, col in enumerate(planner_all.columns, start=1)}
        nrows = len(planner_all) + 1

        red = wb.add_format({"bg_color": "#FFC7CE"})
        yellow = wb.add_format({"bg_color": "#FFEB9C"})
        green = wb.add_format({"bg_color": "#C6EFCE"})

        if all(h in headers for h in ["BO", "Total_OnHand", "Rec_Pieces_Est"]):
            cBO = _col_letter(headers["BO"])
            cTOH = _col_letter(headers["Total_OnHand"])
            cREC = _col_letter(headers["Rec_Pieces_Est"])
            rng = f"A2:{_col_letter(len(headers))}{nrows}"

            ws.conditional_format(rng, {"type": "formula", "criteria": f"=${cBO}2>${cTOH}2", "format": red})
            ws.conditional_format(rng, {"type": "formula",
                                        "criteria": f"=AND(${cBO}2<={cTOH}2, ${cTOH}2<${cREC}2)",
                                        "format": yellow})
            ws.conditional_format(rng, {"type": "formula", "criteria": f"=${cTOH}2>={cREC}2", "format": green})

    buffer.seek(0)
    return buffer.getvalue()


# ==============================
# Pipeline
# ==============================

def run_pipeline(
    bo_df: pd.DataFrame,
    mg4_df: Optional[pd.DataFrame],
    spm_df: pd.DataFrame,
    fcst_export_df: pd.DataFrame,
    target_flag: str = "R",
) -> Tuple[bytes, bytes]:
    db = bo_df.copy()
    db = ensure_columns(db)

    tf = str(target_flag).strip().upper()

    # If BOTH: do NOT use MG4 tagging/filtering
    if tf == "BOTH":
        parts_df = db[["PART NBR", "SHIP UNIT", "ACCT UNIT"]].drop_duplicates()
    else:
        if mg4_df is None or mg4_df.empty:
            raise ValueError("MG4 file is required for Factory Direct (R) or Vendor Direct (V).")
        db = mg4Copilot(db, mg4_df)
        parts_df = partNumbers(db, target_flag=tf)

    fcst_avg = spm_search_by_mtl_fct_prts_Avg(fcst_export_df)

    depot_lines_collector: list = []
    issues_collector: list = []

    for part in sorted(parts_df["PART NBR"].unique().tolist()):
        db = build_and_apply_comments_for_part(db, fcst_avg, part, depot_lines_collector, issues_collector, spm_df)

    db = backfill_missing_recommendations(db, fcst_avg, spm_df, issues_collector, target_flag=tf)

    buf_data = BytesIO()
    db.to_excel(buf_data, index=False)
    buf_data.seek(0)
    data_with_comments_bytes = buf_data.getvalue()

    bo_outputs_bytes = build_outputs_workbook(db, depot_lines_collector, issues_collector)
    return data_with_comments_bytes, bo_outputs_bytes


# ==============================
# Streamlit UI
# ==============================

def main():
    st.title("Global Back Order Recommendation Tool")

    st.sidebar.header("1) Download latest input files")
    st.sidebar.markdown(
        "- [Global Backorder Report](%s)\n"
        "- [MG4 Material Mapping](%s)\n"
        "- [SPM Search by Material Tool](%s)\n"
        "- [Forecast Export](%s)"
        % (
            DOWNLOAD_LINKS["Global Backorder"],
            DOWNLOAD_LINKS["MG4 Tool"],
            DOWNLOAD_LINKS["SPM Search by Material"],
            DOWNLOAD_LINKS["Forecast Export (spm_search_by_mtl_fct_prts)"],
        )
    )
    st.sidebar.markdown("---")
    if st.sidebar.button("Clear cache"):
        st.cache_data.clear()
        st.sidebar.success("Cache cleared.")

    st.subheader("Step 1 – Download latest input files")
    st.markdown(
        f"""
- [Global Backorder Report]({DOWNLOAD_LINKS["Global Backorder"]})
- [MG4 Material Mapping]({DOWNLOAD_LINKS["MG4 Tool"]})
- [SPM Search by Material Tool]({DOWNLOAD_LINKS["SPM Search by Material"]})
- [Forecast Export]({DOWNLOAD_LINKS["Forecast Export (spm_search_by_mtl_fct_prts)"]})
"""
    )

    st.markdown(
        """
Upload your latest files and generate:
- **data_with_comments.xlsx** – original BO file + Comments/Recommendations
- **bo_outputs.xlsx** – multi-sheet planner workbook
"""
    )

    col1, col2, col3 = st.columns([1.2, 1.2, 1.6])
    with col1:
        target_flag = st.radio(
            "Process parts type",
            options=["R", "V", "BOTH"],
            format_func=lambda x: (
                "Factory Direct (R)" if x == "R"
                else "Vendor Direct (V)" if x == "V"
                else "Both (R + V) — MG4 not required"
            ),
            horizontal=True
        )

    with col3:
        st.caption("Uploads accept Excel (.xlsx/.xls) or CSV (.csv).")

    with st.form("bo_form"):
        bo_file = st.file_uploader("Backorder export (Excel/CSV)", type=["xlsx", "xls", "csv"])

        # Only ask for MG4 if R or V
        mg4_file = None
        if str(target_flag).strip().upper() != "BOTH":
            mg4_file = st.file_uploader("MG4 file (Excel/CSV)", type=["xlsx", "xls", "csv"])
        else:
            st.info("MG4 upload skipped because you selected BOTH.")

        spm_file = st.file_uploader("SPM export (Excel/CSV)", type=["xlsx", "xls", "csv"])
        fcst_file = st.file_uploader("Forecast export (Excel/CSV)", type=["xlsx", "xls", "csv"])
        submitted = st.form_submit_button("Run BO Copilot")

    if not submitted:
        return

    tf = str(target_flag).strip().upper()

    required_ok = all([bo_file, spm_file, fcst_file]) and (tf == "BOTH" or mg4_file is not None)
    if not required_ok:
        if tf == "BOTH":
            st.error("Please upload Backorder, SPM, and Forecast files.")
        else:
            st.error("Please upload **all four** required files (Backorder, MG4, SPM, Forecast).")
        return

    try:
        with st.spinner("Reading files…"):
            bo_df = read_table_cached(bo_file)

            mg4_df = None
            if tf != "BOTH":
                mg4_df = read_table_cached(mg4_file)

            spm_df = read_table_cached(spm_file)

            # Forecast: if Excel, try Export sheet; if CSV, just read
            fcst_name = (fcst_file.name or "").lower()
            if fcst_name.endswith(".csv"):
                fcst_export_df = read_table_cached(fcst_file)
            else:
                try:
                    fcst_export_df = read_table_cached(fcst_file, sheet_name="Export")
                except Exception:
                    xls = pd.ExcelFile(fcst_file)
                    first_sheet = xls.sheet_names[0]
                    fcst_export_df = read_table_cached(fcst_file, sheet_name=first_sheet)
                    st.warning(f"Forecast sheet 'Export' not found. Used '{first_sheet}' instead.")

        with st.spinner("Validating & normalizing inputs…"):
            bo_df, mg4_df, spm_df, fcst_export_df = validate_and_prepare_inputs(
                bo_df, mg4_df, spm_df, fcst_export_df,
                require_mg4=(tf != "BOTH")
            )

        # Optional caching warm-up
        _ = forecast_avg_cached(fcst_export_df)

        with st.spinner("Processing recommendations…"):
            data_with_comments_bytes, bo_outputs_bytes = run_pipeline(
                bo_df, mg4_df, spm_df, fcst_export_df,
                target_flag=tf,
            )

        st.success("Done ✅")

        st.download_button(
            label="Download data_with_comments.xlsx",
            data=data_with_comments_bytes,
            file_name=f"data_with_comments_{todays_date()}_{tf}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button(
            label="Download bo_outputs.xlsx",
            data=bo_outputs_bytes,
            file_name=f"bo_outputs_{todays_date()}_{tf}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Error during processing: {e}")


if __name__ == "__main__":
    main()

