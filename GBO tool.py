# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import streamlit as st

import pandas as pd

from datetime import datetime

from io import BytesIO

 

# ==============================

# Class-based recommendation factors

# ==============================

 

CLASS_COLUMN = "INV CLASS"

 

CLASS_REC_FACTORS = {

    "A": 0.90,

    "B": 0.80,

    "C": 0.75,

    "D": 0.70,

    "E": 0.70,

}

 

DEFAULT_REC_FACTOR = 0.70  # fallback if class missing/unknown

 

 

def get_rec_factor_for_part(df: pd.DataFrame, part: str) -> float:

    """

    Look up INV CLASS for a part in df and return its recommendation factor.

    If INV CLASS is missing / not present, fall back to DEFAULT_REC_FACTOR.

    """

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

 

    cls = classes.iloc[0]

    return CLASS_REC_FACTORS.get(cls, DEFAULT_REC_FACTOR)

 

 

# ==============================

# Utility helpers

# ==============================

 

def todays_date() -> str:

    now = datetime.now()

    return f"{now.month}/{now.day}"  # e.g., 8/17

 

 

def ensure_columns(df: pd.DataFrame) -> pd.DataFrame:

    # make sure comment columns exist and are placed side-by-side

    if "Comments" not in df.columns:

        insert_at = min(8, len(df.columns))

        df.insert(insert_at, "Comments", "")

    if "Last Week Comments" not in df.columns:

        idx = df.columns.get_loc("Comments") + 1

        df.insert(idx, "Last Week Comments", "")

    if "Recommendation" not in df.columns:

        df.insert(8, "Recommendation", "")

    return df

 

 

def move_comments_to_last_week(df: pd.DataFrame) -> pd.DataFrame:

    df["Last Week Comments"] = df["Comments"].fillna("")

    df["Comments"] = ""  # rebuild fresh for this week

    return df

 

 

# ==============================

# MG4 tagging

# ==============================

 

def mg4Copilot(db: pd.DataFrame, mg4_df: pd.DataFrame) -> pd.DataFrame:

    if "MATERIAL" not in mg4_df.columns or "PAG/ PDX MG 4" not in mg4_df.columns:

        raise ValueError("MG4 file missing required columns: MATERIAL, PAG/ PDX MG 4")

    lookup = dict(zip(mg4_df["MATERIAL"], mg4_df["PAG/ PDX MG 4"]))

    db["MG4 Result"] = db["PART NBR"].map(lookup)

    return db

 

 

def partNumbers(db: pd.DataFrame) -> pd.DataFrame:

    # filter for Factory Direct 'R' parts

    data_R = db[db["MG4 Result"] == "R"]

    return data_R[["PART NBR", "SHIP UNIT", "ACCT UNIT"]].drop_duplicates()

 

 

# ==============================

# SPM & Forecast inputs

# ==============================

 

def spm_inventory_data(part: str, dist_ctr: str, df_spm: pd.DataFrame):

    """

    Return tuple: (AVAIL, INHOUSE, WIP, INTRANSIT, ON_ORDER, missing_flag)

    """

    for col in ["PRT NUM", "DIST_CTR"]:

        if col not in df_spm.columns:

            raise ValueError(f"SPM file is missing column: {col}")

 

    # region-specific alias: in SPM 3Q01 equals 2003

    if dist_ctr == "3Q01":

        dist_ctr = "2003"

 

    result = df_spm[(df_spm["PRT NUM"] == part) & (df_spm["DIST_CTR"] == dist_ctr)]

    if result.empty:

        return (0, 0, 0, 0, 0, True)

    row = result.iloc[0].copy()

    for inventory in ["AVAIL", "INHOUSE", "WIP", "INTRANSIT", "ON_ORDER"]:

        row[inventory] = pd.to_numeric(row.get(inventory, 0), errors="coerce")

    return (

        int(row["AVAIL"] or 0),

        int(row["INHOUSE"] or 0),

        int(row["WIP"] or 0),

        int(row["INTRANSIT"] or 0),

        int(row["ON_ORDER"] or 0),

        False,

    )

 

 

def spm_search_by_mtl_fct_prts_Avg(fcst_df_raw: pd.DataFrame) -> pd.DataFrame:

    """

    Takes the forecast export DataFrame (sheet 'Export') and

    averages F_M_1..F_M_3 per PRT NUM.

    """

    df = fcst_df_raw.copy()

    df.columns = df.columns.str.strip().str.replace("|", "", regex=False)

    if "Month" not in df.columns or "PRT NUM" not in df.columns or "Value" not in df.columns:

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

    ) & (global_bo_df["SHIP UNIT"].astype(str).str.strip() == str(dist_ctr).strip())

    return pd.to_numeric(global_bo_df.loc[mask, "BO QTY"], errors="coerce").fillna(0).sum()

 

 

def format_depot_entry(date_str, depot, bo, avail, inhouse, wip, intransit, onorder, note=None, note2=None):

    base = (

        f"{date_str}: {depot}: {bo} BO - AVAIL({avail}) - INHOUSE({inhouse}) - "

        f"WIP ({wip})- In Transit({intransit}) - On Order({onorder})"

    )

    if note:

        base += f" - {note}"

    if note2:

        base += f" - {note2}"

    return base

 

 

def build_and_apply_comments_for_part(db_R: pd.DataFrame,

                                      fcst_avg: pd.DataFrame,

                                      part: str,

                                      depot_lines_collector: list,

                                      issues_collector: list,

                                      df_spm: pd.DataFrame) -> pd.DataFrame:

    """

    For a PART NBR, make one new 'Comments' string that concatenates all depot lines for today,

    then append last week's text at the end. Also collect normalized depot rows and issues.

    """

 

    today = todays_date()

    part_rows = db_R[db_R["PART NBR"] == part]

 

    if part_rows.empty:

        return db_R

 

    depots = sorted(part_rows["SHIP UNIT"].dropna().unique().astype(str).tolist())

 

    # forecast lookup for this part

    try:

        fcst_row = fcst_avg.loc[fcst_avg["PRT NUM"] == part, "Average_F_M_1_to_3"]

        fcst_val = float(fcst_row.iloc[0]) if not fcst_row.empty else 0.0

    except Exception:

        fcst_val = 0.0

 

    # Use INV CLASS to choose the factor (A=0.9, B=0.8, C=0.75, D/E=0.7)

    rec_factor = get_rec_factor_for_part(db_R, part)

 

    depot_lines = []

    for depot in depots:

        avail, inhouse, wip, intransit, onorder, missing = spm_inventory_data(part, depot, df_spm)

        total_onhand = avail + inhouse + wip + intransit

        bo_total = calc_total_bo_for_part_depot(db_R, part, depot)

 

        rec_pcs = int(max(0, (bo_total + fcst_val)) * rec_factor)

 

        if missing:

            note = "Review SPM (missing row)"

            issues_collector.append(

                {"PART NBR": part, "SHIP UNIT": depot, "Issue": "SPM missing row", "When": today}

            )

            note2 = None

 

        elif bo_total > total_onhand:

            gap = bo_total - total_onhand

            note = f"Recommend ship ~{max(gap, rec_pcs)} pcs"

            note2 = None

 

        else:

            note = "Covered"

            note2 = None

 

        if str(depot) == "2003" and note != "Covered":

            DY_bo_total = calc_total_bo_for_part_depot(db_R, part, depot)

            if DY_bo_total > total_onhand:

                gap = DY_bo_total - total_onhand

                note2 = f"Recommend ship ~{max(gap, rec_pcs)} pcs"

            elif DY_bo_total <= 0 and total_onhand >= rec_pcs:

                surplus = total_onhand - rec_pcs

                note2 = f"Healthy; potential surplus {surplus} pcs"

            else:

                note2 = "Covered"

 

        # line for concatenated Comments

        line = format_depot_entry(

            today, depot, int(bo_total), avail, inhouse, wip, intransit, onorder, note, note2

        )

        depot_lines.append(line)

 

        # normalized output row

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

 

    combined = "; ".join(depot_lines)

    last_week_any = part_rows["Last Week Comments"].dropna().astype(str)

    last_week_blob = last_week_any.iloc[0] if not last_week_any.empty else ""

    final_comment = combined if not last_week_blob else f"{combined}; {last_week_blob}"

 

    db_R.loc[db_R["PART NBR"] == part, "Comments"] = final_comment

 

    # Add recommendation note

    for depot_row in depot_lines_collector:

        if depot_row["PART NBR"] == part:

            mask = (db_R["PART NBR"] == part) & (db_R["SHIP UNIT"] == depot_row["SHIP UNIT"])

            db_R.loc[mask, "Recommendation"] = depot_row["Recommendation_Note"]

 

    return db_R

 

 

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

 

    # per-part factor based on INV CLASS

    agg["Rec_Factor"] = agg["PART NBR"].apply(lambda p: get_rec_factor_for_part(full_db, p))

    agg["Rec_Pieces_Est"] = (agg["BO"] + agg["Forecast_Avg_F_M_1_to_3"]).clip(lower=0) * agg["Rec_Factor"]

 

    agg["Status"] = agg.apply(

        lambda r: "Ship more" if r["BO"] > r["Total_OnHand"]

        else ("Healthy" if r["Total_OnHand"] >= r["Rec_Pieces_Est"] else "Covered"),

        axis=1

    )

 

    return agg[[

        "PART NBR", "BO", "AVAIL", "INHOUSE", "WIP", "INTRANSIT", "ON_ORDER",

        "Total_OnHand", "Forecast_Avg_F_M_1_to_3", "Rec_Pieces_Est", "Status"

    ]]

 

 

def build_planner_view_all(db_data: pd.DataFrame, depot_lines_df: pd.DataFrame) -> pd.DataFrame:

    """

    Returns a table with **EVERY original column** from BO data for each row,

    merged with depot metrics & status computed for (PART NBR, SHIP UNIT).

    """

    if depot_lines_df.empty:

        out = db_data.copy()

        for c in ["BO", "AVAIL", "INHOUSE", "WIP", "INTRANSIT", "ON_ORDER",

                  "Total_OnHand", "Forecast_Avg_F_M_1_to_3", "Rec_Pieces_Est", "Status"]:

            out[c] = 0 if c != "Status" else ""

        return out

 

    temp = depot_lines_df.copy()

    temp["Total_OnHand"] = temp[["AVAIL", "INHOUSE", "WIP"]].sum(axis=1)

 

    # Map each part to its INV CLASS-based factor

    rec_factor_map = {

        part: get_rec_factor_for_part(db_data, part)

        for part in temp["PART NBR"].unique()

    }

    temp["Rec_Factor"] = temp["PART NBR"].map(rec_factor_map)

    temp["Rec_Pieces_Est"] = (temp["BO"] + temp["Forecast_Avg_F_M_1_to_3"]).clip(lower=0) * temp["Rec_Factor"]

 

    temp["Status"] = temp.apply(

        lambda r: "Ship more" if r["BO"] > r["Total_OnHand"]

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

        if c in merged.columns:

            merged[c] = pd.to_numeric(merged[c], errors="coerce").fillna(0).astype(float)

    if "Status" in merged.columns:

        merged["Status"] = merged["Status"].fillna("")

 

    return merged

 

 

# ==============================

# Excel writer -> memory bytes

# ==============================

 

def _col_letter(n: int) -> str:

    import string

    letters = ""

    while n:

        n, rem = divmod(n-1, 26)

        letters = string.ascii_uppercase[rem] + letters

    return letters

 

 

def build_outputs_workbook(db: pd.DataFrame,

                           depot_lines_collector: list,

                           issues_collector: list) -> bytes:

    """

    Build the multi-sheet bo_outputs.xlsx in memory and return bytes.

    """

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

 

        headers = {col: idx for idx, col in enumerate(planner_all.columns, start=1)}  # 1-based

        nrows = len(planner_all) + 1  # header row included

 

        red = wb.add_format({"bg_color": "#FFC7CE"})

        yellow = wb.add_format({"bg_color": "#FFEB9C"})

        green = wb.add_format({"bg_color": "#C6EFCE"})

 

        if all(h in headers for h in ["BO", "Total_OnHand", "Rec_Pieces_Est"]):

            cBO = _col_letter(headers["BO"])

            cTOH = _col_letter(headers["Total_OnHand"])

            cREC = _col_letter(headers["Rec_Pieces_Est"])

            rng = f"A2:{_col_letter(len(headers))}{nrows}"

 

            ws.conditional_format(rng, {

                "type": "formula",

                "criteria": f"=${cBO}2>${cTOH}2",

                "format": red

            })

            ws.conditional_format(rng, {

                "type": "formula",

                "criteria": f"=AND(${cBO}2<={cTOH}2, ${cTOH}2<${cREC}2)",

                "format": yellow

            })

            ws.conditional_format(rng, {

                "type": "formula",

                "criteria": f"=${cTOH}2>={cREC}2",

                "format": green

            })

 

    buffer.seek(0)

    return buffer.getvalue()

 

 

# ==============================

# Pipeline

# ==============================

 

def run_pipeline(bo_df: pd.DataFrame,

                 mg4_df: pd.DataFrame,

                 spm_df: pd.DataFrame,

                 fcst_export_df: pd.DataFrame) -> tuple[bytes, bytes]:

    """

    Core pipeline:

      - takes four DataFrames

      - returns (data_with_comments_bytes, bo_outputs_bytes)

    """

 

    db = bo_df.copy()

    db = ensure_columns(db)

    # db = move_comments_to_last_week(db)  # enable if you want to roll comments weekly

 

    db = mg4Copilot(db, mg4_df)

 

    parts_df = partNumbers(db)

 

    fcst_avg = spm_search_by_mtl_fct_prts_Avg(fcst_export_df)

 

    depot_lines_collector: list = []

    issues_collector: list = []

 

    for part in sorted(parts_df["PART NBR"].unique().tolist()):

        db = build_and_apply_comments_for_part(

            db, fcst_avg, part, depot_lines_collector, issues_collector, spm_df

        )

 

    # data_with_comments.xlsx

    buf_data = BytesIO()

    db.to_excel(buf_data, index=False)

    buf_data.seek(0)

    data_with_comments_bytes = buf_data.getvalue()

 

    # bo_outputs.xlsx

    bo_outputs_bytes = build_outputs_workbook(db, depot_lines_collector, issues_collector)

 

    return data_with_comments_bytes, bo_outputs_bytes

 

 

# ==============================

# Streamlit UI

# ==============================

 

def main():

    st.title("Global Back Order reccemendation")

   

    DOWNLOAD_LINKS = {

    #  URLs 

    "Global Backorder": "url"
,

    "MG4 Tool": "Url"
,

    "SPM Search by Material": "url",

    "Forecast Export (spm_search_by_mtl_fct_prts)": "url",


}  

 

#🔹 Sidebar links (Step 1)

    st.sidebar.header("1. Download latest input files")

 

    # NOTE: these use the DOWNLOAD_LINKS dict defined at the top.

    st.sidebar.markdown(

        "- [Global Backorder Report](%s)\n"

        "- [MG4 Material Mapping](%s)\n"

        "- [SPM Search by Material Tool](%s)\n"

        "- [Forecast Export (spm_search_by_mtl_fct_prts)](%s)"

        % (

            DOWNLOAD_LINKS["Global Backorder"],

            DOWNLOAD_LINKS["MG4 Tool"],

            DOWNLOAD_LINKS["SPM Search by Material"],

            DOWNLOAD_LINKS["Forecast Export (spm_search_by_mtl_fct_prts)"],

        )

    )

    st.sidebar.markdown("---")

    st.sidebar.caption("Open each link, export/download the Excel, then upload below 👇")

   

    #populate

    st.subheader("Step 1 – Download latest input files")

    st.markdown(

        f"""

        - [Global Backorder Report]({DOWNLOAD_LINKS["Global Backorder"]}) 

        - [MG4 Material Mapping]({DOWNLOAD_LINKS["MG4 Tool"]}) 

        - [SPM Search by Material Tool]({DOWNLOAD_LINKS["SPM Search by Material"]}) 

        - [Forecast Export (spm_search_by_mtl_fct_prts)]({DOWNLOAD_LINKS["Forecast Export (spm_search_by_mtl_fct_prts)"]}) 

        """

    )

 

 

    st.markdown(

        """

        Upload your latest files and generate:

        - **data_with_comments.xlsx** – original BO file + Comments/Recommendations 

        - **bo_outputs.xlsx** – multi-sheet planner workbook 

        """

    )

 

    with st.form("bo_form"):

        bo_file = st.file_uploader("Backorder export (e.g. data.xlsx)", type=["xlsx"])

        mg4_file = st.file_uploader("MG4 file (MATERIAL / PAG/ PDX MG 4)", type=["xlsx"])

        spm_file = st.file_uploader("SPM search by material tool", type=["xlsx"])

        fcst_file = st.file_uploader("Forecast export (spm_search_by_mtl_fct_prts)", type=["xlsx"])

 

        submitted = st.form_submit_button("Run BO Copilot")

 

    if submitted:

        if not all([bo_file, mg4_file, spm_file, fcst_file]):

            st.error("Please upload **all four** required files.")

            return

 

        try:

            with st.spinner("Processing…"):

                bo_df = pd.read_excel(bo_file)

                mg4_df = pd.read_excel(mg4_file)

 

                # SPM – default sheet

                spm_df = pd.read_excel(spm_file)

 

                # Forecast – sheet 'Export' like your original

                fcst_export_df = pd.read_excel(fcst_file, sheet_name="Export")

 

                data_with_comments_bytes, bo_outputs_bytes = run_pipeline(

                    bo_df, mg4_df, spm_df, fcst_export_df

                )

 

            st.success("Done")

 

            st.download_button(

                label="Download data_with_comments.xlsx",

                data=data_with_comments_bytes,

                file_name="data_with_comments.xlsx",

                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

            )

 

            st.download_button(

                label="Download bo_outputs.xlsx",

                data=bo_outputs_bytes,

                file_name="bo_outputs.xlsx",

                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

            )

 

        except Exception as e:

            st.error(f"Error during processing: {e}")

 

 

if __name__ == "__main__":


    main()
