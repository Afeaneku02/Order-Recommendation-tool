# Order-Recommendation-tool

## Overview

app replaces hours of manual cross checking with a single automated workflow

App utilizes 4 reports:
- Global Backorder report  
- MG4 part classification  
- SPM Search by Material inventory  
- Forecast export (F_M_1–F_M_3)


It then generates:

- Automated comments per part & depot  
- Recommendation logic based on **INV CLASS** (A/B/C/D/E)  
- Two output Excel files:
1.) data_with_comments.xlsx
2.) bo_outputs.xlsx



Before running the tool, download the latest versions of:

1. **Global Backorder Report**  
2. **MG4 Material Mapping** (Factory Direct / Vendor)  
3. **SPM Search by Material** (inventory export)  
4. **Forecast Export**  
   - Must contain a sheet named `Export`
   - Includes forecast values for `F_M_1, F_M_2, F_M_3


## 🧠 Key Features

- ✅ Auto-calculates total on-hand inventory per part & depot  
- ✅ Combines **backorders + inventory + forecast** in one place  
- ✅ Uses **INV CLASS-based factors** for recommendations (A, B, C, D/E classes)  
- ✅ Generates clear comments like _“Recommend ship ~X pcs”_ or _“Covered”_  
- ✅ Builds a planner workbook with traffic-light conditional formatting  
- ✅ Flags missing SPM rows and other data issues 


  
## 🖥 How to Use the App

1. **Open the app**  
   - Via the hosted Streamlit URL, or  
   - Locally with `streamlit run app.py`

2. **Download the input files**  
   - Use the links in the left sidebar (if configured), or  
   - Pull them manually from your usual systems (Global BO, MG4, SPM, Forecast)

3. **Upload the four files** in the main app:
   - Backorder export  
   - MG4 file  
   - SPM Search by Material export  
   - Forecast export (`Export` sheet)

4. Click **“Run BO Copilot”**.

5. When processing is complete, click the buttons to:
   - Download `data_with_comments.xlsx`  
   - Download `bo_outputs.xlsx`  

---

## 📄 Generated Outputs

### 1. `data_with_comments.xlsx`
- Original backorder export  
- `Comments` column populated with depot-level notes  
- `Recommendation` column summarizing suggested actions  
- Optional: `Last Week Comments` can be rolled forward if enabled in code  

### 2. `bo_outputs.xlsx`
Multi-sheet workbook including:

- **Data** – processed BO data  
- **Depot_Lines** – one row per part & depot with BO, inventory, forecast, recommendation  
- **Part_Summary** – part-level rollup with BO, on-hand, forecast, and status  
- **Planner_View_All** – designed for daily/weekly planner review with conditional formatting  
- **Issues** – parts/depot combinations with missing SPM rows or other flags  
