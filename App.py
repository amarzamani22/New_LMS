import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import re
import os
import plotly.graph_objects as go
import plotly.express as px
from typing import Dict, List, Tuple # Added for type hinting

# --- Page Configuration (Set first for a professional look) ---
st.set_page_config(
    page_title="LMS Analysis Dashboard",
    page_icon="🏦",
    layout="wide"
)

# --- Modern CSS Styling ---
st.markdown("""
<style>
    /* Main app background */
    .stApp {
        background-color: #F0F2F6;
    }
    /* Metric cards styling */
    .stMetric {
        border-radius: 10px;
        padding: 20px;
        background-color: #FFFFFF;
        border: 1px solid #E0E0E0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.04);
    }
    /* Avoid brittle emotion class names; keep base styling lean */
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# --- Global Constants (Added for maintainability) ---
SHEET_MAP = {'Q1A: Employees': 'QC_Q1A_Main', 'Q2A: Salary': 'QC_Q2A_Main', 'Q3: Hours Worked': 'QC_Q3', 'Q4: Vacancies': 'QC_Q4', 'Q5: Separations': 'QC_Q5'}
ALL_MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
ROLLUP_KEY = "All Financial Institutions"
ENTITY_COL = "Entity / Group"
SUBQ_COL = "Subquestion"
WC_COL = "Worker Category"

# ========= NEW: VR config & helpers (ADD-ONLY; no change to your logic) =========
VR_PATH  = "submission/vr_staging.xlsx"   # <-- put your VR consolidated file path here
VR_SHEET = "Variance"
_VR_MONTH_END = {"Q1":"Mar","Q2":"Jun","Q3":"Sep","Q4":"Dec"}
VR_THRESHOLD_PCT = 25  # ±25% for "Required justification"

@st.cache_data(ttl=3600)
def _vr_load_variance(vr_path: str = VR_PATH, sheet_name: str = VR_SHEET) -> pd.DataFrame:
    """Load VR Variance and build a JOIN_KEY."""
    if not os.path.exists(vr_path):
        return pd.DataFrame(columns=["JOIN_KEY","%Growth","Justification"])
    df = pd.read_excel(vr_path, sheet_name=sheet_name)
    if df.empty:
        return pd.DataFrame(columns=["JOIN_KEY","%Growth","Justification"])
    for c in ["Entity Name","Question","Subquestion","Worker Category","Month"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip().str.upper()
        else:
            df[c] = ""
    df["JOIN_KEY"] = (
        df["Entity Name"] + "|" +
        df["Year"].astype(str) + "|" +
        df["Month"] + "|" +
        df["Question"] + "|" +
        df["Subquestion"] + "|" +
        df["Worker Category"]
    )
    # keep last per JOIN_KEY (latest)
    keep = ["JOIN_KEY","%Growth","Justification","Year","Month"]
    df = (df[keep]
          .sort_values(["JOIN_KEY","Year","Month"])
          .groupby("JOIN_KEY", as_index=False)
          .tail(1))
    return df[["JOIN_KEY","%Growth","Justification"]]

def _vr_infer_year_month_from_period(period_label: str, default_year: int) -> tuple[int, str]:
    """
    Accepts: 'Mar 2025', '2025-Mar', 'Q2 2025', 'Q2' (uses default_year), '2025-Mar', 'Mar', '2025-Mar'
    Returns: (year:int, month_abbr:'Jan'..'Dec')
    """
    s = str(period_label).strip()
    # 2025-Mar or 2025-March
    if "-" in s and s.split("-")[0].isdigit():
        yr, mon = s.split("-")[0], s.split("-")[1][:3].title()
        return int(yr), mon
    # Mar 2025 or March 2025
    parts = s.split()
    if len(parts) == 2 and parts[1].isdigit():
        return int(parts[1]), parts[0][:3].title()
    # Qx 2025
    if s.startswith("Q") and " " in s and s.split()[1].isdigit():
        q, yr = s.split()
        return int(yr), _VR_MONTH_END.get(q, "Mar")
    # Qx (no year)
    if s.startswith("Q") and s[:2] in _VR_MONTH_END:
        return int(default_year), _VR_MONTH_END.get(s[:2], "Mar")
    # YYYY-Mon was handled above; else just month string
    if len(s) >= 3:
        return int(default_year), s[:3].title()
    return int(default_year), "Mar"

def _vr_build_join_key(entity, year, month, question, subq, wc) -> str:
    return (
        str(entity).upper().strip() + "|" +
        str(int(year)) + "|" +
        str(month).upper().strip() + "|" +
        str(question).upper().strip() + "|" +
        str(subq).upper().strip() + "|" +
        str(wc).upper().strip()
    )

def _vr_to_percent_number(x):
    """Return numeric percent (e.g., '41%' -> 41.0, '41' -> 41.0, 0.41 -> 41.0)."""
    if pd.isna(x): return np.nan
    s = str(x).strip()
    try:
        if s.endswith("%"): return float(s[:-1])
        v = float(s)
        return v * 100 if 0 <= v <= 1 else v
    except Exception:
        return np.nan

def _vr_is_blank_just(s):
    if s is None: return True
    t = str(s).strip().upper()
    return t in {"", "-", "N/A"}

# --- 1. Helper Functions (Preserved from your original code) ---
@st.cache_data
def load_data(year, sheet_name):
    """Loads data for a specific year and sheet."""
    file_path = f"submission/qc_workbook_{year}.xlsx"
    if not os.path.exists(file_path):
        return f"Error: File not found. Please ensure '{file_path}' exists in a 'submission' subfolder."
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=5)
        df.dropna(axis=1, how='all', inplace=True)
        # Keep your current rename behavior
        rename_map = {'Q1.1': 'Q1_Total', 'Q2.1': 'Q2_Total', 'Q3.1': 'Q3_Total', 'Q4.1': 'Q4_Total'}
        df.rename(columns=rename_map, inplace=True)
        return df
    except Exception as e:
        return f"Error reading sheet '{sheet_name}' from {file_path}: {e}"

@st.cache_data
def get_reporting_quarter(year):
    """Reads the _About sheet from the specified year's workbook to get the quarter."""
    file_path = f"submission/qc_workbook_{year}.xlsx"
    try:
        about_df = pd.read_excel(file_path, sheet_name="_About", header=None)
        quarter_row = about_df[about_df[0] == 'Quarter']
        if not quarter_row.empty:
            return int(re.search(r'\d+', str(quarter_row.iloc[0, 1])).group())
    except Exception:
        pass
    return 4 # Default to Q4 if sheet is missing or malformed

def _row_filter(df, entity, worker_cat, subq):
    """Robust row filter that tolerates missing Subquestion col."""
    if df.empty:
        return pd.DataFrame()
    cond = (df['Entity / Group'] == entity) & (df['Worker Category'] == worker_cat)
    if 'Subquestion' in df.columns and subq != "N/A":
        cond &= (df['Subquestion'] == subq)
    return df[cond]

# --- MODIFIED: find_outliers (FIXED KeyError) ---
def find_outliers(data_series, prior_year_series, pct_thresh, abs_thresh, iqr_multiplier, yoy_thresh):
    """
    Your original outlier detection function from App.py.
    FIXED: Added check for 'if not outliers' to prevent KeyError on empty list.
    """
    outliers = []
    # If data is insufficient, return empty DataFrame with expected columns
    if data_series.isnull().all() or len(data_series) < 2: 
        return pd.DataFrame(columns=["Period", "Value", "Reason(s)"]).set_index("Period")

    q1, q3 = data_series.quantile(0.25), data_series.quantile(0.75)
    iqr = q3 - q1 if len(data_series) >= 4 else 0
    iqr_lower_bound, iqr_upper_bound = q1 - (iqr_multiplier * iqr), q3 + (iqr_multiplier * iqr)

    for i, (period_name, current_value) in enumerate(data_series.items()):
        reasons = []
        if pd.isna(current_value): continue

        if i > 0:
            prev_val = data_series.iloc[i-1]
            if not pd.isna(prev_val) and prev_val != 0:
                abs_change = current_value - prev_val; pct_change = abs_change / prev_val
                # Uses the pct_thresh variable passed into the function
                if abs(pct_change) > pct_thresh and abs(abs_change) > abs_thresh: 
                    reasons.append(f"High Volatility ({pct_change:+.1%})")
        
        if iqr > 0 and (current_value < iqr_lower_bound or current_value > iqr_upper_bound): 
            reasons.append("IQR Anomaly")
            
        if prior_year_series is not None and period_name in prior_year_series.index:
            prior_value = prior_year_series.get(period_name) # Use .get() for safety
            if not pd.isna(prior_value) and prior_value != 0:
                yoy_change = (current_value - prior_value) / prior_value
                if abs(yoy_change) > yoy_thresh: 
                    reasons.append(f"YoY Anomaly ({yoy_change:+.1%})")
        
        if reasons:
            outliers.append({"Period": period_name, "Value": f"{current_value:,.2f}", "Reason(s)": ", ".join(reasons)})

    # --- FIX for KeyError ---
    if not outliers:
        # Return an empty, indexed DataFrame if no outliers were found
        return pd.DataFrame(columns=["Period", "Value", "Reason(s)"]).set_index("Period")
    # --- End FIX ---
            
    return pd.DataFrame(outliers).set_index("Period") # Original return

def _quarter_labels(year, upto_q):
    return [f"Q{i} {year}" for i in range(1, upto_q + 1)]

def _compute_quarter_totals_from_months(row, upto_q, months_order):
    """Compute quarter totals from monthly columns if Qx_Total is missing."""
    q_map = {1: months_order[0:3], 2: months_order[3:6], 3: months_order[6:9], 4: months_order[9:12]}
    vals = []
    labels = []
    for q in range(1, upto_q + 1):
        mlist = [m for m in q_map[q] if m in row.index]
        if mlist:
            vals.append(row[mlist].astype(float).sum())
            labels.append(f"Q{q}")
    return labels, vals

# --- RESTORED: build_quarter_series (from your 08:31:07 App.py) ---
def build_quarter_series(data_row, year, reporting_quarter):
    """
    Returns a pd.Series indexed by 'Qx YYYY' using existing Qx_Total if present,
    otherwise computes from months available.
    """
    if data_row.empty:
        return pd.Series(dtype=float)

    # Prefer explicit totals if available
    q_cols = [f'Q{i}_Total' for i in range(1, reporting_quarter + 1)]
    have_any_qtot = any(col in data_row.columns for col in q_cols)
    series_vals = []
    labels = []

    if have_any_qtot:
        for q in range(1, reporting_quarter + 1):
            col = f'Q{q}_Total'
            if col in data_row.columns:
                val = pd.to_numeric(data_row.iloc[0][col], errors='coerce')
                if not pd.isna(val):
                    series_vals.append(float(val))
                    labels.append(f"Q{q} {year}")

    # If none (or partial) are available, try computing from months for missing quarters
    months_order = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    # Build a quick lookup of which quarters already added
    present_quarters = {int(l.split()[0][1:]) for l in labels}
    for q in range(1, reporting_quarter + 1):
        if q in present_quarters:
            continue
        # compute
        q_to_months = {1: months_order[0:3], 2: months_order[3:6], 3: months_order[6:9], 4: months_order[9:12]}
        mlist = [m for m in q_to_months[q] if m in data_row.columns]
        if mlist:
            vals = pd.to_numeric(data_row.iloc[0][mlist], errors='coerce').astype(float)
            if not np.all(np.isnan(vals)):
                series_vals.append(float(np.nansum(vals)))
                labels.append(f"Q{q} {year}")

    if not labels:
        return pd.Series(dtype=float)
    # sort by quarter number
    sort_idx = np.argsort([int(l.split()[0][1:]) for l in labels])
    labels = [labels[i] for i in sort_idx]
    series_vals = [series_vals[i] for i in sort_idx]

    s = pd.Series(series_vals, index=labels, dtype=float)
    return s

# --- RESTORED: generate_full_report (from your 08:31:07 App.py) ---
def generate_full_report(
   sheet_map,
   years,
   questions_to_scan,
   thresholds,
   all_months,
   report_filters=None,
   quarter_mode="current",          # "current" | "up_to_selected" | "exact_selected"
   selected_quarter=None            # 1..4 when quarter_mode != "current"
):
   def months_for_quarter(q: int):
       q_map = {1: all_months[0:3], 2: all_months[3:6], 3: all_months[6:9], 4: all_months[9:12]}
       return q_map.get(int(q), [])
   # ---- detection window (what we compute on) vs focus window (what we show) ----
   if quarter_mode == "current":
       rq = get_reporting_quarter(years['current'])
       months_to_use  = all_months[: rq * 3]       # detect on Q1..Q(rq) so April sees March
       focus_months   = months_for_quarter(rq)     # BUT only show Q(rq)
   elif quarter_mode == "up_to_selected" and selected_quarter:
       months_to_use  = all_months[: int(selected_quarter) * 3]  # detect on Q1..Qn
       focus_months   = months_to_use                             # and show all up to Qn
   elif quarter_mode == "exact_selected" and selected_quarter:
       months_to_use  = all_months[: int(selected_quarter) * 3]  # detect on Q1..Qn (keeps March for Apr MoM)
       focus_months   = months_for_quarter(int(selected_quarter)) # but only show Qn
   else:
       months_to_use  = all_months
       focus_months   = all_months
   focus_set = set(focus_months)

   master_outlier_list = []
   progress_bar = st.progress(0, text="Initializing Scan...")
   total_scans = max(1, len(questions_to_scan))
   for i, q_name in enumerate(questions_to_scan):
       sheet_name = sheet_map[q_name]
       progress_bar.progress(i / total_scans, text=f"Scanning: {q_name}")
       df_current = load_data(years['current'], sheet_name)
       df_prior   = load_data(years['prior'], sheet_name) if years.get('prior') else None
       if isinstance(df_current, str):   # file error
           continue
       # align by months present in sheet + quarter scope
       actual_months = [m for m in months_to_use if m in df_current.columns]
       if not actual_months:
           continue
       # read filter for this dataset (if any)
       f_cfg = (report_filters or {}).get(q_name, {"subq": "ALL", "wc": "ALL"})
       subq_filter = f_cfg.get("subq", "ALL")
       wc_filter   = f_cfg.get("wc", "ALL")
       # iterate rows with optional filters
       for _, row in df_current.iterrows():
           entity = row['Entity / Group']
           wc     = row['Worker Category']
           subq   = row['Subquestion'] if 'Subquestion' in df_current.columns else 'N/A'
           # apply filter(s) if specified
           if subq_filter != "ALL" and subq not in subq_filter:
               continue
           if wc_filter   != "ALL" and wc   not in wc_filter:
               continue
           # current series
           monthly_series = pd.to_numeric(row[actual_months], errors='coerce').astype(float)
           # build prior aligned series (same months)
           prior_series = None
           if not isinstance(df_prior, str) and df_prior is not None:
               prior_row = _row_filter(df_prior, entity, wc, subq)
               if not prior_row.empty:
                   prior_series = pd.to_numeric(prior_row[actual_months].iloc[0], errors='coerce').astype(float)
           
           # Call the *same* find_outliers function
           outliers_monthly = find_outliers(monthly_series, prior_series, **thresholds)
           
           # Loop over the returned DataFrame's index and rows
           for period, o_row in outliers_monthly.iterrows():
            if period not in focus_set:
                continue  # show only the selected quarter/month window
            master_outlier_list.append([
                q_name, entity, subq, wc,
                'Monthly', period, o_row['Value'], o_row['Reason(s)']
            ])
       # (Optional) You can later extend here to also write a quarterly view if needed.
   progress_bar.progress(1.0, text="Scan Complete!")
   return pd.DataFrame(
       master_outlier_list,
       columns=['Question', 'Entity / Group', 'Subquestion', 'Worker Category', 'View', 'Period', 'Value', 'Reason(s)']
   )

# --- RESTORED: Multi-Year Helpers (from your 08:31:07 App.py) ---
def _year_range(start_year: int, end_year: int):
    return [y for y in range(int(start_year), int(end_year) + 1)]

def build_multi_year_monthly_series(entity, worker_cat, subq, selected_question, start_year, end_year, exclusions, all_months, sheet_map):
    """
    Returns (series, prior_series_for_yoy) where:
      - series index is ['YYYY-MMM', ...]
      - prior_series aligns keys (YYYY-MMM) -> previous year's same month value, if available
    """
    vals, idx = [], []
    for yr in _year_range(start_year, end_year):
        df_y = load_data(yr, sheet_map[selected_question])
        if isinstance(df_y, str):
            continue
        row_y = _row_filter(df_y, entity, worker_cat, subq)
        if row_y.empty:
            continue
        upto_q_y = get_reporting_quarter(yr)
        months_y = [m for m in all_months[:upto_q_y * 3] if m in row_y.columns]
        if not months_y:
            continue
        ser_y = pd.to_numeric(row_y[months_y].iloc[0], errors='coerce').astype(float)
        for m in months_y:
            label = f"{yr}-{m}"
            if label in exclusions:
                continue
            v = ser_y.get(m, np.nan)
            vals.append(v)
            idx.append(label)

    if not idx:
        return pd.Series(dtype=float), None

    s = pd.Series(vals, index=idx, dtype=float)

    prior_vals, prior_idx = [], []
    for label in s.index:
        try:
            yr_str, mon = label.split('-')
            prev_label = f"{int(yr_str)-1}-{mon}"
            if prev_label in s.index and not pd.isna(s.loc[prev_label]):
                prior_vals.append(float(s.loc[prev_label]))
                prior_idx.append(label)
        except Exception:
            continue

    prior = pd.Series(prior_vals, index=prior_idx, dtype=float) if prior_idx else None
    return s, prior

def build_multi_year_quarterly_series(entity, worker_cat, subq, selected_question, start_year, end_year, exclusions, sheet_map):
    """
    Returns (series, prior_series_for_yoy) where:
      - series index is ['Qx YYYY', ...]
      - prior_series aligns keys (Qx YYYY) -> (Qx YYYY-1) value, if available
    """
    parts = []
    for yr in _year_range(start_year, end_year):
        df_y = load_data(yr, sheet_map[selected_question])
        if isinstance(df_y, str):
            continue
        row_y = _row_filter(df_y, entity, worker_cat, subq)
        if row_y.empty:
            continue
        rq = get_reporting_quarter(yr)
        s_y = build_quarter_series(row_y, yr, rq)
        if s_y.empty:
            continue
        keep = [lab for lab in s_y.index if lab not in exclusions]
        parts.append(s_y.loc[keep] if keep else pd.Series(dtype=float))

    if not parts:
        return pd.Series(dtype=float), None

    s = pd.concat(parts)
    order = sorted(s.index, key=lambda lab: (int(lab.split()[1]), int(lab.split()[0][1:])))
    s = s.loc[order]

    prior_vals, prior_idx = [], []
    for label in s.index:
        try:
            q, yr_str = label.split()
            prev_label = f"{q} {int(yr_str)-1}"
            if prev_label in s.index and not pd.isna(s.loc[prev_label]):
                prior_vals.append(float(s.loc[prev_label]))
                prior_idx.append(label)
        except Exception:
            continue

    prior = pd.Series(prior_vals, index=prior_idx, dtype=float) if prior_idx else None
    return s, prior

# --- NEW Helper for Attribution Panel (FIXED) ---
def prepare_attribution_data(
    df: pd.DataFrame, 
    period_label: str, 
    time_view: str
) -> pd.DataFrame:
    """
    Identifies the correct difference/change columns (Diff Mar, Diff Q2)
    and renames them for consistent attribution analysis.
    FIXED: Logic now correctly identifies quarter label from the period_label.
    """
    if time_view == 'Monthly':
        # Monthly label is just the month (e.g., 'Mar')
        month = period_label.split(' ')[-1] 
        diff_col = f"Diff {month}"
        pct_col = f"MoM {month}"
    else: # Quarterly
        # Quarterly index label is 'Qx YYYY', but the column name is just 'Diff Qx'.
        q_label = period_label.split(' ')[0] # <--- THE FIX: Extracts 'Qx' from 'Qx YYYY'
        diff_col = f"Diff {q_label}"
        pct_col = f"%Diff {q_label}"

    # Check if the required columns exist in the loaded DataFrame
    if diff_col not in df.columns or pct_col not in df.columns:
        return pd.DataFrame() 
        
    df_out = df[[ENTITY_COL, SUBQ_COL, WC_COL, diff_col, pct_col]].copy()
    
    # Clean and convert the percentage column
    df_out[pct_col] = (df_out[pct_col].astype(str).str.replace('%', '', regex=False)
                                     .replace("N/A", np.nan).astype(float) / 100)
    
    df_out.rename(columns={
        diff_col: "Absolute Change (Contribution)", 
        pct_col: "Percentage Change"
    }, inplace=True)

    return df_out.dropna(subset=["Absolute Change (Contribution)"])

# --- Main Dashboard Interface ---
def main():
    st.title("🏦 LMS Analysis Dashboard")

    # --- Global Time Context & Sidebar Controls ---
    st.sidebar.title("Analysis Controls")
    current_year = st.sidebar.selectbox("Current Year:", [2025, 2024, 2023, 2022, 2021])
    comparison_year = st.sidebar.selectbox("Comparison Year:", [2024, 2023, 2022, 2021, 2020, None], index=0)
    st.sidebar.markdown("---")
    analysis_view = st.sidebar.radio("Select View:", ('Interactive Analysis', 'Full Outlier Report'), label_visibility="collapsed")
    st.sidebar.markdown("---")

    reporting_quarter = get_reporting_quarter(current_year)
    months_to_analyze = ALL_MONTHS[:reporting_quarter * 3]
    quarters_to_analyze = [f'Q{i+1}' for i in range(reporting_quarter)]
    st.sidebar.info(f"Analyzing **{current_year}** data up to **Q{reporting_quarter}**.")

    # --- View 1: Interactive Analysis ---
    if analysis_view == 'Interactive Analysis':
        st.sidebar.subheader("Thresholds")
        # ADDED: MoM/QoQ Threshold Slider (to pass to your find_outliers function)
        mom_pct_thresh = st.sidebar.slider("MoM/QoQ % Threshold", 0, 100, 25, 5, format="%d%%") / 100.0
        abs_thresh = st.sidebar.slider("Absolute Change", 10, 1000, 50, 10)
        iqr_mult = st.sidebar.slider("IQR Sensitivity", 1.0, 3.0, 1.5, 0.1)
        yoy_thresh_pct = st.sidebar.slider("YoY % Threshold", 0, 100, 30, 5, format="%d%%")
        yoy_thresh = yoy_thresh_pct / 100.0
        
        st.sidebar.markdown("---")
        enable_multi = st.sidebar.checkbox("Enable multi-year timeline", value=False)

        start_year = current_year
        end_year = current_year
        exclude_months = []
        exclude_quarters = []

        time_view = st.sidebar.radio("Frequency:", ('Monthly', 'Quarterly'), horizontal=True)
        selected_question = st.sidebar.selectbox("Dataset:", options=list(SHEET_MAP.keys()))

        if enable_multi:
            start_year = st.sidebar.number_input("Start year", min_value=2019, max_value=current_year, value=min(2021, current_year))
            end_year = st.sidebar.number_input("End year", min_value=int(start_year), max_value=current_year, value=current_year)

            if time_view == 'Monthly':
                month_opts = []
                for yr in range(int(start_year), int(end_year) + 1):
                    rq = get_reporting_quarter(yr)
                    month_opts.extend([f"{yr}-{m}" for m in ALL_MONTHS[:rq * 3]])
                exclude_months = st.sidebar.multiselect("Exclude months (e.g., 2023-Feb):", options=month_opts, default=[])
            else:
                qtr_opts = []
                for yr in range(int(start_year), int(end_year) + 1):
                    rq = get_reporting_quarter(yr)
                    qtr_opts.extend([f"Q{q} {yr}" for q in range(1, rq + 1)])
                exclude_quarters = st.sidebar.multiselect("Exclude quarters (e.g., Q3 2023):", options=qtr_opts, default=[])

        # --- Data Loading (Moved top-level) ---
        df_current = load_data(current_year, SHEET_MAP[selected_question])
        df_prior = load_data(comparison_year, SHEET_MAP[selected_question]) if comparison_year else None

        if isinstance(df_current, str):
            st.error(df_current)
        else:
            # --- Detection View (Your original logic) ---
            st.header(f"1. Outlier Detection: {selected_question}")
            
            # --- Entity/Category Selection (Your original logic) ---
            entity = st.selectbox("Select an Entity / Group:", options=df_current['Entity / Group'].unique(), 
                                  index=list(df_current['Entity / Group'].unique()).index(ROLLUP_KEY) if ROLLUP_KEY in df_current['Entity / Group'].unique() else 0)
            
            if 'Subquestion' in df_current.columns and df_current[df_current['Entity / Group'] == entity]['Subquestion'].nunique() > 1:
                default_subq_index = np.where(df_current[df_current['Entity / Group'] == entity]['Subquestion'].unique() == 'Employment = A+B(i)+B(ii)')[0][0] if 'Employment = A+B(i)+B(ii)' in df_current[df_current['Entity / Group'] == entity]['Subquestion'].unique() else 0
                subquestion = st.selectbox("Select a Subquestion:", options=df_current[df_current['Entity / Group'] == entity]['Subquestion'].unique(), index=int(default_subq_index))
                
                default_wc_index = np.where(df_current[(df_current['Entity / Group'] == entity) & (df_current['Subquestion'] == subquestion)]['Worker Category'].unique() == 'Total Employment')[0][0] if 'Total Employment' in df_current[(df_current['Entity / Group'] == entity) & (df_current['Subquestion'] == subquestion)]['Worker Category'].unique() else 0
                worker_cat = st.selectbox("Select a Worker Category:", options=df_current[(df_current['Entity / Group'] == entity) & (df_current['Subquestion'] == subquestion)]['Worker Category'].unique(), index=int(default_wc_index))
                
                data_row = _row_filter(df_current, entity, worker_cat, subquestion)
            else:
                subquestion = "N/A"
                default_wc_index = np.where(df_current[df_current['Entity / Group'] == entity]['Worker Category'].unique() == 'Total Employment')[0][0] if 'Total Employment' in df_current[df_current['Entity / Group'] == entity]['Worker Category'].unique() else 0
                worker_cat = st.selectbox("Select a Worker Category:", options=df_current[df_current['Entity / Group'] == entity]['Worker Category'].unique(), index=int(default_wc_index))
                data_row = _row_filter(df_current, entity, worker_cat, subquestion)

            st.caption(f"Displaying: {entity} | {subquestion} | {worker_cat}")
            
            current_series, prior_series = pd.Series(dtype=float), None
            
            # --- FIXED: Robust check for df_prior (resolves ValueError) ---
            prior_row = pd.DataFrame()
            if isinstance(df_prior, pd.DataFrame) and not df_prior.empty:
                prior_row = _row_filter(df_prior, entity, worker_cat, subquestion)
            # --- END FIXED ---

            if time_view == 'Monthly':
                actual_months = [m for m in months_to_analyze if m in data_row.columns]
                if actual_months:
                    current_series = pd.to_numeric(data_row[actual_months].iloc[0], errors='coerce').astype(float)
                    current_series.index = actual_months

                if not prior_row.empty and all(m in prior_row.columns for m in actual_months):
                    prior_series = pd.to_numeric(prior_row[actual_months].iloc[0], errors='coerce').astype(float)
                    prior_series.index = current_series.index

                title = "Monthly Trend"

                if enable_multi:
                    current_series, prior_series = build_multi_year_monthly_series(
                        entity=entity, worker_cat=worker_cat, subq=subquestion,
                        selected_question=selected_question, start_year=int(start_year),
                        end_year=int(end_year), exclusions=set(exclude_months),
                        all_months=ALL_MONTHS, sheet_map=SHEET_MAP
                    )

            else:  # Quarterly
                current_series = build_quarter_series(data_row, current_year, reporting_quarter)
                title = "Quarterly Trend"

                if not prior_row.empty:
                    prior_rq = get_reporting_quarter(comparison_year)
                    prior_series_full = build_quarter_series(prior_row, comparison_year, prior_rq)
                    # Align indices for YoY comparison (e.g., 'Q1 2025' looks for 'Q1 2024')
                    prior_series = prior_series_full.copy()
                    prior_series.index = current_series.index.map(lambda x: f"{x.split()[0]} {int(x.split()[1])-1}")
                    prior_series = prior_series.loc[prior_series.index.intersection(prior_series_full.index)]


                if enable_multi:
                    current_series, prior_series = build_multi_year_quarterly_series(
                        entity=entity, worker_cat=worker_cat, subq=subquestion,
                        selected_question=selected_question, start_year=int(start_year),
                        end_year=int(end_year), exclusions=set(exclude_quarters),
                        sheet_map=SHEET_MAP
                    )

            if current_series.empty or current_series.isnull().all():
                st.warning("No data found for the selected filters or period.")
            else:
                # Call your find_outliers function, now passing the slider value
                outlier_df = find_outliers(
                    current_series, prior_series, 
                    mom_pct_thresh, # Use variable from slider
                    abs_thresh, iqr_mult, yoy_thresh
                )

                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Latest Value", f"{current_series.iloc[-1]:,.0f}")
                col2.metric("Period Average", f"{current_series.mean():,.0f}")
                col3.metric("Period High", f"{np.nanmax(current_series):,.0f}")
                col4.metric("True Outliers", len(outlier_df))
                
                # --- Plotly chart (Your original logic) ---
                x = list(current_series.index)
                y = current_series.values.astype(float)
                fig = go.Figure()

                fig.add_trace(go.Scatter(
                    x=x, y=y, mode='lines+markers', name=f'{current_year} Trend' if not enable_multi else 'Multi-year Trend',
                    hovertemplate='Period: %{x}<br>Value: %{y:,.0f}<extra></extra>'
                ))

                if len(y) >= 2:
                    upper = [None] + [y[i-1] * (1 + mom_pct_thresh) for i in range(1, len(y))] # Use slider value
                    lower = [None] + [y[i-1] * (1 - mom_pct_thresh) for i in range(1, len(y))] # Use slider value

                    fig.add_trace(go.Scatter(
                        x=x, y=lower, mode='lines', line=dict(width=0),
                        name=f'MoM ±{mom_pct_thresh:.0%} Range', showlegend=False, hoverinfo='skip'
                    ))
                    fig.add_trace(go.Scatter(
                        x=x, y=upper, mode='lines', line=dict(width=0),
                        fill='tonexty', name=f'MoM ±{mom_pct_thresh:.0%} Range', opacity=0.15,
                        hovertemplate='Prev×1.25: %{y:,.0f}<extra></extra>'
                    ))

                if not outlier_df.empty:
                    o_x = [p for p in outlier_df.index if p in current_series.index]
                    o_y = [current_series[p] for p in o_x]
                    o_reason = [outlier_df.loc[p, 'Reason(s)'] for p in o_x]
                    fig.add_trace(go.Scatter(
                        x=o_x, y=o_y, mode='markers', name='True Outlier',
                        marker=dict(symbol='x', size=14, color='red'), # Simplified color
                        hovertemplate='Period: %{x}<br>Value: %{y:,.0f}<br>%{customdata}<extra></extra>',
                        customdata=o_reason
                    ))

                fig.update_layout(
                    title=dict(text=title + (" — Multi-year" if enable_multi else ""), font=dict(size=16), x=0.5, xanchor='center'),
                    hovermode='x unified',
                    legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='left', x=0),
                    margin=dict(l=10, r=10, t=70, b=10),
                )
                fig.update_xaxes(tickangle=45)
                st.plotly_chart(fig, use_container_width=True)
                
                if not outlier_df.empty:
                    st.error("🚨 True Outlier(s) Detected!")
                    st.dataframe(outlier_df, use_container_width=True)

                    # ===== NEW: VR merge for Interactive Analysis (MoM only) =====
                    # Only meaningful in Monthly view (VR is MoM). For multi-year, Period looks like 'YYYY-MMM'.
                    try:
                        vr_df = _vr_load_variance()  # cached
                        # Build JOIN_KEY for each outlier row
                        keys = []
                        for per in outlier_df.index.tolist():
                            if enable_multi:
                                # per like '2024-Mar'
                                yr, mon = _vr_infer_year_month_from_period(per, default_year=current_year)
                            else:
                                # per like 'Jan'..'Dec' or 'Qx YYYY' (skip QoQ)
                                if time_view == 'Monthly':
                                    yr, mon = _vr_infer_year_month_from_period(per, default_year=current_year)
                                else:
                                    # Quarterly view: VR is MoM → skip
                                    continue
                            k = _vr_build_join_key(
                                entity=entity,
                                year=yr,
                                month=mon,
                                question=selected_question,
                                subq=subquestion if subquestion else "N/A",
                                wc=worker_cat
                            )
                            keys.append((per, k))

                        if keys:
                            out_vr = outlier_df.copy().reset_index()
                            out_vr.rename(columns={"index":"Period"}, inplace=True)
                            out_vr["JOIN_KEY"] = [k for (_, k) in keys]
                            merged = out_vr.merge(vr_df, on="JOIN_KEY", how="left")

                            # statuses
                            merged["growth_pct_num"] = merged["%Growth"].apply(_vr_to_percent_number)
                            merged["vr_submitted"]   = merged["%Growth"].notna() | merged["Justification"].notna()
                            merged["breaches"]       = merged["growth_pct_num"].abs() >= VR_THRESHOLD_PCT

                            def _status(r):
                                if not r["vr_submitted"]:
                                    return "⌛ Pending"
                                if r["breaches"] and _vr_is_blank_just(r["Justification"]):
                                    return "⚠️ Required justification"
                                return "✅ Matched" if not _vr_is_blank_just(r["Justification"]) else "—"

                            merged["Status"] = merged.apply(_status, axis=1)

                            st.subheader("VR Justification (MoM) — Matched to Outliers")
                            st.caption("Legend: ✅ Matched · ⚠️ Required justification · ⌛ Pending")
                            st.dataframe(
                                merged[["Period","Value","Reason(s)","%Growth","Status","Justification"]],
                                use_container_width=True
                            )
                    except Exception as e:
                        st.warning(f"VR merge skipped: {e}")
                    # ===== END VR merge =====

                else:
                    st.success("✅ No significant outliers detected.")
                
            # --- NEW: Outlier Attribution Panel ---
            # This panel only appears if we are viewing the top aggregate AND outliers were found
            if entity == ROLLUP_KEY and not outlier_df.empty:
                st.header("2. Outlier Attribution Panel 🔍")
                
                outlier_periods = outlier_df.index.tolist()
                selected_outlier_period = st.selectbox(
                    "Select Outlier Period to Analyze (Attribution):", 
                    options=outlier_periods
                )
                
                # Prepare data for the selected period
                full_attribution_df = prepare_attribution_data(
                    df=df_current, # Use the full, unfiltered data
                    period_label=selected_outlier_period,
                    time_view=time_view
                )

                if full_attribution_df.empty:
                    st.error(f"Could not find the necessary 'Diff' (Difference) column in the data for {selected_outlier_period}. Attribution failed.")
                else:
                    # --- Step A: Entity/Group Contribution ---
                    st.subheader("A. Entity / Group Contribution")
                    st.caption(f"Contribution of **ALL** entities/rollups to the change in **{selected_outlier_period}**")
                    
                    # Filter for the *major view's* context
                    df_entity_contrib = full_attribution_df[
                        (full_attribution_df[SUBQ_COL] == subquestion) &
                        (full_attribution_df[WC_COL] == worker_cat)
                    ].copy()
                    
                    df_entity_contrib = df_entity_contrib.sort_values("Absolute Change (Contribution)", ascending=False).reset_index(drop=True)
                    
                    st.dataframe(df_entity_contrib.style.format({
                        "Absolute Change (Contribution)": "{:,.0f}",
                        "Percentage Change": "{:+.1%}"
                    }), use_container_width=True)
                    
                    if not df_entity_contrib.empty:
                        driving_entity = df_entity_contrib.iloc[0][ENTITY_COL]
                        st.info(f"**Primary Driver:** {driving_entity} contributed the largest absolute change.")
                        
                        # --- Step B: Breakdown ---
                        st.subheader("B. Sub-Metric & Worker Category Breakdown for Driving Entity")
                        
                        col_driver, col_breakdown = st.columns([1, 2])
                        
                        selected_driver = col_driver.selectbox(
                            "Analyze Contribution Breakdown for:",
                            options=df_entity_contrib[ENTITY_COL].unique(),
                            index=df_entity_contrib[ENTITY_COL].unique().tolist().index(driving_entity)
                        )
                        
                        breakdown_dim = col_driver.radio("Breakdown Dimension:", options=[WC_COL, SUBQ_COL], index=0)

                        df_breakdown = pd.DataFrame()
                        if breakdown_dim == WC_COL:
                            df_breakdown = full_attribution_df[
                                (full_attribution_df[ENTITY_COL] == selected_driver) &
                                (full_attribution_df[SUBQ_COL] == subquestion) # Filter by major subquestion
                            ]
                        elif breakdown_dim == SUBQ_COL:
                            df_breakdown = full_attribution_df[
                                (full_attribution_df[ENTITY_COL] == selected_driver) &
                                (full_attribution_df[WC_COL] == worker_cat) # Filter by major worker cat
                            ]
                            
                        df_breakdown = df_breakdown.sort_values("Absolute Change (Contribution)", ascending=False).reset_index(drop=True)
                        df_breakdown = df_breakdown[[breakdown_dim, "Absolute Change (Contribution)", "Percentage Change"]]

                        if not df_breakdown.empty:
                            col_breakdown.dataframe(df_breakdown.style.format({
                                "Absolute Change (Contribution)": "{:,.0f}",
                                "Percentage Change": "{:+.1%}"
                            }), use_container_width=True)
                            
                            breakdown_chart = px.bar(
                                df_breakdown, x="Absolute Change (Contribution)", y=breakdown_dim, 
                                orientation='h', color="Absolute Change (Contribution)",
                                color_continuous_scale=px.colors.diverging.RdYlGn,
                                title=f"{breakdown_dim} Contribution in {selected_driver}"
                            )
                            breakdown_chart.update_yaxes(categoryorder='total ascending')
                            col_breakdown.plotly_chart(breakdown_chart, use_container_width=True)
                        else:
                            col_breakdown.warning("No breakdown data available.")

    # --- RESTORED: View 2: Full Outlier Report (from your 08:31:07 App.py) ---
    elif analysis_view == 'Full Outlier Report':
        st.header("Master Outlier Report Generator")
        st.write("Scan selected workbook(s) with precise scope: dataset(s), subquestion(s), worker category(ies), and quarter window.")
        
        st.sidebar.subheader("Report Thresholds")
        abs_thresh_report = st.sidebar.slider("Absolute Change", 10, 1000, 50, 10)
        iqr_mult_report = st.sidebar.slider("IQR Sensitivity", 1.0, 3.0, 1.5, 0.1)
        yoy_thresh_report_pct = st.sidebar.slider("YoY % Threshold", 0, 100, 30, 5, format="%d%%")
        yoy_thresh_report = yoy_thresh_report_pct / 100.0
        # ADDED: MoM/QoQ Threshold for the report generator
        mom_pct_thresh_report = st.sidebar.slider("MoM/QoQ % Threshold", 0, 100, 25, 5, format="%d%%") / 100.0
        
        questions_to_scan = st.multiselect(
            "Select datasets to include in the report:",
            options=list(SHEET_MAP.keys()),
            default=list(SHEET_MAP.keys())
        )
        
        st.markdown("### Quarter Scope")
        quarter_mode = st.radio(
            "Analyze period",
            options=["Current reporting quarter", "Up to selected quarter", "Only selected quarter"],
            index=0,
            horizontal=False
        )
        selected_quarter = None
        if quarter_mode != "Current reporting quarter":
            selected_quarter = st.selectbox("Quarter:", options=[1, 2, 3, 4], index=get_reporting_quarter(current_year)-1)
        
        qmode_arg = "current" if quarter_mode == "Current reporting quarter" else \
                    ("up_to_selected" if quarter_mode == "Up to selected quarter" else "exact_selected")
        
        st.markdown("### Scope by Question / Worker Category (optional)")
        report_filters = {}
        if questions_to_scan:
            for ds in questions_to_scan:
                with st.expander(f"Filter: {ds}", expanded=False):
                    df_options = load_data(current_year, SHEET_MAP[ds])
                    if isinstance(df_options, str):
                        st.warning(df_options)
                        report_filters[ds] = {"subq": "ALL", "wc": "ALL"}
                        continue
                    
                    wc_options = sorted(list(df_options['Worker Category'].dropna().unique()))
                    wc_selected = st.multiselect(
                        f"Worker Category for {ds} (leave empty = ALL)",
                        options=wc_options, default=[], key=f"wc_{ds}"
                    )
                    wc_value = wc_selected if wc_selected else "ALL"
                    
                    if 'Subquestion' in df_options.columns:
                        subq_options = sorted(list(df_options['Subquestion'].dropna().unique()))
                        subq_selected = st.multiselect(
                            f"Subquestion for {ds} (leave empty = ALL)",
                            options=subq_options, default=[], key=f"subq_{ds}"
                        )
                        subq_value = subq_selected if subq_selected else "ALL"
                    else:
                        st.caption("No 'Subquestion' column in this dataset. Using N/A.")
                        subq_value = "ALL"
                    report_filters[ds] = {"subq": subq_value, "wc": wc_value}
        
        if st.button("🚀 Generate Full Report", use_container_width=True):
            if not questions_to_scan:
                st.warning("Please select at least one dataset to scan.")
            else:
                with st.spinner("Analyzing workbook(s)..."):
                    report_thresholds = {
                        'pct_thresh': mom_pct_thresh_report, # Use the new slider value
                        'abs_thresh': abs_thresh_report,
                        'iqr_multiplier': iqr_mult_report,
                        'yoy_thresh': yoy_thresh_report
                    }
                    years = {'current': current_year, 'prior': comparison_year}
                    final_report = generate_full_report(
                        sheet_map=SHEET_MAP,
                        years=years,
                        questions_to_scan=questions_to_scan,
                        thresholds=report_thresholds,
                        all_months=ALL_MONTHS,
                        report_filters=report_filters,
                        quarter_mode=qmode_arg,
                        selected_quarter=selected_quarter
                    )
                st.success(f"Scan complete! Found **{len(final_report)}** potential outliers.")
                if not final_report.empty:
                    st.dataframe(final_report, use_container_width=True)
                    csv = final_report.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Download Report as CSV",
                        data=csv,
                        file_name="master_outlier_report.csv",
                        mime="text/csv",
                        use_container_width=True
                    )

                    # ===== NEW: Full Report with VR merge (Monthly MoM) =====
                    try:
                        vr_df = _vr_load_variance()  # cached
                        fr = final_report.copy()

                        # Build JOIN_KEY from Period (monthly or YYYY-MMM). Skip quarterly.
                        def _fr_key(r):
                            per = r.get("Period","")
                            yr, mon = _vr_infer_year_month_from_period(per, default_year=current_year)
                            return _vr_build_join_key(
                                entity=r.get("Entity / Group",""),
                                year=yr, month=mon,
                                question=r.get("Question",""),
                                subq=r.get("Subquestion","") or "N/A",
                                wc=r.get("Worker Category","")
                            )

                        # Only consider Monthly view rows for VR (your generator uses 'Monthly' in View)
                        fr_m = fr[fr["View"].astype(str).str.upper() == "MONTHLY"].copy()
                        if not fr_m.empty:
                            fr_m["JOIN_KEY"] = fr_m.apply(_fr_key, axis=1)
                            fr_m = fr_m.merge(vr_df, on="JOIN_KEY", how="left")

                            fr_m["growth_pct_num"] = fr_m["%Growth"].apply(_vr_to_percent_number)
                            fr_m["vr_submitted"]   = fr_m["%Growth"].notna() | fr_m["Justification"].notna()
                            fr_m["breaches"]       = fr_m["growth_pct_num"].abs() >= VR_THRESHOLD_PCT
                            fr_m["Status"]         = fr_m.apply(
                                lambda r: "⌛ Pending" if not r["vr_submitted"]
                                else ("⚠️ Required justification" if r["breaches"] and _vr_is_blank_just(r["Justification"])
                                      else ("✅ Matched" if not _vr_is_blank_just(r["Justification"]) else "—")),
                                axis=1
                            )

                            st.subheader("Master Outlier Report (with VR)")
                            st.caption("Legend: ✅ Matched · ⚠️ Required justification · ⌛ Pending")
                            cols = ["Question","Entity / Group","Subquestion","Worker Category","View","Period","Value","Reason(s)","%Growth","Status","Justification"]
                            show_cols = [c for c in cols if c in fr_m.columns]
                            st.dataframe(fr_m[show_cols], use_container_width=True)

                            csv_vr = fr_m[show_cols].to_csv(index=False).encode("utf-8")
                            st.download_button(
                                label="📥 Download Report (with VR)",
                                data=csv_vr,
                                file_name="master_outlier_report_with_vr.csv",
                                mime="text/csv",
                                use_container_width=True
                            )
                    except Exception as e:
                        st.warning(f"VR merge skipped: {e}")
                    # ===== END VR for full report =====

if __name__ == "__main__":
    main()
