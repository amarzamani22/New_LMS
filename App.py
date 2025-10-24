import streamlit as st
import pandas as pd
import numpy as np
import re
import os
import plotly.graph_objects as go
import plotly.express as px
from typing import Dict, List, Tuple

# --- Global Configuration ---
SHEET_MAP = {'Q1A: Employees': 'QC_Q1A_Main', 'Q2A: Salary': 'QC_Q2A_Main', 'Q3: Hours Worked': 'QC_Q3', 'Q4: Vacancies': 'QC_Q4', 'Q5: Separations': 'QC_Q5'}
ALL_MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
ROLLUP_KEY = "All Financial Institutions"
ENTITY_COL = "Entity / Group"
SUBQ_COL = "Subquestion"
WC_COL = "Worker Category"

# --- Page Configuration and Styling ---
st.set_page_config(
    page_title="LMS Analysis Dashboard",
    page_icon="🏦",
    layout="wide"
)

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
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# --- 1. Helper Functions (Consolidated from your original App.py) ---

@st.cache_data
def load_data(year: int, sheet_name: str):
    """Loads data for a specific year and sheet with necessary cleaning."""
    file_path = f"submission/qc_workbook_{year}.xlsx"
    if not os.path.exists(file_path):
        return f"Error: File not found. Please ensure '{file_path}' exists in a 'submission' subfolder."
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=5)
        df.dropna(axis=1, how='all', inplace=True)
        rename_map = {'Q1.1': 'Q1_Total', 'Q2.1': 'Q2_Total', 'Q3.1': 'Q3_Total', 'Q4.1': 'Q4_Total'}
        df.rename(columns=rename_map, inplace=True)
        return df
    except Exception as e:
        return f"Error reading sheet '{sheet_name}' from {file_path}: {e}"

@st.cache_data
def get_reporting_quarter(year: int):
    """Reads the _About sheet from the specified year's workbook to get the quarter."""
    file_path = f"submission/qc_workbook_{year}.xlsx"
    try:
        about_df = pd.read_excel(file_path, sheet_name="_About", header=None)
        quarter_row = about_df[about_df[0] == 'Quarter']
        if not quarter_row.empty:
            q_label = str(quarter_row.iloc[0, 1]).strip().upper()
            return int(re.search(r'\d+', q_label).group())
    except Exception:
        pass
    return 4 # Default to Q4

def _row_filter(df: pd.DataFrame, entity: str, worker_cat: str, subq: str):
    """Robust row filter that tolerates missing Subquestion col."""
    cond = (df[ENTITY_COL] == entity) & (df[WC_COL] == worker_cat)
    if SUBQ_COL in df.columns and subq != "N/A":
        cond &= (df[SUBQ_COL] == subq)
    return df[cond]

# --- MODIFIED: build_quarter_series (from your original App.py logic) ---
def build_quarter_series(data_row, year, reporting_quarter):
    """
    Returns a pd.Series indexed by 'Qx YYYY' using existing Qx_Total if present,
    otherwise computes from months available. (Simplified for this full code)
    """
    if data_row.empty:
        return pd.Series(dtype=float)

    months_order = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    q_to_months = {1: months_order[0:3], 2: months_order[3:6], 3: months_order[6:9], 4: months_order[9:12]}
    series_vals, labels = [], []

    for q in range(1, reporting_quarter + 1):
        col = f'Q{q}_Total'
        q_label = f"Q{q} {year}"
        
        # 1. Prefer explicit totals
        if col in data_row.columns and not pd.isna(data_row.iloc[0][col]):
            series_vals.append(float(pd.to_numeric(data_row.iloc[0][col], errors='coerce')))
            labels.append(q_label)
            continue
            
        # 2. Compute from months if total is missing
        mlist = [m for m in q_to_months[q] if m in data_row.columns]
        if mlist:
            vals = pd.to_numeric(data_row.iloc[0][mlist], errors='coerce').astype(float)
            if not np.all(np.isnan(vals)):
                series_vals.append(float(np.nansum(vals)))
                labels.append(q_label)

    if not labels:
        return pd.Series(dtype=float)
        
    # Keep order Q1..Qn
    s = pd.Series(series_vals, index=labels, dtype=float)
    order = sorted(s.index, key=lambda lab: (int(lab.split()[1]), int(lab.split()[0][1:])))
    return s.loc[order]


# ====================================================================
# MODIFIED: find_outliers - FIXED THE KEYERROR AND ADDED FLAG DETAIL
# ====================================================================
def find_outliers(
    data_series: pd.Series, 
    prior_year_series: pd.Series, 
    mom_pct_thresh: float, 
    yoy_pct_thresh: float, 
    abs_cutoff: float, 
    iqr_multiplier: float
):
    """
    Detects outliers based on three pillars (MoM, IQR, YoY) and 
    returns a DataFrame of flagged periods and reasons.
    """
    outliers = []
    if data_series.isnull().all() or len(data_series) < 2: 
        # Return an empty DataFrame with expected index if no data to analyze
        return pd.DataFrame(columns=["Value", "Reason(s)", "MoM_Diff", "MoM_Pct"])

    q1, q3 = data_series.quantile(0.25), data_series.quantile(0.75)
    iqr = q3 - q1 if len(data_series) >= 4 else 0
    iqr_lower, iqr_upper = q1 - (iqr_multiplier * iqr), q3 + (iqr_multiplier * iqr)

    for i, (period_name, current_value) in enumerate(data_series.items()):
        reasons = []
        if pd.isna(current_value): continue
        
        abs_change, pct_change = np.nan, np.nan
        
        # --- 1. MoM Volatility + Absolute Cutoff ---
        if i > 0:
            prev_val = data_series.iloc[i-1]
            if not pd.isna(prev_val) and prev_val != 0:
                abs_change = current_value - prev_val
                pct_change = abs_change / prev_val
                
                flag = ""
                # Match Excel QC Logic: High Volatility (Red/Yellow)
                if abs(pct_change) >= mom_pct_thresh:
                    if abs(abs_change) >= abs_cutoff:
                        flag = f"High Volatility - RED ({pct_change:+.1%})"
                    elif abs(abs_change) < abs_cutoff:
                        flag = f"High Volatility - YELLOW ({pct_change:+.1%})"
                if flag: reasons.append(flag)

        # --- 2. IQR Anomaly ---
        if iqr > 0 and (current_value < iqr_lower or current_value > iqr_upper): 
            reasons.append("IQR Anomaly")
        
        # --- 3. YoY Anomaly (Simplified to use prior_year_series index) ---
        if prior_year_series is not None and period_name in prior_year_series.index:
            prior_value = prior_year_series.get(period_name)
            if not pd.isna(prior_value) and prior_value != 0:
                yoy_change = (current_value - prior_value) / prior_value
                if abs(yoy_change) > yoy_pct_thresh: 
                    reasons.append(f"YoY Anomaly ({yoy_change:+.1%})")
        
        if reasons:
            outliers.append({
                "Period": period_name, # Temporary column name
                "Value": current_value, 
                "Reason(s)": "; ".join(reasons),
                "MoM_Diff": abs_change,
                "MoM_Pct": pct_change
            })
            
    # FIXED: Check if the list is empty before creating DataFrame to avoid KeyError
    if not outliers:
        return pd.DataFrame(columns=["Value", "Reason(s)", "MoM_Diff", "MoM_Pct"])

    return pd.DataFrame(outliers).set_index("Period")


# ====================================================================
# NEW: Data Preparation for Attribution Panel
# ====================================================================
def prepare_attribution_data(
    df: pd.DataFrame, 
    period_label: str, 
    time_view: str
) -> pd.DataFrame:
    """
    Identifies the correct difference/change columns for the selected period
    and renames them for consistent attribution analysis.
    """
    # Determine the columns based on the selected period
    if time_view == 'Monthly':
        # Handles period_label like 'Mar' (from series index) or 'Q1 Mar' (if passed from other source)
        month = period_label.split(' ')[-1] 
        diff_col = f"Diff {month}"
        pct_col = f"MoM {month}"
    else: # Quarterly
        # Handles period_label like 'Q2 2025' (from series index)
        q_label = period_label.split(' ')[0] 
        diff_col = f"Diff {q_label}"
        pct_col = f"%Diff {q_label}"

    if diff_col not in df.columns:
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
    
    # --- Sidebar Controls ---
    st.sidebar.title("Analysis Controls")
    current_year = st.sidebar.selectbox("Current Year:", [2025, 2024, 2023, 2022, 2021], index=0)
    comparison_year = st.sidebar.selectbox("Comparison Year:", [2024, 2023, 2022, 2021, 2020, None], index=0)
    st.sidebar.markdown("---")
    
    analysis_view = st.sidebar.radio("Select View:", ('Interactive Analysis', 'Full Outlier Report'), label_visibility="collapsed")
    st.sidebar.markdown("---")

    reporting_quarter = get_reporting_quarter(current_year)
    months_to_analyze = ALL_MONTHS[:reporting_quarter * 3]
    st.sidebar.info(f"Analyzing **{current_year}** data up to **Q{reporting_quarter}**.")
    
    if analysis_view == 'Interactive Analysis':
        
        # --- Thresholds ---
        st.sidebar.subheader("Detection Thresholds")
        mom_pct_thresh_pct = st.sidebar.slider("MoM/QoQ % Threshold", 1, 100, 25, 1, format="%d%%") / 100.0 
        yoy_thresh_pct = st.sidebar.slider("YoY % Threshold", 1, 100, 30, 1, format="%d%%") / 100.0
        abs_cutoff = st.sidebar.slider("Absolute Cutoff (for RED flag)", 10, 1000, 50, 10)
        iqr_mult = st.sidebar.slider("IQR Sensitivity", 1.0, 3.0, 1.5, 0.1)
        st.sidebar.markdown("---")

        # --- Data Selection ---
        selected_question = st.sidebar.selectbox("Dataset (Question):", options=list(SHEET_MAP.keys()))
        time_view = st.sidebar.radio("Frequency:", ('Monthly', 'Quarterly'), horizontal=True)
        st.sidebar.markdown("---")
        
        # Load full data (includes all entities and rollups)
        df_full = load_data(current_year, SHEET_MAP[selected_question])
        df_prior_full = load_data(comparison_year, SHEET_MAP[selected_question]) if comparison_year else None
        
        if isinstance(df_full, str):
            st.error(df_full)
            return

        # --- 1. Outlier Detection (Major View) ---
        st.header("1. Outlier Detection (Major View) 📈")
        st.caption(f"Showing trend for **{ROLLUP_KEY}**")

        # Filters for the "Major View" (Top Aggregates)
        df_rollup = df_full[df_full[ENTITY_COL] == ROLLUP_KEY]
        subquestion_options = df_rollup[SUBQ_COL].unique()
        worker_cat_options = df_rollup[WC_COL].unique()
        
        col_major_subq, col_major_wc = st.columns(2)
        
        default_subq_index = np.where(subquestion_options == 'Employment = A+B(i)+B(ii)')[0][0] if 'Employment = A+B(i)+B(ii)' in subquestion_options else 0
        default_wc_index = np.where(worker_cat_options == 'Total Employment')[0][0] if 'Total Employment' in worker_cat_options else 0
        
        major_subquestion = col_major_subq.selectbox(f"Select Major Subquestion:", options=subquestion_options, index=int(default_subq_index))
        major_worker_cat = col_major_wc.selectbox(f"Select Major Worker Category:", options=worker_cat_options, index=int(default_wc_index))

        # Get the single time series for the major rollup (Current and Prior)
        data_row = _row_filter(df_full, ROLLUP_KEY, major_worker_cat, major_subquestion)
        prior_row = _row_filter(df_prior_full, ROLLUP_KEY, major_worker_cat, major_subquestion) if df_prior_full else pd.DataFrame()

        # --- Series Building & Outlier Running ---
        current_series, prior_series = pd.Series(dtype=float), None
        
        if time_view == 'Monthly':
            actual_months = [m for m in months_to_analyze if m in data_row.columns]
            current_series = pd.to_numeric(data_row[actual_months].iloc[0], errors='coerce').astype(float)
            current_series.index = actual_months
            if not prior_row.empty:
                 prior_series = pd.to_numeric(prior_row[[m for m in actual_months if m in prior_row.columns]].iloc[0], errors='coerce').astype(float)
                 prior_series.index = current_series.index
                 
        elif time_view == 'Quarterly':
            current_series = build_quarter_series(data_row, current_year, reporting_quarter)
            if not prior_row.empty:
                prior_series = build_quarter_series(prior_row, comparison_year, get_reporting_quarter(comparison_year) if comparison_year else 4)
                # Align prior series index keys (e.g., 'Q1 2024') with current series index keys ('Q1 2025') for YoY
                prior_series.index = current_series.index.map(lambda x: f"{x.split()[0]} {int(x.split()[1])-1}")
                prior_series = prior_series.loc[prior_series.index.intersection(current_series.index.map(lambda x: f"{x.split()[0]} {int(x.split()[1])-1}"))] # Keep only matching quarters

        # Run the detection engine
        outlier_df = find_outliers(
            current_series, prior_series, 
            mom_pct_thresh_pct, yoy_thresh_pct, 
            abs_cutoff, iqr_mult
        )

        # --- Plotting the Detection Chart ---
        if not current_series.empty and not current_series.isnull().all():
            x = list(current_series.index)
            y = current_series.values.astype(float)
            fig = go.Figure()

            # Current year trend line
            fig.add_trace(go.Scatter(x=x, y=y, mode='lines+markers', name=f'{current_year} Trend'))
            
            # MoM ± Threshold band
            if len(y) >= 2:
                upper = [None] + [y[i-1] * (1 + mom_pct_thresh_pct) for i in range(1, len(y))]
                lower = [None] + [y[i-1] * (1 - mom_pct_thresh_pct) for i in range(1, len(y))]
                fig.add_trace(go.Scatter(x=x, y=lower, mode='lines', line=dict(width=0), showlegend=False, hoverinfo='skip'))
                fig.add_trace(go.Scatter(x=x, y=upper, mode='lines', line=dict(width=0), fill='tonexty', name=f'MoM/QoQ Threshold', opacity=0.15))
                
            # Outlier markers
            if not outlier_df.empty:
                o_periods = outlier_df.index
                o_values = outlier_df["Value"]
                o_reasons = outlier_df["Reason(s)"]
                colors = ['red' if 'RED' in r else 'orange' for r in o_reasons]

                fig.add_trace(go.Scatter(
                    x=o_periods, y=o_values, mode='markers', name='Detected Outlier',
                    marker=dict(symbol='x', size=14, color=colors, line=dict(width=2, color='darkred')),
                    hovertemplate='Period: %{x}<br>Value: %{y:,.0f}<br>Reason: %{customdata}<extra></extra>',
                    customdata=o_reasons
                ))

            fig.update_layout(title=f"Trend for {ROLLUP_KEY}", hovermode='x unified')
            fig.update_xaxes(tickangle=45)
            st.plotly_chart(fig, use_container_width=True)
            
        else:
            st.warning("No time series data available for plotting.")
            
        # --- 2. Outlier Attribution Panel ---
        st.header("2. Outlier Attribution Panel 🔍")
        
        if not outlier_df.empty:
            outlier_periods = outlier_df.index.tolist()
            
            # Interactive Trigger: Select the period to drill down on
            selected_outlier_period = st.selectbox(
                "Select Outlier Period to Analyze (Attribution):", 
                options=outlier_periods
            )
            
            # --- Prepare the Full Attribution Data for the Selected Period ---
            full_attribution_df = prepare_attribution_data(
                df=df_full,
                period_label=selected_outlier_period,
                time_view=time_view
            )

            if full_attribution_df.empty:
                st.error("Could not find the necessary 'Diff' (Difference) column in the data for attribution.")
            else:
                
                # --- Step A: Entity/Group Contribution Table ---
                st.subheader("A. Entity / Group Contribution")
                st.caption(f"Contribution of **ALL** entities/rollups to the change in **{selected_outlier_period}**")
                
                # Filter A: Match Major View's Subquestion and Worker Category
                df_entity_contrib = full_attribution_df[
                    (full_attribution_df[SUBQ_COL] == major_subquestion) &
                    (full_attribution_df[WC_COL] == major_worker_cat)
                ].copy()
                
                # Sort by Absolute Change
                df_entity_contrib = df_entity_contrib.sort_values("Absolute Change (Contribution)", ascending=False).reset_index(drop=True)
                
                # Display table and determine driving entity
                st.dataframe(df_entity_contrib.style.format({
                    "Absolute Change (Contribution)": "{:,.0f}",
                    "Percentage Change": "{:+.1%}"
                }), use_container_width=True)
                
                if not df_entity_contrib.empty:
                    driving_entity = df_entity_contrib.iloc[0][ENTITY_COL]
                    st.info(f"**Primary Driver:** {driving_entity} contributed the largest absolute change.")
                    
                    # --- Step B: Sub-Metric and Category Breakdown ---
                    st.subheader("B. Sub-Metric & Worker Category Breakdown for Driving Entity")
                    st.caption(f"Analyze the specific dimensions that caused the change in **{driving_entity}**.")

                    col_driver, col_breakdown = st.columns([1, 2])
                    
                    # Interactive selection of the entity to break down
                    selected_driver = col_driver.selectbox(
                        "Analyze Contribution Breakdown for:",
                        options=df_entity_contrib[ENTITY_COL].unique(),
                        index=df_entity_contrib[ENTITY_COL].unique().tolist().index(driving_entity)
                    )
                    
                    breakdown_dim = col_driver.radio("Breakdown Dimension:", options=[WC_COL, SUBQ_COL], index=0)

                    df_breakdown = None
                    if breakdown_dim == WC_COL:
                        # Filter for the major subquestion, but all worker categories
                        df_breakdown = full_attribution_df[
                            (full_attribution_df[ENTITY_COL] == selected_driver) &
                            (full_attribution_df[SUBQ_COL] == major_subquestion)
                        ]
                        
                    elif breakdown_dim == SUBQ_COL:
                        # Filter for the major worker category, but all subquestions
                        df_breakdown = full_attribution_df[
                            (full_attribution_df[ENTITY_COL] == selected_driver) &
                            (full_attribution_df[WC_COL] == major_worker_cat)
                        ]
                        
                    df_breakdown = df_breakdown.sort_values("Absolute Change (Contribution)", ascending=False).reset_index(drop=True)
                    df_breakdown = df_breakdown[[breakdown_dim, "Absolute Change (Contribution)", "Percentage Change"]]

                    if not df_breakdown.empty:
                        # Display data table
                        col_breakdown.dataframe(df_breakdown.style.format({
                            "Absolute Change (Contribution)": "{:,.0f}",
                            "Percentage Change": "{:+.1%}"
                        }), use_container_width=True)
                        
                        # Display bar chart
                        breakdown_chart = px.bar(
                            df_breakdown, 
                            x="Absolute Change (Contribution)", 
                            y=breakdown_dim, 
                            orientation='h', 
                            color="Absolute Change (Contribution)",
                            color_continuous_scale=px.colors.diverging.RdYlGn,
                            title=f"{breakdown_dim} Contribution in {selected_driver}"
                        )
                        breakdown_chart.update_yaxes(categoryorder='total ascending')
                        col_breakdown.plotly_chart(breakdown_chart, use_container_width=True)
                        
                    else:
                        col_breakdown.warning(f"No breakdown data available for {selected_driver} on dimension {breakdown_dim}.")

        else:
            st.success("✅ No significant outliers detected in the major view.")
            
    # --- View 2: Full Outlier Report (Placeholder) ---
    elif analysis_view == 'Full Outlier Report':
        st.header("Master Outlier Report Generator")
        st.warning("Please integrate your existing Full Outlier Report logic here.")

if __name__ == "__main__":
    main()
