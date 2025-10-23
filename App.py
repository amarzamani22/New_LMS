# app.py
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import re
import os
import plotly.graph_objects as go

# --- Page Configuration ---
st.set_page_config(
    page_title="LMS Analysis Dashboard",
    page_icon="🏦",
    layout="wide"
)

# --- Minimal CSS ---
st.markdown("""
<style>
    .stApp { background-color: #F0F2F6; }
    .stMetric { border-radius: 10px; padding: 20px; background-color: #FFFFFF;
                border: 1px solid #E0E0E0; box-shadow: 0 4px 6px rgba(0,0,0,0.04); }
    .stButton>button { border-radius: 8px; font-weight: 600; }
    .muted { color:#6b7280; }
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------
# 1) Helper Functions (unchanged from your code unless marked)
# ------------------------------------------------------------
@st.cache_data
def load_data(year, sheet_name):
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
def get_reporting_quarter(year):
    file_path = f"submission/qc_workbook_{year}.xlsx"
    try:
        about_df = pd.read_excel(file_path, sheet_name="_About", header=None)
        quarter_row = about_df[about_df[0] == 'Quarter']
        if not quarter_row.empty:
            return int(re.search(r'\d+', str(quarter_row.iloc[0, 1])).group())
    except Exception:
        pass
    return 4 # default

def _row_filter(df, entity, worker_cat, subq):
    """Robust row filter that tolerates missing Subquestion col."""
    cond = (df['Entity / Group'] == entity) & (df['Worker Category'] == worker_cat)
    if 'Subquestion' in df.columns:
        cond &= (df['Subquestion'] == subq)
    return df[cond]

def find_outliers(data_series, prior_year_series, pct_thresh, abs_thresh, iqr_multiplier, yoy_thresh):
    outliers = []
    if data_series.isnull().all() or len(data_series) < 2: return pd.DataFrame()

    q1, q3 = data_series.quantile(0.25), data_series.quantile(0.75)
    iqr = q3 - q1 if len(data_series) >= 4 else 0
    iqr_lower_bound, iqr_upper_bound = q1 - (iqr_multiplier * iqr), q3 + (iqr_multiplier * iqr)

    for i, (period_name, current_value) in enumerate(data_series.items()):
        reasons = []
        if pd.isna(current_value): continue

        if i > 0:
            prev_val = data_series.iloc[i-1]
            if not pd.isna(prev_val) and prev_val != 0:
                abs_change = current_value - prev_val
                pct_change = abs_change / prev_val
                if abs(pct_change) > pct_thresh and abs(abs_change) > abs_thresh:
                    reasons.append(f"High Volatility ({pct_change:+.1%})")
        if iqr > 0 and (current_value < iqr_lower_bound or current_value > iqr_upper_bound):
            reasons.append("IQR Anomaly")
        if prior_year_series is not None and period_name in prior_year_series.index:
            prior_value = prior_year_series[period_name]
            if not pd.isna(prior_value) and prior_value != 0:
                yoy_change = (current_value - prior_value) / prior_value
                if abs(yoy_change) > yoy_thresh:
                    reasons.append(f"YoY Anomaly ({yoy_change:+.1%})")
        
        if reasons:
            outliers.append({"Period": period_name, "Value": f"{current_value:,.2f}", "Reason(s)": ", ".join(reasons)})
            
    return pd.DataFrame(outliers)

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

def build_quarter_series(data_row, year, reporting_quarter):
    """
    Returns a pd.Series indexed by 'Qx YYYY' using existing Qx_Total if present,
    otherwise computes from months available.
    """
    if data_row.empty:
        return pd.Series(dtype=float)

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

    months_order = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    present_quarters = {int(l.split()[0][1:]) for l in labels}
    for q in range(1, reporting_quarter + 1):
        if q in present_quarters:
            continue
        q_to_months = {1: months_order[0:3], 2: months_order[3:6], 3: months_order[6:9], 4: months_order[9:12]}
        mlist = [m for m in q_to_months[q] if m in data_row.columns]
        if mlist:
            vals = pd.to_numeric(data_row.iloc[0][mlist], errors='coerce').astype(float)
            if not np.all(np.isnan(vals)):
                series_vals.append(float(np.nansum(vals)))
                labels.append(f"Q{q} {year}")

    if not labels:
        return pd.Series(dtype=float)
    sort_idx = np.argsort([int(l.split()[0][1:]) for l in labels])
    labels = [labels[i] for i in sort_idx]
    series_vals = [series_vals[i] for i in sort_idx]

    s = pd.Series(series_vals, index=labels, dtype=float)
    return s

def generate_full_report(
   sheet_map,
   years,
   questions_to_scan,
   thresholds,
   all_months,
   report_filters=None,
   quarter_mode="current",
   selected_quarter=None
):
   def months_for_quarter(q: int):
       q_map = {1: all_months[0:3], 2: all_months[3:6], 3: all_months[6:9], 4: all_months[9:12]}
       return q_map.get(int(q), [])
   if quarter_mode == "current":
       rq = get_reporting_quarter(years['current'])
       months_to_use  = all_months[: rq * 3]
       focus_months   = months_for_quarter(rq)
   elif quarter_mode == "up_to_selected" and selected_quarter:
       months_to_use  = all_months[: int(selected_quarter) * 3]
       focus_months   = months_to_use
   elif quarter_mode == "exact_selected" and selected_quarter:
       months_to_use  = all_months[: int(selected_quarter) * 3]
       focus_months   = months_for_quarter(int(selected_quarter))
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
       if isinstance(df_current, str):
           continue
       actual_months = [m for m in all_months[: get_reporting_quarter(years['current']) * 3] if m in df_current.columns]
       if not actual_months:
           continue
       f_cfg = (report_filters or {}).get(q_name, {"subq": "ALL", "wc": "ALL"})
       subq_filter = f_cfg.get("subq", "ALL")
       wc_filter   = f_cfg.get("wc", "ALL")
       for _, row in df_current.iterrows():
           entity = row['Entity / Group']
           wc     = row['Worker Category']
           subq   = row['Subquestion'] if 'Subquestion' in df_current.columns else 'N/A'
           if subq_filter != "ALL" and subq not in subq_filter:
               continue
           if wc_filter   != "ALL" and wc   not in wc_filter:
               continue
           monthly_series = pd.to_numeric(row[actual_months], errors='coerce').astype(float)
           prior_series = None
           if not isinstance(df_prior, str) and df_prior is not None:
               prior_row = _row_filter(df_prior, entity, wc, subq)
               if not prior_row.empty:
                   prior_series = pd.to_numeric(prior_row[actual_months].iloc[0], errors='coerce').astype(float)
           outliers_monthly = find_outliers(monthly_series, prior_series, **thresholds)
           for _, o_row in outliers_monthly.iterrows():
            if o_row['Period'] not in focus_set:
                continue
            master_outlier_list.append([
                q_name, entity, subq, wc,
                'Monthly', o_row['Period'], o_row['Value'], o_row['Reason(s)']
            ])
   progress_bar.progress(1.0, text="Scan Complete!")
   return pd.DataFrame(
       master_outlier_list,
       columns=['Question', 'Entity / Group', 'Subquestion', 'Worker Category', 'View', 'Period', 'Value', 'Reason(s)']
   )

# =========================
# >>> NEW: utils for quarter and contributors
# =========================
ALL_MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

def months_for_quarter(q: int):
    q = int(q)
    return ALL_MONTHS[(q-1)*3 : q*3]

def is_all_fi(name: str) -> bool:
    return str(name).strip() in ("All FI", "All Financial Institutions", "All RE")

@st.cache_data(show_spinner=False)
def _contributors_for_period(
    period_label,
    *,
    df_current, df_prior,
    current_year, comparison_year,
    subquestion, worker_cat,
    time_view, actual_months,
    reporting_quarter,
    abs_thresh, iqr_mult, yoy_thresh
):
    """
    Returns a list of strings like 'Maybank — IQR Anomaly; YoY (+32.4%)'
    for REs flagged (by your 3 pillars) at the same period.
    """
    scope = df_current.copy()
    if 'Subquestion' in scope.columns:
        scope = scope[scope['Subquestion'] == subquestion]
    scope = scope[scope['Worker Category'] == worker_cat]

    # exclude roll-ups
    EXCLUDE = {"All FI", "All Financial Institutions", "All RE",
               "Banking Institution", "Insurance/Takaful", "DFI"}
    re_names = [x for x in scope['Entity / Group'].dropna().unique().tolist() if x not in EXCLUDE]

    contrib = []
    for re_name in re_names:
        row_re = _row_filter(df_current, re_name, worker_cat, subquestion)
        if row_re.empty: 
            continue

        # build series same as main plot
        if time_view == "Monthly":
            if not actual_months:
                continue
            ser = pd.to_numeric(row_re[actual_months].iloc[0], errors='coerce').astype(float)
            prior_ser = None
            if not isinstance(df_prior, str) and df_prior is not None:
                prior_row = _row_filter(df_prior, re_name, worker_cat, subquestion)
                if not prior_row.empty and all(m in prior_row.columns for m in actual_months):
                    prior_ser = pd.to_numeric(prior_row[actual_months].iloc[0], errors='coerce').astype(float)
        else:  # Quarterly
            ser = build_quarter_series(row_re, current_year, reporting_quarter)
            prior_ser = None
            if not isinstance(df_prior, str) and df_prior is not None:
                prior_row = _row_filter(df_prior, re_name, worker_cat, subquestion)
                if not prior_row.empty:
                    prior_ser = build_quarter_series(prior_row, comparison_year, get_reporting_quarter(comparison_year))

        if ser is None or ser.empty:
            continue

        odf = find_outliers(ser, prior_ser, 0.25, abs_thresh, iqr_mult, yoy_thresh)
        if odf.empty or period_label not in set(odf['Period']):
            continue

        reason = odf.set_index('Period').loc[period_label, 'Reason(s)']
        contrib.append(f"{re_name} — {reason}")

    contrib.sort()
    return contrib[:5]  # keep hover compact

# ---------------------------------------
# 2) Main Dashboard Interface (unchanged)
# ---------------------------------------
st.title("🏦 LMS Analysis Dashboard")

SHEET_MAP = {
    'Q1A: Employees': 'QC_Q1A_Main',
    'Q2A: Salary': 'QC_Q2A_Main',
    'Q3: Hours Worked': 'QC_Q3',
    'Q4: Vacancies': 'QC_Q4',
    'Q5: Separations': 'QC_Q5'
}

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
    abs_thresh = st.sidebar.slider("Absolute Change", 10, 1000, 50, 10)
    iqr_mult = st.sidebar.slider("IQR Sensitivity", 1.0, 3.0, 1.5, 0.1)
    yoy_thresh_pct = st.sidebar.slider("YoY % Threshold", 0, 100, 30, 5, format="%d%%")
    yoy_thresh = yoy_thresh_pct / 100.0
    
    # Always allow multi-year via the existing controls (kept as-is)
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

    df_current = load_data(current_year, SHEET_MAP[selected_question])
    df_prior = load_data(comparison_year, SHEET_MAP[selected_question]) if comparison_year else None

    if isinstance(df_current, str):
        st.error(df_current)
    else:
        entity = st.selectbox("Select an Entity / Group:", options=df_current['Entity / Group'].unique())
        
        if 'Subquestion' in df_current.columns and df_current[df_current['Entity / Group'] == entity]['Subquestion'].nunique() > 1:
            subquestion = st.selectbox("Select a Subquestion:", options=df_current[df_current['Entity / Group'] == entity]['Subquestion'].unique())
            worker_cat = st.selectbox("Select a Worker Category:", options=df_current[(df_current['Entity / Group'] == entity) & (df_current['Subquestion'] == subquestion)]['Worker Category'].unique())
            data_row = _row_filter(df_current, entity, worker_cat, subquestion)
        else:
            subquestion = "N/A"
            worker_cat = st.selectbox("Select a Worker Category:", options=df_current[df_current['Entity / Group'] == entity]['Worker Category'].unique())
            data_row = _row_filter(df_current, entity, worker_cat, subquestion)

        st.header(f"Analysis for: {entity}"); st.caption(f"Category: {worker_cat}")
        
        current_series, prior_series = pd.Series(dtype=float), None

        if time_view == 'Monthly':
            actual_months = [m for m in months_to_analyze if m in data_row.columns]
            if actual_months:
                current_series = pd.to_numeric(data_row[actual_months].iloc[0], errors='coerce').astype(float)
                current_series.index = actual_months

            if not isinstance(df_prior, str) and df_prior is not None:
                prior_row = _row_filter(df_prior, entity, worker_cat, subquestion)
                if not prior_row.empty and all(m in prior_row.columns for m in actual_months):
                    prior_series = pd.to_numeric(prior_row[actual_months].iloc[0], errors='coerce').astype(float)

            title = "Monthly Trend"

            if enable_multi:
                # rebuild series across years (kept from your code)
                vals, idx = [], []
                for yr in range(int(start_year), int(end_year)+1):
                    df_y = load_data(yr, SHEET_MAP[selected_question])
                    if isinstance(df_y, str): continue
                    row_y = _row_filter(df_y, entity, worker_cat, subquestion)
                    if row_y.empty: continue
                    upto_q_y = get_reporting_quarter(yr)
                    months_y = [m for m in ALL_MONTHS[:upto_q_y * 3] if m in row_y.columns]
                    if not months_y: continue
                    ser_y = pd.to_numeric(row_y[months_y].iloc[0], errors='coerce').astype(float)
                    for m in months_y:
                        label = f"{yr}-{m}"
                        if label in exclude_months: continue
                        v = ser_y.get(m, np.nan)
                        vals.append(v); idx.append(label)
                if idx:
                    current_series = pd.Series(vals, index=idx, dtype=float)
                    # build prior for YoY
                    p_vals, p_idx = [], []
                    for label in current_series.index:
                        try:
                            yr_str, mon = label.split('-')
                            prev_label = f"{int(yr_str)-1}-{mon}"
                            if prev_label in current_series.index and not pd.isna(current_series.loc[prev_label]):
                                p_vals.append(float(current_series.loc[prev_label])); p_idx.append(label)
                        except Exception:
                            pass
                    prior_series = pd.Series(p_vals, index=p_idx, dtype=float) if p_idx else None

        else:  # Quarterly
            current_series = build_quarter_series(data_row, current_year, reporting_quarter)
            title = "Quarterly Trend"

            if not isinstance(df_prior, str) and df_prior is not None:
                prior_row = _row_filter(df_prior, entity, worker_cat, subquestion)
                if not prior_row.empty:
                    prior_rq = get_reporting_quarter(comparison_year)
                    prior_series = build_quarter_series(prior_row, comparison_year, prior_rq)

            if enable_multi:
                parts = []
                for yr in range(int(start_year), int(end_year)+1):
                    df_y = load_data(yr, SHEET_MAP[selected_question])
                    if isinstance(df_y, str): continue
                    row_y = _row_filter(df_y, entity, worker_cat, subquestion)
                    if row_y.empty: continue
                    rq = get_reporting_quarter(yr)
                    s_y = build_quarter_series(row_y, yr, rq)
                    if s_y.empty: continue
                    keep = [lab for lab in s_y.index if lab not in set(exclude_quarters)]
                    parts.append(s_y.loc[keep] if keep else pd.Series(dtype=float))
                if parts:
                    current_series = pd.concat(parts)
                    order = sorted(current_series.index, key=lambda lab: (int(lab.split()[1]), int(lab.split()[0][1:])))
                    current_series = current_series.loc[order]
                    p_vals, p_idx = [], []
                    for label in current_series.index:
                        try:
                            q, yr_str = label.split()
                            prev_label = f"{q} {int(yr_str)-1}"
                            if prev_label in current_series.index and not pd.isna(current_series.loc[prev_label]):
                                p_vals.append(float(current_series.loc[prev_label])); p_idx.append(label)
                        except Exception:
                            pass
                    prior_series = pd.Series(p_vals, index=p_idx, dtype=float) if p_idx else None

        if current_series.empty or current_series.isnull().all():
            st.warning("No data found for the selected filters or period.")
        else:
            # Outliers (your engine)
            outlier_df = find_outliers(current_series, prior_series, 0.25, abs_thresh, iqr_mult, yoy_thresh)

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Latest Value", f"{current_series.iloc[-1]:,.0f}")
            col2.metric("Period Average", f"{current_series.mean():,.0f}")
            col3.metric("Period High", f"{np.nanmax(current_series):,.0f}")
            col4.metric("True Outliers", len(outlier_df))
            
            # --- Plotly interactive chart ---
            x = list(current_series.index)
            y = current_series.values.astype(float)

            fig = go.Figure()

            fig.add_trace(go.Scatter(
                x=x, y=y, mode='lines+markers', name='Trend',
                hovertemplate='Period: %{x}<br>Value: %{y:,.0f}<extra></extra>'
            ))

            # MoM ±25% band (only useful for monthly sequences with >=2 points)
            if time_view == 'Monthly' and len(y) >= 2:
                upper = [None] + [y[i-1] * 1.25 for i in range(1, len(y))]
                lower = [None] + [y[i-1] * 0.75 for i in range(1, len(y))]
                fig.add_trace(go.Scatter(x=x, y=lower, mode='lines', line=dict(width=0),
                                         name='MoM −25% Range', showlegend=False, hoverinfo='skip'))
                fig.add_trace(go.Scatter(x=x, y=upper, mode='lines', line=dict(width=0),
                                         fill='tonexty', name='MoM ±25% Range', opacity=0.15,
                                         hovertemplate='Prev×1.25: %{y:,.0f}<extra></extra>'))

            # >>> NEW: Outlier markers with contributor hover (for All FI)
            if not outlier_df.empty:
                o_x = [p for p in outlier_df['Period'] if p in current_series.index]
                o_y = [current_series[p] for p in o_x]

                # build custom hover with contributors when All FI is selected
                o_custom = []
                for p in o_x:
                    base_reason = outlier_df.set_index('Period').loc[p, 'Reason(s)']
                    if is_all_fi(entity):
                        contrib = _contributors_for_period(
                            p,
                            df_current=df_current, df_prior=df_prior,
                            current_year=current_year, comparison_year=comparison_year,
                            subquestion=subquestion, worker_cat=worker_cat,
                            time_view=time_view, actual_months=x if time_view=='Monthly' else [],
                            reporting_quarter=reporting_quarter,
                            abs_thresh=abs_thresh, iqr_mult=iqr_mult, yoy_thresh=yoy_thresh
                        )
                        if contrib:
                            contrib_text = "<br>".join(contrib)
                            o_custom.append(f"{base_reason}<br><b>Contributors:</b><br>{contrib_text}")
                        else:
                            o_custom.append(f"{base_reason}<br><b>Contributors:</b> None flagged")
                    else:
                        o_custom.append(base_reason)

                fig.add_trace(go.Scatter(
                    x=o_x, y=o_y, mode='markers', name='True Outlier',
                    marker=dict(symbol='x', size=14),
                    hovertemplate='Period: %{x}<br>Value: %{y:,.0f}<br>%{customdata}<extra></extra>',
                    customdata=o_custom
                ))

            fig.update_layout(
                title=dict(
                    text=(title + (" — Multi-year" if enable_multi else "")),
                    font=dict(size=16),
                    x=0.5,
                    xanchor='center'
                ),
                hovermode='x unified',
                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='left', x=0),
                margin=dict(l=10, r=10, t=70, b=10),
            )
            fig.update_xaxes(tickangle=45)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": True})

            # Context caption for active range/exclusions
            if enable_multi:
                if time_view == "Monthly":
                    st.caption(f"Range: {int(start_year)}–{int(end_year)} (monthly). Excluded: {', '.join(exclude_months) if exclude_months else 'None'}")
                else:
                    st.caption(f"Range: {int(start_year)}–{int(end_year)} (quarterly). Excluded: {', '.join(exclude_quarters) if exclude_quarters else 'None'}")
            
            if not outlier_df.empty:
                st.error("🚨 True Outlier(s) Detected!")
                st.dataframe(outlier_df, use_container_width=True)
            else:
                st.success("✅ No significant outliers detected.")
            
            # >>> NEW: Contributor table (only when entity == All FI) scoped to current reporting quarter
            if is_all_fi(entity):
                st.subheader("🔎 Contributors in Current Reporting Quarter")
                # focus months/quarter
                focus_months = months_for_quarter(reporting_quarter)
                focus_set = set(focus_months) if time_view == "Monthly" else {f"Q{reporting_quarter} {current_year}"}

                # build RE universe under same subquestion/worker
                scope = df_current.copy()
                if 'Subquestion' in scope.columns:
                    scope = scope[scope['Subquestion'] == subquestion]
                scope = scope[scope['Worker Category'] == worker_cat]

                # Exclude roll-ups
                EXCLUDE = {"All FI", "All Financial Institutions", "All RE",
                           "Banking Institution", "Insurance/Takaful", "DFI"}
                re_names = [x for x in scope['Entity / Group'].dropna().unique().tolist() if x not in EXCLUDE]

                rows = []
                for re_name in re_names:
                    row_re = _row_filter(df_current, re_name, worker_cat, subquestion)
                    if row_re.empty: continue

                    if time_view == "Monthly":
                        months_cols = [m for m in focus_months if m in row_re.columns]  # only quarter months
                        if not months_cols: continue
                        ser = pd.to_numeric(row_re[months_cols].iloc[0], errors='coerce').astype(float)
                        ser.index = months_cols
                        # Build prior aligned on same months if available
                        prior_ser = None
                        if not isinstance(df_prior, str) and df_prior is not None:
                            prior_row = _row_filter(df_prior, re_name, worker_cat, subquestion)
                            if not prior_row.empty and all(m in prior_row.columns for m in months_cols):
                                prior_ser = pd.to_numeric(prior_row[months_cols].iloc[0], errors='coerce').astype(float)
                    else:
                        ser = build_quarter_series(row_re, current_year, reporting_quarter)
                        ser = ser.loc[[f"Q{reporting_quarter} {current_year}"]] if f"Q{reporting_quarter} {current_year}" in ser.index else pd.Series(dtype=float)
                        prior_ser = None
                        if not isinstance(df_prior, str) and df_prior is not None:
                            prior_row = _row_filter(df_prior, re_name, worker_cat, subquestion)
                            if not prior_row.empty:
                                pser_full = build_quarter_series(prior_row, comparison_year, get_reporting_quarter(comparison_year))
                                prior_ser = pser_full.loc[[f"Q{reporting_quarter} {current_year}"]] if f"Q{reporting_quarter} {current_year}" in pser_full.index else None

                    if ser is None or ser.empty: 
                        continue

                    odf = find_outliers(ser, prior_ser, 0.25, abs_thresh, iqr_mult, yoy_thresh)
                    if odf.empty: 
                        continue

                    for _, r in odf.iterrows():
                        if r["Period"] not in focus_set:
                            continue
                        rows.append({
                            "Entity": re_name,
                            "Flagged Period": r["Period"],
                            "Reason(s)": r["Reason(s)"]
                        })

                contrib_df = pd.DataFrame(rows).drop_duplicates()
                if not contrib_df.empty:
                    st.dataframe(contrib_df.sort_values(["Entity","Flagged Period"]), use_container_width=True)
                else:
                    st.info("No RE flagged within the current reporting quarter under All FI.")

            # --- Raw Data Section (unchanged) ---
            with st.expander("Show Raw Data for Comparison"):
                st.subheader(f"Data for {current_year}")
                st.dataframe(data_row)
                
                if comparison_year:
                    st.subheader(f"Data for {comparison_year}")
                    if not isinstance(df_prior, str) and df_prior is not None:
                        prior_row_display = _row_filter(df_prior, entity, worker_cat, subquestion)
                        if not prior_row_display.empty:
                            st.dataframe(prior_row_display)
                        else:
                            st.warning(f"No matching data found for this selection in the {comparison_year} workbook.")
                    else:
                        st.error(f"Could not load data for {comparison_year}. Please check the file.")

# --- View 2: Full Outlier Report (unchanged) ---
elif analysis_view == 'Full Outlier Report':
   st.header("Master Outlier Report Generator")
   st.write("Scan selected workbook(s) with precise scope: dataset(s), subquestion(s), worker category(ies), and quarter window.")
   st.sidebar.subheader("Report Thresholds")
   abs_thresh_report = st.sidebar.slider("Absolute Change", 10, 1000, 50, 10)
   iqr_mult_report = st.sidebar.slider("IQR Sensitivity", 1.0, 3.0, 1.5, 0.1)
   yoy_thresh_report_pct = st.sidebar.slider("YoY % Threshold", 0, 100, 30, 5, format="%d%%")
   yoy_thresh_report = yoy_thresh_report_pct / 100.0

   SHEET_MAP = {
       'Q1A: Employees': 'QC_Q1A_Main',
       'Q2A: Salary': 'QC_Q2A_Main',
       'Q3: Hours Worked': 'QC_Q3',
       'Q4: Vacancies': 'QC_Q4',
       'Q5: Separations': 'QC_Q5'
   }
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
               wc_selected = st.multiselect("Worker Category (leave empty = ALL)", options=wc_options, default=[])
               wc_value = wc_selected if wc_selected else "ALL"
               if 'Subquestion' in df_options.columns:
                   subq_options = sorted(list(df_options['Subquestion'].dropna().unique()))
                   subq_selected = st.multiselect("Subquestion (leave empty = ALL)", options=subq_options, default=[])
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
                   'pct_thresh': 0.25,
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
