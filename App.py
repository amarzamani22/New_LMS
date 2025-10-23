# app.py — RLMS QC Dashboard (Outlier-first redesign)

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import re
import os
import plotly.graph_objects as go

# === Page Config ===
st.set_page_config(page_title="LMS Analysis Dashboard", page_icon="🏦", layout="wide")

# === Minimal CSS ===
st.markdown("""
<style>
  .stApp { background-color: #F7F9FC; }
  .pill { display:inline-block; padding:4px 8px; border-radius:12px; font-size:12px; margin-right:6px; }
  .pill-thr { background:#FFF4E5; border:1px solid #FFD9A0; }
  .pill-iqr { background:#EAF6FF; border:1px solid #B5E0FF; }
  .pill-yoy { background:#EAFCEF; border:1px solid #BFE5C8; }
  .flag-yes { color:#B00020; font-weight:600; }
  .flag-no { color:#20804E; font-weight:600; }
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------------------------------
# 1) Your existing helpers (UNCHANGED)
# -------------------------------------------------------------------------------------
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
    return 4

def _row_filter(df, entity, worker_cat, subq):
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
                abs_change = current_value - prev_val; pct_change = abs_change / prev_val
                if abs(pct_change) > pct_thresh and abs(abs_change) > abs_thresh: reasons.append(f"High Volatility ({pct_change:+.1%})")
        if iqr > 0 and (current_value < iqr_lower_bound or current_value > iqr_upper_bound): reasons.append("IQR Anomaly")
        if prior_year_series is not None and period_name in prior_year_series.index:
            prior_value = prior_year_series[period_name]
            if not pd.isna(prior_value) and prior_value != 0:
                yoy_change = (current_value - prior_value) / prior_value
                if abs(yoy_change) > yoy_thresh: reasons.append(f"YoY Anomaly ({yoy_change:+.1%})")
        if reasons:
            outliers.append({"Period": period_name, "Value": f"{current_value:,.2f}", "Reason(s)": ", ".join(reasons)})
    return pd.DataFrame(outliers)

def _quarter_labels(year, upto_q):
    return [f"Q{i} {year}" for i in range(1, upto_q + 1)]

def _compute_quarter_totals_from_months(row, upto_q, months_order):
    q_map = {1: months_order[0:3], 2: months_order[3:6], 3: months_order[6:9], 4: months_order[9:12]}
    vals, labels = [], []
    for q in range(1, upto_q + 1):
        mlist = [m for m in q_map[q] if m in row.index]
        if mlist:
            vals.append(row[mlist].astype(float).sum())
            labels.append(f"Q{q}")
    return labels, vals

def build_quarter_series(data_row, year, reporting_quarter):
    if data_row.empty:
        return pd.Series(dtype=float)
    q_cols = [f'Q{i}_Total' for i in range(1, reporting_quarter + 1)]
    have_any_qtot = any(col in data_row.columns for col in q_cols)
    series_vals, labels = [], []
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
        if q in present_quarters: continue
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

def _year_range(start_year: int, end_year: int):
    return [y for y in range(int(start_year), int(end_year) + 1)]

def build_multi_year_monthly_series(entity, worker_cat, subq, selected_question, start_year, end_year, exclusions, all_months, sheet_map):
    vals, idx = [], []
    for yr in _year_range(start_year, end_year):
        df_y = load_data(yr, sheet_map[selected_question])
        if isinstance(df_y, str): continue
        row_y = _row_filter(df_y, entity, worker_cat, subq)
        if row_y.empty: continue
        upto_q_y = get_reporting_quarter(yr)
        months_y = [m for m in all_months[:upto_q_y * 3] if m in row_y.columns]
        if not months_y: continue
        ser_y = pd.to_numeric(row_y[months_y].iloc[0], errors='coerce').astype(float)
        for m in months_y:
            label = f"{yr}-{m}"
            if label in exclusions: continue
            v = ser_y.get(m, np.nan)
            vals.append(v); idx.append(label)
    if not idx: return pd.Series(dtype=float), None
    s = pd.Series(vals, index=idx, dtype=float)
    prior_vals, prior_idx = [], []
    for label in s.index:
        try:
            yr_str, mon = label.split('-')
            prev_label = f"{int(yr_str)-1}-{mon}"
            if prev_label in s.index and not pd.isna(s.loc[prev_label]):
                prior_vals.append(float(s.loc[prev_label])); prior_idx.append(label)
        except Exception:
            continue
    prior = pd.Series(prior_vals, index=prior_idx, dtype=float) if prior_idx else None
    return s, prior

def build_multi_year_quarterly_series(entity, worker_cat, subq, selected_question, start_year, end_year, exclusions, sheet_map):
    parts = []
    for yr in _year_range(start_year, end_year):
        df_y = load_data(yr, sheet_map[selected_question])
        if isinstance(df_y, str): continue
        row_y = _row_filter(df_y, entity, worker_cat, subq)
        if row_y.empty: continue
        rq = get_reporting_quarter(yr)
        s_y = build_quarter_series(row_y, yr, rq)
        if s_y.empty: continue
        keep = [lab for lab in s_y.index if lab not in exclusions]
        parts.append(s_y.loc[keep] if keep else pd.Series(dtype=float))
    if not parts: return pd.Series(dtype=float), None
    s = pd.concat(parts)
    order = sorted(s.index, key=lambda lab: (int(lab.split()[1]), int(lab.split()[0][1:])))
    s = s.loc[order]
    prior_vals, prior_idx = [], []
    for label in s.index:
        try:
            q, yr_str = label.split()
            prev_label = f"{q} {int(yr_str)-1}"
            if prev_label in s.index and not pd.isna(s.loc[prev_label]):
                prior_vals.append(float(s.loc[prev_label])); prior_idx.append(label)
        except Exception:
            continue
    prior = pd.Series(prior_vals, index=prior_idx, dtype=float) if prior_idx else None
    return s, prior

# -------------------------------------------------------------------------------------
# 2) NEW: App-wide constants & light helpers
# -------------------------------------------------------------------------------------
SHEET_MAP = {
    'Q1A: Employees': 'QC_Q1A_Main',
    'Q2A: Salary': 'QC_Q2A_Main',
    'Q3: Hours Worked': 'QC_Q3',
    'Q4: Vacancies': 'QC_Q4',
    'Q5: Separations': 'QC_Q5'
}
ALL_MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

# Adjust these to your exact labels in the workbook
ENTITY_MAP = {
    "ALL_FI": {"labels": ["All FI", "All RE", "All Financial Institutions"]},
    "GROUP":  {"labels": ["Banking Institution", "Insurance/Takaful", "DFI"]},
    "SUBGROUP_BANKING": {"labels": ["Commercial Banks","Investment Banks","Digital Banks","Foreign Banks","Islamic Banks"]},
    "SUBGROUP_INS_TAK": {"labels": ["Insurers","Takaful Operators"]},
}

def infer_level(entity_name: str) -> str:
    if entity_name in ENTITY_MAP["ALL_FI"]["labels"]: return "ALL_FI"
    if entity_name in ENTITY_MAP["GROUP"]["labels"]: return "GROUP"
    if entity_name in ENTITY_MAP["SUBGROUP_BANKING"]["labels"] or entity_name in ENTITY_MAP["SUBGROUP_INS_TAK"]["labels"]:
        return "SUBGROUP"
    return "RE"

def months_for_quarter(q: int):
    q = int(q)
    return ALL_MONTHS[(q-1)*3 : q*3]

def period_in_selected_quarter(period_label: str, year: int, quarter: int) -> bool:
    # Monthly label like "2025-Feb" or quarterly like "Q2 2025"
    if " " in period_label and period_label.startswith("Q"):
        try:
            q_lab, y_lab = period_label.split()
            return int(y_lab) == int(year) and int(q_lab[1:]) == int(quarter)
        except Exception:
            return False
    if "-" in period_label:
        y_str, m = period_label.split("-")
        return int(y_str) == int(year) and m in months_for_quarter(quarter)
    return False

# Store a clicked outlier for Drilldown
if "drill_ctx" not in st.session_state:
    st.session_state["drill_ctx"] = None

# -------------------------------------------------------------------------------------
# 3) Sidebar (GLOBAL) — Quarter-first, multi-year always on
# -------------------------------------------------------------------------------------
st.sidebar.title("QC Controls")

current_year = st.sidebar.selectbox("Current Year:", [2025, 2024, 2023, 2022, 2021], index=0)
reporting_quarter = get_reporting_quarter(current_year)
sel_quarter = st.sidebar.selectbox("Quarter to QC:", [1,2,3,4], index=reporting_quarter-1)

# Pillars (all three by default)
st.sidebar.subheader("Pillars")
use_thr_abs = st.sidebar.checkbox("Threshold + Absolute Change", value=True)
use_iqr     = st.sidebar.checkbox("IQR", value=True)
use_yoy     = st.sidebar.checkbox("YoY %", value=True)

# Thresholds (kept same defaults as your app)
st.sidebar.subheader("Detection Thresholds")
abs_thresh = st.sidebar.slider("Absolute Change", 10, 1000, 50, 10)
iqr_mult   = st.sidebar.slider("IQR Sensitivity", 1.0, 3.0, 1.5, 0.1)
yoy_pct    = st.sidebar.slider("YoY % Threshold", 0, 100, 30, 5, format="%d%%")
yoy_thresh = yoy_pct / 100.0

# Dataset & scope
st.sidebar.subheader("Dataset & Scope")
selected_question = st.sidebar.selectbox("Dataset:", options=list(SHEET_MAP.keys()))
df_current = load_data(current_year, SHEET_MAP[selected_question])
if isinstance(df_current, str):
    st.error(df_current)
    st.stop()

# Subquestion & worker defaults
entities_all = sorted(df_current['Entity / Group'].dropna().unique().tolist())
has_subq = 'Subquestion' in df_current.columns
default_subq = None
if has_subq:
    # Try to auto-pick a main roll-up by name; fall back to first
    subq_list = sorted(df_current['Subquestion'].dropna().unique().tolist())
    candidates = [s for s in subq_list if "A+B+C" in str(s) or "Total" in str(s)]
    default_subq = candidates[0] if candidates else (subq_list[0] if subq_list else "N/A")
else:
    subq_list = []
selected_subq = st.sidebar.selectbox("Subquestion:", options=(subq_list if has_subq else ["N/A"]), index= (subq_list.index(default_subq) if (has_subq and default_subq in subq_list) else 0))

wc_list = sorted(df_current[df_current['Subquestion']==selected_subq]['Worker Category'].dropna().unique().tolist()) if has_subq else sorted(df_current['Worker Category'].dropna().unique().tolist())
# Default “Total” if present
default_wc = next((w for w in wc_list if "Total" in str(w)), (wc_list[0] if wc_list else "N/A"))
selected_wc = st.sidebar.selectbox("Worker Category:", options=wc_list, index= wc_list.index(default_wc) if default_wc in wc_list else 0)

# Multi-year always ON + Exclusions
st.sidebar.markdown("---")
st.sidebar.subheader("Timeline Exclusions")
# Build options dynamically for exclusions
# Monthly exclude options
month_opts = []
for yr in range(2019, current_year+1):
    rq = get_reporting_quarter(yr)
    month_opts.extend([f"{yr}-{m}" for m in ALL_MONTHS[:rq*3]])
exclude_months = st.sidebar.multiselect("Exclude months (YYYY-MMM):", options=month_opts, default=[])

# Quarterly exclude options
qtr_opts = []
for yr in range(2019, current_year+1):
    rq = get_reporting_quarter(yr)
    qtr_opts.extend([f"Q{q} {yr}" for q in range(1, rq+1)])
exclude_quarters = st.sidebar.multiselect("Exclude quarters (Qx YYYY):", options=qtr_opts, default=[])

# View switcher
st.sidebar.markdown("---")
view = st.sidebar.radio("View", ["Outlier Inbox", "Drilldown", "Report & Export"], index=0)

# Threshold pack for reuse
THRESH = {'pct_thresh': 0.25, 'abs_thresh': abs_thresh, 'iqr_multiplier': iqr_mult, 'yoy_thresh': yoy_thresh}

# -------------------------------------------------------------------------------------
# 4) Core engines to assemble series and scan (NEW)
# -------------------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def build_series(entity_name: str, subq: str, wc: str, question_key: str, freq: str, start_year: int, end_year: int, excl_months: set, excl_quarters: set):
    """Always multi-year. Returns (series, prior) monthly or quarterly."""
    if freq == "Monthly":
        return build_multi_year_monthly_series(
            entity=entity_name, worker_cat=wc, subq=subq,
            selected_question=question_key, start_year=start_year, end_year=end_year,
            exclusions=excl_months, all_months=ALL_MONTHS, sheet_map=SHEET_MAP
        )
    else:
        return build_multi_year_quarterly_series(
            entity=entity_name, worker_cat=wc, subq=subq,
            selected_question=question_key, start_year=start_year, end_year=end_year,
            exclusions=excl_quarters, sheet_map=SHEET_MAP
        )

def scan_entities_for_quarter(entities: list, freq: str):
    """Run your find_outliers on each entity and keep only hits in the selected quarter."""
    rows = []
    start_year = 2019  # you can lower/raise this if needed
    end_year = current_year
    for ent in entities:
        s, prior = build_series(
            entity_name=ent, subq=selected_subq, wc=selected_wc, question_key=selected_question,
            freq=freq, start_year=start_year, end_year=end_year,
            excl_months=set(exclude_months), excl_quarters=set(exclude_quarters)
        )
        if s.empty or s.isnull().all(): 
            continue
        # Run your detector once and then filter to the selected quarter
        out_df = find_outliers(s, prior, **THRESH)
        if out_df.empty:
            continue
        # Pillar gating (apply which pillars the user enabled)
        # We infer pillars from reason strings
        for _, r in out_df.iterrows():
            period_label = r["Period"]
            if not period_in_selected_quarter(period_label, current_year, sel_quarter):
                continue
            reasons = str(r["Reason(s)"])
            hits = []
            if use_thr_abs and ("High Volatility" in reasons): hits.append("Threshold")
            if use_iqr and ("IQR Anomaly" in reasons): hits.append("IQR")
            if use_yoy and ("YoY Anomaly" in reasons): hits.append("YoY")
            if not hits:
                continue  # user disabled this pillar set
            level = infer_level(ent)
            rows.append({
                "Question": selected_question,
                "Subquestion": selected_subq,
                "Worker Category": selected_wc,
                "Level": level,
                "Entity / Group": ent,
                "Pillar(s)": ", ".join(sorted(set(hits))),
                "Reason(s)": reasons,
                "Period": period_label,
                "Value": s.loc[period_label] if period_label in s.index else np.nan
            })
    return pd.DataFrame(rows)

# -------------------------------------------------------------------------------------
# 5) Outlier Inbox (NEW)
# -------------------------------------------------------------------------------------
if view == "Outlier Inbox":
    st.title("📥 Outlier Inbox")
    st.caption(f"Quarter: **Q{sel_quarter} {current_year}** · Dataset: **{selected_question}** · Subquestion: **{selected_subq}** · Worker: **{selected_wc}**")

    # Frequency auto choice:
    # If your sheet is fundamentally monthly, we use Monthly; else you can add a toggle.
    freq = "Monthly"  # you can infer from sheet if needed

    # Build entity universe from the sheet (current year) under chosen subq/wc
    df_scope = df_current.copy()
    if has_subq:
        df_scope = df_scope[df_scope['Subquestion'] == selected_subq]
    df_scope = df_scope[df_scope['Worker Category'] == selected_wc]
    ent_list = sorted(df_scope['Entity / Group'].dropna().unique().tolist())

    # Scan
    with st.spinner("Scanning entities for selected quarter..."):
        inbox = scan_entities_for_quarter(ent_list, freq=freq)

    # KPIs
    c1, c2, c3, c4 = st.columns(4)
    total_hits = len(inbox)
    c1.metric("Outliers this quarter", f"{total_hits}")
    c2.metric("Threshold hits", str((inbox["Pillar(s)"].str.contains("Threshold")).sum()))
    c3.metric("IQR hits", str((inbox["Pillar(s)"].str.contains("IQR")).sum()))
    c4.metric("YoY hits", str((inbox["Pillar(s)"].str.contains("YoY")).sum()))

    if total_hits == 0:
        st.success("✅ No outliers for this quarter based on selected pillars.")
    else:
        # Sort: All FI first, then Group, Sub-group, then RE; and by severity hint (length of reasons)
        level_order = {"ALL_FI":0, "GROUP":1, "SUBGROUP":2, "RE":3}
        inbox["_lvl_sort"] = inbox["Level"].map(level_order).fillna(9)
        inbox["_sev"] = inbox["Reason(s)"].str.len()
        inbox = inbox.sort_values(by=["_lvl_sort","_sev"], ascending=[True, False])

        # Quick drill buttons: store in session_state
        def make_key(i): return f"drill_{i}"
        st.dataframe(
            inbox[["Question","Subquestion","Worker Category","Level","Entity / Group","Pillar(s)","Reason(s)","Period","Value"]],
            use_container_width=True
        )
        st.caption("Tip: Click a row below to open Drilldown for detail & contribution tree.")

        # Simple interactive table replacement: show buttons per row
        for i, r in inbox.iterrows():
            colA, colB = st.columns([0.85, 0.15])
            with colA:
                st.write(f"**{r['Entity / Group']}** — _{r['Level']}_ · {r['Pillar(s)']} · {r['Reason(s)']} · **{r['Period']}**")
            with colB:
                if st.button("🔎 Drilldown", key=f"rowbtn_{i}"):
                    st.session_state["drill_ctx"] = dict(r)
                    st.switch_page  # Streamlit 1.39+ supports st.switch_page; fallback: change radio
                    st.session_state["__view__"] = "Drilldown"

# -------------------------------------------------------------------------------------
# 6) Drilldown (NEW)
# -------------------------------------------------------------------------------------
if view == "Drilldown":
    st.title("🔎 Drilldown")
    ctx = st.session_state.get("drill_ctx", None)
    if not ctx:
        st.info("Pick an outlier from the **Outlier Inbox** to drill into.")
        st.stop()

    st.caption(f"Context · Q{sel_quarter} {current_year} · Dataset: **{ctx['Question']}** · Subquestion: **{ctx['Subquestion']}** · Worker: **{ctx['Worker Category']}**")
    st.markdown(
        f'<span class="pill pill-thr">Threshold</span> <span class="pill pill-iqr">IQR</span> <span class="pill pill-yoy">YoY</span>',
        unsafe_allow_html=True
    )

    # One combined timeline chart for the selected entity
    ent = ctx["Entity / Group"]
    freq = "Monthly"

    s, prior = build_series(
        entity_name=ent, subq=ctx["Subquestion"], wc=ctx["Worker Category"],
        question_key=ctx["Question"], freq=freq, start_year=2019, end_year=current_year,
        excl_months=set(exclude_months), excl_quarters=set(exclude_quarters)
    )
    if s.empty or s.isnull().all():
        st.warning("No data available for this entity.")
        st.stop()

    # Re-run detector for this single entity to mark points
    out_df = find_outliers(s, prior, **THRESH)
    x = list(s.index); y = s.values.astype(float)

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=x, y=y, mode='lines+markers', name='Timeline',
                             hovertemplate='Period: %{x}<br>Value: %{y:,.0f}<extra></extra>'))

    # Mark quarter window
    q_mask_x = [lab for lab in x if period_in_selected_quarter(lab, current_year, sel_quarter)]
    q_mask_y = [s.loc[l] for l in q_mask_x]
    if q_mask_x:
        fig.add_trace(go.Scatter(x=q_mask_x, y=q_mask_y, mode='markers', name=f'Q{sel_quarter} {current_year}',
                                 marker=dict(size=10, line=dict(width=1)), hoverinfo='skip'))

    # Outliers markers (all time, dim), highlight this quarter brighter
    if not out_df.empty:
        o_all_x = [p for p in out_df['Period'] if p in s.index]
        o_all_y = [s[p] for p in o_all_x]
        fig.add_trace(go.Scatter(x=o_all_x, y=o_all_y, mode='markers', name='Outliers (All)',
                                 marker=dict(symbol='x', size=12), opacity=0.35))
        o_q_x = [p for p in o_all_x if period_in_selected_quarter(p, current_year, sel_quarter)]
        o_q_y = [s[p] for p in o_q_x]
        if o_q_x:
            fig.add_trace(go.Scatter(x=o_q_x, y=o_q_y, mode='markers', name='Outliers (Selected Quarter)',
                                     marker=dict(symbol='x', size=14)))

    fig.update_layout(title=f"Timeline — {ent}", margin=dict(l=10,r=10,t=60,b=10), hovermode='x unified')
    fig.update_xaxes(tickangle=45)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": True})

    # Contributor tree — reuse same engine at each level on the same slice
    st.subheader("Contributor Tree")
    df_scope = df_current.copy()
    if has_subq: df_scope = df_scope[df_scope['Subquestion'] == ctx["Subquestion"]]
    df_scope = df_scope[df_scope['Worker Category'] == ctx["Worker Category"]]
    ent_universe = sorted(df_scope['Entity / Group'].dropna().unique().tolist())

    # Level 0: All FI
    all_fi_labels = ENTITY_MAP["ALL_FI"]["labels"]
    all_fi = [e for e in ent_universe if e in all_fi_labels]
    def render_level(title, entity_list):
        if not entity_list: 
            st.caption(f"*No {title.lower()} rows were found in this dataset.*")
            return pd.DataFrame()
        table = scan_entities_for_quarter(entity_list, freq="Monthly")
        if table.empty:
            st.markdown(f"**{title}:** <span class='flag-no'>OK</span>", unsafe_allow_html=True)
        else:
            st.markdown(f"**{title}:** <span class='flag-yes'>FLAGGED</span>", unsafe_allow_html=True)
            st.dataframe(table[["Entity / Group","Pillar(s)","Reason(s)","Period","Value"]], use_container_width=True)
        return table

    # Level 0
    lvl0 = render_level("All FI", all_fi)

    # Level 1: Groups
    groups = [e for e in ent_universe if e in ENTITY_MAP["GROUP"]["labels"]]
    st.markdown("---")
    st.markdown("**Groups**")
    lvl1 = render_level("Groups", groups)

    # Choose a group to inspect subgroups
    chosen_group = None
    if not lvl1.empty:
        opt_groups = lvl1["Entity / Group"].unique().tolist()
        chosen_group = st.selectbox("Choose a group to expand:", options=opt_groups, index=0)

    # Level 2: Sub-groups (based on chosen group)
    st.markdown("---")
    st.markdown("**Sub-groups**")
    subgrp_pool = []
    if chosen_group in ("Banking Institution",):
        subgrp_pool = ENTITY_MAP["SUBGROUP_BANKING"]["labels"]
    elif chosen_group in ("Insurance/Takaful",):
        subgrp_pool = ENTITY_MAP["SUBGROUP_INS_TAK"]["labels"]
    elif chosen_group in ("DFI",):
        subgrp_pool = []  # if you have DFI sub-buckets, add them to ENTITY_MAP
    subgroups = [e for e in ent_universe if e in subgrp_pool]
    lvl2 = render_level("Sub-groups", subgroups)

    chosen_bucket = None
    if not lvl2.empty:
        opt_buckets = lvl2["Entity / Group"].unique().tolist()
        if opt_buckets:
            chosen_bucket = st.selectbox("Choose a sub-group to expand:", options=opt_buckets, index=0)

    # Level 3: REs — heuristics: RE = universe minus known labels
    st.markdown("---")
    st.markdown("**Reporting Entities (REs)**")
    known_labels = set(sum([v["labels"] for v in ENTITY_MAP.values()], []))
    re_list = [e for e in ent_universe if e not in known_labels]
    # If a bucket chosen, further filter REs by same sheet naming convention if applicable (left generic)
    lvl3 = render_level("REs", re_list)

    st.caption("Note: Adjust ENTITY_MAP label lists at the top to perfectly match your workbook naming.")

# -------------------------------------------------------------------------------------
# 7) Report & Export (kept, quarter-scoped)
# -------------------------------------------------------------------------------------
if view == "Report & Export":
    st.title("📄 Report & Export")
    st.write("Scan workbook(s) and export flagged rows. Scope is set to the selected quarter above.")

    # Reuse your existing generator, but set defaults to current quarter only
    st.subheader("Thresholds")
    abs_thresh_report = st.slider("Absolute Change (Report)", 10, 1000, abs_thresh, 10)
    iqr_mult_report   = st.slider("IQR Sensitivity (Report)", 1.0, 3.0, iqr_mult, 0.1)
    yoy_thresh_report = st.slider("YoY % Threshold (Report)", 0, 100, int(yoy_thresh*100), 5, format="%d%%")/100.0

    questions_to_scan = st.multiselect(
        "Datasets:", options=list(SHEET_MAP.keys()), default=[selected_question]
    )

    # Per-dataset filters
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

    # Use your existing generate_full_report (copied from your code)
    def generate_full_report(
       sheet_map, years, questions_to_scan, thresholds, all_months,
       report_filters=None, quarter_mode="exact_selected", selected_quarter=None
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
        else:  # exact_selected
            months_to_use  = all_months[: int(selected_quarter) * 3] if selected_quarter else all_months
            focus_months   = months_for_quarter(int(selected_quarter)) if selected_quarter else all_months
        focus_set = set(focus_months)

        master_outlier_list = []
        progress_bar = st.progress(0, text="Initializing Scan...")
        total_scans = max(1, len(questions_to_scan))
        for i, q_name in enumerate(questions_to_scan):
            sheet_name = sheet_map[q_name]
            progress_bar.progress(i / total_scans, text=f"Scanning: {q_name}")
            df_current = load_data(years['current'], sheet_name)
            df_prior   = load_data(years['prior'], sheet_name) if years.get('prior') else None
            if isinstance(df_current, str): continue
            actual_months = [m for m in months_to_use if m in df_current.columns]
            if not actual_months: continue
            f_cfg = (report_filters or {}).get(q_name, {"subq": "ALL", "wc": "ALL"})
            subq_filter = f_cfg.get("subq", "ALL"); wc_filter = f_cfg.get("wc", "ALL")
            for _, row in df_current.iterrows():
                entity = row['Entity / Group']
                wc     = row['Worker Category']
                subq   = row['Subquestion'] if 'Subquestion' in df_current.columns else 'N/A'
                if subq_filter != "ALL" and subq not in subq_filter: continue
                if wc_filter   != "ALL" and wc   not in wc_filter: continue
                monthly_series = pd.to_numeric(row[actual_months], errors='coerce').astype(float)
                prior_series = None
                if not isinstance(df_prior, str) and df_prior is not None:
                    prior_row = _row_filter(df_prior, entity, wc, subq)
                    if not prior_row.empty:
                        prior_series = pd.to_numeric(prior_row[actual_months].iloc[0], errors='coerce').astype(float)
                outliers_monthly = find_outliers(monthly_series, prior_series, **thresholds)
                for _, o_row in outliers_monthly.iterrows():
                    if o_row['Period'] not in focus_set: continue
                    master_outlier_list.append([
                        q_name, entity, subq, wc,
                        'Monthly', o_row['Period'], o_row['Value'], o_row['Reason(s)']
                    ])
        progress_bar.progress(1.0, text="Scan Complete!")
        return pd.DataFrame(master_outlier_list, columns=['Question','Entity / Group','Subquestion','Worker Category','View','Period','Value','Reason(s)'])

    years = {'current': current_year, 'prior': current_year-1}
    report_thresholds = {'pct_thresh': 0.25, 'abs_thresh': abs_thresh_report, 'iqr_multiplier': iqr_mult_report, 'yoy_thresh': yoy_thresh_report}

    if st.button("🚀 Generate Report (Selected Quarter Only)", use_container_width=True):
        with st.spinner("Analyzing..."):
            final_report = generate_full_report(
                sheet_map=SHEET_MAP, years=years, questions_to_scan=questions_to_scan,
                thresholds=report_thresholds, all_months=ALL_MONTHS,
                report_filters=report_filters, quarter_mode="exact_selected", selected_quarter=sel_quarter
            )
        st.success(f"Scan complete! Found **{len(final_report)}** potential outliers.")
        if not final_report.empty:
            st.dataframe(final_report, use_container_width=True)
            csv = final_report.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download CSV", data=csv,
                file_name=f"qc_outliers_Q{sel_quarter}_{current_year}.csv",
                mime="text/csv", use_container_width=True
            )
