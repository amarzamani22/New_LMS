import os
import re
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

# -------------------- Page & Styles --------------------
st.set_page_config(page_title="LMS Analysis Dashboard", page_icon="🏦", layout="wide")
st.markdown("""
<style>
.stApp { background-color:#F0F2F6; }
.stMetric { border-radius:10px;padding:20px;background:#FFF;border:1px solid #E0E0E0;box-shadow:0 4px 6px rgba(0,0,0,0.04); }
.stButton>button { border-radius:8px;font-weight:600; }
.badge {
  display:inline-block; padding:4px 10px; border-radius:12px;
  background:#EEF5FF; border:1px solid #CFE2FF; color:#003366; font-weight:600;
  margin: 6px 0 10px 0;
}
</style>
""", unsafe_allow_html=True)

# -------------------- Constants --------------------
SHEET_MAP: Dict[str, str] = {
    'Q1A: Employees': 'QC_Q1A_Main',
    'Q2A: Salary': 'QC_Q2A_Main',
    'Q3: Hours Worked': 'QC_Q3',
    'Q4: Vacancies': 'QC_Q4',
    'Q5: Separations': 'QC_Q5'
}
ALL_MONTHS: List[str] = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
ENTITY_COL = "Entity / Group"
SUBQ_COL = "Subquestion"
WC_COL = "Worker Category"
ROLLUP_KEY = "All Financial Institutions"

# -------------------- Data Access --------------------
@st.cache_data
def load_qc_sheet(year: int, sheet_name: str) -> pd.DataFrame | str:
    path = fr"C:\Users\ttamarz\OneDrive - Bank Negara Malaysia\RLMS\Output\QC Template\qc_workbook_{year}.xlsx"
    if not os.path.exists(path):
        return f"Error: File not found → {path}"
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, header=5)
        df.dropna(axis=1, how='all', inplace=True)
        df.rename(columns={'Q1.1':'Q1_Total','Q2.1':'Q2_Total','Q3.1':'Q3_Total','Q4.1':'Q4_Total'}, inplace=True)
        return df
    except Exception as e:
        return f"Error reading sheet '{sheet_name}' from {path}: {e}"

@st.cache_data
def get_reporting_quarter(year: int) -> int:
    """Read _About sheet to know the current quarter in that workbook."""
    path = fr"C:\Users\ttamarz\OneDrive - Bank Negara Malaysia\RLMS\Output\QC Template\qc_workbook_{year}.xlsx"
    try:
        about = pd.read_excel(path, sheet_name="_About", header=None)
        row = about[about[0] == "Quarter"]
        if not row.empty:
            qn = int(re.search(r'\d+', str(row.iloc[0,1])).group())
            return max(1, min(4, qn))
    except Exception:
        pass
    return 4

def months_for_q(q: int) -> List[str]:
    return {1:['Jan','Feb','Mar'], 2:['Apr','May','Jun'], 3:['Jul','Aug','Sep'], 4:['Oct','Nov','Dec']}[int(q)]

def current_q_month_labels(current_year: int, reporting_q: int) -> List[str]:
    return [f"{current_year}-{m}" for m in months_for_q(reporting_q)]

def qc_row_slice(df: pd.DataFrame, entity: str, wc: str, subq: str) -> pd.DataFrame:
    cond = (df[ENTITY_COL] == entity) & (df[WC_COL] == wc)
    if SUBQ_COL in df.columns and subq != "N/A":
        cond &= (df[SUBQ_COL] == subq)
    return df[cond]

# -------------------- Multi-year Series Builders --------------------
def build_multi_year_monthly_series(
    entity: str, wc: str, subq: str, sheet_key: str,
    start_year: int, end_year: int, excluded_years: set[int]
) -> Tuple[pd.Series, Optional[pd.Series]]:
    vals, idx = [], []
    for yr in range(int(start_year), int(end_year)+1):
        if yr in excluded_years:
            continue
        df = load_qc_sheet(yr, SHEET_MAP[sheet_key])
        if isinstance(df, str):
            continue
        row = qc_row_slice(df, entity, wc, subq)
        if row.empty:
            continue
        rq = get_reporting_quarter(yr)
        months = [m for m in ALL_MONTHS[:rq*3] if m in row.columns]
        if not months:
            continue
        s = pd.to_numeric(row[months].iloc[0], errors="coerce").astype(float)
        for m in months:
            idx.append(f"{yr}-{m}")
            vals.append(s.get(m, np.nan))
    series = pd.Series(vals, index=idx, dtype=float)

    # Align YoY: previous year's same month, but only if present in our built series
    yoy_vals, yoy_idx = [], []
    for lab in series.index:
        yr, mon = lab.split('-')
        prev = f"{int(yr)-1}-{mon}"
        if prev in series.index and not pd.isna(series[prev]):
            yoy_idx.append(lab); yoy_vals.append(series[prev])
    yoy_series = pd.Series(yoy_vals, index=yoy_idx, dtype=float) if yoy_idx else None
    return series, yoy_series

def build_multi_year_quarterly_series(
    entity: str, wc: str, subq: str, sheet_key: str,
    start_year: int, end_year: int, excluded_years: set[int]
) -> Tuple[pd.Series, Optional[pd.Series]]:
    labels, values = [], []
    for yr in range(int(start_year), int(end_year)+1):
        if yr in excluded_years:
            continue
        df = load_qc_sheet(yr, SHEET_MAP[sheet_key])
        if isinstance(df, str):
            continue
        row = qc_row_slice(df, entity, wc, subq)
        if row.empty:
            continue
        rq = get_reporting_quarter(yr)

        # prefer Qx_Total; otherwise sum months for that quarter
        for q in range(1, rq+1):
            col = f"Q{q}_Total"
            if col in row.columns and not pd.isna(row.iloc[0][col]):
                v = float(pd.to_numeric(row.iloc[0][col], errors="coerce"))
            else:
                mlist = [m for m in months_for_q(q) if m in row.columns]
                v = float(pd.to_numeric(row[mlist].iloc[0], errors="coerce").astype(float).sum()) if mlist else np.nan
            if not np.isnan(v):
                labels.append(f"Q{q} {yr}")
                values.append(v)

    series = pd.Series(values, index=labels, dtype=float)

    # Align YoY for quarters (same Q, previous year)
    yoy_vals, yoy_idx = [], []
    for lab in series.index:
        q, yr_str = lab.split()
        prev = f"{q} {int(yr_str)-1}"
        if prev in series.index and not pd.isna(series[prev]):
            yoy_idx.append(lab); yoy_vals.append(series[prev])
    yoy_series = pd.Series(yoy_vals, index=yoy_idx, dtype=float) if yoy_idx else None
    return series, yoy_series

# -------------------- Outlier Engine --------------------
def find_outliers_v2(
    series: pd.Series,
    yoy_series: Optional[pd.Series],
    pct_thresh: float,      # MoM % threshold
    abs_cutoff: float,      # absolute change threshold
    iqr_k: float,           # IQR multiplier
    yoy_thresh: float       # YoY % threshold
) -> pd.DataFrame:
    """
    Outlier logic:
    1. High MoM% if |MoM| >= pct_thresh AND |Δ| >= abs_cutoff
    2. Outlier flagged only if (High MoM%) AND (YoY >= yoy_thresh OR IQR anomaly)
    Reasons include High MoM%, High YoY%, IQR Detect Outlier.
    """

    out = []
    clean = series.dropna()
    if clean.size < 2:
        return pd.DataFrame(columns=["Period", "Value", "Statistical Reasons"]).set_index("Period")

    # Compute IQR bounds
    q1, q3 = clean.quantile(0.25), clean.quantile(0.75)
    iqr = q3 - q1
    lb, ub = ((q1 - iqr_k * iqr, q3 + iqr_k * iqr) if iqr > 0 else (None, None))

    for i, (period, cur) in enumerate(series.items()):
        if pd.isna(cur) or i == 0:
            continue
        prev = series.iloc[i - 1]
        if pd.isna(prev) or prev == 0:
            continue

        abs_chg = cur - prev
        mom = abs_chg / prev

        # Step 1: High MoM check
        if abs(mom) >= pct_thresh and abs(abs_chg) >= abs_cutoff:
            reasons = [f"High MoM% ({mom:+.0%})"]
            extra_reason = False
            # Step 2: Additional significance
            if iqr > 0 and (cur < lb or cur > ub):
                reasons.append("IQR Detect Outlier")
                extra_reason = True

            if yoy_series is not None and period in yoy_series.index:
                py = yoy_series.get(period)
                if not pd.isna(py) and py != 0:
                    yoy = (cur - py) / py
                    if abs(yoy) >= yoy_thresh:
                        reasons.append(f"High YoY% ({yoy:+.0%})")
                        extra_reason = True

            # Only flag if extra_reason is True
            if extra_reason:
                out.append({
                    "Period": period,
                    "Value": f"{cur:,.2f}",
                    "Statistical Reasons": ", ".join(reasons)
                })

    return pd.DataFrame(out).set_index("Period") if out else pd.DataFrame(columns=["Period", "Value", "Statistical Reasons"]).set_index("Period")

# -------------------- Chart (Dual-axis: Bar + Line + Outliers) --------------------
def plot_dual_axis_with_outliers(
    series: pd.Series,
    growth_pct: pd.Series,
    outliers_focus: pd.DataFrame,
    title: str,
    left_title: str = "Value",
    right_title: str = "% Change (MoM)"
):
    x = list(series.index)
    y = series.values.astype(float)
    g = growth_pct.reindex(series.index).fillna(0)

    fig = go.Figure()

    # Bars for values (left axis)
    fig.add_trace(go.Bar(
        x=x, y=y, name=left_title, yaxis="y1", opacity=0.75,
        hovertemplate='Period: %{x}<br>Value: %{y:,.0f}<extra></extra>'
    ))

    # Line for % growth (right axis)
    fig.add_trace(go.Scatter(
        x=x, y=g.values, name=right_title, yaxis="y2",
        mode="lines+markers", line=dict(width=2, dash="dot", color="#003366"),
        hovertemplate='Period: %{x}<br>% Growth: %{y:+.0f}%<extra></extra>'
    ))

    # Outlier markers (red X) over the bars (left axis)
    if not outliers_focus.empty:
        ox = [p for p in outliers_focus.index if p in series.index]
        oy = [series[p] for p in ox]
        oreason = [outliers_focus.loc[p, "Statistical Reasons"] for p in ox]
        fig.add_trace(go.Scatter(
            x=ox, y=oy, mode='markers', name='True Outlier',
            yaxis="y1",
            marker=dict(symbol='x', size=14, color='red', line_width=2),
            hovertemplate='Period: %{x}<br>Value: %{y:,.0f}<br>%{customdata}<extra></extra>',
            customdata=oreason
        ))

    fig.update_layout(
        title=dict(text=title, font=dict(size=16), x=0.5, xanchor='center'),
        hovermode='x unified',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='left', x=0),
        margin=dict(l=10, r=10, t=70, b=10),
        xaxis=dict(tickangle=45),
        yaxis=dict(title=left_title, side='left'),
        yaxis2=dict(
            title=right_title, overlaying='y', side='right', showgrid=False,
            tickformat=".0f"  # already % units in trace hover; axis shows integers (±20, ±40 …)
        )
    )
    st.plotly_chart(fig, use_container_width=True)

# -------------------- UI Helpers --------------------
def sidebar_controls():
    st.sidebar.title("Analysis Controls")
    # years
    years_available = list(range(2019, 2031))
    current_year = st.sidebar.selectbox("Current Year (detect current quarter from this workbook):", options=sorted(years_available, reverse=True), index=years_available.index(2025))
    start_year = st.sidebar.selectbox("Start Year (timeline):", options=years_available, index=years_available.index(2022))
    end_year = st.sidebar.selectbox("End Year (timeline):", options=years_available, index=years_available.index(2025))
    if end_year < start_year:
        st.sidebar.error("End Year must be ≥ Start Year.")
    exclude_years = st.sidebar.multiselect("Exclude Years (optional):", options=[y for y in range(start_year, end_year+1)], default=[])

    # thresholds
    st.sidebar.markdown("---")
    st.sidebar.subheader("Thresholds")
    mom_pct = st.sidebar.slider("MoM/QoQ % Threshold (Gate)", 0, 100, 25, 5, format="%d%%") / 100.0
    abs_cut = st.sidebar.slider("Absolute Change (Significance)", 10, 1000, 50, 10)
    iqr_k = st.sidebar.slider("IQR Sensitivity (Significance)", 1.0, 3.0, 1.5, 0.1)
    yoy_pct = st.sidebar.slider("YoY % Threshold (Significance)", 0, 100, 30, 5, format="%d%%") / 100.0

    # view & dataset
    st.sidebar.markdown("---")
    time_view = st.sidebar.radio("Frequency:", options=['Monthly','Quarterly'], horizontal=True)
    dataset = st.sidebar.selectbox("Dataset:", options=list(SHEET_MAP.keys()))

    # --- Outlier focus window (NEW) ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("Outlier Focus")
    focus_mode = st.sidebar.radio(
        "Show outliers for:",
        options=["Current quarter", "Pick year & quarter"],
        index=0,
        horizontal=False
    )
    focus_year = None
    focus_quarter = None
    if focus_mode == "Pick year & quarter":
        focus_year = st.sidebar.selectbox(
            "Focus Year:",
            options=list(range(start_year, end_year + 1)),
            index=list(range(start_year, end_year + 1)).index(min(end_year, int(current_year)))
        )
        focus_quarter = st.sidebar.selectbox("Focus Quarter:", options=[1, 2, 3, 4], index=0)

    return {
        "current_year": int(current_year),
        "start_year": int(start_year),
        "end_year": int(end_year),
        "exclude_years": set(map(int, exclude_years)),
        "mom_pct": float(mom_pct),
        "abs_cut": float(abs_cut),
        "iqr_k": float(iqr_k),
        "yoy_pct": float(yoy_pct),
        "time_view": time_view,
        "dataset": dataset,
        "focus_mode": focus_mode,
        "focus_year": int(focus_year) if focus_year is not None else None,
        "focus_quarter": int(focus_quarter) if focus_quarter is not None else None,
    }

# -------------------- Main Interactive View --------------------
def main_view():
    st.title("🏦 LMS Analysis Dashboard")

    cfg = sidebar_controls()
    rq = get_reporting_quarter(cfg["current_year"])
    st.sidebar.info(f"Analyzing **{cfg['start_year']}–{cfg['end_year']}** (excl: {', '.join(map(str,cfg['exclude_years'])) or 'none'}) • Current workbook: **{cfg['current_year']} Q{rq}**")

    # Decide which period's outliers to show (NEW)
    if cfg["focus_mode"] == "Current quarter":
        focus_q = rq
        focus_year = cfg["current_year"]
    else:
        focus_q = cfg["focus_quarter"] or rq
        focus_year = cfg["focus_year"] or cfg["current_year"]
    focus_month_labels = set([f"{focus_year}-{m}" for m in months_for_q(focus_q)])
    focus_quarter_label = f"Q{focus_q} {focus_year}"

    # Badge (NEW)
    month_str = "–".join(months_for_q(focus_q))
    st.markdown(f'<div class="badge">Outlier focus: <b>Q{focus_q} {focus_year}</b> <span style="opacity:.7">({month_str})</span></div>', unsafe_allow_html=True)

    df_cur = load_qc_sheet(cfg["current_year"], SHEET_MAP[cfg["dataset"]])
    if isinstance(df_cur, str):
        st.error(df_cur)
        return

    # Entity/Subquestion/Worker Category selection
    st.header(f"Outlier Detection: {cfg['dataset']}")
    entity = st.selectbox("Entity / Group:", options=df_cur[ENTITY_COL].unique(),
                          index=list(df_cur[ENTITY_COL].unique()).index(ROLLUP_KEY) if ROLLUP_KEY in df_cur[ENTITY_COL].unique() else 0)

    if SUBQ_COL in df_cur.columns and df_cur[df_cur[ENTITY_COL]==entity][SUBQ_COL].nunique() > 1:
        u_subq = df_cur[df_cur[ENTITY_COL]==entity][SUBQ_COL].unique()
        subq_default = np.where(u_subq == 'Employment = A+B(i)+B(ii)')[0][0] if 'Employment = A+B(i)+B(ii)' in u_subq else 0
        subq = st.selectbox("Subquestion:", options=u_subq, index=int(subq_default))
        u_wc = df_cur[(df_cur[ENTITY_COL]==entity) & (df_cur[SUBQ_COL]==subq)][WC_COL].unique()
        wc_default = np.where(u_wc == 'Total Employment')[0][0] if 'Total Employment' in u_wc else 0
        wc = st.selectbox("Worker Category:", options=u_wc, index=int(wc_default))
    else:
        subq = "N/A"
        u_wc = df_cur[df_cur[ENTITY_COL]==entity][WC_COL].unique()
        wc_default = np.where(u_wc == 'Total Employment')[0][0] if 'Total Employment' in u_wc else 0
        wc = st.selectbox("Worker Category:", options=u_wc, index=int(wc_default))

    st.caption(f"Displaying: {entity} | {subq} | {wc}")

    # Build multi-year series
    if cfg["time_view"] == "Monthly":
        series, yoy_series = build_multi_year_monthly_series(
            entity, wc, subq, cfg["dataset"],
            cfg["start_year"], cfg["end_year"], cfg["exclude_years"]
        )
        if series.empty or series.isnull().all():
            st.warning("No data found for the selected timeline/filters.")
            return

        # Detection on full multi-year series
        out_all = find_outliers_v2(series, yoy_series, cfg["mom_pct"], cfg["abs_cut"], cfg["iqr_k"], cfg["yoy_pct"])

        # Focus markers: selected quarter (default = current quarter)  (NEW)
        out_focus = out_all.loc[out_all.index.intersection(focus_month_labels)]

        # Growth line (%MoM), rounded like VR for display (unchanged)
        growth_pct = (series.pct_change() * 100).round()  # integers like +41, -8, …
        plot_dual_axis_with_outliers(
            series=series,
            growth_pct=growth_pct,
            outliers_focus=out_focus,
            title=f"Monthly Trend ({cfg['start_year']}–{cfg['end_year']})"
        )

    else:
        series, yoy_series = build_multi_year_quarterly_series(
            entity, wc, subq, cfg["dataset"],
            cfg["start_year"], cfg["end_year"], cfg["exclude_years"]
        )
        if series.empty or series.isnull().all():
            st.warning("No data found for the selected timeline/filters.")
            return

        # Detection on full multi-year series
        out_all = find_outliers_v2(series, yoy_series, cfg["mom_pct"], cfg["abs_cut"], cfg["iqr_k"], cfg["yoy_pct"])

        # Focus markers: selected quarter (default = current quarter) (NEW)
        out_focus = out_all.loc[out_all.index.intersection({focus_quarter_label})]

        # %QoQ growth line (same logic)
        growth_pct = (series.pct_change() * 100).round()
        plot_dual_axis_with_outliers(
            series=series,
            growth_pct=growth_pct,
            outliers_focus=out_focus,
            title=f"Quarterly Trend ({cfg['start_year']}–{cfg['end_year']})",
            right_title="% Change (QoQ)"
        )

    # Metrics & Table
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Latest Value", f"{series.iloc[-1]:,.0f}")
    c2.metric("Average", f"{series.mean():,.0f}")
    c3.metric("Highest Value", f"{np.nanmax(series):,.0f}")
    c4.metric("Current-Q Outliers", len(out_focus))

    if not out_focus.empty:
        st.error("🚨 True Outlier(s) in Current Quarter")
        st.dataframe(out_focus, use_container_width=True)
    else:
        st.success("✅ No current-quarter outliers")

# -------------------- Run --------------------
if __name__ == "__main__":
    main_view()
