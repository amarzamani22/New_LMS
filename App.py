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

# -------------------- VR loader & matching helpers --------------------
def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())

@st.cache_data
def load_vr_variance(vr_path: str) -> pd.DataFrame | str:
    """Load VR staging (sheet 'Variance') and normalize key columns."""
    if not vr_path or not os.path.exists(vr_path):
        return "PENDING"  # special marker to show 'Pending submission'
    try:
        df = pd.read_excel(vr_path, sheet_name="Variance")
        needed = ["Entity Name","Year","Quarter","Month","Question","Subquestion","Worker Category","%Growth","Justification"]
        missing = [c for c in needed if c not in df.columns]
        if missing:
            return f"VR file missing columns: {missing}"
        # normalize helper cols
        df["_ent"]  = df["Entity Name"].map(_norm)
        df["_subq"] = df["Subquestion"].map(_norm)
        df["_wc"]   = df["Worker Category"].map(_norm)
        df["_qnum"] = df["Quarter"].astype(str).str.extract(r"(\d)").astype(float)
        df["_month"]= df["Month"].astype(str).str[:3]  # Jan/Feb/...
        return df
    except Exception as e:
        return f"Error reading VR file: {e}"

def _question_code_from_dataset(dataset_key: str) -> str:
    # 'Q1A: Employees' -> 'Q1A'
    return dataset_key.split(":")[0].replace(" ", "")

def find_vr_just_for_periods(
    vr_df: pd.DataFrame | str,
    dataset_key: str,
    entity_name: str,
    subq: str,
    wc: str,
    periods: List[str],
) -> str:
    """Return concatenated justification(s) for given periods."""
    if isinstance(vr_df, str):
        return "Pending submission" if vr_df == "PENDING" else vr_df

    qcode = _question_code_from_dataset(dataset_key)
    ent_norm = _norm(entity_name)
    subq_norm = _norm(subq)
    wc_norm = _norm(wc)
    justs = []

    for p in periods:
        if "-" in p and "Q" not in p:
            yr_str, mon = p.split("-")
            try:
                yr = int(yr_str)
            except:
                continue
            sub = vr_df[
                (vr_df["Year"] == yr) &
                (vr_df["_month"].str[:3].str.lower() == mon[:3].lower()) &
                (vr_df["Question"].astype(str).str.upper() == qcode.upper()) &
                (vr_df["_ent"] == ent_norm) &
                (vr_df["_wc"] == wc_norm)
            ]
            if subq_norm and subq_norm != _norm("N/A"):
                sub = sub[sub["_subq"] == subq_norm]
            if sub.empty:
                q_from_mon = {"jan":1,"feb":1,"mar":1,"apr":2,"may":2,"jun":2,"jul":3,"aug":3,"sep":3,"oct":4,"nov":4,"dec":4}[mon[:3].lower()]
                sub = vr_df[
                    (vr_df["Year"] == yr) &
                    (vr_df["_qnum"] == q_from_mon) &
                    (vr_df["Question"].astype(str).str.upper() == qcode.upper()) &
                    (vr_df["_ent"] == ent_norm) &
                    (vr_df["_wc"] == wc_norm)
                ]
                if subq_norm and subq_norm != _norm("N/A"):
                    sub = sub[sub["_subq"] == subq_norm]
        else:
            try:
                qlab, yr_str = p.split()
                qn = int(qlab[1:])
                yr = int(yr_str)
            except:
                continue
            sub = vr_df[
                (vr_df["Year"] == yr) &
                (vr_df["_qnum"] == qn) &
                (vr_df["Question"].astype(str).str.upper() == qcode.upper()) &
                (vr_df["_ent"] == ent_norm) &
                (vr_df["_wc"] == wc_norm)
            ]
            if subq_norm and subq_norm != _norm("N/A"):
                sub = sub[sub["_subq"] == subq_norm]

        js = [j for j in sub["Justification"].astype(str).tolist()
              if str(j).strip() and str(j).strip().lower() != "nan"]
        if js:
            justs.append(" | ".join(sorted(set(js))))

    if not justs:
        return "—"
    return " | ".join(sorted(set(justs)))

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
    out = []
    clean = series.dropna()
    if clean.size < 2:
        return pd.DataFrame(columns=["Period", "Value", "Statistical Reasons"]).set_index("Period")

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

        if abs(mom) >= pct_thresh and abs(abs_chg) >= abs_cutoff:
            reasons = [f"High MoM% ({mom:+.0%})"]
            extra_reason = False
            if iqr > 0 and (cur < lb or cur > ub):
                reasons.append("IQR Detect Outlier"); extra_reason = True
            if yoy_series is not None and period in yoy_series.index:
                py = yoy_series.get(period)
                if not pd.isna(py) and py != 0:
                    yoy = (cur - py) / py
                    if abs(yoy) >= yoy_thresh:
                        reasons.append(f"High YoY% ({yoy:+.0%})")
                        extra_reason = True
            if extra_reason:
                out.append({
                    "Period": period,
                    "Value": f"{cur:,.2f}",
                    "Statistical Reasons": ", ".join(reasons)
                })

    return pd.DataFrame(out).set_index("Period") if out else pd.DataFrame(columns=["Period", "Value", "Statistical Reasons"]).set_index("Period")

# -------------------- Chart (Dual-axis) --------------------
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

    fig.add_trace(go.Bar(
        x=x, y=y, name=left_title, yaxis="y1", opacity=0.75,
        hovertemplate='Period: %{x}<br>Value: %{y:,.0f}<extra></extra>'
    ))

    fig.add_trace(go.Scatter(
        x=x, y=g.values, name=right_title, yaxis="y2",
        mode="lines+markers", line=dict(width=2, dash="dot", color="#003366"),
        hovertemplate='Period: %{x}<br>% Growth: %{y:+.0f}%<extra></extra>'
    ))

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
            tickformat=".0f"
        )
    )
    st.plotly_chart(fig, use_container_width=True)

# -------------------- NEW: VR Justification for all Worker Categories (helpers) --------------------
def _norm_key(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip()).lower()

def collect_vr_by_worker(vr_df: pd.DataFrame | str,
                         dataset_key: str,
                         entity_name: str,
                         subq: str,
                         periods: list[str]) -> pd.DataFrame:
    """Return a tidy DF of all Worker Categories (incl. Total) with %Growth and Justification for the given filters."""
    if isinstance(vr_df, str):
        return pd.DataFrame({"Worker Category": ["—"], "%Growth": [""], "Justification": [("Pending submission" if vr_df=="PENDING" else vr_df)]})
    if vr_df is None or vr_df.empty:
        return pd.DataFrame({"Worker Category": ["—"], "%Growth": [""], "Justification": ["—"]})

    qcode = _question_code_from_dataset(dataset_key).upper()
    ent_norm = _norm_key(entity_name)
    subq_norm = _norm_key(subq)
    rows = []

    for p in periods:
        if "-" in p and "Q" not in p:
            try:
                yr_str, mon = p.split("-")
                yr = int(yr_str)
            except Exception:
                continue
            mon3 = mon[:3]
            sub = vr_df[
                (vr_df["Year"] == yr) &
                (vr_df["Question"].astype(str).str.upper() == qcode) &
                (vr_df["_ent"] == ent_norm)
            ]
            if subq_norm and subq_norm != _norm_key("N/A"):
                sub = sub[sub["_subq"] == subq_norm]
            sub = sub[sub["_month"].str[:3].str.lower() == mon3.lower()]
        else:
            try:
                qlab, yr_str = p.split()
                qn = int(qlab[1:])
                yr = int(yr_str)
            except Exception:
                continue
            sub = vr_df[
                (vr_df["Year"] == yr) &
                (vr_df["_qnum"] == qn) &
                (vr_df["Question"].astype(str).str.upper() == qcode) &
                (vr_df["_ent"] == ent_norm)
            ]
            if subq_norm and subq_norm != _norm_key("N/A"):
                sub = sub[sub["_subq"] == subq_norm]

        if sub.empty:
            continue
        take = sub[["Worker Category","%Growth","Justification"]].copy()
        take["Worker Category"] = take["Worker Category"].astype(str).str.strip()
        rows.append(take)

    if not rows:
        return pd.DataFrame({"Worker Category": ["—"], "%Growth": [""], "Justification": ["—"]})

    out = pd.concat(rows, ignore_index=True)
    out = (out
           .groupby(["Worker Category","%Growth"], dropna=False)["Justification"]
           .apply(lambda s: " | ".join(sorted(set(str(v) for v in s if str(v).strip()))))
           .reset_index())
    out["_r"] = out["Worker Category"].str.contains("total", case=False, na=False).map(lambda x: 0 if x else 1)
    out = out.sort_values(["_r","Worker Category"]).drop(columns=["_r"]).reset_index(drop=True)
    return out

def render_contributors_with_vr(contrib_df: pd.DataFrame,
                                vr_df: pd.DataFrame | str,
                                dataset_key: str,
                                subq_display: str,
                                periods_for_focus: list[str]):
    """Render contributors table + per-FI popover/expander showing all Worker Categories’ justifications."""
    if contrib_df is None or contrib_df.empty:
        st.info("No contributor data available.")
        return

    col_entity = next((c for c in contrib_df.columns if c.lower().startswith("entity")), "Entity / Group")
    col_prev   = next((c for c in contrib_df.columns if "prev" in c.lower()), None)
    col_curr   = next((c for c in contrib_df.columns if "curr" in c.lower()), None)
    col_diff   = next((c for c in contrib_df.columns if "absolute" in c.lower() or "Δ" in c or "diff" in c.lower()), None)
    col_mom    = next((c for c in contrib_df.columns if "% change" in c.lower() or "mom" in c.lower() or "qoq" in c.lower()), None)
    col_share  = next((c for c in contrib_df.columns if "contribution" in c.lower() or "share" in c.lower()), None)

    show_cols = [c for c in [col_entity, col_prev, col_curr, col_diff, col_mom, col_share] if c in contrib_df.columns]
    st.subheader("Contribution by FI")
    st.dataframe(contrib_df[show_cols].style.format({
        col_prev: "{:,.0f}" if col_prev else "",
        col_curr: "{:,.0f}" if col_curr else "",
        col_diff: "{:,.0f}" if col_diff else "",
        col_mom:  "{:+.1%}" if col_mom else "",
        col_share:"{:.1f}%" if col_share else "",
    }), use_container_width=True)

    st.caption("Click a row below to view all Worker Categories (incl. Total) justifications for the same period(s).")
    for i, row in contrib_df.iterrows():
        ent = str(row.get(col_entity, ""))
        cols = st.columns([0.35, 0.2, 0.2, 0.25])
        cols[0].markdown(f"**{ent}**")
        if col_diff:
            try:
                cols[1].markdown(f"Δ: **{float(row[col_diff]):,.0f}**")
            except Exception:
                cols[1].markdown(f"Δ: **{row[col_diff]}**")
        if col_share:
            try:
                cols[2].markdown(f"Share: **{float(row[col_share]):.1f}%**")
            except Exception:
                cols[2].markdown(f"Share: **{row[col_share]}**")

        try:
            pop = cols[3].popover("FI Justification", use_container_width=True, key=f"pop_{i}")
            with pop:
                vr_tbl = collect_vr_by_worker(
                    vr_df=vr_df,
                    dataset_key=dataset_key,
                    entity_name=ent,
                    subq=subq_display,
                    periods=periods_for_focus
                )
                st.dataframe(vr_tbl, use_container_width=True)
        except Exception:
            with cols[3].expander("FI Justification", expanded=False):
                vr_tbl = collect_vr_by_worker(
                    vr_df=vr_df,
                    dataset_key=dataset_key,
                    entity_name=ent,
                    subq=subq_display,
                    periods=periods_for_focus
                )
                st.dataframe(vr_tbl, use_container_width=True)

# -------------------- UI Helpers --------------------
def sidebar_controls():
    st.sidebar.title("Analysis Controls")
    years_available = list(range(2019, 2031))
    current_year = st.sidebar.selectbox("Current Year (detect current quarter from this workbook):", options=sorted(years_available, reverse=True), index=years_available.index(2025))
    start_year = st.sidebar.selectbox("Start Year (timeline):", options=years_available, index=years_available.index(2022))
    end_year = st.sidebar.selectbox("End Year (timeline):", options=years_available, index=years_available.index(2025))
    if end_year < start_year:
        st.sidebar.error("End Year must be ≥ Start Year.")
    exclude_years = st.sidebar.multiselect("Exclude Years (optional):", options=[y for y in range(start_year, end_year+1)], default=[])

    st.sidebar.markdown("---")
    st.sidebar.subheader("Thresholds")
    mom_pct = st.sidebar.slider("MoM/QoQ % Threshold (Gate)", 0, 100, 25, 5, format="%d%%") / 100.0
    abs_cut = st.sidebar.slider("Absolute Change (Significance)", 10, 1000, 50, 10)
    iqr_k = st.sidebar.slider("IQR Sensitivity (Significance)", 1.0, 3.0, 1.5, 0.1)
    yoy_pct = st.sidebar.slider("YoY % Threshold (Significance)", 0, 100, 30, 5, format="%d%%") / 100.0

    st.sidebar.markdown("---")
    time_view = st.sidebar.radio("Frequency:", options=['Monthly','Quarterly'], horizontal=True)
    dataset = st.sidebar.selectbox("Dataset:", options=list(SHEET_MAP.keys()))

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

    st.sidebar.markdown("---")
    st.sidebar.subheader("VR Staging")
    vr_path = st.sidebar.text_input(
        "Full path to VR staging Excel (sheet 'Variance'):",
        value="", placeholder=r"C:\...\VR_Consol_2025_Quarter1.xlsx"
    )

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
        "vr_path": vr_path.strip()
    }

# -------------------- Main Interactive View --------------------
def main_view():
    st.title("🏦 LMS Analysis Dashboard")

    cfg = sidebar_controls()
    rq = get_reporting_quarter(cfg["current_year"])
    st.sidebar.info(f"Analyzing **{cfg['start_year']}–{cfg['end_year']}** (excl: {', '.join(map(str,cfg['exclude_years'])) or 'none'}) • Current workbook: **{cfg['current_year']} Q{rq}**")

    if cfg["focus_mode"] == "Current quarter":
        focus_q = rq
        focus_year = cfg["current_year"]
    else:
        focus_q = cfg["focus_quarter"] or rq
        focus_year = cfg["focus_year"] or cfg["current_year"]
    focus_month_labels = set([f"{focus_year}-{m}" for m in months_for_q(focus_q)])
    focus_quarter_label = f"Q{focus_q} {focus_year}"

    month_str = "–".join(months_for_q(focus_q))
    st.markdown(f'<div class="badge">Outlier focus: <b>Q{focus_q} {focus_year}</b> <span style="opacity:.7">({month_str})</span></div>', unsafe_allow_html=True)

    df_cur = load_qc_sheet(cfg["current_year"], SHEET_MAP[cfg["dataset"]])
    if isinstance(df_cur, str):
        st.error(df_cur)
        return
    
    # Normalization for selection display (unchanged behavior, internal only)
    def _norm(s: str) -> str:
        return re.sub(r"\s+", " ", str(s).strip()).casefold()
    def _canonical(series: pd.Series) -> pd.Series:
        tmp = series.dropna().astype(str)
        norm = tmp.map(_norm)
        canon_map = (
            pd.DataFrame({"orig": tmp, "norm": norm})
            .groupby("norm")["orig"]
            .agg(lambda s: s.value_counts().idxmax())
        )
        return norm.map(canon_map)
    df_cur["_ent_norm"]  = df_cur[ENTITY_COL].map(_norm)
    df_cur["_ent_disp"]  = _canonical(df_cur[ENTITY_COL])
    if SUBQ_COL in df_cur.columns:
        df_cur["_subq_norm"] = df_cur[SUBQ_COL].map(_norm)
        df_cur["_subq_disp"] = _canonical(df_cur[SUBQ_COL])
    df_cur["_wc_norm"]   = df_cur[WC_COL].map(_norm)
    df_cur["_wc_disp"]   = _canonical(df_cur[WC_COL])

    st.header(f"Outlier Detection: {cfg['dataset']}")
    entity_options = sorted(df_cur["_ent_disp"].dropna().unique().tolist())
    entity_disp = st.selectbox("Entity / Group:", options=entity_options,
        index=entity_options.index(ROLLUP_KEY) if ROLLUP_KEY in entity_options else 0)
    entity_norm = _norm(entity_disp)

    if SUBQ_COL in df_cur.columns and \
        df_cur[df_cur["_ent_norm"]==entity_norm]["_subq_disp"].nunique() > 1:
        subq_options = sorted(df_cur.loc[df_cur["_ent_norm"]==entity_norm, "_subq_disp"].dropna().unique().tolist())
        subq_disp = st.selectbox("Subquestion:", options=subq_options, index=0)
        subq_norm = _norm(subq_disp)

        wc_options = sorted(
            df_cur.loc[
                (df_cur["_ent_norm"]==entity_norm) & (df_cur["_subq_norm"]==subq_norm),
                "_wc_disp"
            ].dropna().unique().tolist()
        )

        wc_disp = st.selectbox("Worker Category:", options=wc_options, index=0)
        wc_norm = _norm(wc_disp)
        data_row = df_cur[
            (df_cur["_ent_norm"]==entity_norm) &
            (df_cur["_subq_norm"]==subq_norm) &
            (df_cur["_wc_norm"]==wc_norm)
        ]
    else:
        subq_disp = "N/A"; subq_norm = _norm(subq_disp)
        wc_options = sorted(
            df_cur.loc[df_cur["_ent_norm"]==entity_norm, "_wc_disp"].dropna().unique().tolist()
        )
        wc_disp = st.selectbox("Worker Category:", options=wc_options, index=0)
        wc_norm = _norm(wc_disp)
        data_row = df_cur[
            (df_cur["_ent_norm"]==entity_norm) &
            (df_cur["_wc_norm"]==wc_norm)
        ]
    st.caption(f"Displaying: {entity_disp} | {subq_disp} | {wc_disp}")

    # Load VR
    vr_df = load_vr_variance(cfg["vr_path"])

    # Build series + outliers
    if cfg["time_view"] == "Monthly":
        series, yoy_series = build_multi_year_monthly_series(
            entity_disp, wc_disp, subq_disp, cfg["dataset"],
            cfg["start_year"], cfg["end_year"], cfg["exclude_years"]
        )
        if series.empty or series.isnull().all():
            st.warning("No data found for the selected timeline/filters.")
            return

        out_all = find_outliers_v2(series, yoy_series, cfg["mom_pct"], cfg["abs_cut"], cfg["iqr_k"], cfg["yoy_pct"])
        out_focus = out_all.loc[out_all.index.intersection(focus_month_labels)]

        # Growth line (%MoM), rounded like VR for display
        growth_pct = series.pct_change()
        growth_pct = growth_pct.replace([np.inf, -np.inf], np.nan)
        growth_pct = (growth_pct * 100).round().fillna(0)
        plot_dual_axis_with_outliers(
            series=series,
            growth_pct=growth_pct,
            outliers_focus=out_focus,
            title=f"Monthly Trend ({cfg['start_year']}–{cfg['end_year']})"
        )

    else:
        series, yoy_series = build_multi_year_quarterly_series(
            entity_disp, wc_disp, subq_disp, cfg["dataset"],
            cfg["start_year"], cfg["end_year"], cfg["exclude_years"]
        )
        if series.empty or series.isnull().all():
            st.warning("No data found for the selected timeline/filters.")
            return

        out_all = find_outliers_v2(series, yoy_series, cfg["mom_pct"], cfg["abs_cut"], cfg["iqr_k"], cfg["yoy_pct"])
        out_focus = out_all.loc[out_all.index.intersection({focus_quarter_label})]

        growth_pct = series.pct_change()
        growth_pct = growth_pct.replace([np.inf, -np.inf], np.nan)
        growth_pct = (growth_pct * 100).round().fillna(0)
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
        # Attach per-row FI justification (your previous per-row logic)
        out_focus = out_focus.copy()
        if cfg["time_view"] == "Monthly":
            per_row_just = []
            for p in out_focus.index.tolist():
                per_row_just.append(find_vr_just_for_periods(vr_df, cfg["dataset"], entity_disp, subq_disp, wc_disp, [p]))
            out_focus["FI Justification"] = per_row_just
        else:
            qn = int(focus_quarter_label.split()[0][1:])
            qy = int(focus_quarter_label.split()[1])
            q_months = [f"{qy}-{m}" for m in months_for_q(qn)]
            out_focus["FI Justification"] = [
                find_vr_just_for_periods(vr_df, cfg["dataset"], entity_disp, subq_disp, wc_disp, q_months)
                for _ in out_focus.index.tolist()
            ]
        st.dataframe(out_focus, use_container_width=True)
    else:
        st.success("✅ No current-quarter outliers")

    # -------------- NEW: Contribution by FI + VR popover --------------
    # Only meaningful for rollup "All Financial Institutions"
    if entity_disp == ROLLUP_KEY and not out_focus.empty:
        # Build an attribution frame from the current QC sheet for the selected period
        # We re-use the same approach you used earlier via a miniature attribution prep.
        df_full = df_cur.copy()
        # Identify columns for current period
        if cfg["time_view"] == "Monthly":
            # take the last month that is inside focus window & exists in df
            focus_months_sorted = sorted(list(focus_month_labels))
            # choose the last label by timeline order within df columns
            # fallback to any intersection present
            mon_labels = [lab.split("-")[1] for lab in focus_months_sorted]
            mon_in_df = [m for m in mon_labels if m in df_full.columns]
            if not mon_in_df:
                return
            m = mon_in_df[-1]
            prev_m_idx = ALL_MONTHS.index(m) - 1
            if prev_m_idx < 0:
                return
            prev_m = ALL_MONTHS[prev_m_idx]
            if prev_m not in df_full.columns:
                return

            # restrict to same Subquestion & Worker Cat to match the major view
            if SUBQ_COL in df_full.columns:
                sub_part = df_full[df_full[SUBQ_COL] == subq_disp]
            else:
                sub_part = df_full.copy()
            sub_part = sub_part[[ENTITY_COL, prev_m, m]].dropna(subset=[prev_m, m], how="all").copy()
            sub_part.rename(columns={prev_m: "Prev", m: "Curr"}, inplace=True)
        else:
            # Quarterly
            qlab = focus_quarter_label.split()[0]  # e.g., "Q2"
            pq_map = {"Q1": None, "Q2": "Q1", "Q3": "Q2", "Q4": "Q3"}
            pq = pq_map.get(qlab)
            if qlab not in df_full.columns or (pq and pq not in df_full.columns):
                return
            if SUBQ_COL in df_full.columns:
                sub_part = df_full[df_full[SUBQ_COL] == subq_disp]
            else:
                sub_part = df_full.copy()
            keep_cols = [ENTITY_COL, qlab] + ([pq] if pq else [])
            sub_part = sub_part[keep_cols].copy()
            if pq:
                sub_part.rename(columns={pq: "Prev", qlab: "Curr"}, inplace=True)
            else:
                # if no previous quarter (Q1), skip contributors
                return

        # Compute contributors (preserve your logic style)
        sub_part["Prev"] = pd.to_numeric(sub_part.get("Prev", np.nan), errors="coerce")
        sub_part["Curr"] = pd.to_numeric(sub_part.get("Curr", np.nan), errors="coerce")
        sub_part["Absolute Change (Contribution)"] = (sub_part["Curr"].fillna(0) - sub_part["Prev"].fillna(0))
        sub_part["% Change"] = (sub_part["Curr"] - sub_part["Prev"]) / sub_part["Prev"].replace(0, np.nan)
        total_diff = sub_part["Absolute Change (Contribution)"].sum(skipna=True)
        if total_diff == 0 or np.isnan(total_diff):
            total_diff = np.nan
        sub_part["% Contribution"] = (sub_part["Absolute Change (Contribution)"] / total_diff) * 100

        # Sort by magnitude of absolute change
        contrib_df = sub_part.rename(columns={ENTITY_COL: "Entity / Group"}).copy()
        contrib_df = contrib_df.sort_values("Absolute Change (Contribution)", ascending=False).reset_index(drop=True)

        # periods list to pass into popover
        periods_for_focus = (sorted(list(focus_month_labels)) if cfg["time_view"] == "Monthly" else [focus_quarter_label])

        render_contributors_with_vr(
            contrib_df=contrib_df,
            vr_df=vr_df,
            dataset_key=cfg["dataset"],
            subq_display=subq_disp,
            periods_for_focus=periods_for_focus
        )

# -------------------- Run --------------------
if __name__ == "__main__":
    main_view()
