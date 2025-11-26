#!/usr/bin/env python3
from __future__ import annotations
import sys
from pathlib import Path
import pandas as pd

from qc_common import (
    new_empty_workbook, append_about_sheet_last, write_qc_sheet, QCArgs,
    normalize_quarter_label, months_up_to, MONTHS_FULL
)

def _load_sheet(xlsx: Path, name: str) -> pd.DataFrame:
    try:
        return pd.read_excel(xlsx, sheet_name=name, engine="openpyxl")
    except Exception:
        return pd.DataFrame()

def _pick_year(*dfs: pd.DataFrame, explicit: int | None) -> int | None:
    if explicit is not None:
        return int(explicit)
    years = []
    for d in dfs:
        if d is not None and not d.empty and "year" in d.columns:
            years.append(d["year"])
    if not years:
        return None
    s = pd.concat(years, ignore_index=True)
    return int(s.mode().iat[0]) if not s.empty else None

def _months_present(df: pd.DataFrame, up_to_quarter: str) -> list[str]:
    # show only months that exist in df & not beyond current quarter
    allowed = set(months_up_to(up_to_quarter))
    return [m for m in MONTHS_FULL if m in df.columns and m in allowed]

def _ask_path(prompt: str, required: bool = False) -> Path | None:
    while True:
        s = input(prompt).strip()
        if not s and not required:
            return None
        p = Path(s)
        if p.exists():
            return p
        print("  → Path not found. Try again.")

def _ask_float(prompt: str, default: float) -> float:
    s = input(f"{prompt} [{default}]: ").strip()
    if not s:
        return default
    try:
        return float(s)
    except Exception:
        print("  → Not a number. Using default.")
        return default

def _ask_int(prompt: str, default: int | None) -> int | None:
    s = input(f"{prompt}{' ['+str(default)+']' if default is not None else ''}: ").strip()
    if not s:
        return default
    try:
        return int(s)
    except Exception:
        print("  → Not an integer. Leaving empty.")
        return None

def qc_make_all(
    stage: Path,
    out: Path,
    prior_q1: Path | None,
    prior_q2: Path | None,
    prior_q3: Path | None,
    prior_q4: Path | None,
    prior_q5: Path | None,
    th_q1: tuple[float,float,float,float],
    th_q2: tuple[float,float,float,float],
    th_q3: tuple[float,float,float,float],
    th_q4: tuple[float,float,float,float],
    th_q5: tuple[float,float,float,float],
    year: int | None,
):
    # ---- Load available staging sheets ----
    q1a   = _load_sheet(stage, "Q1A_Main")
    q1jf  = _load_sheet(stage, "Q1A_JobFunc_Q4")
    q1b   = _load_sheet(stage, "Q1B")

    q2a   = _load_sheet(stage, "Q2A_Main")
    q2jf  = _load_sheet(stage, "Q2A_JobFunc_Q4")
    q2b   = _load_sheet(stage, "Q2B")

    q3    = _load_sheet(stage, "Q3")
    q4    = _load_sheet(stage, "Q4")
    q5    = _load_sheet(stage, "Q5")

    if all(d.empty for d in [q1a,q1jf,q1b,q2a,q2jf,q2b,q3,q4,q5]):
        print("[WARN] No QC-able sheets in the consolidated workbook.")
        return

    # current quarter from whatever is present
    qseries = pd.concat([d["quarter"] for d in [q1a,q1jf,q1b,q2a,q2jf,q2b,q3,q4,q5] if not d.empty], ignore_index=True)
    current_q = normalize_quarter_label(qseries)

    # decide year (default = modal year across available sheets)
    year = _pick_year(q1a,q1jf,q1b,q2a,q2jf,q2b,q3,q4,q5, explicit=year)
    if year is None:
        print("[ERROR] Could not determine the target year.")
        return

    # filter current year
    def _fy(d):
        return d[d["year"]==year] if (d is not None and not d.empty and "year" in d.columns) else d

    q1a  = _fy(q1a);  q1jf = _fy(q1jf); q1b  = _fy(q1b)
    q2a  = _fy(q2a);  q2jf = _fy(q2jf); q2b  = _fy(q2b)
    q3   = _fy(q3);   q4   = _fy(q4);   q5   = _fy(q5)

    # prior data (prior year only, per question)
    def _prior_df(p: Path | None, sheet: str) -> pd.DataFrame | None:
        if not p: return None
        df = _load_sheet(p, sheet)
        if df.empty: return None
        if "year" in df.columns:
            return df[df["year"]==(year-1)]
        return df

    # thresholds (mom, qoq, yoy, abs)
    mom1,qoq1,yoy1,abs1 = th_q1
    mom2,qoq2,yoy2,abs2 = th_q2
    mom3,qoq3,yoy3,abs3 = th_q3
    mom4,qoq4,yoy4,abs4 = th_q4
    mom5,qoq5,yoy5,abs5 = th_q5

    # create single workbook, write all QC sheets using qc_common so style is identical
    wb = new_empty_workbook()

    # ===== Q1 =====
    if not q1a.empty:
        mshow = [m for m in _months_present(q1a, current_q)]
        write_qc_sheet(QCArgs(
            sheet_name="QC_Q1A_Main",
            df=q1a, wb=wb, year=year, current_q=current_q,
            prior_df=_prior_df(prior_q1, "Q1A_Main"),
            months_to_show=mshow,
            include_job_function=False,
            mom_pct_threshold=mom1, qoq_pct_threshold=qoq1, yoy_pct_threshold=yoy1, abs_cutoff=abs1
        ))
    if not q1jf.empty:
        write_qc_sheet(QCArgs(
            sheet_name="QC_Q1A_JobFunc_Q4",
            df=q1jf, wb=wb, year=year, current_q=current_q,
            prior_df=_prior_df(prior_q1, "Q1A_JobFunc_Q4"),
            months_to_show=[],  # value-based (Q4) in df['value']
            include_job_function=True, yoy_quarters=["Q4"],
            mom_pct_threshold=mom1, qoq_pct_threshold=qoq1, yoy_pct_threshold=yoy1, abs_cutoff=abs1
        ))
    if not q1b.empty:
        poss = [m for m in ["Jun","Dec"] if m in q1b.columns]
        mshow = [m for m in poss if m in months_up_to(current_q)]
        write_qc_sheet(QCArgs(
            sheet_name="QC_Q1B",
            df=q1b, wb=wb, year=year, current_q=current_q,
            prior_df=_prior_df(prior_q1, "Q1B"),
            months_to_show=mshow,
            include_job_function=False,
            mom_pct_threshold=mom1, qoq_pct_threshold=qoq1, yoy_pct_threshold=yoy1, abs_cutoff=abs1
        ))

    # ===== Q2 =====
    if not q2a.empty:
        mshow = [m for m in _months_present(q2a, current_q)]
        write_qc_sheet(QCArgs(
            sheet_name="QC_Q2A_Main",
            df=q2a, wb=wb, year=year, current_q=current_q,
            prior_df=_prior_df(prior_q2, "Q2A_Main"),
            months_to_show=mshow,
            include_job_function=False,
            mom_pct_threshold=mom2, qoq_pct_threshold=qoq2, yoy_pct_threshold=yoy2, abs_cutoff=abs2
        ))
    if not q2jf.empty:
        write_qc_sheet(QCArgs(
            sheet_name="QC_Q2A_JobFunc_Q4",
            df=q2jf, wb=wb, year=year, current_q=current_q,
            prior_df=_prior_df(prior_q2, "Q2A_JobFunc_Q4"),
            months_to_show=[], include_job_function=True, yoy_quarters=["Q4"],
            mom_pct_threshold=mom2, qoq_pct_threshold=qoq2, yoy_pct_threshold=yoy2, abs_cutoff=abs2
        ))
    if not q2b.empty:
        poss = [m for m in ["Jun","Dec"] if m in q2b.columns]
        mshow = [m for m in poss if m in months_up_to(current_q)]
        write_qc_sheet(QCArgs(
            sheet_name="QC_Q2B",
            df=q2b, wb=wb, year=year, current_q=current_q,
            prior_df=_prior_df(prior_q2, "Q2B"),
            months_to_show=mshow,
            include_job_function=False,
            mom_pct_threshold=mom2, qoq_pct_threshold=qoq2, yoy_pct_threshold=yoy2, abs_cutoff=abs2
        ))

    # ===== Q3 (no subquestion, no job function) =====
    if not q3.empty:
        mshow = [m for m in _months_present(q3, current_q)]
        if "subquestion" not in q3.columns:
            q3 = q3.copy(); q3["subquestion"] = ""
        write_qc_sheet(QCArgs(
            sheet_name="QC_Q3",
            df=q3, wb=wb, year=year, current_q=current_q,
            prior_df=_prior_df(prior_q3, "Q3"),
            months_to_show=mshow, include_job_function=False,
            mom_pct_threshold=mom3, qoq_pct_threshold=qoq3, yoy_pct_threshold=yoy3, abs_cutoff=abs3
        ))

    # ===== Q4 =====
    if not q4.empty:
        mshow = [m for m in _months_present(q4, current_q)]
        write_qc_sheet(QCArgs(
            sheet_name="QC_Q4",
            df=q4, wb=wb, year=year, current_q=current_q,
            prior_df=_prior_df(prior_q4, "Q4"),
            months_to_show=mshow, include_job_function=False,
            mom_pct_threshold=mom4, qoq_pct_threshold=qoq4, yoy_pct_threshold=yoy4, abs_cutoff=abs4
        ))

    # ===== Q5 =====
    if not q5.empty:
        mshow = [m for m in _months_present(q5, current_q)]
        write_qc_sheet(QCArgs(
            sheet_name="QC_Q5",
            df=q5, wb=wb, year=year, current_q=current_q,
            prior_df=_prior_df(prior_q5, "Q5"),
            months_to_show=mshow, include_job_function=False,
            mom_pct_threshold=mom5, qoq_pct_threshold=qoq5, yoy_pct_threshold=yoy5, abs_cutoff=abs5
        ))

    # Single About sheet at the end
    append_about_sheet_last(wb, "RLMS – QC (All Questions)", year, current_q)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"\n[DONE] {out}\n")

def main() -> int:
    print("=== RLMS QC – Build ALL Questions (Interactive) ===")

    stage = _ask_path("Path to CURRENT consolidated RLMS (.xlsx): ", required=True)
    out   = Path(input("Path to OUTPUT QC template (.xlsx): ").strip() or "QC_All_Output.xlsx")

    # prior staging (optional, press Enter to skip)
    print("\n--- Optional: prior staging workbooks (used for YoY) ---")
    prior_q1 = _ask_path("  Prior staging for Q1 (.xlsx) [Enter to skip]: ", required=False)
    prior_q2 = _ask_path("  Prior staging for Q2 (.xlsx) [Enter to skip]: ", required=False)
    prior_q3 = _ask_path("  Prior staging for Q3 (.xlsx) [Enter to skip]: ", required=False)
    prior_q4 = _ask_path("  Prior staging for Q4 (.xlsx) [Enter to skip]: ", required=False)
    prior_q5 = _ask_path("  Prior staging for Q5 (.xlsx) [Enter to skip]: ", required=False)

    # thresholds (defaults shown; you can customize per question)
    print("\n--- Thresholds (press Enter to accept defaults) ---")
    def ask_block(tag: str, d_mom=0.25, d_qoq=0.25, d_yoy=0.25, d_abs=50.0):
        print(f"  {tag}:")
        mom = _ask_float("    MoM % threshold", d_mom)
        qoq = _ask_float("    QoQ % threshold", d_qoq)
        yoy = _ask_float("    YoY % threshold", d_yoy)
        abv = _ask_float("    Absolute cutoff", d_abs)
        return (mom, qoq, yoy, abv)

    th_q1 = ask_block("Q1", 0.25, 0.25, 0.25, 50.0)
    th_q2 = ask_block("Q2", 0.25, 0.25, 0.25, 50.0)
    th_q3 = ask_block("Q3", 0.25, 0.25, 0.25, 50.0)
    th_q4 = ask_block("Q4", 0.25, 0.25, 0.25, 50.0)
    th_q5 = ask_block("Q5", 0.25, 0.25, 0.25, 50.0)

    # optional: lock to a specific year (else auto from staging)
    y = _ask_int("\nTarget YEAR (leave blank to auto-detect from staging)", None)

    # run
    qc_make_all(
        stage=stage, out=out,
        prior_q1=prior_q1, prior_q2=prior_q2, prior_q3=prior_q3, prior_q4=prior_q4, prior_q5=prior_q5,
        th_q1=th_q1, th_q2=th_q2, th_q3=th_q3, th_q4=th_q4, th_q5=th_q5,
        year=y
    )
    return 0

if __name__ == "__main__":
    sys.exit(main())
