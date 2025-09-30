# staging_all_fixed.py
from __future__ import annotations
import argparse, time
from pathlib import Path
from typing import List, Tuple
import pandas as pd

# Import each per-question extractor
from extract_q1 import extract_q1_from_file
from extract_q2 import extract_q2_from_file
from extract_q3 import extract_q3_from_file
from extract_q4 import extract_q4_from_file
from extract_q5 import extract_q5_from_file


def main() -> int:
    ap = argparse.ArgumentParser(description="Extract RLMS Q1–Q5 in one pass into a staging workbook (standardized schema).")
    ap.add_argument("--input", required=True, help="Folder with submissions (.xlsx/.xlsm)")
    ap.add_argument("--out",   required=True, help="Output staging workbook (.xlsx)")
    ap.add_argument("--limit", type=int, default=None, help="Limit number of files (debug)")
    args = ap.parse_args()

    root = Path(args.input)
    if not root.exists():
        print(f"[ERROR] Folder not found: {root}")
        return 2

    files: List[Path] = []
    for ext in ("*.xlsx", "*.xlsm"):
        files.extend(p for p in root.rglob(ext) if not p.name.startswith("~$"))
    files.sort()
    if args.limit:
        files = files[:args.limit]
    print(f"[INFO] Files: {len(files)}")

    t0 = time.perf_counter()

    # Collect frames
    q1a_main_acc, q1a_jf_q4_acc, q1b_acc = [], [], []
    q2a_main_acc, q2a_jf_q4_acc, q2b_acc = [], [], []
    q3_acc, q4_acc, q5_acc = [], [], []

    for i, p in enumerate(files, 1):
        # Q1
        try:
            a, jf, b = extract_q1_from_file(p)
            if not a.empty: q1a_main_acc.append(a)
            if not jf.empty: q1a_jf_q4_acc.append(jf)
            if not b.empty: q1b_acc.append(b)
        except Exception as e:
            print(f"[WARN] Q1 skip {p.name}: {e}")

        # Q2
        try:
            a, jf, b = extract_q2_from_file(p)
            if not a.empty: q2a_main_acc.append(a)
            if not jf.empty: q2a_jf_q4_acc.append(jf)
            if not b.empty: q2b_acc.append(b)
        except Exception as e:
            print(f"[WARN] Q2 skip {p.name}: {e}")

        # Q3
        try:
            df = extract_q3_from_file(p)
            if not df.empty: q3_acc.append(df)
        except Exception as e:
            print(f"[WARN] Q3 skip {p.name}: {e}")

        # Q4
        try:
            df = extract_q4_from_file(p)
            if not df.empty: q4_acc.append(df)
        except Exception as e:
            print(f"[WARN] Q4 skip {p.name}: {e}")

        # Q5
        try:
            df = extract_q5_from_file(p)
            if not df.empty: q5_acc.append(df)
        except Exception as e:
            print(f"[WARN] Q5 skip {p.name}: {e}")

        if i % 25 == 0:
            print(f"  processed {i}/{len(files)}")

    # Concatenate
    def _cat(frames: List[pd.DataFrame]) -> pd.DataFrame:
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

    out_q1a_main  = _cat(q1a_main_acc)
    out_q1a_jf_q4 = _cat(q1a_jf_q4_acc)
    out_q1b       = _cat(q1b_acc)

    out_q2a_main  = _cat(q2a_main_acc)
    out_q2a_jf_q4 = _cat(q2a_jf_q4_acc)
    out_q2b       = _cat(q2b_acc)

    out_q3        = _cat(q3_acc)
    out_q4        = _cat(q4_acc)
    out_q5        = _cat(q5_acc)

    # Sorter
    def _sort(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty: return df
        month_cols = [c for c in ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","M"] if c in df.columns]
        base_cols = [c for c in ["entity_name","year","quarter","question","subquestion","worker_category","job_function"] if c in df.columns]
        return df.sort_values(base_cols + month_cols, kind="mergesort")

    out_q1a_main  = _sort(out_q1a_main)
    out_q1a_jf_q4 = _sort(out_q1a_jf_q4)
    out_q1b       = _sort(out_q1b)
    out_q2a_main  = _sort(out_q2a_main)
    out_q2a_jf_q4 = _sort(out_q2a_jf_q4)
    out_q2b       = _sort(out_q2b)
    out_q3        = _sort(out_q3)
    out_q4        = _sort(out_q4)
    out_q5        = _sort(out_q5)

    # Write output
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        if not out_q1a_main.empty:  out_q1a_main.to_excel(xw, index=False, sheet_name="Q1A_Main")
        if not out_q1a_jf_q4.empty: out_q1a_jf_q4.to_excel(xw, index=False, sheet_name="Q1A_JobFunc_Q4")
        if not out_q1b.empty:       out_q1b.to_excel(xw, index=False, sheet_name="Q1B")
        if not out_q2a_main.empty:  out_q2a_main.to_excel(xw, index=False, sheet_name="Q2A_Main")
        if not out_q2a_jf_q4.empty: out_q2a_jf_q4.to_excel(xw, index=False, sheet_name="Q2A_JobFunc_Q4")
        if not out_q2b.empty:       out_q2b.to_excel(xw, index=False, sheet_name="Q2B")
        if not out_q3.empty:        out_q3.to_excel(xw, index=False, sheet_name="Q3")
        if not out_q4.empty:        out_q4.to_excel(xw, index=False, sheet_name="Q4")
        if not out_q5.empty:        out_q5.to_excel(xw, index=False, sheet_name="Q5")
        if all(df.empty for df in [out_q1a_main,out_q1a_jf_q4,out_q1b,out_q2a_main,out_q2a_jf_q4,out_q2b,out_q3,out_q4,out_q5]):
            pd.DataFrame({"message":["No data extracted"]}).to_excel(xw, index=False, sheet_name="EMPTY")

    print(f"[DONE] Wrote staging workbook → {out_path}")
    print(f"[TIME] {time.perf_counter()-t0:0.2f}s")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())