    # ---- SUBQUESTION LABELS (exact strings used in QC templates) ----
    # Use the real subquestion labels used by your QC generator / sheets.
    # If your sheets use slightly different text, adjust only these strings.
    SUB = {
        "emp_A": "A. Number of Employees",
        "emp_B1": "B(i). Malaysian Employees",
        "emp_B2": "B(ii). Non-Malaysian Employees",

        "VAC": "A. Number of Job Vacancies as at End of the Month",
        "NEWJOB": "Number of Job Vacancies Due to New Jobs Created During the Month",

        "HIRE": "New Hires and Recalls",

        "QUIT": "A. Quits and resignation (except retirement)",
        "LAYOFF": "B. Total Layoffs and Discharges",
        "OTHER": "C. Other Separation"
    }

    # ---- Robust QC extraction helper ----
    def qc(df, sub_label, month_idx:int):
        """
        Returns the numeric sum for `sub_label` for the given month index (1..12).
        Logic:
          - Prefer rows where the Entity column indicates 'All' (All Financial Institution / All FI)
          - Prefer Worker Category rows that look like 'All', 'Total', 'All workers'
          - If no worker-total row exists, sum across available worker-category rows
        """
        if df is None or isinstance(df, str):
            return 0.0

        # helper to find entity column (common names)
        entity_col = None
        for cand in ["Entity / Group", "Entity", "Entity/Group", "Entity / group"]:
            if cand in df.columns:
                entity_col = cand
                break

        wc_col = None
        for cand in ["Worker Category", "Worker_Category", "WC", "Worker category"]:
            if cand in df.columns:
                wc_col = cand
                break

        month_col = ALL_MONTHS[month_idx - 1]  # "Jan".."Dec"
        # 1) exact subquestion match
        sel = df[df["Subquestion"].astype(str).str.strip() == sub_label] if "Subquestion" in df.columns else df.iloc[0:0]

        # 2) fallback: contains
        if sel.empty:
            sel = df[df["Subquestion"].astype(str).str.contains(sub_label.split()[0], na=False, case=False)]

        if sel.empty:
            # nothing matches
            return 0.0

        # 3) prefer rows where Entity indicates "All Financial Institution"
        if entity_col:
            all_entity_mask = sel[entity_col].astype(str).str.contains(r"all\s*financial|all\s*fi|all\s*institution|^all\b", case=False, na=False)
            if all_entity_mask.any():
                sel = sel[all_entity_mask]

        # 4) if Worker Category column present, prefer rows that are totals
        if wc_col and wc_col in sel.columns:
            wc_vals = sel[wc_col].astype(str).str.strip().fillna("")
            # patterns that likely indicate total row
            total_mask = wc_vals.str.contains(r"all|total|all workers|total workers", case=False, na=False)
            if total_mask.any():
                sel = sel[total_mask]
                # sum numeric month column across matching total rows
                try:
                    return float(pd.to_numeric(sel[month_col], errors="coerce").fillna(0).sum())
                except Exception:
                    return 0.0
            else:
                # no explicit total row — sum across available worker categories
                try:
                    return float(pd.to_numeric(sel[month_col], errors="coerce").fillna(0).sum())
                except Exception:
                    return 0.0
        else:
            # no worker category column — just sum the month column across sel
            try:
                return float(pd.to_numeric(sel[month_col], errors="coerce").fillna(0).sum())
            except Exception:
                return 0.0

    # ---- COLLECT VALUES FOR 3 MONTHS (use correct subquestions and All-FI totals) ----
    month_rows = []
    for m in q_months:
        # Employment: sum of A + B(i) + B(ii)
        emp_A = qc(df_q1, SUB["emp_A"], m)
        emp_B1 = qc(df_q1, SUB["emp_B1"], m)
        emp_B2 = qc(df_q1, SUB["emp_B2"], m)
        EMP = emp_A + emp_B1 + emp_B2

        VAC = qc(df_q4, SUB["VAC"], m)
        NEWJ = qc(df_q4, SUB["NEWJOB"], m)
        HIRE = qc(df_q5, SUB["HIRE"], m)

        QUIT = qc(df_q5, SUB["QUIT"], m)
        LAY = qc(df_q5, SUB["LAYOFF"], m)
        OTH = qc(df_q5, SUB["OTHER"], m)
        SEP = QUIT + LAY + OTH

        month_rows.append([
            m, EMP, VAC, NEWJ, HIRE, SEP, QUIT, LAY, OTH
        ])
