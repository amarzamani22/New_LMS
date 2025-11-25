def mhs_page():
    """
    FINAL VERSION:
    User selects YEAR + QUARTER.
    System auto-fills:
        - 3 months
        - quarter row
        - year row if Q4
    Shows preview table (month breakdown + quarter summary).
    Prevents overwriting existing months.
    """
    import io
    import openpyxl
    from openpyxl.utils import get_column_letter
    from copy import copy
    from datetime import datetime

    st.header("MHS Table Generator — Quarter Auto-Fill")

    # ---- USER INPUTS ----
    current_year = datetime.now().year
    years = list(range(current_year - 5, current_year + 2))
    year = st.selectbox("Select Year", years, index=years.index(current_year))
    quarter = st.selectbox("Select Quarter", ["Q1", "Q2", "Q3", "Q4"])

    quarter_map = {
        "Q1": [1, 2, 3],
        "Q2": [4, 5, 6],
        "Q3": [7, 8, 9],
        "Q4": [10, 11, 12],
    }
    q_months = quarter_map[quarter]

    mhs_file = st.file_uploader("Upload MHS Template (.xlsx)", type=["xlsx"])
    if not mhs_file:
        return

    # ---- TEMPLATE ANCHORS ----
    ANCHOR = {
        "year_block": 15,        # B15
        "quarter_block": 55,     # C55
        "month_block": 165       # C165
    }
    start_col = 4  # Column D

    # ---- LOAD TEMPLATE ----
    wb = openpyxl.load_workbook(mhs_file)
    ws = wb[wb.sheetnames[0]]

    # ---- HELPER: FIND NEXT EMPTY ROW ----
    def next_empty(sheet, start_row, col):
        r = start_row
        while sheet[f"{col}{r}"].value not in (None, ""):
            r += 1
        return r

    # ---- HELPER: CHECK IF MONTH ALREADY EXISTS ----
    def month_exists(sheet, y, m):
        r = ANCHOR["month_block"]
        while r < 5000:
            cell = sheet[f"C{r}"].value
            if cell == m:
                # check nearest year above
                for back in range(0, 10):
                    yc = sheet[f"B{r-back}"].value
                    if yc == y:
                        return True
            r += 1
        return False

    # ---- STOP IF ANY MONTH ALREADY EXISTS ----
    for m in q_months:
        if month_exists(ws, year, m):
            st.error(f"❌ Month {m} of {year} already exists in MHS table.")
            return

    # ---- LOAD QC SHEETS ----
    try:
        df_q1 = load_qc_sheet(year, SHEET_MAP["Q1A: Employees"])
        df_q4 = load_qc_sheet(year, SHEET_MAP["Q4: Vacancies"])
        df_q5 = load_qc_sheet(year, SHEET_MAP["Q5: Separations"])
    except Exception as e:
        st.error(f"Failed to load QC sheets: {e}")
        return

    # ---- SUBQUESTION LABELS ----
    SUB = {
        "A": "A. Number of Employees",
        "B1": "B(i). Malaysian Employees",
        "B2": "B(ii). Non-Malaysian Employees",
        "VAC": "A. Number of Job Vacancies as at End of the Month",
        "NEWJOB": "Number of Job Vacancies Due to New Jobs Created During the Month",
        "HIRE": "New Hires and Recalls",
        "QUIT": "A. Quits and resignation (except retirement)",
        "LAYOFF": "B. Total Layoffs and Discharges",
        "OTHER": "C. Other Separation"
    }

    # HELPER: extract QC monthly
    def qc(df, sub, m):
        try:
            row = df[df["Subquestion"] == sub]
            col = ALL_MONTHS[m - 1]
            return float(pd.to_numeric(row[col], errors="coerce").fillna(0).sum())
        except:
            return 0

    # ---- COLLECT VALUES FOR 3 MONTHS ----
    month_rows = []
    for m in q_months:
        A = qc(df_q1, SUB["A"], m)
        B1 = qc(df_q1, SUB["B1"], m)
        B2 = qc(df_q1, SUB["B2"], m)
        EMP = A + B1 + B2

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

    # ---- PREVIEW (MONTH BREAKDOWN) ----
    st.subheader("📅 Monthly Breakdown")
    df_prev_months = pd.DataFrame(month_rows, columns=[
        "Month", "Employment", "Vacancies", "New Jobs",
        "Hires", "Separations Total", "Quits", "Layoff", "Other"
    ])
    st.dataframe(df_prev_months)

    # ---- PREVIEW (QUARTER SUMMARY) ----
    st.subheader("📊 Quarter Summary")
    quarter_totals = df_prev_months.iloc[:, 1:].sum()
    st.dataframe(quarter_totals.to_frame("Value"))

    # ---- WRITE TO EXCEL ----
    if st.button("Write Quarter to MHS Template"):
        # Write 3 months
        for m, *vals in month_rows:
            r = next_empty(ws, ANCHOR["month_block"], "C")
            ws[f"B{r}"] = year
            ws[f"C{r}"] = m
            for i, v in enumerate(vals):
                ws[f"{get_column_letter(start_col + i)}{r}"] = v

        # Write quarter
        qr = next_empty(ws, ANCHOR["quarter_block"], "C")
        ws[f"B{qr}"] = year
        ws[f"C{qr}"] = quarter
        for i, v in enumerate(quarter_totals):
            ws[f"{get_column_letter(start_col + i)}{qr}"] = float(v)

        # If Q4 → write Year row
        if quarter == "Q4":
            yr_r = next_empty(ws, ANCHOR["year_block"], "B")
            year_totals = df_prev_months.iloc[:, 1:].sum()
            ws[f"B{yr_r}"] = year
            for i, v in enumerate(year_totals):
                ws[f"{get_column_letter(start_col + i)}{yr_r}"] = float(v)

        # Export
        out = io.BytesIO()
        wb.save(out)
        st.success("Quarter inserted successfully.")
        st.download_button(
            "Download Updated MHS File",
            data=out.getvalue(),
            file_name=f"MHS_{year}_{quarter}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
