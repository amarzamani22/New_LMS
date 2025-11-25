# -----------------------------
# NEW PAGE: MHS Table Generator
# -----------------------------
import io
import openpyxl
from openpyxl.utils import get_column_letter
from datetime import datetime

def mhs_page():
    """
    Streamlit page: MHS Table Generator (separate tab).
    Relies on existing functions/consts in your app:
      - load_qc_sheet(year, sheet_name)
      - SHEET_MAP  (mapping of question -> sheet name)
      - ALL_MONTHS (['Jan','Feb',...])
    Paste this function in App.py and call it from your nav when 'MHS Table' selected.
    """

    st.header("MHS Table Generator — Monthly Highlight Statistics")
    st.write("Select a month to fill into an MHS Excel template. The system will pull values directly from QC sheets (Q1A, Q4, Q5) for the selected year and write Month → Quarter → Year rows automatically.")

    # Allow uploading the MHS template (the same template you use)
    mhs_file = st.file_uploader("Upload MHS template (.xlsx) — the same template structure used by reports", type=["xlsx"])
    if not mhs_file:
        st.info("Upload your MHS template file (e.g. 3.5.12a.xlsx).")
        return

    # Optionally show example path to test (we keep sample path variable if you want to test quickly)
    SAMPLE_TEMPLATE_PATH = "/mnt/data/3.5.12a.xlsx"  # <--- local sample, optional

    # Target month input
    ym = st.text_input("Target month to add (YYYY-MM)", value=datetime.now().strftime("%Y-%m"))
    try:
        target_year = int(ym.split("-")[0])
        target_month = int(ym.split("-")[1])
        if not (1 <= target_month <= 12):
            raise ValueError
    except Exception:
        st.error("Target month must be in YYYY-MM format (e.g. 2025-03).")
        return

    # Fill preference
    fill_option = st.selectbox("Fill behaviour", [
        "Month only",
        "Month + Quarter if complete",
        "Month + Quarter + Year if complete"
    ])

    # Anchor cells as per your template
    ANCHOR = {
        "year_block_row": 15,        # B15 anchors the Year block label area
        "quarter_year_row": 53,      # B53 anchors the Quarter-year area label
        "quarter_first_row": 55,     # C55 first quarter row anchor
        "month_year_row": 165,       # B165 anchors Year in Month block
        "month_first_row": 165       # C165 first month cell
    }

    # Convenience: month name from month number (use your ALL_MONTHS constant)
    try:
        month_name = ALL_MONTHS[target_month - 1]
    except Exception:
        # fallback
        month_name = datetime(2000, target_month, 1).strftime("%b")

    st.write(f"Preparing to add **{month_name} {target_year}**")

    # Utility helpers --------------------------------------------------
    def next_empty_row(sheet, start_row:int, col_letter:str):
        r = start_row
        while sheet[f"{col_letter}{r}"].value not in (None, ""):
            r += 1
            if r > 5000:
                raise RuntimeError("No empty row found - template shape unexpected.")
        return r

    def month_exists_in_template(sheet, year:int, month:int) -> bool:
        # scan month column (C) in month block to see (month,year) pair exists
        r = ANCHOR["month_first_row"]
        while r < 5000:
            cell = sheet[f"C{r}"].value
            if cell is None or cell == "":
                r += 1
                continue
            # if cell equals numeric month label (1..12) or name
            try:
                # if month stored as number
                if int(cell) == month:
                    # find nearest year above (col B)
                    up = r
                    found_year = None
                    for back in range(0, 12):
                        ycell = sheet[f"B{up-back}"].value
                        if isinstance(ycell, int):
                            found_year = ycell
                            break
                        if isinstance(ycell, str) and ycell.strip().isdigit() and len(ycell.strip())==4:
                            found_year = int(ycell.strip()); break
                    if found_year == year:
                        return True
            except Exception:
                # maybe stored as month name (e.g., "Jan"), compare names
                if isinstance(cell, str) and cell.strip().lower().startswith(ALL_MONTHS[month-1][:3].lower()):
                    up = r
                    found_year = None
                    for back in range(0, 12):
                        ycell = sheet[f"B{up-back}"].value
                        if isinstance(ycell, int):
                            found_year = ycell; break
                        if isinstance(ycell, str) and ycell.strip().isdigit() and len(ycell.strip())==4:
                            found_year = int(ycell.strip()); break
                    if found_year == year:
                        return True
            r += 1
        return False

    # Load MHS template workbook via openpyxl (preserve formatting)
    try:
        wb = openpyxl.load_workbook(mhs_file)
        ws = wb[wb.sheetnames[0]]
    except Exception as e:
        st.error(f"Failed to read uploaded template: {e}")
        return

    # Guard: abort if month exists (Option C)
    if month_exists_in_template(ws, target_year, target_month):
        st.error(f"{ym} already exists in the MHS template. Operation aborted (per Option C).")
        return

    # Fetch QC sheets for target year using your existing loader -----------------
    # Uses your load_qc_sheet function and SHEET_MAP values defined in App.py
    try:
        q1a_sheetname = SHEET_MAP['Q1A: Employees']
        q4_sheetname  = SHEET_MAP['Q4: Vacancies']
        q5_sheetname  = SHEET_MAP['Q5: Separations']

        df_q1 = load_qc_sheet(target_year, q1a_sheetname)
        df_q4 = load_qc_sheet(target_year, q4_sheetname)
        df_q5 = load_qc_sheet(target_year, q5_sheetname)
    except Exception as e:
        st.error(f"Failed to load QC sheets for year {target_year}: {e}")
        return

    # Subquestion labels (exact strings extracted from your mapping)
    SUBQ_MAP = {
        "emp_A": "A. Number of Employees",
        "emp_B1": "B(i). Malaysian Employees",
        "emp_B2": "B(ii). Non-Malaysian Employees",

        "vac_total": "A. Number of Job Vacancies as at End of the Month",
        "vac_newjob": "Number of Job Vacancies Due to New Jobs Created During the Month",

        "hires_recalls": "New Hires and Recalls",

        "sep_quit": "A. Quits and resignation (except retirement)",
        "sep_layoff": "B. Total Layoffs and Discharges",
        "sep_other": "C. Other Separation"
    }

    # Helper to pick a numeric value from a QC DF for a subquestion and month
    def qc_month_value(df, subq_label, month_idx:int):
        # df expected to have 'Subquestion' column and Jan..Dec columns
        if isinstance(df, str):
            # load_qc_sheet may return error string
            return 0
        try:
            sel = df[df["Subquestion"].astype(str).str.strip() == subq_label]
            if sel.empty:
                # Sometimes subquestion labels differ slightly; try contains fallback
                sel = df[df["Subquestion"].astype(str).str.contains(subq_label.split()[0], na=False)]
            if sel.empty:
                return 0
            colname = ALL_MONTHS[month_idx-1]  # "Jan".."Dec"
            # sum across any matching rows and return numeric (tolerant)
            vals = pd.to_numeric(sel[colname], errors='coerce').fillna(0).sum()
            return float(vals)
        except Exception:
            return 0

    # Compute monthly indicator values (use qc_month_value)
    emp_A = qc_month_value(df_q1, SUBQ_MAP["emp_A"], target_month)
    emp_B1 = qc_month_value(df_q1, SUBQ_MAP["emp_B1"], target_month)
    emp_B2 = qc_month_value(df_q1, SUBQ_MAP["emp_B2"], target_month)
    total_employment = emp_A + emp_B1 + emp_B2

    vac_total = qc_month_value(df_q4, SUBQ_MAP["vac_total"], target_month)
    vac_newjob = qc_month_value(df_q4, SUBQ_MAP["vac_newjob"], target_month)

    hires = qc_month_value(df_q5, SUBQ_MAP["hires_recalls"], target_month)

    sep_quit = qc_month_value(df_q5, SUBQ_MAP["sep_quit"], target_month)
    sep_layoff = qc_month_value(df_q5, SUBQ_MAP["sep_layoff"], target_month)
    sep_other = qc_month_value(df_q5, SUBQ_MAP["sep_other"], target_month)
    sep_total = sep_quit + sep_layoff + sep_other

    # Final ordered list of indicator values to write horizontally.
    # NOTE: adjust order if your MHS template expects different column order.
    indicator_values = [
        total_employment,
        vac_total,
        vac_newjob,
        hires,
        sep_total,
        sep_quit,
        sep_layoff,
        sep_other
    ]

    st.write("Computed values (will be written left→right into the template indicator columns):")
    preview_df = pd.DataFrame({
        "indicator": [
            "Total Employment", "Vacancies (total)", "Vacancies (new job)", "New hires & recalls",
            "Separations (total)", "Separations (quit)", "Separations (layoff)", "Separations (other)"
        ],
        "value": [round(x, 2) for x in indicator_values]
    })
    st.dataframe(preview_df, use_container_width=True)

    # Where to write indicators in the MHS template?
    # We assume indicators start at column D (col index 4). If your template differs,
    # change start_col_idx accordingly.
    start_col_idx = 4  # Column D

    # Confirm and write
    if st.button("Write MHS row and generate file"):
        # write year label row if not present near month block
        # find first empty row in column B after month_year_row anchor
        try:
            yrow = next_empty_row(ws, ANCHOR["month_year_row"], "B")
            ws[f"B{yrow}"].value = target_year
        except Exception:
            # if B cell already contains year on top of month block, it's fine
            pass

        # write month entry
        mrow = next_empty_row(ws, ANCHOR["month_first_row"], "C")
        ws[f"C{mrow}"].value = target_month

        # write indicator values horizontally starting at start_col_idx
        for i, v in enumerate(indicator_values):
            colletter = get_column_letter(start_col_idx + i)
            ws[f"{colletter}{mrow}"].value = v

        # QUARTER: if requested and quarter complete -> sum quarter months and insert
        def present_months_for_year(sheet, year):
            months_present = set()
            r = ANCHOR["month_first_row"]
            while r < 5000:
                val = sheet[f"C{r}"].value
                if val in (None, ""):
                    r += 1; continue
                try:
                    mm = int(val)
                except Exception:
                    # maybe month name
                    try:
                        mm = ALL_MONTHS.index(str(val)[:3]) + 1
                    except Exception:
                        r += 1; continue
                # get year above
                found_year = None
                for back in range(0, 12):
                    ycell = sheet[f"B{r-back}"].value
                    if isinstance(ycell, int):
                        found_year = ycell; break
                    if isinstance(ycell, str) and ycell.strip().isdigit() and len(ycell.strip())==4:
                        found_year = int(ycell.strip()); break
                if found_year == year:
                    months_present.add(mm)
                r += 1
            return months_present

        if fill_option != "Month only":
            quarter_idx = (target_month - 1) // 3 + 1
            qmonths = {(quarter_idx - 1) * 3 + i for i in (1,2,3)}
            months_present = present_months_for_year(ws, target_year)
            if qmonths.issubset(months_present):
                # compute quarter sums by reading month rows we just wrote (or previously present)
                qvals = []
                for i in range(len(indicator_values)):
                    total = 0
                    # scan for each month in qmonths and add the value at column start_col_idx + i
                    for m in qmonths:
                        r = ANCHOR["month_first_row"]
                        while r < 5000:
                            cellm = ws[f"C{r}"].value
                            if cellm == m:
                                colletter = get_column_letter(start_col_idx + i)
                                total += ws[f"{colletter}{r}"].value or 0
                                break
                            r += 1
                    qvals.append(total)
                # insert quarter row at next empty row under quarter_first_row (col C)
                qrow = next_empty_row(ws, ANCHOR["quarter_first_row"], "C")
                ws[f"C{qrow}"].value = f"{quarter_idx}Q"
                ws[f"B{qrow}"].value = target_year
                for i, v in enumerate(qvals):
                    colletter = get_column_letter(start_col_idx + i)
                    ws[f"{colletter}{qrow}"].value = v

        # YEAR: if requested and all 12 months present -> sum and insert at year block
        if fill_option == "Month + Quarter + Year if complete":
            months_present = present_months_for_year(ws, target_year)
            if set(range(1,13)).issubset(months_present):
                yvals = []
                for i in range(len(indicator_values)):
                    total = 0
                    for m in range(1,13):
                        r = ANCHOR["month_first_row"]
                        while r < 5000:
                            if ws[f"C{r}"].value == m:
                                total += ws[f"{get_column_letter(start_col_idx + i)}{r}"].value or 0
                                break
                            r += 1
                    yvals.append(total)
                yrow = next_empty_row(ws, ANCHOR["year_block_row"], "B")
                ws[f"B{yrow}"].value = target_year
                for i, v in enumerate(yvals):
                    ws[f"{get_column_letter(start_col_idx + i)}{yrow}"].value = v

        # Save workbook to bytes and provide download
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        st.success("MHS template updated. Download ready.")
        st.download_button(
            label="Download updated MHS Excel",
            data=out.getvalue(),
            file_name=f"MHS_updated_{target_year}_{target_month}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# -----------------------------
# How to add this mhs_page to your app navigation:
#
# If you have a page selector in your app (e.g. st.sidebar.selectbox or tabs),
# add an option "MHS Table" and call mhs_page() when selected.
#
# Example:
# page = st.sidebar.selectbox("Page", ["Main","Contribution by FI","MHS Table"])
# if page == "MHS Table":
#     mhs_page()
# else:
#     ...existing logic...
#
# Paste the function above into App.py and call it from your navigation.
# -----------------------------
