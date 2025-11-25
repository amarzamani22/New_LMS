def mhs_page():
    """
    Replaced mhs_page() — quarter-first workflow.
    Paste/replace this function in your App.py where the previous mhs_page lived.
    Relies on existing functions/consts in your app:
      - load_qc_sheet(year, sheet_name)
      - SHEET_MAP  (mapping of question -> sheet name)
      - ALL_MONTHS (['Jan','Feb',...])
    """
    import io
    import openpyxl
    from openpyxl.utils import get_column_letter
    from copy import copy

    st.header("MHS Table Generator — Quarter-based")
    st.write("Choose Year + Quarter. System will add the 3 months for the quarter and the Quarter row automatically.")

    # ------------- UI: Year + Quarter selection ----------------
    # Year selector: show a reasonable range (adjust if needed)
    this_year = datetime.now().year
    years = list(range(this_year - 5, this_year + 3))
    target_year = st.selectbox("Year", years, index=years.index(this_year))
    quarter_label = st.selectbox("Quarter", ["Q1", "Q2", "Q3", "Q4"], index=2)

    st.info("This will insert the three months for the chosen quarter and a Quarter summary row. If any of the 3 months already exist in the MHS template, the operation will abort.")

    # ------------- upload template ----------------
    mhs_file = st.file_uploader("Upload MHS template (.xlsx) — same template used by reports", type=["xlsx"])
    if not mhs_file:
        st.info("Please upload your MHS template file.")
        return

    # anchors (same as before)
    ANCHOR = {
        "year_block_row": 15,        # B15
        "quarter_year_row": 53,      # B53
        "quarter_first_row": 55,     # C55
        "month_year_row": 165,       # B165
        "month_first_row": 165       # C165
    }

    # indicator start column (adjust if your template uses a different start)
    start_col_idx = 4  # Column D

    # convert quarter label to month numbers
    q_idx = int(quarter_label[1])
    quarter_months = [(q_idx - 1) * 3 + i for i in (1, 2, 3)]  # e.g., Q3 -> [7,8,9]

    # Load template via openpyxl (preserve formatting)
    try:
        wb = openpyxl.load_workbook(mhs_file)
        ws = wb[wb.sheetnames[0]]
    except Exception as e:
        st.error(f"Failed to read uploaded template: {e}")
        return

    # utility helpers
    def next_empty_row(sheet, start_row:int, col_letter:str):
        r = start_row
        # If the anchor cell itself already empty, use it (rare)
        while sheet[f"{col_letter}{r}"].value not in (None, ""):
            r += 1
            if r > 5000:
                raise RuntimeError("No empty row found - template unexpected.")
        return r

    def find_month_row(sheet, year:int, month:int):
        """Return row number if a month entry for (year,month) exists; else None."""
        r = ANCHOR["month_first_row"]
        while r < 5000:
            cell = sheet[f"C{r}"].value
            if cell in (None, ""):
                r += 1
                continue
            # try numeric month
            try:
                mm = int(cell)
            except Exception:
                # try match by month name
                try:
                    mm = ALL_MONTHS.index(str(cell)[:3]) + 1
                except Exception:
                    r += 1
                    continue
            # find year above
            found_year = None
            for back in range(0, 12):
                yc = sheet[f"B{r-back}"].value
                if isinstance(yc, int):
                    found_year = yc; break
                if isinstance(yc, str) and yc.strip().isdigit() and len(yc.strip())==4:
                    found_year = int(yc.strip()); break
            if found_year == year and mm == month:
                return r
            r += 1
        return None

    def any_quarter_month_exists(sheet, year:int, months:list) -> bool:
        for m in months:
            if find_month_row(sheet, year, m) is not None:
                return True
        return False

    def copy_row_style(src_row:int, dst_row:int, start_col:int, end_col:int):
        """Copy row height and cell style from src_row to dst_row for columns in [start_col, end_col]."""
        try:
            # row height
            if sheet.row_dimensions.get(src_row) and sheet.row_dimensions.get(src_row).height:
                ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
        except Exception:
            pass
        # copy cell style (font, fill, alignment, border)
        for c in range(start_col, end_col + 1):
            sc = ws.cell(row=src_row, column=c)
            dc = ws.cell(row=dst_row, column=c)
            try:
                dc.font = copy(sc.font)
                dc.fill = copy(sc.fill)
                dc.border = copy(sc.border)
                dc.alignment = copy(sc.alignment)
                # copy number format
                dc.number_format = sc.number_format
            except Exception:
                pass

    # Quick guard: if ANY of the 3 months already exist -> abort (Option C)
    if any_quarter_month_exists(ws, target_year, quarter_months):
        st.error(f"One or more months for {quarter_label} {target_year} already exist in the MHS template. Operation aborted.")
        return

    # Load QC sheets for the selected year using your loader
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

    # Subquestion labels used to select rows (exact strings)
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

    def qc_month_value(df, subq_label, month_idx:int):
        # df expected to have 'Subquestion' column and Jan..Dec columns
        if isinstance(df, str):
            return 0
        try:
            sel = df[df["Subquestion"].astype(str).str.strip() == subq_label]
            if sel.empty:
                sel = df[df["Subquestion"].astype(str).str.contains(subq_label.split()[0], na=False)]
            if sel.empty:
                return 0
            colname = ALL_MONTHS[month_idx-1]  # "Jan".."Dec"
            vals = pd.to_numeric(sel[colname], errors='coerce').fillna(0).sum()
            return float(vals)
        except Exception:
            return 0

    # Prepare all indicator values for each month of the quarter
    months_indicator_rows = []
    for m in quarter_months:
        emp_A = qc_month_value(df_q1, SUBQ_MAP["emp_A"], m)
        emp_B1 = qc_month_value(df_q1, SUBQ_MAP["emp_B1"], m)
        emp_B2 = qc_month_value(df_q1, SUBQ_MAP["emp_B2"], m)
        total_employment = emp_A + emp_B1 + emp_B2

        vac_total = qc_month_value(df_q4, SUBQ_MAP["vac_total"], m)
        vac_newjob = qc_month_value(df_q4, SUBQ_MAP["vac_newjob"], m)
        hires = qc_month_value(df_q5, SUBQ_MAP["hires_recalls"], m)
        sep_quit = qc_month_value(df_q5, SUBQ_MAP["sep_quit"], m)
        sep_layoff = qc_month_value(df_q5, SUBQ_MAP["sep_layoff"], m)
        sep_other = qc_month_value(df_q5, SUBQ_MAP["sep_other"], m)
        sep_total = sep_quit + sep_layoff + sep_other

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
        months_indicator_rows.append((m, indicator_values))

    # PREVIEW
    st.write("Preview (month → values):")
    preview = []
    for m, vals in months_indicator_rows:
        preview.append({"month": m, **{f"c{i+1}": round(v,2) for i,v in enumerate(vals)}})
    st.dataframe(pd.DataFrame(preview), use_container_width=True)

    # Write button
    if st.button(f"Insert {quarter_label} {target_year} (add 3 months + quarter row)"):
        # find a source row to copy styles from (prefer the row just above month_first_row)
        style_src_row = None
        if ws[f"C{ANCHOR['month_first_row'] - 1}"].value not in (None, ""):
            style_src_row = ANCHOR['month_first_row'] - 1
        else:
            # search upward for first non-empty row to copy style
            rtmp = ANCHOR['month_first_row'] - 2
            while rtmp > 1:
                if ws[f"C{rtmp}"].value not in (None, ""):
                    style_src_row = rtmp
                    break
                rtmp -= 1

        # write the three months in chronological order
        write_rows = []
        for m, vals in months_indicator_rows:
            row_to_write = next_empty_row(ws, ANCHOR["month_first_row"], "C")
            # write year label if not present in the block immediately above
            # we only write year label once, when first month of block is added
            # check cell B at row_to_write - 1, if blank, fill target_year at nearest empty B
            # But simpler: ensure target_year exists at the first inserted month row's B
            if ws[f"B{row_to_write}"].value in (None, ""):
                # try to set the year in the row above if empty; otherwise set at row_to_write
                ws[f"B{row_to_write}"].value = target_year

            ws[f"C{row_to_write}"].value = m
            # write indicators starting at start_col_idx (D)
            for i, v in enumerate(vals):
                ws[f"{get_column_letter(start_col_idx + i)}{row_to_write}"].value = v

            # copy style/height from style_src_row if available
            if style_src_row:
                try:
                    copy_row_style(style_src_row, row_to_write, start_col_idx, start_col_idx + len(vals) - 1)
                except Exception:
                    pass

            write_rows.append(row_to_write)

        # After months written, compute quarter sums and insert quarter row
        q_row = next_empty_row(ws, ANCHOR["quarter_first_row"], "C")
        ws[f"C{q_row}"].value = f"{q_idx}Q"
        ws[f"B{q_row}"].value = target_year

        # compute quarter totals for each indicator by summing the just-written month rows
        qvals = []
        for i in range(len(months_indicator_rows[0][1])):
            tot = 0
            for rr in write_rows:
                tot += ws[f"{get_column_letter(start_col_idx + i)}{rr}"].value or 0
            qvals.append(tot)
        for i, v in enumerate(qvals):
            ws[f"{get_column_letter(start_col_idx + i)}{q_row}"].value = v

        # Copy style/height for quarter row from style_src_row if possible
        if style_src_row:
            try:
                copy_row_style(style_src_row, q_row, start_col_idx, start_col_idx + len(qvals) - 1)
            except Exception:
                pass

        st.success(f"{quarter_label} {target_year} months + quarter row inserted successfully.")

    # Year insertion section
    st.markdown("---")
    st.write("Year insertion (optional): insert Year row when all 12 months are present for the selected year.")
    if st.button("Insert Year row (if year complete)"):
        # check present months
        present_months = set()
        r = ANCHOR["month_first_row"]
        while r < 5000:
            val = ws[f"C{r}"].value
            if val in (None, ""):
                r += 1; continue
            try:
                mm = int(val)
            except Exception:
                try:
                    mm = ALL_MONTHS.index(str(val)[:3]) + 1
                except Exception:
                    r += 1; continue
            # find year above
            found_year = None
            for back in range(0, 12):
                ycell = ws[f"B{r-back}"].value
                if isinstance(ycell, int):
                    found_year = ycell; break
                if isinstance(ycell, str) and ycell.strip().isdigit() and len(ycell.strip())==4:
                    found_year = int(ycell.strip()); break
            if found_year == target_year:
                present_months.add(mm)
            r += 1

        if set(range(1,13)).issubset(present_months):
            # compute year totals across months
            yvals = []
            for i in range(len(months_indicator_rows[0][1])):
                s = 0
                r = ANCHOR["month_first_row"]
                while r < 5000:
                    if ws[f"C{r}"].value in (None, ""):
                        r += 1; continue
                    try:
                        mm = int(ws[f"C{r}"].value)
                    except Exception:
                        try:
                            mm = ALL_MONTHS.index(str(ws[f"C{r}"].value)[:3]) + 1
                        except Exception:
                            r += 1; continue
                    # check year for row
                    found_year = None
                    for back in range(0, 12):
                        ycell = ws[f"B{r-back}"].value
                        if isinstance(ycell, int):
                            found_year = ycell; break
                        if isinstance(ycell, str) and ycell.strip().isdigit() and len(ycell.strip())==4:
                            found_year = int(ycell.strip()); break
                    if found_year == target_year:
                        s += ws[f"{get_column_letter(start_col_idx + i)}{r}"].value or 0
                    r += 1
                yvals.append(s)
            yrow = next_empty_row(ws, ANCHOR["year_block_row"], "B")
            ws[f"B{yrow}"].value = target_year
            for i, v in enumerate(yvals):
                ws[f"{get_column_letter(start_col_idx + i)}{yrow}"].value = v
            st.success(f"Year {target_year} row inserted.")
        else:
            st.error(f"Year {target_year} is not complete. Present months: {sorted(list(present_months))}. Cannot insert year row.")
            return

    # final: save to bytes and provide download if any changes were made
    if st.button("Save current workbook and download (no further changes)"):
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        st.success("Workbook saved (download ready).")
        st.download_button(
            label="Download MHS Excel",
            data=out.getvalue(),
            file_name=f"MHS_updated_{target_year}_{quarter_label}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
