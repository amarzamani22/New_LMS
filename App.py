# --- VR helpers for "FI Justification" popover (no change to your logic) ---
import re
def _vr_norm_key(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip()).lower()

def vr_collect_by_worker(
    vr_df,               # your loaded VR "Variance" DataFrame (or "PENDING"/error str)
    dataset_key: str,    # e.g. "Q2A: Salary"
    entity_name: str,    # FI name shown in your contributors table
    subq_display: str,   # current Subquestion shown in the page
    periods: list[str],  # e.g. ["2025-Feb", "2025-Mar"] or ["Q1 2025"]
):
    """Return Worker Category × (%Growth, Justification) for this FI+SubQ+periods."""
    if isinstance(vr_df, str):
        # "PENDING" or error message
        msg = "Pending submission" if vr_df == "PENDING" else vr_df
        return pd.DataFrame({"Worker Category": ["—"], "%Growth": [""], "Justification": [msg]})

    if vr_df is None or vr_df.empty:
        return pd.DataFrame({"Worker Category": ["—"], "%Growth": [""], "Justification": ["—"]})

    # Derive Q-code from dataset (e.g., "Q2A: Salary" -> "Q2A")
    qcode = dataset_key.split(":")[0].replace(" ", "").upper()
    ent_norm = _vr_norm_key(entity_name)
    subq_norm = _vr_norm_key(subq_display)

    out_rows = []
    for p in periods:
        if "-" in p and "Q" not in p:
            # monthly: "YYYY-Mmm"
            try:
                yr_str, mon = p.split("-"); yr = int(yr_str); mon3 = mon[:3].lower()
            except Exception:
                continue
            sub = vr_df[
                (vr_df["Year"] == yr) &
                (vr_df["Question"].astype(str).str.upper() == qcode) &
                (vr_df["_ent"] == ent_norm)
            ]
            if subq_norm and subq_norm != _vr_norm_key("N/A"):
                sub = sub[sub["_subq"] == subq_norm]
            sub = sub[sub["_month"].str[:3].str.lower() == mon3]
        else:
            # quarterly: "Qx YYYY"
            try:
                qlab, yr_str = p.split(); qn = int(qlab[1:]); yr = int(yr_str)
            except Exception:
                continue
            sub = vr_df[
                (vr_df["Year"] == yr) &
                (vr_df["_qnum"] == qn) &
                (vr_df["Question"].astype(str).str.upper() == qcode) &
                (vr_df["_ent"] == ent_norm)
            ]
            if subq_norm and subq_norm != _vr_norm_key("N/A"):
                sub = sub[sub["_subq"] == subq_norm]

        if sub.empty:
            continue
        take = sub[["Worker Category","%Growth","Justification"]].copy()
        take["Worker Category"] = take["Worker Category"].astype(str).str.strip()
        out_rows.append(take)

    if not out_rows:
        return pd.DataFrame({"Worker Category": ["—"], "%Growth": [""], "Justification": ["—"]})

    out = pd.concat(out_rows, ignore_index=True)
    # collapse duplicates across months in the same focus
    out = (
        out.groupby(["Worker Category","%Growth"], dropna=False)["Justification"]
           .apply(lambda s: " | ".join(sorted(set(str(v) for v in s if str(v).strip()))))
           .reset_index()
    )
    # sort Total first if present
    out["_r"] = out["Worker Category"].str.contains("total", case=False, na=False).map(lambda x: 0 if x else 1)
    out = out.sort_values(["_r","Worker Category"]).drop(columns=["_r"]).reset_index(drop=True)
    return out

def vr_render_contrib_popovers(
    contrib_df: pd.DataFrame,   # your already-prepared contributors table
    vr_df,                      # VR "Variance" DF (or "PENDING"/error str)
    dataset_key: str,           # same dataset you display (e.g., "Q2A: Salary")
    subq_display: str,          # current Subquestion display text
    periods_for_focus: list[str]# months in focus quarter OR ["Qx YYYY"]
):
    """
    Renders a clickable 'FI Justification' popover/expander beside each FI row,
    showing all Worker Categories (+ Total) justifications for that FI and period(s).
    """
    if contrib_df is None or contrib_df.empty:
        st.info("No contributor data available.")
        return

    # We keep your table render untouched; this only adds a per-row popover area:
    st.caption("Click a row’s **FI Justification** to view all Worker Categories (incl. Total).")

    # Identify the entity column (robust to naming)
    col_entity = next((c for c in contrib_df.columns if c.lower().startswith("entity")), "Entity / Group")
    for i, row in contrib_df.iterrows():
        ent = str(row.get(col_entity, ""))
        cols = st.columns([0.65, 0.35])
        cols[0].markdown(f"**{ent}**")

        # popover if available; fallback to expander for older Streamlit
        try:
            pop = cols[1].popover("FI Justification", use_container_width=True, key=f"vr_pop_{i}")
            with pop:
                vr_tbl = vr_collect_by_worker(
                    vr_df=vr_df,
                    dataset_key=dataset_key,
                    entity_name=ent,
                    subq_display=subq_display,
                    periods=periods_for_focus
                )
                st.dataframe(vr_tbl, use_container_width=True)
        except Exception:
            with cols[1].expander("FI Justification", expanded=False):
                vr_tbl = vr_collect_by_worker(
                    vr_df=vr_df,
                    dataset_key=dataset_key,
                    entity_name=ent,
                    subq_display=subq_display,
                    periods=periods_for_focus
                )
                st.dataframe(vr_tbl, use_container_width=True)



# --- AFTER your existing st.dataframe(contrib_df) for the "Contribution by FI" table ---
# Build the list of periods you used for the focus window (monthly or quarterly)
# Example you likely already have:
# periods_for_focus = sorted(list(focus_month_labels))   # monthly
# periods_for_focus = [focus_quarter_label]              # quarterly

vr_render_contrib_popovers(
    contrib_df=contrib_df,
    vr_df=vr_df,                        # your loaded VR "Variance" DF or "PENDING"
    dataset_key=selected_question,      # or cfg["dataset"] — the same label used on screen
    subq_display=subquestion,           # or subq_disp — whatever you show in the UI
    periods_for_focus=periods_for_focus
)





