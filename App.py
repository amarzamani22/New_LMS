def build_vr_wc_table_for_entity(vr_df: pd.DataFrame, df_cur: pd.DataFrame, dataset: str, entity: str, subq: str, periods_list: List[str]) -> pd.DataFrame:
    """
    Returns a dataframe of VR justifications by Worker Category for the given entity/subquestion and period list.
    If vr_df is "PENDING" or error string, returns a single-row dataframe.
    """
    if isinstance(vr_df, str):
        return pd.DataFrame([{"Worker Category": "All workers", "FI Justification (selected period)": ("Pending submission" if vr_df == "PENDING" else vr_df)}])

    try:
        mask_ent = (df_cur[ENTITY_COL] == entity)
        if SUBQ_COL in df_cur.columns:
            mask_ent &= (df_cur[SUBQ_COL] == subq)
        wc_list = sorted(df_cur.loc[mask_ent, WC_COL].dropna().unique().tolist())
    except Exception:
        wc_list = ["All workers"]

    rows = []
    for wc_any in wc_list:
        try:
            just_text = find_vr_just_for_periods(
                vr_df=vr_df, dataset=dataset, entity_name=entity, subq=subq, wc=wc_any, periods=periods_list
            )
        except Exception as e:
            just_text = f"(error reading VR: {e})"
        rows.append({"Worker Category": wc_any, "FI Justification (selected period)": just_text})

    show_wc = pd.DataFrame(rows)
    if not show_wc.empty:
        show_wc["__is_total"] = show_wc["Worker Category"].map(lambda x: 0 if _is_total_wc(x) else 1)
        show_wc.sort_values(["__is_total", "Worker Category"], inplace=True)
        show_wc.drop(columns="__is_total", inplace=True)
    return show_wc





# Present a clean table with an interactive "FI Justification" checkbox column
show = dfe_view.copy()
show["Prev"] = show["Prev"].map(lambda v: f"{v:,.0f}")
show["Curr"] = show["Curr"].map(lambda v: f"{v:,.0f}")
show["Delta"] = show["Delta"].map(lambda v: f"{v:+,.0f}")
show["Contribution %"] = show["Contribution %"].map(lambda p: f"{p:+.1f}%" if pd.notna(p) else "–")
show["FI Justification (all WC)"] = False

edited = st.data_editor(
    show,
    use_container_width=True,
    disabled=["Prev","Curr","Delta","Contribution %","FI Justification"],
    column_config={
        "FI Justification (all WC)": st.column_config.CheckboxColumn(
            "FI Justification", help="Click to view all Worker Categories for this FI"
        )
    },
    key=f"fi_contrib_{attrib_period}"
)

# For any rows where the user ticked the FI Justification cell, render a breakdown by Worker Category
selected_rows = edited[edited["FI Justification (all WC)"] == True]
if not selected_rows.empty:
    for _, row in selected_rows.iterrows():
        ent = row["Entity"]
        st.markdown(f"**{ent} — VR Justifications by Worker Category ({attrib_period})**")
        wc_table = build_vr_wc_table_for_entity(
            vr_df=vr_df,
            df_cur=df_cur,
            dataset=cfg["dataset"],
            entity=ent,
            subq=subq,
            periods_list=[attrib_period]
        )
        st.dataframe(wc_table, use_container_width=True)





