"""Report page — Range execution & summary analytics."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
import pandas as pd
from datetime import datetime
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar,
    find_col, df_to_xlsx_bytes, df_to_csv_bytes,
)

inject_css()
init_session_state()
render_sidebar("report")
render_topbar("Report")

merged = st.session_state.merged_df
if merged is None:
    st.warning("No data — upload files on **My Files** first.")
    st.stop()

rt1, rt2 = st.tabs(["📦 Range Execution", "📈 Summary & Analytics"])

# ── Tab: Range Execution ──────────────────────────────────────────────────────
with rt1:
    st.subheader("PPOG → SSPOG / Non-SSPOG Split")
    pog_col = find_col(merged, ["pog","ppog","sspog","cluster"])
    if not pog_col:
        pog_col = st.selectbox("Select POG column",
                               ["— none —"] + list(merged.columns))
        if pog_col == "— none —":
            pog_col = None

    sspog_sel = []
    if pog_col:
        st.success(f"POG column: **{pog_col}**")
        pog_vals  = sorted(merged[pog_col].dropna().astype(str).unique().tolist())
        sspog_sel = st.multiselect("Values to count as SSPOG", pog_vals,
                       default=[v for v in pog_vals if "ss" in v.lower() and "non" not in v.lower()])
        sspog_df  = merged[merged[pog_col].astype(str).isin(sspog_sel)]
        non_df    = merged[~merged[pog_col].astype(str).isin(sspog_sel)]
    else:
        sspog_df  = merged.iloc[0:0]
        non_df    = merged

    c1, c2 = st.columns(2)
    c1.metric("SSPOG rows",     f"{len(sspog_df):,}")
    c2.metric("Non-SSPOG rows", f"{len(non_df):,}")

    st.divider()
    st.subheader("Output Files")
    OUTPUTS = [
        ("01", "Range by Item by Store",           merged),
        ("02", "Item Store · Status · Forecast",   merged),
        ("03", "Item No. · MOD-Store · Cluster",   merged),
        ("04", "SSPOG File — Citrix Upload",        sspog_df),
    ]
    for num, title, df_out in OUTPUTS:
        st.write(f"**{num}. {title}** — {len(df_out):,} rows")
        cx, cy = st.columns(2)
        with cx:
            st.download_button(f"⬇️ {num} .xlsx", df_to_xlsx_bytes(df_out), f"output_{num}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{num}x", use_container_width=True)
        with cy:
            st.download_button(f"⬇️ {num} .csv", df_to_csv_bytes(df_out), f"output_{num}.csv",
                "text/csv", key=f"dl_{num}c", use_container_width=True)

# ── Tab: Summary & Analytics ──────────────────────────────────────────────────
with rt2:
    st.subheader("Summary & Analytics")
    SUMMARY_VIEWS = [
        "Summary by Status", "Summary by Cluster POG",
        "Summary by POG", "Summary by Store",
        "Financial by Status", "Financial by POG",
    ]
    sel     = st.selectbox("Select Summary View", SUMMARY_VIEWS)
    col_map = {
        "Status":  find_col(merged, ["status"]),
        "Cluster": find_col(merged, ["cluster"]),
        "POG":     find_col(merged, ["pog"]),
        "Store":   find_col(merged, ["store"]),
    }
    group_key = next((k for k in col_map if k.lower() in sel.lower()), None)
    group_col = col_map.get(group_key) if group_key else None
    val_cols  = [c for c in merged.columns if pd.api.types.is_numeric_dtype(merged[c])][:3]

    if group_col:
        try:
            if val_cols:
                summary_df = merged.groupby(group_col)[val_cols].sum().reset_index()
                summary_df["Count"] = merged.groupby(group_col).size().values
            else:
                summary_df = merged[group_col].value_counts().reset_index()
                summary_df.columns = [group_col, "Count"]
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
            num_cols = [c for c in summary_df.columns
                        if c != group_col and pd.api.types.is_numeric_dtype(summary_df[c])][:2]
            if num_cols:
                st.bar_chart(summary_df.set_index(group_col)[num_cols], height=220)
            st.download_button(f"⬇️ Export: {sel}", df_to_xlsx_bytes(summary_df),
                f"{sel.replace(' ','_')}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Cannot create summary: {e}")
    else:
        st.warning(f"Column matching **{sel}** not found in data.")
