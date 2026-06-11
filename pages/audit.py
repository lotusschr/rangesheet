"""Audit Log page."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar,
    load_audit_log, df_to_xlsx_bytes,
)

inject_css()
init_session_state()
render_sidebar("audit")
render_topbar("Audit Log")

audit_df = load_audit_log()
if audit_df is None or audit_df.empty:
    st.info("No entries yet — every action is logged automatically.")
    st.stop()

c1, c2 = st.columns([2,1])
with c1:
    aq = st.text_input("Search", placeholder="🔎 ID / name / action...",
                       label_visibility="collapsed")
with c2:
    ad = st.text_input("Date", placeholder="📅 dd/mm/yyyy",
                       label_visibility="collapsed")

show = audit_df.copy()
if aq:
    mask = show.apply(lambda r: r.astype(str).str.contains(aq, case=False, na=False).any(), axis=1)
    show = show[mask]
if ad and "date" in show.columns:
    show = show[show["date"].astype(str).str.contains(ad, na=False)]

st.caption(f"Showing {len(show)} of {len(audit_df)} entries")
st.dataframe(show.iloc[::-1].reset_index(drop=True),
             use_container_width=True, hide_index=True)
st.download_button("⬇️ Export Audit Log (.xlsx)", df_to_xlsx_bytes(audit_df),
    "audit_log.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
