"""Audit Log page — dark-header table with ONLINE/OFFLINE badges."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar, render_page_nav,
    load_audit_log, df_to_xlsx_bytes,
)

inject_css()
init_session_state()
render_sidebar("audit")
render_topbar("Audit Log")

st.markdown(
    "<div style='font-size:28px;font-weight:800;color:#1A1A1A;margin-bottom:24px;'>Audit Log</div>",
    unsafe_allow_html=True,
)

# Load from file; fall back to in-memory session entries
import pandas as pd
audit_df = load_audit_log()
if (audit_df is None or audit_df.empty) and st.session_state.get("audit_log"):
    audit_df = pd.DataFrame(st.session_state.audit_log)

if audit_df is None or audit_df.empty:
    st.markdown("""
<div style="background:#fff;border-radius:16px;border:1px solid #E8E3DC;
            padding:60px;text-align:center;">
    <div style="font-size:48px;margin-bottom:16px;">🛡️</div>
    <div style="font-size:16px;font-weight:600;color:#1A1A1A;margin-bottom:8px;">No entries yet</div>
    <div style="font-size:13px;color:#999;">Every action is logged automatically.</div>
</div>
""", unsafe_allow_html=True)
    st.stop()

# ── Search controls ────────────────────────────────────────────────────────────
st.markdown("""
<div style="background:#fff;border-radius:14px;border:1px solid #E8E3DC;
            padding:14px 20px;margin-bottom:20px;">""", unsafe_allow_html=True)

ctrl1, ctrl2, ctrl3 = st.columns([2.5, 1.6, 1])
with ctrl1:
    aq = st.text_input("search", placeholder="🔍  ID / Name / Action",
                       label_visibility="collapsed")
with ctrl2:
    ad_cal = st.date_input(
        "Date filter",
        value=None,
        key="audit_date_cal",
        label_visibility="collapsed",
        format="DD/MM/YYYY",
    )
with ctrl3:
    if st.button("Clear Date", key="audit_clear_date", use_container_width=True):
        st.session_state["audit_date_cal"] = None
        st.rerun()

st.markdown("</div>", unsafe_allow_html=True)

# ── Filter ────────────────────────────────────────────────────────────────────
show = audit_df.copy()
if aq:
    mask = show.apply(
        lambda r: r.astype(str).str.contains(aq, case=False, na=False).any(), axis=1)
    show = show[mask]
if ad_cal is not None and "date" in show.columns:
    _d_str = ad_cal.strftime("%d/%m/%Y")
    show = show[show["date"].astype(str).str.startswith(_d_str)]

show = show.iloc[::-1].reset_index(drop=True)
total = len(audit_df)
shown = len(show)

# ── Table ─────────────────────────────────────────────────────────────────────
COLS = ["employee_id", "name", "session", "date", "status", "action", "detail"]
for c in COLS:
    if c not in show.columns:
        show[c] = ""

# Header
th = ""
HEADERS = ["EMPLOYEE ID", "NAME", "SESSION", "DATE", "STATUS", "ACTION"]
for h in HEADERS:
    th += (f'<th style="padding:12px 16px;text-align:left;font-weight:700;'
           f'color:#fff;font-size:11px;text-transform:uppercase;letter-spacing:0.06em;'
           f'white-space:nowrap;border-right:1px solid rgba(255,255,255,0.06);">{h}</th>')

# Rows
rows_html = ""
for i, row in show.iterrows():
    bg = "#fff" if i % 2 == 0 else "#FAFAF8"

    # Employee ID — teal link style
    emp_id = str(row.get("employee_id", ""))
    td_emp = (f'<td style="padding:12px 16px;border-bottom:1px solid #F0EBE3;'
              f'border-right:1px solid #F0EBE3;font-family:monospace;'
              f'font-weight:700;color:#2BBFA4;font-size:13px;white-space:nowrap;">'
              f'{emp_id}</td>')

    # Name
    name = str(row.get("name", ""))
    td_name = (f'<td style="padding:12px 16px;border-bottom:1px solid #F0EBE3;'
               f'border-right:1px solid #F0EBE3;font-size:13px;color:#1A1A1A;">'
               f'{name}</td>')

    # Session
    session = str(row.get("session", ""))
    td_sess = (f'<td style="padding:12px 16px;border-bottom:1px solid #F0EBE3;'
               f'border-right:1px solid #F0EBE3;font-size:12px;color:#888;white-space:nowrap;">'
               f'{session}</td>')

    # Date
    date = str(row.get("date", ""))
    td_date = (f'<td style="padding:12px 16px;border-bottom:1px solid #F0EBE3;'
               f'border-right:1px solid #F0EBE3;font-size:12px;color:#555;white-space:nowrap;">'
               f'{date}</td>')

    # Status badge
    status_val = str(row.get("status", "OFFLINE")).upper().strip()
    if status_val == "ONLINE":
        badge_bg, badge_c = "#E8F8F5", "#2BBFA4"
    else:
        badge_bg, badge_c = "#FFEBEE", "#E05555"
    td_status = (f'<td style="padding:12px 16px;border-bottom:1px solid #F0EBE3;'
                 f'border-right:1px solid #F0EBE3;">'
                 f'<span style="background:{badge_bg};color:{badge_c};font-size:11px;'
                 f'font-weight:700;padding:3px 10px;border-radius:20px;white-space:nowrap;">'
                 f'{status_val}</span></td>')

    # Action + detail
    action = str(row.get("action", ""))
    detail = str(row.get("detail", ""))
    detail_span = (f'<span style="color:#BBB;font-size:11px;margin-left:8px;">{detail}</span>'
                   if detail and detail not in ("", "nan") else "")
    td_action = (f'<td style="padding:12px 16px;border-bottom:1px solid #F0EBE3;'
                 f'font-size:13px;color:#1A1A1A;">'
                 f'{action}{detail_span}</td>')

    rows_html += (f'<tr style="background:{bg};">'
                  f'{td_emp}{td_name}{td_sess}{td_date}{td_status}{td_action}</tr>')

st.markdown(f"""
<div style="margin-bottom:12px;">
    <span style="font-size:12px;color:#888;">
        Showing <strong>{shown}</strong> of <strong>{total}</strong> entries
    </span>
</div>
<div style="background:#fff;border-radius:14px;border:1px solid #E8E3DC;overflow:hidden;">
    <div style="overflow-x:auto;">
        <table style="width:100%;border-collapse:collapse;min-width:700px;">
            <thead>
                <tr style="background:#1C1C1E;">{th}</tr>
            </thead>
            <tbody>{rows_html}</tbody>
        </table>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)
st.download_button(
    "⬇️ Export Audit Log (.xlsx)", df_to_xlsx_bytes(audit_df),
    "audit_log.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

render_page_nav("audit")
