"""Audit Log page — dark-header table with ONLINE/OFFLINE badges."""
import sys, os
import html
import re
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar, render_page_nav,
    load_audit_log, df_to_xlsx_bytes, ensure_page_access,
)

inject_css()
init_session_state()
ensure_page_access("audit")
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
<style>
div[data-testid="stVerticalBlock"]:has(.audit-clr-marker) {
    position: relative;
}
div[data-testid="stVerticalBlock"]:has(.audit-clr-marker)
    > div[data-testid="stMarkdown"]:has(.audit-clr-marker) {
    position: absolute;
    height: 0;
    overflow: hidden;
}
div[data-testid="stVerticalBlock"]:has(.audit-clr-marker) > div[data-testid="stButton"] {
    position: absolute;
    top: 1px;
    right: 1px;
    z-index: 10;
}
div[data-testid="stVerticalBlock"]:has(.audit-clr-marker) > div[data-testid="stButton"] button {
    background: transparent !important;
    border: none !important;
    box-shadow: none !important;
    color: #aaa !important;
    font-size: 15px !important;
    font-weight: 300 !important;
    width: 30px !important;
    height: 30px !important;
    min-width: 30px !important;
    min-height: 30px !important;
    padding: 0 !important;
    margin-top: 4px !important;
    margin-right: 4px !important;
    border-radius: 50% !important;
    line-height: 30px !important;
    text-align: center !important;
    cursor: pointer !important;
}
div[data-testid="stVerticalBlock"]:has(.audit-clr-marker) > div[data-testid="stButton"] button:hover {
    color: #333 !important;
    background: rgba(0,0,0,0.07) !important;
}
div[data-testid="stVerticalBlock"]:has(.audit-clr-marker) input {
    padding-right: 38px !important;
}
</style>
<div style="background:#fff;border-radius:14px;border:1px solid #E8E3DC;
            padding:14px 20px;margin-bottom:20px;">""", unsafe_allow_html=True)

ctrl1, ctrl2 = st.columns([2.5, 1.6])
with ctrl1:
    aq = st.text_input("search", placeholder="🔍  ID / Name / Action",
                       label_visibility="collapsed", key="audit_sq")
    if aq:
        st.markdown('<div class="audit-clr-marker"></div>', unsafe_allow_html=True)
        if st.button("✕", key="audit_clr_q"):
            st.session_state["audit_sq"] = ""
            st.rerun()
with ctrl2:
    ad_cal = st.date_input(
        "Date filter",
        value=None,
        key="audit_date_cal",
        label_visibility="collapsed",
        format="DD/MM/YYYY",
    )

st.markdown("</div>", unsafe_allow_html=True)

# ── Filter ────────────────────────────────────────────────────────────────────
show = audit_df.copy()
COLS = [
    "employee_id", "name", "role", "session", "date", "status", "action", "detail",
    "page", "tab", "source", "dg_code", "item_id", "item_name", "field",
    "planogram", "old_value", "new_value",
]
for c in COLS:
    if c not in show.columns:
        show[c] = ""

def _is_edit_action(value) -> bool:
    action = str(value or "").strip().lower()
    if not action:
        return False
    view_prefixes = ("view", "preview", "open")
    return not action.startswith(view_prefixes)

if "action" in show.columns:
    show = show[show["action"].map(_is_edit_action)]

total = len(show)
if aq:
    mask = show.apply(
        lambda r: r.astype(str).str.contains(aq, case=False, na=False).any(), axis=1)
    show = show[mask]
if ad_cal is not None and "date" in show.columns:
    _d_str = ad_cal.strftime("%d/%m/%Y")
    show = show[show["date"].astype(str).str.startswith(_d_str)]

def _build_audit_timestamp(df: pd.DataFrame) -> pd.Series:
    raw = (
        df["date"].fillna("").astype(str).str.strip()
        + " "
        + df["session"].fillna("").astype(str).str.strip()
    ).str.strip()
    return pd.to_datetime(raw, dayfirst=True, errors="coerce")

def _format_session_period(start, end, date_value="", session_value="") -> str:
    if pd.isna(start):
        return f"{str(date_value).strip()} {str(session_value).strip()}".strip()
    if pd.isna(end) or end < start:
        end = start
    if start.date() == end.date():
        return f"{start.strftime('%d/%m/%Y')} {start.strftime('%H.%M')}-{end.strftime('%H.%M')}"
    return f"{start.strftime('%d/%m/%Y %H.%M')}-{end.strftime('%d/%m/%Y %H.%M')}"

def _short_planogram_name(value: str, max_len: int = 64) -> str:
    text = str(value or "").strip()
    if len(text) <= max_len:
        return text
    target = re.search(r"(target[_-]?\d+)", text, flags=re.I)
    if target:
        prefix = text[: max(18, max_len - len(target.group(1)) - 5)].rstrip(" _-")
        return f"{prefix}...{target.group(1)}"
    return text[: max_len - 3].rstrip() + "..."

def _change_verb(old_value: str, new_value: str) -> str:
    old_s = str(old_value or "").strip().lower()
    new_s = str(new_value or "").strip().lower()
    if new_s == "delete":
        return "ลบ"
    if new_s == "new":
        return "เพิ่ม"
    if new_s in ("cleared", "", "nan") or old_s in ("delete", "new"):
        return "คืนค่า"
    return "แก้ไข"

def _summarize_planogram_details(details: list[str], max_lines: int = 6) -> tuple[list[str], int]:
    grouped: dict[tuple[str, str, str, str], list[str]] = {}
    status_lines = []
    parsed_changes = 0

    cell_re = re.compile(
        r"\('([^']*)',\s*'([^']*)',\s*'([^']*)'\)\|([^:;]+):\s*'([^']*)'->'([^']*)'"
    )
    status_re = re.compile(
        r"\('([^']*)',\s*'([^']*)',\s*'([^']*)'\):\s*status\s*->\s*([^;|]+)",
        flags=re.I,
    )

    for detail in details:
        for dg, item_id, item_name, planogram, old_v, new_v in cell_re.findall(detail):
            verb = _change_verb(old_v, new_v)
            key = (verb, dg.strip(), item_id.strip(), item_name.strip())
            pog = _short_planogram_name(planogram)
            if pog and pog not in grouped.setdefault(key, []):
                grouped[key].append(pog)
            parsed_changes += 1
        for dg, item_id, item_name, status in status_re.findall(detail):
            line = f"เปลี่ยน status {dg.strip()}-{item_id.strip()}-{item_name.strip()} เป็น {status.strip()}"
            if line not in status_lines:
                status_lines.append(line)
            parsed_changes += 1

    lines = []
    for (verb, dg, item_id, item_name), planograms in grouped.items():
        if verb == "ลบ":
            prep = "ออกจาก"
        elif verb == "เพิ่ม":
            prep = "เข้า"
        else:
            prep = "ใน"
        shown_pogs = planograms[:4]
        more = len(planograms) - len(shown_pogs)
        pog_text = ", ".join(shown_pogs)
        if more > 0:
            pog_text += f" และอีก {more} planogram"
        lines.append(f"{verb} {dg}-{item_id}-{item_name} {prep} {pog_text}")

    lines.extend(status_lines)
    overflow = max(0, len(lines) - max_lines)
    return lines[:max_lines], overflow

def _summarize_actions(group: pd.DataFrame) -> tuple[str, str]:
    actions = group["action"].fillna("").astype(str).str.strip()
    actions = actions[actions != ""]
    counts = actions.value_counts(sort=False)
    action_text = "; ".join(
        f"{action} x{count}" if count > 1 else action
        for action, count in counts.items()
    )
    details = []
    for detail in group["detail"].fillna("").astype(str):
        detail = detail.strip()
        if not detail or detail.lower() == "nan" or detail in details:
            continue
        details.append(detail)

    concise_lines, overflow = _summarize_planogram_details(details)
    if concise_lines:
        detail_text = " | ".join(concise_lines)
        if overflow:
            detail_text += f" | +{overflow} more item/action"
    else:
        preview = details[:3]
        hidden_count = max(0, int(len(group)) - len(preview))
        detail_text = " | ".join(preview)
        if hidden_count:
            detail_text = f"{detail_text} | +{hidden_count} more" if detail_text else f"+{hidden_count} more"
    return action_text or "Edited", detail_text

if not show.empty:
    show = show.copy()
    show["_audit_ts"] = _build_audit_timestamp(show)
    show = show.sort_values(
        ["employee_id", "name", "role", "_audit_ts"],
        ascending=[True, True, True, True],
        na_position="last",
    ).reset_index(drop=True)
    _user_key = (
        show["employee_id"].fillna("").astype(str)
        + "\x1f" + show["name"].fillna("").astype(str)
        + "\x1f" + show["role"].fillna("").astype(str)
    )
    _prev_user = _user_key.shift()
    _prev_ts = show["_audit_ts"].shift()
    _gap = show["_audit_ts"] - _prev_ts
    _auth_row = show["action"].fillna("").astype(str).str.strip().str.lower().isin(
        ("login", "logout")
    )
    _prev_auth_row = _auth_row.shift(fill_value=False)
    # Approximate login-to-logout sessions from available action timestamps.
    # Business edits are deliberately kept as individual rows so no Range Sheet
    # change is hidden inside a Login/Logout session summary.
    _new_session = (
        (_user_key != _prev_user)
        | show["_audit_ts"].isna()
        | _prev_ts.isna()
        | (_gap > pd.Timedelta(hours=1))
        | (show["_audit_ts"].dt.floor("h") != _prev_ts.dt.floor("h"))
        | ~_auth_row
        | ~_prev_auth_row
    )
    show["_session_bucket"] = _new_session.cumsum()
    grouped_rows = []
    group_cols = ["_session_bucket", "employee_id", "name", "role", "status"]
    for _, group in show.groupby(group_cols, dropna=False, sort=False):
        first = group.iloc[0].copy()
        start = group["_audit_ts"].min()
        end = group["_audit_ts"].max()
        first["date"] = _format_session_period(start, end, first.get("date", ""), first.get("session", ""))
        first["session"] = ""
        first["action"], first["detail"] = _summarize_actions(group)
        first["_audit_ts"] = start
        grouped_rows.append(first)
    show = pd.DataFrame(grouped_rows)
    if "_audit_ts" in show.columns:
        show = show.sort_values("_audit_ts", ascending=False, na_position="last")
else:
    show = show.copy()

show = show.reset_index(drop=True)
shown = len(show)

# ── Table ─────────────────────────────────────────────────────────────────────
def _format_author_role(value) -> str:
    role = str(value or "").strip().lower()
    if role in ("admin", "administrator"):
        return "Admin"
    if role in ("editor", "edit"):
        return "Editor"
    if role in ("viewer", "view"):
        return "Viewer"
    return "Viewer"

_SAMPLE_EMPLOYEE_NAMES = [
    "James Carter",
    "Daniel Cheng",
    "Maya Wilson",
    "Olivia Tan",
    "Ethan Brooks",
    "Sofia Lee",
    "Lucas Martin",
    "Emma Collins",
    "Noah Bennett",
    "Ava Morgan",
    "Liam Turner",
    "Grace Chen",
]

def _display_employee_name(value, employee_id) -> str:
    name = str(value or "").strip()
    if name and name.lower() not in ("developer", "admin", "user", "viewer", "editor"):
        return name
    seed = str(employee_id or "").strip()
    idx = sum(ord(ch) for ch in seed) % len(_SAMPLE_EMPLOYEE_NAMES) if seed else 0
    return _SAMPLE_EMPLOYEE_NAMES[idx]

# Header
th = ""
HEADERS = ["TIME PERIOD", "EMPLOYEE ID", "NAME", "AUTHOR", "STATUS", "LOCATION", "CHANGE"]
for h in HEADERS:
    th += (f'<th style="padding:10px 12px;text-align:left;font-weight:700;'
           f'color:#1A1A1A;font-size:11px;'
           f'white-space:nowrap;border-right:1px solid #E0E0E0;'
           f'background:#F1F3F4;">{h}</th>')

# Rows
rows_html = ""
for i, row in show.iterrows():
    bg = "#fff" if i % 2 == 0 else "#FAFAF8"

    # Timestamp: date + session in one audit-style column.
    date = str(row.get("date", "")).strip()
    session = str(row.get("session", "")).strip()
    timestamp = html.escape(f"{date} {session}".strip())
    td_ts = (f'<td style="padding:8px 12px;border-bottom:1px solid #E0E0E0;'
             f'border-right:1px solid #E0E0E0;font-size:12px;color:#1A1A1A;'
             f'white-space:nowrap;">{timestamp}</td>')

    # Employee ID
    emp_id = str(row.get("employee_id", ""))
    emp_id_html = html.escape(emp_id)
    td_emp = (f'<td style="padding:8px 12px;border-bottom:1px solid #E0E0E0;'
              f'border-right:1px solid #E0E0E0;font-family:monospace;'
              f'font-weight:600;color:#1A1A1A;font-size:12px;white-space:nowrap;">'
              f'{emp_id_html}</td>')

    # Name should read as an employee name, not a position like Developer.
    name = html.escape(_display_employee_name(row.get("name", ""), emp_id))
    td_name = (f'<td style="padding:8px 12px;border-bottom:1px solid #E0E0E0;'
               f'border-right:1px solid #E0E0E0;font-size:12px;color:#1A1A1A;">'
               f'{name}</td>')

    # Author is the user's permission role: Admin / Viewer / Editor.
    author = html.escape(_format_author_role(row.get("role", "") or row.get("author", "")))
    td_author = (f'<td style="padding:8px 12px;border-bottom:1px solid #E0E0E0;'
                 f'border-right:1px solid #E0E0E0;font-size:12px;color:#1A1A1A;'
                 f'white-space:nowrap;">{author}</td>')

    # Status badge
    status_val = html.escape(str(row.get("status", "OFFLINE")).upper().strip())
    if status_val == "ONLINE":
        badge_bg, badge_c = "#E8F8F5", "#2BBFA4"
    else:
        badge_bg, badge_c = "#FFEBEE", "#E05555"
    td_status = (f'<td style="padding:8px 12px;border-bottom:1px solid #E0E0E0;'
                 f'border-right:1px solid #E0E0E0;">'
                 f'<span style="background:{badge_bg};color:{badge_c};font-size:11px;'
                 f'font-weight:700;padding:3px 10px;border-radius:20px;white-space:nowrap;">'
                 f'{status_val}</span></td>')

    # Location + exact change. Structured Range Sheet events show DG, item,
    # field/planogram and old -> new values as separate readable information.
    action_raw = str(row.get("action", "") or "").strip()
    is_rangesheet_edit = action_raw.lower().startswith("rangesheet")
    if is_rangesheet_edit:
        tab_raw = str(row.get("tab", "") or row.get("page", "") or "Range Sheet").strip()
        dg_raw = str(row.get("dg_code", "") or "").strip()
        planogram_raw = str(row.get("planogram", "") or "").strip()
        field_raw = str(row.get("field", "") or "").strip()
        location_parts = [tab_raw]
        if dg_raw and dg_raw.lower() != "nan":
            location_parts.append(f"DG {dg_raw}")
        if planogram_raw and planogram_raw.lower() != "nan":
            location_parts.append(_short_planogram_name(planogram_raw, 48))
        elif field_raw and field_raw.lower() != "nan":
            location_parts.append(field_raw)
        location = html.escape(" · ".join(location_parts))

        item_id_raw = str(row.get("item_id", "") or "").strip()
        item_name_raw = str(row.get("item_name", "") or "").strip()
        old_raw = str(row.get("old_value", "") or "")
        new_raw = str(row.get("new_value", "") or "")
        source_raw = str(row.get("source", "") or "Table").strip()
        item_label = " - ".join(
            value for value in (item_id_raw, item_name_raw)
            if value and value.lower() != "nan"
        )
        change_detail = f"{field_raw}: {old_raw} → {new_raw}" if field_raw else f"{old_raw} → {new_raw}"
        action = html.escape(action_raw.replace("RangeSheet ", ""))
        item_html = html.escape(item_label)
        change_html = html.escape(change_detail)
        source_html = html.escape(source_raw)
        detail_span = (
            f'<div style="color:#374151;font-size:12px;margin-top:3px;">{item_html}</div>'
            f'<div style="color:#64748B;font-size:11px;margin-top:2px;">{change_html}</div>'
            f'<div style="color:#94A3B8;font-size:10px;margin-top:2px;">via {source_html}</div>'
        )
    else:
        location = "Authentication" if action_raw.lower() in ("login", "logout") else html.escape(
            str(row.get("page", "") or "System")
        )
        action = html.escape(action_raw)
        detail = html.escape(str(row.get("detail", "")))
        detail_span = (
            f'<span style="color:#94A3B8;font-size:11px;margin-left:8px;">{detail}</span>'
            if detail and detail not in ("", "nan") else ""
        )

    td_location = (
        f'<td style="padding:8px 12px;border-bottom:1px solid #E0E0E0;'
        f'border-right:1px solid #E0E0E0;font-size:11px;color:#475569;'
        f'max-width:260px;overflow-wrap:anywhere;">{location}</td>'
    )
    td_action = (
        f'<td style="padding:8px 12px;border-bottom:1px solid #E0E0E0;'
        f'font-size:12px;color:#1A1A1A;min-width:300px;">'
        f'<strong>{action}</strong>{detail_span}</td>'
    )

    rows_html += (f'<tr style="background:{bg};">'
                  f'{td_ts}{td_emp}{td_name}{td_author}{td_status}'
                  f'{td_location}{td_action}</tr>')

st.markdown(f"""
<div style="margin-bottom:12px;">
    <span style="font-size:12px;color:#888;">
        Showing <strong>{shown}</strong> audit row(s) from <strong>{total}</strong> entries
    </span>
</div>
<div style="background:#fff;border-radius:14px;border:1px solid #E8E3DC;overflow:hidden;">
    <div style="overflow-x:auto;">
        <table style="width:100%;border-collapse:collapse;min-width:1180px;">
            <thead>
                <tr>{th}</tr>
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
