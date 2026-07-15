"""Shared utilities, CSS, auth, data processing — RangeSheet."""

import streamlit as st
import pandas as pd
import re as _re
import os
import io
import json
from datetime import datetime
from difflib import SequenceMatcher

# ── Paths ─────────────────────────────────────────────────────────────────────
_HERE    = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(_HERE)
BASE_DIR = os.path.join(ROOT_DIR, "rangesheet_data")

# ── Constants ─────────────────────────────────────────────────────────────────
APP_CONFIG = {"allowed_extensions": ["csv", "txt", "xlsx", "xls", "xlsb"], "max_file_mb": 0}

RS_SHEETS = [
    "Range Sheet_SSPOG", "Range Sheet_Non-SSPOG",
    "StoreApply_SSPOG", "5.1 ItembyStore", "5.2 ItembyStore_SC",
    "5.3 Upload_product_library", "5.4 Upload to Citrix",
]
RS_TYPE_ORDER = ["MAINTAIN", "DELETE SOME", "DELETE ALL", "NEW SOME", "NEWNEW"]

STATUS_COLORS = {
    "MAINTAIN":        {"bg": "#E8F8F5", "c": "#2BBFA4"},
    "NEW SOME":        {"bg": "#E3F2FD", "c": "#1565C0"},
    "NEW":             {"bg": "#E8F5E9", "c": "#2E7D32"},
    "NEWNEW":          {"bg": "#F3E5F5", "c": "#6A1B9A"},
    "DELETE SOME":     {"bg": "#FEE8E8", "c": "#E05555"},
    "DELETE ALL":      {"bg": "#FFEBEE", "c": "#B71C1C"},
    "NEW DELETE SOME": {"bg": "#FFF3E0", "c": "#E65100"},
}

NAV_ITEMS = [
    ("landpage",  "🗂️", "My Files"),
    ("viewdata",  "📁", "View Data"),
    ("rangesheetreview",  "🔍", "Rangesheet Review"),
    # ("dashboard", "📊", "Dashboard"),
    ("report",    "📋", "Report"),
    ("audit",     "🛡️", "Audit Log"),
]

RS_COL_GROUPS = [
    {"group": "Item Info", "color": "#D9D9D9", "hdr_color": "#333333", "cols": [
        "Department", "Section", "Subclass", "Barcode", "TPNA", "ID",
        "No. of Unit in Case", "No. of Unit in Inner", "Tray total number",
        "Express Picking Type", "HDET Picking Type",
        "EDLP Price by Format", "Item Name",
    ]},
    {"group": "Range Info", "color": "#E8E3DC", "hdr_color": "#444444", "cols": [
        "AS IS planograms applied",
        "TO-BE planograms applied",
        "AS-IS Stores applied",
        "TO-Be stores applied",
        "Avg Units 52wk/ Forecast new item sales",
        "Supplier Pack Size",
        "Range Tail YYYY",
        "AVG Selling Price by Format",
    ]},
    {"group": "Star Line", "color": "#000000", "hdr_color": "#FFFFFF", "cols": [
        "Star Line",
    ]},
    {"group": "Status", "color": "#B8D4E8", "hdr_color": "#1A4A6B", "cols": [
        "Status", "Check Range To-be Waterfall", "Planogram Name",
    ]},
    {"group": "Priority", "color": "#00CC44", "hdr_color": "#003300", "cols": [
        "Item Priority", "JDA vs Actual", "Actual-Actual",
    ]},
    {"group": "Cluster Summary", "color": "#C9A0DC", "hdr_color": "#3D0070", "cols": [
        "To be stores applied count", "AS IS",
        "MODS", "Fixtures", "Range Class", "Total New SKUs", "Total Delete SKUs",
        "%Achieving CRD case (AS is)", "%Achieving LRD (AS is)",
    ]},
]

FILL_COLORS = {
    "display": {"bg": "#FFFDE7", "text": "#5D4037"},
    "mer":     {"bg": "#FCE4EC", "text": "#880E4F"},
    "formula": {"bg": "#F5F5F5", "text": "#616161"},
}

COLUMN_LABELS = {
    "department":                                "Department",
    "section":                                   "Section",
    "subclass":                                  "Subclass",
    "barcode":                                   "Barcode",
    "tpna":                                      "TPNA",
    "id":                                        "ID",
    "no. of unit in case":                       "No. of Unit in Case",
    "no. of unit in inner":                      "No. of Unit in Inner",
    "tray total number":                         "Tray total number",
    "express picking type":                      "Express Picking Type",
    "hdet picking type":                         "HDET Picking Type",
    "edlp price by format":                      "EDLP Price by Format",
    "item name":                                 "Item Name",
    "as is planograms applied":                  "AS IS planograms applied",
    "to-be planograms applied":                  "TO-BE planograms applied",
    "as-is stores applied":                      "AS-IS Stores applied",
    "to-be stores applied":                      "TO-Be stores applied",
    "avg units 52wk/ forecast new item sales":   "Avg Unit 52 wk/forecast new item sales",
    "supplier pack size":                        "Supplier pack size",
    "range tail yyyy":                           "Range Tail YYYY",
    "avg selling price by format":               "AVG selling Price by format",
    "star line":                                 "Star Line",
    "item priority":                             "Item priority",
    "jda vs actual":                             "JDA vs Actual",
    "actual-actual":                             "Actual-Actual",
    "status":                                    "Status",
    "check range to-be waterfall":               "Check Range to be waterfall",
    "name":                                      "Planogram Name",
}

def normalize_col(col):
    return str(col).strip().lower()

TARGET_COLUMNS = list(COLUMN_LABELS.values())

def find_best_match(upload_col, target_cols, threshold=0.75):
    upload_col = normalize_col(upload_col)
    best_score = 0
    best_match = None
    for target in target_cols:
        score = SequenceMatcher(None, upload_col, normalize_col(target)).ratio()
        if score > best_score:
            best_score = score
            best_match = target
    if best_score >= threshold:
        return best_match
    return None

def build_column_mapping(df):
    mapping = {}
    questions = []
    target_cols = list(COLUMN_LABELS.values())
    for col in df.columns:
        match = find_best_match(col, target_cols)
        if match:
            mapping[col] = match
        else:
            questions.append(col)
    return mapping, questions

def create_review_df(df, mapping):
    review_df = pd.DataFrame()
    for source_col, target_col in mapping.items():
        review_df[target_col] = df[source_col]
    return review_df

# ── Auth ──────────────────────────────────────────────────────────────────────
def current_user() -> dict:
    return st.session_state.get("auth_user", {
        "employee_id": "TH111111", "name": "Developer",
        "role": "admin", "login_time": datetime.now().strftime("%H:%M"),
    })

def is_admin() -> bool:
    return current_user().get("role") == "admin"

def logout():
    st.session_state.auth_user = None

def add_audit(action: str, detail: str = ""):
    u = current_user()
    record = {
        "employee_id": u["employee_id"],
        "name":        u["name"],
        "role":        u["role"].upper(),
        "session":     datetime.now().strftime("%H:%M"),
        "date":        datetime.now().strftime("%d/%m/%Y"),
        "status":      "ONLINE",
        "action":      action,
        "detail":      detail,
    }
    if "audit_log" not in st.session_state:
        st.session_state.audit_log = []
    st.session_state.audit_log.append(record)
    _save_audit_record(record)

def _save_audit_record(record: dict):
    try:
        _ensure_dirs()
        path = os.path.join(BASE_DIR, "audit", "audit_log.csv")
        df_new = pd.DataFrame([record])
        if os.path.exists(path):
            df_old = pd.read_csv(path, encoding="utf-8-sig")
            pd.concat([df_old, df_new], ignore_index=True).to_csv(
                path, index=False, encoding="utf-8-sig")
        else:
            df_new.to_csv(path, index=False, encoding="utf-8-sig")
    except Exception:
        pass

# ── Session state ─────────────────────────────────────────────────────────────
def init_session_state():
    defaults = {
        "raw_files": [], "merged_df": None, "merge_log": [], "audit_log": [],
        "display_cols": None, "canvases": {}, "rangesheet_meta": {},
        "view_sheet": RS_SHEETS[1], "view_vis_cols": None,
        "last_uploaded_preview": None,
        "selected_files": [],
        "review_df": None,
        "column_mapping": {},
        "mapping_completed": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v
    if st.session_state.get("merged_df") is None:
        snap = load_merged_snapshot()
        if snap is not None:
            st.session_state.merged_df = snap
            st.session_state.merge_log = [f"💾 Loaded snapshot ({len(snap):,} rows)"]
    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = 0

# ── CSS ───────────────────────────────────────────────────────────────────────
def inject_css():
    st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Sarabun:wght@400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', 'Sarabun', system-ui, sans-serif !important;
}

.main { background: #EDE8DF !important; }
.block-container { padding: 0 !important; max-width: 100% !important; }
[data-testid="stMainBlockContainer"] {
    padding: 0 2rem 2rem 2rem !important;
    margin-top: 0 !important;
}

section[data-testid="stSidebar"] {
    background-color: #1C1C1E !important;
    border-right: none !important;
}
section[data-testid="stSidebar"] > div {
    background-color: #1C1C1E !important;
}
[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
    padding-top: 0 !important;
    margin-top: 0 !important;
}
[data-testid="stSidebar"] p { color: rgba(255,255,255,0.45) !important; margin: 0 !important; }
[data-testid="stSidebar"] .stMarkdown { color: rgba(255,255,255,0.45) !important; }

[data-testid="stSidebar"] .stButton > button {
    background: transparent !important;
    color: rgba(255,255,255,0.7) !important;
    border: none !important;
    border-radius: 10px !important;
    font-size: 14px !important;
    width: 100% !important;
    text-align: left !important;
    justify-content: flex-start !important;
    padding: 10px 16px !important;
    margin: 1px 0 !important;
    font-weight: 500 !important;
    transition: background 0.15s, color 0.15s;
    box-shadow: none !important;
}
[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(255,255,255,0.08) !important;
    color: #fff !important;
    box-shadow: none !important;
}
[data-testid="stSidebar"] .stButton > button:active,
[data-testid="stSidebar"] .stButton > button:focus {
    background: rgba(255,255,255,0.08) !important;
    box-shadow: none !important;
    outline: none !important;
}

.stButton > button {
    background: #2BBFA4 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: 13px !important;
    padding: 8px 18px !important;
}
.stButton > button:hover { background: #22A08A !important; }

.stDownloadButton > button {
    background: transparent !important;
    color: #2BBFA4 !important;
    border: 1.5px solid #2BBFA4 !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 12px !important;
    padding: 5px 14px !important;
}
.stDownloadButton > button:hover {
    background: #E8F8F5 !important;
}

[data-testid="stMetric"] {
    background: #fff !important;
    border: 1px solid #E8E3DC !important;
    border-radius: 16px !important;
    padding: 20px 20px 14px !important;
    box-shadow: 0 1px 4px rgba(0,0,0,.05) !important;
}
[data-testid="stMetricValue"] {
    color: #1A1A1A !important;
    font-weight: 800 !important;
    font-size: 28px !important;
}
[data-testid="stMetricLabel"] {
    color: #888 !important;
    font-size: 11px !important;
    font-weight: 700 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.05em !important;
}
[data-testid="stMetricDelta"] { font-size: 12px !important; }

.stTabs [data-baseweb="tab-list"] {
    border-bottom: 2px solid #E0D9D2 !important;
    gap: 4px;
    background: transparent !important;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px 8px 0 0 !important;
    font-weight: 600 !important;
    color: #888 !important;
    background: transparent !important;
    padding: 8px 16px !important;
}
.stTabs [aria-selected="true"] {
    color: #2BBFA4 !important;
    border-bottom: 2px solid #2BBFA4 !important;
    background: transparent !important;
}

[data-testid="stExpander"] {
    border: 1px solid #E0D9D2 !important;
    border-radius: 12px !important;
    background: #fff !important;
}

[data-testid="stFileUploaderDropzone"] {
    background: white !important;
    border: 2px dashed #D0CAC2 !important;
    border-radius: 16px !important;
    padding: 40px 24px 32px !important;
    text-align: center !important;
    cursor: pointer !important;
    min-height: 200px !important;
    transition: border-color 0.2s, background 0.2s !important;
    display: flex !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center !important;
}
[data-testid="stFileUploaderDropzone"]:hover {
    border-color: #2BBFA4 !important;
    background: #F0FDF9 !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] > div > span {
    font-size: 15px !important;
    font-weight: 700 !important;
    color: #1A1A1A !important;
    text-transform: uppercase !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] > div > small {
    font-size: 11px !important;
    color: #999 !important;
    text-transform: uppercase !important;
    line-height: 2 !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] svg {
    width: 36px !important;
    height: 36px !important;
    opacity: 0.45 !important;
    color: #2BBFA4 !important;
}

[data-testid="stDataFrame"] {
    border-radius: 12px !important;
    border: 1px solid #E0D9D2 !important;
}

[data-testid="stTextInput"] input {
    border-radius: 10px !important;
    border-color: #E0D9D2 !important;
    background: #fff !important;
}
[data-testid="stSelectbox"] > div {
    border-radius: 10px !important;
}

[data-testid="stProgress"] > div {
    border-radius: 99px !important;
}

hr { border-color: #E0D9D2 !important; margin: 8px 0 !important; }

div[data-testid="stRadio"] > div[role="radiogroup"] {
    flex-wrap: nowrap !important;
    gap: 0 !important;
    border-bottom: 2px solid #E0D9D2 !important;
    margin-bottom: 16px !important;
}
div[data-testid="stRadio"] > div[role="radiogroup"] > label {
    border-bottom: 3px solid transparent !important;
    margin-bottom: -2px !important;
    padding: 8px 14px 8px !important;
    background: transparent !important;
    font-size: 12px !important;
    border-radius: 0 !important;
    white-space: nowrap !important;
    cursor: pointer !important;
    color: #888 !important;
    transition: border-color 0.15s, color 0.15s !important;
}
div[data-testid="stRadio"] > div[role="radiogroup"] > label:hover {
    color: #1A1A1A !important;
    border-bottom-color: #AADDD5 !important;
}
div[data-testid="stRadio"] > div[role="radiogroup"] > label:has(input:checked) {
    border-bottom: 3px solid #2BBFA4 !important;
    color: #1A1A1A !important;
    font-weight: 700 !important;
}
div[data-testid="stRadio"] > div[role="radiogroup"] > label > div:first-child {
    display: none !important;
}

.page-nav-btn > button {
    background: transparent !important;
    color: #2BBFA4 !important;
    border: 1.5px solid #2BBFA4 !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: 13px !important;
}
.page-nav-btn > button:hover {
    background: #E8F8F5 !important;
}

.stButton > button,
.stDownloadButton > button { text-transform: uppercase !important; }
.stTabs [data-baseweb="tab"] { text-transform: uppercase !important; }
div[role="radiogroup"] > label > div:last-child p { text-transform: uppercase !important; }
[data-testid="stWidgetLabel"] p,
[data-testid="stWidgetLabel"] label { text-transform: uppercase !important; }
[data-testid="stMetricLabel"] { text-transform: uppercase !important; }
[data-testid="stSidebar"] .stMarkdown p { text-transform: uppercase !important; }

#MainMenu, footer { visibility: hidden; }
header { visibility: hidden; }
header button { visibility: visible !important; }
[data-testid="stSidebarCollapseButton"] { visibility: visible !important; }
[data-testid="collapsedControl"] { visibility: visible !important; display: flex !important; }
[data-testid="stToolbar"] { visibility: hidden; }
.stDeployButton { display: none; }
[data-testid="stSidebarNav"] { display: none !important; }

.dlg-yes-btn button {
    background: #fff !important;
    border: 1.5px solid #C8C0B8 !important;
    color: #1A1A1A !important;
}
.dlg-yes-btn button:hover {
    background: #F5F0EA !important;
    border-color: #A8A09A !important;
}
.dlg-no-btn button {
    background: #E53935 !important;
    border-color: #C62828 !important;
    color: #fff !important;
}
.dlg-no-btn button:hover {
    background: #C62828 !important;
    border-color: #B71C1C !important;
}
.dup-add-btn button {
    background: #fff !important;
    border: 1.5px solid #C8C0B8 !important;
    color: #1A1A1A !important;
}
.dup-add-btn button:hover {
    background: #F5F0EA !important;
    border-color: #A8A09A !important;
}
.dup-replace-btn button {
    background: #2BBFA4 !important;
    border-color: #1EA891 !important;
    color: #fff !important;
}
.dup-replace-btn button:hover {
    background: #1EA891 !important;
    border-color: #178A78 !important;
}
</style>
""", unsafe_allow_html=True)

# ── Page navigation (Next / Previous) ────────────────────────────────────────
# _PAGE_ORDER  = ["landpage", "rawfiles", "viewdata", "dashboard", "report", "audit"]
_PAGE_ORDER  = ["landpage", "viewdata", "rangesheetreview", "report", "audit"]
_PAGE_LABELS = {
    "landpage":  "My Files",
    "viewdata":  "View Data",
    "rangesheetreview":  "Rangesheet Review",
    # "dashboard": "Dashboard",
    "report":    "Report",
    "audit":     "Audit Log",
}

def render_page_nav(current: str):
    idx       = _PAGE_ORDER.index(current) if current in _PAGE_ORDER else 0
    prev_page = _PAGE_ORDER[idx - 1] if idx > 0 else None
    next_page = _PAGE_ORDER[idx + 1] if idx < len(_PAGE_ORDER) - 1 else None

    st.markdown("<div style='height:48px;'></div>", unsafe_allow_html=True)
    st.markdown(
        "<hr style='border-color:#E0D9D2;margin-bottom:16px;'>",
        unsafe_allow_html=True,
    )

    _, c_prev, c_next = st.columns([6, 1.6, 1.6])

    with c_prev:
        if prev_page:
            st.markdown('<div class="page-nav-btn">', unsafe_allow_html=True)
            if st.button(f"← {_PAGE_LABELS[prev_page]}", key="page_nav_prev",
                         use_container_width=True):
                st.switch_page(f"pages/{prev_page}.py")
            st.markdown("</div>", unsafe_allow_html=True)

    with c_next:
        if next_page:
            st.markdown('<div class="page-nav-btn">', unsafe_allow_html=True)
            if st.button(f"{_PAGE_LABELS[next_page]} →", key="page_nav_next",
                         use_container_width=True):
                st.switch_page(f"pages/{next_page}.py")
            st.markdown("</div>", unsafe_allow_html=True)


# ── Sidebar ───────────────────────────────────────────────────────────────────
def render_sidebar(active: str = "landpage"):
    with st.sidebar:
        st.markdown("""
<div style="padding:20px 16px 14px;display:flex;align-items:center;gap:12px;">
    <div style="width:42px;height:42px;background:#E8A020;border-radius:10px;
                display:flex;align-items:center;justify-content:center;
                font-size:22px;flex-shrink:0;">📄</div>
    <div>
        <div style="color:#fff;font-weight:700;font-size:16px;line-height:1.3;">RangeSheet</div>
        <div style="color:rgba(255,255,255,0.4);font-size:11px;margin-top:1px;">Management Platform</div>
    </div>
</div>
<div style="height:1px;background:rgba(255,255,255,0.08);margin:0 14px 10px;"></div>
""", unsafe_allow_html=True)

        for page_id, icon, label in NAV_ITEMS:
            if page_id == active:
                st.markdown(
                    f'<div style="background:#2BBFA4;border-radius:10px;'
                    f'padding:10px 16px;margin:2px 6px;'
                    f'display:flex;align-items:center;gap:10px;">'
                    f'<span style="font-size:15px;">{icon}</span>'
                    f'<span style="color:#fff;font-weight:600;font-size:14px;">{label}</span>'
                    f'</div>',
                    unsafe_allow_html=True)
            else:
                if st.button(f"{icon}  {label}", key=f"nav_{page_id}",
                             use_container_width=True):
                    st.switch_page(f"pages/{page_id}.py")

        st.markdown("""
<div style="height:1px;background:rgba(255,255,255,0.08);margin:10px 14px 8px;"></div>
""", unsafe_allow_html=True)

# ── Logout confirmation dialog ─────────────────────────────────────────────────
@st.dialog("Confirm")
def _logout_dialog():
    st.markdown("""
<div style="text-align:center;padding:6px 0 20px;">
    <div style="width:56px;height:56px;background:#FFE0E0;border-radius:50%;
                display:flex;align-items:center;justify-content:center;
                font-size:26px;margin:0 auto 16px;">⚠️</div>
    <div style="font-size:18px;font-weight:700;color:#1A1A1A;line-height:1.4;margin-bottom:8px;">
        Do you really want to exit the app?
    </div>
    <div style="font-size:13px;color:#888;line-height:1.5;">
        All of the unsaved progress would be lost!
    </div>
</div>""", unsafe_allow_html=True)
    _yc, _nc = st.columns(2)
    with _yc:
        st.markdown('<div class="dlg-yes-btn">', unsafe_allow_html=True)
        if st.button("Yes", use_container_width=True, type="secondary", key="dlg_logout_yes"):
            logout()
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with _nc:
        st.markdown('<div class="dlg-no-btn">', unsafe_allow_html=True)
        if st.button("No", use_container_width=True, type="primary", key="dlg_logout_no"):
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)


# ── Topbar ────────────────────────────────────────────────────────────────────
def render_topbar(page_name: str):
    user = current_user()
    _tl, _tr = st.columns([7, 2.5])
    with _tl:
        st.markdown(f"""
<div style="padding:16px 0 12px;">
    <span style="color:#2BBFA4;font-weight:600;font-size:14px;">RangeSheet</span>
    <span style="margin:0 8px;color:#C0B8B0;font-size:14px;">/</span>
    <span style="color:#1A1A1A;font-weight:600;font-size:14px;">{page_name}</span>
</div>""", unsafe_allow_html=True)
    with _tr:
        _chip_c, _icon_c = st.columns([4, 1])
        with _chip_c:
            st.markdown(f"""
<div style="padding-top:10px;display:flex;justify-content:flex-end;">
    <div style="display:inline-flex;align-items:center;gap:6px;background:#1A1A1A;
                padding:5px 14px 5px 10px;border-radius:20px;">
        <div style="width:6px;height:6px;background:#2BBFA4;border-radius:50%;flex-shrink:0;"></div>
        <span style="color:#fff;font-size:11px;font-weight:700;letter-spacing:.04em;">{user['employee_id']}</span>
    </div>
</div>""", unsafe_allow_html=True)
        with _icon_c:
            st.markdown('<div class="logout-btn-wrap" style="padding-top:8px;display:flex;justify-content:center;">',
                        unsafe_allow_html=True)
            if st.button("Sign Out", key="_tb_logout", help="Sign out"):
                _logout_dialog()
            st.markdown('</div>', unsafe_allow_html=True)
    st.markdown("<hr style='border-color:#D8D2C8;margin:-6px 0 24px;'>", unsafe_allow_html=True)


# ── File I/O helpers ──────────────────────────────────────────────────────────
_HEADER_KEYWORDS = {
    "department", "barcode", "item name", "section", "subclass", "tpna",
    "status", "edlp price by format", "star line", "supplier pack size",
    "no. of unit in inner", "tray total number", "hdet picking type",
    "check range to-be waterfall", "item priority",
}

def _detect_header_row(raw: bytes, encoding: str, sep: str = ",") -> int:
    try:
        preview = pd.read_csv(
            io.BytesIO(raw), encoding=encoding, sep=sep,
            header=None, nrows=60, on_bad_lines="skip",
            dtype=str, low_memory=False,
        )
        for i, row in preview.iterrows():
            cells = {str(c).strip().strip('"').strip("'").lower()
                     for c in row if pd.notna(c) and str(c).strip()}
            if len(cells & _HEADER_KEYWORDS) >= 2:
                return int(i)
    except Exception:
        pass
    return 0

def _detect_header_row_excel(raw: bytes, engine=None) -> int:
    try:
        kwargs = {"engine": engine} if engine else {}
        preview = pd.read_excel(io.BytesIO(raw), header=None, nrows=30, **kwargs)
        for i, row in preview.iterrows():
            cells = {str(c).strip().lower() for c in row if pd.notna(c)}
            if len(cells & _HEADER_KEYWORDS) >= 2:
                return int(i)
    except Exception:
        pass
    return 0

def _extract_rangesheet_meta(raw: bytes, encoding: str, header_row: int) -> dict:
    meta = {
        "dg_code": "—", "dg_name": "—", "minor_live_week": "—",
        "major_live_week": "—", "event_live_date": "—", "event_desc": "—",
        "range_arch": [],
    }
    try:
        lines = raw.decode(encoding, errors="replace").splitlines()
        for line in lines[1:header_row]:
            cells = [c.strip().strip('"') for c in line.split(",")]
            if not cells or not cells[0]:
                continue
            k = cells[0].upper()
            v = cells[1].strip() if len(cells) > 1 else "—"
            if "DG CODE" in k:                     meta["dg_code"] = v or "—"
            elif "DG NAME" in k:                   meta["dg_name"] = v or "—"
            elif "MINOR LIVE" in k:                meta["minor_live_week"] = v or "—"
            elif "MAJOR LIVE" in k:                meta["major_live_week"] = v or "—"
            elif "EVENT LIVE" in k or "LIVE DATE" in k: meta["event_live_date"] = v or "—"
            elif "EVENT DES" in k:                 meta["event_desc"] = v or "—"
            elif any(t in k for t in ["MAINTAIN","DELETE","NEW SOME","NEWNEW","TOTAL SKU"]):
                try:
                    ai = int(cells[1]) if len(cells) > 1 and str(cells[1]).strip().lstrip("-").isdigit() else 0
                    tb = int(cells[2]) if len(cells) > 2 and str(cells[2]).strip().lstrip("-").isdigit() else 0
                    meta["range_arch"].append({"type": cells[0].strip(), "as_is": ai,
                                               "to_be": tb, "diff": tb - ai})
                except Exception:
                    pass
    except Exception:
        pass
    return meta

# ── Column alias mapping ──────────────────────────────────────────────────────
COLUMN_MAPPING = {
    "dept":                                    "Department",
    "department":                              "Department",
    "Department Code&Desc":                  "Department",
    "department code&desc":                    "Department",
    "dg code":                                 "Department",
    "dg_code":                                 "Department",
    "section":                                 "Section",
    "section code&desc":                       "Section",
    "section code & desc":                     "Section",
    "dg name":                                 "Section",
    "dg_name":                                 "Section",
    "department name":                         "Section",
    "sub class":                               "Subclass",
    "sub-class":                               "Subclass",
    "subclass":                                "Subclass",
    "subclass code & desc":                    "Subclass",
    "subclass code&desc":                      "Subclass",
    "barcode":                                 "Barcode",
    "upc":                                     "Barcode",
    "ean":                                     "Barcode",
    "ean code":                                "Barcode",
    "sku":                                     "Barcode",
    "tpna":                                    "TPNA",
    "style number":                            "TPNA",
    "item id":                                 "ID",
    "item_id":                                 "ID",
    "itemid":                                  "ID",
    "no. of unit in case":                     "No. of Unit in Case",
    "no of unit in case":                      "No. of Unit in Case",
    "units per case":                          "No. of Unit in Case",
    "case units":                              "No. of Unit in Case",
    "no_of_unit_in_case":                      "No. of Unit in Case",
    "casetotalnumber":                         "No. of Unit in Case",
    "no. of unit in inner":                    "No. of Unit in Inner",
    "no of unit in inner":                     "No. of Unit in Inner",
    "no_of_unit_in_inner":                     "No. of Unit in Inner",
    "innerqty":                                "No. of Unit in Inner",
    "tray total number":                       "Tray total number",
    "tray_total_number":                       "Tray total number",
    "traytotalnumber":                         "Tray total number",
    "express picking type":                    "Express Picking Type",
    "express_picking_type":                    "Express Picking Type",
    "minipicktype":                            "Express Picking Type",
    "hdet picking type":                       "HDET Picking Type",
    "hdet_picking_type":                       "HDET Picking Type",
    "hyper&superpickingtype":                  "HDET Picking Type",
    "edlp price by format":                    "EDLP Price by Format",
    "edlp_price_by_format":                    "EDLP Price by Format",
    "edlp price":                              "EDLP Price by Format",
    "avg selling price by format":             "AVG Selling Price by Format",
    "avg_selling_price_by_format":             "AVG Selling Price by Format",
    "avg selling price":                       "AVG Selling Price by Format",
    "average selling price":                   "AVG Selling Price by Format",
    "item name":                               "Item Name",
    "item_name":                               "Item Name",
    "product name":                            "Item Name",
    "product_name":                            "Item Name",
    "productdescription":                      "Item Name",
    "description":                             "Item Name",
    "as is planograms applied":                "AS IS planograms applied",
    "asis planograms applied":                 "AS IS planograms applied",
    "as-is planograms applied":                "AS IS planograms applied",
    "to-be planograms applied":                "TO-BE planograms applied",
    "tobe planograms applied":                 "TO-BE planograms applied",
    "to be planograms applied":                "TO-BE planograms applied",
    "as-is stores applied":                    "AS-IS Stores Applied",
    "as is stores applied":                    "AS-IS Stores Applied",
    "asis stores applied":                     "AS-IS Stores Applied",
    "to-be stores applied":                    "TO-Be stores applied",
    "to be stores applied":                    "TO-Be stores applied",
    "tobe stores applied":                     "TO-Be stores applied",
    "avg units 52wk/ forecast new item sales": "Avg Units 52wk/ Forecast new item sales",
    "avg units 52wk/forecast new item sales":  "Avg Units 52wk/ Forecast new item sales",
    "avg unit 52wk":                           "Avg Units 52wk/ Forecast new item sales",
    "avg units 52wk":                          "Avg Units 52wk/ Forecast new item sales",
    "supplier pack size":                      "Supplier Pack Size",
    "supplier_pack_size":                      "Supplier Pack Size",
    "pack size":                               "Supplier Pack Size",
    "range tail yyyy":                         "Range Tail YYYY",
    "range_tail_yyyy":                         "Range Tail YYYY",
    "range tail":                              "Range Tail YYYY",
    "star line":                               "Star Line",
    "star_line":                               "Star Line",
    "starline":                                "Star Line",
    "item priority":                           "Item Priority",
    "item_priority":                           "Item Priority",
    "itempriority":                            "Item Priority",
    "jda vs actual":                           "JDA vs Actual",
    "jda_vs_actual":                           "JDA vs Actual",
    "actual-actual":                           "Actual-Actual",
    "actual_actual":                           "Actual-Actual",
    "status":                                  "Status",
    "check range to-be waterfall":             "Check Range To-be Waterfall",
    "check range to be waterfall":             "Check Range To-be Waterfall",
    "check range":                             "Check Range To-be Waterfall",
    "cluster (planogram name)":                "Cluster (Planogram name)",
    "planogram name":                          "Cluster (Planogram name)",
    "to be stores applied count":              "To be stores applied count",
    "mods":                                    "MODS",
    "fixtures":                                "Fixtures",
    "fixture":                                 "Fixtures",
    "range class":                             "Range Class",
    "total new skus":                          "Total New SKUs",
    "total delete skus":                       "Total Delete SKUs",
    "%achieving crd case (as is)":             "%Achieving CRD case (AS is)",
    "%achieving lrd (as is)":                  "%Achieving LRD (AS is)",
# }

# ── Column alias mapping ──────────────────────────────────────────────────────
# COLUMN_MAPPING = {
    # HDET → WebApp  (keys must be lowercase — apply_column_mapping does .lower() before lookup)
    "department code&desc":          "Department",
    "department code & desc":        "Department",
    "section code&desc":             "Section",
    "section code & desc":           "Section",
    "subclass code & desc":          "Subclass",
    "subclass code&desc":            "Subclass",
    "upc":                           "Barcode",
    "style number":                  "TPNA",
    "casetotalnumber":               "No. of Unit in Case",
    "innerqty":                      "No. of Unit in Inner",
    "traytotalnumber":               "Tray total number",
    "minipicktype":                  "Express Picking Type",
    "hyper&superpickingtype":        "HDET Picking Type",
    "edlp price":                    "EDLP Price by Format",
    "productdescription":            "Item Name",
    "th_tot_sales_volume_52_wk":     "Avg Units 52wk/ Forecast new item sales",
    "forecastsales":                 "Avg Units 52wk/ Forecast new item sales",
    "originalpacksize":              "Supplier Pack Size",
    "starline":                      "Star Line",
    "name":                          "Planogram Name",
}

def _norm_key(s: str) -> str:
    return _re.sub(r'[^a-z0-9]', '', str(s).strip().lower())

# สร้าง lookup แบบ normalized ไว้ล่วงหน้า
_COLUMN_MAPPING_NORM = {_norm_key(k): v for k, v in COLUMN_MAPPING.items()}

def apply_column_mapping(df: pd.DataFrame) -> pd.DataFrame:
    """Rename df columns using COLUMN_MAPPING. Skips rename if the target name
    already exists in the DataFrame (prevents creating duplicate column names)."""
    rename = {}
    taken = set(df.columns)
    for col in df.columns:
        key = str(col).strip().lower()
        target = _COLUMN_MAPPING_NORM.get(_norm_key(col))
        if target and col != target and target not in taken:
            rename[col] = target
            taken.add(target)
    return df.rename(columns=rename) if rename else df

def _detect_delimiter(raw: bytes, encoding: str = "utf-8-sig") -> str:
    """Detect column delimiter in a text file. Returns '|', '\\t', ';', or ','."""
    try:
        text = raw.decode(encoding, errors="replace")
        lines = [ln for ln in text.splitlines()[:20] if ln.strip() and not ln.startswith("#")]
        if not lines:
            return ","
        counts = {d: 0 for d in ("|", "\t", ";", ",")}
        for ln in lines[:8]:
            for d in counts:
                counts[d] += ln.count(d)
        best = max(counts, key=counts.get)
        return best if counts[best] >= 2 else ","
    except Exception:
        return ","

def _dedup_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Rename duplicate column names by appending .1, .2, … so Arrow/Streamlit
    never sees two identical column headers."""
    seen: dict = {}
    new_cols = []
    for c in df.columns:
        s = str(c)
        if s in seen:
            seen[s] += 1
            new_cols.append(f"{s}.{seen[s]}")
        else:
            seen[s] = 0
            new_cols.append(s)
    df = df.copy()
    df.columns = new_cols
    return df

def _clean_df(df: pd.DataFrame) -> pd.DataFrame:
    """Strip column names, drop fully-empty rows, apply column mapping, dedup."""
    df.columns = [str(c).strip().replace('\n', ' ').replace('\r', '') for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)
    df = apply_column_mapping(df)
    df = _dedup_columns(df)
    return df

def read_uploaded_file(uploaded_file):
    name = uploaded_file.name
    ext  = os.path.splitext(name)[-1].lower()
    if "_file_read_errors" not in st.session_state:
        st.session_state._file_read_errors = {}
    st.session_state._file_read_errors.pop(name, None)

    def _store_err(e):
        st.session_state._file_read_errors[name] = f"{type(e).__name__}: {e}"

    try:
        if ext in (".csv", ".txt"):
            raw = uploaded_file.read()
            uploaded_file.seek(0)
            sep = _detect_delimiter(raw)
            last_err = None
            for enc in ["utf-8-sig", "utf-8", "cp874", "latin1"]:
                try:
                    header_row = _detect_header_row(raw, enc, sep=sep)
                    df = pd.read_csv(io.BytesIO(raw), encoding=enc, sep=sep,
                                     skiprows=header_row, header=0,
                                     on_bad_lines="skip")
                    df = _clean_df(df)
                    if header_row > 0 and "rangesheet_meta" in st.session_state:
                        try:
                            st.session_state.rangesheet_meta = _extract_rangesheet_meta(
                                raw, enc, header_row)
                        except Exception:
                            pass
                    return df
                except Exception as e:
                    last_err = e
            if last_err:
                _store_err(last_err)

        elif ext in (".xlsx", ".xls", ".xlsb"):
            raw = uploaded_file.read()
            last_err = None

            if ext == ".xlsb":
                try:
                    header_row = _detect_header_row_excel(raw, engine="pyxlsb")
                    df = pd.read_excel(io.BytesIO(raw), header=header_row, engine="pyxlsb")
                    return _clean_df(df)
                except Exception as e:
                    _store_err(e); return None

            try:
                header_row = _detect_header_row_excel(raw)
                df = pd.read_excel(io.BytesIO(raw), header=header_row)
                df = _clean_df(df)
                if not df.empty:
                    return df
            except Exception as e:
                last_err = e

            try:
                df = pd.read_excel(io.BytesIO(raw), header=0)
                df = _clean_df(df)
                if not df.empty:
                    return df
            except Exception as e:
                last_err = e

            try:
                all_sheets = pd.read_excel(io.BytesIO(raw), sheet_name=None, dtype=str)
                for _sdf in all_sheets.values():
                    _sdf = _clean_df(_sdf)
                    if not _sdf.empty and len(_sdf.columns) > 1:
                        return _sdf
            except Exception as e:
                last_err = e

            try:
                df = pd.read_excel(io.BytesIO(raw), header=0, engine="xlrd")
                df = _clean_df(df)
                if not df.empty:
                    return df
            except Exception as e:
                last_err = e

            try:
                df = pd.read_excel(io.BytesIO(raw), header=0, dtype=str)
                df = _clean_df(df)
                if not df.empty:
                    return df
            except Exception as e:
                last_err = e

            if last_err:
                _store_err(last_err)

        else:
            _store_err(Exception(f"Unsupported file type: {ext}"))

    except Exception as e:
        _store_err(e)

    return None


# ═════════════════════════════════════════════════════════════════════════════
# ── NEW: Large-file support (e.g. HDET, ~6.7M rows / 3GB) ───────────────────
# ═════════════════════════════════════════════════════════════════════════════
# Added so files far bigger than normal range-sheet uploads (HDET-style
# master files) never go through the full pd.read_csv(io.BytesIO(raw), ...)
# path above, and never get pulled into auto_merge()'s pd.concat(). Instead:
#   - View Data shows a cheap preview (row count via line-counting + first
#     N rows), without ever building a DataFrame out of the full file.
#   - RangeSheet Review filters by a single DG/DG_CODE value using a
#     chunked scan, producing only the matching subset (~20k rows expected),
#     cached to Parquet so re-selecting the same DG is near-instant.
# Everything below reuses _detect_delimiter / _detect_header_row /
# apply_column_mapping / _dedup_columns / _clean_df from above so large-file
# results have exactly the same column names/shape as normal-file results.

LARGE_FILE_THRESHOLD_MB = 100
LARGE_FILE_CHUNK_SIZE = 250_000

# Candidate column names for the DG selector — checked in order, first
# match found in the actual file wins. Extend this list if a new file
# uses a different spelling.
DG_COLUMN_CANDIDATES = [
    "Display Group", "Display group", "display group",
    "DG_CODE", "DG_Code", "DG Code", "DG",
    "Department Code&Desc", "Department Code & Desc",
]

DG_NAME_CANDIDATES = [
    "DG_NAME", "DG Name", "dg name", "DG_name",
    "Display group desc", "Display Group Desc", "display group desc",
    "displaygroupdesc",
    "Department Name", "department name",
    "Section", "section",
]

_LARGE_FILE_CACHE_DIR = os.path.join(BASE_DIR, "cache_large")


def _ensure_large_cache_dir():
    try:
        os.makedirs(_LARGE_FILE_CACHE_DIR, exist_ok=True)
    except Exception:
        pass


def _hdet_parquet_path(csv_path: str) -> str:
    file_tag = os.path.splitext(os.path.basename(csv_path))[0]
    return os.path.join(_LARGE_FILE_CACHE_DIR, f"{file_tag}.parquet")


def ensure_hdet_parquet(csv_path: str, status_cb=None) -> str | None:
    """Convert a large CSV/TSV to Parquet once; return parquet path (or None on failure).
    Skips conversion if an up-to-date parquet already exists.
    status_cb(msg): optional callable for progress messages (e.g. st.status.write).
    """
    try:
        import pyarrow as _pa
        import pyarrow.parquet as _pq
    except ImportError:
        return None

    _ensure_large_cache_dir()
    pq_path   = _hdet_parquet_path(csv_path)
    csv_mtime = os.path.getmtime(csv_path)

    if os.path.exists(pq_path) and os.path.getmtime(pq_path) >= csv_mtime:
        return pq_path  # already up-to-date

    sep, encoding, header_row = _detect_large_file_params(csv_path)
    tmp_path = pq_path + ".tmp"
    writer   = None
    try:
        if status_cb:
            status_cb("Converting HDET to Parquet format (one-time, saves to disk)…")
        chunk_n = 0
        for chunk in pd.read_csv(
            csv_path, sep=sep, encoding=encoding,
            skiprows=header_row, header=0,
            chunksize=LARGE_FILE_CHUNK_SIZE,
            dtype=str, low_memory=False, on_bad_lines="skip",
        ):
            dg_col = _find_dg_col(chunk.columns)
            if dg_col:
                try:
                    chunk = chunk.sort_values(dg_col, kind="mergesort")
                except Exception:
                    pass
            table = _pa.Table.from_pandas(chunk.fillna(""), preserve_index=False)
            if writer is None:
                writer = _pq.ParquetWriter(tmp_path, table.schema, compression="snappy")
            writer.write_table(table, row_group_size=8192)
            chunk_n += 1
            if status_cb:
                status_cb(f"  …processed {chunk_n * LARGE_FILE_CHUNK_SIZE:,} rows")
        if writer:
            writer.close()
        os.replace(tmp_path, pq_path)
        return pq_path
    except Exception as _e:
        if writer:
            try: writer.close()
            except Exception: pass
        try: os.remove(tmp_path)
        except Exception: pass
        return None


def is_large_file(path: str) -> bool:
    """True if the file at `path` is at/above LARGE_FILE_THRESHOLD_MB."""
    try:
        size_mb = os.path.getsize(path) / (1024 * 1024)
        return size_mb >= LARGE_FILE_THRESHOLD_MB
    except Exception:
        return False


def _detect_large_file_params(path: str):
    """Sniff delimiter + header row using only a sample of bytes, reusing
    the existing _detect_delimiter / _detect_header_row helpers so behavior
    matches normal-sized files exactly."""
    with open(path, "rb") as f:
        raw_sample = f.read(2_000_000)  # ~2MB sample is plenty for sniffing
    sep = _detect_delimiter(raw_sample)
    encoding = "utf-8-sig"
    try:
        raw_sample.decode(encoding)
    except Exception:
        encoding = "cp874"
        try:
            raw_sample.decode(encoding)
        except Exception:
            encoding = "latin1"
    header_row = _detect_header_row(raw_sample, encoding, sep=sep)
    return sep, encoding, header_row


def get_large_file_preview(path: str, n_rows: int = 200) -> dict:
    """
    Lightweight preview for the View Data page — never loads the full file.
    Returns total row count (via line counting, not parsing), columns,
    and the first n_rows as a small DataFrame for st.dataframe().
    """
    sep, encoding, header_row = _detect_large_file_params(path)

    try:
        with open(path, "rb") as f:
            total_lines = sum(1 for _ in f)
        total_rows = max(total_lines - header_row - 1, 0)
    except Exception:
        total_rows = None

    preview_df = pd.read_csv(
        path, sep=sep, encoding=encoding,
        skiprows=header_row, header=0,
        nrows=n_rows, on_bad_lines="skip",
        dtype=str, low_memory=False,
    )
    preview_df = _clean_df(preview_df)

    return {
        "total_rows": total_rows,
        "columns": list(preview_df.columns),
        "preview_df": preview_df,
        "file_size_mb": round(os.path.getsize(path) / (1024 * 1024), 1),
    }


def read_large_file_head(path: str, n_rows: int = 500) -> pd.DataFrame:
    """Read the first n_rows from a large file without scanning the full file.
    Much faster than get_large_file_preview() for auto-loading because it
    skips the slow full-file line count."""
    sep, encoding, header_row = _detect_large_file_params(path)
    df = pd.read_csv(
        path, sep=sep, encoding=encoding,
        skiprows=header_row, header=0,
        nrows=n_rows, on_bad_lines="skip",
        dtype=str, low_memory=False,
    )
    return _clean_df(df)


def _find_dg_col(columns) -> str | None:
    """Return the DG column name from a list of columns, case-insensitive."""
    _cands_norm = {_re.sub(r'[^a-z0-9]', '', c.strip().lower()) for c in DG_COLUMN_CANDIDATES}
    for col in columns:
        if _re.sub(r'[^a-z0-9]', '', col.strip().lower()) in _cands_norm:
            return col
    # Fallback: column whose stripped name is just "dg" or contains "display" + "group"
    for col in columns:
        n = _re.sub(r'[^a-z0-9]', '', col.strip().lower())
        if n == "dg" or n == "dgcode" or "displaygroup" in n:
            return col
    return None


@st.cache_data(show_spinner=False)
def get_dg_options(path: str, _mtime: float = 0.0) -> tuple:
    """
    Scan a large file in chunks and return (dg_column_name, sorted unique values).
    Cached by file path + mtime — re-scans only when the file changes.
    Pass _mtime=os.path.getmtime(path) to enable cache invalidation.
    """
    sep, encoding, header_row = _detect_large_file_params(path)
    dg_col = None
    seen_values = set()

    for chunk in pd.read_csv(
        path, sep=sep, encoding=encoding,
        skiprows=header_row, header=0,
        chunksize=LARGE_FILE_CHUNK_SIZE,
        dtype=str, low_memory=False, on_bad_lines="skip",
    ):
        if dg_col is None:
            dg_col = _find_dg_col(chunk.columns)
            if dg_col is None:
                return None, []
        seen_values.update(chunk[dg_col].dropna().astype(str).str.strip().unique())

    return dg_col, sorted(seen_values)


def get_dg_index(path: str) -> dict:
    """
    Return {dg_col, dg_name_col, codes, names, code_to_name} for a large file.
    Uses Parquet mirror (via ensure_hdet_parquet) when available for a fast
    column-only scan. Falls back to chunked CSV. Result JSON is saved to disk
    and reused across app restarts — re-scans only when the source file changes.
    """
    import json as _json
    _ensure_large_cache_dir()
    file_tag  = os.path.splitext(os.path.basename(path))[0]
    idx_file  = os.path.join(_LARGE_FILE_CACHE_DIR, f"{file_tag}__dg_index.json")
    try:
        file_mtime = os.path.getmtime(path)
    except OSError:
        return {}

    # Return disk-cached index if mtime matches (survives restarts)
    if os.path.exists(idx_file):
        try:
            with open(idx_file, "r", encoding="utf-8") as _f:
                _cached = _json.load(_f)
            if abs(_cached.get("mtime", 0) - file_mtime) < 1:
                return _cached
        except Exception:
            pass

    _cands_name_norm = {_re.sub(r"[^a-z0-9]", "", c.lower()) for c in DG_NAME_CANDIDATES}
    dg_col       = None
    dg_name_col  = None
    code_set: set      = set()
    code_to_name: dict = {}

    # Fast path: read only needed columns from Parquet
    pq_path = _hdet_parquet_path(path)
    _used_parquet = False
    if os.path.exists(pq_path) and os.path.getmtime(pq_path) >= file_mtime:
        try:
            import pyarrow.parquet as _pq
            _schema = _pq.read_schema(pq_path)
            _cols   = [f.name for f in _schema]
            dg_col  = _find_dg_col(_cols)
            if dg_col:
                dg_name_col = next(
                    (c for c in _cols if _re.sub(r"[^a-z0-9]", "", c.lower()) in _cands_name_norm), None
                )
                _read_cols = [dg_col] + ([dg_name_col] if dg_name_col else [])
                _tbl = _pq.read_table(pq_path, columns=_read_cols)
                _codes_arr = _tbl[dg_col].to_pylist()
                code_set = {str(c).strip() for c in _codes_arr if c and str(c).strip() not in ("nan", "None", "")}
                if dg_name_col:
                    _names_arr = _tbl[dg_name_col].to_pylist()
                    for c, n in zip(_codes_arr, _names_arr):
                        cs, ns = str(c).strip(), str(n).strip()
                        if cs and cs not in ("nan","None","") and ns not in ("nan","None",""):
                            code_to_name.setdefault(cs, ns)
                _used_parquet = True
        except Exception:
            pass

    # Slow path: chunked CSV scan
    if not _used_parquet:
        sep, encoding, header_row = _detect_large_file_params(path)
        for chunk in pd.read_csv(
            path, sep=sep, encoding=encoding,
            skiprows=header_row, header=0,
            chunksize=LARGE_FILE_CHUNK_SIZE,
            dtype=str, low_memory=False, on_bad_lines="skip",
        ):
            if dg_col is None:
                dg_col = _find_dg_col(chunk.columns)
                if dg_col is None:
                    return {}
                dg_name_col = next(
                    (c for c in chunk.columns
                     if _re.sub(r"[^a-z0-9]", "", c.lower()) in _cands_name_norm), None
                )
            codes = chunk[dg_col].astype(str).str.strip()
            code_set.update(c for c in codes.unique() if c and c not in ("nan", "None"))
            if dg_name_col:
                names = chunk[dg_name_col].astype(str).str.strip()
                for c, n in zip(codes, names):
                    if c and c not in ("nan", "None") and n not in ("nan", "None", ""):
                        code_to_name.setdefault(c, n)

    result = {
        "mtime":        file_mtime,
        "dg_col":       dg_col,
        "dg_name_col":  dg_name_col,
        "codes":        sorted(code_set),
        "names":        sorted(set(code_to_name.values())),
        "code_to_name": code_to_name,
    }
    try:
        with open(idx_file, "w", encoding="utf-8") as _f:
            _json.dump(result, _f, ensure_ascii=False)
    except Exception:
        pass
    return result


def scan_hdet_dg_cascade(path: str, col_candidates: dict) -> dict:
    """Scan full HDET in chunks and build a cascade map keyed by DG value.

    col_candidates keys: "dg", "fmt", "div", "cls"  — each maps to a list of
    candidate column names tried in order.

    Returns:
        {
          "dg_vals":  [...],   # all unique DG values
          "fmt_vals": [...],   # all unique fmt values
          "div_vals": [...],   # all unique div values
          "cls_vals": [...],   # all unique cls values
          "cascade":  {dg_val: {"fmt": [...], "div": [...], "cls": [...]}},
        }
    """
    sep, encoding, header_row = _detect_large_file_params(path)

    # Resolve actual column names from first chunk
    first_chunk = next(iter(pd.read_csv(
        path, sep=sep, encoding=encoding,
        skiprows=header_row, header=0,
        chunksize=1000, dtype=str, low_memory=False, on_bad_lines="skip",
    )))
    actual = {}
    low_map = {str(c).strip().lower(): c for c in first_chunk.columns}
    for key, candidates in col_candidates.items():
        for cand in candidates:
            if cand in first_chunk.columns:
                actual[key] = cand
                break
            if cand.strip().lower() in low_map:
                actual[key] = low_map[cand.strip().lower()]
                break

    dg_key = actual.get("dg")
    if not dg_key:
        return {"dg_vals": [], "fmt_vals": [], "div_vals": [], "cls_vals": [], "cascade": {}}

    use_cols = list({v for v in actual.values() if v})
    other_keys = [k for k in ("fmt", "div", "cls") if k in actual]

    # {dg_val: {other_key: set of values}}
    cascade: dict = {}
    all_vals: dict = {k: set() for k in other_keys}
    all_dg: set = set()

    for chunk in pd.read_csv(
        path, sep=sep, encoding=encoding,
        skiprows=header_row, header=0,
        chunksize=LARGE_FILE_CHUNK_SIZE,
        usecols=use_cols,
        dtype=str, low_memory=False, on_bad_lines="skip",
    ):
        chunk = chunk.fillna("").apply(lambda c: c.str.strip())
        dg_series = chunk[dg_key]
        mask = dg_series != ""
        chunk = chunk[mask]
        dg_series = chunk[dg_key]

        all_dg.update(dg_series.unique())

        for k in other_keys:
            col = actual[k]
            all_vals[k].update(chunk[col][chunk[col] != ""].unique())

            # group by DG, collect unique values of this column
            for dg_val, grp in chunk.groupby(dg_key, sort=False):
                if dg_val not in cascade:
                    cascade[dg_val] = {kk: set() for kk in other_keys}
                cascade[dg_val][k].update(grp[col][grp[col] != ""].unique())

    # Finalise dg_cascade with sorted lists
    dg_cascade = {dg: {k: sorted(v) for k, v in sets.items()}
                  for dg, sets in cascade.items()}

    # Derive reverse: cls_val → {dg: [...], fmt: [...], div: [...]}
    cls_cascade: dict = {}
    for dg_val, related in dg_cascade.items():
        for cls_val in related.get("cls", []):
            if cls_val not in cls_cascade:
                cls_cascade[cls_val] = {"dg": set(), "fmt": set(), "div": set()}
            cls_cascade[cls_val]["dg"].add(dg_val)
            cls_cascade[cls_val]["fmt"].update(related.get("fmt", []))
            cls_cascade[cls_val]["div"].update(related.get("div", []))
    cls_cascade = {cls: {k: sorted(v) for k, v in sets.items()}
                   for cls, sets in cls_cascade.items()}

    return {
        "dg_vals":  sorted(all_dg),
        "fmt_vals": sorted(all_vals.get("fmt", [])),
        "div_vals": sorted(all_vals.get("div", [])),
        "cls_vals": sorted(all_vals.get("cls", [])),
        "cascade":     dg_cascade,   # dg  → {fmt, div, cls}
        "cls_cascade": cls_cascade,  # cls → {dg,  fmt, div}
    }


def summarize_hdet_by_cluster(
    path: str,
    col_candidates: dict,
    filter_cols: dict = None,
) -> pd.DataFrame:
    """Scan HDET in chunks, return cluster-level summary DataFrame.

    col_candidates: {
        "cls":   [...],   # ClusterName column candidates
        "store": [...],   # store identifier column candidates (for StoreCount)
        "pog":   [...],   # POG name column candidates (for Count of POGName)
    }
    filter_cols: same format as count_hdet_totals (optional).

    Returns DataFrame: ClusterName | StoreCount | Count of POGName
    with a Total row appended. Totals are COUNT(DISTINCT) across ALL rows,
    not sum of per-cluster counts (stores/POGs can appear in multiple clusters).
    Sorted by StoreCount descending.
    """
    sep, encoding, header_row = _detect_large_file_params(path)

    first = next(iter(pd.read_csv(
        path, sep=sep, encoding=encoding,
        skiprows=header_row, header=0,
        chunksize=1000, dtype=str, low_memory=False, on_bad_lines="skip",
    )))
    low_map = {str(c).strip().lower(): c for c in first.columns}

    def _resolve(candidates):
        for c in candidates:
            if c in first.columns:
                return c
            if str(c).strip().lower() in low_map:
                return low_map[str(c).strip().lower()]
        return None

    cls_col   = _resolve(col_candidates.get("cls",   []))
    store_col = _resolve(col_candidates.get("store", []))
    pog_col   = _resolve(col_candidates.get("pog",   []))
    id_col    = _resolve(col_candidates.get("id",    []))

    if not cls_col:
        return pd.DataFrame({"ClusterName": [], "StoreCount": [],
                             "Count of POGName": [], "ItemCount": []})

    actual_filters = {}
    if filter_cols:
        for _key, (cands, vals) in filter_cols.items():
            col = _resolve(cands)
            if col and vals:
                actual_filters[col] = set(str(v) for v in vals)

    use_cols = list({cls_col}
                    | ({store_col} if store_col else set())
                    | ({pog_col}   if pog_col   else set())
                    | ({id_col}    if id_col    else set())
                    | set(actual_filters.keys()))

    agg: dict = {}

    for chunk in pd.read_csv(
        path, sep=sep, encoding=encoding,
        skiprows=header_row, header=0,
        chunksize=LARGE_FILE_CHUNK_SIZE,
        usecols=use_cols,
        dtype=str, low_memory=False, on_bad_lines="skip",
    ):
        mask = pd.Series(True, index=chunk.index)
        for col, vals in actual_filters.items():
            if col in chunk.columns:
                mask &= chunk[col].str.strip().isin(vals)
        chunk = chunk[mask]
        if chunk.empty:
            continue

        chunk[cls_col] = chunk[cls_col].fillna("").str.strip()

        for cls_val, grp in chunk.groupby(cls_col, sort=False):
            if not cls_val:
                continue
            if cls_val not in agg:
                agg[cls_val] = {"store": set(), "pog": set(), "pog_pairs": set(), "id": set()}
            if store_col and store_col in grp.columns:
                _sv = grp[store_col].fillna("").str.strip()
                agg[cls_val]["store"].update(_sv.unique())
            if pog_col and pog_col in grp.columns:
                _pv = grp[pog_col].fillna("").str.strip()
                agg[cls_val]["pog"].update(_pv.unique())
                if store_col and store_col in grp.columns:
                    agg[cls_val]["pog_pairs"].update(zip(_sv, _pv))
            if id_col and id_col in grp.columns:
                agg[cls_val]["id"].update(grp[id_col].dropna().str.strip().unique())

    if not agg:
        return pd.DataFrame({"ClusterName": [], "StoreCount": [],
                             "Count of POGName": [], "ItemCount": []})

    _null = {"", "nan", "NaN", "none", "None", "NULL", "null", "N/A", "n/a"}

    def _pog_count(v):
        # Use (store, pog) pairs when available — matches Power BI COUNT(POGName)
        if v["pog_pairs"]:
            return sum(1 for s, p in v["pog_pairs"] if s not in _null and p not in _null)
        return len(v["pog"] - _null)

    rows = [
        {
            "ClusterName":      cls_val,
            "StoreCount":       len(v["store"] - _null),
            "Count of POGName": _pog_count(v),
            "ItemCount":        len(v["id"]     - _null),
        }
        for cls_val, v in agg.items()
    ]
    df = (pd.DataFrame(rows)
            .sort_values("StoreCount", ascending=False)
            .reset_index(drop=True))

    # Total = union across all clusters
    _all_pairs = set().union(*(v["pog_pairs"] for v in agg.values()))
    _all_store = set().union(*(v["store"]     for v in agg.values()))
    _all_id    = set().union(*(v["id"]        for v in agg.values()))
    _tot_pog   = (sum(1 for s, p in _all_pairs if s not in _null and p not in _null)
                  if _all_pairs
                  else len(set().union(*(v["pog"] for v in agg.values())) - _null))
    total_row = pd.DataFrame([{
        "ClusterName":      "Total",
        "StoreCount":       len(_all_store - _null),
        "Count of POGName": _tot_pog,
        "ItemCount":        len(_all_id    - _null),
    }])
    return pd.concat([df, total_row], ignore_index=True)


def scan_csv_store_counts(
    path: str,
    filter_cols: dict = None,
) -> tuple:
    """Scan small FP/POG CSV (store-level, <100 MB) for StoreCount per cluster.

    Returns (cluster_dict, global_count):
      cluster_dict  – {cluster_name: COUNT(DISTINCT store_no)}
      global_count  – total COUNT(DISTINCT store_no) across all clusters after filter
    """
    sep, encoding, header_row = _detect_large_file_params(path)
    first = next(iter(pd.read_csv(
        path, sep=sep, encoding=encoding,
        skiprows=header_row, header=0,
        chunksize=1000, dtype=str, low_memory=False, on_bad_lines="skip",
    )))
    low_map = {str(c).strip().lower(): c for c in first.columns}

    def _res(candidates):
        for c in candidates:
            if c in first.columns:
                return c
            if str(c).strip().lower() in low_map:
                return low_map[str(c).strip().lower()]
        return None

    cls_col   = _res(["POG_Cluster", "ClusterName", "Cluster_Name",
                       "cluster_name", "Cluster Name", "Store_Cluster"])
    store_col = _res(["store_no", "StoreNo", "store_id", "Store_No",
                       "store_number", "PG_Store_Number"])

    if not cls_col or not store_col:
        return {}, 0

    actual_filters: dict = {}
    if filter_cols:
        for _key, (cands, vals) in filter_cols.items():
            col = _res(cands)
            if col and vals:
                actual_filters[col] = set(str(v) for v in vals)

    use_cols = list({cls_col, store_col} | set(actual_filters.keys()))

    agg: dict = {}
    global_stores: set = set()

    for chunk in pd.read_csv(
        path, sep=sep, encoding=encoding,
        skiprows=header_row, header=0,
        chunksize=LARGE_FILE_CHUNK_SIZE,
        usecols=use_cols, dtype=str,
        low_memory=False, on_bad_lines="skip",
    ):
        mask = pd.Series(True, index=chunk.index)
        for col, vals in actual_filters.items():
            if col in chunk.columns:
                mask &= chunk[col].str.strip().isin(vals)
        chunk = chunk[mask]
        if chunk.empty:
            continue

        chunk[cls_col] = chunk[cls_col].fillna("").str.strip()
        global_stores.update(chunk[store_col].dropna().str.strip().unique())

        for cls_val, grp in chunk.groupby(cls_col, sort=False):
            if not cls_val:
                continue
            if cls_val not in agg:
                agg[cls_val] = set()
            agg[cls_val].update(grp[store_col].dropna().str.strip().unique())

    return (
        {k: len(v - {""}) for k, v in agg.items()},
        len(global_stores - {""}),
    )


def count_hdet_totals(
    path: str,
    count_cols: dict,
    filter_cols: dict,
) -> dict:
    """Chunked HDET scan: apply filters, return COUNT(DISTINCT) for named columns.

    count_cols  : {key: [candidate_col_names, ...]}
                  e.g. {"id": ["ID","Barcode"], "name": ["FP_Name","FP Name"]}
    filter_cols : {key: ([candidate_col_names], [filter_values])}
                  e.g. {"dg": (["DG","DG_CODE","Display Group"], ["K3K"])}

    Returns: {key: int}  — 0 when column not found in file.
    """
    sep, encoding, header_row = _detect_large_file_params(path)

    first = next(iter(pd.read_csv(
        path, sep=sep, encoding=encoding,
        skiprows=header_row, header=0,
        chunksize=1000, dtype=str, low_memory=False, on_bad_lines="skip",
    )))
    low_map = {str(c).strip().lower(): c for c in first.columns}

    def _resolve(candidates):
        for c in candidates:
            if c in first.columns:
                return c
            if str(c).strip().lower() in low_map:
                return low_map[str(c).strip().lower()]
        return None

    # Resolve count columns
    actual_count = {k: _resolve(v) for k, v in count_cols.items()}
    actual_count = {k: v for k, v in actual_count.items() if v}

    # Resolve filter columns → {actual_col: set_of_values}
    actual_filters = {}
    for _key, (cands, vals) in filter_cols.items():
        col = _resolve(cands)
        if col and vals:
            actual_filters[col] = set(str(v) for v in vals)

    if not actual_count:
        return {k: 0 for k in count_cols}

    use_cols = list(set(actual_count.values()) | set(actual_filters.keys()))
    seen = {k: set() for k in actual_count}

    for chunk in pd.read_csv(
        path, sep=sep, encoding=encoding,
        skiprows=header_row, header=0,
        chunksize=LARGE_FILE_CHUNK_SIZE,
        usecols=use_cols,
        dtype=str, low_memory=False, on_bad_lines="skip",
    ):
        mask = pd.Series(True, index=chunk.index)
        for col, vals in actual_filters.items():
            if col in chunk.columns:
                mask &= chunk[col].str.strip().isin(vals)
        chunk = chunk[mask]
        for k, col in actual_count.items():
            if col in chunk.columns:
                seen[k].update(chunk[col].dropna().str.strip().unique())

    return {k: len(v - {""}) for k, v in seen.items()}


def scan_hdet_unique_vals(path: str, col_candidates: dict) -> dict:
    """Scan the full large file in chunks and return unique values for multiple columns.

    col_candidates: {key: [list of candidate column names in priority order]}
    Returns: {key: sorted list of unique string values}

    Reads only the matched columns (usecols) so memory stays low even on large files.
    """
    sep, encoding, header_row = _detect_large_file_params(path)

    # First chunk: discover which actual column names match each key
    first_chunk = next(iter(pd.read_csv(
        path, sep=sep, encoding=encoding,
        skiprows=header_row, header=0,
        chunksize=1000, dtype=str, low_memory=False, on_bad_lines="skip",
    )))
    actual_cols = {}  # key → actual column name found
    for key, candidates in col_candidates.items():
        for cand in candidates:
            if cand in first_chunk.columns:
                actual_cols[key] = cand
                break
            # case-insensitive fallback
            low_map = {str(c).strip().lower(): c for c in first_chunk.columns}
            if cand.strip().lower() in low_map:
                actual_cols[key] = low_map[cand.strip().lower()]
                break

    if not actual_cols:
        return {k: [] for k in col_candidates}

    use_cols = list(set(actual_cols.values()))
    seen = {key: set() for key in actual_cols}

    for chunk in pd.read_csv(
        path, sep=sep, encoding=encoding,
        skiprows=header_row, header=0,
        chunksize=LARGE_FILE_CHUNK_SIZE,
        usecols=use_cols,
        dtype=str, low_memory=False, on_bad_lines="skip",
    ):
        for key, col in actual_cols.items():
            if col in chunk.columns:
                seen[key].update(chunk[col].dropna().astype(str).str.strip().unique())

    return {key: sorted(v - {""}) for key, v in seen.items()}


def load_large_file_by_dg(path: str, dg_value: str, use_cache: bool = True) -> pd.DataFrame:
    """
    Filter a large file down to rows matching a single DG value.
    Fast path: filter from Parquet mirror using pyarrow (seconds).
    Slow path: chunked CSV scan (fallback).
    Per-DG result is also cached as Parquet for instant re-loads.
    """
    _ensure_large_cache_dir()
    dg_value       = str(dg_value).strip()
    dg_value_upper = dg_value.upper()
    file_tag       = os.path.splitext(os.path.basename(path))[0]
    cache_file     = os.path.join(_LARGE_FILE_CACHE_DIR, f"{file_tag}__dg_{dg_value_upper}.parquet")
    file_mtime     = os.path.getmtime(path)

    # Return per-DG cached parquet if still fresh.
    # NOTE: this file was already run through _clean_df() before being cached
    # below — do not clean it again here. apply_column_mapping() isn't
    # idempotent (e.g. a second pass renames the already-correct
    # "Planogram Name" column to "Cluster (Planogram name)"), so re-cleaning
    # on every cache hit silently corrupted that column after the first load.
    if use_cache and os.path.exists(cache_file) and os.path.getmtime(cache_file) >= file_mtime:
        return pd.read_parquet(cache_file)

    result = pd.DataFrame()

    # Fast path: filter full Parquet with pyarrow (no full CSV re-read)
    pq_path = _hdet_parquet_path(path)
    if os.path.exists(pq_path) and os.path.getmtime(pq_path) >= file_mtime:
        try:
            import pyarrow.parquet as _pq
            import pyarrow.compute as _pc
            _schema  = _pq.read_schema(pq_path)
            _cols    = [f.name for f in _schema]
            dg_col   = _find_dg_col(_cols)
            if dg_col:
                # Try Parquet predicate pushdown first. This can avoid reading most
                # row groups for a new DG and is much faster than loading the full
                # HDET table into memory before filtering.
                _candidate_vals = list(dict.fromkeys([dg_value, dg_value_upper]))
                for _cand in _candidate_vals:
                    try:
                        _tbl = _pq.read_table(
                            pq_path,
                            filters=[(dg_col, "=", _cand)],
                            use_threads=True,
                            memory_map=True,
                        )
                        if _tbl.num_rows:
                            result = _tbl.to_pandas()
                            break
                    except Exception:
                        pass
                if result.empty:
                    _tbl   = _pq.read_table(pq_path, use_threads=True, memory_map=True)
                    _upper = _pc.utf8_upper(_pc.utf8_strip(_tbl[dg_col].cast("string")))
                    _mask  = _pc.equal(_upper, dg_value_upper)
                    result = _tbl.filter(_mask).to_pandas()
        except Exception:
            result = pd.DataFrame()

    # Slow path: chunked CSV scan
    if result.empty:
        sep, encoding, header_row = _detect_large_file_params(path)
        matched_chunks = []
        dg_col = None
        for chunk in pd.read_csv(
            path, sep=sep, encoding=encoding,
            skiprows=header_row, header=0,
            chunksize=LARGE_FILE_CHUNK_SIZE,
            dtype=str, low_memory=False, on_bad_lines="skip",
        ):
            if dg_col is None:
                dg_col = _find_dg_col(chunk.columns)
                if dg_col is None:
                    return pd.DataFrame()
            filtered = chunk[chunk[dg_col].astype(str).str.strip().str.upper() == dg_value_upper]
            if not filtered.empty:
                matched_chunks.append(filtered)
        result = pd.concat(matched_chunks, ignore_index=True) if matched_chunks else pd.DataFrame()

    result = _clean_df(result)
    if use_cache and not result.empty:
        try:
            result.to_parquet(cache_file)
        except Exception:
            pass
    return result


def auto_merge(files_info: list) -> tuple:
    dfs = [(f.get("name", ""), f["df"]) for f in files_info if f.get("df") is not None]
    if not dfs:
        return None, []
    if len(dfs) == 1:
        return _dedup_columns(dfs[0][1].copy()), [f"📄 Single file — {dfs[0][0]}"]
    merged = pd.concat([d for _, d in dfs], ignore_index=True, sort=False)
    merged = _dedup_columns(merged)
    before = len(merged)
    try:
        merged = merged.drop_duplicates(ignore_index=True)
    except Exception:
        pass
    log = [f"✅ Merged {len(dfs)} files → {len(merged):,} rows"]
    if len(merged) < before:
        log.append(f"🧹 Removed {before - len(merged):,} duplicates")
    return merged, log

def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import PatternFill, Font, Alignment
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")
        buf.seek(0)
        wb = load_workbook(buf)
        ws = wb.active
        green_fill  = PatternFill(fill_type="solid", fgColor="2BBFA4")
        white_font  = Font(color="FFFFFF", bold=True)
        mid_align   = Alignment(horizontal="center", vertical="center")
        for cell in ws[1]:
            cell.fill      = green_fill
            cell.font      = white_font
            cell.alignment = mid_align
        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()
    except Exception:
        return df.to_csv(index=False).encode("utf-8-sig")

def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")

def find_col(df, keywords):
    if df is None:
        return None
    for c in df.columns:
        if any(k in str(c).lower() for k in keywords):
            return c
    return None

def get_fill(col: str) -> str:
    c = str(col).strip()
    if "TO-BE" in c or "TO BE" in c:
        return "mer"
    if "AS-IS" in c or "AS IS" in c or "%MOR" in c:
        return "formula"
    _display = {
        "Department", "Section", "Subclass", "Barcode", "TPNA", "ID", "Item Name",
        "No. of unit in case", "No. of unit in inner", "Tray total number",
        "Express Picking type", "HDET picking type", "EDLP Price by Format",
        "Star Line", "Status", "Check Range To-be Waterfall",
    }
    if c in _display:
        return "display"
    _mer = {"Mer Price (incl. vat7%)", "COST", "Supplier Pack Size",
            "Avg Units 52wk/ Forecast new item sales"}
    if c in _mer:
        return "mer"
    return "formula"

# ── Storage helpers ───────────────────────────────────────────────────────────
def _ensure_dirs():
    for sub in ("uploads", "merged", "audit"):
        try:
            os.makedirs(os.path.join(BASE_DIR, sub), exist_ok=True)
        except Exception:
            pass

def save_merged_snapshot(df: pd.DataFrame):
    try:
        _ensure_dirs()
        df.to_csv(os.path.join(BASE_DIR, "merged", "merged_latest.csv"),
                  index=False, encoding="utf-8-sig")
    except Exception:
        pass

def clear_merged_snapshot():
    try:
        path = os.path.join(BASE_DIR, "merged", "merged_latest.csv")
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        pass

def load_merged_snapshot():
    try:
        path = os.path.join(BASE_DIR, "merged", "merged_latest.csv")
        if os.path.exists(path):
            return pd.read_csv(path, encoding="utf-8-sig")
    except Exception:
        pass
    return None

def storage_info() -> dict:
    try:
        _ensure_dirs()
        upload_dir = os.path.join(BASE_DIR, "uploads")
        files = [f for f in os.listdir(upload_dir)
                 if os.path.isfile(os.path.join(upload_dir, f))]
        total = sum(os.path.getsize(os.path.join(upload_dir, f)) for f in files)
        return {"files": len(files), "total_mb": round(total / 1024 / 1024, 1)}
    except Exception:
        return {"files": 0, "total_mb": 0.0}

def save_file(file_bytes: bytes, filename: str) -> str:
    try:
        _ensure_dirs()
        path = os.path.join(BASE_DIR, "uploads", filename)
        with open(path, "wb") as f:
            f.write(file_bytes)
        return path
    except Exception:
        return ""

# ── Admin-pinned file manifest ────────────────────────────────────────────────
_ADMIN_MANIFEST = os.path.join(BASE_DIR, "admin_files.json")

def load_admin_manifest() -> list:
    """Return list of pinned-file metadata dicts from disk."""
    try:
        if os.path.exists(_ADMIN_MANIFEST):
            with open(_ADMIN_MANIFEST, "r", encoding="utf-8") as _f:
                return json.load(_f)
    except Exception:
        pass
    return []

def save_admin_manifest(entries: list):
    """Persist pinned-file metadata list to disk."""
    try:
        _ensure_dirs()
        with open(_ADMIN_MANIFEST, "w", encoding="utf-8") as _f:
            json.dump(entries, _f, indent=2, ensure_ascii=False)
    except Exception:
        pass

def _read_file_from_path(path: str):
    """Read a saved file from disk using the same logic as read_uploaded_file."""
    class _F:
        name = os.path.basename(path)
        _buf = None
        def read(self):
            if self._buf is None:
                with open(path, "rb") as fh:
                    self._buf = fh.read()
            return self._buf
        def seek(self, n): pass
    return read_uploaded_file(_F())

@st.cache_resource(show_spinner=False)
def _parsed_file_cache():
    """Process-wide cache: filename → {mtime, entry}.
    Files are only parsed once per server process (or when mtime changes)."""
    return {}

def _invalidate_file_cache(name: str):
    """Force a specific file to be re-parsed next time."""
    _parsed_file_cache().pop(name, None)

def load_admin_file_df(name: str):
    """Load one admin file's DataFrame on demand, using the process-wide cache.
    Returns None if the file doesn't exist or can't be parsed.

    NOTE: large files (e.g. HDET) should be checked with is_large_file()
    by the caller and routed to get_large_file_preview() /
    load_large_file_by_dg() instead of this function. This function itself
    is unchanged from before, so behavior for normal-sized files is identical.
    """
    cache = _parsed_file_cache()
    _path = os.path.join(BASE_DIR, "uploads", name)
    if not os.path.exists(_path):
        return None
    _mtime = os.path.getmtime(_path)
    hit = cache.get(name)
    if hit and hit["mtime"] == _mtime:
        return hit["entry"].get("df")
    df = _read_file_from_path(_path)
    if df is not None:
        cache[name] = {"mtime": _mtime, "entry": {"name": name, "df": df}}
    return df

def get_admin_file_entries() -> list:
    """Return entry dicts for all admin-pinned files.
    Parsed DataFrames are cached by file mtime — no disk read on repeat calls.

    CHANGED: large files (e.g. HDET) are now skipped here on purpose — they
    never enter auto_merge()'s pd.concat(). They get a metadata-only entry
    (df=None, is_large=True) instead. auto_merge() already ignores entries
    where df is None, so this alone keeps HDET out of merged_df without
    touching auto_merge() itself. The View Data page should check
    entry.get("is_large") and call get_large_file_preview() for those.
    """
    cache   = _parsed_file_cache()
    entries = []
    for meta in load_admin_manifest():
        _name  = meta["name"]
        _path  = os.path.join(BASE_DIR, "uploads", _name)
        if not os.path.exists(_path):
            cache.pop(_name, None)
            continue

        if is_large_file(_path):
            entries.append({
                **meta,
                "df": None,
                "rows": None,
                "cols": None,
                "pinned": True,
                "is_large": True,
            })
            continue

        _mtime = os.path.getmtime(_path)
        hit    = cache.get(_name)
        if hit and hit["mtime"] == _mtime and "name" in hit["entry"]:
            entries.append(hit["entry"])
            continue
        try:
            df = _read_file_from_path(_path)
            if df is not None:
                _entry = {
                    **meta,
                    "df":     df,
                    "rows":   len(df),
                    "cols":   len(df.columns),
                    "pinned": True,
                    "is_large": False,
                }
                cache[_name] = {"mtime": _mtime, "entry": _entry}
                entries.append(_entry)
        except Exception:
            pass
    return entries

def remove_admin_file(name: str):
    """Remove a file from the admin manifest (does not delete the file bytes)."""
    manifest = [m for m in load_admin_manifest() if m["name"] != name]
    save_admin_manifest(manifest)

# ── Shared cross-session database cache ───────────────────────────────────────
# Bump this integer whenever file-parsing logic changes so every running process
# discards its cached data and re-reads from disk automatically.
_CACHE_VERSION = 5  # bumped 4 -> 5: get_admin_file_entries() now routes large files away

def _manifest_sig() -> str:
    """Short fingerprint of the admin manifest — changes when files are added/removed."""
    try:
        m = load_admin_manifest()
        if not m:
            return ""
        raw = "|".join(sorted(f"{x.get('name','')}:{x.get('date','')}" for x in m))
        for x in m:
            _p = os.path.join(BASE_DIR, "uploads", x.get("name", ""))
            if os.path.exists(_p):
                raw += f"|{os.path.getmtime(_p):.0f}"
        import hashlib
        return hashlib.md5(raw.encode()).hexdigest()[:12]
    except Exception:
        return ""

@st.cache_resource(show_spinner=False)
def _shared_db_cache():
    """Single mutable dict shared across ALL user sessions in this process."""
    return {"df": None, "files_info": [], "version": 0, "cache_v": -1, "manifest_sig": ""}

def bump_shared_db():
    """Invalidate the shared database — call after admin pins or unpins a file."""
    _s = _shared_db_cache()
    _s["df"] = None
    _s["files_info"] = []
    _s["version"] += 1

def get_shared_db():
    """Return (merged_df, files_info) from the shared admin database.
    Auto-reloads when: parser code version changes, manifest changes, or after a bump.
    Returns (None, []) when no pinned files exist.

    Large files (is_large=True entries from get_admin_file_entries) carry
    df=None, and auto_merge() already skips entries where df is None — so
    they're automatically excluded from merged_df without any change here.
    """
    _s = _shared_db_cache()

    if _s.get("cache_v") != _CACHE_VERSION:
        _s["df"] = None
        _s["files_info"] = []
        _s["cache_v"] = _CACHE_VERSION

    _sig = _manifest_sig()
    if _sig and _s.get("manifest_sig") != _sig:
        _s["df"] = None
        _s["files_info"] = []
        _s["manifest_sig"] = _sig

    if _s["df"] is None:
        _entries = get_admin_file_entries()
        if _entries:
            _merged, _finfo = auto_merge(_entries)
            _s["df"] = _merged if (_merged is not None and len(_merged) > 0) else pd.DataFrame()
            _s["files_info"] = _finfo
        else:
            _s["df"] = pd.DataFrame()
            _s["files_info"] = []

    _df = _s["df"]
    return (_df if len(_df) > 0 else None), _s["files_info"]

def load_audit_log():
    try:
        path = os.path.join(BASE_DIR, "audit", "audit_log.csv")
        if os.path.exists(path):
            return pd.read_csv(path, encoding="utf-8-sig")
    except Exception:
        pass
    return None
