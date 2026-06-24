"""Shared utilities, CSS, auth, data processing — RangeSheet."""

import streamlit as st
import pandas as pd
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
    "Range Sheet_Non-SSPOG", "Range Sheet_SSPOG",
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
    ("rawfiles",  "📁", "View Data"),
    ("viewdata",  "🔍", "Rangesheet Review"),
    ("dashboard", "📊", "Dashboard"),
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
        "AS-IS Stores Applied",
        "TO-Be stores applied",
        "Avg Units 52wk/ Forecast new item sales",
        "Supplier Pack Size",
        "Range Tail YYYY",
        "AVG Selling Price by Format",
    ]},
    {"group": "Star Line", "color": "#000000", "hdr_color": "#FFFFFF", "cols": [
        "Star Line",
    ]},
    {"group": "Priority", "color": "#00CC44", "hdr_color": "#003300", "cols": [
        "Item Priority", "JDA vs Actual", "Actual-Actual",
    ]},
    {"group": "Status", "color": "#F5F5F5", "hdr_color": "#555555", "cols": [
        "Status", "Check Range To-be Waterfall",
        "(name of the planogram+productname+store)",
    ]},
    {"group": "Cluster Summary", "color": "#C9A0DC", "hdr_color": "#3D0070", "cols": [
        "Cluster (Planogram name)", "To be stores applied count", "AS IS",
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
    "edlp price by format":                      "EDLP Price by Formate",
    "item name":                                 "Item name",
    "as is planograms applied":                  "As IS planograms applied",
    "to-be planograms applied":                  "To-BE planograms applied",
    "as-is stores applied":                      "AS-IS Store applied",
    "to-be stores applied":                      "To-Be store applied",
    "avg units 52wk/ forecast new item sales":   "Avg unit 52 wk/forecast new item sales",
    "supplier pack size":                        "Supplier pack size",
    "range tail yyyy":                           "Range Tail YYYY",
    "avg selling price by format":               "AVG selling Price by format",
    "star line":                                 "Star Line",
    "item priority":                             "Item priority",
    "jda vs actual":                             "JDA vs Actual",
    "actual-actual":                             "Actual-Actual",
    "status":                                    "Status",
    "check range to-be waterfall":               "Check Range to be waterfall",
    "(name of the planogram+productname+store)": "(name of the planogram+productname+store)",
    "cluster (planogram name)":                  "Cluster (Planogram name)",
    "to be stores applied count":                "To be stores applied count",
    "as is":                                     "As IS",
    "mods":                                      "MODS",
    "fixtures":                                  "Fixtures",
    "range class":                               "Range Class",
    "total new skus":                            "Total New SKUS",
    "total delete skus":                         "Total Delete SKUs",
    "%achieving crd case (as is)":               "%Achieving CRD case (AS is)",
    "%achieving lrd (as is)":                    "%Achieving LRD (AS is)",
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

    _ICON_GREY = (
        "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA1"
        "MTIgNTEyIj48cGF0aCBmaWxsPSIjNTU1NTU1IiBkPSJNNTAyLjYgMjc4LjZjMTIuNS0xMi41"
        "IDEyLjUtMzIuOCAwLTQ1LjNsLTEyOC0xMjhjLTEyLjUtMTIuNS0zMi44LTEyLjUtNDUuMyAw"
        "cy0xMi41IDMyLjggMCA0NS4zTDQwMi43IDIyNCAxOTIgMjI0Yy0xNy43IDAtMzIgMTQuMy0z"
        "MiAzMnMxNC4zIDMyIDMyIDMybDIxMC43IDAtNzMuNCA3My40Yy0xMi41IDEyLjUtMTIuNSAz"
        "Mi44IDAgNDUuM3MzMi44IDEyLjUgNDUuMyAwbDEyOC0xMjh6TTE2MCA5NmMxNy43IDAgMzIt"
        "MTQuMyAzMi0zMnMtMTQuMy0zMi0zMi0zMkw5NiAzMkM0MyAzMiAwIDc1IDAgMTI4TDAgMzg0"
        "YzAgNTMgNDMgOTYgOTYgOTZsNjQgMGMxNy43IDAgMzItMTQuMyAzMi0zMnMtMTQuMy0zMi0z"
        "Mi0zMmwtNjQgMGMtMTcuNyAwLTMyLTE0LjMtMzItMzJsMC0yNTZjMC0xNy43IDE0LjMtMzIg"
        "MzItMzJsNjQgMHoiLz48L3N2Zz4="
    )
    _ICON_RED = (
        "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA1"
        "MTIgNTEyIj48cGF0aCBmaWxsPSIjRTUzOTM1IiBkPSJNNTAyLjYgMjc4LjZjMTIuNS0xMi41"
        "IDEyLjUtMzIuOCAwLTQ1LjNsLTEyOC0xMjhjLTEyLjUtMTIuNS0zMi44LTEyLjUtNDUuMyAw"
        "cy0xMi41IDMyLjggMCA0NS4zTDQwMi43IDIyNCAxOTIgMjI0Yy0xNy43IDAtMzIgMTQuMy0z"
        "MiAzMnMxNC4zIDMyIDMyIDMybDIxMC43IDAtNzMuNCA3My40Yy0xMi41IDEyLjUtMTIuNSAz"
        "Mi44IDAgNDUuM3MzMi44IDEyLjUgNDUuMyAwbDEyOC0xMjh6TTE2MCA5NmMxNy43IDAgMzIt"
        "MTQuMyAzMi0zMnMtMTQuMy0zMi0zMi0zMkw5NiAzMkM0MyAzMiAwIDc1IDAgMTI4TDAgMzg0"
        "YzAgNTMgNDMgOTYgOTYgOTZsNjQgMGMxNy43IDAgMzItMTQuMyAzMi0zMnMtMTQuMy0zMi0z"
        "Mi0zMmwtNjQgMGMtMTcuNyAwLTMyLTE0LjMtMzItMzJsMC0yNTZjMC0xNy43IDE0LjMtMzIg"
        "MzItMzJsNjQgMHoiLz48L3N2Zz4="
    )
    st.markdown(f"""
<style>
button[title="Sign out"],
div[title="Sign out"] button,
span[title="Sign out"] button,
button[title="Sign out"]:hover,
div[title="Sign out"] button:hover,
span[title="Sign out"] button:hover,
.stMarkdown:has(.logout-btn-wrap) ~ div button:hover {{
    background-color: #FFF0F0 !important;
    background-image: url("data:image/svg+xml;base64,{_ICON_RED}") !important;
    border-color: #E53935 !important;
    color: #E53935 !important;
}}
</style>
""", unsafe_allow_html=True)


# ── Page navigation (Next / Previous) ────────────────────────────────────────
_PAGE_ORDER  = ["landpage", "rawfiles", "viewdata", "dashboard", "report", "audit"]
_PAGE_LABELS = {
    "landpage":  "My Files",
    "rawfiles":  "View Data",
    "viewdata":  "Rangesheet Review",
    "dashboard": "Dashboard",
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

        m = st.session_state.get("merged_df")
        if m is not None:
            st.markdown(
                f'<p style="padding:2px 16px;color:rgba(255,255,255,0.35);'
                f'font-size:11px;">Merged: {len(m):,} rows · {len(m.columns)} cols</p>',
                unsafe_allow_html=True)
        sinfo = storage_info()
        st.markdown(
            f'<p style="padding:2px 16px;color:rgba(255,255,255,0.35);'
            f'font-size:11px;">📁 {sinfo["files"]} files · {sinfo["total_mb"]} MB</p>',
            unsafe_allow_html=True)


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
    "department code & desc":                  "Department",
    "department code&desc":                    "Department",
    "dg code":                                 "Department",
    "dg_code":                                 "Department",
    "section":                                 "Section",
    "dg name":                                 "Section",
    "dg_name":                                 "Section",
    "department name":                         "Section",
    "sub class":                               "Subclass",
    "sub-class":                               "Subclass",
    "subclass":                                "Subclass",
    "barcode":                                 "Barcode",
    "upc":                                     "Barcode",
    "ean":                                     "Barcode",
    "ean code":                                "Barcode",
    "sku":                                     "Barcode",
    "tpna":                                    "TPNA",
    "item id":                                 "ID",
    "item_id":                                 "ID",
    "itemid":                                  "ID",
    "no. of unit in case":                     "No. of Unit in Case",
    "no of unit in case":                      "No. of Unit in Case",
    "units per case":                          "No. of Unit in Case",
    "case units":                              "No. of Unit in Case",
    "no_of_unit_in_case":                      "No. of Unit in Case",
    "no. of unit in inner":                    "No. of Unit in Inner",
    "no of unit in inner":                     "No. of Unit in Inner",
    "no_of_unit_in_inner":                     "No. of Unit in Inner",
    "tray total number":                       "Tray total number",
    "tray_total_number":                       "Tray total number",
    "express picking type":                    "Express Picking Type",
    "express_picking_type":                    "Express Picking Type",
    "hdet picking type":                       "HDET Picking Type",
    "hdet_picking_type":                       "HDET Picking Type",
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
}

def apply_column_mapping(df: pd.DataFrame) -> pd.DataFrame:
    """Rename df columns using COLUMN_MAPPING. Skips rename if the target name
    already exists in the DataFrame (prevents creating duplicate column names)."""
    rename = {}
    taken = set(df.columns)
    for col in df.columns:
        key = str(col).strip().lower()
        target = COLUMN_MAPPING.get(key)
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
DG_COLUMN_CANDIDATES = ["DG_CODE", "DG_Code", "DG Code", "DG"]

_LARGE_FILE_CACHE_DIR = os.path.join(BASE_DIR, "cache_large")


def _ensure_large_cache_dir():
    try:
        os.makedirs(_LARGE_FILE_CACHE_DIR, exist_ok=True)
    except Exception:
        pass


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


def get_dg_options(path: str) -> tuple:
    """
    Scan a large file in chunks and return (dg_column_name, sorted unique values)
    for populating the DG/DG_CODE dropdown on RangeSheet Review.
    Does NOT load the full file into one DataFrame — chunked scan only.
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
            for cand in DG_COLUMN_CANDIDATES:
                if cand in chunk.columns:
                    dg_col = cand
                    break
            if dg_col is None:
                return None, []  # no DG-like column in this file
        seen_values.update(chunk[dg_col].dropna().astype(str).unique())

    return dg_col, sorted(seen_values)


def load_large_file_by_dg(path: str, dg_value: str, use_cache: bool = True) -> pd.DataFrame:
    """
    Filter a large file (e.g. HDET) down to rows matching a single
    DG/DG_CODE value, reading in chunks so memory stays bounded regardless
    of total file size. Result is cached as Parquet keyed by file + DG value,
    so re-selecting the same DG is near-instant on subsequent loads.
    """
    _ensure_large_cache_dir()
    dg_value = str(dg_value)
    file_tag = os.path.splitext(os.path.basename(path))[0]
    cache_file = os.path.join(_LARGE_FILE_CACHE_DIR, f"{file_tag}__dg_{dg_value}.parquet")

    file_mtime = os.path.getmtime(path)
    if use_cache and os.path.exists(cache_file):
        if os.path.getmtime(cache_file) >= file_mtime:
            return pd.read_parquet(cache_file)

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
            for cand in DG_COLUMN_CANDIDATES:
                if cand in chunk.columns:
                    dg_col = cand
                    break
            if dg_col is None:
                return pd.DataFrame()

        filtered = chunk[chunk[dg_col].astype(str) == dg_value]
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
    dfs = [(f["name"], f["df"]) for f in files_info if f.get("df") is not None]
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
        cache[name] = {"mtime": _mtime, "entry": {"df": df}}
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
        if hit and hit["mtime"] == _mtime:
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

    if (_s["df"] is not None and not _s["df"].empty
            and len(_s["df"].columns) <= 2 and len(_s["df"]) > 10):
        _s["df"] = None
        _s["files_info"] = []

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
