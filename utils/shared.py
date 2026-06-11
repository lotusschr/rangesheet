# """Shared utilities, CSS, auth, data processing — RangeSheet."""

# import streamlit as st
# import pandas as pd
# import os
# import io
# from datetime import datetime

# # ── Paths ───────────────────────────────────────────────────────────────────
# _HERE    = os.path.dirname(os.path.abspath(__file__))
# ROOT_DIR = os.path.dirname(_HERE)
# BASE_DIR = os.path.join(ROOT_DIR, "rangesheet_data")

# # ── Constants ────────────────────────────────────────────────────────────────
# APP_CONFIG = {"allowed_extensions": ["csv", "xlsx", "xls"], "max_file_mb": 1024}

# RS_SHEETS    = ["Range Sheet_Non-SSPOG", "Range Sheet_SSPOG",
#                 "StoreApply_SSPOG", "5.1 ItembyStore", "5.2 ItembyStore_SC"]
# RS_TYPE_ORDER = ["MAINTAIN", "DELETE SOME", "DELETE ALL", "NEW SOME", "NEWNEW"]

# STATUS_COLORS = {
#     "MAINTAIN":    {"bg": "#E8F8F5", "c": "#2BBFA4"},
#     "NEW SOME":    {"bg": "#E3F2FD", "c": "#1565C0"},
#     "NEW":         {"bg": "#E8F5E9", "c": "#2E7D32"},
#     "NEWNEW":      {"bg": "#F3E5F5", "c": "#6A1B9A"},
#     "DELETE SOME": {"bg": "#FEE8E8", "c": "#E05555"},
#     "DELETE ALL":  {"bg": "#FFEBEE", "c": "#B71C1C"},
# }

# NAV_ITEMS = [
#     ("landpage",  "🗂️", "My Files"),
#     ("viewdata",  "🔍", "View Data"),
#     ("dashboard", "📊", "Dashboard"),
#     ("report",    "📋", "Report"),
#     ("audit",     "🛡️", "Audit Log"),
# ]

# RS_COL_GROUPS = [
#     {"group": "Item Info", "color": "#F5F5F5", "cols": [
#         "Department", "Section", "Subclass", "Barcode", "TPNA", "ID", "Item Name"]},
#     {"group": "Pack Info", "color": "#FAFAFA", "cols": [
#         "No. of unit in case", "No. of unit in inner", "Tray total number",
#         "Express Picking type", "HDET picking type"]},
#     {"group": "Price", "color": "#FFF9C4", "cols": [
#         "EDLP Price by Format", "Mer Price (incl. vat7%)", "COST", "%MOR (from EDLP)"]},
#     {"group": "Range Architecture", "color": "#E8F8F5", "cols": [
#         "AS-IS planograms applied", "TO-BE planograms applied",
#         "AS-IS Stores Applied", "TO-BE stores applied"]},
#     {"group": "Sales & Forecast", "color": "#FFF9C4", "cols": [
#         "Avg Units 52wk/ Forecast new item sales", "Supplier Pack Size",
#         "AS-IS Sale Total Units", "TO-BE Total Sale Units", "Total Units change",
#         "AS-IS Total Sales (Ex Vat)", "TO-BE Total Sales (Ex Vat)", "Total Sales change (Ex Vat)"]},
#     {"group": "Margin", "color": "#FCE4EC", "cols": [
#         "AS-IS Total Margin (Ex Vat)", "TO-BE Total Margin (Ex Vat)", "Total Margin change (Ex Vat)"]},
#     {"group": "Status", "color": "#FCE4EC", "cols": ["Status", "Check Range To-be Waterfall"]},
# ]

# FILL_COLORS = {
#     "display": {"bg": "#FFFDE7", "text": "#5D4037"},
#     "mer":     {"bg": "#FCE4EC", "text": "#880E4F"},
#     "formula": {"bg": "#F5F5F5", "text": "#616161"},
# }

# # ── Auth ─────────────────────────────────────────────────────────────────────
# def current_user() -> dict:
#     return st.session_state.get("auth_user", {
#         "employee_id": "DEV001", "name": "Developer",
#         "role": "admin", "login_time": datetime.now().strftime("%H:%M"),
#     })

# def is_admin() -> bool:
#     return current_user().get("role") == "admin"

# def logout():
#     st.session_state.auth_user = None

# def add_audit(action: str, detail: str = ""):
#     u = current_user()
#     record = {
#         "id": u["employee_id"], "name": u["name"], "role": u["role"].upper(),
#         "session": datetime.now().strftime("%H:%M"),
#         "date": datetime.now().strftime("%d/%m/%Y"),
#         "action": action, "detail": detail,
#     }
#     if "audit_log" in st.session_state:
#         st.session_state.audit_log.append(record)

# # ── Session state ─────────────────────────────────────────────────────────────
# def init_session_state():
#     defaults = {
#         "raw_files": [], "merged_df": None, "merge_log": [], "audit_log": [],
#         "display_cols": None, "canvases": {}, "rangesheet_meta": {},
#         "view_sheet": RS_SHEETS[1], "view_vis_cols": None,
#         "last_uploaded_preview": None,
#     }
#     for k, v in defaults.items():
#         if k not in st.session_state:
#             st.session_state[k] = v
#     if st.session_state.get("merged_df") is None:
#         snap = load_merged_snapshot()
#         if snap is not None:
#             st.session_state.merged_df = snap
#             st.session_state.merge_log = [f"💾 Loaded snapshot ({len(snap):,} rows)"]

# # ── CSS ───────────────────────────────────────────────────────────────────────
# def inject_css():
#     st.markdown("""
# <style>
# @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Sarabun:wght@400;500;600;700&display=swap');

# html, body, [class*="css"] { font-family: 'Inter','Sarabun',system-ui,sans-serif !important; }
# .main { background: #F2EDE8 !important; }
# .block-container { padding: 0 !important; max-width: 100% !important; }
# [data-testid="stMainBlockContainer"] { padding: 1.5rem 2rem !important; }

# /* ── Sidebar NAV buttons ── */
# [data-testid="stSidebar"] .stButton > button {
#   background: transparent !important;
#   color: #444 !important;
#   border: none !important;
#   border-radius: 8px !important;
#   font-size: 14px !important;
#   width: 100% !important;
#   text-align: left !important;
#   justify-content: flex-start !important;
#   padding: 9px 14px !important;
#   margin: 2px 0 !important;
#   transition: background 0.15s;
# }
# [data-testid="stSidebar"] .stButton > button:hover {
#   background: #E8F8F5 !important;
#   color: #2BBFA4 !important;
# }

# [data-testid="stMainBlockContainer"] { padding-top: 0 !important; }

# /* ── Main content buttons ── */
# .stButton > button {
#   background: #2BBFA4 !important;
#   color: #fff !important;
#   border: none !important;
#   border-radius: 8px !important;
#   font-weight: 600 !important;
#   font-size: 13px !important;
# }
# .stButton > button:hover { background: #22A08A !important; }

# /* ── Topbar logout button — keep it small/secondary ── */
# button[data-testid="baseButton-secondary"] {
#   background: transparent !important;
#   color: #E05555 !important;
#   border: 1.5px solid #E05555 !important;
#   font-size: 12px !important;
# }
# button[data-testid="baseButton-secondary"]:hover {
#   background: #FEE8E8 !important;
#   color: #B71C1C !important;
#   border-color: #B71C1C !important;
# }

# /* ── Download buttons ── */
# .stDownloadButton > button {
#   background: transparent !important;
#   color: #2BBFA4 !important;
#   border: 1.5px solid #2BBFA4 !important;
#   border-radius: 8px !important;
#   font-weight: 600 !important; font-size: 13px !important;
# }
# .stDownloadButton > button:hover { background: #E8F8F5 !important; }

# /* ── Metric cards ── */
# [data-testid="stMetric"] {
#   background: #fff !important; border: 1px solid #E0D9D2 !important;
#   border-radius: 12px !important; padding: 14px 16px !important;
#   box-shadow: 0 1px 4px rgba(0,0,0,.06) !important;
# }
# [data-testid="stMetricValue"] { color: #1A1A1A !important; font-weight: 800 !important; }
# [data-testid="stMetricLabel"] { color: #888 !important; font-size: 11px !important;
#   font-weight: 700 !important; text-transform: uppercase !important; }

# /* ── Tabs ── */
# .stTabs [data-baseweb="tab-list"] { border-bottom: 2px solid #E0D9D2 !important; gap: 4px; }
# .stTabs [data-baseweb="tab"] { border-radius: 8px 8px 0 0 !important;
#   font-weight: 600 !important; color: #888 !important; }
# .stTabs [aria-selected="true"] { color: #2BBFA4 !important;
#   border-bottom: 2px solid #2BBFA4 !important; }

# /* ── Expander ── */
# [data-testid="stExpander"] { border: 1px solid #E0D9D2 !important;
#   border-radius: 10px !important; background: #fff !important; }

# /* ── File uploader ── */
# [data-testid="stFileUploader"] {
#   background: #F8F4F0 !important; border: 2px dashed #D6CFC8 !important;
#   border-radius: 12px !important;
# }

# /* ── Dataframe ── */
# [data-testid="stDataFrame"] { border-radius: 10px !important;
#   border: 1px solid #E0D9D2 !important; }

# hr { border-color: #E0D9D2 !important; }

# /* ── Hide Streamlit chrome, keep sidebar toggle ── */
# #MainMenu, footer { visibility: hidden; }
# header { visibility: hidden; }
# header button { visibility: visible !important; }
# [data-testid="stSidebarCollapseButton"] { visibility: visible !important; }
# [data-testid="collapsedControl"] { visibility: visible !important; display: flex !important; }
# .stDeployButton { display: none; }

# /* ── Hide default st.navigation sidebar list ── */
# [data-testid="stSidebarNav"] { display: none !important; }
# </style>
# """, unsafe_allow_html=True)

# # ── Shared UI components ───────────────────────────────────────────────────
# def render_sidebar(active: str = "landpage"):
#     user = current_user()
#     role_label = "ADMIN" if is_admin() else "VIEWER"

#     with st.sidebar:
#         st.markdown("### 🗂️ RangeSheet")
#         st.caption("Management Platform")
#         st.divider()

#         for page_id, icon, label in NAV_ITEMS:
#             if page_id == active:
#                 st.markdown(
#                     f'<div style="background:#E8F8F5;border-left:3px solid #2BBFA4;'
#                     f'border-radius:0 8px 8px 0;padding:8px 12px 8px 11px;margin:2px 0;">'
#                     f'<span style="color:#2BBFA4;font-weight:600;font-size:14px;">'
#                     f'{icon} {label}</span></div>',
#                     unsafe_allow_html=True)
#             else:
#                 if st.button(f"{icon} {label}", key=f"nav_{page_id}",
#                              use_container_width=True):
#                     st.switch_page(f"pages/{page_id}.py")

#         st.divider()

#         m = st.session_state.get("merged_df")
#         if m is not None:
#             st.caption(f"**Merged data:** {len(m):,} rows · {len(m.columns)} cols")

#         sinfo = storage_info()
#         st.caption(f"📁 {sinfo['files']} files · {sinfo['total_mb']} MB")

#         st.divider()
#         st.caption(f"👤 **{user['name']}** · {user['employee_id']}")
#         st.caption(f"🕐 {user['login_time']} · `{role_label}`")
#         if st.button("🚪 Logout", key="sidebar_logout", use_container_width=True, type="secondary"):
#             add_audit("LOGOUT")
#             logout()
#             st.rerun()


# def render_topbar(title: str):
#     """Renders page title heading. User info and logout are in the sidebar."""
#     st.markdown(
#         f"<div style='padding:10px 0 4px;font-size:20px;font-weight:700;color:#1A1A1A;'>"
#         f"<span style='color:#2BBFA4;font-size:14px;font-weight:600;'>RangeSheet</span>"
#         f"<span style='color:#BBB;'> · </span>{title}</div>",
#         unsafe_allow_html=True,
#     )
#     st.markdown("<hr style='margin:0 0 14px;border-color:#E0D9D2;'>", unsafe_allow_html=True)


# # ── File I/O helpers ──────────────────────────────────────────────────────────
# _HEADER_KEYWORDS = {"department", "barcode", "item name", "section", "subclass", "tpna"}

# def _detect_header_row(raw: bytes, encoding: str) -> int:
#     try:
#         lines = raw.decode(encoding, errors="replace").splitlines()
#         for i, line in enumerate(lines[:30]):
#             cells = {c.strip().lower() for c in line.split(",")}
#             if len(cells & _HEADER_KEYWORDS) >= 2:
#                 return i
#     except Exception:
#         pass
#     return 0

# def _extract_rangesheet_meta(raw: bytes, encoding: str, header_row: int) -> dict:
#     meta = {"dg_code": "—", "dg_name": "—", "minor_live_week": "—",
#             "major_live_week": "—", "event_live_date": "—", "event_desc": "—",
#             "range_arch": []}
#     try:
#         lines = raw.decode(encoding, errors="replace").splitlines()
#         for line in lines[1:header_row]:
#             cells = [c.strip().strip('"') for c in line.split(",")]
#             if not cells or not cells[0]:
#                 continue
#             k = cells[0].upper()
#             v = cells[1].strip() if len(cells) > 1 else "—"
#             if "DG CODE" in k:         meta["dg_code"] = v or "—"
#             elif "DG NAME" in k:       meta["dg_name"] = v or "—"
#             elif "MINOR LIVE" in k:    meta["minor_live_week"] = v or "—"
#             elif "MAJOR LIVE" in k:    meta["major_live_week"] = v or "—"
#             elif "EVENT LIVE" in k or "LIVE DATE" in k: meta["event_live_date"] = v or "—"
#             elif "EVENT DES" in k:     meta["event_desc"] = v or "—"
#             elif any(t in k for t in ["MAINTAIN","DELETE","NEW SOME","NEWNEW","TOTAL SKU"]):
#                 try:
#                     ai = int(cells[1]) if len(cells) > 1 and str(cells[1]).strip().lstrip("-").isdigit() else 0
#                     tb = int(cells[2]) if len(cells) > 2 and str(cells[2]).strip().lstrip("-").isdigit() else 0
#                     meta["range_arch"].append({"type": cells[0].strip(), "as_is": ai,
#                                                "to_be": tb, "diff": tb - ai})
#                 except Exception:
#                     pass
#     except Exception:
#         pass
#     return meta

# def read_uploaded_file(uploaded_file):
#     name = uploaded_file.name
#     ext  = os.path.splitext(name)[-1].lower()
#     try:
#         if ext == ".csv":
#             raw = uploaded_file.read()
#             uploaded_file.seek(0)
#             for enc in ["utf-8-sig", "utf-8", "cp874", "latin1"]:
#                 try:
#                     header_row = _detect_header_row(raw, enc)
#                     df = pd.read_csv(io.BytesIO(raw), encoding=enc,
#                                      skiprows=header_row, header=0)
#                     df.columns = [str(c).strip().replace('\n',' ').replace('\r','') for c in df.columns]
#                     if "Department" in df.columns:
#                         df = df[df["Department"].notna() & (df["Department"].astype(str).str.strip() != "")]
#                     df = df.reset_index(drop=True)
#                     if header_row > 0 and "rangesheet_meta" in st.session_state:
#                         try:
#                             st.session_state.rangesheet_meta = _extract_rangesheet_meta(raw, enc, header_row)
#                         except Exception:
#                             pass
#                     return df
#                 except Exception:
#                     continue
#         elif ext in (".xlsx", ".xls"):
#             df = pd.read_excel(uploaded_file)
#             df.columns = [str(c).strip().replace('\n',' ').replace('\r','') for c in df.columns]
#             return df
#     except Exception:
#         pass
#     return None

# def auto_merge(files_info: list) -> tuple:
#     dfs = [(f["name"], f["df"]) for f in files_info if f.get("df") is not None]
#     if not dfs:
#         return None, []
#     if len(dfs) == 1:
#         return dfs[0][1].copy(), [f"📄 Single file — {dfs[0][0]}"]
#     merged = pd.concat([d for _, d in dfs], ignore_index=True, sort=False)
#     before = len(merged)
#     merged = merged.drop_duplicates(ignore_index=True)
#     log = [f"✅ Merged {len(dfs)} files → {len(merged):,} rows"]
#     if len(merged) < before:
#         log.append(f"🧹 Removed {before-len(merged):,} duplicates")
#     return merged, log

# def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
#     try:
#         buf = io.BytesIO()
#         with pd.ExcelWriter(buf, engine="openpyxl") as writer:
#             df.to_excel(writer, index=False)
#         return buf.getvalue()
#     except Exception:
#         return df.to_csv(index=False).encode("utf-8-sig")

# def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
#     return df.to_csv(index=False).encode("utf-8-sig")

# def find_col(df, keywords):
#     if df is None:
#         return None
#     for c in df.columns:
#         if any(k in str(c).lower() for k in keywords):
#             return c
#     return None

# def get_fill(col: str) -> str:
#     c = str(col).strip()
#     if "TO-BE" in c or "TO BE" in c: return "mer"
#     if "AS-IS" in c or "AS IS" in c or "%MOR" in c: return "formula"
#     _display = {"Department","Section","Subclass","Barcode","TPNA","ID","Item Name",
#                 "No. of unit in case","No. of unit in inner","Tray total number",
#                 "Express Picking type","HDET picking type","EDLP Price by Format",
#                 "Star Line","Status","Check Range To-be Waterfall"}
#     if c in _display: return "display"
#     _mer = {"Mer Price (incl. vat7%)","COST","Supplier Pack Size",
#             "Avg Units 52wk/ Forecast new item sales"}
#     if c in _mer: return "mer"
#     return "formula"

# # ── Storage helpers ───────────────────────────────────────────────────────────
# def _ensure_dirs():
#     for sub in ("uploads", "merged", "audit"):
#         try:
#             os.makedirs(os.path.join(BASE_DIR, sub), exist_ok=True)
#         except Exception:
#             pass

# def save_merged_snapshot(df: pd.DataFrame):
#     try:
#         _ensure_dirs()
#         df.to_csv(os.path.join(BASE_DIR, "merged", "merged_latest.csv"),
#                   index=False, encoding="utf-8-sig")
#     except Exception:
#         pass

# def load_merged_snapshot():
#     try:
#         path = os.path.join(BASE_DIR, "merged", "merged_latest.csv")
#         if os.path.exists(path):
#             return pd.read_csv(path, encoding="utf-8-sig")
#     except Exception:
#         pass
#     return None

# def storage_info() -> dict:
#     try:
#         _ensure_dirs()
#         upload_dir = os.path.join(BASE_DIR, "uploads")
#         files = [f for f in os.listdir(upload_dir)
#                  if os.path.isfile(os.path.join(upload_dir, f))]
#         total = sum(os.path.getsize(os.path.join(upload_dir, f)) for f in files)
#         return {"files": len(files), "total_mb": round(total/1024/1024, 1)}
#     except Exception:
#         return {"files": 0, "total_mb": 0.0}

# def save_file(file_bytes: bytes, filename: str) -> str:
#     try:
#         _ensure_dirs()
#         path = os.path.join(BASE_DIR, "uploads", filename)
#         with open(path, "wb") as f:
#             f.write(file_bytes)
#         return path
#     except Exception:
#         return ""

# def load_audit_log():
#     try:
#         path = os.path.join(BASE_DIR, "audit", "audit_log.csv")
#         if os.path.exists(path):
#             return pd.read_csv(path, encoding="utf-8-sig")
#     except Exception:
#         pass
#     return None

"""Shared utilities, CSS, auth, data processing — RangeSheet."""

import streamlit as st
import pandas as pd
import os
import io
import base64
from datetime import datetime

# ── Paths ───────────────────────────────────────────────────────────────────
_HERE    = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(_HERE)
BASE_DIR = os.path.join(ROOT_DIR, "rangesheet_data")

# ── Constants ────────────────────────────────────────────────────────────────
APP_CONFIG = {"allowed_extensions": ["csv", "xlsx", "xls"], "max_file_mb": 1024}

RS_SHEETS    = ["Range Sheet_Non-SSPOG", "Range Sheet_SSPOG",
                "StoreApply_SSPOG", "5.1 ItembyStore", "5.2 ItembyStore_SC"]
RS_TYPE_ORDER = ["MAINTAIN", "DELETE SOME", "DELETE ALL", "NEW SOME", "NEWNEW"]

STATUS_COLORS = {
    "MAINTAIN":    {"bg": "#E8F8F5", "c": "#2BBFA4"},
    "NEW SOME":    {"bg": "#E3F2FD", "c": "#1565C0"},
    "NEW":         {"bg": "#E8F5E9", "c": "#2E7D32"},
    "NEWNEW":      {"bg": "#F3E5F5", "c": "#6A1B9A"},
    "DELETE SOME": {"bg": "#FEE8E8", "c": "#E05555"},
    "DELETE ALL":  {"bg": "#FFEBEE", "c": "#B71C1C"},
}

NAV_ITEMS = [
    ("landpage",  "🗂️", "My Files"),
    ("viewdata",  "🔍", "View Data"),
    ("dashboard", "📊", "Dashboard"),
    ("report",    "📋", "Report"),
    ("audit",     "🛡️", "Audit Log"),
]

RS_COL_GROUPS = [
    {"group": "Item Info", "color": "#F5F5F5", "cols": [
        "Department", "Section", "Subclass", "Barcode", "TPNA", "ID", "Item Name"]},
    {"group": "Pack Info", "color": "#FAFAFA", "cols": [
        "No. of unit in case", "No. of unit in inner", "Tray total number",
        "Express Picking type", "HDET picking type"]},
    {"group": "Price", "color": "#FFF9C4", "cols": [
        "EDLP Price by Format", "Mer Price (incl. vat7%)", "COST", "%MOR (from EDLP)"]},
    {"group": "Range Architecture", "color": "#E8F8F5", "cols": [
        "AS-IS planograms applied", "TO-BE planograms applied",
        "AS-IS Stores Applied", "TO-BE stores applied"]},
    {"group": "Sales & Forecast", "color": "#FFF9C4", "cols": [
        "Avg Units 52wk/ Forecast new item sales", "Supplier Pack Size",
        "AS-IS Sale Total Units", "TO-BE Total Sale Units", "Total Units change",
        "AS-IS Total Sales (Ex Vat)", "TO-BE Total Sales (Ex Vat)", "Total Sales change (Ex Vat)"]},
    {"group": "Margin", "color": "#FCE4EC", "cols": [
        "AS-IS Total Margin (Ex Vat)", "TO-BE Total Margin (Ex Vat)", "Total Margin change (Ex Vat)"]},
    {"group": "Status", "color": "#FCE4EC", "cols": ["Status", "Check Range To-be Waterfall"]},
]

FILL_COLORS = {
    "display": {"bg": "#FFFDE7", "text": "#5D4037"},
    "mer":     {"bg": "#FCE4EC", "text": "#880E4F"},
    "formula": {"bg": "#F5F5F5", "text": "#616161"},
}

# ── Auth ─────────────────────────────────────────────────────────────────────
def current_user() -> dict:
    return st.session_state.get("auth_user", {
        "employee_id": "DEV001", "name": "Developer",
        "role": "admin", "login_time": datetime.now().strftime("%H:%M"),
    })

def is_admin() -> bool:
    return current_user().get("role") == "admin"

def logout():
    st.session_state.auth_user = None

def add_audit(action: str, detail: str = ""):
    u = current_user()
    record = {
        "id": u["employee_id"], "name": u["name"], "role": u["role"].upper(),
        "session": datetime.now().strftime("%H:%M"),
        "date": datetime.now().strftime("%d/%m/%Y"),
        "action": action, "detail": detail,
    }
    if "audit_log" in st.session_state:
        st.session_state.audit_log.append(record)

# ── Session state ─────────────────────────────────────────────────────────────
def init_session_state():
    defaults = {
        "raw_files": [], "merged_df": None, "merge_log": [], "audit_log": [],
        "display_cols": None, "canvases": {}, "rangesheet_meta": {},
        "view_sheet": RS_SHEETS[1], "view_vis_cols": None,
        "last_uploaded_preview": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v
    if st.session_state.get("merged_df") is None:
        snap = load_merged_snapshot()
        if snap is not None:
            st.session_state.merged_df = snap
            st.session_state.merge_log = [f"💾 Loaded snapshot ({len(snap):,} rows)"]

# ── CSS ───────────────────────────────────────────────────────────────────────
def inject_css():
    st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Sarabun:wght@400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter','Sarabun',system-ui,sans-serif !important; }
.main { background: #F2EDE8 !important; }

/* Eliminate all default top padding/margins from Streamlit's structural blocks */
.block-container { padding: 0 !important; max-width: 100% !important; }
[data-testid="stMainBlockContainer"] { padding: 0rem 2rem 1.5rem 2rem !important; margin-top: 0 !important; }
[data-testid="stHeader"] { display: none !important; }

/* ── Sticky Fixed Left Sidebar ── */
[data-testid="stSidebar"] {
    position: fixed !important;
    top: 0 !important;
    left: 0 !important;
    height: 100vh !important;
    background-color: #ffffff !important;
    z-index: 999,999 !important;
}

/* Aggressive elimination of top padding inside the sidebar container */
[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
    padding-top: 0rem !important;
    margin-top: 0rem !important;
}

/* ── Sidebar NAV buttons ── */
[data-testid="stSidebar"] .stButton > button {
  background: transparent !important;
  color: #444 !important;
  border: none !important;
  border-radius: 8px !important;
  font-size: 14px !important;
  width: 100% !important;
  text-align: left !important;
  justify-content: flex-start !important;
  padding: 9px 14px !important;
  margin: 2px 0 !important;
  transition: background 0.15s;
}
[data-testid="stSidebar"] .stButton > button:hover {
  background: #E8F8F5 !important;
  color: #2BBFA4 !important;
}

/* ── Main content buttons ── */
.stButton > button {
  background: #2BBFA4 !important;
  color: #fff !important;
  border: none !important;
  border-radius: 8px !important;
  font-weight: 600 !important;
  font-size: 13px !important;
}
.stButton > button:hover { background: #22A08A !important; }

/* Navbar Clean Logout Style */
.nav-logout-btn > button {
  background: transparent !important;
  color: #E05555 !important;
  border: 1.5px solid #E05555 !important;
  font-size: 12px !important;
  padding: 4px 10px !important;
  border-radius: 6px !important;
}
.nav-logout-btn > button:hover {
  background: #FEE8E8 !important;
  color: #B71C1C !important;
  border-color: #B71C1C !important;
}

/* ── Download buttons ── */
.stDownloadButton > button {
  background: transparent !important;
  color: #2BBFA4 !important;
  border: 1.5px solid #2BBFA4 !important;
  border-radius: 8px !important;
  font-weight: 600 !important; font-size: 13px !important;
}
.stDownloadButton > button:hover { background: #E8F8F5 !important; }

/* ── Metric cards ── */
[data-testid="stMetric"] {
  background: #fff !important; border: 1px solid #E0D9D2 !important;
  border-radius: 12px !important; padding: 14px 16px !important;
  box-shadow: 0 1px 4px rgba(0,0,0,.06) !important;
}
[data-testid="stMetricValue"] { color: #1A1A1A !important; font-weight: 800 !important; }
[data-testid="stMetricLabel"] { color: #888 !important; font-size: 11px !important;
  font-weight: 700 !important; text-transform: uppercase !important; }

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] { border-bottom: 2px solid #E0D9D2 !important; gap: 4px; }
.stTabs [data-baseweb="tab"] { border-radius: 8px 8px 0 0 !important;
  font-weight: 600 !important; color: #888 !important; }
.stTabs [aria-selected="true"] { color: #2BBFA4 !important;
  border-bottom: 2px solid #2BBFA4 !important; }

/* ── Expander ── */
[data-testid="stExpander"] { border: 1px solid #E0D9D2 !important;
  border-radius: 10px !important; background: #fff !important; }

/* ── File uploader ── */
[data-testid="stFileUploader"] {
  background: #F8F4F0 !important; border: 2px dashed #D6CFC8 !important;
  border-radius: 12px !important;
}

/* ── Dataframe ── */
[data-testid="stDataFrame"] { border-radius: 10px !important;
  border: 1px solid #E0D9D2 !important; }

hr { border-color: #E0D9D2 !important; margin: 8px 0 !important;}

/* ── Hide Streamlit chrome, keep sidebar toggle ── */
#MainMenu, footer { visibility: hidden; }
[data-testid="stSidebarCollapseButton"] { visibility: visible !important; }
[data-testid="collapsedControl"] { visibility: visible !important; display: flex !important; }
.stDeployButton { display: none; }
[data-testid="stSidebarNav"] { display: none !important; }
</style>
""", unsafe_allow_html=True)

# ── Shared UI components ───────────────────────────────────────────────────
def render_sidebar(active: str = "landpage"):
    with st.sidebar:
        # Logo and Title stacked directly at the bleeding top of the sidebar
        st.markdown(
            f"""
            <div style="display: flex; align-items: center; gap: 10px; padding: 15px 12px 5px 12px;">
                <span style="font-size: 24px;">📊</span>
                <div>
                    <h3 style="margin: 0; font-size: 18px; color: #1A1A1A;">RangeSheet</h3>
                    <p style="margin: 0; font-size: 11px; color: #888; font-weight: 600; text-transform: uppercase;">Management Platform</p>
                </div>
            </div>
            """, 
            unsafe_allow_html=True
        )
        st.divider()

        # Navigation Elements
        for page_id, icon, label in NAV_ITEMS:
            if page_id == active:
                st.markdown(
                    f'<div style="background:#E8F8F5;border-left:3px solid #2BBFA4;'
                    f'border-radius:0 8px 8px 0;padding:8px 12px 8px 11px;margin:2px 0;">'
                    f'<span style="color:#2BBFA4;font-weight:600;font-size:14px;">'
                    f'{icon} {label}</span></div>',
                    unsafe_allow_html=True)
            else:
                if st.button(f"{icon} {label}", key=f"nav_{page_id}", use_container_width=True):
                    st.switch_page(f"pages/{page_id}.py")

        st.divider()

        # Meta Context Indicators 
        m = st.session_state.get("merged_df")
        if m is not None:
            st.caption(f"**Merged data:** {len(m):,} rows · {len(m.columns)} cols")

        sinfo = storage_info()
        st.caption(f"📁 {sinfo['files']} files · {sinfo['total_mb']} MB")


def render_topbar(title: str):
    """Renders page heading alongside user identity details and logout inside the navbar row."""
    user = current_user()
    role_label = "ADMIN" if is_admin() else "VIEWER"
    
    # Grid column framework to push user info and layout nicely into the top bar
    col1, col2 = st.columns([0.6, 0.4])
    
    with col1:
        st.markdown(
            f"<div style='padding:14px 0 0px; font-size:20px; font-weight:700; color:#1A1A1A;'>"
            f"<span style='color:#2BBFA4; font-size:14px; font-weight:600;'>RangeSheet</span>"
            f"<span style='color:#BBB;'> · </span>{title}</div>",
            unsafe_allow_html=True,
        )
        
    with col2:
        # Right aligned layout for username details & the clean logout button
        st.markdown(
            f"""
            <div style='display: flex; justify-content: flex-end; align-items: center; gap: 15px; padding-top: 14px; height: 100%;'>
                <div style='text-align: right; line-height: 1.2;'>
                    <span style='font-size: 13px; font-weight: 700; color: #1A1A1A;'>{user['name']}</span><br/>
                    <span style='font-size: 11px; color: #888;'>{user['employee_id']} · <code style='background:#F2EDE8; padding:1px 4px; border-radius:4px;'>{role_label}</code></span>
                </div>
                <div class="nav-logout-btn">
            """,
            unsafe_allow_html=True
        )
        if st.button("🚪 Logout", key="navbar_logout"):
            add_audit("LOGOUT")
            logout()
            st.rerun()
        st.markdown("</div></div>", unsafe_allow_html=True)
        
    st.markdown("<hr style='margin:4px 0 14px;'>", unsafe_allow_html=True)


# ── File I/O helpers ──────────────────────────────────────────────────────────
_HEADER_KEYWORDS = {
    "department", "barcode", "item name", "section", "subclass", "tpna",
    "status", "edlp price by format", "star line", "supplier pack size",
    "no. of unit in inner", "tray total number", "hdet picking type",
    "check range to-be waterfall", "item priority",
}

def _detect_header_row(raw: bytes, encoding: str) -> int:
    try:
        lines = raw.decode(encoding, errors="replace").splitlines()
        for i, line in enumerate(lines[:30]):
            cells = {c.strip().lower() for c in line.split(",")}
            if len(cells & _HEADER_KEYWORDS) >= 2:
                return i
    except Exception:
        pass
    return 0

def _detect_header_row_excel(raw: bytes) -> int:
    try:
        preview = pd.read_excel(io.BytesIO(raw), header=None, nrows=30)
        for i, row in preview.iterrows():
            cells = {str(c).strip().lower() for c in row if pd.notna(c)}
            if len(cells & _HEADER_KEYWORDS) >= 2:
                return int(i)
    except Exception:
        pass
    return 0

def _extract_rangesheet_meta(raw: bytes, encoding: str, header_row: int) -> dict:
    meta = {"dg_code": "—", "dg_name": "—", "minor_live_week": "—",
            "major_live_week": "—", "event_live_date": "—", "event_desc": "—",
            "range_arch": []}
    try:
        lines = raw.decode(encoding, errors="replace").splitlines()
        for line in lines[1:header_row]:
            cells = [c.strip().strip('"') for c in line.split(",")]
            if not cells or not cells[0]:
                continue
            k = cells[0].upper()
            v = cells[1].strip() if len(cells) > 1 else "—"
            if "DG CODE" in k:          meta["dg_code"] = v or "—"
            elif "DG NAME" in k:       meta["dg_name"] = v or "—"
            elif "MINOR LIVE" in k:    meta["minor_live_week"] = v or "—"
            elif "MAJOR LIVE" in k:    meta["major_live_week"] = v or "—"
            elif "EVENT LIVE" in k or "LIVE DATE" in k: meta["event_live_date"] = v or "—"
            elif "EVENT DES" in k:     meta["event_desc"] = v or "—"
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

def read_uploaded_file(uploaded_file):
    name = uploaded_file.name
    ext  = os.path.splitext(name)[-1].lower()
    try:
        if ext == ".csv":
            raw = uploaded_file.read()
            uploaded_file.seek(0)
            for enc in ["utf-8-sig", "utf-8", "cp874", "latin1"]:
                try:
                    header_row = _detect_header_row(raw, enc)
                    df = pd.read_csv(io.BytesIO(raw), encoding=enc,
                                     skiprows=header_row, header=0)
                    df.columns = [str(c).strip().replace('\n',' ').replace('\r','') for c in df.columns]
                    if "Department" in df.columns:
                        df = df[df["Department"].notna() & (df["Department"].astype(str).str.strip() != "")]
                    df = df.reset_index(drop=True)
                    if header_row > 0 and "rangesheet_meta" in st.session_state:
                        try:
                            st.session_state.rangesheet_meta = _extract_rangesheet_meta(raw, enc, header_row)
                        except Exception:
                            pass
                    return df
                except Exception:
                    continue
        elif ext in (".xlsx", ".xls"):
            raw = uploaded_file.read()
            header_row = _detect_header_row_excel(raw)
            df = pd.read_excel(io.BytesIO(raw), header=header_row)
            df.columns = [str(c).strip().replace('\n',' ').replace('\r','') for c in df.columns]
            if "Department" in df.columns:
                df = df[df["Department"].notna() & (df["Department"].astype(str).str.strip() != "")]
            df = df.reset_index(drop=True)
            return df
    except Exception:
        pass
    return None

def auto_merge(files_info: list) -> tuple:
    dfs = [(f["name"], f["df"]) for f in files_info if f.get("df") is not None]
    if not dfs:
        return None, []
    if len(dfs) == 1:
        return dfs[0][1].copy(), [f"📄 Single file — {dfs[0][0]}"]
    merged = pd.concat([d for _, d in dfs], ignore_index=True, sort=False)
    before = len(merged)
    merged = merged.drop_duplicates(ignore_index=True)
    log = [f"✅ Merged {len(dfs)} files → {len(merged):,} rows"]
    if len(merged) < before:
        log.append(f"🧹 Removed {before-len(merged):,} duplicates")
    return merged, log

def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    try:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        return buf.getvalue()
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
    if "TO-BE" in c or "TO BE" in c: return "mer"
    if "AS-IS" in c or "AS IS" in c or "%MOR" in c: return "formula"
    _display = {"Department","Section","Subclass","Barcode","TPNA","ID","Item Name",
                "No. of unit in case","No. of unit in inner","Tray total number",
                "Express Picking type","HDET picking type","EDLP Price by Format",
                "Star Line","Status","Check Range To-be Waterfall"}
    if c in _display: return "display"
    _mer = {"Mer Price (incl. vat7%)","COST","Supplier Pack Size",
            "Avg Units 52wk/ Forecast new item sales"}
    if c in _mer: return "mer"
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
        return {"files": len(files), "total_mb": round(total/1024/1024, 1)}
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

def load_audit_log():
    try:
        path = os.path.join(BASE_DIR, "audit", "audit_log.csv")
        if os.path.exists(path):
            return pd.read_csv(path, encoding="utf-8-sig")
    except Exception:
        pass
    return None