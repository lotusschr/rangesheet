"""
RangeSheet — Management Platform
รัน: streamlit run APP.py
   หรือ: python APP.py
"""
import streamlit as st
import pandas as pd
import os
import io
from datetime import datetime

# ═══════════════════════════════════════════════════════════════════════════
# INLINE IMPLEMENTATIONS (no external modules needed)
# ═══════════════════════════════════════════════════════════════════════════
APP_CONFIG = {
    "allowed_extensions": ["csv", "xlsx", "xls"],
    "max_file_mb": 1024,
    "merge_similarity_threshold": 0.5,
}

def current_user() -> dict:
    return st.session_state.get("auth_user", {
        "employee_id": "DEV001",
        "name": "Developer",
        "role": "admin",
        "login_time": datetime.now().strftime("%H:%M"),
    })

def is_admin() -> bool:
    return current_user().get("role") == "admin"

def has_permission(action: str) -> bool:
    return is_admin()

def logout():
    st.session_state.auth_user = None

def add_audit(action: str, detail: str = ""):
    u = current_user()
    record = {
        "id": u["employee_id"], "name": u["name"],
        "role": u["role"].upper(),
        "session": datetime.now().strftime("%H:%M"),
        "date": datetime.now().strftime("%d/%m/%Y"),
        "action": action, "detail": detail,
    }
    if "audit_log" in st.session_state:
        st.session_state.audit_log.append(record)

_RANGESHEET_HEADER_KEYWORDS = {"department", "barcode", "item name", "section", "subclass", "tpna"}

def _detect_header_row(raw: bytes, encoding: str) -> int:
    try:
        lines = raw.decode(encoding, errors="replace").splitlines()
        for i, line in enumerate(lines[:30]):
            cells = {c.strip().lower() for c in line.split(",")}
            if len(cells & _RANGESHEET_HEADER_KEYWORDS) >= 2:
                return i
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
            if "DG CODE" in k:       meta["dg_code"] = v or "—"
            elif "DG NAME" in k:     meta["dg_name"] = v or "—"
            elif "MINOR LIVE" in k:  meta["minor_live_week"] = v or "—"
            elif "MAJOR LIVE" in k:  meta["major_live_week"] = v or "—"
            elif "EVENT LIVE" in k or "LIVE DATE" in k: meta["event_live_date"] = v or "—"
            elif "EVENT DES" in k:   meta["event_desc"] = v or "—"
            elif any(t in k for t in ["MAINTAIN","DELETE","NEW SOME","NEWNEW","TOTAL SKU"]):
                try:
                    as_is = int(cells[1]) if len(cells) > 1 and str(cells[1]).strip().lstrip("-").isdigit() else 0
                    to_be = int(cells[2]) if len(cells) > 2 and str(cells[2]).strip().lstrip("-").isdigit() else 0
                    meta["range_arch"].append({"type": cells[0].strip(), "as_is": as_is,
                                               "to_be": to_be, "diff": to_be - as_is})
                except Exception:
                    pass
    except Exception:
        pass
    return meta

# ── Column group / fill definitions (mirrors RangeSheet Excel layout) ─────────
_FILL_COLORS = {
    "display":  {"bg": "#FFFDE7", "text": "#5D4037"},
    "mer":      {"bg": "#FCE4EC", "text": "#880E4F"},
    "formula":  {"bg": "#F5F5F5", "text": "#616161"},
    "starline": {"bg": "#000000", "text": "#FFFFFF"},
    "green":    {"bg": "#CCFF90", "text": "#1B5E20"},
}
_STATUS_COLORS = {
    "MAINTAIN":    {"bg": "#E8F8F5", "c": "#2BBFA4"},
    "NEW SOME":    {"bg": "#E3F2FD", "c": "#1565C0"},
    "NEW":         {"bg": "#E8F5E9", "c": "#2E7D32"},
    "NEWNEW":      {"bg": "#F3E5F5", "c": "#6A1B9A"},
    "DELETE SOME": {"bg": "#FEE8E8", "c": "#E05555"},
    "DELETE ALL":  {"bg": "#FFEBEE", "c": "#B71C1C"},
}
_RS_SHEETS = ["Range Sheet_Non-SSPOG", "Range Sheet_SSPOG",
              "StoreApply_SSPOG", "5.1 ItembyStore", "5.2 ItembyStore_SC"]
_RS_TYPE_ORDER = ["MAINTAIN", "DELETE SOME", "DELETE ALL", "NEW SOME", "NEWNEW"]
_RS_COL_GROUPS = [
    {"group": "Item Info", "color": "#F5F5F5", "cols": [
        "Department", "Section", "Subclass", "Barcode", "TPNA", "ID", "Item Name",
    ]},
    {"group": "Pack Info", "color": "#FAFAFA", "cols": [
        "No. of unit in case", "No. of unit in inner", "Tray total number",
        "Express Picking type", "HDET picking type",
    ]},
    {"group": "Price", "color": "#FFF9C4", "cols": [
        "EDLP Price by Format", "AVG Selling Price by Format",
        "Mer Price (incl. vat7%)", "COST", "%MOR (from EDLP)",
    ]},
    {"group": "Range Architecture", "color": "#E8F8F5", "cols": [
        "AS-IS planograms applied", "TO-BE planograms applied",
        "AS-IS Stores Applied", "TO-BE stores applied",
    ]},
    {"group": "Sales & Forecast", "color": "#FFF9C4", "cols": [
        "Avg Units 52wk/ Forecast new item sales",
        "Supplier Pack Size", "Range Tail YYYY",
        "AS-IS Sale Total Units", "TO-BE Total Sale Units", "Total Units change",
        "AS-IS Total Sales (Ex Vat)", "TO-BE Total Sales (Ex Vat)", "Total Sales change (Ex Vat)",
    ]},
    {"group": "Margin", "color": "#FCE4EC", "cols": [
        "AS-IS Total Margin (Ex Vat)", "TO-BE Total Margin (Ex Vat)",
        "Total Margin change (Ex Vat)",
    ]},
    {"group": "Star Line", "color": "#000000", "text_color": "#FFFFFF", "cols": [
        "Star Line",
    ]},
    {"group": "Performance", "color": "#00E676", "text_color": "#1B5E20", "cols": [
        "Item Priority", "JDA vs Actual", "Actual-Actual",
    ]},
    {"group": "Status", "color": "#9E9E9E", "text_color": "#FFFFFF", "cols": [
        "Status",
    ]},
    {"group": "Range Check", "color": "#BDBDBD", "text_color": "#424242", "cols": [
        "Check Range To-be Waterfall",
    ]},
]

def _get_fill(col: str) -> str:
    c = str(col).strip()
    if c == "Star Line": return "starline"
    if c in ("Item Priority", "JDA vs Actual", "Actual-Actual"): return "green"
    if "TO-BE" in c or "TO BE" in c: return "mer"
    if "AS-IS" in c or "AS IS" in c or "%MOR" in c: return "formula"
    _display = {"Department","Section","Subclass","Barcode","TPNA","ID","Item Name",
                "No. of unit in case","No. of unit in inner","Tray total number",
                "Express Picking type","HDET picking type","EDLP Price by Format",
                "AVG Selling Price by Format","Range Tail YYYY",
                "Status","Check Range To-be Waterfall"}
    if c in _display: return "display"
    _mer = {"Mer Price (incl. vat7%)","COST","Supplier Pack Size",
            "Avg Units 52wk/ Forecast new item sales"}
    if c in _mer: return "mer"
    return "formula"

def read_uploaded_file(uploaded_file):
    name = uploaded_file.name
    ext = os.path.splitext(name)[-1].lower()
    try:
        if ext == ".csv":
            raw = uploaded_file.read()
            uploaded_file.seek(0)
            for enc in ["utf-8-sig", "utf-8", "cp874", "latin1"]:
                try:
                    header_row = _detect_header_row(raw, enc)
                    df = pd.read_csv(io.BytesIO(raw), encoding=enc,
                                     skiprows=header_row, header=0)
                    # normalize column names (strip whitespace + internal newlines)
                    df.columns = [str(c).strip().replace('\n', ' ').replace('\r', '') for c in df.columns]
                    if "Department" in df.columns:
                        df = df[df["Department"].notna() & (df["Department"].astype(str).str.strip() != "")]
                    df = df.reset_index(drop=True)
                    # side-effect: extract RangeSheet metadata when format is detected
                    if header_row > 0 and "rangesheet_meta" in st.session_state:
                        try:
                            st.session_state.rangesheet_meta = _extract_rangesheet_meta(raw, enc, header_row)
                        except Exception:
                            pass
                    return df
                except Exception:
                    continue
        elif ext in (".xlsx", ".xls"):
            df = pd.read_excel(uploaded_file)
            df.columns = [str(c).strip().replace('\n', ' ').replace('\r', '') for c in df.columns]
            return df
    except Exception:
        pass
    return None

def auto_merge(files_info: list) -> tuple:
    dfs = [(f["name"], f["df"]) for f in files_info if f.get("df") is not None]
    if not dfs:
        return None, []
    if len(dfs) == 1:
        return dfs[0][1].copy(), [f"📄 ไฟล์เดียว — {dfs[0][0]}"]
    merged = pd.concat([d for _, d in dfs], ignore_index=True, sort=False)
    before = len(merged)
    merged = merged.drop_duplicates(ignore_index=True)
    log = [f"✅ รวม {len(dfs)} ไฟล์ → {len(merged):,} แถว"]
    if len(merged) < before:
        log.append(f"🧹 ลบแถวซ้ำ {before - len(merged):,} แถว")
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

BASE_DIR = os.path.join(os.path.dirname(__file__), "rangesheet_data")

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

def load_audit_log():
    try:
        path = os.path.join(BASE_DIR, "audit", "audit_log.csv")
        if os.path.exists(path):
            return pd.read_csv(path, encoding="utf-8-sig")
    except Exception:
        pass
    return None

# ═══════════════════════════════════════════════════════════════════════════
# PAGE CONFIG + CSS
# ═══════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="RangeSheet",
    page_icon="🗂️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Sarabun:wght@400;500;600;700&display=swap');

/* ── Global ── */
html, body, [class*="css"] {
  font-family: 'Inter','Sarabun',system-ui,sans-serif !important;
}
.main { background: #F2EDE8 !important; }
.block-container { padding: 0 !important; max-width: 100% !important; }

/* ── Sidebar ── */
[data-testid="stSidebar"] {
  background: #1C1C1E !important;
  min-width: 240px !important; max-width: 240px !important;
}
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] div {
  color: rgba(255,255,255,0.75) !important;
}
/* Radio nav items */
[data-testid="stSidebar"] [data-testid="stRadio"] label {
  color: rgba(255,255,255,0.6) !important;
  font-size: 14px !important;
  padding: 8px 12px !important;
  border-radius: 8px !important;
  display: block !important;
}
[data-testid="stSidebar"] [data-testid="stRadio"] label:hover {
  background: #2A2A2D !important;
  color: #fff !important;
}
[data-testid="stSidebar"] [data-testid="stRadio"] [aria-checked="true"] + div label,
[data-testid="stSidebar"] [data-testid="stRadio"] input:checked + div {
  color: #fff !important;
}
/* Sidebar buttons */
[data-testid="stSidebar"] .stButton > button {
  background: rgba(255,255,255,0.07) !important;
  color: rgba(255,255,255,0.7) !important;
  border: 1px solid rgba(255,255,255,0.12) !important;
  border-radius: 8px !important;
  font-size: 13px !important;
  width: 100% !important;
}
[data-testid="stSidebar"] .stButton > button:hover {
  background: rgba(255,255,255,0.14) !important;
  color: #fff !important;
}
/* Sidebar divider */
[data-testid="stSidebar"] hr {
  border-color: rgba(255,255,255,0.1) !important;
}

/* ── Top content area padding ── */
[data-testid="stMainBlockContainer"] {
  padding: 1.5rem 2rem !important;
}

/* ── Buttons (main area) ── */
.stButton > button {
  background: #2BBFA4 !important;
  color: #fff !important;
  border: none !important;
  border-radius: 8px !important;
  font-weight: 600 !important;
  font-size: 13px !important;
}
.stButton > button:hover { background: #22A08A !important; }

/* ── Download buttons ── */
.stDownloadButton > button {
  background: transparent !important;
  color: #2BBFA4 !important;
  border: 1.5px solid #2BBFA4 !important;
  border-radius: 8px !important;
  font-weight: 600 !important;
  font-size: 13px !important;
}
.stDownloadButton > button:hover {
  background: #E8F8F5 !important;
}

/* ── Metric cards ── */
[data-testid="stMetric"] {
  background: #fff !important;
  border: 1px solid #E0D9D2 !important;
  border-radius: 12px !important;
  padding: 14px 16px !important;
  box-shadow: 0 1px 4px rgba(0,0,0,.06) !important;
}
[data-testid="stMetricValue"] { color: #1A1A1A !important; font-weight: 800 !important; }
[data-testid="stMetricLabel"] { color: #888 !important; font-size: 11px !important; font-weight: 700 !important; text-transform: uppercase !important; }

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] { border-bottom: 2px solid #E0D9D2 !important; gap: 4px; }
.stTabs [data-baseweb="tab"] { border-radius: 8px 8px 0 0 !important; font-weight: 600 !important; color: #888 !important; }
.stTabs [aria-selected="true"] { color: #2BBFA4 !important; border-bottom: 2px solid #2BBFA4 !important; }

/* ── Expander ── */
[data-testid="stExpander"] { border: 1px solid #E0D9D2 !important; border-radius: 10px !important; background: #fff !important; }

/* ── File uploader ── */
[data-testid="stFileUploader"] {
  background: #F8F4F0 !important;
  border: 2px dashed #D6CFC8 !important;
  border-radius: 12px !important;
}

/* ── Dataframe ── */
[data-testid="stDataFrame"] { border-radius: 10px !important; border: 1px solid #E0D9D2 !important; }

/* ── Divider ── */
hr { border-color: #E0D9D2 !important; }

/* ── Info/warn/error boxes ── */
.rs-info  { background:#E8F8F5; border-left:4px solid #2BBFA4; border-radius:8px; padding:10px 14px; font-size:13px; margin:8px 0; }
.rs-warn  { background:#FEF3E0; border-left:4px solid #E08A20; border-radius:8px; padding:10px 14px; font-size:13px; margin:8px 0; }
.rs-error { background:#FEE8E8; border-left:4px solid #E05555; border-radius:8px; padding:10px 14px; font-size:13px; margin:8px 0; }

/* ── Nav bar at bottom of page ── */
.nav-bar {
  background: #fff; border-top: 1px solid #E0D9D2;
  padding: 12px 0; margin-top: 24px;
}

/* ── Hide Streamlit chrome ── */
#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none; }

/* ── Topbar row (anchor-targeted: element after .rs-tb-anch marker) ── */
[data-testid="element-container"]:has(.rs-tb-anch)
  + [data-testid="element-container"]
  [data-testid="stHorizontalBlock"] {
  background: #fff !important;
  border-bottom: 1px solid #E0D9D2 !important;
  margin: -1.5rem -2rem 1.5rem !important;
  padding: 0 28px !important;
  align-items: center !important;
  min-height: 52px !important;
}
/* Logout pill inside topbar */
[data-testid="element-container"]:has(.rs-tb-anch)
  + [data-testid="element-container"]
  .stButton > button {
  background: transparent !important;
  border: 1px solid rgba(43,191,164,.38) !important;
  border-radius: 20px !important;
  color: #1A1A1A !important;
  font-size: 11px !important;
  font-weight: 600 !important;
  padding: 3px 14px !important;
  white-space: nowrap !important;
  box-shadow: none !important;
}
[data-testid="element-container"]:has(.rs-tb-anch)
  + [data-testid="element-container"]
  .stButton > button:hover {
  background: #E8F8F5 !important;
  border-color: #2BBFA4 !important;
  color: #2BBFA4 !important;
}

/* ── Sheet tab inactive buttons (anchor-targeted) ── */
[data-testid="element-container"]:has(.rs-tabs-anch)
  + [data-testid="element-container"]
  .stButton > button {
  background: #F8F4F0 !important;
  color: #888 !important;
  border: 1px solid #E0D9D2 !important;
  border-top: 2px solid transparent !important;
  border-radius: 6px 6px 0 0 !important;
  font-size: 12px !important;
  font-weight: 400 !important;
  padding: 5px 8px !important;
  box-shadow: none !important;
}
[data-testid="element-container"]:has(.rs-tabs-anch)
  + [data-testid="element-container"]
  .stButton > button:hover {
  background: #EDEAE7 !important;
  color: #555 !important;
  border-top-color: #B0A8A0 !important;
}
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ═══════════════════════════════════════════════════════════════════════════
_defaults = {
    "page": "files",
    "raw_files": [],
    "merged_df": None,
    "merge_log": [],
    "audit_log": [],
    "display_cols": None,
    "canvases": {},
    "rangesheet_meta": {},
    "view_sheet": _RS_SHEETS[1],
    "view_vis_cols": None,
    "sidebar_open": True,
}
for _k, _v in _defaults.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

if st.session_state.merged_df is None:
    _snap = load_merged_snapshot()
    if _snap is not None:
        st.session_state.merged_df = _snap
        st.session_state.merge_log = [f"💾 โหลดข้อมูลล่าสุด ({len(_snap):,} แถว)"]

# Sidebar visibility CSS
if not st.session_state.get("sidebar_open", True):
    st.markdown("""<style>
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="stMainBlockContainer"] { margin-left: 0 !important; }
    </style>""", unsafe_allow_html=True)

user = current_user()

# ═══════════════════════════════════════════════════════════════════════════
# NAVIGATION HELPERS
# ═══════════════════════════════════════════════════════════════════════════
PAGES = ["files", "view", "dash", "report", "audit"]
PAGE_LABELS = {
    "files":  "🗂️ My Files",
    "view":   "🔍 View Data",
    "dash":   "📊 Dashboard",
    "report": "📋 Report",
    "audit":  "🛡️ Audit Log",
}

def go(page: str):
    st.session_state.page = page
    st.rerun()

def nav_buttons():
    """Prev / Next navigation bar at the bottom of every page."""
    idx = PAGES.index(st.session_state.page)
    st.markdown("<div style='margin-top:32px;'></div>", unsafe_allow_html=True)
    st.markdown(
        "<div style='background:#fff;border-top:1px solid #E0D9D2;"
        "padding:12px 0;border-radius:0 0 12px 12px;'></div>",
        unsafe_allow_html=True)
    col_prev, col_mid, col_next = st.columns([1, 2, 1])
    with col_prev:
        if idx > 0:
            if st.button(f"◀  {PAGE_LABELS[PAGES[idx-1]]}", use_container_width=True,
                         key="nav_prev"):
                go(PAGES[idx - 1])
    with col_mid:
        st.markdown(
            f"<div style='text-align:center;color:#888;font-size:12px;padding:8px 0;'>"
            f"{'·  ' * idx}"
            f"<span style='color:#2BBFA4;font-weight:700;'>●</span>"
            f"{'  ·' * (len(PAGES)-idx-1)}"
            f"</div>",
            unsafe_allow_html=True)
    with col_next:
        if idx < len(PAGES) - 1:
            if st.button(f"{PAGE_LABELS[PAGES[idx+1]]}  ▶", use_container_width=True,
                         key="nav_next"):
                go(PAGES[idx + 1])

# ═══════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════
with st.sidebar:
    # ── Brand ──
    st.markdown("""
    <div style="padding:20px 16px 14px;border-bottom:1px solid rgba(255,255,255,.08);margin-bottom:8px;">
      <div style="display:flex;align-items:center;gap:10px;">
        <div style="width:36px;height:36px;background:#2BBFA4;border-radius:8px;
             display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0;">🗂️</div>
        <div>
          <div style="color:#fff;font-weight:700;font-size:15px;line-height:1.2;">RangeSheet</div>
          <div style="color:rgba(255,255,255,.4);font-size:11px;">Management Platform</div>
        </div>
      </div>
    </div>""", unsafe_allow_html=True)

    # ── Navigation ──
    selected = st.radio(
        "nav",
        options=PAGES,
        format_func=lambda k: PAGE_LABELS[k],
        index=PAGES.index(st.session_state.page),
        label_visibility="collapsed",
    )
    if selected != st.session_state.page:
        go(selected)

    # ── User info ──
    role_color = "#E08A20" if is_admin() else "#3B82F6"
    role_label = "ADMIN" if is_admin() else "VIEWER"
    st.markdown(f"""
    <div style="border-top:1px solid rgba(255,255,255,.08);margin:12px 0 8px;padding-top:12px;">
      <div style="display:flex;justify-content:space-between;align-items:center;padding:8px 12px;
           background:rgba(255,255,255,.05);border-radius:8px;">
        <div>
          <div style="color:#fff;font-size:13px;font-weight:600;">{user['name']}</div>
          <div style="color:rgba(255,255,255,.4);font-size:10px;margin-top:1px;">{user['employee_id']} · {user['login_time']}</div>
        </div>
        <span style="background:{role_color};color:#fff;padding:2px 8px;border-radius:10px;
              font-size:9px;font-weight:800;">{role_label}</span>
      </div>
    </div>""", unsafe_allow_html=True)

    if st.button("🚪  Logout", use_container_width=True, key="sidebar_logout"):
        add_audit("LOGOUT")
        logout()
        st.rerun()

    # ── Data summary ──
    if st.session_state.merged_df is not None:
        m = st.session_state.merged_df
        st.markdown(f"""
        <div style="margin:10px 0 0;padding:10px 12px;background:rgba(43,191,164,.12);
             border-radius:8px;border:1px solid rgba(43,191,164,.25);">
          <div style="color:rgba(255,255,255,.5);font-size:10px;font-weight:700;
               text-transform:uppercase;letter-spacing:.06em;">Merged Data</div>
          <div style="color:#fff;font-size:12px;font-weight:600;margin-top:3px;">
            {len(m):,} rows · {len(m.columns)} cols</div>
        </div>""", unsafe_allow_html=True)

    sinfo = storage_info()
    st.markdown(f"""
    <div style="margin:8px 0 0;padding:8px 12px;background:rgba(255,255,255,.04);border-radius:8px;">
      <div style="color:rgba(255,255,255,.4);font-size:10px;font-weight:700;
           text-transform:uppercase;letter-spacing:.06em;">Storage</div>
      <div style="color:rgba(255,255,255,.7);font-size:11px;margin-top:3px;">
        📁 {sinfo['files']} files · {sinfo['total_mb']} MB</div>
    </div>""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# TOP BAR
# ═══════════════════════════════════════════════════════════════════════════
_role_lbl = "ADMIN" if is_admin() else "VIEWER"
# Anchor marker — CSS targets the columns row that immediately follows this
st.markdown('<div class="rs-tb-anch"></div>', unsafe_allow_html=True)
_tb0, _tb1, _tb2 = st.columns([0.4, 5, 2])
with _tb0:
    _sb_icon = "✕" if st.session_state.get("sidebar_open", True) else "☰"
    if st.button(_sb_icon, key="topbar_sb_toggle", use_container_width=True):
        st.session_state.sidebar_open = not st.session_state.get("sidebar_open", True)
        st.rerun()
with _tb1:
    st.markdown(f"""
    <div style="padding:11px 0;font-size:13px;color:#888;">
      <span style="color:#2BBFA4;font-weight:700;">RangeSheet</span>
      <span style="margin:0 6px;">·</span>
      <span style="color:#1A1A1A;font-weight:600;">{PAGE_LABELS[st.session_state.page]}</span>
    </div>""", unsafe_allow_html=True)
with _tb2:
    _badge = "🔓" if is_admin() else "🔒"
    if st.button(f"{_badge} {user['employee_id']} · {_role_lbl}  ↩",
                 key="topbar_logout", use_container_width=True):
        add_audit("LOGOUT")
        logout()
        st.rerun()

# ═══════════════════════════════════════════════════════════════════════════
# PAGE 1 — MY FILES
# ═══════════════════════════════════════════════════════════════════════════
if st.session_state.page == "files":
    st.markdown("<div style='font-size:22px;font-weight:700;margin-bottom:2px;'>My Files</div>", unsafe_allow_html=True)
    st.markdown("<div style='color:#888;font-size:13px;margin-bottom:18px;'>อัปโหลด .csv / .xlsx → auto-merge → บันทึกลงเครื่อง</div>", unsafe_allow_html=True)

    col_left, col_right = st.columns([1, 1.8], gap="large")

    with col_left:
        st.subheader("Upload Files")
        uploaded = st.file_uploader(
            "ลากไฟล์มาวาง หรือ Browse",
            type=APP_CONFIG["allowed_extensions"],
            accept_multiple_files=True,
        )
        if uploaded:
            new_files, failed = [], []
            for f in uploaded:
                if any(x["name"] == f.name for x in st.session_state.raw_files):
                    continue
                df = read_uploaded_file(f)
                if df is None:
                    failed.append(f.name)
                    continue
                f.seek(0)
                path = save_file(f.read(), f.name)
                new_files.append({
                    "name": f.name, "df": df,
                    "size": f"{round(f.size/1024,1)} KB",
                    "date": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "rows": len(df), "cols": len(df.columns),
                    "saved": bool(path),
                })
            if failed:
                st.error(f"อ่านไม่ได้: {', '.join(failed)}")
            if new_files:
                st.session_state.raw_files.extend(new_files)
                merged, log = auto_merge(st.session_state.raw_files)
                st.session_state.merged_df = merged
                st.session_state.merge_log = log
                st.session_state.display_cols = None
                save_merged_snapshot(merged)
                add_audit("UPLOAD", f"{len(new_files)} files → {len(merged):,} rows")
                st.success(f"✅ อัปโหลด {len(new_files)} ไฟล์ — {len(merged):,} แถว")
                st.rerun()

        if st.session_state.merge_log:
            st.subheader("Merge Log")
            for entry in st.session_state.merge_log:
                st.write(entry)

        if st.session_state.raw_files:
            if st.button("🗑️ Clear all files", key="clear_all"):
                st.session_state.raw_files = []
                st.session_state.merged_df = None
                st.session_state.merge_log = []
                st.session_state.display_cols = None
                add_audit("CLEAR FILES")
                st.rerun()

    with col_right:
        st.subheader("Uploaded Files")
        if not st.session_state.raw_files:
            st.info("ยังไม่มีไฟล์ — ลากไฟล์มาวางด้านซ้าย")
        else:
            for f in st.session_state.raw_files:
                st.write(f"📄 **{f['name']}** — {f['rows']:,} rows · {f['cols']} cols · {f['size']}")

        if st.session_state.merged_df is not None:
            m = st.session_state.merged_df
            st.divider()
            c1, c2, c3 = st.columns(3)
            c1.metric("Files", len(st.session_state.raw_files))
            c2.metric("Merged Rows", f"{len(m):,}")
            c3.metric("Columns", len(m.columns))

            with st.expander(f"👁️ Preview (50 แถวแรก จาก {len(m):,})"):
                st.dataframe(m.head(50), use_container_width=True, hide_index=True)

            d1, d2 = st.columns(2)
            with d1:
                st.download_button("⬇️ Export .xlsx", df_to_xlsx_bytes(m),
                    file_name=f"merged_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
            with d2:
                st.download_button("⬇️ Export .csv", df_to_csv_bytes(m),
                    file_name=f"merged_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv", use_container_width=True)

    nav_buttons()

# ═══════════════════════════════════════════════════════════════════════════
# PAGE 2 — VIEW DATA
# ═══════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "view":
    merged = st.session_state.merged_df
    if merged is None:
        st.warning("ยังไม่มีข้อมูล — ไปที่หน้า My Files เพื่ออัปโหลดไฟล์ก่อน")
    else:
        all_cols = list(merged.columns)

        # ── Match _RS_COL_GROUPS definitions to actual df columns ──────────────
        def _nc(s): return str(s).lower().strip().replace('\n', ' ').replace('  ', ' ')
        _dcm = {_nc(c): c for c in all_cols}
        def _fc(key): return _dcm.get(_nc(key))

        active_groups = []
        for _grp in _RS_COL_GROUPS:
            _matched = [_fc(k) for k in _grp["cols"] if _fc(k) is not None]
            if _matched:
                active_groups.append({**_grp, "matched": _matched})
        _all_grouped = [c for g in active_groups for c in g["matched"]]
        _remaining   = [c for c in all_cols if c not in _all_grouped]
        if _remaining:
            active_groups.append({"group": "Other", "color": "#FAFAFA", "matched": _remaining})

        if st.session_state.view_vis_cols is None:
            st.session_state.view_vis_cols = _all_grouped[:] or all_cols[:]

        # ── Sheet tabs — active=HTML div, inactive=st.button (CSS makes them tab-shaped)
        st.markdown('<div class="rs-tabs-anch"></div>', unsafe_allow_html=True)
        _tab_cols = st.columns(len(_RS_SHEETS))
        for _ti, (_tcol, _s) in enumerate(zip(_tab_cols, _RS_SHEETS)):
            with _tcol:
                if _s == st.session_state.view_sheet:
                    st.markdown(
                        f'<div style="padding:5px 4px;text-align:center;'
                        f'border:1px solid #E0D9D2;border-top:2px solid #2BBFA4;'
                        f'border-radius:6px 6px 0 0;background:#fff;'
                        f'color:#1A1A1A;font-weight:700;font-size:12px;'
                        f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">'
                        f'{_s}</div>', unsafe_allow_html=True)
                else:
                    if st.button(_s, key=f"stab_{_ti}", use_container_width=True):
                        st.session_state.view_sheet = _s
                        st.rerun()
        st.markdown('<div style="border-top:1px solid #E0D9D2;margin-bottom:14px;"></div>',
                    unsafe_allow_html=True)

        tab_all, tab_canvas = st.tabs(["📊 All Data", "🎨 New Canvas"])

        with tab_all:
            _meta = st.session_state.rangesheet_meta

            # ── DG Card | Range Architecture | Color Legend ────────────────────
            _dg_c, _arch_c, _leg_c = st.columns([1.1, 2.8, 0.75])

            with _dg_c:
                _dg_rows = ""
                for _k, _v, _hl in [
                    ("DG CODE",         _meta.get("dg_code",        "—"), True),
                    ("DG NAME",         _meta.get("dg_name",        "—"), True),
                    ("MINOR LIVE WEEK", _meta.get("minor_live_week","—"), False),
                    ("MAJOR LIVE WEEK", _meta.get("major_live_week","—"), False),
                    ("Event Live Date", _meta.get("event_live_date","—"), False),
                    ("Event Des",       _meta.get("event_desc",     "—"), False),
                ]:
                    _bg = "#FFF9C4" if _hl else "transparent"
                    _dg_rows += (
                        f'<div style="display:flex;justify-content:space-between;'
                        f'margin-bottom:5px;font-size:12px;">'
                        f'<span style="color:#888;">{_k}</span>'
                        f'<span style="font-weight:600;background:{_bg};padding:0 4px;'
                        f'border-radius:3px;">{_v}</span></div>'
                    )
                st.markdown(f'''
                <div style="background:#fff;border-radius:10px;border:1px solid #E0D9D2;padding:14px 16px;">
                  <div style="font-size:10px;font-weight:700;color:#888;text-transform:uppercase;
                       letter-spacing:.07em;margin-bottom:10px;">Display Group</div>
                  {_dg_rows}
                  <div style="margin-top:12px;display:flex;gap:6px;">
                    <div style="flex:1;background:#2BBFA4;color:#fff;border-radius:8px;padding:8px;
                         text-align:center;font-size:11px;font-weight:700;">1. SELECT DG</div>
                    <div style="flex:1;background:#E08A20;color:#fff;border-radius:8px;padding:8px;
                         text-align:center;font-size:11px;font-weight:700;">2. SUBMIT RANGE</div>
                  </div>
                </div>''', unsafe_allow_html=True)

            with _arch_c:
                _arch_rows = list(_meta.get("range_arch", []))
                if not _arch_rows:
                    _sc = next((c for c in all_cols if c.strip().lower() == "status"), None)
                    if _sc:
                        for _t in _RS_TYPE_ORDER:
                            _cnt = int((merged[_sc].astype(str).str.strip() == _t).sum())
                            _ai = _cnt if _t == "MAINTAIN" or "DELETE" in _t else 0
                            _tb = _cnt if _t == "MAINTAIN" or "NEW" in _t else 0
                            _arch_rows.append({"type": _t, "as_is": _ai, "to_be": _tb, "diff": _tb - _ai})
                        _tai = sum(r["as_is"] for r in _arch_rows)
                        _ttb = sum(r["to_be"] for r in _arch_rows)
                        _arch_rows.append({"type": "TOTAL SKU", "as_is": _tai, "to_be": _ttb,
                                           "diff": _ttb - _tai, "_total": True})
                _atbody = ""
                for _r in _arch_rows:
                    _tot = _r.get("_total") or "TOTAL" in str(_r.get("type","")).upper()
                    _pct = "%" in str(_r.get("type",""))
                    _rbg = "#1C1C1E" if _tot else ("#F5F5F5" if _pct else "transparent")
                    _rc  = "#fff" if _tot else "#1A1A1A"
                    _fw  = "700" if _tot else "400"
                    _atbody += (
                        f'<tr style="background:{_rbg};">'
                        f'<td style="padding:6px 10px;font-weight:{_fw};color:{_rc};">{_r.get("type","")}</td>'
                        f'<td style="padding:6px 8px;text-align:center;color:{_rc};">{_r.get("as_is","")}</td>'
                        f'<td style="padding:6px 8px;text-align:center;color:{_rc};">{_r.get("to_be","")}</td>'
                        f'<td style="padding:6px 8px;text-align:center;color:{"#fff" if _tot else "#888"};">{_r.get("diff","")}</td>'
                        f'</tr>'
                    )
                st.markdown(f'''
                <div style="background:#fff;border-radius:10px;border:1px solid #E0D9D2;overflow:hidden;">
                  <div style="padding:10px 14px;border-bottom:1px solid #E0D9D2;
                       display:flex;justify-content:space-between;align-items:center;">
                    <span style="font-weight:700;font-size:13px;">Range Architecture</span>
                    <span style="font-size:11px;color:#2BBFA4;font-weight:600;">AVG LRD CASE %</span>
                  </div>
                  <div style="overflow-x:auto;">
                    <table style="width:100%;border-collapse:collapse;font-size:11px;">
                      <thead><tr style="background:#F5F5F5;">
                        <th style="padding:7px 10px;text-align:left;font-weight:700;color:#888;
                             border-bottom:1px solid #E0D9D2;min-width:120px;">TYPE</th>
                        <th style="padding:7px 8px;text-align:center;font-weight:700;color:#888;
                             border-bottom:1px solid #E0D9D2;font-size:10px;min-width:55px;">AS IS</th>
                        <th style="padding:7px 8px;text-align:center;font-weight:700;color:#888;
                             border-bottom:1px solid #E0D9D2;font-size:10px;min-width:55px;">TO BE</th>
                        <th style="padding:7px 8px;text-align:center;font-weight:700;color:#888;
                             border-bottom:1px solid #E0D9D2;font-size:10px;min-width:55px;">DIFF</th>
                      </tr></thead>
                      <tbody>{_atbody}</tbody>
                    </table>
                  </div>
                </div>''', unsafe_allow_html=True)

            with _leg_c:
                st.markdown('''
                <div style="background:#fff;border-radius:10px;border:1px solid #E0D9D2;padding:14px 16px;">
                  <div style="font-size:10px;font-weight:700;color:#888;text-transform:uppercase;
                       letter-spacing:.07em;margin-bottom:10px;">Color Note</div>
                  <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
                    <div style="width:24px;height:14px;background:#FFFDE7;border:1px solid #ddd;border-radius:3px;"></div>
                    <span style="font-size:11px;">Display fill</span>
                  </div>
                  <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
                    <div style="width:24px;height:14px;background:#FCE4EC;border:1px solid #ddd;border-radius:3px;"></div>
                    <span style="font-size:11px;">Merchandiser fill</span>
                  </div>
                  <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
                    <div style="width:24px;height:14px;background:#F5F5F5;border:1px solid #ddd;border-radius:3px;"></div>
                    <span style="font-size:11px;">Formula</span>
                  </div>
                  <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
                    <div style="width:24px;height:14px;background:#000;border-radius:3px;"></div>
                    <span style="font-size:11px;">Star Line</span>
                  </div>
                  <div style="display:flex;align-items:center;gap:8px;">
                    <div style="width:24px;height:14px;background:#CCFF90;border:1px solid #ddd;border-radius:3px;"></div>
                    <span style="font-size:11px;">Performance</span>
                  </div>
                </div>''', unsafe_allow_html=True)

            st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

            # ── Controls bar ───────────────────────────────────────────────────
            _ctrl1, _ctrl2, _ctrl3 = st.columns([3, 3, 1.5])
            with _ctrl1:
                _filter_mode = st.radio("FILTER:", ["ALL", "SSPOG", "NON-SSPOG"],
                                        horizontal=True, label_visibility="visible")
            with _ctrl2:
                _search_q = st.text_input("Search", placeholder="🔎  Search item, barcode, status...",
                                          label_visibility="collapsed")
            with _ctrl3:
                _edit_on = st.toggle("⚙️ Edit Columns", key="vw_edit_cols")

            # ── Edit columns panel ─────────────────────────────────────────────
            if _edit_on:
                st.markdown('''<div style="background:#F8F4F0;border-radius:10px;padding:10px 16px 6px;
                    border:1px solid #E0D9D2;margin-bottom:10px;">
                  <div style="font-size:12px;font-weight:700;margin-bottom:6px;">
                    ⚙️ Column visibility — <span style="color:#2BBFA4;">display only, raw data unchanged</span>
                  </div></div>''', unsafe_allow_html=True)
                _new_vis = []
                for _g in active_groups:
                    st.caption(f"— {_g['group']} —")
                    _sel = st.multiselect(
                        _g["group"], _g["matched"],
                        default=[c for c in _g["matched"]
                                 if c in (st.session_state.view_vis_cols or _g["matched"])],
                        key=f"vv_{_g['group']}", label_visibility="collapsed")
                    _new_vis.extend(_sel)
                _ec1, _ec2 = st.columns(2)
                with _ec1:
                    if st.button("✅ Apply", key="vv_apply"):
                        st.session_state.view_vis_cols = _new_vis or _all_grouped[:]
                        st.rerun()
                with _ec2:
                    if st.button("↺ Show All", key="vv_reset"):
                        st.session_state.view_vis_cols = _all_grouped[:]
                        st.rerun()

            # ── Resolve display columns ────────────────────────────────────────
            _vis    = st.session_state.view_vis_cols or _all_grouped or all_cols
            disp_cols = [c for c in _vis if c in merged.columns] or all_cols[:20]

            _disp_groups = []
            for _g in active_groups:
                _gc = [c for c in _g["matched"] if c in disp_cols]
                if _gc:
                    _disp_groups.append({**_g, "disp_cols": _gc})

            # ── Apply filter & search ──────────────────────────────────────────
            df_view   = merged.copy()
            _pog_col  = next((c for c in all_cols
                              if "pog" in c.lower() and "cluster" in c.lower()), None)
            if _filter_mode == "SSPOG" and _pog_col:
                df_view = df_view[
                    df_view[_pog_col].astype(str).str.contains("SSPOG", na=False) &
                    ~df_view[_pog_col].astype(str).str.contains("Non", na=False)]
            elif _filter_mode == "NON-SSPOG" and _pog_col:
                df_view = df_view[
                    df_view[_pog_col].astype(str).str.contains("Non-SSPOG", na=False)]
            if _search_q:
                _mask = df_view.apply(
                    lambda r: r.astype(str).str.contains(_search_q, case=False, na=False).any(),
                    axis=1)
                df_view = df_view[_mask]
            df_view = df_view.reset_index(drop=True)

            # ── Metric tiles ───────────────────────────────────────────────────
            _stc = next((c for c in all_cols if c.strip().lower() == "status"), None)
            _m_maintain  = int((df_view[_stc].astype(str).str.strip() == "MAINTAIN").sum()) if _stc else 0
            _m_sspog     = int(df_view[_pog_col].astype(str).str.contains("SSPOG", na=False).sum() -
                               df_view[_pog_col].astype(str).str.contains("Non-SSPOG", na=False).sum()
                               ) if _pog_col else 0
            _m_nonsspog  = int(df_view[_pog_col].astype(str).str.contains("Non-SSPOG", na=False).sum()
                               ) if _pog_col else 0
            _m_null      = int(df_view[disp_cols].isnull().sum().sum())

            _tm1, _tm2, _tm3, _tm4, _tm5 = st.columns(5)
            _tm1.metric("Total SKUs",  f"{len(df_view):,}")
            _tm2.metric("MAINTAIN",    f"{_m_maintain:,}")
            _tm3.metric("SSPOG",       f"{_m_sspog:,}")
            _tm4.metric("Non-SSPOG",   f"{_m_nonsspog:,}")
            _tm5.metric("Null Values", f"{_m_null:,}")

            # ── Main HTML table ────────────────────────────────────────────────
            _MAX = 200
            _tdf = df_view[disp_cols].head(_MAX)

            def _cw(col): return max(80, min(200, len(str(col)) * 7 + 16))

            _h = ['<div style="overflow-x:auto;border-radius:10px;'
                  'border:1px solid #E0D9D2;margin-top:12px;">',
                  '<table style="border-collapse:collapse;font-size:11px;min-width:100%;">',
                  '<thead>']
            # group header row
            _h.append('<tr>')
            for _g in _disp_groups:
                _gtc = _g.get("text_color", "#555")
                _gbc = _g.get("color", "#F5F5F5")
                _h.append(
                    f'<th colspan="{len(_g["disp_cols"])}" style="padding:5px 8px;text-align:center;'
                    f'background:{_gbc};border:1px solid rgba(255,255,255,.15);font-size:9px;font-weight:700;'
                    f'color:{_gtc};letter-spacing:.05em;text-transform:uppercase;">{_g["group"]}</th>')
            _h.append('</tr>')
            # column header row (dark)
            _h.append('<tr style="background:#1C1C1E;">')
            for _g in _disp_groups:
                for _c in _g["disp_cols"]:
                    _fill = _get_fill(_c)
                    _dot  = _FILL_COLORS.get(_fill, {}).get("bg", "transparent") if _fill else "transparent"
                    _w    = _cw(_c)
                    _h.append(
                        f'<th style="padding:8px 10px;text-align:left;font-weight:700;'
                        f'color:rgba(255,255,255,.75);font-size:10px;letter-spacing:.04em;'
                        f'white-space:nowrap;min-width:{_w}px;max-width:{_w}px;'
                        f'border-right:1px solid rgba(255,255,255,.08);">'
                        f'<div style="display:flex;align-items:center;gap:4px;">'
                        f'<span style="width:5px;height:5px;border-radius:50%;background:{_dot};'
                        f'border:1px solid rgba(255,255,255,.3);flex-shrink:0;display:inline-block;">'
                        f'</span><span style="overflow:hidden;text-overflow:ellipsis;">{_c}</span>'
                        f'</div></th>')
            _h.append('</tr></thead><tbody>')
            # data rows
            for _i, _row in _tdf.iterrows():
                _rb = "#FFFFFF" if _i % 2 == 0 else "#F8F4F0"
                _h.append(f'<tr style="background:{_rb};">')
                for _g in _disp_groups:
                    for _c in _g["disp_cols"]:
                        _val   = _row.get(_c, "")
                        _sv    = "" if pd.isna(_val) or str(_val) == "nan" else str(_val)
                        _fill  = _get_fill(_c)
                        _w     = _cw(_c)
                        _cl    = _c.strip().lower()
                        # Cell background / text color by fill type
                        if _fill == "starline":
                            _cbg = "background:#000000;"
                            _ctxt = "color:#FFFFFF;font-weight:700;"
                        elif _fill == "green":
                            _cbg = "background:#CCFF90;"
                            _ctxt = "color:#1B5E20;font-weight:600;"
                        elif _fill == "display":
                            _cbg = "background:#FFFDE740;"
                            _ctxt = ""
                        elif _fill == "mer":
                            _cbg = "background:#FCE4EC40;"
                            _ctxt = ""
                        else:
                            _cbg = ""
                            _ctxt = ""
                        if _cl == "status" and _sv:
                            _sc2 = _STATUS_COLORS.get(_sv.strip(), {"bg": "#F5F5F5", "c": "#888"})
                            _inner = (f'<span style="background:{_sc2["bg"]};color:{_sc2["c"]};'
                                      f'padding:2px 8px;border-radius:4px;font-size:10px;'
                                      f'font-weight:700;white-space:nowrap;">{_sv}</span>')
                        elif _cl in ("barcode", "id") and _sv:
                            _inner = (f'<span style="color:#2BBFA4;font-family:monospace;'
                                      f'font-weight:600;">{_sv}</span>')
                        else:
                            _inner = f'<span style="{_ctxt}">{_sv}</span>'
                        _h.append(
                            f'<td style="padding:7px 10px;border-bottom:1px solid #E0D9D2;'
                            f'border-right:1px solid #E0D9D2;{_cbg}min-width:{_w}px;max-width:{_w}px;'
                            f'overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">'
                            f'{_inner}</td>')
                _h.append('</tr>')
            _h.append('</tbody></table></div>')
            st.markdown(''.join(_h), unsafe_allow_html=True)

            if len(df_view) > _MAX:
                st.caption(f"Showing first {_MAX} of {len(df_view):,} rows — export to see all")

            # footer
            _fc1, _fc2, _fc3 = st.columns([3, 1, 1])
            with _fc1:
                st.caption(f"{len(df_view):,} rows · {len(disp_cols)} columns · raw: {len(merged):,} rows")
            with _fc2:
                st.download_button("⬇️ Export display (.xlsx)",
                    df_to_xlsx_bytes(df_view[disp_cols]),
                    file_name=f"view_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
            with _fc3:
                st.download_button("⬇️ Raw data (.csv)", df_to_csv_bytes(merged),
                    file_name=f"raw_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv", use_container_width=True)

        with tab_canvas:
            st.markdown('''
            <div style="font-size:16px;font-weight:700;margin-bottom:4px;">New Canvas</div>
            <div style="font-size:13px;color:#888;margin-bottom:16px;">
              Select columns, filter rows, save as a named canvas. Raw data is never modified.
            </div>''', unsafe_allow_html=True)
            _cv1, _cv2 = st.columns(2)
            with _cv1:
                _cnv_srch = st.text_input("Filter", placeholder="🔎  Filter rows...",
                                          label_visibility="collapsed")
            with _cv2:
                _cnv_name = st.text_input("Name", placeholder="Canvas name e.g. SSPOG-A Items",
                                          label_visibility="collapsed")
            _cnv_cols = []
            for _g in active_groups:
                st.caption(f"— {_g['group']} —")
                _cnv_cols.extend(st.multiselect(
                    _g["group"], _g["matched"],
                    default=_g["matched"][:min(3, len(_g["matched"]))],
                    key=f"cnv_{_g['group']}", label_visibility="collapsed"))
            _cdf = merged[_cnv_cols].copy() if _cnv_cols else merged.copy()
            if _cnv_srch:
                _cdf = _cdf[_cdf.apply(
                    lambda r: r.astype(str).str.contains(_cnv_srch, case=False, na=False).any(),
                    axis=1)]
            st.dataframe(_cdf.reset_index(drop=True), use_container_width=True,
                         height=300, hide_index=True)
            st.caption(f"{len(_cdf):,} rows · {len(_cnv_cols)} columns")
            _ccb1, _ccb2 = st.columns(2)
            with _ccb1:
                if st.button("💾 Save Canvas"):
                    _nm = _cnv_name.strip() or f"Canvas {len(st.session_state.canvases)+1}"
                    st.session_state.canvases[_nm] = _cdf.copy()
                    add_audit("SAVE CANVAS", _nm)
                    st.success(f'✅ Canvas "{_nm}" saved!')
            with _ccb2:
                if _cnv_cols:
                    st.download_button("⬇️ Export Canvas", df_to_xlsx_bytes(_cdf),
                        file_name=f"{_cnv_name or 'canvas'}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)
            if st.session_state.canvases:
                st.markdown("**Saved Canvases**")
                for _cnm, _cs in st.session_state.canvases.items():
                    _sc1, _sc2 = st.columns([4, 1])
                    with _sc1:
                        st.write(f"🎨 **{_cnm}** — {len(_cs):,} rows · {len(_cs.columns)} cols")
                    with _sc2:
                        st.download_button("⬇️", df_to_xlsx_bytes(_cs),
                            file_name=f"{_cnm}.xlsx", key=f"cs_dl_{_cnm}")

    nav_buttons()

# ═══════════════════════════════════════════════════════════════════════════
# PAGE 3 — DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "dash":
    st.markdown("<div style='font-size:22px;font-weight:700;margin-bottom:18px;'>Dashboard</div>", unsafe_allow_html=True)
    merged = st.session_state.merged_df
    if merged is None:
        st.warning("ยังไม่มีข้อมูล — อัปโหลดไฟล์ที่หน้า My Files ก่อน")
    else:
        try:
            import plotly.graph_objects as go
            import plotly.express as px
            HAS_PLOTLY = True
        except ImportError:
            HAS_PLOTLY = False

        # ── Color map for TYPE values ─────────────────────────────────────
        TYPE_COLOR = {
            "MAINTAIN":    "#2BBFA4",
            "DELETE SOME": "#E05555",
            "DELETE ALL":  "#C0392B",
            "NEW SOME":    "#3B82F6",
            "NEWNEW":      "#8B5CF6",
            "NEW":         "#3B82F6",
            "DELETE":      "#E05555",
            "REVIEW":      "#E08A20",
        }

        # ── Detect TYPE column ────────────────────────────────────────────
        type_col = find_col(merged, ["type"])
        # also check if any column contains the known TYPE values
        if not type_col:
            for c in merged.columns:
                vals = merged[c].dropna().astype(str).str.upper().unique()
                if any(v in ["MAINTAIN","DELETE SOME","DELETE ALL","NEW SOME","NEWNEW","NEW","DELETE"] for v in vals):
                    type_col = c
                    break

        # ── Detect AS IS / TO BE columns ─────────────────────────────────
        def _find_cols(keywords):
            return [c for c in merged.columns
                    if any(k in str(c).upper() for k in keywords)]

        as_is_cols = _find_cols(["AS IS", "ASIS", "AS_IS"])
        to_be_cols = _find_cols(["TO BE", "TOBE", "TO_BE"])

        # ══════════════════════════════════════════════════════════════════
        # CASE A: Detected Range Architecture table structure
        # ══════════════════════════════════════════════════════════════════
        if type_col and (as_is_cols or to_be_cols):
            # Group by TYPE
            grp = merged.copy()
            grp[type_col] = grp[type_col].astype(str).str.strip().str.upper()

            # Exclude TOTAL / % rows
            data_rows = grp[~grp[type_col].isin(["TOTAL SKU","% IMPACT","TOTAL","NAN",""])]

            # Pick first AS IS / TO BE columns for Range (SKU count proxy)
            asis_col  = as_is_cols[0]  if as_is_cols  else None
            tobe_col  = to_be_cols[0]  if to_be_cols  else None

            def safe_num(series):
                return pd.to_numeric(series, errors="coerce").fillna(0)

            if asis_col:
                data_rows = data_rows.copy()
                data_rows["_asis"] = safe_num(data_rows[asis_col])
            else:
                data_rows["_asis"] = 0

            if tobe_col:
                data_rows["_tobe"] = safe_num(data_rows[tobe_col])
            else:
                data_rows["_tobe"] = 0

            summary = data_rows.groupby(type_col)[["_asis","_tobe"]].sum().reset_index()
            summary.columns = ["TYPE","AS IS","TO BE"]
            summary["CHANGE"] = summary["TO BE"] - summary["AS IS"]
            summary["COLOR"]  = summary["TYPE"].map(lambda t: TYPE_COLOR.get(t, "#888"))

            total_asis = int(summary["AS IS"].sum())
            total_tobe = int(summary["TO BE"].sum())
            total_change = total_tobe - total_asis
            pct_change = (total_change / total_asis * 100) if total_asis else 0

            # ── KPI row ──────────────────────────────────────────────────
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Total SKU (AS IS)", f"{total_asis:,}")
            k2.metric("Total SKU (TO BE)", f"{total_tobe:,}",
                      delta=f"{total_change:+,}")
            k3.metric("Net Change", f"{total_change:+,}")
            k4.metric("% Impact", f"{pct_change:.1f}%")

            st.divider()

            # ── Charts row ───────────────────────────────────────────────
            left, right = st.columns([1.6, 1], gap="large")

            with left:
                st.markdown("**📊 AS IS vs TO BE by Type**")
                if HAS_PLOTLY:
                    fig = go.Figure()
                    fig.add_trace(go.Bar(
                        name="AS IS", x=summary["TYPE"], y=summary["AS IS"],
                        marker_color="#B0BEC5", text=summary["AS IS"],
                        textposition="outside"))
                    fig.add_trace(go.Bar(
                        name="TO BE", x=summary["TYPE"], y=summary["TO BE"],
                        marker_color=summary["COLOR"].tolist(),
                        text=summary["TO BE"], textposition="outside"))
                    fig.update_layout(
                        barmode="group", height=320,
                        plot_bgcolor="#F2EDE8", paper_bgcolor="#F2EDE8",
                        margin=dict(t=20, b=20, l=10, r=10),
                        font=dict(family="Inter,sans-serif", size=12),
                        legend=dict(orientation="h", y=1.08),
                        xaxis=dict(tickfont=dict(size=11)),
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    chart_data = summary.set_index("TYPE")[["AS IS","TO BE"]]
                    st.bar_chart(chart_data, height=300)

            with right:
                st.markdown("**🍩 TO BE Distribution**")
                pie_data = summary[summary["TO BE"] > 0]
                if len(pie_data):
                    if HAS_PLOTLY:
                        fig2 = px.pie(
                            pie_data, names="TYPE", values="TO BE",
                            color="TYPE",
                            color_discrete_map={t: c for t,c in zip(pie_data["TYPE"], pie_data["COLOR"])},
                            hole=0.45,
                        )
                        fig2.update_layout(
                            height=320, showlegend=True,
                            plot_bgcolor="#F2EDE8", paper_bgcolor="#F2EDE8",
                            margin=dict(t=20, b=20, l=10, r=10),
                            font=dict(family="Inter,sans-serif", size=11),
                        )
                        fig2.update_traces(textposition="inside", textinfo="percent+label")
                        st.plotly_chart(fig2, use_container_width=True)
                    else:
                        st.bar_chart(pie_data.set_index("TYPE")["TO BE"], height=300)
                else:
                    st.info("TO BE ทั้งหมดเป็น 0")

            st.divider()

            # ── Change bar (net change per type) ─────────────────────────
            st.markdown("**📈 Net Change per Type (TO BE − AS IS)**")
            if HAS_PLOTLY:
                colors_change = ["#2BBFA4" if v >= 0 else "#E05555"
                                 for v in summary["CHANGE"]]
                fig3 = go.Figure(go.Bar(
                    x=summary["TYPE"], y=summary["CHANGE"],
                    marker_color=colors_change,
                    text=[f"{v:+}" for v in summary["CHANGE"]],
                    textposition="outside",
                ))
                fig3.add_hline(y=0, line_color="#888", line_width=1)
                fig3.update_layout(
                    height=240, plot_bgcolor="#F2EDE8", paper_bgcolor="#F2EDE8",
                    margin=dict(t=20, b=20, l=10, r=10),
                    font=dict(family="Inter,sans-serif", size=12),
                    yaxis_title="Change",
                )
                st.plotly_chart(fig3, use_container_width=True)
            else:
                st.bar_chart(summary.set_index("TYPE")["CHANGE"], height=220)

            st.divider()

            # ── Summary table (color-coded) ───────────────────────────────
            st.markdown("**📋 Summary Table**")
            for _, row in summary.iterrows():
                color = row["COLOR"]
                change_html = (
                    f'<span style="color:#2BBFA4;font-weight:700;">+{int(row["CHANGE"])}</span>'
                    if row["CHANGE"] > 0 else
                    f'<span style="color:#E05555;font-weight:700;">{int(row["CHANGE"])}</span>'
                    if row["CHANGE"] < 0 else
                    f'<span style="color:#888;">0</span>'
                )
                st.markdown(f"""
                <div style="display:grid;grid-template-columns:160px 100px 100px 100px 1fr;
                     align-items:center;padding:10px 16px;margin-bottom:4px;
                     background:#fff;border-radius:8px;border-left:4px solid {color};
                     border:1px solid #E0D9D2;border-left:4px solid {color};">
                  <span style="font-weight:600;color:{color};">{row['TYPE']}</span>
                  <span style="text-align:center;color:#555;">AS IS: <b>{int(row['AS IS'])}</b></span>
                  <span style="text-align:center;color:#555;">TO BE: <b>{int(row['TO BE'])}</b></span>
                  <span style="text-align:center;">Change: {change_html}</span>
                  <div style="height:6px;background:#E0D9D2;border-radius:4px;margin:0 8px;">
                    <div style="height:6px;background:{color};border-radius:4px;
                         width:{min(int(row['TO BE'])/max(total_tobe,1)*100,100):.0f}%;"></div>
                  </div>
                </div>""", unsafe_allow_html=True)

            # ── Additional numeric columns (Sale Impact, Margin Impact) ──
            num_cols = [c for c in merged.columns
                        if pd.api.types.is_numeric_dtype(merged[c])
                        and c not in [asis_col or "", tobe_col or ""]]
            if num_cols and type_col:
                st.divider()
                st.markdown("**💰 Impact Columns by Type**")
                try:
                    impact_df = merged[merged[type_col].isin(summary["TYPE"].tolist())]
                    impact_sum = impact_df.groupby(type_col)[num_cols[:4]].sum().reset_index()
                    st.dataframe(impact_sum, use_container_width=True, hide_index=True)
                except Exception:
                    pass

        # ══════════════════════════════════════════════════════════════════
        # CASE B: Generic data — show general stats
        # ══════════════════════════════════════════════════════════════════
        else:
            n_rows  = len(merged)
            n_cols  = len(merged.columns)
            n_files = len(st.session_state.raw_files)
            n_nulls = int(merged.isnull().sum().sum())

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Total Rows",    f"{n_rows:,}")
            k2.metric("Files Merged",  max(n_files, 1))
            k3.metric("Null Values",   f"{n_nulls:,}")
            k4.metric("Columns",       n_cols)

            st.divider()

            # Any status/type column
            for kw in [["type"],["status"],["category"],["group"]]:
                sc = find_col(merged, kw)
                if sc:
                    vc = merged[sc].astype(str).value_counts().reset_index()
                    vc.columns = [sc, "Count"]
                    st.markdown(f"**Distribution — {sc}**")
                    if HAS_PLOTLY:
                        fig = px.bar(vc, x=sc, y="Count", color=sc,
                                     color_discrete_sequence=px.colors.qualitative.Set2)
                        fig.update_layout(height=280, showlegend=False,
                                          plot_bgcolor="#F2EDE8", paper_bgcolor="#F2EDE8",
                                          margin=dict(t=10,b=10,l=10,r=10))
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.bar_chart(vc.set_index(sc), height=250)

            # Numeric column distributions
            num_cols = [c for c in merged.columns if pd.api.types.is_numeric_dtype(merged[c])]
            if num_cols:
                st.divider()
                st.markdown("**Numeric Column Totals**")
                totals = merged[num_cols[:8]].sum().reset_index()
                totals.columns = ["Column", "Total"]
                st.bar_chart(totals.set_index("Column"), height=250)

    nav_buttons()

# ═══════════════════════════════════════════════════════════════════════════
# PAGE 4 — REPORT
# ═══════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "report":
    st.markdown("<div style='font-size:22px;font-weight:700;margin-bottom:18px;'>Report</div>", unsafe_allow_html=True)
    merged = st.session_state.merged_df
    if merged is None:
        st.warning("ยังไม่มีข้อมูล — อัปโหลดไฟล์ที่หน้า My Files ก่อน")
    else:
        rt1, rt2 = st.tabs(["📦 Range Execution", "📈 Summary & Analytics"])

        with rt1:
            st.subheader("PPOG → SSPOG / Non-SSPOG Split")
            pog_col = find_col(merged, ["pog", "ppog", "sspog", "cluster"])
            if not pog_col:
                pog_col = st.selectbox("เลือก POG column",
                    ["— ไม่เลือก —"] + list(merged.columns))
                if pog_col == "— ไม่เลือก —":
                    pog_col = None

            sspog_sel = []
            if pog_col:
                st.success(f"POG column: **{pog_col}**")
                pog_vals = sorted(merged[pog_col].dropna().astype(str).unique().tolist())
                sspog_sel = st.multiselect("ค่าที่นับเป็น SSPOG", pog_vals,
                    default=[v for v in pog_vals if "ss" in v.lower() and "non" not in v.lower()])
                sspog_df = merged[merged[pog_col].astype(str).isin(sspog_sel)]
                non_df   = merged[~merged[pog_col].astype(str).isin(sspog_sel)]
            else:
                sspog_df = merged.iloc[0:0]
                non_df   = merged

            c1, c2 = st.columns(2)
            c1.metric("SSPOG rows", f"{len(sspog_df):,}")
            c2.metric("Non-SSPOG rows", f"{len(non_df):,}")

            st.divider()
            st.subheader("Output Files")
            OUTPUTS = [
                ("01", "Range by Item by Store", merged),
                ("02", "Item Store · Status · Forecast", merged),
                ("03", "Item No. · MOD-Store · Cluster", merged),
                ("04", "SSPOG File — Citrix Upload", sspog_df),
            ]
            for num, title, df_out in OUTPUTS:
                st.write(f"**{num}. {title}** — {len(df_out):,} rows")
                cx, cy = st.columns(2)
                with cx:
                    st.download_button(f"⬇️ {num} .xlsx",
                        df_to_xlsx_bytes(df_out), f"output_{num}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{num}x", use_container_width=True)
                with cy:
                    st.download_button(f"⬇️ {num} .csv",
                        df_to_csv_bytes(df_out), f"output_{num}.csv",
                        "text/csv", key=f"dl_{num}c", use_container_width=True)

        with rt2:
            st.subheader("Summary & Analytics")
            SUMMARY_VIEWS = [
                "Summary by Status", "Summary by Cluster POG",
                "Summary by POG", "Summary by Store",
                "Financial by Status", "Financial by POG",
            ]
            sel = st.selectbox("เลือก Summary View", SUMMARY_VIEWS)
            col_map = {
                "Status":  find_col(merged, ["status"]),
                "Cluster": find_col(merged, ["cluster"]),
                "POG":     find_col(merged, ["pog"]),
                "Store":   find_col(merged, ["store"]),
            }
            group_key = next((k for k in col_map if k.lower() in sel.lower()), None)
            group_col = col_map.get(group_key) if group_key else None
            val_cols = [c for c in merged.columns
                        if pd.api.types.is_numeric_dtype(merged[c])][:3]
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
                    st.error(f"สร้าง summary ไม่ได้: {e}")
            else:
                st.warning(f"ไม่พบ column ที่ตรงกับ **{sel}**")

    nav_buttons()

# ═══════════════════════════════════════════════════════════════════════════
# PAGE 5 — AUDIT LOG
# ═══════════════════════════════════════════════════════════════════════════
elif st.session_state.page == "audit":
    st.markdown("<div style='font-size:22px;font-weight:700;margin-bottom:18px;'>Audit Log</div>", unsafe_allow_html=True)
    audit_df = load_audit_log()
    if audit_df is None or audit_df.empty:
        st.info("ยังไม่มีบันทึก — ทุก action จะถูกเก็บอัตโนมัติ")
    else:
        c1, c2 = st.columns([2, 1])
        with c1:
            aq = st.text_input("ค้นหา", placeholder="🔎 ID / ชื่อ / action...",
                label_visibility="collapsed")
        with c2:
            ad = st.text_input("วันที่", placeholder="📅 dd/mm/yyyy",
                label_visibility="collapsed")

        show = audit_df.copy()
        if aq:
            mask = show.apply(lambda r: r.astype(str).str.contains(
                aq, case=False, na=False).any(), axis=1)
            show = show[mask]
        if ad and "date" in show.columns:
            show = show[show["date"].astype(str).str.contains(ad, na=False)]

        st.caption(f"Showing {len(show)} of {len(audit_df)} entries")
        st.dataframe(show.iloc[::-1].reset_index(drop=True),
            use_container_width=True, hide_index=True)
        st.download_button("⬇️ Export Audit Log (.xlsx)", df_to_xlsx_bytes(audit_df),
            "audit_log.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    nav_buttons()

# ═══════════════════════════════════════════════════════════════════════════
# LAUNCHER — `python APP.py` works without remembering `streamlit run`
# ═══════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    try:
        from streamlit.runtime import exists as _st_exists
        if not _st_exists():
            import subprocess, sys
            subprocess.run([sys.executable, "-m", "streamlit", "run", __file__])
    except Exception:
        import subprocess, sys
        subprocess.run([sys.executable, "-m", "streamlit", "run", __file__])
