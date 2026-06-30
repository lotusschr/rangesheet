"""View Data page — Minor dashboard tab + one tab per uploaded file."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
import pandas as pd
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar, render_page_nav,
    load_admin_manifest, load_admin_file_df, get_shared_db,
    is_large_file, get_large_file_preview, BASE_DIR,
    scan_hdet_dg_cascade,
    summarize_hdet_by_cluster,
    _detect_large_file_params, LARGE_FILE_CHUNK_SIZE,
)

# Inline: scan small FP/POG CSV for StoreCount per cluster
def _scan_csv_store_counts(path, filter_cols=None):
    sep, enc, hr = _detect_large_file_params(path)
    first = next(iter(pd.read_csv(path, sep=sep, encoding=enc, skiprows=hr,
        header=0, chunksize=1000, dtype=str, low_memory=False, on_bad_lines="skip")))
    low_map = {str(c).strip().lower(): c for c in first.columns}
    def _res(cands):
        for c in cands:
            if c in first.columns: return c
            if str(c).strip().lower() in low_map: return low_map[str(c).strip().lower()]
        return None
    cls_col   = _res(["POG_Cluster","ClusterName","Cluster_Name","cluster_name","Cluster Name"])
    store_col = _res(["store_no","StoreNo","store_id","Store_No","store_number","PG_Store_Number"])
    if not cls_col or not store_col:
        return {}, 0
    actual_filters = {}
    if filter_cols:
        for _k, (cands, vals) in filter_cols.items():
            col = _res(cands)
            if col and vals:
                actual_filters[col] = set(str(v) for v in vals)
    use_cols = list({cls_col, store_col} | set(actual_filters.keys()))
    agg, global_stores = {}, set()
    for chunk in pd.read_csv(path, sep=sep, encoding=enc, skiprows=hr, header=0,
            chunksize=LARGE_FILE_CHUNK_SIZE, usecols=use_cols, dtype=str,
            low_memory=False, on_bad_lines="skip"):
        mask = pd.Series(True, index=chunk.index)
        for col, vals in actual_filters.items():
            if col in chunk.columns:
                mask &= chunk[col].str.strip().isin(vals)
        chunk = chunk[mask]
        if chunk.empty: continue
        chunk[cls_col] = chunk[cls_col].fillna("").str.strip()
        global_stores.update(chunk[store_col].dropna().str.strip().unique())
        for cv, grp in chunk.groupby(cls_col, sort=False):
            if not cv: continue
            if cv not in agg: agg[cv] = set()
            agg[cv].update(grp[store_col].dropna().str.strip().unique())
    return {k: len(v-{""}) for k, v in agg.items()}, len(global_stores-{""})


@st.cache_data(show_spinner="Building POG table…")
def _scan_hdet_for_pivot(path: str, cl_sel: tuple, dg_sel: tuple,
                          div_sel: tuple, fmt_sel: tuple, mtime: int = 0):
    """Scan HDET (chunked) filtered by ClusterName — return pivot-ready DataFrame."""
    _sep, _enc, _hr = _detect_large_file_params(path)
    _c = {k: None for k in ("cls","dg","id","desc","pog","fmt","div","val")}
    _rows = []
    for _ck in pd.read_csv(path, sep=_sep, encoding=_enc, skiprows=_hr,
                            chunksize=LARGE_FILE_CHUNK_SIZE, dtype=str,
                            low_memory=False, on_bad_lines="skip"):
        if _c["cls"] is None:
            _lm = {str(x).strip().lower(): x for x in _ck.columns}
            def _fc(*cs, _lm=_lm, _ck=_ck):
                for c in cs:
                    if c in _ck.columns: return c
                    if c.lower() in _lm: return _lm[c.lower()]
                return None
            _c["cls"]  = _fc("ClusterName","Cluster_Name","cluster_name","Cluster Name")
            _c["dg"]   = _fc("DG","DG_CODE","dg_code","Display Group","Display_Group")
            _c["id"]   = _fc("ID","id","Barcode","barcode","TPNA")
            _c["desc"] = _fc("Item Name","item_name","ProductDescription","Description",
                              "Name_TH","product_name","Item_Name","TPNA_Name")
            _c["pog"]  = _fc("Name","FP_Name","FP Name","POGName","pog_name","FPName")
            _c["fmt"]  = _fc("store_Format","store_format","StoreFormat","Format")
            _c["div"]  = _fc("Div Code&Desc","Div Code & Desc","DivCode&Desc")
            _c["val"]  = _fc("Capacity","capacity","ForecastSales","forecast_new_item_sales",
                              "Avg_unit","avg_unit","Facing","facings","Value","value","qty","Qty")
        if not _c["cls"] or not _c["pog"]:
            continue
        _ck = _ck[_ck[_c["cls"]].astype(str).str.strip().isin(set(cl_sel))]
        if _ck.empty: continue
        if fmt_sel and _c["fmt"]:
            _ck = _ck[_ck[_c["fmt"]].astype(str).str.strip().isin(set(fmt_sel))]
        if div_sel and _c["div"]:
            _ck = _ck[_ck[_c["div"]].astype(str).str.strip().isin(set(div_sel))]
        if dg_sel and _c["dg"]:
            _ck = _ck[_ck[_c["dg"]].astype(str).str.strip().isin(set(dg_sel))]
        if _ck.empty: continue
        _keep = list(dict.fromkeys(v for v in _c.values() if v and v in _ck.columns))
        _rows.append(_ck[_keep].copy())
    if not _rows:
        return pd.DataFrame()
    _df = pd.concat(_rows, ignore_index=True)
    _rename = {v: k2 for k2, v in [
        ("DG_CODE", _c["dg"]), ("ID", _c["id"]),
        ("ProductDescription", _c["desc"]), ("POGName", _c["pog"]), ("Value", _c["val"]),
    ] if v and v != k2}
    return _df.rename(columns=_rename)


inject_css()
init_session_state()
render_sidebar("rawfiles")
render_topbar("View Data")

st.markdown("""
<style>
/* ── Tab bar ─────────────────────────────────────────────────────── */
[data-baseweb="tab-list"] {
    overflow-x: auto !important; overflow-y: hidden !important;
    flex-wrap: nowrap !important; scrollbar-width: thin !important;
    scrollbar-color: #2BBFA4 #E8E3DC !important;
    padding-bottom: 4px !important; gap: 2px !important;
}
[data-baseweb="tab-list"]::-webkit-scrollbar { height: 4px !important; }
[data-baseweb="tab-list"]::-webkit-scrollbar-thumb { background: #2BBFA4 !important; border-radius: 2px !important; }
[data-baseweb="tab-list"]::-webkit-scrollbar-track { background: #F0EBE3 !important; }
button[data-testid="stTabScrollLeft"],
button[data-testid="stTabScrollRight"] { display: none !important; }
[data-baseweb="tab"] {
    flex-shrink: 0 !important; white-space: nowrap !important;
    border-radius: 8px 8px 0 0 !important; background: transparent !important;
    color: #666 !important; font-weight: 600 !important;
    padding: 9px 20px !important; border: none !important;
    font-size: 14px !important; transition: background 0.15s, color 0.15s !important;
}
[data-baseweb="tab"]:hover { background: rgba(43,191,164,0.12) !important; color: #2BBFA4 !important; }
[data-baseweb="tab"][aria-selected="true"] { background: #2BBFA4 !important; color: #fff !important; }
[data-baseweb="tab-highlight"] { display: none !important; }
[data-baseweb="tab-border"] { background: #D0CAC2 !important; height: 1px !important; }

/* ── Minor dashboard ─────────────────────────────────────────────── */
.pog-cluster-header {
    background: #2BBFA4; color: #fff; font-weight: 700;
    font-size: 13px; padding: 6px 12px;
    border-radius: 4px 4px 0 0; letter-spacing: 0.05em;
    margin-bottom: 0;
}
.minor-sub-label {
    font-size: 11px; font-weight: 700; color: #333;
    margin: 10px 0 4px; text-transform: uppercase; letter-spacing: 0.04em;
}
.mini-tbl { border-collapse: collapse; width: 100%; font-size: 11px; }
.mini-tbl th {
    border: 1px solid #c8c0b8; padding: 3px 9px;
    background: #E8E3DC; text-align: center; font-weight: 700; color: #333;
}
.mini-tbl th.total-col { background: #D0CAC2; }
.mini-tbl td {
    border: 1px solid #c8c0b8; padding: 3px 9px;
    text-align: right; color: #1A1A1A;
}
.mini-tbl td.total-col { font-weight: 700; background: #F5F2EE; }
.nav-arrows {
    display: flex; gap: 2px; justify-content: flex-end;
    align-items: center; margin-bottom: 8px;
}
.nav-arrow {
    width: 26px; height: 26px; border: 1px solid #ccc;
    border-radius: 3px; background: #f5f5f5;
    display: flex; align-items: center; justify-content: center;
    font-size: 13px; color: #555; cursor: pointer;
}
.filter-icon {
    width: 26px; height: 26px; border: 1px solid #ccc;
    border-radius: 3px; background: #f5f5f5;
    display: flex; align-items: center; justify-content: center;
    font-size: 12px; color: #555;
}
.large-file-badge {
    font-size: 11px; background: #FFF3E0; color: #E65100;
    border-radius: 4px; padding: 2px 8px; font-weight: 700; margin-left: 6px;
}

/* ── Multiselect: plain text tags (no red pill background) ──────────── */
[data-testid="stMultiSelect"] [data-baseweb="tag"] {
    background-color: transparent !important;
    border: none !important;
    padding: 0 2px 0 0 !important;
    margin: 1px 2px !important;
}
[data-testid="stMultiSelect"] [data-baseweb="tag"] span:first-child {
    color: #333 !important;
    font-weight: 600 !important;
    font-size: 13px !important;
}
[data-testid="stMultiSelect"] [data-baseweb="tag"] [role="presentation"] {
    color: #888 !important;
}
</style>
""", unsafe_allow_html=True)

# ── Sync admin-pinned files from manifest — no file parsing, instant ─────────
if not st.session_state.get("_admin_synced"):
    _existing_names = {f["name"] for f in st.session_state.raw_files}
    for _meta in load_admin_manifest():
        if _meta["name"] not in _existing_names:
            st.session_state.raw_files.insert(0, {**_meta, "df": None, "pinned": True})
            _existing_names.add(_meta["name"])
    st.session_state._admin_synced = True

_all_files = st.session_state.raw_files  # include df=None; lazy-loaded per tab

# ── Shared helpers ────────────────────────────────────────────────────────────
def _dedup(df: pd.DataFrame) -> pd.DataFrame:
    seen: dict[str, int] = {}
    cols = []
    for c in df.columns:
        s = str(c)
        cols.append(f"{s}.{seen[s]}" if s in seen else s)
        seen[s] = seen.get(s, 0) + 1
    if cols != list(df.columns):
        df = df.copy(); df.columns = cols
    return df

def _short(name: str, maxlen: int = 22) -> str:
    base = os.path.splitext(name)[0]
    return (base[:maxlen] + "…") if len(base) > maxlen else base

def _find_col(df, *candidates):
    low = {str(c).strip().lower(): c for c in df.columns}
    for cd in candidates:
        if cd in df.columns:
            return cd
        if cd.lower() in low:
            return low[cd.lower()]
    return None

def _render_data(entry: dict, tab_key: str):
    _fname = entry.get("name", "")
    _path = os.path.join(BASE_DIR, "uploads", _fname)

    # ── Large-file branch (e.g. HDET, ~6.7M rows) ────────────────────────────
    # Never call load_admin_file_df() for these — that would trigger a full
    # parse of a multi-GB file. Show a cheap preview instead: total row
    # count via line-counting + first N rows only.
    if os.path.exists(_path) and is_large_file(_path):
        with st.spinner(f"Scanning {_fname}…"):
            _info = get_large_file_preview(_path)
        st.markdown(
            f'<div style="font-size:13px;font-weight:700;color:#1A1A1A;padding:8px 0 2px;">'
            f'{_fname}<span class="large-file-badge">large file — preview only</span></div>'
            f'<div style="font-size:12px;color:#888;margin-bottom:10px;">'
            f'{_info["total_rows"]:,} rows total &nbsp;·&nbsp; '
            f'{len(_info["columns"])} columns &nbsp;·&nbsp; '
            f'{_info["file_size_mb"]:,} MB &nbsp;·&nbsp; '
            f'showing first {len(_info["preview_df"]):,} rows'
            f'</div>', unsafe_allow_html=True,
        )
        st.dataframe(_dedup(_info["preview_df"]), use_container_width=True,
                     height=580, hide_index=True)
        st.caption(
            "This file is too large to load in full here. Use "
            "**Rangesheet Review** to filter it down by DG / DG_CODE."
        )
        return

    # ── Normal-size file path — unchanged from before ────────────────────────
    _df = entry.get("df")
    # Lazy-load: only parse the file when this tab is actually opened
    if _df is None:
        with st.spinner(f"Loading {_fname}…"):
            _df = load_admin_file_df(_fname)
            if _df is not None:
                entry["df"]   = _df
                entry["rows"] = len(_df)
                entry["cols"] = len(_df.columns)
    if _df is None or len(_df) == 0:
        st.warning("No data could be read from this file.")
        return
    _rows, _cols = len(_df), len(_df.columns)
    _badge = (
        '<span style="font-size:11px;background:#E8F8F5;color:#2BBFA4;'
        'border-radius:4px;padding:2px 8px;font-weight:700;margin-left:6px;">pinned</span>'
        if entry.get("pinned") else ""
    )
    st.markdown(
        f'<div style="font-size:13px;font-weight:700;color:#1A1A1A;padding:8px 0 2px;">'
        f'{_fname}{_badge}</div>'
        f'<div style="font-size:12px;color:#888;margin-bottom:10px;">'
        f'{_rows:,} rows &nbsp;·&nbsp; {_cols} columns'
        f'{"&nbsp;·&nbsp;" + entry.get("size","") if entry.get("size") else ""}'
        f'{"&nbsp;·&nbsp;" + entry.get("date","") if entry.get("date") else ""}'
        f'</div>', unsafe_allow_html=True,
    )
    _s1, _s2 = st.columns([4, 2])
    with _s1:
        _q = st.text_input("Search", placeholder="🔎  Search any column…",
                           key=f"q_{tab_key}", label_visibility="collapsed")
    with _s2:
        _col_sel = st.selectbox("In column", ["All columns"] + list(_df.columns),
                                key=f"cs_{tab_key}", label_visibility="collapsed")
    _dv = _df
    if _q.strip():
        _qt = _q.strip()
        if _col_sel != "All columns" and _col_sel in _dv.columns:
            _mask = _dv[_col_sel].astype(str).str.contains(_qt, case=False, na=False)
        else:
            _mask = pd.Series(False, index=_dv.index)
            for _c in _dv.columns:
                _mask = _mask | _dv[_c].astype(str).str.contains(_qt, case=False, na=False)
        _dv = _dv[_mask].reset_index(drop=True)
    st.dataframe(_dedup(_dv), use_container_width=True, height=580, hide_index=True)
    if _q.strip():
        st.caption(f"Found {len(_dv):,} of {_rows:,} rows matching \"{_q.strip()}\"")
    else:
        st.caption(f"{_rows:,} rows · {_cols} columns")


# ── Minor dashboard ───────────────────────────────────────────────────────────
def _render_minor():
    _db_df, _ = get_shared_db()

    if _db_df is None or _db_df.empty:
        st.info("Upload or pin files on **My Files** to populate this dashboard.")
        return

    _df = _dedup(_db_df.copy())

    # ── Column detection ──────────────────────────────────────────────────────
    # Never fall back to "Department" for div — that's a different column.
    # Never mix "Display Group" into dg_col — it conflicts with disp_col.
    _fmt_col   = _find_col(_df, "store_Format", "store_format", "StoreFormat", "Format")
    _div_col   = _find_col(_df, "Div Code&Desc", "Div Code & Desc", "DivCode&Desc", "DivCode")
    _dg_col    = _find_col(_df, "DG", "DG_CODE", "dg_code")
    _cl_col    = _find_col(_df, "ClusterName", "cluster_name", "Cluster_Name",
                           "Cluster (Planogram name)", "Cluster", "FP_Name", "pog_name")
    _store_col = _find_col(_df, "store_no", "StoreNo", "store_id", "Store_No")
    _pog_col   = _find_col(_df, "FP_Name", "pog_name", "PogName", "Cluster (Planogram name)")
    _id_col    = _find_col(_df, "ID", "id", "Barcode", "barcode", "TPNA")
    _desc_col  = _find_col(_df, "Item Name", "item_name", "ProductDescription", "Description")
    _disp_col  = _find_col(_df, "displaygroup", "Display Group", "DisplayGroup",
                           "Display_Group", "displayGroup")

    # ── DIAGNOSTIC — expand to confirm column names, then remove this block ───
    with st.expander("🔍 Column diagnostic (remove when confirmed)", expanded=False):
        st.write("**All columns in _df:**", list(_df.columns))
        st.write(f"fmt=`{_fmt_col}` | div=`{_div_col}` | dg=`{_dg_col}` "
                 f"| cl=`{_cl_col}` | disp=`{_disp_col}` | store=`{_store_col}`")

    # ── HDET cascade scan ─────────────────────────────────────────────────────
    # Scans HDET once per session and builds a bidirectional map:
    #   cascade[dg]  → {div:[...], cls:[...], fmt:[...]}
    #   cls_cascade[cls] → {dg:[...], div:[...], fmt:[...]}
    # This powers cross-filtering even for columns not in _df.
    _RF_HDET_KEY  = "rawfiles_hdet_cascade_v1"
    _RF_HDET_PATH = "rawfiles_hdet_path"
    if _RF_HDET_KEY not in st.session_state:
        _rf_hdet_path = None
        for _rfm in load_admin_manifest():
            _rfp = os.path.join(BASE_DIR, "uploads", _rfm["name"])
            if "hdet" in _rfm["name"].lower() and os.path.exists(_rfp) and is_large_file(_rfp):
                _rf_hdet_path = _rfp
                break
        st.session_state[_RF_HDET_PATH] = _rf_hdet_path  # persist for totals lookup
        if _rf_hdet_path:
            with st.spinner("Scanning HDET cascade for filter options…"):
                _scanned = scan_hdet_dg_cascade(_rf_hdet_path, {
                    "dg":  ["DG", "DG_CODE", "dg_code", "Display Group", "Display_Group"],
                    "div": ["Div Code&Desc", "Div Code & Desc", "DivCode&Desc"],
                    "cls": ["ClusterName", "Cluster_Name", "Cluster Name"],
                    "fmt": ["store_Format", "store_format", "StoreFormat", "Format"],
                })
            st.session_state[_RF_HDET_KEY] = _scanned
        else:
            st.session_state[_RF_HDET_KEY] = {
                "dg_vals": [], "fmt_vals": [], "div_vals": [], "cls_vals": [],
                "cascade": {}, "cls_cascade": {},
            }
    _rf_hdet      = st.session_state[_RF_HDET_KEY]
    _hdet_path    = st.session_state.get(_RF_HDET_PATH)
    _h_casc    = _rf_hdet.get("cascade",     {})
    _h_cls_c   = _rf_hdet.get("cls_cascade", {})

    # ── Cross-filter helpers ──────────────────────────────────────────────────
    _cur_fmt = st.session_state.get("minor_fmt", [])
    _cur_div = st.session_state.get("minor_div", [])
    _cur_dg  = st.session_state.get("minor_dg",  [])
    _cur_cl  = st.session_state.get("minor_cl",  [])

    def _xf(skip: str) -> pd.DataFrame:
        """Filter _df by all dropdowns except `skip` (only cols present in _df)."""
        _f = _df
        if skip != "fmt" and _cur_fmt and _fmt_col:
            _f = _f[_f[_fmt_col].astype(str).isin(_cur_fmt)]
        if skip != "div" and _cur_div and _div_col:
            _f = _f[_f[_div_col].astype(str).isin(_cur_div)]
        if skip != "dg"  and _cur_dg  and _dg_col:
            _f = _f[_f[_dg_col].astype(str).isin(_cur_dg)]
        if skip != "cl"  and _cur_cl  and _cl_col:
            _f = _f[_f[_cl_col].astype(str).isin(_cur_cl)]
        return _f

    def _uniq(df, col):
        if not col: return []
        return sorted(df[col].dropna().astype(str).replace("", pd.NA).dropna().unique())

    def _hdet_opts(key: str) -> list:
        """Compute valid options for `key` using HDET cascade given current selections.

        key is one of "dg", "div", "cls", "fmt".
        Builds the intersection of allowed values from each active selection.
        """
        _full = {
            "dg":  _rf_hdet.get("dg_vals",  []),
            "div": _rf_hdet.get("div_vals", []),
            "cls": _rf_hdet.get("cls_vals", []),
            "fmt": _rf_hdet.get("fmt_vals", []),
        }
        _sels = {"dg": _cur_dg, "div": _cur_div, "cls": _cur_cl, "fmt": _cur_fmt}
        result = set(_full.get(key, []))

        for src, vals in _sels.items():
            if src == key or not vals:
                continue
            allowed: set = set()
            for v in vals:
                if src == "dg":
                    allowed.update(_h_casc.get(v, {}).get(key, []))
                elif src == "cls":
                    allowed.update(_h_cls_c.get(v, {}).get(key, []))
                elif src in ("div", "fmt"):
                    # Reverse-lookup: walk cascade to find entries with this value
                    for dg_v, rel in _h_casc.items():
                        if v in rel.get(src, []):
                            if key == "dg":
                                allowed.add(dg_v)
                            else:
                                allowed.update(rel.get(key, []))
                    if src == "div" and key == "cls":
                        for cl_v, cd in _h_cls_c.items():
                            if v in cd.get("div", []):
                                allowed.add(cl_v)
            if allowed:
                result &= allowed

        return sorted(result)

    # ── Compute each dropdown's options ───────────────────────────────────────
    # If the column is in _df: cross-filter via _xf() (fast, exact row match).
    # If not in _df: cross-filter via HDET cascade (handles HDET-only columns).
    _fmt_opts = _uniq(_xf("fmt"), _fmt_col) if _fmt_col else _hdet_opts("fmt")
    _div_opts = _uniq(_xf("div"), _div_col) if _div_col else _hdet_opts("div")
    _dg_opts  = _uniq(_xf("dg"),  _dg_col)  if _dg_col  else _hdet_opts("dg")
    _cl_opts  = _uniq(_xf("cl"),  _cl_col)  if _cl_col  else _hdet_opts("cls")

    # Drop stale selections no longer present in narrowed option lists
    for _ss, _opts in [("minor_fmt", _fmt_opts), ("minor_div", _div_opts),
                        ("minor_dg",  _dg_opts),  ("minor_cl",  _cl_opts)]:
        if _ss in st.session_state:
            _valid = [v for v in st.session_state[_ss] if v in _opts]
            if _valid != st.session_state[_ss]:
                st.session_state[_ss] = _valid

    # ── Filter bar (4 dropdowns) ──────────────────────────────────────────────
    _fc1, _fc2, _fc3, _fc4 = st.columns(4)
    with _fc1:
        _fmt_sel = st.multiselect("store_Format", _fmt_opts, key="minor_fmt",
                                  placeholder="All formats")
    with _fc2:
        _div_sel = st.multiselect("Div Code&Desc", _div_opts, key="minor_div",
                                  placeholder="All divisions")
    with _fc3:
        _dg_sel = st.multiselect("DG", _dg_opts, key="minor_dg",
                                 placeholder="All DGs")
    with _fc4:
        _cl_sel = st.multiselect("ClusterName", _cl_opts, key="minor_cl",
                                 placeholder="All clusters")

    # Apply all 4 filters → _fdf used by every panel below
    _fdf = _df.copy()
    if _fmt_sel and _fmt_col:  _fdf = _fdf[_fdf[_fmt_col].astype(str).isin(_fmt_sel)]
    if _div_sel and _div_col:  _fdf = _fdf[_fdf[_div_col].astype(str).isin(_div_sel)]
    if _dg_sel  and _dg_col:   _fdf = _fdf[_fdf[_dg_col].astype(str).isin(_dg_sel)]
    if _cl_sel  and _cl_col:   _fdf = _fdf[_fdf[_cl_col].astype(str).isin(_cl_sel)]

    # ── Cluster summary (computed before columns — feeds both pivot + left panel) ─
    _summary_df = None
    if _hdet_path:
        _sum_cache_key = (f"hdet_sum_v8|dg={_dg_sel}|cl={_cl_sel}"
                          f"|div={_div_sel}|fmt={_fmt_sel}")
        if _sum_cache_key not in st.session_state:
            _sf_txt: dict = {}
            if _dg_sel:
                _sf_txt["dg"] = (["Display Group", "DG", "DG_CODE", "dg_code",
                                   "Display_Group"], _dg_sel)
            if _cl_sel:
                _sf_txt["cls"] = (["ClusterName", "Cluster_Name",
                                    "cluster_name", "Cluster Name"], _cl_sel)
            if _div_sel:
                _sf_txt["div"] = (["Div Code&Desc", "Div Code & Desc",
                                    "DivCode&Desc"], _div_sel)
            if _fmt_sel:
                _sf_txt["fmt"] = (["store_Format", "store_format",
                                    "StoreFormat", "Format"], _fmt_sel)

            with st.spinner("Loading cluster summary…"):
                _pog_df = summarize_hdet_by_cluster(
                    _hdet_path,
                    col_candidates={
                        "cls":   ["ClusterName", "Cluster_Name", "cluster_name",
                                  "Cluster Name"],
                        "store": ["PG_Store_Number","store_no","StoreNo","Store_No",
                                  "store_number","Store_Number","StoreID","store_id"],
                        "pog":   ["Name", "FP_Name", "FP Name", "FPName",
                                  "pog_name", "POGName"],
                        "id":    ["ID", "id", "Barcode", "barcode", "TPNA"],
                    },
                    filter_cols=_sf_txt if _sf_txt else None,
                )
                _tot_mask = _pog_df["ClusterName"] == "Total"
                _d = (_pog_df[~_tot_mask]
                      .sort_values("StoreCount", ascending=False)
                      .reset_index(drop=True))
                st.session_state[_sum_cache_key] = pd.concat(
                    [_d, _pog_df[_tot_mask]], ignore_index=True)

        _summary_df = st.session_state[_sum_cache_key]
        if "ItemCount" not in _summary_df.columns:
            _summary_df = _summary_df.copy()
            _summary_df["ItemCount"] = 0

    # ── Horizontal pivot: Total Store Apply + Total Item per cluster ───────────
    if _summary_df is not None and not _summary_df.empty:
        _is_tot2   = _summary_df["ClusterName"] == "Total"
        _data_p    = _summary_df[~_is_tot2]
        _tot_p     = _summary_df[_is_tot2]
        _cls_names = list(_data_p["ClusterName"])
        _store_map = dict(zip(_data_p["ClusterName"], _data_p["StoreCount"].astype(int)))
        _item_map  = dict(zip(_data_p["ClusterName"], _data_p["ItemCount"].astype(int)))
        if not _tot_p.empty:
            _cls_names.append("Total")
            _store_map["Total"] = int(_tot_p.iloc[0]["StoreCount"])
            _item_map["Total"]  = int(_tot_p.iloc[0]["ItemCount"])
        _th = "".join(
            f"<th style='padding:3px 10px;white-space:nowrap;text-align:right;"
            f"border-right:1px solid #2d3350;'>{c}</th>"
            for c in _cls_names
        )
        _sr = "".join(
            f"<td style='padding:3px 10px;text-align:right;"
            f"border-right:1px solid #2d3350;'>{_store_map[c]:,}</td>"
            for c in _cls_names
        )
        _ir = "".join(
            f"<td style='padding:3px 10px;text-align:right;"
            f"border-right:1px solid #2d3350;'>{_item_map[c]:,}</td>"
            for c in _cls_names
        )
        st.markdown(
            f"<div style='overflow-x:auto;font-size:0.78rem;margin-bottom:6px;'>"
            f"<table style='border-collapse:collapse;'>"
            f"<thead><tr style='background:#1e2130;color:#9ba3c2;'>"
            f"<th style='padding:3px 10px;text-align:left;white-space:nowrap;"
            f"border-right:1px solid #2d3350;'></th>{_th}</tr></thead>"
            f"<tbody>"
            f"<tr style='background:#161b2e;color:#e0e4f7;'>"
            f"<td style='padding:3px 10px;font-weight:600;white-space:nowrap;"
            f"border-right:1px solid #2d3350;'>Total Store Apply</td>{_sr}</tr>"
            f"<tr style='background:#1a1f33;color:#e0e4f7;'>"
            f"<td style='padding:3px 10px;font-weight:600;white-space:nowrap;"
            f"border-right:1px solid #2d3350;'>Total Item</td>{_ir}</tr>"
            f"</tbody></table></div>",
            unsafe_allow_html=True,
        )

    # ── 2-panel body: cluster table [2] | main table [5] ─────────────────────
    _col_l, _col_r = st.columns([2, 5], gap="small")

    # ── LEFT: Cluster table (data already computed above) ─────────────────────
    with _col_l:
        if _summary_df is not None:
            _is_total  = _summary_df["ClusterName"] == "Total"
            _data_rows = _summary_df[~_is_total].reset_index(drop=True)
            _total_row = _summary_df[_is_total]
            _disp_cols3 = ["ClusterName","StoreCount","Count of POGName"]
            _disp_rows3 = _data_rows[[c for c in _disp_cols3 if c in _data_rows.columns]]
            st.dataframe(
                _disp_rows3,
                hide_index=True,
                use_container_width=True,
                height=520,
                column_config={
                    "ClusterName":      st.column_config.TextColumn(
                                            "ClusterName", width="medium"),
                    "StoreCount":       st.column_config.NumberColumn(
                                            "StoreCount", width="small", format="%d"),
                    "Count of POGName": st.column_config.NumberColumn(
                                            "Count of POGName", width="small", format="%d"),
                },
            )
            if not _total_row.empty:
                _tr = _total_row.iloc[0]
                st.markdown(
                    f"<div style='font-weight:700;border-top:2px solid #555;"
                    f"padding:4px 2px;font-size:0.82rem;display:flex;gap:8px;'>"
                    f"<span style='flex:2'>Total</span>"
                    f"<span style='flex:1;text-align:right'>"
                    f"{int(_tr['StoreCount']):,}</span>"
                    f"<span style='flex:1;text-align:right'>"
                    f"{int(_tr['Count of POGName']):,}</span>"
                    f"</div>",
                    unsafe_allow_html=True,
                )
        else:
            st.caption("HDET file not found.")

    # ── RIGHT: POG_Cluster pivot from A5 file ─────────────────────────────────
    with _col_r:
        st.markdown('<div class="pog-cluster-header">POG_Cluster</div>',
                    unsafe_allow_html=True)

        if not _hdet_path:
            st.caption("HDET file not found — upload and pin the HDET file.")
        elif not _cl_sel:
            st.info("Select a **ClusterName** to display the POG_Cluster table.")
        else:
            _h_mtime = int(os.path.getmtime(_hdet_path))
            _raw_pvt = _scan_hdet_for_pivot(
                _hdet_path,
                cl_sel=tuple(sorted(_cl_sel)),
                dg_sel=tuple(sorted(_dg_sel)),
                div_sel=tuple(sorted(_div_sel)),
                fmt_sel=tuple(sorted(_fmt_sel)),
                mtime=_h_mtime,
            )
            if _raw_pvt.empty or "POGName" not in _raw_pvt.columns:
                st.caption("No data for current filters (or POGName column not found in HDET).")
            else:
                _row_keys   = [c for c in ["DG_CODE","ID","ProductDescription"]
                               if c in _raw_pvt.columns]
                _raw_pvt["POGName"] = _raw_pvt["POGName"].fillna("").str.strip()
                if "Value" in _raw_pvt.columns:
                    _raw_pvt["Value"] = pd.to_numeric(_raw_pvt["Value"], errors="coerce")
                if _row_keys and "Value" in _raw_pvt.columns:
                    try:
                        _pvt = (
                            _raw_pvt.groupby(_row_keys + ["POGName"], sort=False)["Value"]
                            .sum().unstack("POGName")
                        ).reset_index()
                        _pvt.columns.name = None
                    except Exception:
                        _pvt = _raw_pvt[_row_keys].drop_duplicates().reset_index(drop=True)
                else:
                    _pvt = (_raw_pvt[_row_keys].drop_duplicates().reset_index(drop=True)
                            if _row_keys else pd.DataFrame())
                _row_labels = _row_keys
                if _row_labels:
                    _pvt = _pvt.sort_values(_row_labels[0]).reset_index(drop=True)

<<<<<<< HEAD
                _pog_cols = [c for c in _pvt.columns if c not in _row_labels]
                _N        = len(_pog_cols)
                _has_desc = "ProductDescription" in _pvt.columns
                # ── Fixed-column widths & sticky left offsets ─────────────────
                _TH = "#1e2130"; _TC = "#2BBFA4"; _TC2 = "#1a9e8b"; _B = "#2d3350"
                _COL_W = {"DG_CODE": 74, "ID": 108, "ProductDescription": 230}
                _left_px = {}; _acc = 0
                for _cn in ["DG_CODE", "ID", "ProductDescription"]:
                    if _cn in _pvt.columns:
                        _left_px[_cn] = _acc
                        _acc += _COL_W[_cn]
                def _stkH(col, top="0px"):
                    _w = _COL_W.get(col, 100)
                    _l = _left_px.get(col, 0)
                    return (f"position:sticky;left:{_l}px;top:{top};z-index:5;"
                            f"background:{_TH};color:#9ba3c2;padding:4px 8px;"
                            f"border:1px solid {_B};font-weight:700;"
                            f"text-align:left;white-space:nowrap;min-width:{_w}px;")
                def _stkB(col, bg):
                    _w = _COL_W.get(col, 100)
                    _l = _left_px.get(col, 0)
                    return (f"position:sticky;left:{_l}px;z-index:1;"
                            f"background:{bg};padding:3px 8px;"
                            f"border:1px solid {_B};white-space:nowrap;min-width:{_w}px;")
                _grS = (f"position:sticky;top:0;z-index:3;"
                        f"background:{_TC};color:#fff;padding:4px 8px;"
                        f"border:1px solid #1a8a74;font-weight:700;text-align:center;")
                _pgS = (f"position:sticky;top:33px;z-index:3;"
                        f"background:{_TC2};color:#fff;padding:3px 6px;"
                        f"border:1px solid #1a8a74;font-weight:600;text-align:center;"
                        f"font-size:0.70rem;max-width:130px;overflow:hidden;"
                        f"text-overflow:ellipsis;white-space:nowrap;")
                # ── Build HTML ───────────────────────────────────────────────
                _ht = [
                    "<div style='overflow-x:auto;overflow-y:auto;"
                    "max-height:520px;font-size:0.78rem;'>",
                    "<table style='border-collapse:collapse;'>",
                    "<thead><tr>",
                ]
                # Row 1: DG_CODE + ID (rowspan=2, sticky) + group headers
                for _lbl in [c for c in ["DG_CODE","ID"] if c in _pvt.columns]:
                    _ht.append(f"<th rowspan='2' style='{_stkH(_lbl)}'>{_lbl}</th>")
                if _has_desc:
                    _ht.append(f"<th colspan='1' style='{_grS}'>POG CLUSTER</th>")
                if _N:
                    _cl_label = " / ".join(_cl_sel) if _cl_sel else "POG Cluster MOD fixture"
                    _ht.append(f"<th colspan='{_N}' style='{_grS}'>{_cl_label}</th>")
                _ht.append("</tr><tr>")
                # Row 2: ProductDescription (sticky) + each POGName
                if _has_desc:
                    _desc_h_style = _stkH("ProductDescription", "33px")
                    _ht.append(f"<th style='{_desc_h_style}'>ProductDescription</th>")
                for _pc in _pog_cols:
                    _spc = str(_pc).replace("<","&lt;").replace(">","&gt;")
                    _ht.append(f"<th title='{_spc}' style='{_pgS}'>{_spc}</th>")
                _ht.append("</tr></thead><tbody>")
                # Data rows — DG_CODE shown only on first row of each DG group
                _prev_dg4 = object()
                for _ri, _row in _pvt.iterrows():
                    _rbg = "#0e1120" if _ri % 2 == 0 else "#141829"
                    _ht.append(f"<tr style='background:{_rbg};'>")
                    if "DG_CODE" in _pvt.columns:
                        _dv4 = str(_row["DG_CODE"])
                        _bs = _stkB("DG_CODE", _rbg)
                        if _dv4 != _prev_dg4:
                            _ht.append(
                                f"<td style='{_bs}color:#e0e4f7;font-weight:600;"
                                f"vertical-align:top;'>{_dv4}</td>")
                            _prev_dg4 = _dv4
                        else:
                            _ht.append(f"<td style='{_bs}'></td>")
                    if "ID" in _pvt.columns:
                        _id_b_style = _stkB("ID", _rbg)
                        _ht.append(
                            f"<td style='{_id_b_style}color:#c8cde8;'>"
                            f"{_row['ID']}</td>")
                    if _has_desc:
                        _dsc = str(_row["ProductDescription"]).replace("<","&lt;").replace(">","&gt;")
                        _desc_b_style = _stkB("ProductDescription", _rbg)
                        _ht.append(
                            f"<td style='{_desc_b_style}color:#c8cde8;"
                            f"max-width:230px;overflow:hidden;text-overflow:ellipsis;'>"
                            f"{_dsc}</td>")
                    for _pc in _pog_cols:
                        _v = _row.get(_pc, float("nan"))
                        _ht.append(
                            f"<td style='color:#e0e4f7;padding:3px 8px;"
                            f"border:1px solid {_B};text-align:right;'>"
                            f"{'%.2f' % _v if pd.notna(_v) and _v != 0 else ''}</td>")
                    _ht.append("</tr>")
                _ht.append("</tbody></table></div>")
                st.markdown("".join(_ht), unsafe_allow_html=True)
                st.caption(f"{len(_pvt):,} items · {_N} POGs")
=======
        # Empty POG-cluster columns — leave blank, to be filled from data later
        if not _main_df.empty and _cl_col:
            _clusters = sorted(_fdf[_cl_col].dropna().astype(str).unique())
            if _clusters:
                _main_df = pd.concat(
                    [_main_df, pd.DataFrame(None, index=_main_df.index, columns=_clusters)],
                    axis=1,
                )

        st.dataframe(
            _dedup(_main_df),
            use_container_width=True,
            height=540,
            hide_index=True,
        )
        st.caption(f"{len(_main_df):,} items")
>>>>>>> 23daa6d634948c79f295ffa572955f0250fffda5


# ── Tab label list: [Minor] + one per uploaded file ───────────────────────────
_file_labels: list[str] = []
_seen_labels: dict[str, int] = {}
for _f in _all_files:
    _lbl = _short(_f["name"])
    if _lbl in _seen_labels:
        _seen_labels[_lbl] += 1
        _file_labels.append(f"{_lbl} ({_seen_labels[_lbl]})")
    else:
        _seen_labels[_lbl] = 0
        _file_labels.append(_lbl)

_all_tab_labels = ["Minor"] + _file_labels

# ── No files yet — still show Minor tab ──────────────────────────────────────
if not _all_files:
    (_t_minor,) = st.tabs(["Minor"])
    with _t_minor:
        _render_minor()
    render_page_nav("rawfiles")
    st.stop()

# ── Render all tabs ───────────────────────────────────────────────────────────
_all_tabs = st.tabs(_all_tab_labels)

with _all_tabs[0]:
    _render_minor()

for _i, (_tab, _entry) in enumerate(zip(_all_tabs[1:], _all_files)):
    with _tab:
        _render_data(_entry, f"file_{_i}")

render_page_nav("rawfiles")
