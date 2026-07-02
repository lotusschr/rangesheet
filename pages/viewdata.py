"""View Data page — Minor dashboard tab + one tab per uploaded file."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
import pandas as pd
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar, render_page_nav,
    load_admin_manifest, load_admin_file_df, get_shared_db,
    is_large_file, get_large_file_preview, BASE_DIR,
)

inject_css()
init_session_state()
render_sidebar("viewdata")
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

    # Detect columns
    _fmt_col  = _find_col(_df, "store_Format", "store_format", "StoreFormat", "Format")
    _div_col  = _find_col(_df, "Department", "Section", "Div", "DivCode")
    _dg_col   = _find_col(_df, "Department", "DG_CODE", "dg_code", "DG", "Section")
    _cl_col   = _find_col(_df, "Cluster (Planogram name)", "ClusterName", "cluster_name",
                          "Cluster", "FP_Name", "pog_name")
    _store_col = _find_col(_df, "store_no", "StoreNo", "store_id")
    _pog_col  = _find_col(_df, "FP_Name", "pog_name", "PogName", "Cluster (Planogram name)")
    _id_col   = _find_col(_df, "ID", "id", "Barcode", "barcode", "TPNA")
    _desc_col = _find_col(_df, "Item Name", "item_name", "ProductDescription", "Description")

    # ── Filter bar (4 columns across full width) ──────────────────────────────
    _fc1, _fc2, _fc3, _fc4 = st.columns(4)

    _fmt_opts = sorted(_df[_fmt_col].dropna().astype(str).unique()) if _fmt_col else []
    _div_opts = sorted(_df[_div_col].dropna().astype(str).unique()) if _div_col else []
    _dg_opts  = sorted(_df[_dg_col].dropna().astype(str).unique())  if _dg_col  else []
    _cl_opts  = sorted(_df[_cl_col].dropna().astype(str).unique())  if _cl_col  else []

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

    # Apply filters
    _fdf = _df.copy()
    if _fmt_sel and _fmt_col:  _fdf = _fdf[_fdf[_fmt_col].astype(str).isin(_fmt_sel)]
    if _div_sel and _div_col:  _fdf = _fdf[_fdf[_div_col].astype(str).isin(_div_sel)]
    if _dg_sel  and _dg_col:   _fdf = _fdf[_fdf[_dg_col].astype(str).isin(_dg_sel)]
    if _cl_sel  and _cl_col:   _fdf = _fdf[_fdf[_cl_col].astype(str).isin(_cl_sel)]

    # ── 3-panel body ──────────────────────────────────────────────────────────
    # Proportions: cluster table [2] | stats [1] | main table [4]
    _col_l, _col_c, _col_r = st.columns([2, 1, 4], gap="small")

    # ── LEFT: Cluster summary table ───────────────────────────────────────────
    with _col_l:
        if _cl_col:
            _grp = _fdf.groupby(_cl_col, dropna=False)
            _s_cnt = (_grp[_store_col].nunique()
                      if _store_col else _grp.size())
            _p_cnt = (_grp[_pog_col].nunique()
                      if _pog_col and _pog_col != _cl_col else _grp.size())
            _summary = pd.DataFrame({
                "ClusterName":      _s_cnt.index.astype(str),
                "StoreCount":       _s_cnt.values.astype(int),
                "Count of POGName": _p_cnt.values.astype(int),
            }).reset_index(drop=True)
            _total_row = pd.DataFrame([{
                "ClusterName":      "Total",
                "StoreCount":       int(_summary["StoreCount"].sum()),
                "Count of POGName": int(_summary["Count of POGName"].sum()),
            }])
            _disp = pd.concat([_summary, _total_row], ignore_index=True)
            st.dataframe(
                _disp, hide_index=True, use_container_width=True,
                height=min(36 * (len(_disp) + 1) + 3, 520),
            )
        else:
            st.caption("No cluster column found in data.")

    # ── CENTER: nav arrows + Total Store Apply + Total Item ───────────────────
    with _col_c:
        # Decorative navigation arrows (matching the Excel-style controls)
        st.markdown("""
        <div class="nav-arrows" style="margin-top:4px;">
            <div class="nav-arrow">↓</div>
            <div class="nav-arrow">↓</div>
            <div class="nav-arrow">↑</div>
            <div class="filter-icon">▽</div>
            <div class="filter-icon">⊞</div>
            <div class="filter-icon">⋯</div>
        </div>
        """, unsafe_allow_html=True)

        # Determine which cluster is "active" for the stats
        _active_cl = _cl_sel[0] if _cl_sel else (_cl_opts[0] if _cl_opts else None)

        # Total Store Apply
        if _store_col and _cl_col and _active_cl:
            _sub_cl = _fdf[_fdf[_cl_col].astype(str) == _active_cl]
            _store_val = int(_sub_cl[_store_col].nunique())
        elif _store_col:
            _store_val = int(_fdf[_store_col].nunique())
        else:
            _store_val = len(_fdf)
        _store_total = (_fdf[_store_col].nunique()
                        if _store_col else len(_fdf))
        _cl_label = _active_cl if _active_cl else "All"

        st.markdown(f"""
        <div class="minor-sub-label">Total Store Apply</div>
        <table class="mini-tbl">
            <tr>
                <th>{_cl_label}</th>
                <th class="total-col">Total</th>
            </tr>
            <tr>
                <td>{_store_val:,}</td>
                <td class="total-col">{int(_store_total):,}</td>
            </tr>
        </table>
        """, unsafe_allow_html=True)

        # Total Item
        if _id_col and _cl_col and _active_cl:
            _sub_cl2 = _fdf[_fdf[_cl_col].astype(str) == _active_cl]
            _item_val = int(_sub_cl2[_id_col].nunique())
        elif _id_col:
            _item_val = int(_fdf[_id_col].nunique())
        else:
            _item_val = len(_fdf)
        _item_total = (_fdf[_id_col].nunique() if _id_col else len(_fdf))

        st.markdown(f"""
        <div class="minor-sub-label">Total Item</div>
        <table class="mini-tbl">
            <tr>
                <th>{_cl_label}</th>
                <th class="total-col">Total</th>
            </tr>
            <tr>
                <td>{_item_val:,}</td>
                <td class="total-col">{int(_item_total):,}</td>
            </tr>
        </table>
        """, unsafe_allow_html=True)

    # ── RIGHT: POG_Cluster main table ─────────────────────────────────────────
    with _col_r:
        st.markdown('<div class="pog-cluster-header">POG_Cluster</div>',
                    unsafe_allow_html=True)

        # Fixed columns: DG_CODE | ID | ProductDescription
        _fixed: dict[str, pd.Series] = {}
        if _dg_col:   _fixed["DG_CODE"]             = _fdf[_dg_col].astype(str)
        if _id_col:   _fixed["ID"]                  = _fdf[_id_col].astype(str)
        if _desc_col: _fixed["ProductDescription"]  = _fdf[_desc_col].astype(str)

        if _fixed:
            _main_df = pd.DataFrame(_fixed).drop_duplicates().reset_index(drop=True)
        else:
            _main_df = pd.DataFrame({"DG_CODE": [], "ID": [], "ProductDescription": []})

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
    render_page_nav("viewdata")
    st.stop()

# ── Render all tabs ───────────────────────────────────────────────────────────
_all_tabs = st.tabs(_all_tab_labels)

with _all_tabs[0]:
    _render_minor()

for _i, (_tab, _entry) in enumerate(zip(_all_tabs[1:], _all_files)):
    with _tab:
        _render_data(_entry, f"file_{_i}")

render_page_nav("viewdata")
