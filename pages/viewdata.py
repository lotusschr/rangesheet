"""View Data page — Minor dashboard tab + one tab per uploaded file."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
import pandas as pd
from utils.shared import *
from utils.shared import _detect_large_file_params


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


@st.cache_data(show_spinner="Loading HDET (first load only, ~3 min)…")
def _build_hdet_mini(path: str, _mtime: int = 0) -> pd.DataFrame:
    """Scan HDET once; return compact categorical DataFrame with key columns.

    Replaces the cascade scan + cluster summary + pivot scan.
    After first load everything runs in-memory (instant).
    """
    _sep, _enc, _hr = _detect_large_file_params(path)
    _WANT = {
        "ClusterName":        ["ClusterName","Cluster_Name","cluster_name","Cluster Name"],
        "store_Format":       ["store_Format","Store_Format","StoreFormat","STORE_FORMAT",
                               "Store Format","store format","Format"],
        "Div Code&Desc":      ["Div Code&Desc","Div Code & Desc","DivCode&Desc"],
        "Display Group":      ["Display Group","DG","DG_CODE","Display_Group"],
        "DG_Desc":            ["Display group desc","Display Group Desc","DisplayGroupDesc",
                               "DG_Desc","DGDesc","dg_desc","Display_Group_Desc"],
        "PG_Store_Number":    ["PG_Store_Number","store_no","StoreNo","Store_No",
                               "store_number","Store_Number"],
        "Name":               ["Name","FP_Name","FP Name","FPName","pog_name","POGName"],
        "ID":                 ["ID","id","Barcode","barcode","TPNA"],
        "ProductDescription": ["ProductDescription","Product Description",
                               "item_name","Item Name","Description"],
        "ForecastSales":      ["ForecastSales","forecast_new_item_sales","Avg_unit",
                               "avg_unit","Value","value"],
        "_hdet_tsa":          ["Total Store Apply","TotalStoreApply","total_store_apply",
                               "Total_Store_Apply","store_apply"],
    }
    _first = next(iter(pd.read_csv(
        path, sep=_sep, encoding=_enc, skiprows=_hr, header=0,
        chunksize=1000, dtype=str, low_memory=False, on_bad_lines="skip"
    )))
    _low = {str(c).strip().lower(): c for c in _first.columns}
    _col_map: dict = {}
    for _std, _cands in _WANT.items():
        for _cand in _cands:
            if _cand in _first.columns:
                _col_map[_std] = _cand; break
            if _cand.lower() in _low:
                _col_map[_std] = _low[_cand.lower()]; break
    _use = list(set(_col_map.values()))
    _ren = {v: k for k, v in _col_map.items()}
    _cat = [k for k in _col_map if k not in ("ForecastSales", "_hdet_tsa")]
    _chunks = []
    for _ck in pd.read_csv(
        path, sep=_sep, encoding=_enc, skiprows=_hr, header=0,
        chunksize=LARGE_FILE_CHUNK_SIZE, usecols=_use,
        dtype=str, low_memory=False, on_bad_lines="skip"
    ):
        _ck = _ck.rename(columns=_ren)
        for _col in _cat:
            if _col in _ck.columns:
                _ck[_col] = _ck[_col].fillna("").str.strip().astype("category")
        if "ForecastSales" in _ck.columns:
            _ck["ForecastSales"] = (pd.to_numeric(_ck["ForecastSales"], errors="coerce")
                                    .fillna(0.0).astype("float32"))
        if "_hdet_tsa" in _ck.columns:
            _ck["_hdet_tsa"] = (pd.to_numeric(_ck["_hdet_tsa"], errors="coerce")
                                .fillna(0).astype("int32"))
        _chunks.append(_ck)
    if not _chunks:
        return pd.DataFrame(columns=list(_WANT.keys()))
    _out = pd.concat(_chunks, ignore_index=True)
    for _col in _out.select_dtypes("category").columns:
        _out[_col] = _out[_col].cat.remove_unused_categories()
    return _out


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

# A5 and HDET are source files used by the Minor dashboard and downstream
# calculations.  Keep them available in session state, but do not expose their
# raw-data tabs in the View Data navigation.
def _show_raw_file_tab(entry: dict) -> bool:
    _base_name = os.path.splitext(str(entry.get("name", "")))[0].strip().lower()
    return not (_base_name.startswith("a5") or _base_name.startswith("hdet"))


_visible_files = [_entry for _entry in _all_files if _show_raw_file_tab(_entry)]

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


def _build_pog_cluster_html(pvt, cl_sel, max_height="calc(100vh - 250px)"):
    _TH = "#f0f2f6"; _TC = "#2BBFA4"; _TC2 = "#1a9e8b"; _B = "#dee2e6"
    _COL_W = {"DG_CODE": 74, "ID": 108, "ProductDescription": 230}
    _row_labels = [c for c in ["DG_CODE", "ID", "ProductDescription"] if c in pvt.columns]
    _pog_cols   = [c for c in pvt.columns if c not in set(_row_labels)]
    _N          = len(_pog_cols)
    _has_desc   = "ProductDescription" in pvt.columns
    _left_px = {}; _acc = 0
    for _cn in ["DG_CODE", "ID", "ProductDescription"]:
        if _cn in pvt.columns:
            _left_px[_cn] = _acc; _acc += _COL_W[_cn]

    def _stkH(col, top="0px"):
        _w = _COL_W.get(col, 100); _l = _left_px.get(col, 0)
        return (f"position:sticky;left:{_l}px;top:{top};z-index:5;"
                f"background:{_TH};color:#374151;padding:4px 8px;"
                f"border:1px solid {_B};font-weight:700;"
                f"text-align:left;white-space:nowrap;min-width:{_w}px;")

    def _stkB(col, bg):
        _w = _COL_W.get(col, 100); _l = _left_px.get(col, 0)
        return (f"position:sticky;left:{_l}px;z-index:1;"
                f"background:{bg};padding:3px 8px;"
                f"border:1px solid {_B};white-space:nowrap;min-width:{_w}px;")

    _cl_label = str(cl_sel) if cl_sel else "POG Cluster MOD fixture"
    # Cluster name: sticky left at the first POG column position so it's always visible
    _clS = (f"position:sticky;left:{_acc}px;top:0;z-index:4;"
            f"background:{_TC};color:#fff;padding:4px 12px;"
            f"border:1px solid #1a8a74;font-weight:700;text-align:left;white-space:nowrap;")
    _grS = (f"position:sticky;top:0;z-index:3;"
            f"background:{_TC};color:#fff;padding:4px 8px;"
            f"border:1px solid #1a8a74;font-weight:700;text-align:center;")
    _pgS = (f"position:sticky;top:33px;z-index:3;"
            f"background:{_TC2};color:#fff;padding:3px 6px;"
            f"border:1px solid #1a8a74;font-weight:600;text-align:center;"
            f"font-size:0.70rem;min-width:120px;white-space:normal;word-break:break-word;")

    _ht = [
        "<style>"
        ".pog-tbl-wrap::-webkit-scrollbar{height:12px;width:8px}"
        ".pog-tbl-wrap::-webkit-scrollbar-track{background:#E8E3DC;border-radius:6px}"
        ".pog-tbl-wrap::-webkit-scrollbar-thumb{background:#2BBFA4;border-radius:6px;border:2px solid #E8E3DC}"
        ".pog-tbl-wrap::-webkit-scrollbar-thumb:hover{background:#1a9e8b}"
        ".pog-tbl-wrap::-webkit-scrollbar-corner{background:#E8E3DC}"
        "</style>"
        f"<div class='pog-tbl-wrap' style='overflow-x:auto;overflow-y:auto;"
        f"max-height:{max_height};font-size:0.78rem;"
        f"scrollbar-width:auto;scrollbar-color:#2BBFA4 #E8E3DC;'>",
        "<table style='border-collapse:collapse;'>",
        "<thead style='position:sticky;top:0;z-index:4;'><tr>",
    ]
    for _lbl in [c for c in ["DG_CODE", "ID"] if c in pvt.columns]:
        _disp = "Item no." if _lbl == "ID" else _lbl
        _ht.append(f"<th rowspan='2' style='{_stkH(_lbl)}'>{_disp}</th>")
    if _has_desc:
        _ht.append(f"<th colspan='1' style='{_grS}'>POG CLUSTER</th>")
    if _N:
        _ht.append(f"<th colspan='{_N}' style='{_clS}'>{_cl_label}</th>")
    _ht.append("</tr><tr>")
    if _has_desc:
        _dh = _stkH("ProductDescription", "33px")
        _ht.append(f"<th style='{_dh}'>ProductDescription</th>")
    for _pc in _pog_cols:
        _spc = str(_pc).replace("<", "&lt;").replace(">", "&gt;")
        _ht.append(f"<th title='{_spc}' style='{_pgS}'>{_spc}</th>")
    _ht.append("</tr></thead><tbody>")

    _prev_dg4 = object()
    for _ri, _row in pvt.iterrows():
        _rbg = "#ffffff" if _ri % 2 == 0 else "#f7f8fb"
        _ht.append(f"<tr style='background:{_rbg};'>")
        if "DG_CODE" in pvt.columns:
            _dv4 = str(_row["DG_CODE"]); _bs = _stkB("DG_CODE", _rbg)
            _is_first = _dv4 != _prev_dg4; _prev_dg4 = _dv4
            _dc = "color:#111827;font-weight:600;" if _is_first else "color:#9ca3af;"
            _ht.append(f"<td style='{_bs}{_dc}'>{_dv4}</td>")
        if "ID" in pvt.columns:
            _id_v = str(_row["ID"]).strip()
            _id_val = _id_v.zfill(9) if _id_v not in ("", "nan", "None") else ""
            _ib = _stkB("ID", _rbg)
            _ht.append(f"<td style='{_ib}color:#374151;'>{_id_val}</td>")
        if _has_desc:
            _dsc = str(_row["ProductDescription"]).replace("<", "&lt;").replace(">", "&gt;")
            _db = _stkB("ProductDescription", _rbg)
            _ht.append(f"<td style='{_db}color:#374151;"
                       f"max-width:230px;overflow:hidden;text-overflow:ellipsis;'>{_dsc}</td>")
        for _pc in _pog_cols:
            _v = _row.get(_pc, float("nan"))
            _ht.append(f"<td style='color:#111827;padding:3px 8px;"
                       f"border:1px solid {_B};text-align:right;'>"
                       f"{'%.2f' % _v if pd.notna(_v) and _v != 0 else ''}</td>")
        _ht.append("</tr>")
    _ht.append("</tbody></table></div>")
    return "".join(_ht)


@st.dialog("POG Cluster Table", width="large")
def _pog_fullscreen_dialog():
    st.markdown("""
    <style>
    div[data-testid="stDialog"] > div > div[role="dialog"] {
        width:  100vw !important; max-width:  100vw !important;
        height: 100vh !important; max-height: 100vh !important;
        margin: 0 !important; border-radius: 0 !important;
        top: 0 !important; left: 0 !important; transform: none !important;
        padding: 14px 18px !important;
    }
    div[data-testid="stDialog"] > div {
        width: 100vw !important; height: 100vh !important;
        align-items: flex-start !important;
    }
    </style>""", unsafe_allow_html=True)
    _lbl       = st.session_state.get("_pog_fs_label", "")
    _pvt_fs    = st.session_state.get("_pog_fs_pvt")
    _cl_sel_fs = st.session_state.get("_pog_fs_cl_sel")
    _ready_html = st.session_state.get("_pog_fs_html", "")
    if _lbl:
        st.caption(_lbl)
    _id_srch_fs = st.text_input(
        "", placeholder="🔎  Search by Item no. (ID)…",
        label_visibility="collapsed", key="minor_id_srch_fs"
    )
    if _pvt_fs is not None:
        if _id_srch_fs and "ID" in _pvt_fs.columns:
            _pvt_fs = _pvt_fs.copy()
            _si = _pvt_fs["ID"].astype(str).str.strip().str.zfill(9)
            _sq = (str(_id_srch_fs).strip().zfill(9)
                   if str(_id_srch_fs).strip().isdigit()
                   else str(_id_srch_fs).strip())
            _im = _si.str.contains(_sq, case=False, na=False)
            _pvt_fs = pd.concat([_pvt_fs[_im], _pvt_fs[~_im]]).reset_index(drop=True)
            _ready_html = _build_pog_cluster_html(
                _pvt_fs, _cl_sel_fs, max_height="calc(100vh - 78px)"
            )
        elif not _ready_html:
            _ready_html = _build_pog_cluster_html(
                _pvt_fs, _cl_sel_fs, max_height="calc(100vh - 78px)"
            )
        st.markdown(_ready_html, unsafe_allow_html=True)


def _render_pog_cluster_panel(pvt, cl_sel, label: str, table_html: str):
    """Render the table controls in a fragment so Full screen opens quickly."""
    _exp_c, _btn_c = st.columns([8, 1])
    with _exp_c:
        with st.expander(label, expanded=True):
            st.markdown(table_html, unsafe_allow_html=True)
    with _btn_c:
        if st.button(
            "⛶", key="minor_fs_btn", help="Full screen", use_container_width=True
        ):
            # Reuse the table and its already-rendered HTML.  This avoids rebuilding
            # thousands of cells once more just to open the dialog.
            st.session_state["_pog_fs_pvt"]    = pvt
            st.session_state["_pog_fs_cl_sel"] = cl_sel
            st.session_state["_pog_fs_label"]  = label
            st.session_state["_pog_fs_html"]   = table_html.replace(
                "max-height:calc(100vh - 250px)",
                "max-height:calc(100vh - 78px)",
                1,
            )
            _pog_fullscreen_dialog()


if hasattr(st, "fragment"):
    _render_pog_cluster_panel = st.fragment(_render_pog_cluster_panel)


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
    _fmt_col   = _find_col(_df, "store_Format", "Store_Format", "store_format",
                           "STORE_FORMAT", "StoreFormat", "Store Format",
                           "store format", "Format")
    _div_col   = _find_col(_df, "Div Code&Desc", "Div Code & Desc", "DivCode&Desc", "DivCode")
    _dg_col    = _find_col(_df, "DG", "DG_CODE", "dg_code")
    _cl_col    = _find_col(_df, "ClusterName", "cluster_name", "Cluster_Name",
                           "Cluster (Planogram name)", "Cluster", "FP_Name", "pog_name")
    _store_col = _find_col(_df, "store_no", "StoreNo", "store_id", "Store_No",
                           "PG_Store_Number", "pg_store_number",
                           "store_number", "StoreNumber", "storenum", "Store Number")
    _pog_col   = _find_col(_df, "FP_Name", "pog_name", "PogName", "Cluster (Planogram name)")
    _id_col    = _find_col(_df, "ID", "id", "Barcode", "barcode", "TPNA")
    _desc_col  = _find_col(_df, "Item Name", "item_name", "ProductDescription", "Description")
    _disp_col  = _find_col(_df, "displaygroup", "Display Group", "DisplayGroup",
                           "Display_Group", "displayGroup")

    # ── DIAGNOSTIC — expand to confirm column names, then remove this block ───
    if False:
        st.write("**All columns in _df:**", list(_df.columns))
        st.write(f"fmt=`{_fmt_col}` | div=`{_div_col}` | dg=`{_dg_col}` "
                 f"| cl=`{_cl_col}` | disp=`{_disp_col}` | store=`{_store_col}`")
        st.write("**store cols in _df:**",
                 [c for c in _df.columns if "store" in str(c).lower()])
        _pog_cl_check = _find_col(_df, "POG_Cluster", "pog_cluster")
        _pog_nm_check = _find_col(_df, "planogramname", "POGName", "pog_name")
        st.write(f"POG_Cluster col in _df: `{_pog_cl_check}` | "
                 f"planogramname col in _df: `{_pog_nm_check}`")

    # ── HDET mini-table (one scan; everything else is instant in-memory) ────────
    _RF_HDET_PATH = "rawfiles_hdet_path"
    if _RF_HDET_PATH not in st.session_state:
        _rf_hdet_path = None
        for _rfm in load_admin_manifest():
            _rfp = os.path.join(BASE_DIR, "uploads", _rfm["name"])
            if "hdet" in _rfm["name"].lower() and os.path.exists(_rfp) and is_large_file(_rfp):
                _rf_hdet_path = _rfp
                break
        st.session_state[_RF_HDET_PATH] = _rf_hdet_path
    _hdet_path = st.session_state.get(_RF_HDET_PATH)
    _mini = None
    _dg_desc_map: dict = {}
    if _hdet_path:
        _h_mtime = int(os.path.getmtime(_hdet_path))
        _mini = _build_hdet_mini(_hdet_path, _h_mtime)
        if _mini is not None and "DG_Desc" in _mini.columns and "Display Group" in _mini.columns:
            _dg_desc_map = (
                _mini[["Display Group", "DG_Desc"]]
                .drop_duplicates()
                .dropna(subset=["Display Group"])
                .set_index("Display Group")["DG_Desc"]
                .astype(str)
                .to_dict()
            )
    _mini_fmt_col = _find_col(_mini, "store_Format", "Store_Format", "store_format",
                              "STORE_FORMAT", "StoreFormat", "Store Format",
                              "store format", "Format") if _mini is not None else None
    _mini_div_col = _find_col(_mini, "Div Code&Desc", "Div Code & Desc",
                              "DivCode&Desc", "DivCode") if _mini is not None else None
    _mini_dg_col = _find_col(_mini, "Display Group", "DG", "DG_CODE",
                             "Display_Group", "displayGroup") if _mini is not None else None
    _mini_cl_col = _find_col(_mini, "ClusterName", "Cluster_Name", "cluster_name",
                             "Cluster Name") if _mini is not None else None

    # ── Migrate old multiselect lists → single value ──────────────────────────
    for _ss in ("minor_fmt", "minor_div", "minor_dg", "minor_cl"):
        _v = st.session_state.get(_ss)
        if isinstance(_v, list):
            st.session_state[_ss] = _v[0] if _v else None

    _cur_fmt = st.session_state.get("minor_fmt")
    _cur_div = st.session_state.get("minor_div")
    _cur_dg  = st.session_state.get("minor_dg")
    _cur_cl  = st.session_state.get("minor_cl")

    def _cascade(skip: str) -> list:
        """Valid options for `skip` given all other current single-select filters."""
        if _mini is None:
            return []
        _f = _mini
        if skip != "fmt" and _cur_fmt and _mini_fmt_col in _f.columns:
            _f = _f[_f[_mini_fmt_col].astype(str).str.strip() == str(_cur_fmt).strip()]
        if skip != "div" and _cur_div and _mini_div_col in _f.columns:
            _f = _f[_f[_mini_div_col].astype(str).str.strip() == str(_cur_div).strip()]
        if skip != "dg" and _cur_dg and _mini_dg_col in _f.columns:
            _f = _f[_f[_mini_dg_col].astype(str).str.strip() == str(_cur_dg).strip()]
        if skip != "cl" and _cur_cl and _mini_cl_col in _f.columns:
            _f = _f[_f[_mini_cl_col].astype(str).str.strip() == str(_cur_cl).strip()]
        _col = {"fmt": _mini_fmt_col, "div": _mini_div_col,
                "dg": _mini_dg_col, "cl": _mini_cl_col}[skip]
        if _col not in _f.columns:
            return []
        return sorted(v for v in _f[_col].dropna().astype(str).str.strip().unique()
                      if v not in ("", "nan", "None", "none", "null"))

    def _current_options():
        _a5_sfmt_local = st.session_state.get("_a5_sfmt_opts", [])
        return (
            sorted(set(_cascade("fmt")) | set(_a5_sfmt_local),
                   key=lambda x: (x not in _a5_sfmt_local, x)),
            _cascade("div"),
            _cascade("dg"),
            _cascade("cl"),
        )

    # Drop stale values and auto-select when a cascaded dropdown has only one
    # possible value. This keeps filters visibly connected without extra clicks.
    for _ in range(4):
        _fmt_opts, _div_opts, _dg_opts, _cl_opts = _current_options()
        _changed = False
        for _ss, _opts in [("minor_fmt", _fmt_opts), ("minor_div", _div_opts),
                           ("minor_dg",  _dg_opts),  ("minor_cl",  _cl_opts)]:
            _cur = st.session_state.get(_ss)
            if _cur not in (None, *_opts):
                st.session_state[_ss] = None
                _cur = None
                _changed = True
            if _cur is None and len(_opts) == 1:
                st.session_state[_ss] = _opts[0]
                _changed = True
        _cur_fmt = st.session_state.get("minor_fmt")
        _cur_div = st.session_state.get("minor_div")
        _cur_dg  = st.session_state.get("minor_dg")
        _cur_cl  = st.session_state.get("minor_cl")
        if not _changed:
            break
    _fmt_opts, _div_opts, _dg_opts, _cl_opts = _current_options()

    # ── Filter bar (4 single-select dropdowns) ────────────────────────────────
    _fmt_sel = st.session_state.get("minor_fmt")
    _fc2, _fc3, _fc4 = st.columns(3)
    with _fc2:
        _div_sel = st.selectbox("DIV CODE&DESC", [None] + _div_opts, key="minor_div",
                                format_func=lambda x: "All divisions" if x is None else x)
    with _fc3:
        _dg_sel = st.selectbox("DG", [None] + _dg_opts, key="minor_dg",
                               format_func=lambda x: "All DGs" if x is None else
                               (f"{x}  {_dg_desc_map[x]}" if x in _dg_desc_map and _dg_desc_map[x] not in ("", "nan", "None") else x))
    with _fc4:
        _cl_sel = st.selectbox("CLUSTERNAME", [None] + _cl_opts, key="minor_cl",
                               format_func=lambda x: "All clusters" if x is None else x)

    # _fdf: filtered small DB (for any downstream use)
    _fdf = _df.copy()
    if _fmt_sel and _fmt_col:  _fdf = _fdf[_fdf[_fmt_col].astype(str) == _fmt_sel]
    if _div_sel and _div_col:  _fdf = _fdf[_fdf[_div_col].astype(str) == _div_sel]
    if _dg_sel  and _dg_col:   _fdf = _fdf[_fdf[_dg_col].astype(str)  == _dg_sel]
    if _cl_sel  and _cl_col:   _fdf = _fdf[_fdf[_cl_col].astype(str)  == _cl_sel]

    def _filter_eq(_frame: pd.DataFrame, _col: str | None, _val):
        if not _val or not _col or _col not in _frame.columns:
            return _frame
        return _frame[_frame[_col].astype(str).str.strip() == str(_val).strip()]

    # ── Cluster summary from mini-table (instant pandas, no re-scan) ──────────
    _summary_df = None
    if _mini is not None:
        # _null_s = {"", "0", "0.0"}
        _null_s = {"", "nan", "None", "NaN", "none", "null"}
        # Base filter: format + div + cluster (NO DG — keeps StoreCount stable)
        _mf = _mini
        _mf = _filter_eq(_mf, _mini_fmt_col, _fmt_sel)
        _mf = _filter_eq(_mf, _mini_div_col, _div_sel)
        _mf = _filter_eq(_mf, _mini_cl_col, _cl_sel)
        # TotalStoreApply: add DG filter (counts DG entries on planogram)
        _mf_dg = _filter_eq(_mf, _mini_dg_col, _dg_sel)
        # Left cluster table: Format + Div + DG only — NO ClusterName filter so all
        # clusters connected to the selected DG are shown (mirrors PBI behaviour)
        _mf_for_clusters = _mini
        _mf_for_clusters = _filter_eq(_mf_for_clusters, _mini_fmt_col, _fmt_sel)
        _mf_for_clusters = _filter_eq(_mf_for_clusters, _mini_div_col, _div_sel)
        _mf_for_clusters = _filter_eq(_mf_for_clusters, _mini_dg_col, _dg_sel)

        # StoreCount = DISTINCTCOUNT(POG_Store[store_no])
        # PBI source: A5_2_POG_FP_HDET_LIVE_*.csv  →  store_no + Property_Store_Cluster + POGName
        _null_sv      = {"", "nan", "NaN", "None", "none", "null", "NULL", "N/A", "n/a"}
        _sc_store_col = None
        _sc_cl_col    = None
        _sc_src_df    = None
        _sc_fmt_col   = None
        _sc_src_label = "none"
        _sc_base      = pd.DataFrame()

        # ── Direct scan: find a CSV with store_no + Property_Store_Cluster + POGName ──
        # This is the exact signature of the A5 POG_Store CSV that PBI uses.
        # We read only the 4 needed columns (usecols) so it's fast even for large CSVs.
        _A5_S  = ["store_no", "StoreNo", "Store_No", "store_number", "StoreNumber"]
        # POG_Cluster = "A"/"N_G_A"/... (the PBI cluster column); planogramname = PBI's POGName
        _A5_C  = ["POG_Cluster", "pog_cluster", "Property_Store_Cluster", "property_store_cluster"]
        _A5_P  = ["planogramname", "POGName", "pog_name", "POG_Name", "POG Name",
                  "planogram_name", "Planogramname"]
        _A5_F  = ["store_Format", "store_format", "StoreFormat", "store format"]

        for _am in load_admin_manifest():
            _ap = os.path.join(BASE_DIR, "uploads", _am["name"])
            if not os.path.exists(_ap):
                continue
            if os.path.splitext(_am["name"])[-1].lower() not in (".csv", ".txt"):
                continue
            try:
                _sep2, _enc2, _hdr2 = _detect_large_file_params(_ap)
                _peek = pd.read_csv(_ap, sep=_sep2, encoding=_enc2,
                                    skiprows=_hdr2, nrows=0, dtype=str, low_memory=False)
                _lmap = {str(c).strip().lower(): str(c).strip() for c in _peek.columns}
                def _rc(_cands, _m=_lmap):
                    for _cc in _cands:
                        if _cc.lower() in _m:
                            return _m[_cc.lower()]
                    return None
                _a5_s = _rc(_A5_S)
                _a5_c = _rc(_A5_C)
                _a5_p = _rc(_A5_P)
                if not (_a5_s and _a5_c and _a5_p):
                    continue
                _a5_f  = _rc(_A5_F)
                _use   = list({_a5_s, _a5_c, _a5_p} | ({_a5_f} if _a5_f else set()))
                _a5_df = pd.read_csv(_ap, sep=_sep2, encoding=_enc2, skiprows=_hdr2,
                                     usecols=_use, dtype=str, low_memory=False,
                                     on_bad_lines="skip")
                _a5_df.columns = [str(c).strip() for c in _a5_df.columns]
                _sc_store_col = _a5_s
                _sc_cl_col    = _a5_c
                _sc_src_df    = _a5_df
                _sc_fmt_col   = _a5_f
                _sc_src_label = f"A5:{_am['name']} store={_a5_s}, cluster={_a5_c}"
                break
            except Exception:
                pass

        # ── A5 store_Format: cache options & apply format filter ─────────────────
        if _sc_fmt_col and _sc_src_df is not None and _sc_fmt_col in _sc_src_df.columns:
            _a5_sfmt_all = sorted(
                v for v in _sc_src_df[_sc_fmt_col].dropna().astype(str).str.strip().unique()
                if v.lower() not in {"", "nan", "none", "null"}
            )
            if _a5_sfmt_all != st.session_state.get("_a5_sfmt_opts", []):
                st.session_state["_a5_sfmt_opts"] = _a5_sfmt_all
                st.rerun()
            if _fmt_sel:
                _sc_src_df = _sc_src_df[
                    _sc_src_df[_sc_fmt_col].astype(str).str.strip() == _fmt_sel
                ]

        # ── JOIN: HDET[Name] = A5[planogramname] → DISTINCTCOUNT(store_no) per ClusterName ──
        # This matches PBI's model exactly:
        #   StoreCount      = DISTINCTCOUNT(POG_Store[store_no])  per Item_POG[ClusterName]
        #   Count of POGName = COUNT(POG_Store[planogramname])     per Item_POG[ClusterName]
        _pog_nm_col = (_find_col(_sc_src_df, "planogramname", "POGName", "pog_name",
                                 "POG_Name", "planogram_name", "Planogramname")
                       if _sc_src_df is not None else None)
        _sc_base    = pd.DataFrame()
        _pc         = pd.Series(dtype=int)
        _pc_tot     = 0

        if (_sc_store_col and _pog_nm_col
                and "Name" in _mf_for_clusters.columns and "ClusterName" in _mf_for_clusters.columns):
            # Step 1: unique (ClusterName, POG_Name) — Format/Div/DG filtered but NO ClusterName
            # filter so left table shows ALL clusters connected to the selected DG
            _hdet_pogs = _mf_for_clusters[["ClusterName", "Name"]].copy()
            _hdet_pogs["ClusterName"] = _hdet_pogs["ClusterName"].astype(str).str.strip()
            _hdet_pogs["Name"]        = _hdet_pogs["Name"].astype(str).str.strip()
            _hdet_pogs = _hdet_pogs[
                (_hdet_pogs["ClusterName"] != "") & (_hdet_pogs["ClusterName"] != "nan")
                & (_hdet_pogs["Name"] != "")      & (_hdet_pogs["Name"] != "nan")
            ].drop_duplicates()

            # Step 2: A5 (store_no, planogramname) — one row per store × POG assignment
            _a5_sub = _sc_src_df[[_pog_nm_col, _sc_store_col]].copy()
            _a5_sub[_pog_nm_col]   = _a5_sub[_pog_nm_col].astype(str).str.strip()
            _a5_sub[_sc_store_col] = _a5_sub[_sc_store_col].astype(str).str.strip()
            _a5_sub = _a5_sub[
                ~_a5_sub[_pog_nm_col].isin(_null_sv)
                & ~_a5_sub[_sc_store_col].isin(_null_sv)
            ]

            # Step 3: JOIN → each row = (ClusterName, Name, planogramname, store_no)
            _joined = _hdet_pogs.merge(
                _a5_sub,
                left_on="Name",
                right_on=_pog_nm_col,
                how="left"
            )

            _sc     = _joined.groupby("ClusterName")[_sc_store_col].nunique()
            _sc.index.name = None
            _pc     = (_joined[_joined[_pog_nm_col].notna()]
                       .groupby("ClusterName")[_pog_nm_col].nunique())
            _pc.index.name = None
            # Total = DISTINCTCOUNT(planogramname) where Name exists in both HDET and A5
            _null_low    = {v.lower() for v in _null_sv}
            _hdet_nm_all = _mf_for_clusters["Name"].astype(str).str.strip()
            _hdet_nm_set = set(_hdet_nm_all[~_hdet_nm_all.str.lower().isin(_null_low)].unique())
            _a5_nm       = _sc_src_df[_pog_nm_col].astype(str).str.strip()
            _a5_nm_set   = set(_a5_nm[~_a5_nm.str.lower().isin(_null_low)].unique())
            _pc_tot      = len(_hdet_nm_set & _a5_nm_set)
            _sc_base = _joined
            _sc_src_label += f" | JOIN on planogramname → {len(_sc)} clusters"
        else:
            # Fallback: HDET-only (no A5 found or missing planogramname col)
            _sc_src_label = "HDET fallback"
            _sc = (_mf[~_mf["PG_Store_Number"].isin(_null_s)]
                   .groupby("ClusterName")["PG_Store_Number"].nunique()
                   if "PG_Store_Number" in _mf.columns else pd.Series(dtype=int))
            _pc = (_mf[_mf["Name"] != ""]
                   .groupby("ClusterName")["Name"].nunique()
                   if "Name" in _mf.columns else pd.Series(dtype=int))
            _pc_tot = (int(_mf[_mf["Name"] != ""]["Name"].nunique())
                       if "Name" in _mf.columns else 0)

        if False:
            st.caption(f"Source: **{_sc_src_label}** | clusters: {len(_sc)}")
        if _cl_sel:
            # ClusterName selected: use _mf_dg (has ClusterName filter) so horizontal
            # pivot stays focused on the selected cluster — original behaviour
            _ic_cols = [c for c in ["ClusterName","Display Group","ID","ProductDescription"]
                        if c in _mf_dg.columns]
            _ic = (_mf_dg[_ic_cols].drop_duplicates()
                   .groupby("ClusterName").size()
                   if len(_ic_cols) > 1 else pd.Series(dtype=int))
            _ta = (_mf_dg[_mf_dg["Name"] != ""]
                   .groupby("ClusterName")["Name"].nunique()
                   if "Name" in _mf_dg.columns else pd.Series(dtype=int))
        else:
            # No ClusterName filter: align with left table values so horizontal pivot
            # matches PBI before filter is applied
            # TotalStoreApply = Count of POGName (same as _pc, already confirmed correct)
            _ta = _pc.copy() if not _pc.empty else pd.Series(dtype=int)
            # TotalItem = DISTINCTCOUNT(ID) per cluster, null IDs excluded — matches PBI
            if "ID" in _mf_for_clusters.columns:
                _ic_src = _mf_for_clusters[~_mf_for_clusters["ID"].isin(_null_sv)]
                _ic = _ic_src.groupby("ClusterName")["ID"].nunique()
                _ic.index.name = None
            else:
                _ic = pd.Series(dtype=int)

        _cls_list = sorted(set(_sc.index) | set(_pc.index) | set(_ta.index))
        _summary_df = pd.DataFrame({
            "ClusterName":      _cls_list,
            "StoreCount":       [int(_sc.get(c, 0)) for c in _cls_list],
            "Count of POGName": [int(_pc.get(c, 0)) for c in _cls_list],
            "ItemCount":        [int(_ic.get(c, 0)) for c in _cls_list],
            "TotalStoreApply":  [int(_ta.get(c, 0)) for c in _cls_list],
        })
        _summary_df = (_summary_df.sort_values("StoreCount", ascending=False)
                       .reset_index(drop=True))
        _ic_tot_cols = [c for c in ["Display Group","ID","ProductDescription"]
                        if c in _mf_dg.columns]
        _tot_row = pd.DataFrame([{
            "ClusterName":      "Total",
            # "StoreCount":       int(_mf[~_mf["PG_Store_Number"].isin(_null_s)]
            #                         ["PG_Store_Number"].nunique())
            #                     if "PG_Store_Number" in _mf.columns else 0,
            # Total StoreCount: when DG filter active use joined result (filter-aware),
            # otherwise use full A5 distinct count (matches PBI unfiltered grand total = 491)
            "StoreCount":       (int(_joined[_sc_store_col].nunique())
                                 if _dg_sel and not _sc_base.empty and _sc_store_col
                                 else (int(_sc_src_df[_sc_store_col].nunique())
                                       if _sc_src_df is not None and _sc_store_col
                                       else (int(_mf[~_mf["PG_Store_Number"].isin(_null_s)]
                                                 ["PG_Store_Number"].nunique())
                                             if "PG_Store_Number" in _mf.columns else 0))),
            "Count of POGName": _pc_tot,
            "ItemCount":        int(_mf_dg[_ic_tot_cols].drop_duplicates().shape[0])
                                if _ic_tot_cols else 0,
            "TotalStoreApply":  int(_mf_dg[_mf_dg["Name"] != ""]["Name"].nunique())
                                if "Name" in _mf_dg.columns else 0,
        }])
        _summary_df = pd.concat([_summary_df, _tot_row], ignore_index=True)

    # ── Horizontal pivot: Total Store Apply + Total Item per cluster ───────────
    if (_summary_df is not None and not _summary_df.empty
            and (_fmt_sel or _div_sel or _dg_sel or _cl_sel)):
        _is_tot2   = _summary_df["ClusterName"] == "Total"
        _data_p    = _summary_df[~_is_tot2]
        # Only show clusters that have actual TotalStoreApply or ItemCount data
        _data_p    = _data_p[(_data_p["TotalStoreApply"] > 0) | (_data_p["ItemCount"] > 0)]
        _tot_p     = _summary_df[_is_tot2]
        _cls_names = list(_data_p["ClusterName"])
        _store_map = dict(zip(_data_p["ClusterName"], _data_p["TotalStoreApply"].astype(int)))
        _item_map  = dict(zip(_data_p["ClusterName"], _data_p["ItemCount"].astype(int)))
        if not _tot_p.empty:
            _cls_names.append("Total")
            _store_map["Total"] = int(_tot_p.iloc[0]["TotalStoreApply"])
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
            _tbl_h = max(60, min(len(_disp_rows3) * 35 + 38, 800))
            st.dataframe(
                _disp_rows3,
                hide_index=True,
                use_container_width=True,
                height=_tbl_h,
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

        if _mini is None:
            st.caption("HDET file not found — upload and pin the HDET file.")
        elif not _cl_sel:
            st.info("Select a **ClusterName** to display the POG_Cluster table.")
        else:
            # Build pivot from mini-table (instant — no additional HDET scan)
            _f = _filter_eq(_mini, _mini_cl_col, _cl_sel).copy()
            _f = _filter_eq(_f, _mini_fmt_col, _fmt_sel)
            _f = _filter_eq(_f, _mini_div_col, _div_sel)
            _f = _filter_eq(_f, _mini_dg_col, _dg_sel)
            _ren = {"Display Group": "DG_CODE", "Name": "POGName", "ForecastSales": "Value"}
            _raw_pvt = _f.rename(columns=_ren)
            _pvt_cols = [c for c in ["DG_CODE","ID","ProductDescription","POGName","Value"]
                         if c in _raw_pvt.columns]
            _raw_pvt = _raw_pvt[_pvt_cols].copy()
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

                # ID search bar — matched row floats to top
                _id_srch = st.text_input(
                    "", placeholder="🔎  Search by Item no. (ID)…",
                    label_visibility="collapsed", key="minor_id_srch"
                )
                if _id_srch and "ID" in _pvt.columns:
                    _si = _pvt["ID"].astype(str).str.strip().str.zfill(9)
                    _sq = (str(_id_srch).strip().zfill(9)
                           if str(_id_srch).strip().isdigit()
                           else str(_id_srch).strip())
                    _im = _si.str.contains(_sq, case=False, na=False)
                    _pvt = pd.concat([_pvt[_im], _pvt[~_im]]).reset_index(drop=True)

                _pog_cols = [c for c in _pvt.columns if c not in _row_labels]
                _N        = len(_pog_cols)
                _right_html  = _build_pog_cluster_html(_pvt, _cl_sel)
                _right_label = f"{len(_pvt):,} items · {_N} POGs"
                _render_pog_cluster_panel(
                    _pvt, _cl_sel, _right_label, _right_html
                )


# ── Tab label list: [Minor] + one per uploaded file ───────────────────────────
_file_labels: list[str] = []
_seen_labels: dict[str, int] = {}
for _f in _visible_files:
    _lbl = _short(_f["name"])
    if _lbl in _seen_labels:
        _seen_labels[_lbl] += 1
        _file_labels.append(f"{_lbl} ({_seen_labels[_lbl]})")
    else:
        _seen_labels[_lbl] = 0
        _file_labels.append(_lbl)

_all_tab_labels = ["Minor"] + _file_labels

# ── No files yet — still show Minor tab ──────────────────────────────────────
if not _visible_files:
    (_t_minor,) = st.tabs(["Minor"])
    with _t_minor:
        _render_minor()
    render_page_nav("viewdata")
    st.stop()

# ── Render all tabs ───────────────────────────────────────────────────────────
_all_tabs = st.tabs(_all_tab_labels)

with _all_tabs[0]:
    _render_minor()

for _i, (_tab, _entry) in enumerate(zip(_all_tabs[1:], _visible_files)):
    with _tab:
        _render_data(_entry, f"file_{_i}")

render_page_nav("viewdata")
