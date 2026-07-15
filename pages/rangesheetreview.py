"""Rangesheet review page — sheet tabs, Range Architecture card, data table."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


import streamlit as st
import pandas as pd
import re as _re
import json as _json
import math
import difflib
import html as _html
import hashlib as _hashlib
import io
from datetime import datetime
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar, render_page_nav,
    RS_SHEETS, RS_COL_GROUPS, STATUS_COLORS, FILL_COLORS, COLUMN_LABELS, COLUMN_MAPPING,
    get_fill, df_to_xlsx_bytes, df_to_csv_bytes, add_audit,
    get_shared_db,
    load_large_file_by_dg, get_dg_options, get_dg_index, ensure_hdet_parquet,
    is_large_file, BASE_DIR, load_admin_manifest, read_large_file_head,
    apply_column_mapping, DG_COLUMN_CANDIDATES,
)

inject_css()
init_session_state()
render_sidebar("rangesheetreview")
render_topbar("Rangesheet Review")

selected_files = st.session_state.get("selected_files", [])
_user_upload = st.session_state.get("upload_df")

# Always read the shared admin database — it refreshes for all users whenever
# the admin pins or unpins a file (bump_shared_db invalidates the cache).
_shared_df, _ = get_shared_db()

_RS_NEW_ROWS_DIR = os.path.join(BASE_DIR, "rangesheet_data", ".autosave", "rangesheetreview")


def _rs_new_rows_path(scope: str) -> str:
    os.makedirs(_RS_NEW_ROWS_DIR, exist_ok=True)
    digest = _hashlib.sha1(str(scope).encode("utf-8")).hexdigest()
    return os.path.join(_RS_NEW_ROWS_DIR, f"new_rows_{digest}.json")


def _rs_load_new_rows(scope: str) -> list[dict]:
    path = _rs_new_rows_path(scope)
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = _json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def _rs_save_new_rows(scope: str, rows: list[dict]) -> None:
    path = _rs_new_rows_path(scope)
    tmp = f"{path}.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        _json.dump(rows or [], f, ensure_ascii=False, indent=2, default=str)
    os.replace(tmp, path)


def _rs_edit_state_path(scope: str) -> str:
    os.makedirs(_RS_NEW_ROWS_DIR, exist_ok=True)
    digest = _hashlib.sha1(str(scope).encode("utf-8")).hexdigest()
    return os.path.join(_RS_NEW_ROWS_DIR, f"edit_state_{digest}.json")


def _rs_table_base_path(scope: str) -> str:
    os.makedirs(_RS_NEW_ROWS_DIR, exist_ok=True)
    digest = _hashlib.sha1(str(scope).encode("utf-8")).hexdigest()
    return os.path.join(_RS_NEW_ROWS_DIR, f"table_base_{digest}.pkl")


def _rs_report_packets_path(p: str) -> str:
    os.makedirs(_RS_NEW_ROWS_DIR, exist_ok=True)
    return os.path.join(_RS_NEW_ROWS_DIR, f"report_packets_{p}.pkl")


def _rs_load_report_packets(p: str) -> dict:
    path = _rs_report_packets_path(p)
    if not os.path.exists(path):
        return {}
    try:
        data = pd.read_pickle(path)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _rs_save_report_packets(p: str, packets: dict) -> None:
    path = _rs_report_packets_path(p)
    tmp = f"{path}.tmp"
    pd.to_pickle(packets or {}, tmp)
    os.replace(tmp, path)


def _rs_load_table_base(scope: str) -> dict:
    path = _rs_table_base_path(scope)
    if not os.path.exists(path):
        return {}
    try:
        data = pd.read_pickle(path)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _rs_save_table_base(scope: str, data: dict) -> None:
    path = _rs_table_base_path(scope)
    tmp = f"{path}.tmp"
    pd.to_pickle(data or {}, tmp)
    os.replace(tmp, path)


def _rs_pack_keyed_dict(d: dict) -> list[dict]:
    out = []
    for k, v in (d or {}).items():
        key = list(k) if isinstance(k, tuple) else [str(k)]
        out.append({"key": key, "value": v})
    return out


def _rs_unpack_keyed_dict(items) -> dict:
    out = {}
    if not isinstance(items, list):
        return out
    for item in items:
        if not isinstance(item, dict):
            continue
        key = item.get("key", [])
        if not isinstance(key, list):
            key = [str(key)]
        out[tuple(str(x) for x in key)] = item.get("value", {})
    return out


def _rs_load_edit_state(scope: str) -> dict:
    path = _rs_edit_state_path(scope)
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = _json.load(f)
        if not isinstance(data, dict):
            return {}
        pog_actions = _rs_unpack_keyed_dict(data.get("pog_actions", []))
        pog_edits = _rs_unpack_keyed_dict(data.get("pog_edits", []))
        status_overrides = {
            k: str(v)
            for k, v in _rs_unpack_keyed_dict(data.get("status_overrides", [])).items()
        }
        for rk, status in list(status_overrides.items()):
            if str(status).strip().upper() == "DELETE ALL":
                pog_actions.pop(rk, None)
                pog_edits.pop(rk, None)
        return {
            "pog_actions": pog_actions,
            "pog_edits": pog_edits,
            "avg_u_edits": _rs_unpack_keyed_dict(data.get("avg_u_edits", [])),
            "data_edits": _rs_unpack_keyed_dict(data.get("data_edits", [])),
            "status_overrides": status_overrides,
        }
    except Exception:
        return {}


def _rs_save_edit_state(
    scope: str,
    pog_actions: dict,
    pog_edits: dict,
    avg_u_edits: dict,
    status_overrides: dict,
    data_edits: dict | None = None,
) -> None:
    path = _rs_edit_state_path(scope)
    tmp = f"{path}.tmp"
    pog_actions = dict(pog_actions or {})
    pog_edits = dict(pog_edits or {})
    status_overrides = dict(status_overrides or {})
    for rk, status in list(status_overrides.items()):
        if str(status).strip().upper() == "DELETE ALL":
            pog_actions.pop(rk, None)
            pog_edits.pop(rk, None)
    data = {
        "pog_actions": _rs_pack_keyed_dict(pog_actions),
        "pog_edits": _rs_pack_keyed_dict(pog_edits),
        "avg_u_edits": _rs_pack_keyed_dict(avg_u_edits),
        "data_edits": _rs_pack_keyed_dict(data_edits or {}),
        "status_overrides": _rs_pack_keyed_dict(status_overrides),
    }
    with open(tmp, "w", encoding="utf-8") as f:
        _json.dump(data, f, ensure_ascii=False, separators=(",", ":"), default=str)
    os.replace(tmp, path)

# User's own file selection takes priority; fall back to shared admin database.
if selected_files and _user_upload is not None and len(_user_upload) > 0:
    merged = _user_upload
else:
    merged = _shared_df

# Re-apply column mapping here so that files already in session state (uploaded
# before the mapping was updated) are also correctly renamed without re-upload.
if merged is not None and len(merged) > 0:
    merged = apply_column_mapping(merged)

# Treat both None AND empty DataFrame as "no data"
_no_data = merged is None or (hasattr(merged, "__len__") and len(merged) == 0)

if _no_data:
    st.markdown("""
<div style="background:#fff;border-radius:16px;border:1px solid #E8E3DC;
            padding:60px;text-align:center;">
    <div style="font-size:48px;margin-bottom:16px;">📂</div>
    <div style="font-size:16px;font-weight:600;color:#1A1A1A;margin-bottom:8px;">No data yet</div>
    <div style="font-size:13px;color:#999;">Go to My Files and upload a file first.</div>
</div>
""", unsafe_allow_html=True)
    st.stop()

# Deduplicate column names in the merged df so Arrow / Streamlit never crashes.
def _dedup(df: pd.DataFrame) -> pd.DataFrame:
    seen: dict[str, int] = {}
    cols = []
    for c in df.columns:
        s = str(c)
        if s in seen:
            seen[s] += 1
            cols.append(f"{s}.{seen[s]}")
        else:
            seen[s] = 0
            cols.append(s)
    if cols != list(df.columns):
        df = df.copy()
        df.columns = cols
    return df

merged = _dedup(merged)
all_cols = list(merged.columns)

# ── Column group matching ──────────────────────────────────────────────────────
def _nc(s):  return str(s).lower().strip().replace('\n', ' ').replace('  ', ' ')
def _nca(s): return _re.sub(r'[^a-z0-9]', '', _nc(s))   # letters+digits only

@st.cache_data(show_spinner=False)
def _load_a5_subset(amp: str, mtime: float, dyn_pog_cols_key: tuple):
    """Load only A5 columns/rows needed for the visible planograms."""
    _ext = os.path.splitext(amp)[1].lower()
    _encodings = ("utf-8-sig", "cp874", "latin1")
    _header = None
    _enc_used = None
    _warn = None

    try:
        if _ext in (".csv", ".txt"):
            _last_err = None
            for _enc in _encodings:
                try:
                    _header = pd.read_csv(amp, sep="|", encoding=_enc, nrows=0)
                    _enc_used = _enc
                    break
                except Exception as _ce:
                    _last_err = _ce
            if _header is None and _last_err is not None:
                raise _last_err
        else:
            _header = pd.read_excel(amp, nrows=0)
    except Exception as _e:
        return None, {}, str(_e)

    _cols = list(_header.columns)
    _cluster_c = (
        next((c for c in _cols if _nca(c) == _nca("POG_Cluster")), None)
        or next((c for c in _cols if _nca(c) in ("pogcluster", "clustername")), None)
    )
    _pog_candidates = [
        c for c in _cols
        if c != _cluster_c and (
            _nca(c) in ("planogramname", "pogname", "planogram")
            or "planogram" in _nca(c)
            or ("pog" in _nca(c) and "cluster" not in _nca(c))
        )
    ]
    _mod_c = next((c for c in _cols if _nca(c) == _nca("no_of_mod")), None)
    _fix_c = next((c for c in _cols if _nca(c) == _nca("Fixture_code")), None)
    _rng_c = next((c for c in _cols if _nca(c) == _nca("Range class")), None)
    _usecols = list(dict.fromkeys([c for c in (
        _cluster_c, *_pog_candidates, _mod_c, _fix_c, _rng_c
    ) if c]))

    if not _usecols:
        return None, {}, "A5 loaded but required columns were not found."

    try:
        if _ext in (".csv", ".txt"):
            _a5 = pd.read_csv(amp, sep="|", encoding=_enc_used, usecols=_usecols)
        else:
            _a5 = pd.read_excel(amp, usecols=_usecols)
    except Exception as _e:
        return None, {}, str(_e)

    _target = {
        _nca(v) for v in dyn_pog_cols_key
        if str(v).strip() not in ("", "nan", "None")
    }
    _best_col, _best_hit = (_pog_candidates[0] if _pog_candidates else None), -1
    if _target:
        for _cand in _pog_candidates:
            _vals = set(_a5[_cand].dropna().astype(str).str.strip().map(_nca))
            _hit = len(_vals & _target)
            if _hit > _best_hit:
                _best_col, _best_hit = _cand, _hit
        if _best_col and _best_hit > 0:
            _mask = _a5[_best_col].astype(str).str.strip().map(_nca).isin(_target)
            _a5 = _a5.loc[_mask].copy()

    _meta = {
        "pog": _best_col,
        "cluster": _cluster_c,
        "mod": _mod_c,
        "fixture": _fix_c,
        "range": _rng_c,
    }
    return _a5, _meta, _warn

def _cast_text_cols(df: pd.DataFrame, text_cols: list) -> pd.DataFrame:
    """Cast specified columns to str so TextColumn editors don't crash on int data."""
    for c in text_cols:
        if c in df.columns:
            df[c] = df[c].where(df[c].isna(), df[c].astype(str))
    return df

# Columns that should default to 0 (not blank) when not present in the source file.
_ZERO_DEFAULT_COLS = {
    _nca("AS IS planograms applied"),
    _nca("TO-BE planograms applied"),
    _nca("AS-IS Stores Applied"),
    _nca("TO-Be stores applied"),
}

def _fill_from_db(col_list: list, source_df) -> pd.DataFrame:
    """Build a DataFrame matching col_list by scanning source_df for same-named columns."""
    if source_df is None or len(source_df) == 0:
        return pd.DataFrame(columns=col_list)
    col_map = {col: next((c for c in source_df.columns if _nca(c) == _nca(col)), None)
               for col in col_list}
    matched = [col for col, src in col_map.items() if src is not None]
    if not matched:
        return pd.DataFrame(columns=col_list)
    n = len(source_df)
    result = {col: (source_df[col_map[col]].reset_index(drop=True)
                    if col_map[col]
                    else pd.Series([0] * n) if _nca(col) in _ZERO_DEFAULT_COLS
                    else pd.Series([None] * n))
              for col in col_list}
    df = pd.DataFrame(result)
    # Drop rows where every matched column is null
    df = df[~df[matched].isnull().all(axis=1)].reset_index(drop=True)
    return df

# Signature used to detect when the database (pinned files) changes
_db_sig = (len(merged), tuple(merged.columns.tolist()))

_dcm  = {_nc(c):  c for c in all_cols}   # normalized → actual col name
_dcm2 = {_nca(c): c for c in all_cols}   # aggressively normalized → actual col name

def _fc(key):
    return _dcm.get(_nc(key)) or _dcm2.get(_nca(key))

_CL_AGG = {_nca(k): v for k, v in COLUMN_LABELS.items()}
def _col_label(c):
    return (COLUMN_LABELS.get(_nc(c))
            or COLUMN_LABELS.get(_nc(c).replace('_', ' ').replace('-', ' '))
            or _CL_AGG.get(_nca(c))
            or c)

active_groups = []
for grp in RS_COL_GROUPS:
    matched = [_fc(k) for k in grp["cols"] if _fc(k) is not None]
    if matched:
        active_groups.append({**grp, "matched": matched})
_all_grouped = [c for g in active_groups for c in g["matched"]]
# Non-RS columns are intentionally excluded from Rangesheet Review — they appear on the View Data page instead.

# view_vis_cols stores which cols to SHOW (set), order always follows _all_grouped
if st.session_state.view_vis_cols is not None:
    if not any(c in all_cols for c in st.session_state.view_vis_cols):
        st.session_state.view_vis_cols = None

# ── MINOR / MAJOR — top-level header tabs ─────────────────────────────────────
st.markdown("""
<style>
/* ── MINOR/MAJOR radio — big header tabs ────────────────────────────────── */
div[data-testid="stVerticalBlock"] > div:nth-child(1) div[role="radiogroup"] > label {
    font-size: 20px !important;
    font-weight: 900 !important;
    padding: 14px 44px !important;
    letter-spacing: 0.06em !important;
    min-width: 130px !important;
    text-align: center !important;
    border-bottom-width: 4px !important;
}
div[data-testid="stVerticalBlock"] > div:nth-child(1) div[role="radiogroup"] {
    border-bottom-width: 3px !important;
}

/* ── Sheet tab bar — drag-scroll with thin teal scrollbar ───────────────── */
[data-baseweb="tab-list"] {
    overflow-x: auto !important;
    overflow-y: hidden !important;
    flex-wrap: nowrap !important;
    scrollbar-width: thin !important;
    scrollbar-color: #2BBFA4 #E8E3DC !important;
    padding-bottom: 4px !important;
    gap: 2px !important;
    -webkit-overflow-scrolling: touch !important;
}
[data-baseweb="tab-list"]::-webkit-scrollbar {
    height: 4px !important;
}
[data-baseweb="tab-list"]::-webkit-scrollbar-track {
    background: #F0EBE3 !important;
    border-radius: 2px !important;
}
[data-baseweb="tab-list"]::-webkit-scrollbar-thumb {
    background: #2BBFA4 !important;
    border-radius: 2px !important;
}
/* Hide Streamlit's built-in left/right scroll arrows */
button[data-testid="stTabScrollLeft"],
button[data-testid="stTabScrollRight"] {
    display: none !important;
}

/* ── Individual tab buttons ──────────────────────────────────────────────── */
[data-baseweb="tab"] {
    flex-shrink: 0 !important;
    white-space: nowrap !important;
    border-radius: 8px 8px 0 0 !important;
    background: transparent !important;
    color: #666 !important;
    font-weight: 600 !important;
    padding: 9px 16px !important;
    border: none !important;
    transition: background 0.15s, color 0.15s !important;
}
[data-baseweb="tab"]:hover {
    background: rgba(43, 191, 164, 0.12) !important;
    color: #2BBFA4 !important;
}

/* ── Active tab — teal fill instead of underline ────────────────────────── */
[data-baseweb="tab"][aria-selected="true"] {
    background: #2BBFA4 !important;
    color: #ffffff !important;
}
/* Remove default underline highlight bar */
[data-baseweb="tab-highlight"] {
    display: none !important;
}
/* Bottom border of tab strip */
[data-baseweb="tab-border"] {
    background: #D0CAC2 !important;
    height: 1px !important;
}
</style>
""", unsafe_allow_html=True)

if "vw_live_type" not in st.session_state:
    st.session_state.vw_live_type = "MINOR"

_live_type = st.radio(
    "live_type_sel", ["MINOR", "MAJOR", "Refresh"],
    horizontal=True, label_visibility="collapsed",
    key="vw_live_type",
)

# ── MAJOR — Department / Section / Class tabs ──────────────────────────────────
if _live_type == "MAJOR":
    _dept_col  = next((c for c in all_cols if "department" in _nc(c)), None)
    _sect_col  = next((c for c in all_cols if _nc(c) == "section" or ("section" in _nc(c) and "sub" not in _nc(c))), None)
    _class_col = next((c for c in all_cols if "class" in _nc(c) or "subclass" in _nc(c)), None)

    _mj1, _mj2, _mj3 = st.tabs(["Department", "Section", "Class"])

    def _group_summary(col):
        if col and col in merged.columns:
            _g = merged.groupby(col).size().reset_index(name="SKUs")
            _g = _g.sort_values("SKUs", ascending=False).reset_index(drop=True)
            return _g
        return None

    with _mj1:
        _df_dept = _group_summary(_dept_col)
        if _df_dept is not None:
            st.dataframe(_df_dept, use_container_width=True, hide_index=True, height=420)
            st.caption(f"{len(_df_dept)} departments · {len(merged):,} total SKUs")
        else:
            st.info("Department column not found in the uploaded data.")

    with _mj2:
        _df_sect = _group_summary(_sect_col)
        if _df_sect is not None:
            st.dataframe(_df_sect, use_container_width=True, hide_index=True, height=420)
            st.caption(f"{len(_df_sect)} sections · {len(merged):,} total SKUs")
        else:
            st.info("Section column not found in the uploaded data.")

    with _mj3:
        _df_cls = _group_summary(_class_col)
        if _df_cls is not None:
            st.dataframe(_df_cls, use_container_width=True, hide_index=True, height=420)
            st.caption(f"{len(_df_cls)} classes · {len(merged):,} total SKUs")
        else:
            st.info("Class / Subclass column not found in the uploaded data.")

    render_page_nav("viewdata")
    st.stop()

# ── REFRESH — Store No. filter ─────────────────────────────────────────────────
if _live_type == "Refresh":
    _store_col = (
        next((c for c in all_cols if "store" in _nc(c) and "no" in _nc(c)), None)
        or next((c for c in all_cols if _nc(c) in ("store", "store no", "store no.", "store number")), None)
        or next((c for c in all_cols if "store" in _nc(c)), None)
    )

    st.markdown("""
<div style="background:#fff;border-radius:14px;border:1px solid #E8E3DC;
            padding:20px 24px 16px;">
    <div style="font-size:15px;font-weight:700;color:#1A1A1A;margin-bottom:4px;">
        Store Filter
    </div>
    <div style="font-size:12px;color:#888;margin-bottom:14px;">
        Select one or more Store No. values — the table below shows only matching rows.
    </div>
</div>""", unsafe_allow_html=True)

    if _store_col:
        _store_opts = sorted(
            merged[_store_col].dropna().astype(str).str.strip().unique().tolist()
        )

        _sel_stores = st.multiselect(
            "STORE NO.",
            _store_opts,
            key="vw_store_sel",
            placeholder="TYPE TO SEARCH OR SELECT STORES…",
        )

        if _sel_stores:
            _store_df = merged[
                merged[_store_col].astype(str).str.strip().isin(_sel_stores)
            ].reset_index(drop=True)

            st.markdown(
                f"<div style='font-size:12px;color:#888;margin:12px 0 6px;'>"
                f"Showing <strong>{len(_store_df):,}</strong> rows for "
                f"<strong>{len(_sel_stores)}</strong> store(s)</div>",
                unsafe_allow_html=True,
            )
            st.dataframe(_store_df, use_container_width=True, hide_index=True, height=460)

            _r1, _r2 = st.columns(2)
            with _r1:
                st.download_button(
                    "⬇️ Export .xlsx", df_to_xlsx_bytes(_store_df),
                    file_name=f"store_filter_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with _r2:
                st.download_button(
                    "⬇️ Export .csv", df_to_csv_bytes(_store_df),
                    file_name=f"store_filter_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv", use_container_width=True,
                )
        else:
            st.info("Select at least one Store No. above to view filtered data.")
    else:
        st.warning("No Store No. column found in the uploaded data.")

    render_page_nav("viewdata")
    st.stop()

def _render_sheet_content(df_src, p, dg_col_hint=None, large_file_path=None, dg_options=None,
                           fixed_dg_code=None, fixed_dg_name=None):
    """Full All Data + New Canvas tab pair, keyed with prefix p.
    dg_col_hint    : raw DG column name — shown as first column in Table view.
    large_file_path: when set, DG code search loads the full slice from this file.
    dg_options     : list of known DG codes — renders a searchable combobox instead of text input.
    fixed_dg_code/fixed_dg_name: when set, df_src is already filtered to this DG by the caller
    (e.g. a picker rendered above the tab) — skip the redundant Search DG Code/DG Name widgets
    here and just display the already-chosen DG."""
    _src_cols = list(df_src.columns)
    _meta = st.session_state.rangesheet_meta

    # Build column lookups from df_src directly — not from the module-level
    # merged/all_cols — so this works correctly when df_src is HDET or any
    # other file with different column names than merged.
    _s_dcm  = {_nc(c): c for c in _src_cols}
    _s_dcm2 = {_nca(c): c for c in _src_cols}
    def _src_fc(key):
        return _s_dcm.get(_nc(key)) or _s_dcm2.get(_nca(key))

    _sag = []
    for _grp in RS_COL_GROUPS:
        _gm = [_src_fc(k) for k in _grp["cols"] if _src_fc(k) is not None]
        if _gm:
            _sag.append({**_grp, "matched": _gm})
    _sag_flat = [c for g in _sag for c in g["matched"]]
    # Only RS_COL_GROUPS columns are shown here; all other columns belong to View Data page.

    # _subview = st.radio("sub_view",["📋 Table","🏪 Cluster","📊 Status"],
    #                     horizontal=True, label_visibility="collapsed", key=f"{p}_subview")

    _subview = "📋 Table"
    _tab_all = st.container()

    # ── ALL DATA ──────────────────────────────────────────────────────────────
    with _tab_all:
        _dg_code_col = (
            next((c for c in _src_cols if _nc(c) in ("dg code", "dg_code", "dg")), None)
            or next((c for c in _src_cols if _nca(c) in [_nca(x) for x in DG_COLUMN_CANDIDATES]), None)
            or next((c for c in _src_cols if "dg" in _nc(c) and "code" in _nc(c)), None)
            or next((c for c in _src_cols if "department" in _nc(c)), None)
        )
        _dg_name_col = (
            next((c for c in _src_cols if _nc(c) in ("dg name", "dg_name")), None)
            or next((c for c in _src_cols if "dg" in _nc(c) and "name" in _nc(c)), None)
            or next((c for c in _src_cols if "display" in _nc(c) and ("group desc" in _nc(c) or "desc" in _nc(c))), None)
            or next((c for c in _src_cols if _nc(c) == "section"), None)
            or next((c for c in _src_cols if "section" in _nc(c)), None)
        )
        _dg_c, _arch_c, _leg_c = st.columns([1.05, 2.9, 0.75])
        _sel_dg_code = ""
        _sel_dg_name = "ALL"

        with _dg_c:
            st.markdown("""<div style="background:#fff;border-radius:14px;border:1px solid #E8E3DC;
                        padding:14px 14px 0 14px;">
            <div style="font-size:10px;font-weight:700;color:#888;text-transform:uppercase;
                        letter-spacing:.08em;margin-bottom:10px;">Display Group</div></div>""",
                unsafe_allow_html=True)
            with st.container():
                if fixed_dg_code is not None or fixed_dg_name is not None:
                    # Already filtered to one DG by the picker above this tab —
                    # don't render another search box, just show the choice.
                    # _sel_dg_name stays "ALL" so the name-filter below is a no-op:
                    # df_src was loaded by DG Code already, and the name column this
                    # function detects locally may not be the same one the picker
                    # above used to resolve fixed_dg_name.
                    _sel_dg_code = fixed_dg_code or ""
                    st.session_state[f"{p}_selected_dg_code"] = _sel_dg_code
                    st.session_state[f"{p}_selected_dg_name"] = fixed_dg_name or "ALL"
                    st.markdown(
                        f'<div style="font-size:12px;color:#555;margin-bottom:8px;">'
                        f'DG Code: <strong>{_sel_dg_code or "—"}</strong><br>'
                        f'DG Name: <strong>{fixed_dg_name or "—"}</strong></div>',
                        unsafe_allow_html=True,
                    )
                else:
                    _ALL_OPT = "— All (preview 500 rows) —"
                    if dg_options:
                        _dg_combo_opts = [_ALL_OPT] + list(dg_options)
                        _prev_code = str(st.session_state.get(f"{p}_selected_dg_code", "") or "")
                        _default_code = _prev_code if _prev_code in _dg_combo_opts else _ALL_OPT
                        if st.session_state.get(f"{p}_dg_code") not in _dg_combo_opts:
                            st.session_state[f"{p}_dg_code"] = _default_code
                        _dg_combo_val  = st.selectbox(
                            "Search DG Code", _dg_combo_opts,
                            index=_dg_combo_opts.index(_default_code),
                            key=f"{p}_dg_code", label_visibility="visible",
                        )
                        _sel_dg_code = "" if _dg_combo_val == _ALL_OPT else _dg_combo_val
                    else:
                        _sel_dg_code = st.text_input(
                            "Search DG Code", placeholder="Type DG code to filter…",
                            key=f"{p}_dg_code", label_visibility="visible")
                    st.session_state[f"{p}_selected_dg_code"] = _sel_dg_code
                    if _dg_code_col and _dg_name_col and _sel_dg_code:
                        _ns = df_src[df_src[_dg_code_col].astype(str).str.contains(_sel_dg_code, case=False, na=False)]
                        _dg_name_opts = ["ALL"] + sorted(_ns[_dg_name_col].dropna().astype(str).str.strip().unique().tolist())
                    elif _dg_name_col:
                        _dg_name_opts = ["ALL"] + sorted(df_src[_dg_name_col].dropna().astype(str).str.strip().unique().tolist())
                    else:
                        _dg_name_opts = ["ALL"]
                    _prev_name = st.session_state.get(f"{p}_selected_dg_name", st.session_state.get(f"_{p}_dg_name_prev", "ALL"))
                    if _prev_name not in _dg_name_opts:
                        _prev_name = "ALL"
                    if st.session_state.get(f"{p}_dg_name") not in _dg_name_opts:
                        st.session_state[f"{p}_dg_name"] = _prev_name
                    _sel_dg_name = st.selectbox(
                        "DG NAME", _dg_name_opts,
                        index=(_dg_name_opts.index(_prev_name) if _prev_name in _dg_name_opts else 0),
                        key=f"{p}_dg_name", label_visibility="visible")
                    st.session_state[f"_{p}_dg_name_prev"] = _sel_dg_name
                    st.session_state[f"{p}_selected_dg_name"] = _sel_dg_name
                _static_rows = ""
                for _sk, _sv2 in [
                    ("MINOR LIVE WEEK", _meta.get("minor_live_week", "—")),
                    ("MAJOR LIVE WEEK", _meta.get("major_live_week", "—")),
                    ("Event Live Date", _meta.get("event_live_date", "—")),
                    ("Event Des",       _meta.get("event_desc",       "—")),
                ]:
                    _static_rows += (
                        f'<div style="display:flex;justify-content:space-between;align-items:center;'
                        f'margin-bottom:5px;font-size:11px;gap:4px;">'
                        f'<span style="color:#888;white-space:nowrap;">{_sk}</span>'
                        f'<span style="font-weight:600;color:#555;text-align:right;">{_sv2}</span></div>')
                st.markdown(f"""<div style="background:#fff;border-radius:0 0 14px 14px;
                            border:1px solid #E8E3DC;border-top:none;padding:10px 14px 10px;">
                    {_static_rows}</div>""", unsafe_allow_html=True)

        # Persisted dropdown edits must be available before summary panels render.
        # The main grid loads this state later too, but range architecture/cluster
        # are above the grid and otherwise miss edits after a browser refresh.
        _pre_sticky_for_state = [c for c in ["DG Code", "ID", "Item Name"] if c in df_src.columns]
        _pre_scope_candidates = []
        for _candidate_sticky in (
            _pre_sticky_for_state,
            ["DG Code", "ID", "Item Name"],
            ["ID", "Item Name"],
        ):
            if _candidate_sticky not in _pre_scope_candidates:
                _pre_scope_candidates.append(_candidate_sticky)
        _pre_edit_state_scope = (
            f"{p}|{large_file_path or ''}|"
            f"{getattr(df_src, 'shape', ('', ''))}|"
            f"{_sel_dg_code or ''}|{_sel_dg_name or ''}|"
            f"{','.join(map(str, _pre_sticky_for_state))}"
        )
        _pre_edit_state_scope_key = f"{p}_edit_state_scope"
        if st.session_state.get(_pre_edit_state_scope_key) != _pre_edit_state_scope:
            _loaded_edit_state = {}
            _loaded_edit_scope = _pre_edit_state_scope
            for _scope_sticky in _pre_scope_candidates:
                _candidate_scope = (
                    f"{p}|{large_file_path or ''}|"
                    f"{getattr(df_src, 'shape', ('', ''))}|"
                    f"{_sel_dg_code or ''}|{_sel_dg_name or ''}|"
                    f"{','.join(map(str, _scope_sticky))}"
                )
                _candidate_state = _rs_load_edit_state(_candidate_scope)
                if any(_candidate_state.get(_k) for _k in ("pog_actions", "pog_edits", "avg_u_edits", "data_edits", "status_overrides")):
                    _loaded_edit_state = _candidate_state
                    _loaded_edit_scope = _candidate_scope
                    break
                _legacy_scope = (
                    f"{p}|{large_file_path or ''}|"
                    f"{getattr(df_src, 'shape', ('', ''))}|{','.join(map(str, _scope_sticky))}"
                )
                _legacy_state = _rs_load_edit_state(_legacy_scope)
                if any(_legacy_state.get(_k) for _k in ("pog_actions", "pog_edits", "avg_u_edits", "data_edits", "status_overrides")):
                    _loaded_edit_state = _legacy_state
                    _loaded_edit_scope = _candidate_scope
                    break
            st.session_state[_pre_edit_state_scope_key] = _loaded_edit_scope
            st.session_state[f"{p}_pog_actions"] = _loaded_edit_state.get("pog_actions", {})
            st.session_state[f"{p}_pog_edits"] = _loaded_edit_state.get("pog_edits", {})
            st.session_state[f"{p}_avg_u_edits"] = _loaded_edit_state.get("avg_u_edits", {})
            st.session_state[f"{p}_data_edits"] = _loaded_edit_state.get("data_edits", {})
            st.session_state[f"{p}_status_overrides"] = _loaded_edit_state.get("status_overrides", {})

        with _arch_c:
            _type_col = (
                next((c for c in _src_cols if _nc(c) in ("item priority","itempriority")), None)
                or next((c for c in _src_cols if "item" in _nc(c) and "priority" in _nc(c)), None)
                or next((c for c in _src_cols if _nc(c) == "status"), None)
            )
            _stc = next((c for c in _src_cols if _nc(c) == "status"), None)
            _price_col = next((c for c in _src_cols if "avg selling price" in _nc(c) or ("selling price" in _nc(c) and "avg" in _nc(c))), None)
            _edlp_col  = next((c for c in _src_cols if "edlp price" in _nc(c)), None)
            _asis_stc  = next((c for c in _src_cols if ("as-is stores applied" in _nc(c) or ("as is" in _nc(c) and "stores applied" in _nc(c))) and "to" not in _nc(c)[:4]), None)
            _tobe_stc  = next((c for c in _src_cols if "to-be stores applied" in _nc(c) or "to be stores applied" in _nc(c) or "to-be stores" in _nc(c)), None)
            _ab = df_src.copy()
            if _sel_dg_code and _dg_code_col and _dg_code_col in _ab.columns:
                _ab = _ab[_ab[_dg_code_col].astype(str).str.contains(_sel_dg_code, case=False, na=False)]
            if _sel_dg_name != "ALL" and _dg_name_col and _dg_name_col in _ab.columns:
                _ab = _ab[_ab[_dg_name_col].astype(str).str.strip() == _sel_dg_name]
            _is_sspog = (p == "ss")
            TYPES = ["MAINTAIN"] + (["NEW DELETE SOME"] if _is_sspog else []) + ["DELETE SOME","DELETE ALL","NEW SOME","NEWNEW"]
            _status_ov_live = st.session_state.get(f"{p}_status_overrides", {})
            _arch_action_store = {}
            for _store in (
                st.session_state.get(f"{p}_pog_actions", {}),
                st.session_state.get(f"{p}_pog_edits", {}),
            ):
                if isinstance(_store, dict):
                    for _rk_a, _acts_a in _store.items():
                        if isinstance(_acts_a, dict):
                            _arch_action_store.setdefault(_rk_a, {}).update(_acts_a)
            def _arch_status_from_actions(_acts):
                if not isinstance(_acts, dict) or not _acts:
                    return ""
                _vals = [str(v).strip().lower() for v in _acts.values() if str(v).strip()]
                if not _vals:
                    return ""
                _del_n = sum(1 for v in _vals if v == "delete")
                _new_n = sum(1 for v in _vals if v == "new")
                if _del_n and _new_n:
                    return "NEW DELETE SOME" if _is_sspog else "DELETE SOME"
                if _del_n:
                    return "DELETE SOME"
                if _new_n:
                    return "NEW SOME"
                return ""
            _arch_action_status_by_rk = {}
            _arch_action_status_by_id = {}
            for _rk_a, _acts_a in _arch_action_store.items():
                _st_a = _arch_status_from_actions(_acts_a)
                if not _st_a:
                    continue
                _arch_action_status_by_rk[_rk_a] = _st_a
                if isinstance(_rk_a, tuple) and len(_rk_a) > 1:
                    _arch_action_status_by_id[str(_rk_a[1])] = _st_a
            _id_col_arch = next((c for c in _ab.columns if _nca(c) == _nca("ID")), None)
            _item_col_arch = (
                next((c for c in _ab.columns if _nca(c) == _nca("Item Name")), None)
                or next((c for c in _ab.columns if "item" in _nc(c) and "name" in _nc(c)), None)
            )
            _sticky_for_arch = [c for c in [_dg_code_col, _id_col_arch, _item_col_arch] if c and c in _ab.columns]
            if _id_col_arch and _id_col_arch in _ab.columns:
                _ab = _ab.drop_duplicates(subset=[_id_col_arch], keep="first").copy()
            def _arch_status_series(base: pd.DataFrame):
                _asis = pd.Series(["MAINTAIN"] * len(base), index=base.index)
                _tobe = _asis.copy()
                _ov_by_id = {}
                for _rk, _ov in _status_ov_live.items():
                    if not isinstance(_rk, tuple):
                        continue
                    if len(_rk) >= 3 and _sel_dg_code and str(_rk[0]).strip() != str(_sel_dg_code).strip():
                        continue
                    if len(_rk) > 1:
                        _ov_by_id[str(_rk[1])] = _ov
                if (_status_ov_live or _arch_action_status_by_rk) and _id_col_arch and _id_col_arch in base.columns:
                    _mapped_status = base[_id_col_arch].astype(str).map(
                        lambda _idv: _ov_by_id.get(_idv) or _arch_action_status_by_id.get(_idv) or ""
                    )
                    _mapped_status = _mapped_status.astype(str).str.strip().str.upper()
                    _mapped_mask = _mapped_status.ne("")
                    if _mapped_mask.any():
                        _tobe.loc[_mapped_mask] = _mapped_status.loc[_mapped_mask]
                elif _status_ov_live or _arch_action_status_by_rk:
                    for _idx2, _row2 in base.iterrows():
                        _rk2 = tuple(str(_row2.get(c, "")) for c in _sticky_for_arch) if _sticky_for_arch else None
                        _ov2 = _status_ov_live.get(_rk2) if _rk2 else None
                        if not _ov2 and _rk2:
                            _ov2 = _arch_action_status_by_rk.get(_rk2)
                        if not _ov2 and _id_col_arch:
                            _id2 = str(_row2.get(_id_col_arch, ""))
                            _ov2 = _ov_by_id.get(_id2) or _arch_action_status_by_id.get(_id2)
                        if _ov2:
                            _tobe.at[_idx2] = str(_ov2).strip().upper()
                return _asis, _tobe
            _asis_status, _tobe_status = _arch_status_series(_ab)
            def _fmt(v):
                if v == 0: return "0"
                try: return f"{int(v):,}" if v == int(v) else f"{v:,.1f}"
                except: return str(v)
            _arch_rows = []
            for _t in TYPES:
                _idx_ai = _ab.index[_asis_status == _t]
                _idx_tb = _ab.index[_tobe_status == _t]
                _idx = _idx_tb
                _ai = len(_idx_ai)
                _tb = len(_idx_tb)
                _prices = pd.to_numeric(_ab.loc[_idx, _price_col], errors="coerce").fillna(0) if (_price_col and len(_idx) > 0) else pd.Series([], dtype=float)
                _asis_s = pd.to_numeric(_ab.loc[_idx, _asis_stc], errors="coerce").fillna(0) if (_asis_stc and len(_idx) > 0) else pd.Series([1]*len(_idx), dtype=float)
                _tobe_s = pd.to_numeric(_ab.loc[_idx, _tobe_stc], errors="coerce").fillna(0) if (_tobe_stc and len(_idx) > 0) else pd.Series([1]*len(_idx), dtype=float)
                _arch_rows.append({"type":_t,"as_is":_ai,"to_be":_tb,
                    "sale_ai":float((_prices*_asis_s).sum()) if len(_prices) else 0.0,
                    "sale_tb":float((_prices*_tobe_s).sum()) if len(_prices) else 0.0,
                    "sale_diff":float((_prices*_tobe_s).sum()-(_prices*_asis_s).sum()) if len(_prices) else 0.0,
                    "marg_ai":0.0})
            _t_ai=sum(r["as_is"] for r in _arch_rows)
            _delete_all_tb = next((r["to_be"] for r in _arch_rows if r["type"] == "DELETE ALL"), 0)
            _newnew_tb = next((r["to_be"] for r in _arch_rows if r["type"] == "NEWNEW"), 0)
            _t_tb=_t_ai - _delete_all_tb + _newnew_tb
            _t_sai=sum(r["sale_ai"] for r in _arch_rows); _t_stb=sum(r["sale_tb"] for r in _arch_rows)
            _t_sdiff=sum(r["sale_diff"] for r in _arch_rows); _t_mai=sum(r["marg_ai"] for r in _arch_rows)
            _pct_sale=f"{(_t_sdiff/_t_sai*100):.1f}%" if _t_sai else "0.0%"
            _B="border:1px solid #B8B8B8;"; _BR="border-right:2px solid #999;"
            _TH="padding:7px 8px;text-align:center;font-size:10px;font-weight:700;border:1px solid #B8B8B8;background:#D9D9D9;color:#333;"

            if p == "sa":
                # ── StoreApply: simplified TYPE | AS IS | TO BE only ──────────
                tbody=""
                for _r in _arch_rows:
                    _t=_r["type"]
                    tbody+=(f'<tr style="border-bottom:1px solid #D8D8D8;">'
                        f'<td style="padding:5px 10px;font-size:11px;color:#1A1A1A;{_B}">{_t}</td>'
                        f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}">{_fmt(_r["as_is"])}</td>'
                        f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}">{_fmt(_r["to_be"])}</td>'
                        f'</tr>')
                tbody+=(f'<tr><td style="padding:6px 10px;font-size:11px;font-weight:800;{_B}">TOTAL SKU</td>'
                    f'<td style="padding:6px 8px;text-align:center;font-size:11px;font-weight:700;{_B}">{_fmt(_t_ai)}</td>'
                    f'<td style="padding:6px 8px;text-align:center;font-size:11px;font-weight:700;{_B}">{_fmt(_t_tb)}</td></tr>'
                    f'<tr><td style="padding:5px 10px;font-size:11px;font-style:italic;{_B}">% Impact</td>'
                    f'<td style="{_B}"></td>'
                    f'<td style="text-align:center;color:#00AA00;font-weight:700;font-size:11px;{_B}">0.0%</td></tr>')
                st.markdown(f"""<div style="background:#fff;overflow:hidden;border:1px solid #B8B8B8;">
                  <div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;font-size:11px;">
                    <thead>
                      <tr><th colspan="3" style="{_TH}text-align:left;min-width:130px;">Range architecture</th></tr>
                      <tr>
                        <th style="{_TH}text-align:left;min-width:140px;">TYPE</th>
                        <th style="{_TH}min-width:60px;">AS IS</th>
                        <th style="{_TH}min-width:60px;">TO BE</th>
                      </tr>
                    </thead><tbody>{tbody}</tbody></table></div>
                </div>""", unsafe_allow_html=True)

            else:
                # ── Non-SSPOG / SSPOG: full table with Sale + Margin Impact ───
                tbody=""
                for _r in _arch_rows:
                    _t=_r["type"]
                    tbody+=(f'<tr style="border-bottom:1px solid #D8D8D8;">'
                        f'<td style="padding:5px 10px;font-size:11px;color:#1A1A1A;{_B}">{_t}</td>'
                        f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}">{_fmt(_r["as_is"])}</td>'
                        f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}{_BR}">{_fmt(_r["to_be"])}</td>'
                        f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}">{_fmt(_r["sale_ai"])}</td>'
                        f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}">{_fmt(_r["sale_tb"])}</td>'
                        f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}{_BR}">{_fmt(_r["sale_diff"])}</td>'
                        f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}">{_fmt(_r["marg_ai"])}</td></tr>')
                tbody+=(f'<tr style="font-weight:800;"><td style="padding:6px 10px;font-size:11px;font-weight:800;{_B}">TOTAL SKU</td>'
                    f'<td style="padding:6px 8px;text-align:center;font-size:11px;font-weight:700;{_B}">{_fmt(_t_ai)}</td>'
                    f'<td style="padding:6px 8px;text-align:center;font-size:11px;font-weight:700;{_B}{_BR}">{_fmt(_t_tb)}</td>'
                    f'<td style="padding:6px 8px;text-align:center;font-size:11px;font-weight:700;{_B}">{_fmt(_t_sai)}</td>'
                    f'<td style="padding:6px 8px;text-align:center;font-size:11px;font-weight:700;{_B}">{_fmt(_t_stb)}</td>'
                    f'<td style="padding:6px 8px;text-align:center;font-size:11px;font-weight:700;{_B}{_BR}">{_fmt(_t_sdiff)}</td>'
                    f'<td style="padding:6px 8px;text-align:center;font-size:11px;font-weight:700;{_B}">{_fmt(_t_mai)}</td></tr>'
                    f'<tr><td style="padding:5px 10px;font-size:11px;font-style:italic;{_B}">% Impact</td>'
                    f'<td style="{_B}"></td>'
                    f'<td style="text-align:center;color:#00AA00;font-weight:700;font-size:11px;{_B}{_BR}">0.0%</td>'
                    f'<td style="{_B}"></td><td style="{_B}"></td>'
                    f'<td style="text-align:center;color:#00AA00;font-weight:700;font-size:11px;{_B}{_BR}">{_pct_sale}</td>'
                    f'<td style="{_B}"></td></tr>')
                st.markdown(f"""<div style="background:#fff;overflow:hidden;border:1px solid #B8B8B8;">
                  <div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;font-size:11px;">
                    <thead><tr>
                      <th colspan="3" style="{_TH}text-align:left;min-width:130px;">Range architecture</th>
                      <th colspan="3" style="{_TH}">Sale Impact ( ex.vat) / Week<br>calcualte from Mer Price</th>
                      <th colspan="1" style="{_TH}">Margin Impact ( ex.vat) / Week<br>calcualte from EDLP Price</th>
                    </tr><tr>
                      <th style="{_TH}text-align:left;">TYPE</th>
                      <th style="{_TH}min-width:52px;">AS IS</th>
                      <th style="{_TH}min-width:52px;border-right:2px solid #999;">TO BE</th>
                      <th style="{_TH}min-width:60px;">AS IS</th>
                      <th style="{_TH}min-width:60px;">TO BE</th>
                      <th style="{_TH}min-width:52px;border-right:2px solid #999;">DIFF</th>
                      <th style="{_TH}min-width:60px;">AS IS</th>
                    </tr></thead><tbody>{tbody}</tbody></table></div>
                  <div style="text-align:right;padding:3px 8px;font-size:9px;color:#888;
                              border-top:1px solid #E0E0E0;background:#F8F8F8;">
                    display mgr = Avg selling price from format</div></div>""", unsafe_allow_html=True)

        with _leg_c:
            st.empty()

        st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)

        _new_rows_scope = f"{p}|{_sel_dg_code or 'ALL'}|{_sel_dg_name or 'ALL'}|{large_file_path or ''}"
        _new_rows_key = f"{p}_new_item_rows"
        _new_rows_placeholder_key = f"{p}_new_item_placeholder_rows"
        _new_rows_scope_key = f"{p}_new_item_rows_scope"
        _new_rows_scroll_key = f"{p}_new_item_scroll_key"
        _new_rows_pending_key = f"{p}_new_item_pending_add"
        if st.session_state.get(_new_rows_scope_key) != _new_rows_scope:
            st.session_state[_new_rows_scope_key] = _new_rows_scope
            st.session_state[_new_rows_key] = _rs_load_new_rows(_new_rows_scope)
            st.session_state[_new_rows_placeholder_key] = []
            st.session_state[_new_rows_scroll_key] = ""

        _c2, _c3, _cadd, _c4 = st.columns([4, 1, 1.25, 1.2])
        with _c2:
            _search_q = st.text_input("Search", placeholder="🔎  Search item, barcode, status...",
                                      label_visibility="collapsed", key=f"{p}_search")
        with _c3:
            _n_rows = st.number_input("Rows", min_value=10, max_value=100000, value=500, step=100,
                                      help="Max rows to display", key=f"{p}_nrows")
        with _cadd:
            if st.button("Add new item", key=f"{p}_add_new_item", use_container_width=True):
                st.session_state[_new_rows_pending_key] = int(st.session_state.get(_new_rows_pending_key, 0)) + 10
        with _c4:
            _edit_on = st.toggle("⚙️ Edit Columns", key=f"{p}_edit_cols")

        _pvis_key = f"{p}_vis_cols"
        if _pvis_key not in st.session_state:
            st.session_state[_pvis_key] = None
        if st.session_state[_pvis_key] is not None:
            if not any(c in _src_cols for c in st.session_state[_pvis_key]):
                st.session_state[_pvis_key] = None

        if _edit_on:
            st.markdown("""<div style="background:#fff;border-radius:12px;border:1px solid #E8E3DC;
                        padding:12px 16px 8px;margin-bottom:12px;">
              <div style="font-size:12px;font-weight:700;color:#1A1A1A;margin-bottom:8px;">⚙️ Column Visibility</div>
            </div>""", unsafe_allow_html=True)
            _new_vis = []
            for _g in _sag:
                st.caption(f"— {_g['group']} —")
                _sel = st.multiselect(_g["group"], _g["matched"],
                    default=[c for c in _g["matched"] if c in (st.session_state[_pvis_key] or _g["matched"])],
                    key=f"{p}_vv_{_g['group']}", label_visibility="collapsed")
                _new_vis.extend(_sel)
            _ec1, _ec2 = st.columns(2)
            with _ec1:
                if st.button("✅ Apply", key=f"{p}_vv_apply"):
                    st.session_state[_pvis_key] = _new_vis or _sag_flat[:]
                    st.rerun()
            with _ec2:
                if st.button("↺ Show All", key=f"{p}_vv_reset"):
                    st.session_state[_pvis_key] = _sag_flat[:]
                    st.rerun()

        _vis_set = set(st.session_state[_pvis_key] or [])
        _matched_cols = [c for c in _sag_flat if c in df_src.columns and (not _vis_set or c in _vis_set)]
        disp_cols = _matched_cols  # Only RS_COL_GROUPS columns — no fallback to all cols

        if not disp_cols:
            st.info("No columns in this dataset match the RangeSheet standard headers. "
                    "Check the raw file on the View Data page.")
            return  # exit only this function, not the whole page

        # Filter + search
        _pog_c  = next((c for c in _src_cols if "pog" in c.lower() and "cluster" in c.lower()), None)

        if _sel_dg_code and large_file_path and not fixed_dg_code:
            # Search full HDET file — the 500-row preview may not contain this DG
            _q   = _sel_dg_code.strip()
            _sig = f"{large_file_path}|{_q.upper()}"
            _sk  = f"{p}_dg_search_sig"
            _dk  = f"{p}_dg_search_df"
            if st.session_state.get(_sk) != _sig:
                with st.spinner(f"Searching entire HDET for DG '{_q}'…"):
                    _result = load_large_file_by_dg(large_file_path, _q)
                st.session_state[_sk] = _sig
                st.session_state[_dk] = _result
            df_view = st.session_state[_dk].copy(deep=False)
            # Keep the lookup quiet; users only need the filtered table here.
        else:
            df_view = df_src.copy(deep=False)
            if _sel_dg_code:
                _q = _sel_dg_code.strip()
                if _dg_code_col and _dg_code_col in df_view.columns:
                    df_view = df_view[
                        df_view[_dg_code_col].astype(str).str.strip()
                        .str.contains(_q, case=False, na=False)
                    ]
                else:
                    _scan_cols = [c for c in df_view.columns if c in df_view.columns]
                    if _scan_cols:
                        _scan_mask = df_view[_scan_cols].astype(str).apply(
                            lambda _col: _col.str.contains(_q, case=False, na=False, regex=False)
                        ).any(axis=1)
                        df_view = df_view[_scan_mask]

        if _sel_dg_name != "ALL" and _dg_name_col and _dg_name_col in df_view.columns:
            df_view = df_view[df_view[_dg_name_col].astype(str).str.strip() == _sel_dg_name]
        _search_q_text = str(_search_q or "").strip()
        _search_q_l = _search_q_text.lower()
        _search_q_n = _nca(_search_q_text)
        _status_search_terms = (
            "maintain", "delete", "new", "newnew", "some", "all",
            "delist", "inactive", "สถานะ", "ลบ", "เพิ่ม"
        )
        _apply_source_search = (
            bool(_search_q_text)
            and not any(_t in _search_q_l for _t in _status_search_terms)
        )
        if _search_q_text:
            _pog_search_src_col = (
                _src_fc("Planogram Name")
                or next((c for c in df_view.columns
                         if _nca(c) in ("name", "planogramname", "pogname", "planogram")), None)
            )
            if _pog_search_src_col and _search_q_n:
                _pog_search_vals = df_view[_pog_search_src_col].astype(str).map(_nca)
                if _pog_search_vals.str.contains(_search_q_n, case=False, na=False, regex=False).any():
                    _apply_source_search = False

        if _apply_source_search:
            _search_cols = [c for c in df_view.columns]
            if _search_cols:
                _search_mask_src = df_view[_search_cols].astype(str).apply(
                    lambda _col: _col.str.contains(
                        _search_q_text, case=False, na=False, regex=False
                    )
                ).any(axis=1)
                df_view = df_view[_search_mask_src]
        df_view = df_view.reset_index(drop=True)

        # _m_maintain = int((df_view[_stc].astype(str).str.strip()=="MAINTAIN").sum()) if _stc else 0
        # _m_sspog    = int(df_view[_pog_c].astype(str).str.contains("SSPOG",na=False).sum() - df_view[_pog_c].astype(str).str.contains("Non-SSPOG",na=False).sum()) if _pog_c else 0
        # _m_nonsspog = int(df_view[_pog_c].astype(str).str.contains("Non-SSPOG",na=False).sum()) if _pog_c else 0
        # _m_null     = int(df_view[disp_cols].isnull().sum().sum())
        # m1,m2,m3,m4,m5 = st.columns(5)
        # m1.metric("Total SKUs",f"{len(df_view):,}"); m2.metric("MAINTAIN",f"{_m_maintain:,}")
        # m3.metric("SSPOG",f"{_m_sspog:,}"); m4.metric("Non-SSPOG",f"{_m_nonsspog:,}"); m5.metric("Null Values",f"{_m_null:,}")

        # ── Table ─────────────────────────────────────────────────────────────
        if _subview == "📋 Table":
            _MAX = int(_n_rows)
            # Fixed predefined column structure — headers never change.
            # Data is pulled from df_view by _nca() name matching; empty where no match.
            _std_all = [c for grp in RS_COL_GROUPS for c in grp["cols"]]

            df_view = df_view.rename(columns=lambda c: COLUMN_MAPPING.get(_nca(c), c))
            _t_col_map = {
                sc: next((c for c in df_view.columns if _nca(c) == _nca(sc)), None)
                for sc in _std_all
            }
            # Visibility filter: if user applied one, keep only cols whose matched
            # actual name (or standard name itself) is in the visible set.
            if _vis_set:
                _std_all = [sc for sc in _std_all
                            if (_t_col_map.get(sc) in _vis_set) or sc in _vis_set]
            _n = len(df_view)   # use full df_view for pivot; cap at _MAX products after
            def _safe_col(src_df, col_name, n):
                s = src_df[col_name]
                if isinstance(s, pd.DataFrame):   # duplicate col names → take first
                    s = s.iloc[:, 0]
                return s.iloc[:n].reset_index(drop=True)
            _STATUS_COL_NCA = _nca("Status")
            _tdf = pd.DataFrame({
                sc: (_safe_col(df_view, _t_col_map[sc], _n)
                     if _t_col_map.get(sc)
                     else pd.Series([0] * _n) if _nca(sc) in _ZERO_DEFAULT_COLS
                     else pd.Series(["MAINTAIN"] * _n) if _nca(sc) == _STATUS_COL_NCA
                     else pd.Series([""] * _n))
                for sc in _std_all
            })
            # Prepend DG Code as first column if available
            _dg_raw_col = next(
                (c for c in df_view.columns if dg_col_hint and _nca(c) == _nca(dg_col_hint)), None
            ) if dg_col_hint else None
            if _dg_raw_col and "DG Code" not in _tdf.columns:
                _tdf.insert(0, "DG Code", _safe_col(df_view, _dg_raw_col, _n))

            # Dynamic planogram columns: one column per unique NAME — no count limit
            _pog_src_col = _t_col_map.get("Planogram Name") or next(
                (c for c in df_view.columns if _nca(c) in ("name","planogramname","pogname","planogram")), None
            )

            # Sales-value column — match TH_Tot_Sales_Value_52WK (Value, not Volume).
            # _nca() strips underscores/spaces so "TH_Tot_Sales_Value_52WK" normalises
            # to "thtotsalesvalue52wk", distinct from Volume's "thtotsalesvolume52wk".
            _sv_nca = _nca("TH_Tot_Sales_Value_52WK")   # → "thtotsalesvalue52wk"
            _sv_col = next((c for c in df_view.columns if _nca(c) == _sv_nca), None)
            if not _sv_col:
                st.warning(
                    "Column **TH_Tot_Sales_Value_52WK** (Value) not found — "
                    "planogram value cells will be blank. Check the column name in your file."
                )
            else:
                # Internal source detail; do not show on the review page.
                pass

            _dyn_pog_cols = []
            _piv_pog_cols = []   # set inside pivot block; guards outer debug expanders
            _pk_std       = []
            _pk_raw       = []
            _pk_map       = {}
            _cs_base_pog_counts = {}
            if _pog_src_col:
                # All planogram names from full df_view (no row cap — cap products after pivot)
                _pog_vals_all = df_view[_pog_src_col].astype(str).str.strip()
                _dyn_pog_cols = sorted(v for v in _pog_vals_all.unique() if v and v not in ("nan","None",""))

                if "Planogram Name" in _tdf.columns:
                    _tdf = _tdf.drop(columns=["Planogram Name"])

                # Product-key standard names present in _tdf → their raw df_view counterparts
                _proto_sticky = [c for c in ["DG Code", "ID", "Item Name"] if c in _tdf.columns]
                _pk_map = {}
                for _std in _proto_sticky:
                    _raw = next((c for c in df_view.columns if _nca(c) == _nca(_std)), None)
                    if _raw:
                        _pk_map[_std] = _raw
                _pk_std = list(_pk_map.keys())   # standard names with a raw counterpart
                _pk_raw = list(dict.fromkeys(_pk_map.values()))  # raw names, deduplicated

                _pivot_cache_key = f"{p}_table_pivot_cache"
                _source_mtime = (
                    os.path.getmtime(large_file_path)
                    if large_file_path and os.path.exists(large_file_path)
                    else 0
                )
                _pivot_sig = (
                    "pivot_cluster_count_v2",
                    str(large_file_path or ""),
                    float(_source_mtime),
                    str(_sel_dg_code or ""),
                    str(_sel_dg_name or ""),
                    str(_search_q_text or "") if _apply_source_search else "",
                    int(_MAX),
                    tuple(str(c) for c in _std_all),
                    tuple(str(c) for c in df_view.columns),
                    tuple(df_view.shape),
                    str(_pog_src_col or ""),
                    str(_sv_col or ""),
                    tuple(str(c) for c in _pk_raw),
                    tuple(str(c) for c in _dyn_pog_cols),
                )
                _pivot_cache = st.session_state.get(_pivot_cache_key)
                _pivot_scope = repr(_pivot_sig)
                if not (
                    isinstance(_pivot_cache, dict)
                    and _pivot_cache.get("sig") == _pivot_sig
                    and isinstance(_pivot_cache.get("tdf"), pd.DataFrame)
                ):
                    _disk_cache = _rs_load_table_base(_pivot_scope)
                    if (
                        isinstance(_disk_cache, dict)
                        and _disk_cache.get("sig") == _pivot_sig
                        and isinstance(_disk_cache.get("tdf"), pd.DataFrame)
                    ):
                        _pivot_cache = _disk_cache
                        st.session_state[_pivot_cache_key] = _disk_cache
                if (
                    isinstance(_pivot_cache, dict)
                    and _pivot_cache.get("sig") == _pivot_sig
                    and isinstance(_pivot_cache.get("tdf"), pd.DataFrame)
                ):
                    _tdf = _pivot_cache["tdf"].copy()
                    _piv_pog_cols = list(_pivot_cache.get("piv_pog_cols", []))
                    _cs_base_pog_counts = dict(_pivot_cache.get("cs_as_is_sku_counts", {}) or {})
                elif _dyn_pog_cols and _pk_raw:
                    # Pivot from FULL df_view: index=product key, columns=planogram, values=sales vol.
                    # aggfunc="sum" handles the case where a (product × planogram) pair appears in
                    # more than one row (e.g. multi-store long format).
                    _sv_series = (
                        pd.to_numeric(df_view[_sv_col], errors="coerce")
                        if _sv_col
                        else pd.Series([float("nan")] * len(df_view), index=df_view.index)
                    )
                    _piv = (
                        df_view
                        .assign(_pog_=_pog_vals_all, _sv_=_sv_series)
                        .pivot_table(index=_pk_raw, columns="_pog_", values="_sv_", aggfunc="sum")
                        .reset_index()
                    )
                    _piv.columns.name = None
                    _piv_pog_cols = [c for c in _piv.columns if c in set(_dyn_pog_cols)]

                    # Deduplicate _tdf to one row per product (take first occurrence of each attr)
                    _tdf_dedup = _tdf.groupby(_pk_std, sort=False).first().reset_index()

                    # Rename pivot's raw key columns to standard names, then merge
                    _piv_std = (
                        _piv.rename(columns=dict(zip(_pk_raw, _pk_std)))[_pk_std + _piv_pog_cols]
                    )
                    _cs_base_pog_counts = {
                        str(_pc): int(pd.to_numeric(_piv_std[_pc], errors="coerce").notna().sum())
                        for _pc in _piv_pog_cols
                    }
                    _tdf = _tdf_dedup.merge(_piv_std, on=_pk_std, how="left")

                    # Cap at _MAX products (not long rows) and reset index for clean row numbers
                    _tdf = _tdf.head(_MAX).reset_index(drop=True)

                    # Check Range To-be Waterfall = number of planograms this product appears in
                    _tdf["Check Range To-be Waterfall"] = _tdf[_piv_pog_cols].notna().sum(axis=1)
                    st.session_state[_pivot_cache_key] = {
                        "sig": _pivot_sig,
                        "tdf": _tdf.copy(),
                        "piv_pog_cols": list(_piv_pog_cols),
                        "cs_as_is_sku_counts": dict(_cs_base_pog_counts),
                    }
                    _rs_save_table_base(
                        _pivot_scope,
                        {
                            "sig": _pivot_sig,
                            "tdf": _tdf.copy(),
                            "piv_pog_cols": list(_piv_pog_cols),
                            "cs_as_is_sku_counts": dict(_cs_base_pog_counts),
                        },
                    )

                    # # ── TEMPORARY DEBUG EXPANDER ─────────────────────────────────
                    # with st.expander("🔎 debug — pivot diagnostics", expanded=False):

                    #     # ── 1. Column matching ──────────────────────────────────
                    #     st.markdown("**1 · Column matching**")
                    #     st.write({
                    #         "_pog_src_col (planogram name col)": _pog_src_col,
                    #         "_sv_col (sales VALUE col)":         _sv_col,
                    #         "_pk_map (std → raw key cols)":      _pk_map,
                    #         "_dyn_pog_cols (planogram labels)":  _dyn_pog_cols,
                    #     })

                    #     # ── 2. Pivot integrity ──────────────────────────────────
                    #     st.markdown("**2 · Pivot integrity**")
                    #     _n_unique_products = df_view[_pk_raw].drop_duplicates().shape[0]
                    #     st.write({
                    #         "df_view shape (long rows × cols)":     df_view.shape,
                    #         "unique products in df_view":            _n_unique_products,
                    #         "_piv shape (products × pog+key cols)": _piv.shape,
                    #     })
                    #     st.markdown("Pivot columns:")
                    #     st.write(list(_piv.columns))
                    #     st.markdown("_dyn_pog_cols (expected headers):")
                    #     st.write(_dyn_pog_cols)
                    #     st.markdown("_piv_pog_cols (intersection — values that actually landed in pivot):")
                    #     st.write(_piv_pog_cols)
                    #     _missing_in_pivot = [c for c in _dyn_pog_cols if c not in set(_piv.columns)]
                    #     if _missing_in_pivot:
                    #         st.warning(f"⚠️ These _dyn_pog_cols are NOT in the pivot columns "
                    #                    f"(name mismatch?): {_missing_in_pivot}")

                    #     # ── 3. Merge check ──────────────────────────────────────
                    #     st.markdown("**3 · Merge check (join key alignment)**")
                    #     # Rows where ALL pog columns are NaN → no pivot match
                    #     if _piv_pog_cols:
                    #         _blank_mask = _tdf[_piv_pog_cols].isna().all(axis=1)
                    #         _n_blank = int(_blank_mask.sum())
                    #         st.write(f"Products with ALL pog columns blank after merge: {_n_blank} / {len(_tdf)}")
                    #         if _n_blank > 0:
                    #             _blank_keys = _tdf.loc[_blank_mask, _pk_std].head(5)
                    #             st.markdown("Sample blank product keys in `_tdf` (after merge):")
                    #             st.dataframe(_blank_keys)
                    #             # Pivot index keys — compare dtypes / repr
                    #             _piv_keys_sample = _piv_std[_pk_std].head(5)
                    #             st.markdown("Sample pivot key rows in `_piv_std` (should match above):")
                    #             st.dataframe(_piv_keys_sample)
                    #             # Dtype comparison
                    #             _dtype_tdf = {c: str(_tdf[c].dtype) for c in _pk_std}
                    #             _dtype_piv = {c: str(_piv_std[c].dtype) for c in _pk_std}
                    #             st.write("_tdf key dtypes:", _dtype_tdf)
                    #             st.write("_piv_std key dtypes:", _dtype_piv)

                    #     # ── 4. Raw source check ─────────────────────────────────
                    #     st.markdown("**4 · Raw source check — df_view rows for blank products**")
                    #     if _piv_pog_cols:
                    #         _blank_mask2 = _tdf[_piv_pog_cols].isna().all(axis=1)
                    #         _sample_blank_stds = _tdf.loc[_blank_mask2, _pk_std].head(3)
                    #         if not _sample_blank_stds.empty:
                    #             for _, _brow in _sample_blank_stds.iterrows():
                    #                 # Build a filter against df_view using raw pk col names
                    #                 _filt = pd.Series([True] * len(df_view), index=df_view.index)
                    #                 for _s, _r in _pk_map.items():
                    #                     _filt &= (df_view[_r].astype(str).str.strip() == str(_brow[_s]).strip())
                    #                 _raw_rows = df_view.loc[_filt, _pk_raw + [_pog_src_col] + ([_sv_col] if _sv_col else [])]
                    #                 st.markdown(f"Raw rows for key `{tuple(_brow[c] for c in _pk_std)}`:")
                    #                 if _raw_rows.empty:
                    #                     st.warning("No matching rows found in df_view — key mismatch in the filter itself.")
                    #                 else:
                    #                     st.dataframe(_raw_rows)
                    #         else:
                    #             st.success("No blank-product rows found — all products matched the pivot.")

                    #     # ── 5. Pre-render cell comparison ───────────────────────
                    #     st.markdown("**5 · Pre-render cell comparison (pre-overlay, pre-reorder)**")
                    #     st.write(f"_MAX = {_MAX} | len(_tdf) = {len(_tdf)}")

                    #     # Any overlay data already in session state?
                    #     _dbg_ov_key  = f"{p}_pog_actions"
                    #     _dbg_ov_data = st.session_state.get(_dbg_ov_key, {})
                    #     st.write(f"_pog_actions entries in session state: {len(_dbg_ov_data)}")
                    #     if _dbg_ov_data:
                    #         st.write("Sample overlay entries (first 3):",
                    #                  dict(list(_dbg_ov_data.items())[:3]))

                    #     # Auto-pick 3 products that have the most non-NaN values in _piv
                    #     _dbg_nonnull_counts = _piv[_piv_pog_cols].notna().sum(axis=1)
                    #     _dbg_top_idx = _dbg_nonnull_counts.nlargest(3).index.tolist()

                    #     for _dbg_pi in _dbg_top_idx:
                    #         _dbg_prow  = _piv.loc[_dbg_pi]
                    #         _dbg_kstd  = {s: _dbg_prow[_pk_map[s]] for s in _pk_std}

                    #         # Locate product in _tdf by matching key values as strings
                    #         _dbg_mask = pd.Series([True] * len(_tdf), index=_tdf.index)
                    #         for _dbg_s in _pk_std:
                    #             _dbg_mask &= (
                    #                 _tdf[_dbg_s].astype(str).str.strip()
                    #                 == str(_dbg_kstd[_dbg_s]).strip()
                    #             )
                    #         _dbg_tmatch = _tdf[_dbg_mask]

                    #         _dbg_cap_flag = (
                    #             "no" if _dbg_tmatch.empty
                    #             else ("YES ⚠️" if _dbg_tmatch.index.min() >= _MAX else "no")
                    #         )
                    #         st.markdown(
                    #             f"**Product key:** `{_dbg_kstd}` — "
                    #             f"_tdf row(s): `{_dbg_tmatch.index.tolist()}` "
                    #             f"(row index ≥ _MAX={_MAX}? {_dbg_cap_flag})"
                    #         )

                    #         _dbg_comp = []
                    #         for _dbg_pc in _piv_pog_cols[:40]:
                    #             _piv_val = _dbg_prow.get(_dbg_pc, float("nan"))
                    #             if _dbg_tmatch.empty:
                    #                 _tdf_val  = "PRODUCT NOT IN _tdf"
                    #                 _fmt      = ""
                    #             elif _dbg_pc not in _dbg_tmatch.columns:
                    #                 _tdf_val  = "COLUMN MISSING"
                    #                 _fmt      = ""
                    #             else:
                    #                 _tdf_val = _dbg_tmatch.iloc[0][_dbg_pc]
                    #                 try:
                    #                     _fmt = (f"{int(float(_tdf_val)):,}"
                    #                             if pd.notna(_tdf_val) else "(blank)")
                    #                 except (TypeError, ValueError):
                    #                     _fmt = f"(err: {_tdf_val!r})"
                    #             _dbg_comp.append({
                    #                 "planogram col": _dbg_pc,
                    #                 "_piv value":    _piv_val,
                    #                 "_tdf value":    _tdf_val,
                    #                 "type(_tdf)":    type(_tdf_val).__name__,
                    #                 "pd.notna":      pd.notna(_tdf_val)
                    #                                  if not isinstance(_tdf_val, str)
                    #                                  else True,
                    #                 "render output": _fmt,
                    #             })
                    #         st.dataframe(pd.DataFrame(_dbg_comp), height=300, use_container_width=True)
                    # # ── END TEMPORARY DEBUG EXPANDER ─────────────────────────────

                else:
                    _cs_base_pog_counts = {}
                    _tdf["Check Range To-be Waterfall"] = ""

            # Column ordering: sticky left | data cols | Status | planogram cols rightmost
            _STICKY      = [c for c in ["DG Code", "ID", "Item Name"] if c in _tdf.columns]
            _dyn_pog_set = set(_dyn_pog_cols)
            _LAST        = [c for c in ["Status", "Check Range To-be Waterfall", "Planogram Name"] if c in _tdf.columns] + _dyn_pog_cols
            _REST        = [c for c in _tdf.columns if c not in _STICKY and c not in set(_LAST)]
            _tdf         = _tdf[_STICKY + _REST + _LAST]
            if "Status" in _tdf.columns:
                _tdf["Status"] = "MAINTAIN"

            # ── Persistent planogram-edit overlay ────────────────────────────────
            # Key = (DG Code, ID, Item Name) — after pivot, one row per product so
            # the sticky columns alone are unique.
            _pog_edits_key = f"{p}_pog_edits"
            _data_edits_key = f"{p}_data_edits"

            def _make_rk(ri: int) -> tuple:
                # After pivot _tdf has one row per product — _STICKY alone is unique
                return tuple(str(_tdf.at[ri, c]) for c in _STICKY)

            _ROW_KEY_COL = "__rs_row_key"
            _NEW_ROW_COL = "__rs_added_row"
            _EDIT_COL = "__rs_last_edit"
            _tdf[_ROW_KEY_COL] = [
                _json.dumps(_make_rk(_ri), ensure_ascii=False)
                for _ri in range(len(_tdf))
            ]
            _tdf[_NEW_ROW_COL] = False
            _tdf[_EDIT_COL] = ""
            _cs_as_is_sku_counts: dict[str, int] = {}
            _cs_original_present: dict[str, set] = {}
            for _pc in _dyn_pog_cols:
                if _pc in _tdf.columns:
                    _num_s = pd.to_numeric(_tdf[_pc], errors="coerce")
                    _present = set(_tdf.index[_num_s.notna()].tolist())
                    _cs_original_present[_pc] = _present
                    _cs_as_is_sku_counts[_pc] = int(
                        _cs_base_pog_counts.get(str(_pc), len(_present))
                    )
            for _pc in _dyn_pog_cols:
                if _pc in _tdf.columns:
                    _tdf[_pc] = _tdf[_pc].astype(object)

            def _row_to_rk(row) -> tuple:
                raw_key = row.get(_ROW_KEY_COL, "")
                if raw_key not in ("", None) and not (
                    isinstance(raw_key, float) and pd.isna(raw_key)
                ):
                    try:
                        parsed = _json.loads(str(raw_key))
                        if isinstance(parsed, list):
                            return tuple(str(x) for x in parsed)
                    except Exception:
                        pass
                return tuple(str(row.get(c, "")) for c in _STICKY)

            _pog_edits = st.session_state.setdefault(_pog_edits_key, {})
            _data_edits = st.session_state.setdefault(_data_edits_key, {})
            _pog_actions_live_for_originals = st.session_state.get(f"{p}_pog_actions", {})
            _status_overrides_live_for_originals = st.session_state.get(f"{p}_status_overrides", {})
            _orig_action_values = {}
            _orig_action_rows = set(_pog_edits.keys()) | set(_pog_actions_live_for_originals.keys())
            _orig_action_rows |= {
                _rk0 for _rk0, _st0 in (_status_overrides_live_for_originals or {}).items()
                if str(_st0).strip().upper() == "DELETE ALL"
            }
            _orig_action_cells = set()
            for _store in (_pog_edits, _pog_actions_live_for_originals):
                for _rk0, _acts0 in (_store or {}).items():
                    if isinstance(_acts0, dict):
                        for _pc0, _act0 in _acts0.items():
                            if _act0 in ("Delete", "New"):
                                _orig_action_cells.add((_rk0, _pc0))
            if _orig_action_rows or _orig_action_cells:
                _orig_cells_by_rk = {}
                for _rk_cell, _pc0 in _orig_action_cells:
                    _orig_cells_by_rk.setdefault(_rk_cell, set()).add(_pc0)
                for _ri0 in range(len(_tdf)):
                    _rk0 = _make_rk(_ri0)
                    _restore_cols = (
                        set(_dyn_pog_cols)
                        if _rk0 in _orig_action_rows
                        else set(_orig_cells_by_rk.get(_rk0, ()))
                    )
                    for _pc0 in _restore_cols:
                        if _pc0 not in _tdf.columns:
                            continue
                        _ov0 = _tdf.at[_ri0, _pc0]
                        if _ov0 is None or (isinstance(_ov0, float) and pd.isna(_ov0)):
                            _ov0 = ""
                        _orig_action_values[
                            f"{_tdf.at[_ri0, _ROW_KEY_COL]}\u241f{_pc0}"
                        ] = "" if str(_ov0) in ("nan", "None") else str(_ov0)
                        _rk_join = "\u241e".join(str(_x) for _x in _rk0)
                        _orig_action_values[
                            f"{_rk_join}\u241f{_pc0}"
                        ] = "" if str(_ov0) in ("nan", "None") else str(_ov0)
            if _pog_edits and _dyn_pog_cols and _STICKY:
                for _ri in range(len(_tdf)):
                    _rk = _make_rk(_ri)
                    if _rk in _pog_edits:
                        for _pc, _pv in _pog_edits[_rk].items():
                            if _pc in _tdf.columns:
                                _tdf.at[_ri, _pc] = _pv

            if "Check Range To-be Waterfall" in _tdf.columns and _dyn_pog_cols:
                _pog_cols_in_tdf = [c for c in _dyn_pog_cols if c in _tdf.columns]
                if _pog_cols_in_tdf:
                    _pog_actions_live = st.session_state.get(f"{p}_pog_actions", {})
                    _status_overrides_live = st.session_state.get(f"{p}_status_overrides", {})
                    _waterfall_counts = []
                    for _ri in range(len(_tdf)):
                        _rk = _make_rk(_ri)
                        _count = sum(
                            1 for _pc in _pog_cols_in_tdf
                            if _ri in _cs_original_present.get(_pc, set())
                        )
                        if str(_status_overrides_live.get(_rk, "")).strip().upper() == "DELETE ALL":
                            _waterfall_counts.append(0)
                            continue
                        _acts = {}
                        _acts.update(_pog_edits.get(_rk, {}))
                        _acts.update(_pog_actions_live.get(_rk, {}))
                        for _pc, _act in _acts.items():
                            if _pc not in _pog_cols_in_tdf:
                                continue
                            _was_present = _ri in _cs_original_present.get(_pc, set())
                            _act_s = str(_act).strip().lower()
                            if _act_s == "delete" and _was_present:
                                _count -= 1
                            elif _act_s == "new" and not _was_present:
                                _count += 1
                        _waterfall_counts.append(max(0, _count))
                    _tdf["Check Range To-be Waterfall"] = _waterfall_counts

            # # ── TEMPORARY: post-overlay debug ────────────────────────────────────
            # if _dyn_pog_cols:
            #     with st.expander("🔎 debug — post-overlay / render-ready state", expanded=False):
            #         st.markdown("**Values in `_tdf` AFTER overlay applied, AFTER column reorder — "
            #                     "this is exactly what the render loop sees.**")
            #         # Pick first 3 rows of _tdf; show all pog columns
            #         _pog_in_tdf = [c for c in _dyn_pog_cols if c in _tdf.columns]
            #         _dbg_act_ov = st.session_state.get(f"{p}_pog_actions", {})
            #         st.write(f"_dyn_pog_cols count: {len(_dyn_pog_cols)} | "
            #                  f"pog cols present in _tdf: {len(_pog_in_tdf)} | "
            #                  f"action overlay entries: {len(_dbg_act_ov)}")

            #         for _dbg_ri in range(min(3, len(_tdf))):
            #             _dbg_rk = _make_rk(_dbg_ri)
            #             _ov_for_row = _dbg_act_ov.get(_dbg_rk, {})
            #             st.markdown(f"**Row {_dbg_ri}** | key={_dbg_rk} | "
            #                         f"overlay entries for this key: {_ov_for_row}")
            #             _post_rows = []
            #             for _pc in _pog_in_tdf[:40]:
            #                 _raw = _tdf.at[_dbg_ri, _pc]
            #                 try:
            #                     _fmt = (f"{int(float(_raw)):,}"
            #                             if pd.notna(_raw) else "(blank)")
            #                 except (TypeError, ValueError):
            #                     _fmt = f"(err: {_raw!r})"
            #                 _post_rows.append({
            #                     "planogram col": _pc,
            #                     "raw value":     _raw,
            #                     "type":          type(_raw).__name__,
            #                     "pd.notna":      pd.notna(_raw)
            #                                      if not isinstance(_raw, str) else True,
            #                     "render output": _fmt,
            #                     "wiped by overlay?": (
            #                         "YES ⚠️" if _pc in _ov_for_row
            #                         and pd.isna(_ov_for_row[_pc]) else "no"
            #                     ),
            #                 })
            #             st.dataframe(pd.DataFrame(_post_rows), height=300, use_container_width=True)
            # # ── END TEMPORARY post-overlay debug ─────────────────────────────────

            # # ── TEMPORARY: automated reconciliation (proves pivot ≡ raw data) ──
            # if _piv_pog_cols and _sv_col and _pog_src_col:
            #     with st.expander("🔍 reconciliation — pivot vs raw data", expanded=False):
            #         _SEP = "\x00\x01"   # separator unlikely to appear in product IDs

            #         # ── Pre-compute shared key series (used by all 4 checks) ──
            #         # pk string key in df_view (raw col names)
            #         _rc_dfv_pk = df_view[_pk_raw[0]].astype(str)
            #         for _c in _pk_raw[1:]:
            #             _rc_dfv_pk = _rc_dfv_pk + _SEP + df_view[_c].astype(str)

            #         # pk string key in _tdf (standard col names, same join order)
            #         _rc_tdf_pk = _tdf[_pk_std[0]].astype(str)
            #         for _c in _pk_std[1:]:
            #             _rc_tdf_pk = _rc_tdf_pk + _SEP + _tdf[_c].astype(str)

            #         _rc_displayed_pks = set(_rc_tdf_pk)                          # products in table
            #         _rc_pog_str       = df_view[_pog_src_col].astype(str).str.strip()
            #         _rc_sv_num        = pd.to_numeric(df_view[_sv_col], errors="coerce")
            #         _rc_pk_mask       = _rc_dfv_pk.isin(_rc_displayed_pks)        # rows for displayed products
            #         _rc_pog_mask      = _rc_pog_str.isin(set(_dyn_pog_cols))      # rows for displayed planograms

            #         # ── Check 1: Grand total ────────────────────────────────────
            #         st.markdown("#### Check 1 — Grand total")
            #         _c1_tdf = float(_tdf[_piv_pog_cols].sum(skipna=True).sum())
            #         _c1_raw = float(_rc_sv_num[_rc_pk_mask & _rc_pog_mask].sum())
            #         _c1_diff = abs(_c1_tdf - _c1_raw)
            #         _c1_ok   = _c1_diff < 0.5
            #         st.write({
            #             "Pivot _tdf total (all pog cols, all displayed products)":
            #                 f"{_c1_tdf:,.2f}",
            #             "Raw df_view total (same products + planograms)":
            #                 f"{_c1_raw:,.2f}",
            #             "Absolute difference":
            #                 f"{_c1_diff:.4f}",
            #         })
            #         if _c1_ok:
            #             st.success("✓ Grand totals match")
            #         else:
            #             st.error(f"✗ Grand total mismatch — diff = {_c1_diff:,.2f}")

            #         # ── Check 2: Cell spot-check (8 random non-NaN cells) ───────
            #         st.markdown("#### Check 2 — Cell spot-check (8 random cells)")
            #         _c2_melt = (
            #             _tdf[_pk_std + _piv_pog_cols]
            #             .melt(id_vars=_pk_std, var_name="_pog_", value_name="_val_")
            #             .dropna(subset=["_val_"])
            #             .reset_index(drop=True)
            #         )
            #         _c2_sample = (
            #             _c2_melt.sample(min(8, len(_c2_melt)), random_state=42)
            #             if not _c2_melt.empty else _c2_melt
            #         )
            #         _c2_rows = []
            #         for _, _sr in _c2_sample.iterrows():
            #             _pog_lbl  = _sr["_pog_"]
            #             _tdf_val  = float(_sr["_val_"])
            #             _pk_key   = _SEP.join(str(_sr[c]) for c in _pk_std)
            #             _c2_filt  = (_rc_dfv_pk == _pk_key) & (_rc_pog_str == _pog_lbl)
            #             _raw_sum  = float(_rc_sv_num[_c2_filt].sum())
            #             _c2_rows.append({
            #                 **{c: _sr[c] for c in _pk_std},
            #                 "planogram":  _pog_lbl,
            #                 "raw_sum":    f"{_raw_sum:,.0f}",
            #                 "tdf_value":  f"{_tdf_val:,.0f}",
            #                 "match?":     "✓" if abs(_raw_sum - _tdf_val) < 0.5 else "✗",
            #             })
            #         _c2_df = pd.DataFrame(_c2_rows)
            #         st.dataframe(_c2_df, hide_index=True, use_container_width=True)
            #         if _c2_df.empty or (_c2_df["match?"] == "✓").all():
            #             st.success("✓ All sampled cells match")
            #         else:
            #             _c2_bad = int((~(_c2_df["match?"] == "✓")).sum())
            #             st.error(f"✗ {_c2_bad} cell(s) mismatch")

            #         # ── Check 3: Marginals (5 products, row-total) ───────────────
            #         st.markdown("#### Check 3 — Marginals (5 sample products, row-total)")
            #         _c3_sample = _tdf.sample(min(5, len(_tdf)), random_state=7)
            #         _c3_rows = []
            #         for _, _mr in _c3_sample.iterrows():
            #             _pivot_tot = sum(
            #                 float(_mr[c]) for c in _piv_pog_cols
            #                 if c in _mr.index and pd.notna(_mr[c])
            #             )
            #             _pk_key  = _SEP.join(str(_mr[c]) for c in _pk_std)
            #             _c3_filt = (_rc_dfv_pk == _pk_key) & _rc_pog_mask
            #             _raw_tot = float(_rc_sv_num[_c3_filt].sum())
            #             _c3_rows.append({
            #                 **{c: _mr[c] for c in _pk_std},
            #                 "tdf row-total":    f"{_pivot_tot:,.0f}",
            #                 "raw df_view sum":  f"{_raw_tot:,.0f}",
            #                 "match?": "✓" if abs(_pivot_tot - _raw_tot) < 0.5 else "✗",
            #             })
            #         _c3_df = pd.DataFrame(_c3_rows)
            #         st.dataframe(_c3_df, hide_index=True, use_container_width=True)
            #         if _c3_df.empty or (_c3_df["match?"] == "✓").all():
            #             st.success("✓ All row-totals match")
            #         else:
            #             st.error(f"✗ {int((~(_c3_df['match?'] == '✓')).sum())} row-total(s) mismatch")

            #         # ── Check 4: aggfunc sanity ─────────────────────────────────
            #         st.markdown("#### Check 4 — aggfunc sanity (repeated product×planogram pairs?)")
            #         _c4_df = (
            #             df_view[_rc_pk_mask & _rc_pog_mask]
            #             .assign(_pk_=_rc_dfv_pk, _pog_=_rc_pog_str)
            #             .groupby(["_pk_", "_pog_"], sort=False)
            #             .size()
            #             .reset_index(name="_count_")
            #         )
            #         _c4_repeated = _c4_df[_c4_df["_count_"] > 1]
            #         st.write(
            #             f"Total (product, planogram) pairs in displayed subset: "
            #             f"**{len(_c4_df):,}** | "
            #             f"Pairs appearing >1 time: **{len(_c4_repeated):,}**"
            #         )
            #         if _c4_repeated.empty:
            #             st.success("✓ All pairs unique — aggfunc='sum' ≡ first (no multi-row summing)")
            #         else:
            #             st.warning(
            #                 f"⚠️ {len(_c4_repeated):,} pairs repeat. "
            #                 f"aggfunc='sum' is adding their rows. "
            #                 f"Confirm this is intended (e.g. multi-store rows)."
            #             )
            #             for _, _rr in _c4_repeated.head(3).iterrows():
            #                 _c4_filt = (_rc_dfv_pk == _rr["_pk_"]) & (_rc_pog_str == _rr["_pog_"])
            #                 _c4_detail = df_view.loc[_c4_filt, _pk_raw + [_pog_src_col, _sv_col]].copy()
            #                 _c4_detail["→ summed_to"] = (
            #                     pd.to_numeric(_c4_detail[_sv_col], errors="coerce").sum()
            #                 )
            #                 st.markdown(
            #                     f"**pk=`{_rr['_pk_']}`  pog=`{_rr['_pog_']}`** "
            #                     f"({int(_rr['_count_'])} rows)"
            #                 )
            #                 st.dataframe(_c4_detail, hide_index=True)
            # # ── END TEMPORARY reconciliation ──────────────────────────────────

            # ── Layout helpers ─────────────────────────────────────────────────────
            # Column visibility is managed entirely by the AG Grid sidebar + header menu.
            # State (hide/width/pin) persisted in session state survives DG changes; colIds
            # for columns that no longer exist in the new DG are silently ignored by AG Grid.
            _STICKY      = [c for c in ["DG Code", "ID", "Item Name"] if c in _tdf.columns]
            _dyn_pog_set = set(c for c in _dyn_pog_cols if c in _tdf.columns)
            _pinned_rows: list = []

            # ── Cluster Summary Panel — pinned above the planogram grid ───────────
            if _dyn_pog_cols:
                _CS_ROWS = [
                    ("TO-BE Stores applied count", "#F4A460", "#000000"),
                    ("AS-IS Stores applied count",  "#F4A460", "#000000"),
                    ("MODs",                         "#DCDCDC", "#000000"),
                    ("FIXTURE",                      "#DCDCDC", "#000000"),
                    ("RANGE CLASS",                  "#DCDCDC", "#000000"),
                    ("Total NEW SKUs",               "#808080", "#FFFFFF"),
                    ("Total DELETE SKUs",            "#808080", "#FFFFFF"),
                    ("TO-BE SKUs count",             "#808080", "#FFFFFF"),
                    ("AS-IS SKUs count",             "#F0F0F0", "#1565C0"),
                    ("%Achieving CRD case (AS is)",  "#FFFFFF", "#000000"),
                    ("%Achieving LRD SALES (AS is)", "#FFFFFF", "#000000"),
                ]

                # ── Load A5 file: cached, column-pruned, and filtered to visible POGs ──
                _a5_warn = None
                _a5 = None
                _a5_meta = {}
                for _am in load_admin_manifest():
                    _amp = os.path.join(BASE_DIR, "uploads", _am["name"])
                    if (os.path.exists(_amp)
                            and "a5" in _am["name"].lower()
                            and not is_large_file(_amp)):
                        _a5, _a5_meta, _a5_warn = _load_a5_subset(
                            _amp,
                            os.path.getmtime(_amp),
                            tuple(str(c) for c in _dyn_pog_cols),
                        )
                        break

                # ── Locate columns in A5 ─────────────────────────────────────────
                _a5_pog_c = _a5_cluster_c = _a5_mod_c = _a5_fix_c = _a5_rng_c = None
                if _a5 is not None:
                    _a5_pog_c = _a5_meta.get("pog")
                    _a5_cluster_c = _a5_meta.get("cluster")
                    _a5_mod_c = _a5_meta.get("mod")
                    _a5_fix_c = _a5_meta.get("fixture")
                    _a5_rng_c = _a5_meta.get("range")

                # ── Build pog-name → compact A5 lookup ───────────────────────────
                _a5_lkp: dict = {}
                _a5_pair_lkp: dict = {}
                _a5_pog_records: list[dict] = []
                _pog_to_cl: dict = {}
                if _a5 is not None and _a5_pog_c:
                    _cols_for_records = [c for c in (
                        _a5_pog_c, _a5_cluster_c, _a5_mod_c, _a5_fix_c, _a5_rng_c
                    ) if c and c in _a5.columns]
                    for _rec in _a5[_cols_for_records].drop_duplicates().to_dict("records"):
                        _ap = str(_rec.get(_a5_pog_c, "")).strip()
                        if _ap in ("", "nan", "None"):
                            continue
                        _key = _nca(_ap)
                        _cl = str(_rec.get(_a5_cluster_c, "")).strip() if _a5_cluster_c else ""
                        _row_vals = {
                            "mod": _rec.get(_a5_mod_c) if _a5_mod_c else "",
                            "fixture": _rec.get(_a5_fix_c) if _a5_fix_c else "",
                            "range": _rec.get(_a5_rng_c) if _a5_rng_c else "",
                            "cluster": _cl,
                        }
                        _a5_pog_records.append({
                            "pog": _ap,
                            "key": _key,
                            "cluster": _cl,
                        })
                        _a5_lkp[_key] = _row_vals
                        if _cl not in ("", "nan", "None"):
                            _a5_pair_lkp[(_key, _nca(_cl))] = _row_vals
                            _pog_to_cl[_ap] = _cl
                            _pog_to_cl[_key] = _cl

                    def _pog_target_token(_v) -> str:
                        _m = _re.search(r"(?:target[_\-\s]*|_)(\d{4,})", str(_v or ""), flags=_re.I)
                        return _m.group(1) if _m else ""

                    # Main table planogram names can be longer than A5 planogram names
                    # (for example include CRACKER/.../target_5012). Fill the exact
                    # displayed column -> cluster header map by matching normalized
                    # names and target codes, so report pages do not fall back to
                    # "Unspecified".
                    for _dp in _dyn_pog_cols:
                        _dp_s = str(_dp)
                        _dp_key = _nca(_dp_s)
                        if _pog_to_cl.get(_dp_s) or _pog_to_cl.get(_dp_key):
                            continue
                        _dp_target = _pog_target_token(_dp_s)
                        _best_cl = ""
                        for _r in _a5_pog_records:
                            _rk = str(_r.get("key", ""))
                            if _rk and (_rk in _dp_key or _dp_key in _rk):
                                _best_cl = str(_r.get("cluster", "")).strip()
                                break
                        if not _best_cl and _dp_target:
                            for _r in _a5_pog_records:
                                _rk = str(_r.get("key", ""))
                                if _dp_target and _dp_target in _rk:
                                    _best_cl = str(_r.get("cluster", "")).strip()
                                    break
                        if _best_cl and _best_cl not in ("nan", "None"):
                            _pog_to_cl[_dp_s] = _best_cl
                            _pog_to_cl[_dp_key] = _best_cl

                # ── Cell-value resolver ───────────────────────────────────────────
                _cs_actions_live = st.session_state.get(f"{p}_pog_actions", {})
                _cs_new_counts = {str(c): 0 for c in _dyn_pog_cols}
                _cs_delete_counts = {str(c): 0 for c in _dyn_pog_cols}
                _cs_to_be_counts = {
                    str(c): int(_cs_as_is_sku_counts.get(c, 0))
                    for c in _dyn_pog_cols
                }
                _rk_to_idx = {_make_rk(_ari): _ari for _ari in range(len(_tdf))}
                _cs_status_live = st.session_state.get(f"{p}_status_overrides", {})
                _cs_delete_all_rks = set()
                for _srk, _sst in (_cs_status_live or {}).items():
                    if str(_sst).strip().upper() != "DELETE ALL":
                        continue
                    _ari = _rk_to_idx.get(_srk)
                    if _ari is None:
                        continue
                    _cs_delete_all_rks.add(_srk)
                    for _pc in _dyn_pog_cols:
                        _pc_s = str(_pc)
                        if _ari in _cs_original_present.get(_pc, set()):
                            _cs_delete_counts[_pc_s] = _cs_delete_counts.get(_pc_s, 0) + 1
                            _cs_to_be_counts[_pc_s] = max(
                                0, _cs_to_be_counts.get(_pc_s, 0) - 1
                            )
                _all_action_keys = set(_cs_actions_live) | set(_pog_edits)
                for _ark in _all_action_keys:
                    if _ark in _cs_delete_all_rks:
                        continue
                    _ari = _rk_to_idx.get(_ark)
                    if _ari is None:
                        continue
                    _merged_acts = {}
                    _merged_acts.update(_pog_edits.get(_ark, {}))
                    _merged_acts.update(_cs_actions_live.get(_ark, {}))
                    for _pc, _act in _merged_acts.items():
                        _pc_s = str(_pc)
                        _act_s = str(_act).strip().lower()
                        _was_present = _ari in _cs_original_present.get(_pc, set())
                        if _act_s == "new":
                            _cs_new_counts[_pc_s] = _cs_new_counts.get(_pc_s, 0) + 1
                            if not _was_present:
                                _cs_to_be_counts[_pc_s] = _cs_to_be_counts.get(_pc_s, 0) + 1
                        elif _act_s == "delete":
                            _cs_delete_counts[_pc_s] = _cs_delete_counts.get(_pc_s, 0) + 1
                            if _was_present:
                                _cs_to_be_counts[_pc_s] = max(
                                    0, _cs_to_be_counts.get(_pc_s, 0) - 1
                                )

                def _cs_val(row_label: str, pog_col: str, cluster_name: str = "") -> str:
                    _pog_col_s = str(pog_col)
                    _pog_key = _nca(pog_col)
                    _cl_key = _nca(cluster_name)
                    _a5r = (
                        _a5_pair_lkp.get((_pog_key, _cl_key))
                        if _cl_key else None
                    ) or _a5_lkp.get(_pog_key)
                    if row_label == "TO-BE Stores applied count":
                        return "1"
                    if row_label == "AS-IS Stores applied count":
                        return "1"
                    if row_label == "Total NEW SKUs":
                        return str(_cs_new_counts.get(_pog_col_s, 0))
                    if row_label == "Total DELETE SKUs":
                        return str(_cs_delete_counts.get(_pog_col_s, 0))
                    if row_label == "TO-BE SKUs count":
                        return str(_cs_to_be_counts.get(_pog_col_s, 0))
                    if row_label == "AS-IS SKUs count":
                        return str(_cs_as_is_sku_counts.get(pog_col, 0))
                    if row_label == "MODs" and _a5_mod_c:
                        _v = _a5r.get("mod") if _a5r is not None else ""
                        return str(_v) if pd.notna(_v) else ""
                    if row_label == "FIXTURE" and _a5_fix_c:
                        _v = _a5r.get("fixture") if _a5r is not None else ""
                        return str(_v) if pd.notna(_v) else ""
                    if row_label == "RANGE CLASS" and _a5_rng_c:
                        _v = _a5r.get("range") if _a5r is not None else ""
                        return str(_v) if pd.notna(_v) else ""
                    return ""

                # Cluster data source: session state (Cluster tab edits) or auto-fill
                _ct_cls_disp = list(st.session_state.get(f"{p}_ct_cls") or [])
                if not _ct_cls_disp:
                    _ct_cls_disp = sorted(set(_pog_to_cl.values()))
                _ct_dat_disp = st.session_state.get(f"{p}_ct_dat", {})

                # ── HTML cluster summary table above the grid ─────────────────────
                # The top table is split into two zones to match AG Grid:
                #   1. Pinned zone  (matches AG Grid pinned-left panel — doesn't scroll)
                #   2. Body zone    (matches AG Grid body — scrolls in sync)
                # _pinned_w: width of sticky columns (AG Grid pinned-left panel)
                # _rest_w:   width of body columns before Status (REST cols)
                _last_set2   = {"Status", "Check Range To-be Waterfall",
                                "Planogram Name"} | set(_dyn_pog_cols)
                _sticky_wmap = {"DG Code": 86, "ID": 120, "Item Name": 210}
                _pinned_w    = sum(_sticky_wmap.get(c, 110) for c in _STICKY)
                _rest_w      = 110 * len([c for c in _tdf.columns
                                          if c not in set(_STICKY) and c not in _last_set2])
                # AG Grid balham-matched styles
                _LW = 260   # Status + Check Range To-be Waterfall default width
                _CW = 72    # match POG column width in AG Grid
                _BRD    = "1px solid #BDC3C7"
                _HDR_BG = "#f5f7f7"
                _BASE   = (f"border-right:{_BRD};border-bottom:{_BRD};"
                           "padding:4px 8px;font-size:12px;white-space:nowrap;"
                           "overflow:hidden;text-overflow:ellipsis;")
                _td_s   = _BASE
                _th_s   = (_BASE + f"background:{_HDR_BG};font-weight:700;"
                           "text-align:center;color:#222;"
                           f"border-top:{_BRD};border-left:{_BRD};")

                # header row — one column per planogram (aligns with AG Grid body)
                # Each column header shows the cluster name for that planogram.
                _ct_hdr  = (f'<th class="cs-label-col" style="{_th_s}min-width:{_LW}px;'
                             f'width:{_LW}px;max-width:{_LW}px;">'
                             f'Cluster</th>')
                for _pi, _pog in enumerate(_dyn_pog_cols):
                    _cl_lbl = _pog_to_cl.get(_pog) or _pog_to_cl.get(_nca(_pog), "")
                    _ct_hdr += (f'<th class="cs-pog-cell" data-idx="{_pi}" '
                                f'style="{_th_s}min-width:{_CW}px;width:{_CW}px;'
                                f'max-width:{_CW}px;height:128px;padding:2px 4px;'
                                f'text-overflow:clip;" title="{_cl_lbl}">'
                                f'<div style="height:124px;display:flex;align-items:center;'
                                f'justify-content:center;writing-mode:vertical-rl;'
                                f'transform:rotate(180deg);white-space:nowrap;'
                                f'overflow:visible;text-overflow:clip;">{_cl_lbl}</div></th>')

                # data rows — one cell per planogram, value = cluster metric for that planogram
                _ct_body = ""
                for _ri2, (_rl, _rbg, _rtc) in enumerate(_CS_ROWS):
                    _even   = _ri2 % 2 == 0
                    _row_bg = "#ffffff" if _even else "#f9f9f9"
                    _lbl_s  = (_td_s + f"font-weight:700;color:#222;"
                               f"background:{_row_bg};border-left:{_BRD};")
                    _tds  = (f'<td class="cs-label-col" style="{_lbl_s}min-width:{_LW}px;'
                             f'width:{_LW}px;max-width:{_LW}px;">{_rl}</td>')
                    for _pi, _pog in enumerate(_dyn_pog_cols):
                        _cl_key = _pog_to_cl.get(_pog) or _pog_to_cl.get(_nca(_pog), "")
                        if _rl in (
                            "MODs", "FIXTURE", "RANGE CLASS",
                            "Total NEW SKUs", "Total DELETE SKUs",
                            "TO-BE SKUs count", "AS-IS SKUs count",
                        ):
                            _v = _cs_val(_rl, _pog, _cl_key)
                        else:
                            _ssv = str(_ct_dat_disp.get(_rl, {}).get(_cl_key, "") or "")
                            if _ssv in ("nan", "None"): _ssv = ""
                            _v = _ssv if _ssv else _cs_val(_rl, _pog, _cl_key)
                        _dat_s = (_td_s + f"text-align:center;color:#222;"
                                  f"background:{_row_bg};")
                        if _rl == "FIXTURE" and str(_v).strip():
                            _tds += (f'<td class="cs-pog-cell" data-idx="{_pi}" '
                                     f'style="{_dat_s}min-width:{_CW}px;width:{_CW}px;'
                                     f'max-width:{_CW}px;height:112px;padding:2px 4px;'
                                     f'text-overflow:clip;">'
                                     f'<div style="height:108px;display:flex;align-items:center;'
                                     f'justify-content:center;writing-mode:vertical-rl;'
                                     f'transform:rotate(180deg);white-space:nowrap;'
                                     f'overflow:visible;text-overflow:clip;">{_v}</div></td>')
                        else:
                            _tds += (f'<td class="cs-pog-cell" data-idx="{_pi}" '
                                     f'style="{_dat_s}min-width:{_CW}px;width:{_CW}px;'
                                     f'max-width:{_CW}px;">'
                                     f'{_v}</td>')
                    _ct_body += f'<tr>{_tds}</tr>'

                _BRIDGE_H = 178
                _bridge_cells = (
                    f'<td class="cs-label-col" style="{_td_s}min-width:{_LW}px;'
                    f'width:{_LW}px;max-width:{_LW}px;height:{_BRIDGE_H}px;'
                    f'background:transparent;border-left:{_BRD};border-bottom:0;'
                    f'border-top:0;"></td>'
                )
                for _pi, _pog in enumerate(_dyn_pog_cols):
                    _bridge_cells += (
                        f'<td class="cs-pog-cell" data-idx="{_pi}" '
                        f'style="{_td_s}min-width:{_CW}px;width:{_CW}px;'
                        f'max-width:{_CW}px;height:{_BRIDGE_H}px;'
                        f'background:transparent;border-bottom:0;border-top:0;"></td>'
                    )
                _ct_bridge = (
                    '<table class="cs-bridge-table" '
                    'style="border-collapse:collapse;table-layout:fixed;'
                    'pointer-events:none;position:absolute;left:0;top:100%;'
                    'z-index:30;">'
                    f'<tbody><tr>{_bridge_cells}</tr></tbody></table>'
                )

                if _a5_warn:
                    st.caption(f"⚠️ A5 load error: {_a5_warn} — MODs/FIXTURE/RANGE CLASS blank.")
                elif _a5 is None:
                    st.caption("ℹ️ No A5 file pinned — MODs, FIXTURE, RANGE CLASS blank.")
                elif _a5_pog_c is None:
                    st.caption("⚠️ A5 loaded but planogram-name column not found — "
                               "MODs/FIXTURE/RANGE CLASS blank.")

                # Wrapper: pinned cover on the left (static) + scrollable body on the right
                st.markdown(
                    '<style>.cs-scroll-sync::-webkit-scrollbar{height:6px}'
                    '.cs-scroll-sync::-webkit-scrollbar-track{background:#f1f1f1}'
                    '.cs-scroll-sync::-webkit-scrollbar-thumb{background:#BDC3C7;border-radius:3px}'
                    '.cs-label-col{position:sticky;left:0;z-index:7;}'
                    'thead .cs-label-col{z-index:9;}'
                    '</style>'
                    '<div style="position:relative;margin-bottom:0;z-index:30;">'
                    # pinned cover: blank white area matching AG Grid pinned-left panel
                    f'<div class="cs-pin-cover" style="position:absolute;left:0;top:0;'
                    f'width:{_pinned_w}px;height:100%;background:#fff;z-index:5;'
                    f'border-right:2px solid #BDC3C7;box-sizing:border-box;"></div>'
                    # scrollable body zone (starts at _pinned_w, matches AG Grid body)
                    f'<div class="cs-scroll-sync" style="overflow-x:hidden;overflow-y:visible;'
                    f'position:relative;margin-bottom:0;'
                    f'margin-left:{_pinned_w + _rest_w}px;">'
                    '<table style="border-collapse:collapse;table-layout:fixed;">'
                    f'<thead><tr>{_ct_hdr}</tr></thead>'
                    f'<tbody>{_ct_body}</tbody>'
                    f'</table>{_ct_bridge}</div></div>',
                    unsafe_allow_html=True,
                )

            # ── Single AG Grid — always active, no view/edit toggle ────────────────
            try:
                from st_aggrid import (
                    AgGrid, DataReturnMode, GridOptionsBuilder, GridUpdateMode,
                    JsCode,
                )
                _AGGRID_OK = True
            except ImportError:
                _AGGRID_OK = False

            if not _AGGRID_OK:
                st.error("⚠️ `streamlit-aggrid` not installed — run `pip install streamlit-aggrid`.")
            else:
                _ACTIONS         = ["", "Keep", "Delete", "New", "Delist"]
                _ACT_SET         = {"Keep", "Delete", "New", "Delist"}
                _COL_W           = {"DG Code": 86, "ID": 120, "Item Name": 210}
                _AVG_U_STD       = "Avg Units 52wk/ Forecast new item sales"
                _pog_actions_key = f"{p}_pog_actions"
                _avg_u_edits_key = f"{p}_avg_u_edits"
                _data_edits_key = f"{p}_data_edits"
                _status_ov_key   = f"{p}_status_overrides"
                _pending_del_key = f"{p}_pending_top_delete"
                _pending_new_key = f"{p}_pending_low_sales_new"
                _blank_hl_key    = f"{p}_highlight_submit_blanks"
                _edit_state_scope = (
                    f"{p}|{large_file_path or ''}|"
                    f"{getattr(df_src, 'shape', ('', ''))}|"
                    f"{_sel_dg_code or ''}|{_sel_dg_name or ''}|"
                    f"{','.join(map(str, _STICKY))}"
                )
                _edit_state_scope_key = f"{p}_edit_state_scope"
                if st.session_state.get(_edit_state_scope_key) != _edit_state_scope:
                    _loaded_edit_state = _rs_load_edit_state(_edit_state_scope)
                    if not any(_loaded_edit_state.get(_k) for _k in ("pog_actions", "pog_edits", "avg_u_edits", "data_edits", "status_overrides")):
                        _legacy_edit_state_scope = (
                            f"{p}|{large_file_path or ''}|"
                            f"{getattr(df_src, 'shape', ('', ''))}|{','.join(map(str, _STICKY))}"
                        )
                        _legacy_edit_state = _rs_load_edit_state(_legacy_edit_state_scope)
                        if any(_legacy_edit_state.get(_k) for _k in ("pog_actions", "pog_edits", "avg_u_edits", "data_edits", "status_overrides")):
                            _loaded_edit_state = _legacy_edit_state
                    st.session_state[_edit_state_scope_key] = _edit_state_scope
                    st.session_state[_pog_actions_key] = _loaded_edit_state.get("pog_actions", {})
                    st.session_state[_pog_edits_key] = _loaded_edit_state.get("pog_edits", {})
                    st.session_state[_avg_u_edits_key] = _loaded_edit_state.get("avg_u_edits", {})
                    st.session_state[_data_edits_key] = _loaded_edit_state.get("data_edits", {})
                    st.session_state[_status_ov_key] = _loaded_edit_state.get("status_overrides", {})

                # setdefault instead of get: returns the LIVE in-session dict (or creates it).
                # Without this, .get() on a missing key returns a temporary {} that is not
                # the same object as the one _pog_act_store writes into, so _derive_status
                # always sees an empty dict on the first action and returns "MAINTAIN".
                _pog_actions      = st.session_state.setdefault(_pog_actions_key, {})
                _avg_u_edits      = st.session_state.setdefault(_avg_u_edits_key, {})
                _data_edits       = st.session_state.setdefault(_data_edits_key, {})
                _status_overrides = st.session_state.setdefault(_status_ov_key, {})

                def _num_for_rank(v):
                    if v is None or (isinstance(v, float) and pd.isna(v)):
                        return None
                    s = str(v).strip().replace(",", "")
                    if s in ("", "nan", "None") or s in _ACT_SET:
                        return None
                    try:
                        return float(s)
                    except (TypeError, ValueError):
                        return None

                def _find_tdf_col(names: list[str]):
                    targets = {_nca(x) for x in names}
                    for c in _tdf.columns:
                        if _nca(c) in targets:
                            return c
                    return None

                def _low_sales_rank_rows(limit: int = 5):
                    vol_col = _find_tdf_col([
                        "TH_Tot_Sales_Volume_52WK",
                        "TH Tot Sales Volume 52WK",
                    ])
                    fc_col = _find_tdf_col(["ForecastSales", "Forecast Sales"])
                    mapped_col = _find_tdf_col(["Avg Units 52wk/ Forecast new item sales"])
                    if not vol_col:
                        vol_col = mapped_col
                    if not fc_col:
                        fc_col = mapped_col
                    rows = []
                    if not vol_col and not fc_col:
                        return rows
                    for _idx, _row in _tdf.iterrows():
                        _vol = _num_for_rank(_row.get(vol_col)) if vol_col else None
                        _fc = _num_for_rank(_row.get(fc_col)) if fc_col else None
                        _metric = _vol if _vol is not None else _fc
                        if _metric is not None:
                            rows.append((_metric, _idx))
                    return [idx for _, idx in sorted(rows, key=lambda x: x[0])[:limit]]

                _low_sales_row_set = set(_low_sales_rank_rows(5))

                def _is_low_sales_item(row_i: int) -> bool:
                    return row_i in _low_sales_row_set

                def _delete_is_top10(row_i: int, pog_col: str):
                    val = _num_for_rank(_tdf.at[row_i, pog_col])
                    if val is None:
                        return False, None, None
                    nums = []
                    for pc in _dyn_pog_cols:
                        if pc in _tdf.columns:
                            n = _num_for_rank(_tdf.at[row_i, pc])
                            if n is not None:
                                nums.append(n)
                    if not nums:
                        return False, val, None
                    nums.sort(reverse=True)
                    cutoff_idx = max(1, int(math.ceil(len(nums) * 0.10))) - 1
                    cutoff = nums[cutoff_idx]
                    return val >= cutoff, val, cutoff

                # ── Status derivation from action overlay ─────────────────────────
                _rk_to_i = {_make_rk(_ri): _ri for _ri in range(len(_tdf))}

                def _row_numeric_pog_cols(row_i: int) -> list[str]:
                    cols = []
                    for pc in _dyn_pog_cols:
                        if pc in _tdf.columns and _num_for_rank(_tdf.at[row_i, pc]) is not None:
                            cols.append(pc)
                    return cols

                def _derive_status(rk: tuple) -> str:
                    _acts = {c: v for c, v in _pog_actions.get(rk, {}).items() if v}
                    _stored_status = str(_status_overrides.get(rk, "")).strip().upper()
                    if not _acts:
                        if _stored_status in {
                            "DELETE SOME",
                            "DELETE ALL",
                            "NEW SOME",
                            "NEWNEW",
                            "NEW DELETE SOME",
                            "INACTIVE",
                        }:
                            return _stored_status
                        return "MAINTAIN"
                    if "Delist" in _acts.values():
                        return "Inactive"
                    _delete_cols = {c for c, v in _acts.items() if v == "Delete"}
                    _new_cols = {c for c, v in _acts.items() if v == "New"}
                    _nd, _nn = len(_delete_cols), len(_new_cols)
                    _valid_pog_cols = [c for c in _dyn_pog_cols if c in _tdf.columns]
                    _np = len(_valid_pog_cols)
                    _row_i = _rk_to_i.get(rk)
                    _existing_cols = (
                        set(_row_numeric_pog_cols(_row_i))
                        if _row_i is not None else set()
                    )
                    _delete_scope = _existing_cols | _delete_cols
                    if _nd > 0 and _nn > 0:
                        return "NEW DELETE SOME" if _is_sspog else "DELETE SOME"
                    if _nd > 0 and _delete_scope and _delete_cols >= _delete_scope:
                        return "DELETE ALL"
                    if _nd > 0:
                        return "DELETE SOME"
                    if _np > 0 and _nn >= _np:
                        return "NEWNEW"
                    if _nn > 0:
                        return "NEW SOME"
                    return "MAINTAIN"

                _pending_del = st.session_state.get(_pending_del_key)
                if _pending_del:
                    _pending_del_items = _pending_del if isinstance(_pending_del, list) else [_pending_del]
                    _item_nm = _pending_del_items[0].get("item", "") if _pending_del_items else ""
                    _pog_nm = ", ".join(str(x.get("pog", "")) for x in _pending_del_items if x.get("pog"))
                    st.warning(
                        f"Item {_item_nm} is a Top 10% best seller in this row "
                        f"for planogram {_pog_nm}. Do you still want to delete it?"
                    )
                    _c_ok, _c_cancel = st.columns([1, 1])
                    if _c_ok.button("Confirm Delete", key=f"{p}_confirm_top_delete"):
                        for _pd in _pending_del_items:
                            _rk = tuple(_pd.get("rk", ()))
                            _pc = _pd.get("col", "")
                            if _rk and _pc:
                                _pog_actions.setdefault(_rk, {})[_pc] = "Delete"
                                _pog_edits.setdefault(_rk, {})[_pc] = "Delete"
                                _status_overrides[_rk] = _derive_status(_rk)
                        _rs_save_edit_state(
                            _edit_state_scope,
                            _pog_actions,
                            _pog_edits,
                            _avg_u_edits,
                            _status_overrides,
                            _data_edits,
                        )
                        st.session_state.pop(_pending_del_key, None)
                        st.rerun()
                    if _c_cancel.button("Cancel", key=f"{p}_cancel_top_delete"):
                        st.session_state.pop(_pending_del_key, None)
                        st.rerun()

                _pending_new = st.session_state.get(_pending_new_key)
                if _pending_new:
                    _pending_new_items = _pending_new if isinstance(_pending_new, list) else [_pending_new]
                    _item_nm = _pending_new_items[0].get("item", "") if _pending_new_items else ""
                    _pog_nm = ", ".join(str(x.get("pog", "")) for x in _pending_new_items if x.get("pog"))
                    st.warning(
                        f"Item {_item_nm} is a Top 10% lowest seller. "
                        f"Are you sure you want to add it to planogram {_pog_nm}?"
                    )
                    _n_ok, _n_cancel = st.columns([1, 1])
                    if _n_ok.button("Confirm New", key=f"{p}_confirm_low_sales_new"):
                        for _pn in _pending_new_items:
                            _rk = tuple(_pn.get("rk", ()))
                            _pc = _pn.get("col", "")
                            if _rk and _pc:
                                _pog_actions.setdefault(_rk, {})[_pc] = "New"
                                _pog_edits.setdefault(_rk, {})[_pc] = "New"
                                _status_overrides[_rk] = _derive_status(_rk)
                        _rs_save_edit_state(
                            _edit_state_scope,
                            _pog_actions,
                            _pog_edits,
                            _avg_u_edits,
                            _status_overrides,
                            _data_edits,
                        )
                        st.session_state.pop(_pending_new_key, None)
                        st.rerun()
                    if _n_cancel.button("Cancel New", key=f"{p}_cancel_low_sales_new"):
                        st.session_state.pop(_pending_new_key, None)
                        st.rerun()

                # ── Build display DataFrame: apply action + avg-u + status overlays ─
                _tdf_display = _tdf.copy()
                _display_rks = [_make_rk(_dri) for _dri in range(len(_tdf_display))]
                _derived_statuses = [_derive_status(_drk) for _drk in _display_rks]
                _object_cols_needed = set(_dyn_pog_cols)
                _object_cols_needed.update(c for _av in _avg_u_edits.values() for c in (_av or {}))
                _object_cols_needed.update(c for _dv in _data_edits.values() for c in (_dv or {}))
                _object_cols_needed.update(["Status", _ROW_KEY_COL, _NEW_ROW_COL, _EDIT_COL])
                _object_cols_needed = [c for c in _object_cols_needed if c in _tdf_display.columns]
                if _object_cols_needed:
                    _tdf_display[_object_cols_needed] = _tdf_display[_object_cols_needed].astype(object)
                for _dri, _drk in enumerate(_display_rks):
                    _derived_status = _derived_statuses[_dri]
                    if _derived_status == "DELETE ALL":
                        for _dpc in _dyn_pog_cols:
                            if (
                                _dpc in _tdf_display.columns
                                and _dri in _cs_original_present.get(_dpc, set())
                            ):
                                _tdf_display.at[_dri, _dpc] = "Delete"
                    for _dpc, _dact in _pog_actions.get(_drk, {}).items():
                        if _dpc in _tdf_display.columns and _dact:
                            _tdf_display.at[_dri, _dpc] = _dact
                    for _dac, _dav in _avg_u_edits.get(_drk, {}).items():
                        if _dac in _tdf_display.columns:
                            _tdf_display.at[_dri, _dac] = _dav
                    for _ddc, _ddv in _data_edits.get(_drk, {}).items():
                        if _ddc in _tdf_display.columns:
                            _tdf_display.at[_dri, _ddc] = _ddv
                    if "Status" in _tdf_display.columns:
                        _tdf_display.at[_dri, "Status"] = _derived_status

                # Force POG cells to text for AgGrid. If these columns are inferred
                # as numeric, the select editor can visually change but fail to
                # commit action labels such as Delete/New back to Streamlit.
                _display_pog_cols = [c for c in _dyn_pog_cols if c in _tdf_display.columns]
                if _display_pog_cols:
                    _pog_text = _tdf_display[_display_pog_cols]
                    _tdf_display[_display_pog_cols] = _pog_text.where(_pog_text.notna(), "").astype(str)

                _saved_item_rows = list(st.session_state.get(_new_rows_key, []))
                _placeholder_item_rows = list(st.session_state.get(_new_rows_placeholder_key, []))
                _pending_new_items = int(st.session_state.pop(_new_rows_pending_key, 0) or 0)
                if _pending_new_items > 0:
                    _start_seq = len(_saved_item_rows) + len(_placeholder_item_rows)
                    for _j in range(_pending_new_items):
                        _rk_new = ["__NEW__", _new_rows_scope, str(_start_seq + _j)]
                        _row_new = {c: "" for c in _tdf_display.columns}
                        _row_new[_ROW_KEY_COL] = _json.dumps(_rk_new, ensure_ascii=False)
                        _row_new[_NEW_ROW_COL] = True
                        _placeholder_item_rows.append(_row_new)
                    st.session_state[_new_rows_placeholder_key] = _placeholder_item_rows
                    st.session_state[_new_rows_scroll_key] = _placeholder_item_rows[-_pending_new_items].get(_ROW_KEY_COL, "")

                _new_item_rows = _saved_item_rows + _placeholder_item_rows

                if _new_item_rows:
                    def _new_row_has_input(_row: dict) -> bool:
                        _skip_cols = {
                            _ROW_KEY_COL,
                            _NEW_ROW_COL,
                            _EDIT_COL,
                            "Status",
                            "Check Range To-be Waterfall",
                        }
                        for _c, _v in (_row or {}).items():
                            if _c in _skip_cols:
                                continue
                            _s = "" if _v is None else str(_v).strip()
                            if _s and _s.lower() not in ("nan", "none"):
                                return True
                        return False

                    def _new_row_avg_is_numeric(_row: dict) -> bool:
                        _v = (_row or {}).get(_AVG_U_STD, "")
                        _s = "" if _v is None else str(_v).replace(",", "").strip()
                        if not _s or _s.lower() in ("nan", "none"):
                            return False
                        try:
                            float(_s)
                            return True
                        except Exception:
                            return False

                    _normalized_new_rows = []
                    for _nr in _new_item_rows:
                        _row_new = {c: _nr.get(c, "") for c in _tdf_display.columns}
                        _row_new[_ROW_KEY_COL] = _nr.get(_ROW_KEY_COL, "")
                        _row_new[_NEW_ROW_COL] = True
                        _id_val = _re.sub(r"\D+", "", str(_row_new.get("ID", "")).strip())[:9]
                        if "ID" in _row_new:
                            _row_new["ID"] = _id_val
                        if "Status" in _row_new:
                            _row_new["Status"] = (
                                "NEWNEW"
                                if _re.fullmatch(r"\d{9}", _id_val)
                                and _new_row_avg_is_numeric(_row_new)
                                else ""
                            )
                        _normalized_new_rows.append(_row_new)
                    _new_item_rows = _normalized_new_rows
                    _saved_new_rows_display = [
                        _nr for _nr in _new_item_rows
                        if _new_row_has_input(_nr)
                    ]
                    _placeholder_new_rows_display = [
                        _nr for _nr in _new_item_rows
                        if not _new_row_has_input(_nr)
                    ]
                    st.session_state[_new_rows_key] = _saved_new_rows_display
                    st.session_state[_new_rows_placeholder_key] = _placeholder_new_rows_display
                    if len(_saved_new_rows_display) != len(_saved_item_rows):
                        _rs_save_new_rows(_new_rows_scope, _saved_new_rows_display)
                    _tdf_display = pd.concat(
                        [_tdf_display, pd.DataFrame(_new_item_rows)],
                        ignore_index=True,
                    )

                _search_pog_match_cols = [
                    c for c in _dyn_pog_cols
                    if _search_q_n and _search_q_n in _nca(c)
                ]

                if _search_q_text and not _search_pog_match_cols:
                    _search_mask = pd.Series(False, index=_tdf_display.index)
                    for _sc in [c for c in _tdf_display.columns if c != _ROW_KEY_COL]:
                        if _sc in (_NEW_ROW_COL, _EDIT_COL):
                            continue
                        _search_mask |= _tdf_display[_sc].astype(str).str.contains(
                            _search_q_text, case=False, na=False, regex=False
                        )
                    if _NEW_ROW_COL in _tdf_display.columns:
                        _search_mask |= _tdf_display[_NEW_ROW_COL].astype(bool)
                    if _search_mask.any():
                        _tdf_display = _tdf_display.loc[_search_mask]
                    else:
                        _tdf_display = _tdf_display.iloc[0:0]

                def _sspog_layout_xlsx_bytes() -> bytes:
                    try:
                        from openpyxl import Workbook
                        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
                        from openpyxl.utils import get_column_letter
                    except Exception:
                        return df_to_xlsx_bytes(_tdf_display.drop(
                            columns=[c for c in (_ROW_KEY_COL, _NEW_ROW_COL, _EDIT_COL) if c in _tdf_display.columns],
                            errors="ignore",
                        ))

                    def _clean_xl(v):
                        if v is None:
                            return ""
                        try:
                            if pd.isna(v):
                                return ""
                        except Exception:
                            pass
                        if isinstance(v, str) and v in ("nan", "None", "<NA>"):
                            return ""
                        return v

                    def _as_num_or_text(v):
                        v = _clean_xl(v)
                        if isinstance(v, str):
                            s = v.replace(",", "").strip()
                            if s and s not in _ACT_SET:
                                try:
                                    return float(s) if "." in s else int(s)
                                except Exception:
                                    return v
                        return v

                    def _write_matrix(ws, start_row, start_col, data, style_kind="plain"):
                        thin = Side(style="thin", color="B8B8B8")
                        border = Border(left=thin, right=thin, top=thin, bottom=thin)
                        for r_off, row in enumerate(data):
                            for c_off, val in enumerate(row):
                                cell = ws.cell(start_row + r_off, start_col + c_off, _clean_xl(val))
                                cell.border = border
                                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                                if style_kind == "header" or r_off == 0:
                                    cell.fill = PatternFill("solid", fgColor="D9D9D9")
                                    cell.font = Font(bold=True, color="1A1A1A")
                        return start_row + len(data) - 1

                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Range Sheet SSPOG"
                    ws.sheet_view.showGridLines = False
                    thin = Side(style="thin", color="B8B8B8")
                    border = Border(left=thin, right=thin, top=thin, bottom=thin)
                    dark_fill = PatternFill("solid", fgColor="808080")
                    grey_fill = PatternFill("solid", fgColor="D9D9D9")
                    pale_fill = PatternFill("solid", fgColor="F5F7F7")
                    pink_fill = PatternFill("solid", fgColor="FCE4EC")
                    yellow_fill = PatternFill("solid", fgColor="FFF59D")
                    green_fill = PatternFill("solid", fgColor="C6EFCE")
                    delete_fill = PatternFill("solid", fgColor="FDE2E2")
                    new_fill = PatternFill("solid", fgColor="E3F2FD")

                    main_df = _tdf_display.drop(
                        columns=[c for c in (_ROW_KEY_COL, _NEW_ROW_COL, _EDIT_COL) if c in _tdf_display.columns],
                        errors="ignore",
                    ).copy()
                    for _c in main_df.columns:
                        if _c in _dyn_pog_cols:
                            main_df[_c] = main_df[_c].map(_clean_xl)

                    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(8, min(14, len(main_df.columns))))
                    ws.cell(1, 1, "Range Sheet SSPOG").font = Font(bold=True, size=16, color="FFFFFF")
                    ws.cell(1, 1).fill = PatternFill("solid", fgColor="2BBFA4")
                    ws.cell(1, 1).alignment = Alignment(horizontal="left", vertical="center")
                    ws.row_dimensions[1].height = 26
                    _export_dg_name = (
                        fixed_dg_name
                        or st.session_state.get(f"{p}_selected_dg_name")
                        or st.session_state.get("_ns_loaded_dg_name")
                        or _sel_dg_name
                        or "ALL"
                    )
                    if str(_export_dg_name).strip().upper().startswith("— ALL"):
                        _export_dg_name = "ALL"
                    ws.cell(2, 1, f"DG Code: {_sel_dg_code or 'ALL'}")
                    ws.cell(2, 2, f"DG Name: {_export_dg_name or 'ALL'}")
                    ws.cell(2, 4, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

                    arch_headers = [
                        "TYPE", "AS IS", "TO BE",
                        "Sale AS IS", "Sale TO BE", "Sale DIFF",
                        "Margin AS IS",
                    ]
                    try:
                        arch_rows_export = list(_arch_rows)
                    except NameError:
                        arch_rows_export = []
                    try:
                        arch_total_as_is = _t_ai
                    except NameError:
                        arch_total_as_is = sum(r.get("as_is", 0) for r in arch_rows_export)
                    try:
                        arch_total_to_be = _t_tb
                    except NameError:
                        arch_total_to_be = sum(r.get("to_be", 0) for r in arch_rows_export)
                    try:
                        arch_total_sale_as_is = _t_sai
                    except NameError:
                        arch_total_sale_as_is = sum(r.get("sale_ai", 0) for r in arch_rows_export)
                    try:
                        arch_total_sale_to_be = _t_stb
                    except NameError:
                        arch_total_sale_to_be = sum(r.get("sale_tb", 0) for r in arch_rows_export)
                    try:
                        arch_total_sale_diff = _t_sdiff
                    except NameError:
                        arch_total_sale_diff = sum(r.get("sale_diff", 0) for r in arch_rows_export)
                    try:
                        arch_total_margin_as_is = _t_mai
                    except NameError:
                        arch_total_margin_as_is = sum(r.get("marg_ai", 0) for r in arch_rows_export)
                    try:
                        arch_pct_sale = _pct_sale
                    except NameError:
                        arch_pct_sale = "0.0%"
                    arch_data = [["Range architecture", "", "", "Sale Impact / Week", "", "", "Margin Impact / Week"], arch_headers]
                    for _r in arch_rows_export:
                        arch_data.append([
                            _r.get("type", ""),
                            _r.get("as_is", 0),
                            _r.get("to_be", 0),
                            _r.get("sale_ai", 0),
                            _r.get("sale_tb", 0),
                            _r.get("sale_diff", 0),
                            _r.get("marg_ai", 0),
                        ])
                    arch_data.append([
                        "TOTAL SKU",
                        arch_total_as_is,
                        arch_total_to_be,
                        arch_total_sale_as_is,
                        arch_total_sale_to_be,
                        arch_total_sale_diff,
                        arch_total_margin_as_is,
                    ])
                    arch_data.append(["% Impact", "", "0.0%", "", "", arch_pct_sale, ""])
                    arch_end = _write_matrix(ws, 4, 1, arch_data)
                    for row in ws.iter_rows(min_row=4, max_row=arch_end, min_col=1, max_col=7):
                        for cell in row:
                            cell.border = border
                        row[0].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                    for cell in ws[4]:
                        if cell.column <= 7:
                            cell.fill = grey_fill
                            cell.font = Font(bold=True)
                    for cell in ws[5]:
                        if cell.column <= 7:
                            cell.fill = grey_fill
                            cell.font = Font(bold=True)
                    ws.cell(arch_end - 1, 1).font = Font(bold=True)
                    for col in range(1, 8):
                        ws.cell(arch_end - 1, col).font = Font(bold=True)

                    main_cols = list(main_df.columns)
                    first_pog_idx = next((i for i, c in enumerate(main_cols) if c in _dyn_pog_cols), len(main_cols))
                    main_start_row = max(arch_end + 4, 20)
                    main_start_col = 1
                    first_pog_col = main_start_col + first_pog_idx
                    cluster_label_col = max(main_start_col, first_pog_col - 2)
                    cluster_start_row = 4

                    try:
                        cluster_rows_export = list(_CS_ROWS)
                    except NameError:
                        cluster_rows_export = []
                    cluster_cols = [c for c in _dyn_pog_cols if c in main_cols]
                    cluster_header = ["Cluster"] + [
                        _pog_to_cl.get(_pog) or _pog_to_cl.get(_nca(_pog), "") or str(_pog)
                        for _pog in cluster_cols
                    ]
                    cluster_matrix = [cluster_header]
                    for _rl, _rbg, _rtc in cluster_rows_export:
                        row = [_rl]
                        for _pog in cluster_cols:
                            _cl_key = _pog_to_cl.get(_pog) or _pog_to_cl.get(_nca(_pog), "")
                            if _rl in (
                                "MODs", "FIXTURE", "RANGE CLASS",
                                "Total NEW SKUs", "Total DELETE SKUs",
                                "TO-BE SKUs count", "AS-IS SKUs count",
                            ):
                                row.append(_cs_val(_rl, _pog, _cl_key))
                            else:
                                _ssv = str(_ct_dat_disp.get(_rl, {}).get(_cl_key, "") or "")
                                row.append("" if _ssv in ("nan", "None") else (_ssv or _cs_val(_rl, _pog, _cl_key)))
                        cluster_matrix.append(row)
                    _write_matrix(ws, cluster_start_row, cluster_label_col, cluster_matrix)
                    for row in ws.iter_rows(
                        min_row=cluster_start_row,
                        max_row=cluster_start_row + len(cluster_matrix) - 1,
                        min_col=cluster_label_col,
                        max_col=cluster_label_col + len(cluster_matrix[0]) - 1,
                    ):
                        for cell in row:
                            cell.border = border
                            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                            if cell.row == cluster_start_row:
                                cell.fill = grey_fill
                                cell.font = Font(bold=True)
                                if cell.column > cluster_label_col:
                                    cell.alignment = Alignment(textRotation=90, horizontal="center", vertical="center", wrap_text=True)
                    for r in range(cluster_start_row + 1, cluster_start_row + len(cluster_matrix)):
                        ws.cell(r, cluster_label_col).font = Font(bold=True)
                        if ws.cell(r, cluster_label_col).value in ("Total NEW SKUs", "Total DELETE SKUs", "TO-BE SKUs count"):
                            for c in range(cluster_label_col, cluster_label_col + len(cluster_matrix[0])):
                                ws.cell(r, c).fill = PatternFill("solid", fgColor="F2F2F2")
                    ws.row_dimensions[cluster_start_row].height = 54
                    fixture_row = cluster_start_row + 1 + next((i for i, row in enumerate(cluster_matrix[1:]) if row and row[0] == "FIXTURE"), 3)
                    ws.row_dimensions[fixture_row].height = 48
                    for c in range(cluster_label_col + 1, cluster_label_col + len(cluster_matrix[0])):
                        ws.cell(fixture_row, c).alignment = Alignment(textRotation=90, horizontal="center", vertical="center", wrap_text=True)

                    # Bridge grid lines between cluster summary and main table for one-sheet feel.
                    bridge_start = cluster_start_row + len(cluster_matrix)
                    for r in range(bridge_start, main_start_row):
                        for c in range(cluster_label_col, cluster_label_col + len(cluster_matrix[0])):
                            ws.cell(r, c).border = Border(left=thin, right=thin)

                    # Main table
                    for j, col in enumerate(main_cols, start=main_start_col):
                        cell = ws.cell(main_start_row, j, col)
                        cell.fill = pale_fill
                        cell.font = Font(bold=True)
                        cell.border = border
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                        if col in _dyn_pog_cols:
                            cell.alignment = Alignment(textRotation=90, horizontal="center", vertical="center", wrap_text=True)
                    ws.row_dimensions[main_start_row].height = 56
                    for i, (_, row) in enumerate(main_df.iterrows(), start=main_start_row + 1):
                        ws.row_dimensions[i].height = 18
                        for j, col in enumerate(main_cols, start=main_start_col):
                            val = _as_num_or_text(row.get(col, ""))
                            cell = ws.cell(i, j, val)
                            cell.border = border
                            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
                            if col == "Item Name":
                                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True)
                            if col == "Status":
                                sv = str(val).upper()
                                if "DELETE" in sv:
                                    cell.fill = delete_fill
                                elif "NEW" in sv:
                                    cell.fill = yellow_fill
                            elif col in _dyn_pog_cols:
                                if str(val) == "Delete":
                                    cell.fill = delete_fill
                                    cell.font = Font(color="C00000", bold=True)
                                elif str(val) == "New":
                                    cell.fill = new_fill
                                    cell.font = Font(color="0066CC", bold=True)
                    # Do not freeze panes here. Freezing at the main table row locks
                    # the whole summary area in Excel and makes vertical scrolling
                    # appear stuck on the range architecture section.
                    ws.freeze_panes = None
                    ws.auto_filter.ref = (
                        f"{get_column_letter(main_start_col)}{main_start_row}:"
                        f"{get_column_letter(main_start_col + len(main_cols) - 1)}"
                        f"{main_start_row + max(len(main_df), 1)}"
                    )

                    for idx, col in enumerate(main_cols, start=main_start_col):
                        if col in ("DG Code", "Status"):
                            ws.column_dimensions[get_column_letter(idx)].width = 10
                        elif col == "ID":
                            ws.column_dimensions[get_column_letter(idx)].width = 11
                        elif col == "Item Name":
                            ws.column_dimensions[get_column_letter(idx)].width = 26
                        elif col in _dyn_pog_cols:
                            ws.column_dimensions[get_column_letter(idx)].width = 7
                        else:
                            ws.column_dimensions[get_column_letter(idx)].width = 11
                    for col in range(cluster_label_col, cluster_label_col + len(cluster_matrix[0])):
                        ws.column_dimensions[get_column_letter(col)].width = 7 if col > cluster_label_col else 18

                    out = io.BytesIO()
                    wb.save(out)
                    return out.getvalue()

                # ── JsCode: pog cells show action label OR formatted number ────────
                # Rule-based chatbot over the currently filtered rangesheet view.
                def _chat_norm(q: str) -> str:
                    s = str(q or "").lower().strip()
                    repl = {
                        "planogramme": "planogram",
                        "planogrammes": "planogram",
                        "planograms": "planogram",
                        "pog": "planogram",
                        "แพลนโนแกรม": "planogram",
                        "แพลน": "planogram",
                        "มากสุด": "มากที่สุด",
                        "เยอะสุด": "มากที่สุด",
                        "สูงสุด": "มากที่สุด",
                        "น้อยสุด": "น้อยที่สุด",
                        "ต่ำสุด": "น้อยที่สุด",
                        "ต่ําสุด": "น้อยที่สุด",
                        "ลบออก": "ลบ",
                    }
                    for a, b in repl.items():
                        s = s.replace(a, b)
                    return _re.sub(r"\s+", " ", s).strip()

                _CHAT_INTENT_LIBRARY = {
                    "planogram_most_products": {
                        "label": "Planogramme ไหนที่มีสินค้ามากที่สุด",
                        "samples": [
                            "planogramme ไหนที่มีสินค้ามากที่สุด",
                            "planogram ไหนมีสินค้ามากสุด",
                            "แพลนไหนสินค้ามากที่สุด",
                            "which planogram has the most products",
                            "most sku planogram",
                        ],
                    },
                    "planogram_fewest_products": {
                        "label": "Planogramme ไหนที่มีสินค้าน้อยที่สุด",
                        "samples": [
                            "planogramme ไหนที่มีสินค้าน้อยที่สุด",
                            "planogram ไหนมีสินค้าน้อยสุด",
                            "แพลนไหนสินค้าน้อยที่สุด",
                            "which planogram has the fewest products",
                            "least sku planogram",
                        ],
                    },
                    "lowest_sales_items": {
                        "label": "ควรลบสินค้าไหนออก / สินค้าที่มียอดขายต่ำที่สุด",
                        "samples": [
                            "คิดว่าควรลบสินค้าไหนออก",
                            "สินค้าที่มียอดขายต่ำที่สุด",
                            "สินค้าไหนขายต่ำสุด",
                            "recommend delete lowest sales items",
                            "lowest sales product",
                        ],
                    },
                    "planogram_highest_total": {
                        "label": "Planogramme ไหนมียอดขายรวมมากที่สุด",
                        "samples": [
                            "planogramme ไหนมียอดขายรวมมากที่สุด",
                            "แพลนไหนยอดขายรวมสูงสุด",
                            "highest total sales planogram",
                        ],
                    },
                    "planogram_lowest_total": {
                        "label": "Planogramme ไหนมียอดขายรวมน้อยที่สุด",
                        "samples": [
                            "planogramme ไหนมียอดขายรวมน้อยที่สุด",
                            "แพลนไหนยอดขายรวมต่ำสุด",
                            "lowest total sales planogram",
                        ],
                    },
                    "item_most_planograms": {
                        "label": "สินค้าไหนอยู่ใน planogramme มากที่สุด",
                        "samples": [
                            "สินค้าไหนอยู่ใน planogramme มากที่สุด",
                            "item in most planograms",
                            "สินค้าอยู่ในแพลนมากสุด",
                        ],
                    },
                    "item_fewest_planograms": {
                        "label": "สินค้าไหนอยู่ใน planogramme น้อยที่สุด",
                        "samples": [
                            "สินค้าไหนอยู่ใน planogramme น้อยที่สุด",
                            "item in fewest planograms",
                            "สินค้าอยู่ในแพลนน้อยสุด",
                        ],
                    },
                    "planogram_most_blanks": {
                        "label": "Planogramme ไหนมีช่องว่างมากที่สุด",
                        "samples": [
                            "planogramme ไหนมีช่องว่างมากที่สุด",
                            "แพลนไหน blank มากสุด",
                            "planogram with most blanks",
                            "missing most planogram",
                        ],
                    },
                    "action_counts": {
                        "label": "มีสินค้า New/Delete/Delist/Maintain กี่ตัว",
                        "samples": [
                            "มีสินค้า new delete delist maintain กี่ตัว",
                            "count status action",
                            "นับสถานะ",
                        ],
                    },
                    "summary": {
                        "label": "สรุป DG นี้มีสินค้ากี่ตัวและ planogramme กี่ตัว",
                        "samples": [
                            "สรุป DG นี้",
                            "มีสินค้ากี่ตัวและ planogramme กี่ตัว",
                            "dg summary",
                            "count sku and planogram",
                        ],
                    },
                }

                def _chat_tokens(s: str) -> set[str]:
                    return set(_re.findall(r"[a-z0-9\u0E00-\u0E7F]+", _chat_norm(s)))

                def _score_chat_intent(q: str, intent: str) -> float:
                    qn = _chat_norm(q)
                    qt = _chat_tokens(qn)
                    entry = _CHAT_INTENT_LIBRARY.get(intent, {})
                    candidates = [entry.get("label", "")] + list(entry.get("samples", []))
                    best = 0.0
                    for cand in candidates:
                        cn = _chat_norm(cand)
                        ct = _chat_tokens(cn)
                        ratio = difflib.SequenceMatcher(None, qn, cn).ratio()
                        overlap = (len(qt & ct) / max(len(qt | ct), 1)) if qt or ct else 0.0
                        best = max(best, (ratio * 0.55) + (overlap * 0.45))
                    return best

                def _best_chat_intent(q: str):
                    ranked = sorted(
                        ((intent, _score_chat_intent(q, intent)) for intent in _CHAT_INTENT_LIBRARY),
                        key=lambda x: x[1],
                        reverse=True,
                    )
                    if not ranked or ranked[0][1] < 0.18:
                        return "unknown", 0.0
                    return ranked[0]

                def _num_cell(v):
                    if v is None or (isinstance(v, float) and pd.isna(v)):
                        return None
                    s = str(v).strip().replace(",", "")
                    if s in ("", "nan", "None") or s in _ACT_SET:
                        return None
                    try:
                        return float(s)
                    except (TypeError, ValueError):
                        return None

                def _numeric_planogram_matrix(src_df: pd.DataFrame) -> pd.DataFrame:
                    cols = [c for c in _dyn_pog_cols if c in src_df.columns]
                    out = pd.DataFrame(index=src_df.index)
                    for c in cols:
                        out[c] = src_df[c].map(_num_cell)
                    return out

                def _fmt_item(row) -> str:
                    dg = str(row.get("DG Code", "") or "").strip()
                    iid = str(row.get("ID", "") or "").strip()
                    item = str(row.get("Item Name", "") or "").strip()
                    parts = [x for x in (dg, iid, item) if x and x not in ("nan", "None")]
                    return "-".join(parts) if parts else "(unknown item)"

                def _find_col(df: pd.DataFrame, names: list[str]):
                    targets = {_nca(x) for x in names}
                    for c in df.columns:
                        if _nca(c) in targets:
                            return c
                    return None

                def _chat_cmp(s: str) -> str:
                    return _re.sub(r"[^a-z0-9\u0E00-\u0E7F]+", "", str(s or "").lower())

                def _chat_delete_command(q: str) -> bool:
                    s = _chat_norm(q)
                    return (
                        any(k in s for k in ("delete", "ลบ", "ลด", "เอาออก"))
                        and any(k in s for k in (
                            "planogram", "แพลน", "target", "ออกจาก", "ออก",
                            "ทั้งหมด", "ทั้งแถว", "ทุก"
                        ))
                    )

                def _chat_add_command(q: str) -> bool:
                    s = _chat_norm(q)
                    return (
                        any(k in s for k in ("add", "new", "เพิ่ม", "ใส่"))
                        and any(k in s for k in (
                            "planogram", "แพลน", "target", "ลงใน", "เข้า",
                            "ทั้งหมด", "ทั้งแถว", "ทุก"
                        ))
                    )

                def _chat_question_parts(q: str) -> list[str]:
                    raw = str(q or "")
                    parts = [raw]
                    for marker in ("planogramme", "planogram", "pog", "แพลน", "ออกจาก"):
                        if marker in raw.lower():
                            pieces = _re.split(marker, raw, flags=_re.IGNORECASE)
                            parts.extend(pieces)
                    return [p.strip() for p in parts if str(p).strip()]

                def _best_chat_planogram(q: str):
                    if not _dyn_pog_cols:
                        return None, 0.0
                    q_parts = _chat_question_parts(q)
                    q_cmp = _chat_cmp(q)
                    q_tokens = _chat_tokens(q)
                    best_col, best_score = None, 0.0
                    for col in _dyn_pog_cols:
                        col_s = str(col)
                        col_cmp = _chat_cmp(col_s)
                        col_tokens = _chat_tokens(col_s)
                        score = 0.0
                        if col_cmp and col_cmp in q_cmp:
                            score = 1.0
                        for part in q_parts:
                            part_cmp = _chat_cmp(part)
                            if not part_cmp:
                                continue
                            if part_cmp and part_cmp in col_cmp:
                                score = max(score, 0.82)
                            score = max(score, difflib.SequenceMatcher(None, part_cmp, col_cmp).ratio())
                        overlap = (
                            len(q_tokens & col_tokens) / max(len(q_tokens | col_tokens), 1)
                            if q_tokens or col_tokens else 0.0
                        )
                        score = max(score, overlap)
                        if score > best_score:
                            best_col, best_score = col_s, score
                    return best_col, best_score

                def _chat_target_numbers(q: str) -> list[str]:
                    raw = str(q or "")
                    out = []
                    for n in _re.findall(r"\b\d{3,6}\b", raw):
                        if n not in out:
                            out.append(n)
                    return out

                def _chat_planograms_from_targets(q: str) -> list[str]:
                    nums = _chat_target_numbers(q)
                    cols = []
                    for n in nums:
                        n_cmp = _chat_cmp(n)
                        matches = []
                        for col in _dyn_pog_cols:
                            col_s = str(col)
                            col_cmp = _chat_cmp(col_s)
                            if (
                                f"target{n_cmp}" in col_cmp
                                or col_cmp.endswith(n_cmp)
                                or n_cmp in col_cmp
                            ):
                                matches.append(col_s)
                        if matches:
                            chosen = sorted(
                                matches,
                                key=lambda c: (f"target{n_cmp}" not in _chat_cmp(c), len(str(c))),
                            )[0]
                            if chosen not in cols:
                                cols.append(chosen)
                    return cols

                def _chat_planogram_columns(q: str) -> list[str]:
                    cols = _chat_planograms_from_targets(q)
                    if cols:
                        return cols
                    pog_col, pog_score = _best_chat_planogram(q)
                    return [pog_col] if pog_col and pog_score >= 0.28 else []

                def _chat_apply_delete_rows(row_indices: list[int], pog_cols: list[str]):
                    pending = []
                    changed = []
                    valid_pog_cols = [c for c in pog_cols if c in _tdf.columns]
                    for row_i in row_indices:
                        if row_i not in _tdf.index:
                            continue
                        rk = _make_rk(row_i)
                        item_txt = _fmt_item(_tdf.loc[row_i])
                        row_changed = False
                        nums_by_col = {}
                        for pc in _dyn_pog_cols:
                            if pc in _tdf.columns:
                                n = _num_for_rank(_tdf.at[row_i, pc])
                                if n is not None:
                                    nums_by_col[pc] = n
                        cutoff = None
                        if nums_by_col:
                            sorted_nums = sorted(nums_by_col.values(), reverse=True)
                            cutoff_idx = max(1, int(math.ceil(len(sorted_nums) * 0.10))) - 1
                            cutoff = sorted_nums[cutoff_idx]
                        for pog_col in valid_pog_cols:
                            cell_val = nums_by_col.get(pog_col)
                            is_top = cutoff is not None and cell_val is not None and cell_val >= cutoff
                            if is_top:
                                pending.append({
                                    "rk": rk,
                                    "row_i": row_i,
                                    "col": pog_col,
                                    "pog": pog_col,
                                    "item": item_txt,
                                    "value": cell_val,
                                    "cutoff": cutoff,
                                })
                                continue
                            _pog_actions.setdefault(rk, {})[pog_col] = "Delete"
                            _pog_edits.setdefault(rk, {})[pog_col] = "Delete"
                            changed.append((item_txt, pog_col))
                            row_changed = True
                        if row_changed:
                            _status_overrides[rk] = _derive_status(rk)
                    if pending:
                        existing = st.session_state.get(_pending_del_key)
                        existing_items = existing if isinstance(existing, list) else ([existing] if existing else [])
                        st.session_state[_pending_del_key] = existing_items + pending
                    return changed, pending

                def _chat_apply_new_rows(row_indices: list[int], pog_cols: list[str]):
                    changed = []
                    valid_pog_cols = [c for c in pog_cols if c in _tdf.columns]
                    for row_i in row_indices:
                        if row_i not in _tdf.index:
                            continue
                        rk = _make_rk(row_i)
                        item_txt = _fmt_item(_tdf.loc[row_i])
                        row_changed = False
                        for pog_col in valid_pog_cols:
                            _pog_actions.setdefault(rk, {})[pog_col] = "New"
                            _pog_edits.setdefault(rk, {})[pog_col] = "New"
                            changed.append((item_txt, pog_col))
                            row_changed = True
                        if row_changed:
                            _status_overrides[rk] = _derive_status(rk)
                    return changed

                def _best_chat_item_row(q: str):
                    q_cmp = _chat_cmp(q)
                    q_tokens = _chat_tokens(q)
                    best_i, best_score = None, 0.0
                    for idx, row in _tdf.iterrows():
                        item_id = str(row.get("ID", "") or "")
                        item_name = str(row.get("Item Name", "") or "")
                        item_label = f"{item_id} {item_name}".strip()
                        id_cmp = _chat_cmp(item_id)
                        name_cmp = _chat_cmp(item_name)
                        label_cmp = _chat_cmp(item_label)
                        label_tokens = _chat_tokens(item_label)
                        score = 0.0
                        if id_cmp and len(id_cmp) >= 3 and id_cmp in q_cmp:
                            score = 1.15
                        if name_cmp and len(name_cmp) >= 4 and name_cmp in q_cmp:
                            score = max(score, 1.05)
                        if label_cmp and label_cmp in q_cmp:
                            score = max(score, 1.0)
                        overlap = (
                            len(q_tokens & label_tokens) / max(len(q_tokens | label_tokens), 1)
                            if q_tokens or label_tokens else 0.0
                        )
                        score = max(score, overlap)
                        if score > best_score:
                            best_i, best_score = idx, score
                    return best_i, best_score

                def _chat_delete_item_from_planogram(q: str):
                    if not _chat_delete_command(q):
                        return None, False
                    row_i, item_score = _best_chat_item_row(q)
                    pog_cols = _chat_planogram_columns(q)
                    if row_i is None or item_score < 0.30:
                        return (
                            "หา item ที่ต้องการลบไม่เจอในข้อมูลหลัง filter ลองพิมพ์ ID หรือ Item Name ให้ชัดขึ้น",
                            False,
                        )
                    if _chat_all_word(q):
                        pog_cols = [c for c in _dyn_pog_cols if c in _tdf.columns]
                    if not pog_cols:
                        item_txt = _fmt_item(_tdf.loc[row_i])
                        st.session_state[f"{p}_rs_chat_recent_delete"] = {
                            "rows": [row_i],
                            "items": [{"row_i": row_i, "item": item_txt}],
                        }
                        return (
                            "ต้องการให้ลบ item นี้ออกจาก planogramme ไหน? "
                            "พิมพ์ชื่อ planogramme / target เช่น 5012 หรือพิมพ์ 'ทั้งหมด' เพื่อลบทั้งแถว",
                            False,
                        )
                    rk = _make_rk(row_i)
                    item_txt = _fmt_item(_tdf.loc[row_i])
                    changed_pairs, pending = _chat_apply_delete_rows([row_i], pog_cols)
                    changed = [pog for _, pog in changed_pairs]
                    if pending:
                        pending_names = ", ".join(str(x["pog"]) for x in pending)
                        changed_txt = (
                            " เปลี่ยนเป็น Delete แล้ว: " + ", ".join(changed) + "."
                            if changed else ""
                        )
                        return (
                            f"{item_txt} in planogram {pending_names} is a Top 10% best seller. "
                            "Please confirm before deleting by clicking Confirm Delete or Cancel above the table."
                            + changed_txt,
                            False,
                        )
                    scope = "ทุก planogramme" if _chat_all_word(q) else ", ".join(pog_cols)
                    return f"เปลี่ยน {item_txt} ใน {scope} เป็น Delete แล้ว Status จะเปลี่ยนตาม row", True

                def _chat_add_item_to_planogram(q: str):
                    if not _chat_add_command(q):
                        return None, False
                    row_i, item_score = _best_chat_item_row(q)
                    pog_cols = (
                        [c for c in _dyn_pog_cols if c in _tdf.columns]
                        if _chat_all_word(q) else _chat_planogram_columns(q)
                    )
                    if row_i is None or item_score < 0.30:
                        return (
                            "หา item ที่ต้องการเพิ่มไม่เจอในข้อมูลหลัง filter ลองพิมพ์ ID หรือ Item Name ให้ชัดขึ้น",
                            False,
                        )
                    rk = _make_rk(row_i)
                    item_txt = _fmt_item(_tdf.loc[row_i])
                    if not pog_cols:
                        st.session_state[f"{p}_rs_chat_pending_add"] = {
                            "rows": [row_i],
                            "items": [item_txt],
                        }
                        return (
                            "หา planogramme ที่ต้องการเพิ่มไม่เจอ ลองพิมพ์ชื่อ planogramme / target ให้ชัดขึ้น "
                            "หรือพิมพ์ 'ทั้งหมด' เพื่อเพิ่มทั้ง row ของ item นี้",
                            False,
                        )
                    if _is_low_sales_item(row_i):
                        st.session_state[f"{p}_rs_chat_confirm_new"] = {
                            "rows": [row_i],
                            "pog_cols": pog_cols,
                            "items": [item_txt],
                        }
                        scope = "ทุก planogramme" if _chat_all_word(q) else ", ".join(pog_cols)
                        return (
                            f"{item_txt} is a Top 10% lowest seller. "
                            f"Are you sure you want to add this item to {scope}? "
                            "Reply 'ใช่' to confirm or 'ไม่' to cancel.",
                            False,
                        )
                    _chat_apply_new_rows([row_i], pog_cols)
                    scope = "ทุก planogramme" if _chat_all_word(q) else ", ".join(pog_cols)
                    return f"เพิ่ม {item_txt} เข้า {scope} เป็น New แล้ว", True

                def _detect_chat_intent(q: str) -> str:
                    s = _chat_norm(q)
                    ascii_s = _nca(s)
                    has_plan = ("planogram" in s) or ("planogram" in ascii_s)
                    has_max = any(k in s for k in ("มากที่สุด", "สูงสุด", "เยอะสุด", "most", "highest", "max"))
                    has_min = any(k in s for k in ("น้อยที่สุด", "ต่ำสุด", "least", "lowest", "min"))
                    if any(k in s for k in ("สรุป", "summary", "กี่ sku", "กี่สินค้า", "กี่ planogram")):
                        return "summary"
                    if any(k in s for k in ("new", "delete", "delist", "maintain", "status", "สถานะ")) and not any(k in s for k in ("ควรลบ", "ขายต่ำ", "ยอดขายต่ำ")):
                        return "action_counts"
                    if any(k in s for k in ("ช่องว่าง", "blank", "missing", "ว่าง")) and has_plan:
                        return "planogram_most_blanks"
                    if any(k in s for k in ("ควรลบ", "ขายต่ำ", "ยอดขายต่ำ", "สินค้าขายต่ำ", "lowest sales")):
                        return "lowest_sales_items"
                    if any(k in s for k in ("สินค้าไหนอยู่", "อยู่ใน planogram", "อยู่ในแพลน", "item in")):
                        return "item_most_planograms" if has_max else "item_fewest_planograms"
                    if any(k in s for k in ("ยอดขายรวม", "sum", "total sales", "รวมยอด")) and has_plan:
                        return "planogram_highest_total" if has_max else "planogram_lowest_total"
                    if has_plan and any(k in s for k in ("สินค้า", "product", "sku")):
                        return "planogram_fewest_products" if has_min else "planogram_most_products"
                    return _best_chat_intent(q)[0]

                def _planogram_counts_answer(kind: str) -> str:
                    m = _numeric_planogram_matrix(_tdf)
                    if m.empty:
                        return "ไม่พบ planogram columns สำหรับคำนวณ"
                    counts = m.notna().sum().sort_values(ascending=(kind == "min"))
                    if kind == "min":
                        positive = counts[counts > 0]
                        counts = positive if not positive.empty else counts
                    if counts.empty:
                        return "ไม่มีตัวเลขใน planogram columns"
                    pog, val = counts.index[0], int(counts.iloc[0])
                    label = "มากที่สุด" if kind == "max" else "น้อยที่สุด"
                    return f"Planogramme ที่มีสินค้า{label}: {pog} ({val:,} item)"

                def _planogram_total_answer(kind: str) -> str:
                    m = _numeric_planogram_matrix(_tdf)
                    if m.empty:
                        return "ไม่พบ planogram columns สำหรับคำนวณยอดรวม"
                    totals = m.sum(skipna=True).sort_values(ascending=(kind == "min"))
                    if totals.empty:
                        return "ไม่มีตัวเลขใน planogram columns"
                    pog, val = totals.index[0], float(totals.iloc[0])
                    label = "มากที่สุด" if kind == "max" else "น้อยที่สุด"
                    return f"Planogramme ที่มียอดขายรวม{label}: {pog} ({val:,.2f})"

                def _item_planogram_answer(kind: str) -> str:
                    m = _numeric_planogram_matrix(_tdf)
                    if m.empty:
                        return "ไม่พบ planogram columns สำหรับคำนวณ"
                    counts = m.notna().sum(axis=1)
                    if kind == "min":
                        counts = counts[counts > 0]
                    if counts.empty:
                        return "ไม่มีสินค้าใน planogram columns"
                    target = counts.max() if kind == "max" else counts.min()
                    rows = counts[counts == target].index[:5]
                    label = "มากที่สุด" if kind == "max" else "น้อยที่สุด"
                    items = [f"{_fmt_item(_tdf.loc[i])} ({int(target):,} planogramme)" for i in rows]
                    return "สินค้าที่อยู่ใน planogramme" + label + ": " + "; ".join(items)

                def _planogram_blanks_answer() -> str:
                    cols = [c for c in _dyn_pog_cols if c in _tdf.columns]
                    if not cols:
                        return "ไม่พบ planogram columns สำหรับนับช่องว่าง"
                    blanks = {c: int(_tdf[c].map(_num_cell).isna().sum()) for c in cols}
                    ser = pd.Series(blanks).sort_values(ascending=False)
                    return f"Planogramme ที่มีช่องว่างมากที่สุด: {ser.index[0]} ({int(ser.iloc[0]):,} ช่อง)"

                def _action_counts_answer() -> str:
                    vals = []
                    for _, acts in _pog_actions.items():
                        vals.extend([v for v in acts.values() if v])
                    status_counts = (
                        _tdf_display["Status"].astype(str).value_counts().to_dict()
                        if "Status" in _tdf_display.columns else {}
                    )
                    action_counts = {a: vals.count(a) for a in ("New", "Delete", "Delist", "Keep")}
                    return (
                        "Action counts: "
                        + ", ".join(f"{k}={v:,}" for k, v in action_counts.items())
                        + " | Status: "
                        + ", ".join(f"{k}={v:,}" for k, v in status_counts.items())
                    )

                def _summary_answer() -> str:
                    return f"DG/filter ปัจจุบันมี {len(_tdf):,} SKU/item และ {len(_dyn_pog_cols):,} planogramme"

                def _lowest_sales_items_answer() -> str:
                    vol_col = _find_col(_tdf, [
                        "TH_Tot_Sales_Volume_52WK",
                        "TH Tot Sales Volume 52WK",
                    ])
                    fc_col = _find_col(_tdf, [
                        "ForecastSales",
                        "Forecast Sales",
                    ])
                    mapped_col = _find_col(_tdf, [
                        "Avg Units 52wk/ Forecast new item sales",
                    ])
                    if not vol_col:
                        vol_col = mapped_col
                    if not fc_col:
                        fc_col = mapped_col
                    if not vol_col and not fc_col:
                        return "ไม่พบ column TH_Tot_Sales_Volume_52WK หรือ ForecastSales ในข้อมูลหลัง filter"
                    rows = []
                    for row_idx, row in _tdf.iterrows():
                        vol = _num_cell(row.get(vol_col)) if vol_col else None
                        fc = _num_cell(row.get(fc_col)) if fc_col else None
                        metric, source = (vol, "Volume52WK") if vol is not None else (fc, "ForecastSales")
                        if source == "Volume52WK" and vol_col == mapped_col:
                            source = "Volume52WK/Forecast mapped"
                        if metric is None:
                            continue
                        rows.append((metric, source, _fmt_item(row), row_idx))
                    if not rows:
                        return "ไม่พบตัวเลขยอดขาย/forecast สำหรับคำนวณสินค้าแนะนำให้ลบ"
                    rows = sorted(rows, key=lambda x: x[0])[:5]
                    st.session_state[f"{p}_rs_chat_recent_items"] = [
                        {"row_i": int(row_i), "item": item, "metric": float(val), "source": src}
                        for val, src, item, row_i in rows
                    ]
                    lines = [f"{idx}. {item} = {val:,.2f} ({src})" for idx, (val, src, item, _) in enumerate(rows, 1)]
                    return "สินค้าที่มียอดขายต่ำที่สุด 5 ตัวแรก:\n" + "\n".join(lines)

                def _answer_chat(q: str) -> str:
                    intent = _detect_chat_intent(q)
                    if intent == "planogram_most_products":
                        return _planogram_counts_answer("max")
                    if intent == "planogram_fewest_products":
                        return _planogram_counts_answer("min")
                    if intent == "lowest_sales_items":
                        return _lowest_sales_items_answer()
                    if intent == "planogram_highest_total":
                        return _planogram_total_answer("max")
                    if intent == "planogram_lowest_total":
                        return _planogram_total_answer("min")
                    if intent == "item_most_planograms":
                        return _item_planogram_answer("max")
                    if intent == "item_fewest_planograms":
                        return _item_planogram_answer("min")
                    if intent == "planogram_most_blanks":
                        return _planogram_blanks_answer()
                    if intent == "action_counts":
                        return _action_counts_answer()
                    if intent == "summary":
                        return _summary_answer()
                    return "ยังไม่เข้าใจคำถามนี้ ลองถามเช่น: planogramme ไหนมีสินค้ามากที่สุด, ควรลบสินค้าไหน, สรุป DG นี้"

                def _intent_label(intent: str) -> str:
                    return _CHAT_INTENT_LIBRARY.get(intent, {}).get("label", "คำถามอื่น")

                def _answer_intent(intent: str) -> str:
                    if intent == "planogram_most_products":
                        return _planogram_counts_answer("max")
                    if intent == "planogram_fewest_products":
                        return _planogram_counts_answer("min")
                    if intent == "lowest_sales_items":
                        return _lowest_sales_items_answer()
                    if intent == "planogram_highest_total":
                        return _planogram_total_answer("max")
                    if intent == "planogram_lowest_total":
                        return _planogram_total_answer("min")
                    if intent == "item_most_planograms":
                        return _item_planogram_answer("max")
                    if intent == "item_fewest_planograms":
                        return _item_planogram_answer("min")
                    if intent == "planogram_most_blanks":
                        return _planogram_blanks_answer()
                    if intent == "action_counts":
                        return _action_counts_answer()
                    if intent == "summary":
                        return _summary_answer()
                    return "ยังไม่เข้าใจคำถามนี้"

                _chat_key = f"{p}_rs_chat_history"
                _chat_pending_key = f"{p}_rs_chat_pending"
                _chat_input_seq_key = f"{p}_rs_chat_input_seq"
                _chat_recent_items_key = f"{p}_rs_chat_recent_items"
                _chat_recent_delete_key = f"{p}_rs_chat_recent_delete"
                _chat_pending_add_key = f"{p}_rs_chat_pending_add"
                _chat_confirm_new_key = f"{p}_rs_chat_confirm_new"
                st.session_state.setdefault(_chat_key, [])
                st.session_state.setdefault(_chat_input_seq_key, 0)
                st.markdown("""
<style>
div[data-testid="stPopover"] {
    position: fixed !important;
    right: 30px !important;
    bottom: 30px !important;
    left: auto !important;
    top: auto !important;
    z-index: 999999 !important;
}
div[data-testid="stPopover"] > button,
div[data-testid="stPopover"] button {
    width: 78px;
    height: 78px;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 34px;
    font-weight: 800;
    box-shadow: 0 16px 36px rgba(88,43,180,.38);
    background: linear-gradient(135deg, #7C3AED 0%, #C026D3 55%, #14B8A6 100%);
    color: #fff;
    border: 4px solid rgba(255,255,255,.94);
    min-width: 78px;
    padding: 0;
}
div[data-testid="stPopover"] > button:hover,
div[data-testid="stPopover"] button:hover {
    transform: translateY(-2px) scale(1.03);
    color: #fff;
    box-shadow: 0 18px 42px rgba(88,43,180,.48);
}
div[data-testid="stPopover"] > div,
div[data-testid="stPopover"] [data-baseweb="popover"] {
    width: min(420px, calc(100vw - 28px));
}
.rs-chat-head {
    margin: -18px -18px 16px -18px;
    padding: 16px 18px;
    border-radius: 12px 12px 0 0;
    background: linear-gradient(135deg, #7C3AED 0%, #C026D3 100%);
    color: #fff;
}
.rs-chat-title {
    display: flex;
    align-items: center;
    gap: 10px;
    font-size: 18px;
    font-weight: 800;
}
.rs-chat-avatar {
    width: 34px;
    height: 34px;
    border-radius: 50%;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    background: rgba(255,255,255,.22);
    border: 1px solid rgba(255,255,255,.5);
}
.rs-chat-online {
    margin-left: 44px;
    margin-top: -4px;
    color: rgba(255,255,255,.84);
    font-size: 12px;
}
.rs-chat-online::before {
    content: "";
    display: inline-block;
    width: 8px;
    height: 8px;
    margin-right: 6px;
    border-radius: 50%;
    background: #7CFF9B;
}
.rs-chat-log {
    max-height: 360px;
    overflow-y: auto;
    padding: 4px 2px 12px 2px;
}
.rs-bot-row, .rs-user-row {
    display: flex;
    margin: 10px 0;
}
.rs-bot-row {
    justify-content: flex-start;
}
.rs-user-row {
    justify-content: flex-end;
}
.rs-bubble {
    max-width: 82%;
    padding: 10px 13px;
    border-radius: 18px;
    line-height: 1.35;
    font-size: 14px;
    white-space: pre-wrap;
}
.rs-bot {
    background: #F4F5F7;
    color: #1F2937;
    border-top-left-radius: 6px;
}
.rs-user {
    background: #B963D1;
    color: #fff;
    border-top-right-radius: 6px;
}
.rs-suggest {
    border: 1px solid #E7D7F5;
    background: #FCF7FF;
    border-radius: 16px;
    padding: 12px;
    margin: 10px 0 14px 0;
}
</style>
""", unsafe_allow_html=True)

                def _chat_delete_word(q: str) -> bool:
                    s = _chat_norm(q)
                    return any(k in s for k in ("delete", "ลบ", "เอาออก"))

                def _chat_all_word(q: str) -> bool:
                    s = _chat_norm(q)
                    compact = _chat_cmp(q)
                    return (
                        any(k in s for k in ("all", "ทุก planogram", "ทุกแพลน", "ทั้งแถว"))
                        or _chat_cmp("ทั้งหมด") in compact
                        or _chat_cmp("ทั้งแถว") in compact
                        or _chat_cmp("ทุกแพลน") in compact
                        or _chat_cmp("ทุก planogram") in compact
                    )

                def _chat_yes_word(q: str) -> bool:
                    compact = _chat_cmp(q)
                    return any(k in compact for k in ("yes", "confirm", _chat_cmp("ใช่"), _chat_cmp("ยืนยัน"), _chat_cmp("ตกลง")))

                def _chat_no_word(q: str) -> bool:
                    compact = _chat_cmp(q)
                    return any(k in compact for k in ("no", "cancel", _chat_cmp("ไม่"), _chat_cmp("ยกเลิก")))

                def _chat_answer_confirm_new(q: str):
                    pending = st.session_state.get(_chat_confirm_new_key)
                    if not pending:
                        return None
                    if _chat_no_word(q):
                        st.session_state.pop(_chat_confirm_new_key, None)
                        return "ยกเลิกการเพิ่ม item แล้ว"
                    if not _chat_yes_word(q):
                        return "ตอบ 'ใช่' เพื่อยืนยันเพิ่ม item นี้ หรือ 'ไม่' เพื่อยกเลิก"
                    rows = [int(x) for x in pending.get("rows", []) if x in _tdf.index]
                    pog_cols = [c for c in pending.get("pog_cols", []) if c in _tdf.columns]
                    st.session_state.pop(_chat_confirm_new_key, None)
                    if not rows or not pog_cols:
                        return "ไม่พบ item หรือ planogramme สำหรับเพิ่มแล้ว"
                    changed = _chat_apply_new_rows(rows, pog_cols)
                    scope = "ทุก planogramme" if len(pog_cols) == len([c for c in _dyn_pog_cols if c in _tdf.columns]) else ", ".join(pog_cols)
                    return f"ยืนยันแล้ว เพิ่มเป็น New {len(changed):,} ช่อง สำหรับ {scope}"

                def _chat_answer_pending_add_scope(q: str):
                    pending = st.session_state.get(_chat_pending_add_key)
                    if not pending:
                        return None
                    rows = [int(x) for x in pending.get("rows", []) if x in _tdf.index]
                    if not rows:
                        st.session_state.pop(_chat_pending_add_key, None)
                        return "ไม่พบ item เดิมในข้อมูลหลัง filter แล้ว"
                    if _chat_all_word(q):
                        pog_cols = [c for c in _dyn_pog_cols if c in _tdf.columns]
                    else:
                        pog_cols = _chat_planogram_columns(q)
                    if not pog_cols:
                        return "ยังหา planogramme ไม่เจอ กรุณาพิมพ์ชื่อ planogramme / target เช่น 5012 หรือพิมพ์ 'ทั้งหมด'"
                    st.session_state.pop(_chat_pending_add_key, None)
                    low_rows = [r for r in rows if _is_low_sales_item(r)]
                    if low_rows:
                        st.session_state[_chat_confirm_new_key] = {
                            "rows": rows,
                            "pog_cols": pog_cols,
                            "items": [_fmt_item(_tdf.loc[r]) for r in rows],
                        }
                        item_txt = ", ".join(_fmt_item(_tdf.loc[r]) for r in low_rows[:3])
                        scope = "ทุก planogramme" if _chat_all_word(q) else ", ".join(pog_cols)
                        return (
                            f"{item_txt} is a Top 10% lowest seller. "
                            f"Are you sure you want to add this item to {scope}? "
                            "Reply 'ใช่' to confirm or 'ไม่' to cancel."
                        )
                    changed = _chat_apply_new_rows(rows, pog_cols)
                    scope = "ทุก planogramme" if _chat_all_word(q) else ", ".join(pog_cols)
                    return f"เพิ่มเป็น New แล้ว {len(changed):,} ช่อง สำหรับ {scope}"

                def _chat_start_recent_delete(q: str):
                    if not _chat_delete_word(q):
                        return None
                    recent = st.session_state.get(_chat_recent_items_key, [])
                    if not recent:
                        return None
                    rows = [int(x.get("row_i")) for x in recent if x.get("row_i") is not None]
                    rows = [r for r in rows if r in _tdf.index]
                    if not rows:
                        return "ไม่พบ item ล่าสุดในข้อมูลหลัง filter แล้ว"
                    st.session_state[_chat_recent_delete_key] = {"rows": rows, "items": recent}
                    return (
                        "ต้องการให้ลบ item ที่ตอบล่าสุดออกจาก planogramme ไหน? "
                        "พิมพ์ชื่อ planogramme / target เช่น 5029, 5086 หรือพิมพ์ 'ทั้งหมด'"
                    )

                def _chat_answer_recent_delete_scope(q: str):
                    pending = st.session_state.get(_chat_recent_delete_key)
                    if not pending:
                        return None
                    rows = [int(x) for x in pending.get("rows", []) if x in _tdf.index]
                    if not rows:
                        st.session_state.pop(_chat_recent_delete_key, None)
                        return "ไม่พบ item ล่าสุดในข้อมูลหลัง filter แล้ว"
                    if _chat_all_word(q):
                        pog_cols = [c for c in _dyn_pog_cols if c in _tdf.columns]
                    else:
                        pog_cols = _chat_planogram_columns(q)
                    if not pog_cols:
                        return "ยังหา planogramme ไม่เจอ กรุณาพิมพ์ชื่อ planogramme / target หรือพิมพ์ 'ทั้งหมด'"
                    changed, top_pending = _chat_apply_delete_rows(rows, pog_cols)
                    st.session_state.pop(_chat_recent_delete_key, None)
                    changed_count = len(changed)
                    top_count = len(top_pending)
                    scope = "ทุก planogramme" if _chat_all_word(q) else ", ".join(pog_cols)
                    if top_count:
                        msg = (
                            f"ตั้งค่า Delete แล้ว {changed_count:,} ช่อง สำหรับ {scope}. "
                            f"มี {top_count:,} ช่องเป็น Top 10% best seller ต้องกด Confirm Delete หรือ Cancel ด้านบนตารางก่อน"
                        )
                    else:
                        msg = (
                            f"ตั้งค่า Delete แล้ว {changed_count:,} ช่อง สำหรับ {scope}. "
                            "dropdown และ Status จะเปลี่ยนตาม row"
                        )
                    return msg

                def _render_chat_dialog():
                    st.markdown("""
<div class="rs-chat-head">
  <div class="rs-chat-title"><span class="rs-chat-avatar">&#129302;</span><span>RangeBot</span></div>
  <div class="rs-chat-online">Online now</div>
</div>
""", unsafe_allow_html=True)
                    st.markdown('<div class="rs-chat-log">', unsafe_allow_html=True)
                    if not st.session_state[_chat_key]:
                        st.markdown(
                            '<div class="rs-bot-row"><div class="rs-bubble rs-bot">'
                            'สวัสดีครับ ถามข้อมูลจาก rangesheet หลัง apply filter ได้เลย'
                            '</div></div>',
                            unsafe_allow_html=True,
                        )
                    for _msg in st.session_state[_chat_key][-6:]:
                        _uq = _html.escape(str(_msg.get("q", "")))
                        _ba = _html.escape(str(_msg.get("a", "")))
                        st.markdown(
                            f'<div class="rs-user-row"><div class="rs-bubble rs-user">{_uq}</div></div>'
                            f'<div class="rs-bot-row"><div class="rs-bubble rs-bot">{_ba}</div></div>',
                            unsafe_allow_html=True,
                        )
                    st.markdown('</div>', unsafe_allow_html=True)
                    _pending = st.session_state.get(_chat_pending_key)
                    if _pending:
                        _score = _pending.get("score")
                        _score_txt = f" ({_score:.0%} match)" if isinstance(_score, (int, float)) else ""
                        st.markdown(
                            '<div class="rs-suggest">'
                            f"คุณต้องการถามคำถามนี้ไหม:<br><b>{_html.escape(_intent_label(_pending['intent']))}</b>"
                            f"{_html.escape(_score_txt)}"
                            '</div>',
                            unsafe_allow_html=True,
                        )
                        _yc, _nc = st.columns([1, 1])
                        if _yc.button("ใช่", key=f"{p}_rs_chat_yes", use_container_width=True):
                            ans = _answer_intent(_pending["intent"])
                            st.session_state[_chat_key].append({"q": _pending["q"], "a": ans})
                            st.session_state.pop(_chat_pending_key, None)
                            st.session_state[_chat_input_seq_key] += 1
                            st.rerun()
                        if _nc.button("ไม่ใช่", key=f"{p}_rs_chat_no", use_container_width=True):
                            st.session_state.pop(_chat_pending_key, None)
                            st.session_state[_chat_input_seq_key] += 1
                            st.warning("คุณต้องการถามว่าอะไร? กรุณาพิมพ์คำถามใหม่")
                    _input_key = f"{p}_rs_chat_input_{st.session_state[_chat_input_seq_key]}"
                    _q = st.text_input("Ask chatbot", key=_input_key, label_visibility="collapsed", placeholder="Type your rangesheet question...")
                    _qa, _qb = st.columns([1, 1])
                    if _qa.button("Ask", key=f"{p}_rs_chat_ask", use_container_width=True) and _q.strip():
                        _follow_answer = _chat_answer_confirm_new(_q)
                        if _follow_answer is None:
                            _follow_answer = _chat_answer_pending_add_scope(_q)
                        if _follow_answer is None:
                            _follow_answer = _chat_answer_recent_delete_scope(_q)
                        if _follow_answer is not None:
                            st.session_state[_chat_key].append({"q": _q.strip(), "a": _follow_answer})
                            st.session_state[_chat_input_seq_key] += 1
                            st.rerun()
                        else:
                            _cmd_answer, _cmd_changed = _chat_delete_item_from_planogram(_q)
                            if _cmd_answer is None:
                                _cmd_answer, _cmd_changed = _chat_add_item_to_planogram(_q)
                            if _cmd_answer is not None:
                                st.session_state[_chat_key].append({"q": _q.strip(), "a": _cmd_answer})
                                st.session_state.pop(_chat_pending_key, None)
                                st.session_state[_chat_input_seq_key] += 1
                                st.rerun()
                            else:
                                _recent_delete_answer = _chat_start_recent_delete(_q)
                                if _recent_delete_answer is not None:
                                    st.session_state[_chat_key].append({"q": _q.strip(), "a": _recent_delete_answer})
                                    st.session_state[_chat_input_seq_key] += 1
                                    st.rerun()
                                else:
                                    intent = _detect_chat_intent(_q)
                                    if intent == "unknown":
                                        st.warning("คุณต้องการถามว่าอะไร? ลองพิมพ์ใหม่อีกครั้ง")
                                    else:
                                        st.session_state[_chat_pending_key] = {
                                            "q": _q.strip(),
                                            "intent": intent,
                                            "score": _score_chat_intent(_q, intent),
                                        }
                                        st.session_state[_chat_input_seq_key] += 1
                                        st.rerun()
                    if _qb.button("Clear chat", key=f"{p}_rs_chat_clear", use_container_width=True):
                        st.session_state[_chat_key] = []
                        st.session_state.pop(_chat_pending_key, None)
                        st.session_state.pop(_chat_recent_delete_key, None)
                        st.session_state.pop(_chat_pending_add_key, None)
                        st.session_state.pop(_chat_confirm_new_key, None)
                        st.session_state[_chat_input_seq_key] += 1
                        st.rerun()

                if hasattr(st, "popover"):
                    with st.popover("💬", help="Open Rangesheet chat"):
                        _render_chat_dialog()
                else:
                    with st.expander("Rangesheet chatbot", expanded=False):
                        _render_chat_dialog()

                _fmt_pog = JsCode("""
function(params) {
    if (params.value == null || params.value === '') return '';
    if (['Keep','Delete','New','Delist'].indexOf(String(params.value)) >= 0)
        return String(params.value);
    var n = Number(params.value);
    if (isNaN(n)) return String(params.value);
    return n.toLocaleString('en-US', {minimumFractionDigits:0, maximumFractionDigits:2});
}""")
                _highlight_blanks = bool(st.session_state.get(_blank_hl_key))
                _search_q_js = _json.dumps(_search_q_l)
                _search_qn_js = _json.dumps(_search_q_n)
                _style_pog = JsCode("""
function(params) {
    var highlightBlanks = __HL__;
    var q = __Q__;
    var qn = __QN__;
    var base = {backgroundColor:'', color:'', fontWeight:'normal', textAlign:'center'};
    function norm(v) {
        return String(v == null ? '' : v).toLowerCase().replace(/[^a-z0-9]/g, '');
    }
    var vText = (params.value == null ? '' : String(params.value)).toLowerCase();
    var matched = q && vText.indexOf(q) >= 0;
    var fieldMatched = qn && norm(params.colDef && params.colDef.field).indexOf(qn) >= 0;
    var s = {
        'Keep':   {backgroundColor:'#E8F5E9',color:'#2E7D32',fontWeight:'bold',textAlign:'center'},
        'Delete': {backgroundColor:'#FFEBEE',color:'#C62828',fontWeight:'bold',textAlign:'center'},
        'New':    {backgroundColor:'#E3F2FD',color:'#1565C0',fontWeight:'bold',textAlign:'center'},
        'Delist': {backgroundColor:'#F5F5F5',color:'#757575',fontWeight:'bold',textAlign:'center'},
    };
    if (fieldMatched) {
        return {backgroundColor:'#FFF59D',color:'#111',fontWeight:'800',textAlign:'center'};
    }
    if (matched) {
        return {backgroundColor:'#FFF59D',color:'#111',fontWeight:'800',textAlign:'center'};
    }
    if (highlightBlanks) {
        var v = params.value;
        var blank = (v == null || String(v).trim() === '' ||
                     String(v) === 'nan' || String(v) === 'None');
        if (blank) return {backgroundColor:'#FFF59D',color:'#000',textAlign:'center'};
    }
    return s[String(params.value)] || base;
}""".replace("__HL__", "true" if _highlight_blanks else "false").replace("__Q__", _search_q_js).replace("__QN__", _search_qn_js))
                _style_data = JsCode("""
function(params) {
    var q = __Q__;
    var avgCol = __AVG_U_COL__;
    var v = params.value;
    var blank = (v == null || String(v).trim() === '' ||
                 String(v) === 'nan' || String(v) === 'None');
    var field = String((params.colDef && params.colDef.field) || '');
    if (params.data && params.data['__rs_added_row'] && field === avgCol) {
        var sAvg = String(v == null ? '' : v).replace(/,/g, '').trim();
        if (!sAvg || isNaN(Number(sAvg))) {
            return {backgroundColor:'#FFF8E1', color:'#8A5A00'};
        }
    }
    var matched = q && !blank && String(v).toLowerCase().indexOf(q) >= 0;
    if (matched) {
        return {backgroundColor:'#FFF59D',color:'#111',fontWeight:'800'};
    }
    return (__HL__ && blank) ? {backgroundColor:'#FFF59D',color:'#000'} : null;
}""".replace("__HL__", "true" if _highlight_blanks else "false").replace("__Q__", _search_q_js).replace("__AVG_U_COL__", _json.dumps(_AVG_U_STD)))
                _fmt_status = JsCode("""
function(params) {
    if (params.value == null) return '';
    return String(params.value).toUpperCase();
}""")
                _style_status = JsCode("""
function(params) {
    var v = params.value == null ? '' : String(params.value).trim().toUpperCase();
    var q = __Q__;
    var matched = q && v.toLowerCase().indexOf(q) >= 0;
    var base = {
        textAlign: 'center'
    };
    function mark(style) {
        if (matched) {
            style.backgroundColor = '#FFF59D';
            style.color = '#111';
        }
        return style;
    }
    if (v === 'DELETE SOME') {
        return mark(Object.assign(base, {backgroundColor:'#FFF3E0', color:'#E65100'}));
    }
    if (v === 'DELETE ALL') {
        return mark(Object.assign(base, {backgroundColor:'#FFEBEE', color:'#C62828'}));
    }
    if (v === 'NEW SOME') {
        return mark(Object.assign(base, {backgroundColor:'#FFF59D', color:'#6D4C00'}));
    }
    if (v === 'NEWNEW') {
        return mark(Object.assign(base, {backgroundColor:'#E8F5E9', color:'#1B5E20'}));
    }
    return mark(base);
}""".replace("__Q__", _search_q_js))
                _fmt_avg_u = JsCode("""
function(params) {
    if (params.value == null || params.value === '') return '';
    var n = Number(params.value);
    if (isNaN(n)) return String(params.value);
    var s = n.toFixed(5);
    s = s.replace(/(\\.[0-9]*?)0+$/, '$1').replace(/\\.$/, '');
    return s;
}""")
                _editable_new_row = JsCode("""
function(params) {
    return !!(params.data && params.data['__rs_added_row']);
}""")
                _pog_cols_js_for_setter = _json.dumps([str(c) for c in _dyn_pog_cols])
                _orig_action_values_js = _json.dumps(_orig_action_values, ensure_ascii=False)
                _pog_action_editor_params = JsCode("""
function(params) {
    var actions = {'Keep':true, 'Delete':true, 'New':true, 'Delist':true};
    var origMap = __ORIG_MAP__;
    var field = String((params.colDef && params.colDef.field) || '');
    var rowKey = String((params.data && params.data['__rs_row_key']) || '');
    var origKey = rowKey + '\\u241f' + field;
    var origValue = Object.prototype.hasOwnProperty.call(origMap, origKey) ? String(origMap[origKey]) : '';
    function isNumericCell(v) {
        if (v === null || v === undefined) return false;
        var s = String(v).replace(/,/g, '').trim();
        if (!s || s === 'nan' || s === 'None' || actions[s]) return false;
        return !isNaN(Number(s));
    }

    if (params.value === 'Delete') return {values: origValue ? ['Delete', origValue] : ['Delete', '']};
    if (params.value === 'New') return {values: origValue ? ['New', origValue] : ['New', '']};
    return isNumericCell(params.value) ? {values: ['Delete']} : {values: ['New']};
}""".replace("__ORIG_MAP__", _orig_action_values_js))
                _status_action_editor_params = JsCode("""
function(params) {
    return {values: ['MAINTAIN', 'DELETE ALL']};
}""")
                _delete_guard = JsCode("""
function(params) {
    var oldValue = params.oldValue;
    var newValue = params.newValue;
    var pogCols = __POG_COLS__;
    function updateRowStatus() {
        var actions = {'Keep':true, 'Delete':true, 'New':true, 'Delist':true};
        function isNum(v) {
            if (v === null || v === undefined) return false;
            var s = String(v).replace(/,/g, '').trim();
            if (!s || s === 'nan' || s === 'None' || actions[s]) return false;
            return !isNaN(Number(s));
        }
        var deleteCols = {}, newCols = {}, existingCols = {};
        var deleteCount = 0, newCount = 0, pogCount = 0, hasDelist = false;
        for (var i = 0; i < pogCols.length; i++) {
            var c = String(pogCols[i]);
            if (!(c in params.data)) continue;
            pogCount++;
            var v = params.data[c];
            if (isNum(v)) existingCols[c] = true;
            if (v === 'Delist') hasDelist = true;
            if (v === 'Delete') {
                deleteCols[c] = true;
                deleteCount++;
            }
            if (v === 'New') {
                newCols[c] = true;
                newCount++;
            }
        }
        var status = 'MAINTAIN';
        if (hasDelist) {
            status = 'Inactive';
        } else if (deleteCount > 0 && newCount > 0) {
            status = 'DELETE SOME';
        } else if (deleteCount > 0) {
            var deleteAll = true;
            var hasDeleteScope = false;
            for (var dc in deleteCols) existingCols[dc] = true;
            for (var ec in existingCols) {
                hasDeleteScope = true;
                if (!deleteCols[ec]) {
                    deleteAll = false;
                    break;
                }
            }
            status = (hasDeleteScope && deleteAll) ? 'DELETE ALL' : 'DELETE SOME';
        } else if (newCount > 0) {
            status = (pogCount > 0 && newCount >= pogCount) ? 'NEWNEW' : 'NEW SOME';
        }
        params.data['Status'] = status;
        var wf = 0;
        for (var wc = 0; wc < pogCols.length; wc++) {
            var pc = String(pogCols[wc]);
            if (!(pc in params.data)) continue;
            var pv = params.data[pc];
            if (isNum(pv) || pv === 'New') wf++;
        }
        params.data['Check Range To-be Waterfall'] = wf;
        try {
            params.api.refreshCells({
                rowNodes: [params.node],
                columns: ['Status', 'Check Range To-be Waterfall'],
                force: true
            });
        } catch (e) {}
    }
    if (String(newValue) !== 'Delete') {
        params.data[params.colDef.field] = newValue;
        updateRowStatus();
        return true;
    }
    var actions = {'Keep':true, 'Delete':true, 'New':true, 'Delist':true};
    function toNum(v) {
        if (v === null || v === undefined) return null;
        var s = String(v).replace(/,/g, '').trim();
        if (!s || s === 'nan' || s === 'None' || actions[s]) return null;
        var n = Number(s);
        return isNaN(n) ? null : n;
    }
    var current = toNum(oldValue);
    if (current === null) current = toNum(params.data[params.colDef.field]);
    if (current === null) {
        params.data[params.colDef.field] = newValue;
        return true;
    }
    var nums = [];
    for (var i = 0; i < pogCols.length; i++) {
        var n = toNum(params.data[pogCols[i]]);
        if (n !== null) nums.push(n);
    }
    if (!nums.length) {
        params.data[params.colDef.field] = newValue;
        return true;
    }
    nums.sort(function(a, b) { return b - a; });
    var cutoffIndex = Math.max(1, Math.ceil(nums.length * 0.10)) - 1;
    var cutoff = nums[cutoffIndex];
    if (current >= cutoff) {
        var item = params.data['Item Name'] || params.data['ID'] || '';
        var ok = window.confirm(
            'Item ' + item + ' is a Top 10% best seller in this row.\\n' +
            'Do you still want to delete it?'
        );
        if (!ok) {
            params.data[params.colDef.field] = oldValue;
            updateRowStatus();
            return false;
        }
    }
    params.data[params.colDef.field] = newValue;
    updateRowStatus();
    return true;
}
""".replace("__POG_COLS__", _pog_cols_js_for_setter))

                # ── CSS: vertical pog headers + bottom-align horizontal headers ───
                # AG Grid v16 DOM: .ag-header-cell → .ag-cell-label-container
                #   → .ag-header-cell-label → .ag-header-cell-text (rotate this span)
                _custom_css = {
                    ".ag-header-cell.pog-vertical .ag-cell-label-container": {
                        "height":          "100%",
                        "padding":         "4px 0",
                        "justify-content": "center",
                        "align-items":     "center",
                    },
                    ".ag-header-cell.pog-vertical .ag-header-cell-label": {
                        "height":          "100%",
                        "justify-content": "center",
                        "align-items":     "center",
                    },
                    ".ag-header-cell.pog-vertical .ag-header-cell-text": {
                        "writing-mode":  "vertical-rl",
                        "transform":     "rotate(180deg)",
                        "white-space":   "normal",
                        "overflow":      "visible",
                        "text-overflow": "unset",
                    },
                    ".ag-header-cell.pog-vertical": {
                        "padding": "0 !important",
                    },
                    ".ag-header-cell.search-header-match": {
                        "background-color": "#FFF59D !important",
                        "font-weight": "800 !important",
                    },
                    ".ag-header-cell.search-header-match .ag-header-cell-text": {
                        "color": "#111 !important",
                        "font-weight": "800 !important",
                    },
                    ".ag-header-cell:not(.pog-vertical) .ag-cell-label-container": {
                        "align-items":    "flex-end",
                        "padding-bottom": "4px",
                        "width": "100%",
                    },
                    ".ag-header-cell:not(.pog-vertical) .ag-header-cell-label": {
                        "justify-content": "center",
                        "width": "100%",
                        "min-width": "0",
                    },
                    ".ag-header-cell:not(.pog-vertical) .ag-header-cell-text": {
                        "display": "block",
                        "width": "100%",
                        "max-width": "100%",
                        "white-space": "normal",
                        "overflow": "hidden",
                        "text-overflow": "unset",
                        "word-break": "normal",
                        "overflow-wrap": "break-word",
                        "text-align": "center",
                        "line-height": "1.18",
                    },
                    ".ag-cell.pog-value-cell": {
                        "font-size": "10px !important",
                        "font-variant-numeric": "tabular-nums",
                        "padding-left": "1px !important",
                        "padding-right": "1px !important",
                        "text-align": "center",
                        "white-space": "nowrap",
                        "text-overflow": "clip !important",
                    },
                    ".ag-cell.pog-value-cell .ag-cell-value": {
                        "overflow": "visible !important",
                        "text-overflow": "clip !important",
                    },
                    ".ag-overlay": {
                        "display": "none !important",
                        "background": "transparent !important",
                        "pointer-events": "none !important",
                    },
                    ".ag-overlay-loading-wrapper": {
                        "display": "none !important",
                        "background": "transparent !important",
                        "pointer-events": "none !important",
                    },
                    ".ag-overlay-loading-center": {
                        "display": "none !important",
                    },
                    ".ag-root-wrapper, .ag-root, .ag-body-viewport, .ag-center-cols-container, .ag-row, .ag-cell": {
                        "opacity": "1 !important",
                    },
                }

                # ── Build GridOptions ─────────────────────────────────────────────
                # min_column_width must be passed explicitly — omitting it resets to
                # the parameter default (5), overriding the 110 set by from_dataframe.
                gb = GridOptionsBuilder.from_dataframe(_tdf_display)
                gb.configure_default_column(
                    min_column_width=80,
                    sortable=True, filter=False, resizable=True,
                    editable=False, wrapHeaderText=True,
                    suppressMovable=False,
                    menuTabs=["generalMenuTab", "columnsMenuTab"],
                )
                def _header_class(_gc: str, base: str = "") -> str:
                    _classes = [base] if base else []
                    if _search_q_l and (
                        _search_q_l in str(_gc).lower()
                        or (_search_q_n and _search_q_n in _nca(str(_gc)))
                    ):
                        _classes.append("search-header-match")
                    return " ".join(_classes) if _classes else None

                for _gc in _tdf_display.columns:
                    _hidden = False   # visibility controlled by grid sidebar/menu; not here
                    if _gc == _ROW_KEY_COL:
                        gb.configure_column(
                            _gc, hide=True, suppressColumnsToolPanel=True,
                            suppressMovable=True,
                        )
                    elif _gc == _EDIT_COL:
                        gb.configure_column(
                            _gc, hide=True, suppressColumnsToolPanel=True,
                            suppressMovable=True,
                        )
                    elif _gc == _NEW_ROW_COL:
                        gb.configure_column(
                            _gc, hide=True, suppressColumnsToolPanel=True,
                            suppressMovable=True,
                        )
                    elif _gc in set(_STICKY):
                        gb.configure_column(
                            _gc, pinned="left",
                            width=_COL_W.get(_gc, 130), minWidth=_COL_W.get(_gc, 80),
                            editable=True, hide=_hidden, suppressMovable=True,
                            wrapHeaderText=True,
                            headerClass=_header_class(_gc),
                            cellStyle=_style_data,
                        )
                    elif _gc in _dyn_pog_set:
                        # Pog columns: wide enough for dropdown editor and action labels.
                        # type=[] clears any numericColumn type that from_dataframe may
                        # have inferred so our valueFormatter is never overridden.
                        gb.configure_column(
                            _gc,
                            width=72, minWidth=64, maxWidth=80,
                            headerClass=_header_class(_gc, "pog-vertical"),
                            valueFormatter=_fmt_pog,
                            cellStyle=_style_pog,
                            cellClass="pog-value-cell",
                            editable=True,
                            cellEditor="agSelectCellEditor",
                            cellEditorParams=_pog_action_editor_params,
                            valueSetter=_delete_guard,
                            cellDataType=False,
                            singleClickEdit=False,
                            hide=_hidden,
                            type=[],
                        )
                    elif _gc == _AVG_U_STD:
                        gb.configure_column(
                            _gc, valueFormatter=_fmt_avg_u,
                            editable=True, hide=_hidden,
                            wrapHeaderText=True, minWidth=120, width=130,
                            headerClass=_header_class(_gc),
                            cellStyle=_style_data,
                            type=["numericColumn"],
                        )
                    elif _gc == "Status":
                        gb.configure_column(
                            _gc, editable=True, hide=_hidden,
                            wrapHeaderText=True, minWidth=120, width=130,
                            headerClass=_header_class(_gc),
                            valueFormatter=_fmt_status,
                            cellStyle=_style_status,
                            cellEditor="agSelectCellEditor",
                            cellEditorParams=_status_action_editor_params,
                            singleClickEdit=False,
                            type=[],
                        )
                    else:
                        # Data columns: readable widths, word-wrap headers, user-resizable
                        gb.configure_column(
                            _gc, editable=True, hide=_hidden,
                            wrapHeaderText=True, minWidth=120, width=130,
                            headerClass=_header_class(_gc),
                            cellStyle=_style_data,
                        )

                _go = gb.build()
                _go["headerHeight"]              = 260 if _dyn_pog_cols else 56
                _go["rowHeight"]                 = 32
                _pog_cols_js = _json.dumps([str(c) for c in _dyn_pog_cols])
                _sticky_cols_js = _json.dumps([str(c) for c in _STICKY])
                _search_scroll_col_js = _json.dumps(
                    str(_search_pog_match_cols[0]) if _search_pog_match_cols else ""
                )
                _new_row_scroll_js = _json.dumps(str(st.session_state.get(_new_rows_scroll_key, "") or ""))
                _go["onCellValueChanged"] = JsCode("""
function(params) {
  var pogCols = __POG_COLS__;
  var stickyCols = __STICKY_COLS__;
  var origMap = __ORIG_MAP__;
  var avgCol = __AVG_U_COL__;
  var field = String((params.colDef && params.colDef.field) || '');
  if (!params.data) return;
  params.data['__rs_last_edit'] = field;

  if (field === 'Status') {
    var sv = String(params.newValue == null ? '' : params.newValue).trim().toUpperCase();
    if (sv === 'MAINTAIN') {
      var rowKey = String(params.data['__rs_row_key'] || '');
      var stickyKey = stickyCols.map(function(c){
        return String(params.data[c] == null ? '' : params.data[c]);
      }).join('\\u241e');
      function originalValue(col) {
        var k1 = rowKey + '\\u241f' + col;
        if (Object.prototype.hasOwnProperty.call(origMap, k1)) return origMap[k1];
        var k2 = stickyKey + '\\u241f' + col;
        if (Object.prototype.hasOwnProperty.call(origMap, k2)) return origMap[k2];
        return null;
      }
      for (var mi = 0; mi < pogCols.length; mi++) {
        var mc = String(pogCols[mi]);
        if (!(mc in params.data)) continue;
        var mv = params.data[mc];
        var ov = originalValue(mc);
        if (ov !== null) {
          params.data[mc] = ov;
        } else if (isAction(mv)) {
          params.data[mc] = '';
        }
      }
      params.data['Status'] = 'MAINTAIN';
      var mwf = 0;
      for (var mwc = 0; mwc < pogCols.length; mwc++) {
        var mpc = String(pogCols[mwc]);
        if (!(mpc in params.data)) continue;
        if (isNum(params.data[mpc])) mwf++;
      }
      params.data['Check Range To-be Waterfall'] = mwf;
      try {
        params.api.refreshCells({rowNodes: [params.node], force: false});
      } catch (e) {}
    } else if (sv === 'DELETE ALL') {
      for (var si = 0; si < pogCols.length; si++) {
        var sc = String(pogCols[si]);
        if (!(sc in params.data)) continue;
        var cv = params.data[sc];
        if (isNum(cv) || cv === 'Delete') {
          params.data[sc] = 'Delete';
        } else if (cv === 'New') {
          params.data[sc] = '';
        }
      }
      params.data['Status'] = 'DELETE ALL';
      params.data['Check Range To-be Waterfall'] = 0;
      try {
        params.api.refreshCells({rowNodes: [params.node], force: false});
      } catch (e) {}
    }
    return;
  }

  if (params.data['__rs_added_row']) {
    var idVal = String(params.data['ID'] == null ? '' : params.data['ID']).trim();
    var cleanId = idVal.replace(/\D/g, '').slice(0, 9);
    function isRequiredAvgNumber(v) {
      if (v === null || v === undefined) return false;
      var s = String(v).replace(/,/g, '').trim();
      if (!s || s === 'nan' || s === 'None') return false;
      return isFinite(Number(s));
    }
    if (cleanId !== idVal) {
      params.data['ID'] = cleanId;
    }
    params.data['Status'] = (
      /^\d{9}$/.test(cleanId) && isRequiredAvgNumber(params.data[avgCol])
    ) ? 'NEWNEW' : '';
    try {
      params.api.refreshCells({
        rowNodes: [params.node],
        columns: ['ID', 'Status', avgCol],
        force: true
      });
    } catch (e) {}
    if (pogCols.indexOf(field) < 0) return;
  } else if (pogCols.indexOf(field) < 0) {
    return;
  }

  function isAction(v) {
    return v === 'Keep' || v === 'Delete' || v === 'New' || v === 'Delist';
  }
  function isNum(v) {
    if (v === null || v === undefined) return false;
    var s = String(v).replace(/,/g, '').trim();
    if (!s || s === 'nan' || s === 'None' || isAction(s)) return false;
    return !isNaN(Number(s));
  }

  var deleteCols = {};
  var newCols = {};
  var existingCols = {};
  var deleteCount = 0;
  var newCount = 0;
  var pogCount = 0;
  var hasDelist = false;

  for (var i = 0; i < pogCols.length; i++) {
    var c = String(pogCols[i]);
    if (!(c in params.data)) continue;
    pogCount++;
    var v = params.data[c];
    if (isNum(v)) existingCols[c] = true;
    if (v === 'Delist') hasDelist = true;
    if (v === 'Delete') {
      deleteCols[c] = true;
      deleteCount++;
    }
    if (v === 'New') {
      newCols[c] = true;
      newCount++;
    }
  }

  var status = 'MAINTAIN';
  if (hasDelist) {
    status = 'Inactive';
  } else if (deleteCount > 0 && newCount > 0) {
    status = 'DELETE SOME';
  } else if (deleteCount > 0) {
    var deleteAll = true;
    var hasDeleteScope = false;
    for (var dc in deleteCols) {
      existingCols[dc] = true;
    }
    for (var ec in existingCols) {
      hasDeleteScope = true;
      if (!deleteCols[ec]) {
        deleteAll = false;
        break;
      }
    }
    status = (hasDeleteScope && deleteAll) ? 'DELETE ALL' : 'DELETE SOME';
  } else if (newCount > 0) {
    status = (pogCount > 0 && newCount >= pogCount) ? 'NEWNEW' : 'NEW SOME';
  }

  params.data['Status'] = status;
  var wf = 0;
  for (var wc = 0; wc < pogCols.length; wc++) {
    var pc = String(pogCols[wc]);
    if (!(pc in params.data)) continue;
    var pv = params.data[pc];
    if (isNum(pv) || pv === 'New') wf++;
  }
  params.data['Check Range To-be Waterfall'] = wf;
  try {
    params.api.refreshCells({
      rowNodes: [params.node],
      columns: ['Status', 'Check Range To-be Waterfall'],
      force: true
    });
  } catch (e) {}
}
""".replace("__POG_COLS__", _pog_cols_js).replace("__STICKY_COLS__", _sticky_cols_js).replace("__ORIG_MAP__", _orig_action_values_js).replace("__AVG_U_COL__", _json.dumps(_AVG_U_STD)))
                # Post horizontal scroll position to parent page so the cluster
                # summary table above can follow. Uses postMessage because the
                # AgGrid iframe is sandboxed (no allow-same-origin).
                _go["onGridReady"] = JsCode("""
function(params){
  var tries=0;
  var _busy=false;
  var _csPogCols=__POG_COLS__;
  var _searchScrollCol=__SEARCH_SCROLL_COL__;
  var _newRowScrollKey=__NEW_ROW_SCROLL_KEY__;
  var _scrollPostPending=false;
  var _lastScrollLeft=0;
  var _lastPostedScrollLeft=-1;
  function csMetrics(){
    var state=params.api.getColumnState();
    var byId={}, pinnedW=0, nonPinnedOffset=0, statusOffset=null, firstPogOffset=null;
    var pogSet={};
    for(var p=0;p<_csPogCols.length;p++){pogSet[String(_csPogCols[p])]=true;}
    for(var i=0;i<state.length;i++){
      var s=state[i];
      if(s.hide)continue;
      var colId=String(s.colId);
      byId[colId]=s;
      var cw=s.width||110;
      if(s.pinned==='left'){
        pinnedW+=cw;
        continue;
      }
      if(statusOffset===null && colId==='Status'){
        statusOffset=nonPinnedOffset;
      }
      if(firstPogOffset===null && pogSet[colId]){
        firstPogOffset=nonPinnedOffset;
      }
      nonPinnedOffset+=cw;
    }
    if(firstPogOffset===null){firstPogOffset=nonPinnedOffset;}
    var spacer=(statusOffset!==null) ? statusOffset : Math.max(0, firstPogOffset-260);
    var lw=Math.max(80, firstPogOffset-spacer);
    var pogW=[];
    for(var j=0;j<_csPogCols.length;j++){
      var ps=byId[String(_csPogCols[j])];
      pogW.push(ps&&!ps.hide ? (ps.width||72) : 72);
    }
    return {_cs_pinned:pinnedW,_cs_spacer:spacer,_cs_lw:lw,_cs_pog_widths:pogW};
  }
  var iv=setInterval(function(){
    tries++;
    var el=document.querySelector('.ag-body-horizontal-scroll-viewport');
    if(!el&&tries<40){return;}
    clearInterval(iv);
    if(!el)return;

    if(_newRowScrollKey){
      setTimeout(function(){
        try{
          var targetIndex=-1;
          params.api.forEachNode(function(node){
            if(targetIndex<0 && node.data && String(node.data['__rs_row_key'])===String(_newRowScrollKey)){
              targetIndex=node.rowIndex;
            }
          });
          if(targetIndex>=0){
            var viewIndex=Math.max(0,targetIndex-3);
            params.api.ensureIndexVisible(viewIndex,'top');
            params.api.setFocusedCell(targetIndex,'ID');
            params.api.startEditingCell({rowIndex:targetIndex,colKey:'ID'});
          }
        }catch(e){}
      },250);
    }

    if(_searchScrollCol){
      setTimeout(function(){
        try{
          if(params.columnApi && params.columnApi.ensureColumnVisible){
            params.columnApi.ensureColumnVisible(_searchScrollCol, 'middle');
          }else if(params.api && params.api.ensureColumnVisible){
            params.api.ensureColumnVisible(_searchScrollCol, 'middle');
          }
          window.parent.postMessage(csMetrics(),'*');
        }catch(e){}
      },150);
    }

    /* ── Compute: pinned width, REST-body spacer, label width ── */
    try{
      window.parent.postMessage(csMetrics(),'*');
    }catch(e){}

    /* ── Bottom → Top: post scroll position ── */
    el.addEventListener('scroll',function(){
      if(_busy)return;
      _lastScrollLeft=el.scrollLeft;
      if(Math.abs(_lastScrollLeft-_lastPostedScrollLeft)<2)return;
      if(_scrollPostPending)return;
      _scrollPostPending=true;
      setTimeout(function(){
        _scrollPostPending=false;
        if(Math.abs(_lastScrollLeft-_lastPostedScrollLeft)<2)return;
        _lastPostedScrollLeft=_lastScrollLeft;
        window.parent.postMessage({_cs_hscroll:_lastScrollLeft},'*');
      },100);
    },{passive:true});
    params.api.addEventListener('columnResized',function(ev){
      if(ev.finished){
        try{
          window.parent.postMessage(csMetrics(),'*');
        }catch(e){}
      }
    });

    /* ── Top → Bottom: receive scroll command ── */
    window.addEventListener('message',function(e){
      if(!e.data)return;
      if(e.data._cs_agscroll===undefined)return;
      _busy=true;
      var target=Number(e.data._cs_agscroll)||0;
      if(Math.abs(el.scrollLeft-target)>1){
        el.scrollLeft=target;
      }
      setTimeout(function(){_busy=false;},80);
    });
  },250);
}
""".replace("__POG_COLS__", _pog_cols_js).replace("__SEARCH_SCROLL_COL__", _search_scroll_col_js).replace("__NEW_ROW_SCROLL_KEY__", _new_row_scroll_js))
                _go["suppressRowClickSelection"] = True
                _go["suppressCellFocus"]         = False
                _go["suppressLoadingOverlay"]    = True
                _go["suppressNoRowsOverlay"]     = True
                _go["overlayLoadingTemplate"]    = "<span></span>"
                _go["overlayNoRowsTemplate"]     = "<span></span>"
                # Columns tool panel — lets users re-show hidden columns via a sidebar
                _go["sideBar"] = {
                    "toolPanels": [
                        {
                            "id": "columns",
                            "labelDefault": "Columns",
                            "toolPanel": "agColumnsToolPanel",
                            "toolPanelParams": {
                                "suppressRowGroups": True,
                                "suppressValues":    True,
                                "suppressPivots":    True,
                                "suppressPivotMode": True,
                            },
                        }
                    ],
                    "defaultToolPanel": "",
                }

                _col_state_key = f"{p}_col_state"
                _active_col_state = st.session_state.get(_col_state_key)
                if _active_col_state:
                    _active_col_state = [dict(_cs) for _cs in _active_col_state]
                    for _cs in _active_col_state:
                        _cid = str(_cs.get("colId") or _cs.get("field") or "")
                        if _cid == "ID":
                            _cs["width"] = max(int(_cs.get("width") or 0), 120)
                        elif _cid == "Status":
                            _cs["width"] = max(int(_cs.get("width") or 0), 130)
                        elif (
                            _cid
                            and _cid not in (_ROW_KEY_COL, _NEW_ROW_COL, _EDIT_COL)
                            and _cid not in _dyn_pog_set
                        ):
                            _cs["width"] = max(int(_cs.get("width") or 0), 120)
                st.session_state[f"{p}_latest_grid_df"] = _tdf_display
                st.session_state[f"{p}_latest_grid_sig"] = f"{_sel_dg_code}|{_sel_dg_name}"
                _grid_return_js = JsCode("""
function({streamlitRerunEventTriggerName, eventData}) {
  var api = eventData && eventData.api;
  var row = eventData && eventData.data ? eventData.data : null;
  return {
    data: row ? [row] : [],
    columnsState: api && api.getColumnState ? api.getColumnState() : null,
    trigger: streamlitRerunEventTriggerName || ''
  };
}
""")
                _grid_response = AgGrid(
                    _tdf_display,
                    gridOptions=_go,
                    update_mode=GridUpdateMode.VALUE_CHANGED,
                    data_return_mode=DataReturnMode.CUSTOM,
                    custom_jscode_for_grid_return=_grid_return_js,
                    custom_css=_custom_css,
                    theme="balham",
                    height=900 if _dyn_pog_cols else 720,
                    fit_columns_on_grid_load=False,
                    allow_unsafe_jscode=True,
                    enable_enterprise_modules=True,
                    try_to_convert_back_to_original_types=False,
                    columns_state=_active_col_state,
                    update_on=["cellValueChanged"],
                    key=f"{p}_aggrid",
                )
                # Persist column state so hide/width/pin survive reruns and DG changes
                _saved_col_state = getattr(_grid_response, "columns_state", None)
                if _saved_col_state is None and hasattr(_grid_response, "get"):
                    _saved_col_state = _grid_response.get("columnsState")
                if _saved_col_state is not None:
                    st.session_state[_col_state_key] = _saved_col_state
                if st.session_state.get(_new_rows_scroll_key):
                    st.session_state[_new_rows_scroll_key] = ""

                # ── Receive AG Grid scroll messages → update top cluster table ───
                # AgGrid iframe is sandboxed, so we use postMessage.
                # onGridReady (above) posts {_cs_hscroll: x} to window.parent.
                # This iframe listens on window.parent for those messages and
                # updates the top table's scrollLeft accordingly.
                import streamlit.components.v1 as _stcv1
                _stcv1.html("""<script>
(function(){
  var _raf=window.requestAnimationFrame||function(f){setTimeout(f,16);};
  var _doc=window.parent.document;
  var _syncBusy=false;
  var _csPinned=0, _csSpacer=0, _csHScroll=0;
  var _relayFrames=null, _relayFrameTs=0, _lastRelayX=-1, _relayPending=false;

  function applyClusterPosition(){
    var cover=_doc.querySelector('.cs-pin-cover');
    var sync=_doc.querySelector('.cs-scroll-sync');
    if(!sync)return;
    var left=Math.max(_csPinned, _csPinned + _csSpacer - _csHScroll);
    sync.style.marginLeft=left+'px';
    if(cover){cover.style.width=left+'px';}
    var target=Math.max(0, _csHScroll - _csSpacer);
    if(Math.abs(sync.scrollLeft-target)>1){
      _syncBusy=true;
      sync.scrollLeft=target;
      setTimeout(function(){_syncBusy=false;},80);
    }
  }

  /* ── Receive messages from AgGrid (bottom → top scroll + spacer) ── */
  window.parent.addEventListener('message',function(e){
    if(!e.data)return;

    /* Update pinned cover width + scrollable zone margin-left */
    if(e.data._cs_pinned!==undefined){
      _csPinned=e.data._cs_pinned;
      applyClusterPosition();
    }

    /* Update REST cols width so body Cluster aligns with body Status */
    if(e.data._cs_spacer!==undefined){
      _csSpacer=e.data._cs_spacer;
      applyClusterPosition();
    }

    /* Update Cluster label col width = Status + Check Range */
    if(e.data._cs_lw!==undefined){
      var lbls=_doc.querySelectorAll('.cs-label-col');
      for(var i=0;i<lbls.length;i++){
        lbls[i].style.minWidth=e.data._cs_lw+'px';
        lbls[i].style.width=e.data._cs_lw+'px';
        lbls[i].style.maxWidth=e.data._cs_lw+'px';
      }
    }

    /* Update each top POG column to match the lower grid column width */
    if(e.data._cs_pog_widths!==undefined){
      var widths=e.data._cs_pog_widths, cells=_doc.querySelectorAll('.cs-pog-cell'), byIdx={};
      for(var c=0;c<cells.length;c++){
        var idx=parseInt(cells[c].getAttribute('data-idx'),10);
        if(!byIdx[idx])byIdx[idx]=[];
        byIdx[idx].push(cells[c]);
      }
      for(var j=0;j<widths.length;j++){
        var w=widths[j]+'px';
        if(byIdx[j])for(var k=0;k<byIdx[j].length;k++){
          byIdx[j][k].style.minWidth=w;
          byIdx[j][k].style.width=w;
          byIdx[j][k].style.maxWidth=w;
        }
      }
    }

    /* Sync top table scrollLeft from AgGrid scroll event */
    if(e.data._cs_hscroll!==undefined){
      var top=_doc.querySelector('.cs-scroll-sync');
      if(!top)return;
      _csHScroll=e.data._cs_hscroll;
      _syncBusy=true;
      _raf(function(){
        applyClusterPosition();
        setTimeout(function(){_syncBusy=false;},50);
      });
    }
  });

  /* ── Attach top → bottom scroll relay ── */
  function attachTopRelay(){
    var top=_doc.querySelector('.cs-scroll-sync');
    if(!top){setTimeout(attachTopRelay,400);return;}
    top.addEventListener('scroll',function(){
      if(_syncBusy)return;
      var x=_csSpacer + top.scrollLeft;
      if(Math.abs(x-_lastRelayX)<2)return;
      if(_relayPending)return;
      _relayPending=true;
      setTimeout(function(){
        _relayPending=false;
        if(Math.abs(x-_lastRelayX)<2)return;
        _lastRelayX=x;
        var now=Date.now();
        if(!_relayFrames || now-_relayFrameTs>2000){
          _relayFrames=_doc.querySelectorAll('iframe');
          _relayFrameTs=now;
        }
        for(var i=0;i<_relayFrames.length;i++){
          try{_relayFrames[i].contentWindow.postMessage({_cs_agscroll:x},'*');}catch(ex){}
        }
      },80);
    },{passive:true});
  }

  /* ── Close visual gap between cluster summary and AgGrid ── */
  function closeGap(){
    var top=_doc.querySelector('.cs-scroll-sync');
    if(!top)return;
    function ancestor(el,attr,val){
      while(el){if(el.getAttribute&&el.getAttribute(attr)===val)return el;el=el.parentElement;}
      return null;
    }
    var csEl=ancestor(top,'data-testid','element-container');
    if(csEl)csEl.style.cssText+=';margin-bottom:0!important;padding-bottom:0!important;';
    var iframes=_doc.querySelectorAll('iframe');
    for(var i=0;i<iframes.length;i++){
      var el=ancestor(iframes[i],'data-testid','element-container');
      if(el&&el!==csEl){el.style.cssText+=';margin-top:0!important;padding-top:0!important;';break;}
    }
  }

  setTimeout(function(){
    var old=_doc.getElementById('rs-floating-hscroll');
    if(old&&old.parentNode){old.parentNode.removeChild(old);}
    var oldStyle=_doc.getElementById('rs-floating-hscroll-style');
    if(oldStyle&&oldStyle.parentNode){oldStyle.parentNode.removeChild(oldStyle);}
    closeGap();
    attachTopRelay();
  },800);
})();
</script>""", height=0)

                # ── Auto-commit on VALUE_CHANGED ──────────────────────────────────
                _grid_data = (
                    _grid_response.get("data")
                    if hasattr(_grid_response, "get")
                    else _grid_response["data"]
                )
                if _grid_data is not None:
                    try:
                        _new_df = (pd.DataFrame(_grid_data)
                                   .reindex(columns=_tdf_display.columns))
                    except Exception:
                        _new_df = None
                    if _new_df is not None and len(_new_df) > 0:
                        if len(_new_df) >= len(_tdf_display):
                            st.session_state[f"{p}_latest_grid_df"] = _new_df
                            st.session_state[f"{p}_latest_grid_sig"] = f"{_sel_dg_code}|{_sel_dg_name}"
                        _pog_act_store = _pog_actions       # same live dict (setdefault above)
                        _pog_edit_store = _pog_edits
                        _avg_u_store   = _avg_u_edits
                        _data_edit_store = _data_edits
                        _stat_store    = _status_overrides
                        _audit_lines   = []
                        _needs_rerun   = False
                        _needs_confirm_rerun = False
                        _changed_rks   = set()
                        _new_rows_changed = False
                        _new_rows_for_commit = (
                            list(st.session_state.get(_new_rows_key, []))
                            + list(st.session_state.get(_new_rows_placeholder_key, []))
                        )
                        _new_rows_by_key = {
                            str(_nr.get(_ROW_KEY_COL, "")): dict(_nr)
                            for _nr in _new_rows_for_commit
                        }
                        _existing_action_rks = (
                            set(_pog_act_store.keys())
                            | set(_pog_edit_store.keys())
                            | set(_stat_store.keys())
                        )
                        _id_sticky_idx = _STICKY.index("ID") if "ID" in _STICKY else None
                        _existing_rk_by_id = {}
                        if _id_sticky_idx is not None:
                            for _erk in _existing_action_rks:
                                if len(_erk) > _id_sticky_idx:
                                    _existing_rk_by_id[str(_erk[_id_sticky_idx])] = _erk

                        def _resolve_grid_rk(_row):
                            _rk_exact = _row_to_rk(_row)
                            if _rk_exact in _rk_to_i or _rk_exact in _existing_action_rks:
                                return _rk_exact
                            if _id_sticky_idx is not None and "ID" in _row:
                                _rk_by_id = _existing_rk_by_id.get(str(_row.get("ID", "")))
                                if _rk_by_id is not None:
                                    return _rk_by_id
                            return _rk_exact

                        _dyn_cols_in_new = [c for c in _dyn_pog_cols if c in _new_df.columns]
                        if _EDIT_COL in _new_df.columns:
                            _edit_mask = _new_df[_EDIT_COL].fillna("").astype(str).ne("")
                        else:
                            _edit_mask = pd.Series(False, index=_new_df.index)
                        if "Status" in _new_df.columns:
                            _status_s = _new_df["Status"].fillna("").astype(str).str.upper()
                            if bool(_edit_mask.any()):
                                # Fast path: the grid's valueSetter/onCellValueChanged marks
                                # exactly the edited row.  Scanning every existing action row
                                # here makes each dropdown feel like a full-table recalculation.
                                _candidate_df = _new_df.loc[_edit_mask]
                            else:
                                _grid_rks = _new_df.apply(_resolve_grid_rk, axis=1)
                                _existing_mask = _grid_rks.apply(lambda _rk: _rk in _existing_action_rks)
                                _store_status_s = _grid_rks.apply(
                                    lambda _rk: str(_stat_store.get(_rk, "MAINTAIN")).strip().upper()
                                )
                                _status_mismatch_mask = _status_s.ne(_store_status_s)
                                _maintain_mask = _status_s.eq("MAINTAIN") & _existing_mask
                                _action_status_mask = (
                                    _status_s.isin(["DELETE SOME", "DELETE ALL", "NEW SOME", "NEWNEW"])
                                    & _status_mismatch_mask
                                )
                                _candidate_df = _new_df.loc[
                                    _maintain_mask | _action_status_mask
                                ]
                        else:
                            _candidate_df = _new_df.loc[_edit_mask]

                        for _nri, _nrow in _candidate_df.iterrows():
                            _edit_field = str(_nrow.get(_EDIT_COL, "") or "")
                            _status_edit = str(_nrow.get("Status", "") or "").strip().upper()
                            _crk = _resolve_grid_rk(_nrow)
                            _cri = _rk_to_i.get(_crk)
                            if _cri is None:
                                if bool(_nrow.get(_NEW_ROW_COL, False)):
                                    _rk_raw = str(_nrow.get(_ROW_KEY_COL, ""))
                                    if _rk_raw:
                                        _nr = {
                                            c: ("" if _nrow.get(c, "") is None else _nrow.get(c, ""))
                                            for c in _tdf_display.columns
                                            if c not in (_NEW_ROW_COL, _EDIT_COL)
                                        }
                                        _nr[_ROW_KEY_COL] = _rk_raw
                                        _nr[_NEW_ROW_COL] = True
                                        _id_val = _re.sub(r"\D+", "", str(_nr.get("ID", "")).strip())[:9]
                                        if "ID" in _nr:
                                            _nr["ID"] = _id_val
                                        if "Status" in _nr:
                                            _nr["Status"] = (
                                                "NEWNEW"
                                                if _re.fullmatch(r"\d{9}", _id_val)
                                                and _new_row_avg_is_numeric(_nr)
                                                else ""
                                            )
                                        if _new_row_has_input(_nr):
                                            _prev = _new_rows_by_key.get(_rk_raw)
                                            if _prev != _nr:
                                                _new_rows_by_key[_rk_raw] = _nr
                                                _new_rows_changed = True
                                        elif _rk_raw in _new_rows_by_key:
                                            _new_rows_by_key.pop(_rk_raw, None)
                                            _new_rows_changed = True
                                continue
                            if not _edit_field and not (
                                (
                                    _status_edit == "MAINTAIN"
                                    and _crk in _existing_action_rks
                                )
                                or (
                                    _status_edit == "DELETE ALL"
                                    and str(_stat_store.get(_crk, "MAINTAIN")).strip().upper() != "DELETE ALL"
                                )
                                or (
                                    _status_edit in ("DELETE SOME", "NEW SOME", "NEWNEW")
                                    and str(_stat_store.get(_crk, "MAINTAIN")).strip().upper() != _status_edit
                                )
                            ):
                                continue

                            # Pog action changes
                            _action_cols = set()
                            _status_row_edit = _edit_field == "Status" or (
                                _status_edit == "MAINTAIN"
                                and _crk in _existing_action_rks
                            ) or (
                                _status_edit == "DELETE ALL"
                                and str(_stat_store.get(_crk, "MAINTAIN")).strip().upper() != "DELETE ALL"
                            )
                            if _status_row_edit and _status_edit == "MAINTAIN":
                                _action_cols.update(_pog_act_store.get(_crk, {}).keys())
                                _action_cols.update(_pog_edit_store.get(_crk, {}).keys())
                                for _scan_pc in _dyn_cols_in_new:
                                    if str(_nrow.get(_scan_pc, "")) in _ACT_SET:
                                        _action_cols.add(_scan_pc)
                            elif _status_row_edit and _status_edit == "DELETE ALL":
                                pass
                            elif _edit_field in _dyn_pog_set:
                                _action_cols.update(_pog_act_store.get(_crk, {}).keys())
                                _action_cols.update(_pog_edit_store.get(_crk, {}).keys())
                                _action_cols.add(_edit_field)
                            elif _status_edit in ("DELETE SOME", "NEW SOME", "NEWNEW"):
                                _action_cols.update(_pog_act_store.get(_crk, {}).keys())
                                _action_cols.update(_pog_edit_store.get(_crk, {}).keys())
                                for _scan_pc in _dyn_cols_in_new:
                                    if str(_nrow.get(_scan_pc, "")) in _ACT_SET:
                                        _action_cols.add(_scan_pc)
                            if _status_row_edit and _status_edit == "MAINTAIN":
                                if (
                                    _crk in _pog_act_store
                                    or _crk in _pog_edit_store
                                    or _crk in _stat_store
                                    or _action_cols
                                ):
                                    _pog_act_store.pop(_crk, None)
                                    _pog_edit_store.pop(_crk, None)
                                    _stat_store.pop(_crk, None)
                                    _needs_rerun = True
                                    _changed_rks.add(_crk)
                                    _audit_lines.append(f"{_crk}: status -> MAINTAIN")
                                continue
                            if _status_row_edit and _status_edit == "DELETE ALL":
                                if (
                                    _pog_act_store.get(_crk)
                                    or _pog_edit_store.get(_crk)
                                    or _stat_store.get(_crk) != "DELETE ALL"
                                ):
                                    _pog_act_store.pop(_crk, None)
                                    _pog_edit_store.pop(_crk, None)
                                    _stat_store[_crk] = "DELETE ALL"
                                    _needs_rerun = True
                                    _changed_rks.add(_crk)
                                    _audit_lines.append(f"{_crk}: status -> DELETE ALL")
                                continue
                            for _cpc in _action_cols:
                                if _cpc not in _tdf_display.columns:
                                    continue
                                _had_action = (
                                    _cpc in _pog_act_store.get(_crk, {})
                                    or _cpc in _pog_edit_store.get(_crk, {})
                                )
                                _ov = _tdf.at[_cri, _cpc]
                                _nv = _nrow.get(_cpc, "")
                                _ov_s = ("" if _ov is None or
                                         (isinstance(_ov, float) and pd.isna(_ov))
                                         else str(_ov))
                                _nv_s = ("" if _nv is None or
                                         (isinstance(_nv, float) and pd.isna(_nv))
                                         else str(_nv))
                                if _nv_s in ("nan", "None"):
                                    _nv_s = ""
                                if _status_row_edit and _status_edit == "DELETE ALL":
                                    _nv_s = "Delete"
                                elif _status_row_edit and _status_edit == "MAINTAIN":
                                    _nv_s = ""
                                if (
                                    _status_row_edit
                                    and _status_edit == "DELETE ALL"
                                    and _had_action
                                    and _pog_act_store.get(_crk, {}).get(_cpc) == "Delete"
                                    and _pog_edit_store.get(_crk, {}).get(_cpc) == "Delete"
                                ):
                                    continue
                                if _ov_s == _nv_s and not _had_action:
                                    continue
                                _needs_rerun = True
                                _changed_rks.add(_crk)
                                if _nv_s == "" or _nv_s not in _ACT_SET:
                                    # Blank selected → clear action, revert to showing number
                                    if _crk in _pog_act_store:
                                        _pog_act_store[_crk].pop(_cpc, None)
                                        if not _pog_act_store[_crk]:
                                            del _pog_act_store[_crk]
                                    if _crk in _pog_edit_store:
                                        _pog_edit_store[_crk].pop(_cpc, None)
                                        if not _pog_edit_store[_crk]:
                                            del _pog_edit_store[_crk]
                                    _audit_lines.append(
                                        f"{_crk}|{_cpc}: {_ov_s!r}→cleared")
                                else:
                                    if _nv_s == "Delete" and not (
                                        _status_row_edit
                                        and _status_edit == "DELETE ALL"
                                    ):
                                        _is_top, _cell_val, _cutoff = _delete_is_top10(_cri, _cpc)
                                        if _is_top:
                                            _item = (
                                                str(_tdf.at[_cri, "Item Name"])
                                                if "Item Name" in _tdf.columns
                                                and pd.notna(_tdf.at[_cri, "Item Name"])
                                                else str(_crk)
                                            )
                                            st.session_state[_pending_del_key] = {
                                                "rk": _crk,
                                                "row_i": _cri,
                                                "col": _cpc,
                                                "pog": _cpc,
                                                "item": _item,
                                                "value": _cell_val,
                                                "cutoff": _cutoff,
                                            }
                                            _needs_confirm_rerun = True
                                            _audit_lines.append(
                                                f"{_crk}|{_cpc}: Delete blocked for top10 confirm")
                                            continue
                                    if _nv_s == "New" and _is_low_sales_item(_cri):
                                        _item = (
                                            str(_tdf.at[_cri, "Item Name"])
                                            if "Item Name" in _tdf.columns
                                            and pd.notna(_tdf.at[_cri, "Item Name"])
                                            else str(_crk)
                                        )
                                        st.session_state[_pending_new_key] = {
                                            "rk": _crk,
                                            "row_i": _cri,
                                            "col": _cpc,
                                            "pog": _cpc,
                                            "item": _item,
                                        }
                                        _needs_confirm_rerun = True
                                        _audit_lines.append(
                                            f"{_crk}|{_cpc}: New blocked for low-sales confirm")
                                        continue
                                    _pog_act_store.setdefault(_crk, {})[_cpc] = _nv_s
                                    _pog_edit_store.setdefault(_crk, {})[_cpc] = _nv_s
                                    _audit_lines.append(
                                        f"{_crk}|{_cpc}: {_ov_s!r}→{_nv_s!r}")

                            # Avg Units numeric change
                            if _edit_field == _AVG_U_STD and _AVG_U_STD in _tdf_display.columns:
                                _ou = _tdf.at[_cri, _AVG_U_STD]
                                _nu = _nrow.get(_AVG_U_STD, None)
                                try:
                                    _of = (float(_ou) if _ou is not None and not (
                                               isinstance(_ou, float) and pd.isna(_ou))
                                           else None)
                                    _nf = (float(_nu) if _nu is not None and not (
                                               isinstance(_nu, float) and pd.isna(_nu))
                                           else None)
                                except (TypeError, ValueError):
                                    _of = _nf = None
                                if _of != _nf:
                                    _needs_rerun = True
                                    if _nf is None:
                                        _avg_u_store.setdefault(_crk, {}).pop(
                                            _AVG_U_STD, None)
                                        if not _avg_u_store.get(_crk):
                                            _avg_u_store.pop(_crk, None)
                                    else:
                                        _avg_u_store.setdefault(_crk, {})[_AVG_U_STD] = _nf
                                    _audit_lines.append(
                                        f"{_crk}|Avg Units: {_of}→{_nf}")

                            # Plain main-table edits: allow users to delete/type values
                            # in regular columns and keep those edits across refreshes.
                            _plain_blocked = (
                                _ROW_KEY_COL,
                                _NEW_ROW_COL,
                                _EDIT_COL,
                                "Status",
                                "Check Range To-be Waterfall",
                            )
                            if (
                                _edit_field
                                and _edit_field in _tdf_display.columns
                                and _edit_field not in _plain_blocked
                                and _edit_field not in _dyn_pog_set
                                and _edit_field != _AVG_U_STD
                            ):
                                _old_plain = _tdf.at[_cri, _edit_field] if _edit_field in _tdf.columns else ""
                                _new_plain = _nrow.get(_edit_field, "")

                                def _plain_norm(_v):
                                    if _v is None:
                                        return ""
                                    try:
                                        if pd.isna(_v):
                                            return ""
                                    except Exception:
                                        pass
                                    return str(_v)

                                _old_s = _plain_norm(_old_plain)
                                _new_s = _plain_norm(_new_plain)
                                if _old_s != _new_s:
                                    _needs_rerun = True
                                    _data_edit_store.setdefault(_crk, {})[_edit_field] = _new_plain
                                    _audit_lines.append(
                                        f"{_crk}|{_edit_field}: {_old_s!r}→{_new_s!r}"
                                    )
                                elif _edit_field in _data_edit_store.get(_crk, {}):
                                    _needs_rerun = True
                                    _data_edit_store[_crk].pop(_edit_field, None)
                                    if not _data_edit_store.get(_crk):
                                        _data_edit_store.pop(_crk, None)
                                    _audit_lines.append(
                                        f"{_crk}|{_edit_field}: restored original"
                                    )

                        if _needs_rerun:
                            if _new_rows_changed:
                                _session_new_rows = list(_new_rows_by_key.values())
                                _saved_new_rows = [
                                    _r for _r in _session_new_rows
                                    if _new_row_has_input(_r)
                                ]
                                _placeholder_new_rows = [
                                    _r for _r in _session_new_rows
                                    if not _new_row_has_input(_r)
                                ]
                                st.session_state[_new_rows_key] = _saved_new_rows
                                st.session_state[_new_rows_placeholder_key] = _placeholder_new_rows
                                _rs_save_new_rows(_new_rows_scope, _saved_new_rows)
                            # Re-derive Status only for rows touched by this edit.
                            if "Status" in _tdf.columns:
                                for _srk in _changed_rks:
                                    _derived = _derive_status(_srk)
                                    if _derived == "MAINTAIN":
                                        _stat_store.pop(_srk, None)
                                    else:
                                        _stat_store[_srk] = _derived
                            if _audit_lines:
                                add_audit(
                                    "Planogram Action",
                                    f"tab={p}; {len(_audit_lines)} change(s): "
                                    + "; ".join(_audit_lines[:10])
                                    + ("…" if len(_audit_lines) > 10 else ""),
                                )
                            _rs_save_edit_state(
                                _edit_state_scope,
                                _pog_act_store,
                                _pog_edit_store,
                                _avg_u_store,
                                _stat_store,
                                _data_edit_store,
                            )
                            if _needs_confirm_rerun:
                                st.rerun()
                        elif _new_rows_changed:
                            _session_new_rows = list(_new_rows_by_key.values())
                            _saved_new_rows = [
                                _r for _r in _session_new_rows
                                if _new_row_has_input(_r)
                            ]
                            _placeholder_new_rows = [
                                _r for _r in _session_new_rows
                                if not _new_row_has_input(_r)
                            ]
                            st.session_state[_new_rows_key] = _saved_new_rows
                            st.session_state[_new_rows_placeholder_key] = _placeholder_new_rows
                            _rs_save_new_rows(_new_rows_scope, _saved_new_rows)

            # Row-count helper hidden to keep the review page clean.

        # ── Cluster ───────────────────────────────────────────────────────────
        elif _subview == "🏪 Cluster":
            _DM=[("TO-BE Stores applied count","#F4A460","#000000"),
                 ("AS-IS Stores applied count","#F4A460","#000000"),
                 ("MODs","#DCDCDC","#000000"),
                 ("FIXTURE","#DCDCDC","#000000"),
                 ("RANGE CLASS","#DCDCDC","#000000"),
                 ("Total NEW SKUs","#808080","#FFFFFF"),
                 ("Total DELETE SKUs","#808080","#FFFFFF"),
                 ("TO-BE SKUs count","#808080","#FFFFFF"),
                 ("AS-IS SKUs count","#F0F0F0","#1565C0"),
                 ("%Achieving LRD CASE (As Is)","#FFFFFF","#000000"),
                 ("%Achieving LRD SALES (As Is)","#FFFFFF","#000000")]
            _DRC={r:(bg,tc) for r,bg,tc in _DM}
            _kc=f"{p}_ct_cls"; _kd=f"{p}_ct_dat"; _kr=f"{p}_ct_clr"
            # Pull unique POG Cluster Mod Fixture values from data to use as cluster columns
            _pog_name_col = next(
                (c for c in df_view.columns if _nca(c) == _nca("POG Cluster Mod Fixture")),
                None
            ) or next(
                (c for c in df_view.columns if "pogcluster" in _nca(c) or ("cluster" in _nca(c) and "mod" in _nca(c))),
                None
            )
            _auto_clusters = (
                sorted(df_view[_pog_name_col].dropna().astype(str).str.strip()
                       .replace("", pd.NA).dropna().unique().tolist())
                if _pog_name_col else []
            )
            # Reset cluster columns/data when DG filter changes
            _dg_filter_sig = f"{_sel_dg_code}|{_sel_dg_name}"
            _dg_sig_key = f"{p}_ct_dg_sig"
            _filter_changed = st.session_state.get(_dg_sig_key) != _dg_filter_sig
            if _filter_changed:
                st.session_state[_dg_sig_key] = _dg_filter_sig
                st.session_state[_kc] = _auto_clusters[:]
                st.session_state[_kd] = {r: {} for r, _, _ in _DM}
                st.session_state[_kr] = dict(_DRC)
            elif _kc not in st.session_state:
                st.session_state[_kc] = _auto_clusters[:]
            if _kd not in st.session_state: st.session_state[_kd]={r:{} for r,_,_ in _DM}
            if _kr not in st.session_state: st.session_state[_kr]=dict(_DRC)
            _tb1,_tb2,_tb3,_tb4,_tb5=st.columns([3.2,1.1,1.0,0.9,0.8])
            with _tb1: st.text_input("_ct",key=f"{p}_ct_inp",placeholder="Type name → Add as Column or Row…",label_visibility="collapsed")
            with _tb2:
                if st.button("＋ Column",key=f"{p}_ct_add_col",use_container_width=True):
                    _v=st.session_state.get(f"{p}_ct_inp","").strip()
                    if _v and _v not in st.session_state[_kc]:
                        st.session_state[_kc].append(_v)
                        for _rk in st.session_state[_kd]: st.session_state[_kd][_rk].setdefault(_v,"")
                    st.session_state[f"{p}_ct_inp"]=""; st.rerun()
            with _tb3:
                if st.button("＋ Row",key=f"{p}_ct_add_row",use_container_width=True):
                    _v=st.session_state.get(f"{p}_ct_inp","").strip() or f"Metric {len(st.session_state[_kd])+1}"
                    if _v not in st.session_state[_kd]:
                        st.session_state[_kd][_v]={c:"" for c in st.session_state[_kc]}
                        st.session_state[_kr][_v]=("#FFFFFF","#000000")
                    st.session_state[f"{p}_ct_inp"]=""; st.rerun()
            with _tb4:
                if st.button("↺ Reset",key=f"{p}_ct_reset",use_container_width=True):
                    st.session_state[_kc]=_auto_clusters[:]
                    st.session_state[_kd]={r:{} for r,_,_ in _DM}
                    st.session_state[_kr]=dict(_DRC); st.rerun()
            with _tb5: _em=st.toggle("✏️ Edit",key=f"{p}_ct_edit")
            _clsx=st.session_state[_kc]; _rowsx=list(st.session_state[_kd].keys())
            if not _em:
                if _clsx:
                    _vdf = pd.DataFrame(
                        [{"POG Cluster": _rn,
                          **{_c: (lambda v: None if v is None or str(v) in ("nan","None","") else v)(
                              st.session_state[_kd].get(_rn, {}).get(_c, ""))
                             for _c in _clsx}}
                         for _rn in _rowsx]
                    )
                    st.dataframe(_vdf, use_container_width=True, hide_index=True, height=420)
                else:
                    st.info("No cluster columns found in data. Add columns using the input above ↑")
            else:
                _er=[{"POG Cluster":r,**{c:st.session_state[_kd].get(r,{}).get(c,"") for c in _clsx}} for r in _rowsx]
                _edf=pd.DataFrame(_er) if _er else pd.DataFrame(columns=["POG Cluster"]+_clsx)
                _ecfg={"POG Cluster":st.column_config.TextColumn("POG Cluster",width="large"),
                       **{c:st.column_config.TextColumn(c,width="small") for c in _clsx}}
                _edited=st.data_editor(_edf,num_rows="dynamic",use_container_width=True,
                                       hide_index=True,column_config=_ecfg,key=f"{p}_ct_editor")
                _nd={}
                for _,_row in _edited.iterrows():
                    _rn=str(_row.get("POG Cluster","")).strip()
                    if _rn and _rn not in ("nan","None",""):
                        _nd[_rn]={c:_row.get(c,"") for c in _clsx}
                        if _rn not in st.session_state[_kr]: st.session_state[_kr][_rn]=("#FFFFFF","#000000")
                st.session_state[_kd]=_nd
            if _rowsx:
                _expdf=pd.DataFrame([{"POG Cluster":r,**{c:st.session_state[_kd].get(r,{}).get(c,"") for c in _clsx}} for r in _rowsx])
                st.markdown("<div style='height:10px;'></div>",unsafe_allow_html=True)
                st.download_button("⬇️ Export Cluster Table .xlsx",df_to_xlsx_bytes(_expdf),
                    file_name=f"cluster_{p}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"{p}_cl_dl")

    # ── NEW CANVAS ────────────────────────────────────────────────────────────
    if False:
        st.markdown("""<div style="font-size:16px;font-weight:700;color:#1A1A1A;margin-bottom:4px;">New Canvas</div>
        <div style="font-size:13px;color:#888;margin-bottom:16px;">Select columns, filter rows, save as a named canvas.</div>""",
            unsafe_allow_html=True)
        _cv1,_cv2=st.columns(2)
        with _cv1: _cnv_srch=st.text_input("Filter",placeholder="🔎  Filter rows...",label_visibility="collapsed",key=f"{p}_cnv_srch")
        with _cv2: _cnv_name=st.text_input("Canvas name",placeholder="e.g. SSPOG-A Items",label_visibility="collapsed",key=f"{p}_cnv_name")
        _cnv_cols=[]
        for _g in _sag:
            st.caption(f"— {_g['group']} —")
            _cnv_cols.extend(st.multiselect(_g["group"],_g["matched"],
                default=_g["matched"][:min(3,len(_g["matched"]))],
                key=f"{p}_cnv_{_g['group']}",label_visibility="collapsed"))
        _cdf=df_src[_cnv_cols].copy() if _cnv_cols else df_src.copy()
        if _cnv_srch:
            _cdf=_cdf[_cdf.apply(lambda r:r.astype(str).str.contains(_cnv_srch,case=False,na=False).any(),axis=1)]
        st.dataframe(_dedup(_cdf.reset_index(drop=True)),use_container_width=True,height=300,hide_index=True)
        st.caption(f"{len(_cdf):,} rows · {len(_cnv_cols)} columns")
        _ccb1,_ccb2=st.columns(2)
        with _ccb1:
            if st.button("💾 Save Canvas",use_container_width=True,key=f"{p}_save_canvas"):
                _nm=_cnv_name.strip() or f"Canvas {len(st.session_state.canvases)+1}"
                st.session_state.canvases[_nm]=_cdf.copy(); add_audit("Save Canvas",_nm)
                st.success(f'✅ Canvas "{_nm}" saved!')
        with _ccb2:
            if _cnv_cols:
                st.download_button("⬇️ Export Canvas",df_to_xlsx_bytes(_cdf),
                    file_name=f"{_cnv_name or 'canvas'}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,key=f"{p}_dl_canvas")
        if st.session_state.canvases:
            st.markdown("**Saved Canvases**")
            for _cnm,_cs in st.session_state.canvases.items():
                _sc1,_sc2=st.columns([4,1])
                with _sc1: st.write(f"🎨 **{_cnm}** — {len(_cs):,} rows · {len(_cs.columns)} cols")
                with _sc2: st.download_button("⬇️",df_to_xlsx_bytes(_cs),file_name=f"{_cnm}.xlsx",key=f"{p}_cs_dl_{_cnm}")

    st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)
    _submit_missing_key = f"{p}_submit_missing_blanks"
    _blank_hl_key = f"{p}_highlight_submit_blanks"

    def _strip_submit_internal_cols(_df: pd.DataFrame) -> pd.DataFrame:
        _drop_cols = [
            c for c in ("__rs_row_key", "__rs_added_row", "__rs_last_edit")
            if c in _df.columns
        ]
        return _df.drop(columns=_drop_cols) if _drop_cols else _df

    def _submit_payload_df(limit_rows: int | None = None, keep_internal: bool = False):
        _base = st.session_state.get(f"{p}_latest_grid_df")
        if st.session_state.get(f"{p}_latest_grid_sig") != f"{_sel_dg_code}|{_sel_dg_name}":
            _base = None
        if not isinstance(_base, pd.DataFrame):
            _base = locals().get("_tdf_display", None)
        if isinstance(_base, pd.DataFrame):
            _src = _base
        else:
            _src = df_view
        if limit_rows is not None:
            _src = _src.head(limit_rows)
        _out = _src.reset_index(drop=True)
        if not keep_internal:
            _out = _strip_submit_internal_cols(_out)
        return _out

    def _submit_to_report_packet(payload_override: pd.DataFrame | None = None):
        if isinstance(payload_override, pd.DataFrame):
            _payload = _strip_submit_internal_cols(payload_override.reset_index(drop=True))
        else:
            _payload = _submit_payload_df()
        _pog_cluster_map = {}
        try:
            for _pog in locals().get("_dyn_pog_cols", []) or []:
                _cl = locals().get("_pog_to_cl", {}).get(_pog) or locals().get("_pog_to_cl", {}).get(_nca(_pog), "")
                if str(_cl).strip():
                    _pog_cluster_map[str(_pog)] = str(_cl).strip()
                    _pog_cluster_map[_nca(_pog)] = str(_cl).strip()
        except Exception:
            _pog_cluster_map = {}
        _dg_label_parts = []
        if str(_sel_dg_code or "").strip():
            _dg_label_parts.append(str(_sel_dg_code).strip())
        if str(_sel_dg_name or "").strip() and str(_sel_dg_name).strip().upper() != "ALL":
            _dg_label_parts.append(str(_sel_dg_name).strip())
        _dg_label = " - ".join(_dg_label_parts) or "All DG"
        _submitted_dt = datetime.now()
        _dg_key = f"{_sel_dg_code or 'ALL'}|{_sel_dg_name or 'ALL'}"
        _packet_key = f"{_dg_key}|{_submitted_dt.strftime('%Y%m%d%H%M%S%f')}"
        _packets_key = f"vw_submit_reports_{p}"
        _packets = st.session_state.get(_packets_key, {})
        if not isinstance(_packets, dict):
            _packets = {}
        _saved_packets = _rs_load_report_packets(p) if not _packets else {}
        if isinstance(_saved_packets, dict) and _saved_packets:
            _saved_packets.update(_packets)
            _packets = _saved_packets
        _packets[_packet_key] = {
            "label": _dg_label,
            "dg_key": _dg_key,
            "df": _payload,
            "pog_cluster_map": _pog_cluster_map,
            "submitted_at": _submitted_dt.strftime("%Y-%m-%d %H:%M:%S"),
            "submitted_order": _submitted_dt.strftime("%Y%m%d%H%M%S%f"),
        }
        st.session_state[_packets_key] = _packets
        _rs_save_report_packets(p, _packets)
        st.session_state[f"vw_submit_{p}"] = _payload
        st.session_state["management_report_source"] = "all"
        st.session_state["management_report_latest_source"] = p
        st.session_state["management_report_packet"] = _packet_key

    def _scan_blank_cells(_df: pd.DataFrame, max_rows: int = 50, max_cells: int = 250):
        _skip_cols = {"__rs_row_key", "__rs_added_row", "__rs_last_edit"}
        _scan = _df.head(max_rows).reset_index(drop=True)
        _scan = _scan.drop(columns=[c for c in _skip_cols if c in _scan.columns], errors="ignore")
        if _scan.empty:
            return []
        _str = _scan.astype("string")
        _blank_mask = _scan.isna() | _str.apply(
            lambda _col: _col.str.strip().isin(("", "nan", "None", "<NA>"))
        )
        _locs = _blank_mask.stack()
        _idxs = list(_locs[_locs].head(max_cells).index)
        return [
            {"row": int(_ri) + 1, "column": str(_col)}
            for _ri, _col in _idxs
        ]

    def _is_blank_submit_value(_val) -> bool:
        if _val is None:
            return True
        try:
            if pd.isna(_val):
                return True
        except Exception:
            pass
        return str(_val).strip().lower() in ("", "nan", "none", "<na>")

    def _new_item_required_missing(_df: pd.DataFrame) -> list[dict]:
        if not isinstance(_df, pd.DataFrame) or "__rs_added_row" not in _df.columns:
            return []
        _skip_cols = {
            "__rs_row_key",
            "__rs_added_row",
            "__rs_last_edit",
            "Status",
            "Check Range To-be Waterfall",
        }
        _avg_required_name = locals().get(
            "_AVG_U_STD",
            "Avg Units 52wk/ Forecast new item sales",
        )
        _avg_col = _avg_required_name if _avg_required_name in _df.columns else None
        _missing = []
        _added_mask = _df["__rs_added_row"].astype(str).str.lower().isin(("true", "1", "yes"))
        for _ri, _row in _df.loc[_added_mask].reset_index(drop=False).iterrows():
            _has_input = False
            for _col in _df.columns:
                if _col in _skip_cols:
                    continue
                if not _is_blank_submit_value(_row.get(_col, "")):
                    _has_input = True
                    break
            if not _has_input:
                continue

            _row_no = int(_row.get("index", _ri)) + 1
            _id_raw = "" if "ID" not in _df.columns else str(_row.get("ID", "")).strip()
            _id_digits = _re.sub(r"\D+", "", _id_raw)
            if not _re.fullmatch(r"\d{9}", _id_digits):
                _missing.append({
                    "row": _row_no,
                    "column": "ID",
                    "reason": "ID must be 9 digits",
                })

            if _avg_col is None:
                _missing.append({
                    "row": _row_no,
                    "column": _avg_required_name,
                    "reason": "Required 52wk column is missing",
                })
            else:
                _avg_raw = _row.get(_avg_col, "")
                _avg_text = "" if _avg_raw is None else str(_avg_raw).replace(",", "").strip()
                try:
                    _avg_ok = bool(_avg_text) and not pd.isna(float(_avg_text))
                except Exception:
                    _avg_ok = False
                if not _avg_ok:
                    _missing.append({
                        "row": _row_no,
                        "column": _avg_col,
                        "reason": "52wk value is required and must be numeric",
                    })
        return _missing

    _missing_blanks = st.session_state.get(_submit_missing_key)
    if _missing_blanks:
        def _render_missing_prompt():
            st.warning(
                f"Found {len(_missing_blanks):,} blank cell(s). "
                "Blank cells are highlighted in yellow."
            )
            st.write("Do you want to add missing data, or submit anyway?")
            _missing_df = pd.DataFrame(_missing_blanks[:50])
            _sel = st.dataframe(
                _missing_df,
                hide_index=True,
                height=180,
                key=f"{p}_missing_blank_picker",
                on_select="rerun",
                selection_mode="single-row",
            )
            _sel_rows = []
            try:
                _sel_rows = list(_sel.selection.rows)
            except Exception:
                _sel_rows = []
            if _sel_rows:
                _picked = _missing_df.iloc[int(_sel_rows[0])]
                _scan_detail = locals().get("_tdf_display", df_view).reset_index(drop=True)
                _row_pos = int(_picked["row"]) - 1
                _col_name = str(_picked["column"])
                if 0 <= _row_pos < len(_scan_detail):
                    _detail_row = _scan_detail.iloc[_row_pos]
                    _item = (
                        _detail_row.get("Item Name", "")
                        if "Item Name" in _scan_detail.columns else ""
                    )
                    _item_id = (
                        _detail_row.get("ID", "")
                        if "ID" in _scan_detail.columns else ""
                    )
                    st.info(
                        f"Selected blank: row {int(_picked['row'])}, "
                        f"column `{_col_name}` | ID: {_item_id} | Item: {_item}"
                    )
                    _detail = pd.DataFrame({
                        "column": list(_scan_detail.columns),
                        "value": [_detail_row.get(c, "") for c in _scan_detail.columns],
                    })
                    _detail["selected"] = _detail["column"].eq(_col_name).map(
                        {True: "< blank cell", False: ""}
                    )
                    st.dataframe(_detail, hide_index=True, height=220)
            _mb1, _mb2 = st.columns([1, 1])
            if _mb1.button("Add missing data", key=f"{p}_add_missing_data", use_container_width=True):
                st.session_state[_blank_hl_key] = True
                st.session_state.pop(_submit_missing_key, None)
                st.rerun()
            if _mb2.button("Submit anyway", key=f"{p}_submit_anyway", type="primary", use_container_width=True):
                _submit_full_internal = _submit_payload_df(keep_internal=True)
                _required_missing = _new_item_required_missing(_submit_full_internal)
                if _required_missing:
                    st.error("New item rows must have ID and 52wk value before submit.")
                    st.dataframe(pd.DataFrame(_required_missing), hide_index=True, height=180)
                    return
                _submit_to_report_packet(_submit_full_internal)
                st.session_state.pop(_submit_missing_key, None)
                st.session_state[_blank_hl_key] = False
                st.switch_page("pages/report.py")

        if hasattr(st, "dialog"):
            @st.dialog("Missing data found")
            def _missing_data_dialog():
                _render_missing_prompt()
            _missing_data_dialog()
        else:
            _render_missing_prompt()
    _can_export_sspog_layout = (
        p in ("ns", "ss")
        and bool(_dyn_pog_cols)
        and callable(locals().get("_sspog_layout_xlsx_bytes"))
    )
    if _can_export_sspog_layout:
        _submit_col, _export_col = st.columns([1.45, 1.0])
    else:
        _submit_col = st.container()
        _export_col = None

    with _submit_col:
        if st.button("SUBMIT TO REPORT", key=f"{p}_btn_sub", type="primary", use_container_width=True):
            _submit_full_internal = _submit_payload_df(keep_internal=True)
            _required_missing = _new_item_required_missing(_submit_full_internal)
            if _required_missing:
                st.error("New item rows must have ID and 52wk value before submit.")
                st.dataframe(pd.DataFrame(_required_missing), hide_index=True, height=180)
                st.stop()
            _scan_df = _submit_full_internal.head(50)
            _missing = _scan_blank_cells(_scan_df)
            if _missing:
                st.session_state[_submit_missing_key] = _missing
                st.session_state[_blank_hl_key] = True
                st.rerun()
            else:
                st.session_state.pop(_submit_missing_key, None)
                st.session_state[_blank_hl_key] = False
                _submit_to_report_packet(_submit_full_internal)
                st.switch_page("pages/report.py")
    if _export_col is not None:
        with _export_col:
            st.download_button(
                "SAVE AS .XLSX",
                data=_sspog_layout_xlsx_bytes(),
                file_name=f"rangesheet_sspog_layout_{_sel_dg_code or 'ALL'}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"{p}_layout_export_xlsx_submit",
                use_container_width=True,
            )
    if st.session_state.get(f"vw_submit_{p}") is not None:
        st.caption(f"✅ {len(st.session_state[f'vw_submit_{p}']):,} rows submitted to Report")

# Pre-filter SSPOG / Non-SSPOG rows
_pog_col_main = next((c for c in all_cols if "pog" in c.lower() and "cluster" in c.lower()), None)
_sspog_df = (
    merged[
        merged[_pog_col_main].astype(str).str.contains("SSPOG", na=False) &
        ~merged[_pog_col_main].astype(str).str.contains("Non", na=False)
    ] if _pog_col_main else merged
)
# Inverse filter — consistent with .str.contains("Non-SSPOG") used elsewhere in the app
_nonsspog_df = (
    merged[merged[_pog_col_main].astype(str).str.contains("Non-SSPOG", na=False)]
    if _pog_col_main else merged
)

# ── Sheet tabs (scrollable via st.tabs) ───────────────────────────────────────
# Single source of truth for tab labels — defined HERE so Streamlit's file watcher
# picks up changes immediately without stale module-import caches from utils/shared.py.
# RS_SHEETS (imported) stays at 7 entries for other pages; this local list owns the UI.
_TAB_LABELS = [
    "Range Sheet_SSPOG",
    "Range Sheet_Non-SSPOG", # (Simple)
    # "Range Sheet_Non-SSPOG",
    "StoreApply_SSPOG",
    "5.1 ItembyStore",
    "5.2 ItembyStore_SC",
    "5.3 Upload_product_library",
    "5.4 Upload to Citrix",
]
assert len(_TAB_LABELS) == len(set(_TAB_LABELS)), \
    f"Duplicate tab labels: {[l for l in _TAB_LABELS if _TAB_LABELS.count(l) > 1]}"
_sheet_tabs = st.tabs(_TAB_LABELS)
_tab = {name: t for name, t in zip(_TAB_LABELS, _sheet_tabs)}
assert len(_tab) == len(_TAB_LABELS), "Bug: tab dict is shorter than label list"


with _tab["Range Sheet_SSPOG"]:
    # ── Step 1: Find pinned HDET (large) and A5 files ────────────────────────
    _ns_hdet_path = None
    _ns_a5_name   = None
    for _nsmeta in load_admin_manifest():
        _nsp = os.path.join(BASE_DIR, "uploads", _nsmeta["name"])
        if not os.path.exists(_nsp):
            continue
        if "hdet" in _nsmeta["name"].lower() and is_large_file(_nsp):
            _ns_hdet_path = _nsp
        elif "a5" in _nsmeta["name"].lower() and not is_large_file(_nsp):
            _ns_a5_name = _nsmeta["name"]

    # ── HDET controls: DG dropdown + Load button ──────────────────────────────
    if _ns_hdet_path:
        _ns_hfname  = os.path.basename(_ns_hdet_path)
        _ns_src_lbl = st.session_state.get("_ns_hdet_src", "")

        # Source/cache details are intentionally hidden on the review page.

        # Step 1: Convert HDET CSV → Parquet once (makes all future ops much faster)
        if "ns_pq_ready" not in st.session_state:
            ensure_hdet_parquet(_ns_hdet_path)
            st.session_state["ns_pq_ready"] = True

        # Step 2: Build DG index — reads from Parquet (fast) or CSV; result saved to disk JSON
        if "ns_dg_index" not in st.session_state:
            with st.spinner("Building DG dropdown index (one-time, saved to disk)…"):
                _ns_idx = get_dg_index(_ns_hdet_path)
            st.session_state["ns_dg_index"] = _ns_idx
        _ns_idx = st.session_state["ns_dg_index"]

        _ns_codes        = _ns_idx.get("codes", [])
        _ns_names        = _ns_idx.get("names", [])
        _ns_code_to_name = _ns_idx.get("code_to_name", {})
        _ns_name_to_code = {v: k for k, v in _ns_code_to_name.items()}

        # DG Code + DG Name controls: synced and auto-loading.
        _ALL = "— All (preview) —"
        _code_opts = [_ALL] + _ns_codes
        _name_opts = [_ALL] + _ns_names
        _saved_ns_code = st.session_state.get("ns_selected_dg_code")
        _saved_ns_name = st.session_state.get("ns_selected_dg_name")
        if st.session_state.get("ns_sel_dg_code") not in _code_opts and _saved_ns_code in _code_opts:
            st.session_state["ns_sel_dg_code"] = _saved_ns_code
        if st.session_state.get("ns_sel_dg_name") not in _name_opts and _saved_ns_name in _name_opts:
            st.session_state["ns_sel_dg_name"] = _saved_ns_name
        if st.session_state.get("ns_sel_dg_code") not in _code_opts:
            st.session_state["ns_sel_dg_code"] = _ALL
        if st.session_state.get("ns_sel_dg_name") not in _name_opts:
            st.session_state["ns_sel_dg_name"] = _ALL

        def _ns_clear_dg_filter():
            st.session_state["ns_sel_dg_code"] = _ALL
            st.session_state["ns_sel_dg_name"] = _ALL
            st.session_state["ns_selected_dg_code"] = _ALL
            st.session_state["ns_selected_dg_name"] = _ALL
            for _k in ["ns_hdet_df", "_ns_hdet_src", "ns_lf_dg_sig",
                       "_ns_loaded_dg_code", "_ns_loaded_dg_name"]:
                st.session_state.pop(_k, None)

        def _ns_sync_from_code():
            _code = st.session_state.get("ns_sel_dg_code", _ALL)
            if _code == _ALL:
                st.session_state["ns_sel_dg_name"] = _ALL
            else:
                st.session_state["ns_sel_dg_name"] = _ns_code_to_name.get(_code, _ALL) or _ALL
            st.session_state["ns_selected_dg_code"] = st.session_state.get("ns_sel_dg_code", _ALL)
            st.session_state["ns_selected_dg_name"] = st.session_state.get("ns_sel_dg_name", _ALL)

        def _ns_sync_from_name():
            _name = st.session_state.get("ns_sel_dg_name", _ALL)
            if _name == _ALL:
                st.session_state["ns_sel_dg_code"] = _ALL
            else:
                st.session_state["ns_sel_dg_code"] = _ns_name_to_code.get(_name, _ALL) or _ALL
            st.session_state["ns_selected_dg_code"] = st.session_state.get("ns_sel_dg_code", _ALL)
            st.session_state["ns_selected_dg_name"] = st.session_state.get("ns_sel_dg_name", _ALL)

        _dg_c1, _dg_x1, _dg_c2, _dg_x2 = st.columns([2, 0.16, 2, 0.16])
        with _dg_c1:
            _ns_sel_code = st.selectbox(
                "DG Code", _code_opts,
                key="ns_sel_dg_code", label_visibility="visible",
                on_change=_ns_sync_from_code)
        with _dg_x1:
            st.write("")
            st.button("×", key="ns_clear_dg_code", help="Clear DG filter",
                      use_container_width=True, on_click=_ns_clear_dg_filter)
        with _dg_c2:
            _ns_sel_name = st.selectbox(
                "DG Name", _name_opts,
                key="ns_sel_dg_name", label_visibility="visible",
                on_change=_ns_sync_from_name)
        with _dg_x2:
            st.write("")
            st.button("×", key="ns_clear_dg_name", help="Clear DG filter",
                      use_container_width=True, on_click=_ns_clear_dg_filter)

        # Resolve: if DG Name picked, map to code
        _ns_dg_query = None
        _ns_sel_code = st.session_state.get("ns_sel_dg_code", _ALL)
        _ns_sel_name = st.session_state.get("ns_sel_dg_name", _ALL)
        if _ns_sel_code != _ALL:
            _ns_dg_query = _ns_sel_code
            _ns_sel_name = _ns_code_to_name.get(_ns_dg_query, _ns_sel_name)
        elif _ns_sel_name != _ALL:
            _ns_dg_query = _ns_name_to_code.get(_ns_sel_name, _ns_sel_name)
            _ns_sel_code = _ns_dg_query
        st.session_state["ns_selected_dg_code"] = _ns_sel_code
        st.session_state["ns_selected_dg_name"] = _ns_sel_name

        if _ns_dg_query:
            _sig = f"{_ns_hdet_path}|{_ns_dg_query}"
            if st.session_state.get("ns_lf_dg_sig") != _sig:
                with st.spinner(f"Loading DG '{_ns_dg_query}' from HDET…"):
                    _loaded = load_large_file_by_dg(_ns_hdet_path, _ns_dg_query)
                st.session_state["ns_hdet_df"]   = _loaded
                st.session_state["_ns_hdet_src"] = f"DG = {_ns_dg_query} ({len(_loaded):,} rows)"
                st.session_state["ns_lf_dg_sig"] = _sig
                st.session_state["_ns_loaded_dg_code"] = _ns_dg_query
                st.session_state["_ns_loaded_dg_name"] = _ns_code_to_name.get(_ns_dg_query, "")
        else:
            for _k in ["ns_hdet_df", "_ns_hdet_src", "ns_lf_dg_sig",
                       "_ns_loaded_dg_code", "_ns_loaded_dg_name"]:
                st.session_state.pop(_k, None)

    # Require a DG to be picked and loaded before showing any data
    _ns_hdet_df = st.session_state.get("ns_hdet_df")
    if _ns_hdet_path and _ns_hdet_df is None:
        st.info("👆 Select a DG Code or DG Name above to view planogram data.")
    else:
        _ns_combined = _dedup(_ns_hdet_df.copy()) if _ns_hdet_df is not None else merged
        _ns_idx = st.session_state.get("ns_dg_index", {})
        _render_sheet_content(_ns_combined, "ns",
                              dg_col_hint=_ns_idx.get("dg_col"),
                              large_file_path=_ns_hdet_path,
                              dg_options=_ns_idx.get("codes", []),
                              fixed_dg_code=st.session_state.get("_ns_loaded_dg_code"),
                              fixed_dg_name=st.session_state.get("_ns_loaded_dg_name"))

with _tab["Range Sheet_Non-SSPOG"]: # Simple — p="ns2" avoids key collision with HDET tab (p="ns")
    _render_sheet_content(_nonsspog_df, "ns2")
# with _tab["Range Sheet_Non-SSPOG"]:
#     _render_sheet_content(_nonsspog_df, "nsv")
with _tab["StoreApply_SSPOG"]:
    _render_sheet_content(_sspog_df, "sa")
with _tab["5.1 ItembyStore"]:
    _IB_COLS = [
        "store_no", "store_name", "ID", "ProductDescription",
        "POG_STATUS", "Display Group", "Display group desc",
        "ForecastSales", "TH_Tot_Sales_Value_52WK", "TH_Tot_Sales_Volume_52WK",
        "DaysSupply", "MaxDOS",
    ]
    _IB_TEXT_COLS = ["store_no", "store_name", "ID", "ProductDescription",
                     "POG_STATUS", "Display Group", "Display group desc"]
    _IB_COL_CFG = {
        "store_no":                 st.column_config.TextColumn("store_no",                 width="small"),
        "store_name":               st.column_config.TextColumn("store_name",               width="medium"),
        "ID":                       st.column_config.TextColumn("ID",                       width="small"),
        "ProductDescription":       st.column_config.TextColumn("ProductDescription",       width="large"),
        "POG_STATUS":               st.column_config.TextColumn("POG_STATUS",               width="small"),
        "Display Group":            st.column_config.TextColumn("Display Group",            width="small"),
        "Display group desc":       st.column_config.TextColumn("Display group desc",       width="medium"),
        "ForecastSales":            st.column_config.NumberColumn("ForecastSales",          width="small", format="%.2f"),
        "TH_Tot_Sales_Value_52WK":  st.column_config.NumberColumn("TH_Tot_Sales_Value_52WK",  width="small", format="%.2f"),
        "TH_Tot_Sales_Volume_52WK": st.column_config.NumberColumn("TH_Tot_Sales_Volume_52WK", width="small", format="%.0f"),
        "DaysSupply":               st.column_config.NumberColumn("DaysSupply",             width="small", format="%.0f"),
        "MaxDOS":                   st.column_config.NumberColumn("MaxDOS",                 width="small", format="%.0f"),
    }

    # ── Find HDET file (large .txt pinned by admin) ───────────────────────────
    _hdet_path = None
    for _hmeta in load_admin_manifest():
        _hp = os.path.join(BASE_DIR, "uploads", _hmeta["name"])
        if "hdet" in _hmeta["name"].lower() and os.path.exists(_hp) and is_large_file(_hp):
            _hdet_path = _hp
            break

    def _build_ib(hdet_df: pd.DataFrame) -> pd.DataFrame:
        """Map hdet_df columns → _IB_COLS. For columns absent in HDET, backfill
        from the main rangesheet data (merged). Leave null if both are missing."""
        _hmap = {col: next((c for c in hdet_df.columns if _nca(c) == _nca(col)), None)
                 for col in _IB_COLS}
        _mmap = {col: next((c for c in merged.columns if _nca(c) == _nca(col)), None)
                 for col in _IB_COLS}
        _hn = len(hdet_df)
        _out = {}
        for col in _IB_COLS:
            if _hmap[col] is not None:
                _out[col] = hdet_df[_hmap[col]].reset_index(drop=True)
            elif _mmap[col] is not None:
                _mv = merged[_mmap[col]].reset_index(drop=True)
                if len(_mv) >= _hn:
                    _out[col] = _mv.iloc[:_hn].reset_index(drop=True)
                else:
                    _out[col] = pd.concat(
                        [_mv, pd.Series([None] * (_hn - len(_mv)))], ignore_index=True)
            else:
                _out[col] = pd.Series([None] * _hn)
        return _cast_text_cols(pd.DataFrame(_out), _IB_TEXT_COLS)

    # ── Auto-load on first access only (never auto-clear; keep last version) ──
    _hdet_mtime = (f"{os.path.getmtime(_hdet_path):.0f}" if _hdet_path else None)
    if "ib_data" not in st.session_state:
        if _hdet_path:
            with st.spinner("Loading Item by Store data from HDET…"):
                _auto_df = read_large_file_head(_hdet_path, n_rows=500)
            st.session_state.ib_data = _build_ib(_auto_df)
            st.session_state["_ib_hdet_mtime"] = _hdet_mtime
            st.session_state["_ib_source"] = f"HDET preview (500 rows)"
        else:
            st.session_state.ib_data = _cast_text_cols(
                _fill_from_db(_IB_COLS, merged), _IB_TEXT_COLS)
            st.session_state["_ib_source"] = "Rangesheet data"

    # ── HDET controls (DG filter + reset) ────────────────────────────────────
    if _hdet_path:
        _hfname = os.path.basename(_hdet_path)
        _src_label = st.session_state.get("_ib_source", "")
        st.markdown(
            f"<div style='font-size:11px;color:#2BBFA4;margin-bottom:6px;'>"
            f"Source: <strong>{_hfname}</strong>"
            + (f" · {_src_label}" if _src_label else "")
            + " · Enter DG code below to load a full filtered slice.</div>",
            unsafe_allow_html=True,
        )
        _hc1, _hc2, _hc3, _ = st.columns([1.8, 1.4, 1.6, 3.2])
        with _hc1:
            _ib_dg_inp = st.text_input(
                "DG Code", key="ib_hdet_dg",
                placeholder="e.g. 101",
                label_visibility="collapsed",
            )
        with _hc2:
            _ib_reload = st.button("Load by DG", key="ib_hdet_reload", use_container_width=True)
        with _hc3:
            _ib_reset_hdet = st.button("↺ Reload HDET Preview", key="ib_hdet_reset", use_container_width=True)

        if _ib_reset_hdet:
            with st.spinner(f"Reloading preview from {_hfname}…"):
                _auto_df = read_large_file_head(_hdet_path, n_rows=500)
            st.session_state.ib_data = _build_ib(_auto_df)
            st.session_state["_ib_source"] = "HDET preview (500 rows)"
            st.rerun()

        if _ib_reload:
            _dg_val = _ib_dg_inp.strip()
            if _dg_val:
                with st.spinner(f"Loading DG={_dg_val} from {_hfname}…"):
                    _hdet_df = load_large_file_by_dg(_hdet_path, _dg_val)
                if not _hdet_df.empty:
                    st.session_state.ib_data = _build_ib(_hdet_df)
                    st.session_state["_ib_source"] = f"HDET · DG={_dg_val} ({len(_hdet_df):,} rows)"
                    st.rerun()
                else:
                    st.warning(f"No rows found for DG={_dg_val!r} in {_hfname}")
            else:
                st.info("Enter a DG code first, then click Load by DG.")

    # ── Toolbar ───────────────────────────────────────────────────────────────
    _ib_c1, _ib_c2, _ib_c3, _ = st.columns([1.1, 1.0, 1.4, 4.5])
    with _ib_c1:
        if st.button("＋ Add Row", key="ib_add_row", use_container_width=True):
            _empty = pd.DataFrame([{c: None for c in _IB_COLS}])
            st.session_state.ib_data = pd.concat(
                [st.session_state.ib_data, _empty], ignore_index=True)
            st.rerun()
    with _ib_c2:
        if st.button("↺ Reset", key="ib_clear", use_container_width=True):
            st.session_state.ib_data = _cast_text_cols(
                _fill_from_db(_IB_COLS, merged), _IB_TEXT_COLS)
            st.session_state["_ib_source"] = "Rangesheet data"
            st.session_state.pop("vw_submit_51", None)
            st.rerun()
    with _ib_c3:
        if st.button("✅ Submit to Report", key="ib_submit", use_container_width=True,
                     type="primary"):
            _to_send = st.session_state.ib_data.dropna(how="all")
            st.session_state["vw_submit_51"] = _to_send.reset_index(drop=True)
            st.success(f"Submitted {len(_to_send):,} rows → go to Report page to export.")

    _ib_edited = st.data_editor(
        st.session_state.ib_data,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config=_IB_COL_CFG,
        height=420,
        key="ib_editor",
    )
    st.session_state.ib_data = _ib_edited
    st.caption(f"{len(_ib_edited):,} rows")

    if st.session_state.get("vw_submit_51") is not None:
        st.caption(f"✅ {len(st.session_state['vw_submit_51']):,} rows submitted to Report")
with _tab["5.2 ItembyStore_SC"]:
    _IBS_COLS = [
        "store", "item on POG", "Item Name", "DG", "DG_Des",
        "Status", "Forecast", "avg per wk", "Coperate Status",
    ]
    _IBS_COL_CFG = {
        "store":           st.column_config.TextColumn("store",           width="small"),
        "item on POG":     st.column_config.TextColumn("item on POG",     width="small"),
        "Item Name":       st.column_config.TextColumn("Item Name",       width="large"),
        "DG":              st.column_config.TextColumn("DG",              width="small"),
        "DG_Des":          st.column_config.TextColumn("DG_Des",          width="medium"),
        "Status":          st.column_config.TextColumn("Status",          width="small"),
        "Forecast":        st.column_config.NumberColumn("Forecast",      width="small",  format="%.2f"),
        "avg per wk":      st.column_config.NumberColumn("avg per wk",    width="small",  format="%.2f"),
        "Coperate Status": st.column_config.TextColumn("Coperate Status", width="medium"),
    }
    if "ibs_sc_data" not in st.session_state or st.session_state.get("_ibs_db_sig") != _db_sig:
        st.session_state.ibs_sc_data = _fill_from_db(_IBS_COLS, merged)
        st.session_state["_ibs_db_sig"] = _db_sig

    # ── Toolbar ──────────────────────────────────────────────────────────────
    _ibs_c1, _ibs_c2, _ibs_c3, _ibs_c4 = st.columns([1.1, 1.0, 1.4, 4.5])
    with _ibs_c1:
        if st.button("＋ Add Row", key="ibs_add_row", use_container_width=True):
            _empty = pd.DataFrame([{c: None for c in _IBS_COLS}])
            st.session_state.ibs_sc_data = pd.concat(
                [st.session_state.ibs_sc_data, _empty], ignore_index=True
            )
            st.rerun()
    with _ibs_c2:
        if st.button("↺ Reset", key="ibs_clear", use_container_width=True):
            st.session_state.ibs_sc_data = _fill_from_db(_IBS_COLS, merged)
            st.session_state.pop("vw_submit_52", None)
            st.rerun()
    with _ibs_c3:
        if st.button("✅ Submit to Report", key="ibs_submit", use_container_width=True,
                     type="primary"):
            _to_send = st.session_state.ibs_sc_data.dropna(how="all")
            st.session_state["vw_submit_52"] = _to_send.reset_index(drop=True)
            st.success(f"Submitted {len(_to_send):,} rows → go to Report page to export.")

    # ── Editable table ────────────────────────────────────────────────────────
    _ibs_edited = st.data_editor(
        st.session_state.ibs_sc_data,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config=_IBS_COL_CFG,
        height=420,
        key="ibs_editor",
    )
    st.session_state.ibs_sc_data = _ibs_edited

    if st.session_state.get("vw_submit_52") is not None:
        st.caption(f"✅ {len(st.session_state['vw_submit_52']):,} rows submitted to Report")
with _tab["5.3 Upload_product_library"]:
    _PRODLIB_COLS = ["ID", "Product Description", "Mod_structure_fixture"]
    if "vw_prodlib_data" not in st.session_state or st.session_state.get("_pl_db_sig") != _db_sig:
        st.session_state.vw_prodlib_data = _fill_from_db(_PRODLIB_COLS, merged)
        st.session_state["_pl_db_sig"] = _db_sig

    # ── Toolbar ──────────────────────────────────────────────────────────────
    _pl_c1, _pl_c2, _pl_c3, _pl_c4 = st.columns([1.1, 1.0, 1.4, 4.5])
    with _pl_c1:
        if st.button("＋ Add Row", key="pl_add_row", use_container_width=True):
            _empty = pd.DataFrame([{c: None for c in _PRODLIB_COLS}])
            st.session_state.vw_prodlib_data = pd.concat(
                [st.session_state.vw_prodlib_data, _empty], ignore_index=True
            )
            st.rerun()
    with _pl_c2:
        if st.button("↺ Reset", key="pl_clear", use_container_width=True):
            st.session_state.vw_prodlib_data = _fill_from_db(_PRODLIB_COLS, merged)
            st.session_state.pop("vw_submit_53", None)
            st.rerun()
    with _pl_c3:
        if st.button("✅ Submit to Report", key="pl_submit", use_container_width=True,
                     type="primary"):
            _to_send = st.session_state.vw_prodlib_data.dropna(how="all")
            st.session_state["vw_submit_53"] = _to_send.reset_index(drop=True)
            st.success(f"Submitted {len(_to_send):,} rows → go to Report page to export.")

    _prodlib_edited = st.data_editor(
        st.session_state.vw_prodlib_data,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "ID": st.column_config.TextColumn("ID", width="small"),
            "Product Description": st.column_config.TextColumn("Product Description", width="large"),
            "Mod_structure_fixture": st.column_config.TextColumn("Mod_structure_fixture", width="medium"),
        },
        key="vw_prodlib_editor",
        height=460,
    )
    st.session_state.vw_prodlib_data = _prodlib_edited
    st.caption(f"{len(_prodlib_edited):,} rows")

    if st.session_state.get("vw_submit_53") is not None:
        st.caption(f"✅ {len(st.session_state['vw_submit_53']):,} rows submitted to Report")
with _tab["5.4 Upload to Citrix"]:
    _CITRIX_COLS = [
        "POGName", "store_no", "store_name",
        "POG_WIDTH", "POG_HEIGHT", "POG_DEPTH",
        "SQM", "FP_status", "Capacity",
    ]
    _CX_TEXT_COLS = ["POGName", "store_no", "store_name", "FP_status"]
    _CITRIX_COL_CFG = {
        "POGName":    st.column_config.TextColumn("POGName",    width="medium"),
        "store_no":   st.column_config.TextColumn("store_no",   width="small"),
        "store_name": st.column_config.TextColumn("store_name", width="medium"),
        "POG_WIDTH":  st.column_config.NumberColumn("POG_WIDTH",  width="small", format="%.2f"),
        "POG_HEIGHT": st.column_config.NumberColumn("POG_HEIGHT", width="small", format="%.2f"),
        "POG_DEPTH":  st.column_config.NumberColumn("POG_DEPTH",  width="small", format="%.2f"),
        "SQM":        st.column_config.NumberColumn("SQM",        width="small", format="%.2f"),
        "FP_status":  st.column_config.TextColumn("FP_status",  width="small"),
        "Capacity":   st.column_config.NumberColumn("Capacity",   width="small", format="%.0f"),
    }

    # ── Find HDET file ────────────────────────────────────────────────────────
    _cx_hdet_path = None
    for _cxmeta in load_admin_manifest():
        _cxp = os.path.join(BASE_DIR, "uploads", _cxmeta["name"])
        if "hdet" in _cxmeta["name"].lower() and os.path.exists(_cxp) and is_large_file(_cxp):
            _cx_hdet_path = _cxp
            break

    def _build_cx(hdet_df: pd.DataFrame) -> pd.DataFrame:
        """Fill _CITRIX_COLS: Capacity (and any other matching cols) come from
        HDET; remaining cols fall back to merged. Null if neither has them."""
        _hmap = {col: next((c for c in hdet_df.columns if _nca(c) == _nca(col)), None)
                 for col in _CITRIX_COLS}
        _mmap = {col: next((c for c in merged.columns if _nca(c) == _nca(col)), None)
                 for col in _CITRIX_COLS}
        _hn = len(hdet_df)
        _out = {}
        for col in _CITRIX_COLS:
            if _hmap[col] is not None:
                _out[col] = hdet_df[_hmap[col]].reset_index(drop=True)
            elif _mmap[col] is not None:
                _mv = merged[_mmap[col]].reset_index(drop=True)
                if len(_mv) >= _hn:
                    _out[col] = _mv.iloc[:_hn].reset_index(drop=True)
                else:
                    _out[col] = pd.concat(
                        [_mv, pd.Series([None] * (_hn - len(_mv)))], ignore_index=True)
            else:
                _out[col] = pd.Series([None] * _hn)
        return _cast_text_cols(pd.DataFrame(_out), _CX_TEXT_COLS)

    # ── Auto-load on first access; never auto-clear ───────────────────────────
    if "vw_citrix_data" not in st.session_state:
        if _cx_hdet_path:
            with st.spinner("Loading 5.4 data from HDET…"):
                _cx_auto = read_large_file_head(_cx_hdet_path, n_rows=500)
            st.session_state.vw_citrix_data = _build_cx(_cx_auto)
            st.session_state["_cx_source"] = "HDET preview (500 rows)"
        else:
            st.session_state.vw_citrix_data = _cast_text_cols(
                _fill_from_db(_CITRIX_COLS, merged), _CX_TEXT_COLS)
            st.session_state["_cx_source"] = "Rangesheet data"

    # ── HDET controls ─────────────────────────────────────────────────────────
    if _cx_hdet_path:
        _cx_hfname = os.path.basename(_cx_hdet_path)
        _cx_src = st.session_state.get("_cx_source", "")
        st.markdown(
            f"<div style='font-size:11px;color:#2BBFA4;margin-bottom:6px;'>"
            f"Source: <strong>{_cx_hfname}</strong>"
            + (f" · {_cx_src}" if _cx_src else "")
            + " · <em>Capacity</em> column pulled from HDET · enter DG to load filtered slice.</div>",
            unsafe_allow_html=True,
        )
        _cxh1, _cxh2, _cxh3, _ = st.columns([1.8, 1.4, 1.6, 3.2])
        with _cxh1:
            _cx_dg_inp = st.text_input(
                "DG Code", key="cx_hdet_dg",
                placeholder="e.g. 101",
                label_visibility="collapsed",
            )
        with _cxh2:
            _cx_reload = st.button("Load by DG", key="cx_hdet_reload", use_container_width=True)
        with _cxh3:
            _cx_reset_hdet = st.button("↺ Reload HDET Preview", key="cx_hdet_reset", use_container_width=True)

        if _cx_reset_hdet:
            with st.spinner(f"Reloading preview from {_cx_hfname}…"):
                _cx_auto = read_large_file_head(_cx_hdet_path, n_rows=500)
            st.session_state.vw_citrix_data = _build_cx(_cx_auto)
            st.session_state["_cx_source"] = "HDET preview (500 rows)"
            st.rerun()

        if _cx_reload:
            _cx_dg = _cx_dg_inp.strip()
            if _cx_dg:
                with st.spinner(f"Loading DG={_cx_dg} from {_cx_hfname}…"):
                    _cx_hdet_df = load_large_file_by_dg(_cx_hdet_path, _cx_dg)
                if not _cx_hdet_df.empty:
                    st.session_state.vw_citrix_data = _build_cx(_cx_hdet_df)
                    st.session_state["_cx_source"] = f"HDET · DG={_cx_dg} ({len(_cx_hdet_df):,} rows)"
                    st.rerun()
                else:
                    st.warning(f"No rows found for DG={_cx_dg!r} in {_cx_hfname}")
            else:
                st.info("Enter a DG code first, then click Load by DG.")

    # ── Toolbar ──────────────────────────────────────────────────────────────
    _cx_c1, _cx_c2, _cx_c3, _ = st.columns([1.1, 1.0, 1.4, 4.5])
    with _cx_c1:
        if st.button("＋ Add Row", key="cx_add_row", use_container_width=True):
            _empty = pd.DataFrame([{c: None for c in _CITRIX_COLS}])
            st.session_state.vw_citrix_data = pd.concat(
                [st.session_state.vw_citrix_data, _empty], ignore_index=True
            )
            st.rerun()
    with _cx_c2:
        if st.button("↺ Reset", key="cx_clear", use_container_width=True):
            st.session_state.vw_citrix_data = _cast_text_cols(
                _fill_from_db(_CITRIX_COLS, merged), _CX_TEXT_COLS)
            st.session_state["_cx_source"] = "Rangesheet data"
            st.session_state.pop("vw_submit_54", None)
            st.rerun()
    with _cx_c3:
        if st.button("✅ Submit to Report", key="cx_submit", use_container_width=True,
                     type="primary"):
            _to_send = st.session_state.vw_citrix_data.dropna(how="all")
            st.session_state["vw_submit_54"] = _to_send.reset_index(drop=True)
            st.success(f"Submitted {len(_to_send):,} rows → go to Report page to export.")

    _citrix_edited = st.data_editor(
        st.session_state.vw_citrix_data,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config=_CITRIX_COL_CFG,
        height=420,
        key="cx_editor",
    )
    st.session_state.vw_citrix_data = _citrix_edited
    st.caption(f"{len(_citrix_edited):,} rows")

    if st.session_state.get("vw_submit_54") is not None:
        st.caption(f"✅ {len(st.session_state['vw_submit_54']):,} rows submitted to Report")

if False:
    _meta = st.session_state.rangesheet_meta

    _dg_c, _arch_c, _leg_c = st.columns([1.05, 2.9, 0.75])

    # ── Display Group card — connected DG Code / DG Name search ─────────────
    # Detect DG-like columns in the data
    _dg_code_col = (
        next((c for c in all_cols if _nc(c) in ("dg code", "dg_code", "dg")), None)
        or next((c for c in all_cols if "dg" in _nc(c) and "code" in _nc(c)), None)
        or next((c for c in all_cols if "department" in _nc(c)), None)
    )
    _dg_name_col = (
        next((c for c in all_cols if _nc(c) in ("dg name", "dg_name")), None)
        or next((c for c in all_cols if "dg" in _nc(c) and "name" in _nc(c)), None)
        or next((c for c in all_cols if "section" in _nc(c)), None)
    )

    # Unique DG Code options (from data or from meta)
    if _dg_code_col:
        _dg_code_opts = ["ALL"] + sorted(
            merged[_dg_code_col].dropna().astype(str).str.strip().unique().tolist()
        )
    else:
        _cur_code = _meta.get("dg_code", "—")
        _dg_code_opts = ["ALL"] + ([_cur_code] if _cur_code not in ("—", "") else [])

    _sel_dg_code = "ALL"  # default before widget renders
    _sel_dg_name = "ALL"

    with _dg_c:
        st.markdown("""
<div style="background:#fff;border-radius:14px;border:1px solid #E8E3DC;
            padding:14px 14px 0 14px;">
    <div style="font-size:10px;font-weight:700;color:#888;text-transform:uppercase;
                letter-spacing:.08em;margin-bottom:10px;">Display Group</div>
</div>""", unsafe_allow_html=True)

        with st.container():
            # DG Code — searchable
            _prev_code = st.session_state.get("_vw_dg_code_prev", "ALL")
            _sel_dg_code = st.selectbox(
                "DG CODE",
                _dg_code_opts,
                index=(_dg_code_opts.index(_prev_code)
                       if _prev_code in _dg_code_opts else 0),
                key="vw_dg_code",
                label_visibility="visible",
            )

            # DG Name — options filtered by selected DG Code (connected)
            if _dg_code_col and _dg_name_col and _sel_dg_code != "ALL":
                _name_src = merged[merged[_dg_code_col].astype(str).str.strip() == _sel_dg_code]
                _dg_name_opts = ["ALL"] + sorted(
                    _name_src[_dg_name_col].dropna().astype(str).str.strip().unique().tolist()
                )
            elif _dg_name_col:
                _dg_name_opts = ["ALL"] + sorted(
                    merged[_dg_name_col].dropna().astype(str).str.strip().unique().tolist()
                )
            else:
                _cur_name = _meta.get("dg_name", "—")
                _dg_name_opts = ["ALL"] + ([_cur_name] if _cur_name not in ("—", "") else [])

            _prev_name = st.session_state.get("_vw_dg_name_prev", "ALL")
            # If DG Code just changed, reset DG Name to All
            if _sel_dg_code != _prev_code and _prev_name not in _dg_name_opts:
                _prev_name = "ALL"

            _sel_dg_name = st.selectbox(
                "DG NAME",
                _dg_name_opts,
                index=(_dg_name_opts.index(_prev_name)
                       if _prev_name in _dg_name_opts else 0),
                key="vw_dg_name",
                label_visibility="visible",
            )

            # Persist selections for next rerun
            st.session_state["_vw_dg_code_prev"] = _sel_dg_code
            st.session_state["_vw_dg_name_prev"] = _sel_dg_name

            # Update meta
            _meta["dg_code"] = _sel_dg_code if _sel_dg_code != "ALL" else _meta.get("dg_code", "—")
            _meta["dg_name"] = _sel_dg_name if _sel_dg_name != "ALL" else _meta.get("dg_name", "—")

            # Static fields
            _static_rows = ""
            for _sk, _sv2 in [
                ("MINOR LIVE WEEK", _meta.get("minor_live_week", "—")),
                ("MAJOR LIVE WEEK", _meta.get("major_live_week", "—")),
                ("Event Live Date", _meta.get("event_live_date", "—")),
                ("Event Des",       _meta.get("event_desc",       "—")),
            ]:
                _static_rows += (
                    f'<div style="display:flex;justify-content:space-between;'
                    f'align-items:center;margin-bottom:5px;font-size:11px;gap:4px;">'
                    f'<span style="color:#888;white-space:nowrap;">{_sk}</span>'
                    f'<span style="font-weight:600;color:#555;text-align:right;">{_sv2}</span></div>'
                )
            st.markdown(f"""
<div style="background:#fff;border-radius:0 0 14px 14px;border:1px solid #E8E3DC;
            border-top:none;padding:10px 14px 10px;">
    {_static_rows}
</div>""", unsafe_allow_html=True)

            _btn_a, _btn_b = st.columns(2)
            with _btn_a:
                st.button("SELECT", key="btn_select_dg", use_container_width=True,
                          type="primary")
            with _btn_b:
                st.button("SUBMIT TO REPORT", key="btn_submit_range",
                          use_container_width=True)


    # ── Range Architecture card (7-column, two-row merged headers) ────────────
    with _arch_c:
        # Detect TYPE source column (Item Priority first, then Status)
        _type_col = (
            next((c for c in all_cols if _nc(c) in ("item priority", "itempriority")), None)
            or next((c for c in all_cols if "item" in _nc(c) and "priority" in _nc(c)), None)
            or next((c for c in all_cols if _nc(c) == "status"), None)
        )
        _stc = next((c for c in all_cols if _nc(c) == "status"), None)

        # Detect value columns
        _price_col  = next((c for c in all_cols if "avg selling price" in _nc(c) or ("selling price" in _nc(c) and "avg" in _nc(c))), None)
        _edlp_col   = next((c for c in all_cols if "edlp price" in _nc(c)), None)
        _asis_stc   = next((c for c in all_cols if ("as-is stores applied" in _nc(c) or ("as is" in _nc(c) and "stores applied" in _nc(c))) and "to" not in _nc(c)[:4]), None)
        _tobe_stc   = next((c for c in all_cols if "to-be stores applied" in _nc(c) or "to be stores applied" in _nc(c) or "to-be stores" in _nc(c)), None)

        # Base dataframe filtered by DG selections
        _arch_base = merged.copy()
        if _sel_dg_code != "ALL" and _dg_code_col and _dg_code_col in _arch_base.columns:
            _arch_base = _arch_base[_arch_base[_dg_code_col].astype(str).str.strip() == _sel_dg_code]
        if _sel_dg_name != "ALL" and _dg_name_col and _dg_name_col in _arch_base.columns:
            _arch_base = _arch_base[_arch_base[_dg_name_col].astype(str).str.strip() == _sel_dg_name]

        TYPES = ["MAINTAIN", "NEW DELETE SOME", "DELETE SOME", "DELETE ALL", "NEW SOME", "NEWNEW"]
        _status_ov_live = st.session_state.get("ss_status_overrides", {}) or st.session_state.get("ns_status_overrides", {})
        _sticky_for_arch = [c for c in ["DG Code", "ID", "Item Name"] if c in _arch_base.columns]
        _id_col_arch = next((c for c in _arch_base.columns if _nca(c) == _nca("ID")), None)

        def _arch_status_series(base: pd.DataFrame):
            if _type_col and _type_col in base.columns:
                _asis = base[_type_col].astype(str).str.strip().str.upper()
            else:
                _asis = pd.Series(["MAINTAIN"] * len(base), index=base.index)
            _asis = _asis.replace({"": "MAINTAIN", "NAN": "MAINTAIN", "NONE": "MAINTAIN"})
            _tobe = _asis.copy()
            _used_ov = set()
            _ov_by_id = {
                str(_rk[1] if isinstance(_rk, tuple) and len(_rk) > 1 else _rk): _ov
                for _rk, _ov in _status_ov_live.items()
            }
            if _status_ov_live:
                for _idx2, _row2 in base.iterrows():
                    _rk2 = tuple(str(_row2.get(c, "")) for c in _sticky_for_arch) if _sticky_for_arch else None
                    _ov2 = _status_ov_live.get(_rk2) if _rk2 else None
                    if _ov2 and _rk2:
                        _used_ov.add(_rk2)
                    if not _ov2 and _id_col_arch:
                        _id2 = str(_row2.get(_id_col_arch, ""))
                        if _id2 not in _used_ov:
                            _ov2 = _ov_by_id.get(_id2)
                        if _ov2:
                            _used_ov.add(_id2)
                    if _ov2:
                        _tobe.at[_idx2] = str(_ov2).strip().upper()
                for _rk3, _ov3 in _status_ov_live.items():
                    _id3 = str(_rk3[1] if isinstance(_rk3, tuple) and len(_rk3) > 1 else _rk3)
                    if _rk3 in _used_ov or _id3 in _used_ov:
                        continue
                    _maint_idx = _tobe.index[_tobe == "MAINTAIN"]
                    if len(_maint_idx):
                        _tobe.at[_maint_idx[0]] = str(_ov3).strip().upper()
            return _asis, _tobe

        _asis_status, _tobe_status = _arch_status_series(_arch_base)

        def _fmt(v):
            if v == 0: return "0"
            try: return f"{int(v):,}" if v == int(v) else f"{v:,.1f}"
            except: return str(v)

        _arch_rows = []
        for _t in TYPES:
            _idx_ai = _arch_base.index[_asis_status == _t]
            _idx_tb = _arch_base.index[_tobe_status == _t]
            _idx = _idx_tb

            _ai = len(_idx_ai)
            _tb = len(_idx_tb)

            # Sale Impact AS IS / TO BE from Mer Price (AVG Selling Price × stores)
            _prices = pd.to_numeric(_arch_base.loc[_idx, _price_col], errors="coerce").fillna(0) if (_price_col and len(_idx) > 0) else pd.Series([], dtype=float)
            _asis_s = pd.to_numeric(_arch_base.loc[_idx, _asis_stc], errors="coerce").fillna(0) if (_asis_stc and len(_idx) > 0) else pd.Series([1]*len(_idx), dtype=float)
            _tobe_s = pd.to_numeric(_arch_base.loc[_idx, _tobe_stc], errors="coerce").fillna(0) if (_tobe_stc and len(_idx) > 0) else pd.Series([1]*len(_idx), dtype=float)
            _sale_ai = float((_prices * _asis_s).sum()) if len(_prices) else 0.0
            _sale_tb = float((_prices * _tobe_s).sum()) if len(_prices) else 0.0

            _marg_ai = 0.0

            _arch_rows.append({
                "type": _t, "as_is": _ai, "to_be": _tb,
                "sale_ai": _sale_ai, "sale_tb": _sale_tb,
                "sale_diff": _sale_tb - _sale_ai, "marg_ai": _marg_ai,
            })

        _t_ai    = sum(r["as_is"]     for r in _arch_rows)
        _t_tb    = sum(r["to_be"]     for r in _arch_rows)
        _t_sai   = sum(r["sale_ai"]   for r in _arch_rows)
        _t_stb   = sum(r["sale_tb"]   for r in _arch_rows)
        _t_sdiff = sum(r["sale_diff"] for r in _arch_rows)
        _t_mai   = sum(r["marg_ai"]   for r in _arch_rows)
        _pct_sale = f"{(_t_sdiff / _t_sai * 100):.1f}%" if _t_sai else "0.0%"

        # ── Build table rows ──────────────────────────────────────────────────
        _B = "border:1px solid #B8B8B8;"
        _BR = "border-right:2px solid #999;"
        tbody = ""
        for _r in _arch_rows:
            _t     = _r["type"]
            tbody += (
                f'<tr style="border-bottom:1px solid #D8D8D8;">'
                f'<td style="padding:5px 10px;font-size:11px;color:#1A1A1A;{_B}">{_t}</td>'
                f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}">{_fmt(_r["as_is"])}</td>'
                f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}{_BR}">{_fmt(_r["to_be"])}</td>'
                f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}">{_fmt(_r["sale_ai"])}</td>'
                f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}">{_fmt(_r["sale_tb"])}</td>'
                f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}{_BR}">{_fmt(_r["sale_diff"])}</td>'
                f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}">{_fmt(_r["marg_ai"])}</td>'
                f'</tr>')

        tbody += (
            f'<tr style="font-weight:800;">'
            f'<td style="padding:6px 10px;font-size:11px;font-weight:800;{_B}">TOTAL SKU</td>'
            f'<td style="padding:6px 8px;text-align:center;font-size:11px;font-weight:700;{_B}">{_fmt(_t_ai)}</td>'
            f'<td style="padding:6px 8px;text-align:center;font-size:11px;font-weight:700;{_B}{_BR}">{_fmt(_t_tb)}</td>'
            f'<td style="padding:6px 8px;text-align:center;font-size:11px;font-weight:700;{_B}">{_fmt(_t_sai)}</td>'
            f'<td style="padding:6px 8px;text-align:center;font-size:11px;font-weight:700;{_B}">{_fmt(_t_stb)}</td>'
            f'<td style="padding:6px 8px;text-align:center;font-size:11px;font-weight:700;{_B}{_BR}">{_fmt(_t_sdiff)}</td>'
            f'<td style="padding:6px 8px;text-align:center;font-size:11px;font-weight:700;{_B}">{_fmt(_t_mai)}</td>'
            f'</tr>'
            f'<tr>'
            f'<td style="padding:5px 10px;font-size:11px;font-style:italic;{_B}">% Impact</td>'
            f'<td style="{_B}"></td>'
            f'<td style="text-align:center;color:#00AA00;font-weight:700;font-size:11px;{_B}{_BR}">0.0%</td>'
            f'<td style="{_B}"></td><td style="{_B}"></td>'
            f'<td style="text-align:center;color:#00AA00;font-weight:700;font-size:11px;{_B}{_BR}">{_pct_sale}</td>'
            f'<td style="{_B}"></td>'
            f'</tr>'
        )

        _TH = "padding:7px 8px;text-align:center;font-size:10px;font-weight:700;border:1px solid #B8B8B8;background:#D9D9D9;color:#333;"
        st.markdown(f"""
<div style="background:#fff;overflow:hidden;border:1px solid #B8B8B8;">
    <div style="overflow-x:auto;">
        <table style="width:100%;border-collapse:collapse;font-size:11px;">
            <thead>
                <tr>
                    <th colspan="3" style="{_TH}text-align:left;min-width:130px;">Range architecture</th>
                    <th colspan="3" style="{_TH}">Sale Impact ( ex.vat) / Week<br>calcualte from Mer Price</th>
                    <th colspan="1" style="{_TH}">Margin Impact ( ex.vat) / Week<br>calcualte from EDLP Price</th>
                </tr>
                <tr>
                    <th style="{_TH}text-align:left;">TYPE</th>
                    <th style="{_TH}min-width:52px;">AS IS</th>
                    <th style="{_TH}min-width:52px;border-right:2px solid #999;">TO BE</th>
                    <th style="{_TH}min-width:60px;">AS IS</th>
                    <th style="{_TH}min-width:60px;">TO BE</th>
                    <th style="{_TH}min-width:52px;border-right:2px solid #999;">DIFF</th>
                    <th style="{_TH}min-width:60px;">AS IS</th>
                </tr>
            </thead>
            <tbody>{tbody}</tbody>
        </table>
    </div>
    <div style="text-align:right;padding:3px 8px;font-size:9px;color:#888;
                border-top:1px solid #E0E0E0;background:#F8F8F8;">
        display mgr = Avg selling price from format
    </div>
</div>""", unsafe_allow_html=True)

    with _leg_c:
        st.empty()

    st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)

    # ── Controls ──────────────────────────────────────────────────────────────
    _ctrl1, _ctrl2, _ctrl3, _ctrl4 = st.columns([2, 3, 1, 1.2])
    with _ctrl1:
        _filter_mode = st.selectbox("Filter", ["ALL", "SSPOG", "NON-SSPOG"],
                                    label_visibility="collapsed", key="vw_filter")
    with _ctrl2:
        _search_q = st.text_input("Search", placeholder="🔎  Search item, barcode, status...",
                                  label_visibility="collapsed")
    with _ctrl3:
        _n_rows = st.number_input("Rows", min_value=1, max_value=10000, value=10, step=10,
                                  help="Rows to display")
    with _ctrl4:
        _edit_on = st.toggle("⚙️ Edit Columns", key="vw_edit_cols")

    if _edit_on:
        st.markdown("""
<div style="background:#fff;border-radius:12px;border:1px solid #E8E3DC;
            padding:12px 16px 8px;margin-bottom:12px;">
    <div style="font-size:12px;font-weight:700;color:#1A1A1A;margin-bottom:8px;">
        ⚙️ Column Visibility
    </div>
</div>""", unsafe_allow_html=True)
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

    # Order always follows _all_grouped (= RS_COL_GROUPS order); vis_cols is just a visibility filter
    _vis_set  = set(st.session_state.view_vis_cols or [])
    disp_cols = (
        [c for c in _all_grouped if c in merged.columns and (not _vis_set or c in _vis_set)]
        or all_cols[:20]
    )

    _disp_groups = []
    for _g in active_groups:
        _gc = [c for c in _g["matched"] if c in disp_cols]
        if _gc:
            _disp_groups.append({**_g, "disp_cols": _gc})

    # ── Filter + search ───────────────────────────────────────────────────────
    df_view  = merged.copy()
    _pog_col = next((c for c in all_cols if "pog" in c.lower() and "cluster" in c.lower()), None)
    if _filter_mode == "SSPOG" and _pog_col:
        df_view = df_view[
            df_view[_pog_col].astype(str).str.contains("SSPOG", na=False) &
            ~df_view[_pog_col].astype(str).str.contains("Non", na=False)]
    elif _filter_mode == "NON-SSPOG" and _pog_col:
        df_view = df_view[df_view[_pog_col].astype(str).str.contains("Non-SSPOG", na=False)]
    # DG Code / DG Name filter (connected to Display Group selectboxes)
    if _sel_dg_code != "ALL" and _dg_code_col and _dg_code_col in df_view.columns:
        df_view = df_view[df_view[_dg_code_col].astype(str).str.strip() == _sel_dg_code]
    if _sel_dg_name != "ALL" and _dg_name_col and _dg_name_col in df_view.columns:
        df_view = df_view[df_view[_dg_name_col].astype(str).str.strip() == _sel_dg_name]
    if _search_q:
        _mask   = df_view.apply(
            lambda r: r.astype(str).str.contains(_search_q, case=False, na=False).any(), axis=1)
        df_view = df_view[_mask]
    df_view = df_view.reset_index(drop=True)

    # ── Metrics row ───────────────────────────────────────────────────────────
    _m_maintain = int((df_view[_stc].astype(str).str.strip() == "MAINTAIN").sum()) if _stc else 0
    _m_sspog    = int(df_view[_pog_col].astype(str).str.contains("SSPOG", na=False).sum() -
                      df_view[_pog_col].astype(str).str.contains("Non-SSPOG", na=False).sum()
                      ) if _pog_col else 0
    _m_nonsspog = int(df_view[_pog_col].astype(str).str.contains("Non-SSPOG", na=False).sum()
                      ) if _pog_col else 0
    _m_null     = int(df_view[disp_cols].isnull().sum().sum())

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total SKUs",  f"{len(df_view):,}")
    m2.metric("MAINTAIN",    f"{_m_maintain:,}")
    m3.metric("SSPOG",       f"{_m_sspog:,}")
    m4.metric("Non-SSPOG",   f"{_m_nonsspog:,}")
    m5.metric("Null Values", f"{_m_null:,}")

    # ── Sub-view selector ─────────────────────────────────────────────────────
    _subview = st.radio(
        "sub_view", ["📋 Table", "🏪 Cluster", "📊 Status"],
        horizontal=True, label_visibility="collapsed", key="vw_subview",
    )

    # ─────────────────────────────────────────────────────────────────────────
    # SUB-VIEW: Table
    # ─────────────────────────────────────────────────────────────────────────
    if _subview == "📋 Table":
        _MAX = int(_n_rows)

        # ── Hardcoded 24-column spec: (display label, [candidate data-col names])
        _VIEW_SPECS = [
            ("Department",                             ["Department", "Dept", "Department Code&Desc", "Department Code & Desc", "department_code_desc"]),
            ("Section",                                ["Section", "section", "Section Code&Desc", "Section Code & Desc"]),
            ("Subclass",                               ["Subclass", "SubClass", "Sub Class", "subclass", "Subclass Code & Desc", "Subclass Code&Desc"]),
            ("Barcode",                                ["Barcode", "barcode", "UPC", "EAN", "ean"]),
            ("TPNA",                                   ["TPNA", "tpna", "Style Number"]),
            ("ID",                                     ["ID", "id", "Item ID", "ItemID"]),
            ("No. of Unit in Case",                    ["No. of Unit in Case", "No_of_Unit_in_Case", "Units Per Case", "Case Units", "no of unit in case", "CaseTotalNumber"]),
            ("No. of Unit in Inner",                   ["No. of Unit in Inner", "No_of_Unit_in_Inner", "no of unit in inner", "InnerQty"]),
            ("Tray total number",                      ["Tray total number", "Tray_total_number", "Tray Total Number", "tray total", "TrayTotalNumber"]),
            ("Express Picking Type",                   ["Express Picking Type", "Express_Picking_Type", "express picking type", "MiniPickType"]),
            ("HDET Picking Type",                      ["HDET Picking Type", "HDET_Picking_Type", "hdet picking type", "Hyper&SuperPickingType"]),
            ("EDLP Price by Format",                   ["EDLP Price by Format", "EDLP_Price_by_Format", "edlp price by format", "EDLP Price"]),
            ("Item Name",                              ["Item Name", "Item name", "item name", "ProductDescription"]),
            ("As IS planograms applied",               ["AS IS planograms applied", "As IS planograms applied", "AS-IS planograms applied", "ASIS planograms applied", "as is planograms applied"]),
            ("To-BE planograms applied",               ["TO-BE planograms applied", "To-BE planograms applied", "TOBE planograms applied", "to be planograms applied"]),
            ("AS-IS Store applied",                    ["AS-IS Stores Applied", "AS IS Stores Applied", "ASIS Stores Applied", "AS-IS Store applied", "as-is stores applied"]),
            ("To-Be store applied",                    ["TO-Be stores applied", "To-Be stores applied", "TOBE stores applied", "to-be stores applied", "to be stores applied"]),
            ("Avg unit 52 wk/forecast new item sales", ["__avg_unit_coa__", "Avg Units 52wk/ Forecast new item sales", "Avg Units 52wk/Forecast new item sales", "avg units 52wk/ forecast new item sales", "Avg unit 52wk"]),
            ("Supplier pack size",                     ["Supplier Pack Size", "Supplier pack size", "Supplier_Pack_Size", "supplier pack size", "OriginalPackSize"]),
            ("Range Tail YYYY",                        ["Range Tail YYYY", "Range_Tail_YYYY", "range tail yyyy"]),
            ("AVG selling Price by format",            ["AVG Selling Price by Format", "Avg Selling Price by Format", "avg selling price by format", "AVG_Selling_Price_by_Format"]),
            ("Star Line",                              ["Star Line", "Star_Line", "starline", "star line", "StarLine"]),
            ("Item priority",                          ["Item Priority", "Item priority", "Item_Priority", "item priority"]),
            ("JDA vs Actual",                          ["JDA vs Actual", "JDA_vs_Actual", "jda vs actual"]),
            ("Actual-Actual",                          ["Actual-Actual", "Actual_Actual", "actual-actual", "actual actual"]),
        ]
        # Group definitions: (label, bg-color, text-color, column-count)
        _VIEW_GROUPS = [
            ("Item Info",  "#D9D9D9", "#333333", 13),
            ("Range Info", "#E8E3DC", "#444444",  8),
            ("Star Line",  "#000000", "#FFFFFF",   1),
            ("Priority",   "#00CC44", "#003300",   3),
        ]

        _ZERO_SPEC_COLS = {
            "As IS planograms applied", "To-BE planograms applied",
            "AS-IS Store applied", "To-Be store applied",
        }

        def _sel_view(src_df, specs):
            def _vnorm(s): return _re.sub(r'[^a-z0-9]', '', str(s).lower())
            result = {}
            for out_name, candidates in specs:
                src_col = None
                for cand in candidates:
                    src_col = next(
                        (c for c in src_df.columns if _vnorm(c) == _vnorm(cand)), None)
                    if src_col:
                        break
                if src_col:
                    result[out_name] = src_df[src_col].reset_index(drop=True)
                elif out_name in _ZERO_SPEC_COLS:
                    result[out_name] = pd.Series([0] * len(src_df), name=out_name)
                else:
                    result[out_name] = pd.Series([None] * len(src_df), name=out_name)
            return pd.DataFrame(result)

        # Coalesce: "Avg Units 52wk" (primary) → "ForecastSales" (fallback for new items).
        # Builds __avg_unit_coa__ so _sel_view picks it as the first candidate above.
        _vnorm_v = lambda s: _re.sub(r'[^a-z0-9]', '', str(s).lower())
        _avg_u_src_cands  = ["Avg Units 52wk/ Forecast new item sales",
                              "Avg Units 52wk/Forecast new item sales",
                              "avg units 52wk/ forecast new item sales",
                              "Avg unit 52wk", "Avg Units 52wk"]
        _fcast_src_cands  = ["ForecastSales", "Forecast new item sales",
                              "Forecast New Item Sales", "forecast new item sales"]
        _avg_u_raw = next(
            (c for c in df_view.columns
             for cand in _avg_u_src_cands if _vnorm_v(c) == _vnorm_v(cand)), None)
        _fcast_raw = next(
            (c for c in df_view.columns
             for cand in _fcast_src_cands if _vnorm_v(c) == _vnorm_v(cand)), None)
        if _avg_u_raw or _fcast_raw:
            _avg_u_s  = (pd.to_numeric(df_view[_avg_u_raw], errors="coerce")
                         if _avg_u_raw else pd.Series([float("nan")] * len(df_view), dtype="float64"))
            _fcast_s  = (pd.to_numeric(df_view[_fcast_raw], errors="coerce")
                         if _fcast_raw else pd.Series([float("nan")] * len(df_view), dtype="float64"))
            df_view = df_view.copy()
            df_view["__avg_unit_coa__"] = _avg_u_s.where(_avg_u_s.notna(), _fcast_s)

        _tdf      = _sel_view(df_view, _VIEW_SPECS).head(_MAX)
        _hdrs     = list(_tdf.columns)
        def _cw(lbl): return max(90, min(240, len(lbl) * 7 + 16))

        _h = [
            '<div style="overflow-x:auto;border-radius:12px;border:1px solid #E0D9D2;margin-top:14px;">',
            '<table style="border-collapse:collapse;font-size:11px;min-width:100%;">',
            '<thead><tr>',
        ]
        for _gname, _gcolor, _gtcolor, _gcnt in _VIEW_GROUPS:
            _h.append(
                f'<th colspan="{_gcnt}" style="padding:6px 8px;text-align:center;'
                f'background:{_gcolor};border:1px solid #E0D9D2;font-size:9px;font-weight:700;'
                f'color:{_gtcolor};letter-spacing:.05em;text-transform:uppercase;">{_gname}</th>')
        _h.append('</tr><tr style="background:#1C1C1E;">')
        for _lbl in _hdrs:
            _w = _cw(_lbl)
            _h.append(
                f'<th style="padding:9px 10px;text-align:left;font-weight:700;'
                f'color:rgba(255,255,255,.85);font-size:10px;white-space:nowrap;'
                f'min-width:{_w}px;max-width:{_w}px;border-right:1px solid rgba(255,255,255,.07);">'
                f'{_lbl}</th>')
        _h.append('</tr></thead><tbody>')

        for _i, _row in _tdf.iterrows():
            _rb = "#FFFFFF" if _i % 2 == 0 else "#F8F4F0"
            _h.append(f'<tr style="background:{_rb};">')
            for _lbl in _hdrs:
                _val = _row.get(_lbl, "")
                _sv  = "" if pd.isna(_val) or str(_val) in ("nan", "None") else str(_val)
                _w   = _cw(_lbl)
                _ln  = _lbl.lower()
                if _ln in ("barcode", "id") and _sv:
                    _inner = (f'<span style="color:#2BBFA4;font-family:monospace;'
                              f'font-weight:600;">{_sv}</span>')
                elif _lbl == "Avg unit 52 wk/forecast new item sales" and _sv:
                    try:
                        # Round to 5dp then strip trailing zeros (and bare ".").
                        # Avoids %.5f padding (1.5 → "1.50000") and %g scientific
                        # notation on large values (e.g. 123456.0 → "1.23456e+05").
                        _sv = ("%.5f" % round(float(_sv), 5)).rstrip("0").rstrip(".")
                    except (ValueError, TypeError):
                        _sv = ""
                    _inner = f'<span>{_sv}</span>'
                else:
                    _inner = f'<span>{_sv}</span>'
                _h.append(
                    f'<td style="padding:8px 10px;border-bottom:1px solid #E0D9D2;'
                    f'border-right:1px solid #E0D9D2;min-width:{_w}px;max-width:{_w}px;'
                    f'overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">{_inner}</td>')
            _h.append('</tr>')
        _h.append('</tbody></table></div>')
        st.markdown(''.join(_h), unsafe_allow_html=True)

        # Row-count helper hidden to keep the review page clean.

        _fc1, _fc2, _fc3 = st.columns([3, 1, 1])
        with _fc1:
            pass
        with _fc2:
            st.download_button("⬇️ Export .xlsx",
                df_to_xlsx_bytes(_tdf),
                file_name=f"view_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        with _fc3:
            st.download_button("⬇️ Raw .csv", df_to_csv_bytes(merged),
                file_name=f"raw_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv", use_container_width=True)

    # ─────────────────────────────────────────────────────────────────────────
    # SUB-VIEW: Cluster
    # ─────────────────────────────────────────────────────────────────────────
    elif _subview == "🏪 Cluster":
        # Row definitions: (label, bg-color, text-color)
        _DEFAULT_METRICS = [
            ("MAX Stores applied count",         "#F4A460", "#000000"),
            ("count of stores in store cluster", "#F4A460", "#000000"),
            ("MODs",                             "#DCDCDC", "#000000"),
            ("FIXTURE",                          "#DCDCDC", "#000000"),
            ("New Framework",                    "#DCDCDC", "#000000"),
            ("Total NEW SKUs",                   "#808080", "#FFFFFF"),
            ("Total DELETE SKUs",                "#808080", "#FFFFFF"),
            ("TO-BE SKUs count",                 "#808080", "#FFFFFF"),
            ("AS-IS SKUs count",                 "#F0F0F0", "#1565C0"),
            ("%Achieving LRD CASE (As Is)",      "#FFFFFF", "#000000"),
            ("%Achieving LRD SALES (As Is)",     "#FFFFFF", "#000000"),
        ]
        _DEF_ROW_COLORS = {r: (bg, tc) for r, bg, tc in _DEFAULT_METRICS}

        if "ct_clusters"  not in st.session_state:
            st.session_state.ct_clusters  = []
        if "ct_data"      not in st.session_state:
            st.session_state.ct_data      = {r: {} for r, _, _ in _DEFAULT_METRICS}
        if "ct_row_clrs"  not in st.session_state:
            st.session_state.ct_row_clrs  = dict(_DEF_ROW_COLORS)

        # ── Toolbar ───────────────────────────────────────────────────────────
        _tb1, _tb2, _tb3, _tb4, _tb5 = st.columns([3.2, 1.1, 1.0, 0.9, 0.8])
        with _tb1:
            _ct_inp = st.text_input(
                "_ct", key="ct_inp",
                placeholder="Type name → Add as Column or Row…",
                label_visibility="collapsed",
            )
        with _tb2:
            if st.button("＋ Column", key="ct_add_col", use_container_width=True):
                _v = st.session_state.get("ct_inp", "").strip()
                if _v and _v not in st.session_state.ct_clusters:
                    st.session_state.ct_clusters.append(_v)
                    for _rk in st.session_state.ct_data:
                        st.session_state.ct_data[_rk].setdefault(_v, "")
                st.session_state["ct_inp"] = ""
                st.rerun()
        with _tb3:
            if st.button("＋ Row", key="ct_add_row", use_container_width=True):
                _v = st.session_state.get("ct_inp", "").strip() or f"Metric {len(st.session_state.ct_data)+1}"
                if _v not in st.session_state.ct_data:
                    st.session_state.ct_data[_v] = {c: "" for c in st.session_state.ct_clusters}
                    st.session_state.ct_row_clrs[_v] = ("#FFFFFF", "#000000")
                st.session_state["ct_inp"] = ""
                st.rerun()
        with _tb4:
            if st.button("↺ Reset", key="ct_reset", use_container_width=True):
                st.session_state.ct_clusters = []
                st.session_state.ct_data     = {r: {} for r, _, _ in _DEFAULT_METRICS}
                st.session_state.ct_row_clrs = dict(_DEF_ROW_COLORS)
                st.rerun()
        with _tb5:
            _edit_mode = st.toggle("✏️ Edit", key="ct_edit")

        _clusters = st.session_state.ct_clusters
        _all_rows = list(st.session_state.ct_data.keys())

        if not _edit_mode:
            # ── Display: styled HTML table matching the Excel design ──────────
            _ph = [
                '<div style="overflow-x:auto;border-radius:10px;border:1px solid #CCC;margin-top:12px;">',
                '<table style="border-collapse:collapse;font-size:12px;min-width:100%;">',
                '<thead><tr>',
                '<th style="background:#FFD700;color:#000;font-weight:800;padding:14px 18px;'
                'border:1px solid #BBB;text-align:left;min-width:230px;font-size:13px;">'
                'POG Cluster</th>',
            ]
            if _clusters:
                for _c in _clusters:
                    _ph.append(
                        f'<th style="background:#FFD700;color:#000;font-weight:700;'
                        f'padding:10px 14px;border:1px solid #BBB;text-align:center;'
                        f'min-width:100px;white-space:nowrap;">{_c}</th>'
                    )
            else:
                _ph.append(
                    '<th style="background:#FFD700;color:#888;padding:10px 14px;'
                    'border:1px solid #BBB;font-size:11px;font-style:italic;min-width:260px;">'
                    'Add cluster columns using the input above ↑</th>'
                )
            _ph.append('</tr></thead><tbody>')

            for _rname in _all_rows:
                _rbg, _rtc = st.session_state.ct_row_clrs.get(_rname, ("#FFFFFF", "#000000"))
                _ph.append(
                    f'<tr><td style="background:{_rbg};color:{_rtc};font-weight:600;'
                    f'padding:8px 18px;border:1px solid #CCC;white-space:nowrap;">{_rname}</td>'
                )
                if _clusters:
                    for _c in _clusters:
                        _v = st.session_state.ct_data.get(_rname, {}).get(_c, "")
                        _v = "" if _v is None or str(_v) in ("nan", "None") else str(_v)
                        _ph.append(
                            f'<td style="background:{_rbg};color:{_rtc};padding:8px 12px;'
                            f'border:1px solid #CCC;text-align:center;">{_v}</td>'
                        )
                else:
                    _ph.append(f'<td style="background:{_rbg};padding:8px 12px;border:1px solid #CCC;"></td>')
                _ph.append('</tr>')

            _ph.append('</tbody></table></div>')
            st.markdown(''.join(_ph), unsafe_allow_html=True)

        else:
            # ── Edit mode: data_editor (cells + add/delete rows) ──────────────
            _edit_rows = [
                {"POG Cluster": r,
                 **{c: st.session_state.ct_data.get(r, {}).get(c, "") for c in _clusters}}
                for r in _all_rows
            ]
            _edit_df = pd.DataFrame(_edit_rows) if _edit_rows else pd.DataFrame(columns=["POG Cluster"] + _clusters)

            _col_cfg = {
                "POG Cluster": st.column_config.TextColumn("POG Cluster", width="large"),
                **{c: st.column_config.TextColumn(c, width="small") for c in _clusters},
            }
            _edited = st.data_editor(
                _edit_df,
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                column_config=_col_cfg,
                key="ct_editor",
            )
            # Persist edits to session state
            _new_data = {}
            for _, _er in _edited.iterrows():
                _rn = str(_er.get("POG Cluster", "")).strip()
                if _rn and _rn not in ("nan", "None", ""):
                    _new_data[_rn] = {c: _er.get(c, "") for c in _clusters}
                    if _rn not in st.session_state.ct_row_clrs:
                        st.session_state.ct_row_clrs[_rn] = ("#FFFFFF", "#000000")
            st.session_state.ct_data = _new_data

        # ── Export ────────────────────────────────────────────────────────────
        if _all_rows:
            _exp_df = pd.DataFrame([
                {"POG Cluster": r,
                 **{c: st.session_state.ct_data.get(r, {}).get(c, "") for c in _clusters}}
                for r in _all_rows
            ])
            st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
            st.download_button(
                "⬇️ Export Cluster Table .xlsx",
                df_to_xlsx_bytes(_exp_df),
                file_name=f"cluster_table_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ─────────────────────────────────────────────────────────────────────────
    # SUB-VIEW: Status
    # ─────────────────────────────────────────────────────────────────────────
    elif _subview == "📊 Status":
        if _stc:
            _stat_counts = (
                df_view[_stc].astype(str).str.strip()
                .value_counts().reset_index()
            )
            _stat_counts.columns = ["Status", "Count"]
            _stat_counts = _stat_counts[~_stat_counts["Status"].isin(["nan", ""])]
            _total_st = _stat_counts["Count"].sum()

            _sh = [
                '<div style="overflow-x:auto;border-radius:12px;border:1px solid #E0D9D2;margin-top:14px;">',
                '<table style="border-collapse:collapse;font-size:12px;width:100%;">',
                '<thead><tr style="background:#1C1C1E;">',
                '<th style="padding:10px 16px;text-align:left;color:rgba(255,255,255,.75);'
                'font-size:11px;font-weight:700;">STATUS</th>',
                '<th style="padding:10px 16px;text-align:center;color:rgba(255,255,255,.75);'
                'font-size:11px;font-weight:700;">COUNT</th>',
                '<th style="padding:10px 16px;text-align:center;color:rgba(255,255,255,.75);'
                'font-size:11px;font-weight:700;">%</th>',
                '<th style="padding:10px 16px;color:rgba(255,255,255,.75);'
                'font-size:11px;font-weight:700;">BAR</th>',
                '</tr></thead><tbody>',
            ]
            for _si, _sr in _stat_counts.iterrows():
                _sval = str(_sr["Status"])
                _sc2  = STATUS_COLORS.get(_sval, {"bg": "#F5F5F5", "c": "#888"})
                _pct  = _sr["Count"] / _total_st * 100 if _total_st else 0
                _rb2  = "#FFFFFF" if _si % 2 == 0 else "#F8F4F0"
                _sh.append(
                    f'<tr style="background:{_rb2};">'
                    f'<td style="padding:10px 16px;border-bottom:1px solid #F0EBE3;">'
                    f'<span style="background:{_sc2["bg"]};color:{_sc2["c"]};padding:3px 10px;'
                    f'border-radius:4px;font-weight:700;font-size:11px;">{_sval}</span></td>'
                    f'<td style="padding:10px 16px;text-align:center;border-bottom:1px solid #F0EBE3;'
                    f'font-weight:700;font-size:13px;">{_sr["Count"]:,}</td>'
                    f'<td style="padding:10px 16px;text-align:center;border-bottom:1px solid #F0EBE3;'
                    f'color:#888;font-size:12px;">{_pct:.1f}%</td>'
                    f'<td style="padding:10px 16px;border-bottom:1px solid #F0EBE3;">'
                    f'<div style="background:#EDE8DF;border-radius:99px;height:6px;">'
                    f'<div style="background:{_sc2["c"]};border-radius:99px;height:6px;'
                    f'width:{min(_pct,100):.1f}%;"></div></div></td>'
                    f'</tr>'
                )
            _sh.append('</tbody></table></div>')
            st.markdown(''.join(_sh), unsafe_allow_html=True)
            st.caption(f"{len(_stat_counts)} statuses · {_total_st:,} total SKUs")
            st.download_button(
                "⬇️ Export Status Summary",
                df_to_xlsx_bytes(_stat_counts),
                file_name=f"status_summary_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("Status column not found in the data.")

render_page_nav("rangesheetreview")
