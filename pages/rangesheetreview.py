"""Rangesheet review page — sheet tabs, Range Architecture card, data table."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


import streamlit as st
import pandas as pd
import re as _re
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

    _subview = st.radio("sub_view",["📋 Table","🏪 Cluster"],
                        horizontal=True, label_visibility="collapsed", key=f"{p}_subview")

    _tab_all, _tab_cnv = st.tabs(["📊  All Data", "🎨  New Canvas"])

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
                        _dg_combo_val  = st.selectbox(
                            "Search DG Code", _dg_combo_opts,
                            key=f"{p}_dg_code", label_visibility="visible",
                        )
                        _sel_dg_code = "" if _dg_combo_val == _ALL_OPT else _dg_combo_val
                    else:
                        _sel_dg_code = st.text_input(
                            "Search DG Code", placeholder="Type DG code to filter…",
                            key=f"{p}_dg_code", label_visibility="visible")
                    if _dg_code_col and _dg_name_col and _sel_dg_code:
                        _ns = df_src[df_src[_dg_code_col].astype(str).str.contains(_sel_dg_code, case=False, na=False)]
                        _dg_name_opts = ["ALL"] + sorted(_ns[_dg_name_col].dropna().astype(str).str.strip().unique().tolist())
                    elif _dg_name_col:
                        _dg_name_opts = ["ALL"] + sorted(df_src[_dg_name_col].dropna().astype(str).str.strip().unique().tolist())
                    else:
                        _dg_name_opts = ["ALL"]
                    _prev_name = st.session_state.get(f"_{p}_dg_name_prev", "ALL")
                    if _prev_name not in _dg_name_opts:
                        _prev_name = "ALL"
                    _sel_dg_name = st.selectbox(
                        "DG NAME", _dg_name_opts,
                        index=(_dg_name_opts.index(_prev_name) if _prev_name in _dg_name_opts else 0),
                        key=f"{p}_dg_name", label_visibility="visible")
                    st.session_state[f"_{p}_dg_name_prev"] = _sel_dg_name
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
            _ASIS_T = {"MAINTAIN","DELETE SOME","DELETE ALL"} | ({"NEW DELETE SOME"} if _is_sspog else set())
            _TOBE_T = {"MAINTAIN","DELETE SOME","NEW SOME","NEWNEW"} | ({"NEW DELETE SOME"} if _is_sspog else set())
            def _fmt(v):
                if v == 0: return "0"
                try: return f"{int(v):,}" if v == int(v) else f"{v:,.1f}"
                except: return str(v)
            _arch_rows = []
            for _t in TYPES:
                _msk = (_ab[_type_col].astype(str).str.strip() == _t) if (_type_col and _type_col in _ab.columns) else pd.Series([False]*len(_ab), index=_ab.index)
                _idx = _ab.index[_msk]
                _ai = len(_idx) if _t in _ASIS_T else 0
                _tb = len(_idx) if _t in _TOBE_T else 0
                _prices = pd.to_numeric(_ab.loc[_idx, _price_col], errors="coerce").fillna(0) if (_price_col and len(_idx) > 0) else pd.Series([], dtype=float)
                _asis_s = pd.to_numeric(_ab.loc[_idx, _asis_stc], errors="coerce").fillna(0) if (_asis_stc and len(_idx) > 0) else pd.Series([1]*len(_idx), dtype=float)
                _tobe_s = pd.to_numeric(_ab.loc[_idx, _tobe_stc], errors="coerce").fillna(0) if (_tobe_stc and len(_idx) > 0) else pd.Series([1]*len(_idx), dtype=float)
                _edlp_p = pd.to_numeric(_ab.loc[_idx, _edlp_col], errors="coerce").fillna(0) if (_edlp_col and len(_idx) > 0) else pd.Series([], dtype=float)
                _arch_rows.append({"type":_t,"as_is":_ai,"to_be":_tb,
                    "sale_ai":float((_prices*_asis_s).sum()) if len(_prices) else 0.0,
                    "sale_tb":float((_prices*_tobe_s).sum()) if len(_prices) else 0.0,
                    "sale_diff":float((_prices*_tobe_s).sum()-(_prices*_asis_s).sum()) if len(_prices) else 0.0,
                    "marg_ai":float((_edlp_p*_asis_s).sum()) if len(_edlp_p) else 0.0})
            _t_ai=sum(r["as_is"] for r in _arch_rows); _t_tb=sum(r["to_be"] for r in _arch_rows)
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
                    _ai_bg="background:#E53935;color:#fff;" if _t=="NEWNEW" else ""
                    _tb_bg="background:#E53935;color:#fff;" if _t=="DELETE ALL" else ""
                    tbody+=(f'<tr style="border-bottom:1px solid #D8D8D8;">'
                        f'<td style="padding:5px 10px;font-size:11px;color:#1A1A1A;{_B}">{_t}</td>'
                        f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}{_ai_bg}">{_fmt(_r["as_is"])}</td>'
                        f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}{_tb_bg}">{_fmt(_r["to_be"])}</td>'
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
                    _ai_bg="background:#E53935;color:#fff;" if _t=="NEWNEW" else ""
                    _tb_bg="background:#E53935;color:#fff;" if _t=="DELETE ALL" else ""
                    tbody+=(f'<tr style="border-bottom:1px solid #D8D8D8;">'
                        f'<td style="padding:5px 10px;font-size:11px;color:#1A1A1A;{_B}">{_t}</td>'
                        f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}{_ai_bg}">{_fmt(_r["as_is"])}</td>'
                        f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}{_BR}{_tb_bg}">{_fmt(_r["to_be"])}</td>'
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
            st.markdown("""<div style="background:#fff;border-radius:14px;border:1px solid #E8E3DC;padding:16px;">
              <div style="font-size:10px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.08em;margin-bottom:12px;">Color Note</div>
              <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px;">
                <div style="width:22px;height:14px;background:#FFFDE7;border:1px solid #E0D9D2;border-radius:3px;flex-shrink:0;"></div>
                <span style="font-size:11px;color:#555;">Display fill</span></div>
              <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px;">
                <div style="width:22px;height:14px;background:#FCE4EC;border:1px solid #E0D9D2;border-radius:3px;flex-shrink:0;"></div>
                <span style="font-size:11px;color:#555;">Merchandiser fill</span></div>
              <div style="display:flex;align-items:center;gap:8px;">
                <div style="width:22px;height:14px;background:#F5F5F5;border:1px solid #E0D9D2;border-radius:3px;flex-shrink:0;"></div>
                <span style="font-size:11px;color:#555;">Formula</span></div></div>""", unsafe_allow_html=True)

        st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)

        _c2, _c3, _c4 = st.columns([4, 1, 1.2])
        with _c2:
            _search_q = st.text_input("Search", placeholder="🔎  Search item, barcode, status...",
                                      label_visibility="collapsed", key=f"{p}_search")
        with _c3:
            _n_rows = st.number_input("Rows", min_value=10, max_value=100000, value=500, step=100,
                                      help="Max rows to display", key=f"{p}_nrows")
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

        if _sel_dg_code and large_file_path:
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
            df_view = st.session_state[_dk].copy()
            st.caption(f"🔍 Found **{len(df_view):,} rows** matching DG = '{_q}' (full file search)")
        else:
            df_view = df_src.copy()
            if _sel_dg_code:
                _q = _sel_dg_code.strip()
                if _dg_code_col and _dg_code_col in df_view.columns:
                    df_view = df_view[
                        df_view[_dg_code_col].astype(str).str.strip()
                        .str.contains(_q, case=False, na=False)
                    ]
                else:
                    df_view = df_view[df_view.apply(
                        lambda r: r.astype(str).str.contains(_q, case=False, na=False).any(), axis=1)]

        if _sel_dg_name != "ALL" and _dg_name_col and _dg_name_col in df_view.columns:
            df_view = df_view[df_view[_dg_name_col].astype(str).str.strip() == _sel_dg_name]
        if _search_q:
            df_view = df_view[df_view.apply(
                lambda r: r.astype(str).str.contains(_search_q, case=False, na=False).any(), axis=1)]
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
                st.caption(f"Planogram cells → `{_sv_col}` (Value column ✓)")

            _dyn_pog_cols = []
            _piv_pog_cols = []   # set inside pivot block; guards outer debug expanders
            _pk_std       = []
            _pk_raw       = []
            _pk_map       = {}
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

                if _dyn_pog_cols and _pk_raw:
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
                    _tdf = _tdf_dedup.merge(_piv_std, on=_pk_std, how="left")

                    # Cap at _MAX products (not long rows) and reset index for clean row numbers
                    _tdf = _tdf.head(_MAX).reset_index(drop=True)

                    # Check Range To-be Waterfall = number of planograms this product appears in
                    _tdf["Check Range To-be Waterfall"] = _tdf[_piv_pog_cols].notna().sum(axis=1)

                    # ── TEMPORARY DEBUG EXPANDER ─────────────────────────────────
                    with st.expander("🔎 debug — pivot diagnostics", expanded=False):

                        # ── 1. Column matching ──────────────────────────────────
                        st.markdown("**1 · Column matching**")
                        st.write({
                            "_pog_src_col (planogram name col)": _pog_src_col,
                            "_sv_col (sales VALUE col)":         _sv_col,
                            "_pk_map (std → raw key cols)":      _pk_map,
                            "_dyn_pog_cols (planogram labels)":  _dyn_pog_cols,
                        })

                        # ── 2. Pivot integrity ──────────────────────────────────
                        st.markdown("**2 · Pivot integrity**")
                        _n_unique_products = df_view[_pk_raw].drop_duplicates().shape[0]
                        st.write({
                            "df_view shape (long rows × cols)":     df_view.shape,
                            "unique products in df_view":            _n_unique_products,
                            "_piv shape (products × pog+key cols)": _piv.shape,
                        })
                        st.markdown("Pivot columns:")
                        st.write(list(_piv.columns))
                        st.markdown("_dyn_pog_cols (expected headers):")
                        st.write(_dyn_pog_cols)
                        st.markdown("_piv_pog_cols (intersection — values that actually landed in pivot):")
                        st.write(_piv_pog_cols)
                        _missing_in_pivot = [c for c in _dyn_pog_cols if c not in set(_piv.columns)]
                        if _missing_in_pivot:
                            st.warning(f"⚠️ These _dyn_pog_cols are NOT in the pivot columns "
                                       f"(name mismatch?): {_missing_in_pivot}")

                        # ── 3. Merge check ──────────────────────────────────────
                        st.markdown("**3 · Merge check (join key alignment)**")
                        # Rows where ALL pog columns are NaN → no pivot match
                        if _piv_pog_cols:
                            _blank_mask = _tdf[_piv_pog_cols].isna().all(axis=1)
                            _n_blank = int(_blank_mask.sum())
                            st.write(f"Products with ALL pog columns blank after merge: {_n_blank} / {len(_tdf)}")
                            if _n_blank > 0:
                                _blank_keys = _tdf.loc[_blank_mask, _pk_std].head(5)
                                st.markdown("Sample blank product keys in `_tdf` (after merge):")
                                st.dataframe(_blank_keys)
                                # Pivot index keys — compare dtypes / repr
                                _piv_keys_sample = _piv_std[_pk_std].head(5)
                                st.markdown("Sample pivot key rows in `_piv_std` (should match above):")
                                st.dataframe(_piv_keys_sample)
                                # Dtype comparison
                                _dtype_tdf = {c: str(_tdf[c].dtype) for c in _pk_std}
                                _dtype_piv = {c: str(_piv_std[c].dtype) for c in _pk_std}
                                st.write("_tdf key dtypes:", _dtype_tdf)
                                st.write("_piv_std key dtypes:", _dtype_piv)

                        # ── 4. Raw source check ─────────────────────────────────
                        st.markdown("**4 · Raw source check — df_view rows for blank products**")
                        if _piv_pog_cols:
                            _blank_mask2 = _tdf[_piv_pog_cols].isna().all(axis=1)
                            _sample_blank_stds = _tdf.loc[_blank_mask2, _pk_std].head(3)
                            if not _sample_blank_stds.empty:
                                for _, _brow in _sample_blank_stds.iterrows():
                                    # Build a filter against df_view using raw pk col names
                                    _filt = pd.Series([True] * len(df_view), index=df_view.index)
                                    for _s, _r in _pk_map.items():
                                        _filt &= (df_view[_r].astype(str).str.strip() == str(_brow[_s]).strip())
                                    _raw_rows = df_view.loc[_filt, _pk_raw + [_pog_src_col] + ([_sv_col] if _sv_col else [])]
                                    st.markdown(f"Raw rows for key `{tuple(_brow[c] for c in _pk_std)}`:")
                                    if _raw_rows.empty:
                                        st.warning("No matching rows found in df_view — key mismatch in the filter itself.")
                                    else:
                                        st.dataframe(_raw_rows)
                            else:
                                st.success("No blank-product rows found — all products matched the pivot.")

                        # ── 5. Pre-render cell comparison ───────────────────────
                        st.markdown("**5 · Pre-render cell comparison (pre-overlay, pre-reorder)**")
                        st.write(f"_MAX = {_MAX} | len(_tdf) = {len(_tdf)}")

                        # Any overlay data already in session state?
                        _dbg_ov_key  = f"{p}_pog_edits"
                        _dbg_ov_data = st.session_state.get(_dbg_ov_key, {})
                        st.write(f"_pog_edits entries in session state: {len(_dbg_ov_data)}")
                        if _dbg_ov_data:
                            st.write("Sample overlay entries (first 3):",
                                     dict(list(_dbg_ov_data.items())[:3]))

                        # Auto-pick 3 products that have the most non-NaN values in _piv
                        _dbg_nonnull_counts = _piv[_piv_pog_cols].notna().sum(axis=1)
                        _dbg_top_idx = _dbg_nonnull_counts.nlargest(3).index.tolist()

                        for _dbg_pi in _dbg_top_idx:
                            _dbg_prow  = _piv.loc[_dbg_pi]
                            _dbg_kstd  = {s: _dbg_prow[_pk_map[s]] for s in _pk_std}

                            # Locate product in _tdf by matching key values as strings
                            _dbg_mask = pd.Series([True] * len(_tdf), index=_tdf.index)
                            for _dbg_s in _pk_std:
                                _dbg_mask &= (
                                    _tdf[_dbg_s].astype(str).str.strip()
                                    == str(_dbg_kstd[_dbg_s]).strip()
                                )
                            _dbg_tmatch = _tdf[_dbg_mask]

                            _dbg_cap_flag = (
                                "no" if _dbg_tmatch.empty
                                else ("YES ⚠️" if _dbg_tmatch.index.min() >= _MAX else "no")
                            )
                            st.markdown(
                                f"**Product key:** `{_dbg_kstd}` — "
                                f"_tdf row(s): `{_dbg_tmatch.index.tolist()}` "
                                f"(row index ≥ _MAX={_MAX}? {_dbg_cap_flag})"
                            )

                            _dbg_comp = []
                            for _dbg_pc in _piv_pog_cols[:40]:
                                _piv_val = _dbg_prow.get(_dbg_pc, float("nan"))
                                if _dbg_tmatch.empty:
                                    _tdf_val  = "PRODUCT NOT IN _tdf"
                                    _fmt      = ""
                                elif _dbg_pc not in _dbg_tmatch.columns:
                                    _tdf_val  = "COLUMN MISSING"
                                    _fmt      = ""
                                else:
                                    _tdf_val = _dbg_tmatch.iloc[0][_dbg_pc]
                                    try:
                                        _fmt = (f"{int(float(_tdf_val)):,}"
                                                if pd.notna(_tdf_val) else "(blank)")
                                    except (TypeError, ValueError):
                                        _fmt = f"(err: {_tdf_val!r})"
                                _dbg_comp.append({
                                    "planogram col": _dbg_pc,
                                    "_piv value":    _piv_val,
                                    "_tdf value":    _tdf_val,
                                    "type(_tdf)":    type(_tdf_val).__name__,
                                    "pd.notna":      pd.notna(_tdf_val)
                                                     if not isinstance(_tdf_val, str)
                                                     else True,
                                    "render output": _fmt,
                                })
                            st.dataframe(pd.DataFrame(_dbg_comp), height=300, use_container_width=True)
                    # ── END TEMPORARY DEBUG EXPANDER ─────────────────────────────

                else:
                    _tdf["Check Range To-be Waterfall"] = ""

            # Column ordering: sticky left | data cols | Status | planogram cols rightmost
            _STICKY      = [c for c in ["DG Code", "ID", "Item Name"] if c in _tdf.columns]
            _dyn_pog_set = set(_dyn_pog_cols)
            _LAST        = [c for c in ["Status", "Check Range To-be Waterfall", "Planogram Name"] if c in _tdf.columns] + _dyn_pog_cols
            _REST        = [c for c in _tdf.columns if c not in _STICKY and c not in set(_LAST)]
            _tdf         = _tdf[_STICKY + _REST + _LAST]

            # ── Persistent planogram-edit overlay ────────────────────────────────
            # Key = (DG Code, ID, Item Name) — after pivot, one row per product so
            # the sticky columns alone are unique.
            _pog_edits_key = f"{p}_pog_edits"

            def _make_rk(ri: int) -> tuple:
                # After pivot _tdf has one row per product — _STICKY alone is unique
                return tuple(str(_tdf.at[ri, c]) for c in _STICKY)

            _pog_edits = st.session_state.get(_pog_edits_key, {})
            if _pog_edits and _dyn_pog_cols and _STICKY:
                for _ri in range(len(_tdf)):
                    _rk = _make_rk(_ri)
                    if _rk in _pog_edits:
                        for _pc, _pv in _pog_edits[_rk].items():
                            if _pc in _tdf.columns:
                                _tdf.at[_ri, _pc] = _pv

            # ── TEMPORARY: post-overlay debug ────────────────────────────────────
            if _dyn_pog_cols:
                with st.expander("🔎 debug — post-overlay / render-ready state", expanded=False):
                    st.markdown("**Values in `_tdf` AFTER overlay applied, AFTER column reorder — "
                                "this is exactly what the render loop sees.**")
                    # Pick first 3 rows of _tdf; show all pog columns
                    _pog_in_tdf = [c for c in _dyn_pog_cols if c in _tdf.columns]
                    st.write(f"_dyn_pog_cols count: {len(_dyn_pog_cols)} | "
                             f"pog cols present in _tdf: {len(_pog_in_tdf)} | "
                             f"overlay entries: {len(_pog_edits)}")

                    for _dbg_ri in range(min(3, len(_tdf))):
                        _dbg_rk = _make_rk(_dbg_ri)
                        _ov_for_row = _pog_edits.get(_dbg_rk, {})
                        st.markdown(f"**Row {_dbg_ri}** | key={_dbg_rk} | "
                                    f"overlay entries for this key: {_ov_for_row}")
                        _post_rows = []
                        for _pc in _pog_in_tdf[:40]:
                            _raw = _tdf.at[_dbg_ri, _pc]
                            try:
                                _fmt = (f"{int(float(_raw)):,}"
                                        if pd.notna(_raw) else "(blank)")
                            except (TypeError, ValueError):
                                _fmt = f"(err: {_raw!r})"
                            _post_rows.append({
                                "planogram col": _pc,
                                "raw value":     _raw,
                                "type":          type(_raw).__name__,
                                "pd.notna":      pd.notna(_raw)
                                                 if not isinstance(_raw, str) else True,
                                "render output": _fmt,
                                "wiped by overlay?": (
                                    "YES ⚠️" if _pc in _ov_for_row
                                    and pd.isna(_ov_for_row[_pc]) else "no"
                                ),
                            })
                        st.dataframe(pd.DataFrame(_post_rows), height=300, use_container_width=True)
            # ── END TEMPORARY post-overlay debug ─────────────────────────────────

            # ── TEMPORARY: automated reconciliation (proves pivot ≡ raw data) ──
            if _piv_pog_cols and _sv_col and _pog_src_col:
                with st.expander("🔍 reconciliation — pivot vs raw data", expanded=False):
                    _SEP = "\x00\x01"   # separator unlikely to appear in product IDs

                    # ── Pre-compute shared key series (used by all 4 checks) ──
                    # pk string key in df_view (raw col names)
                    _rc_dfv_pk = df_view[_pk_raw[0]].astype(str)
                    for _c in _pk_raw[1:]:
                        _rc_dfv_pk = _rc_dfv_pk + _SEP + df_view[_c].astype(str)

                    # pk string key in _tdf (standard col names, same join order)
                    _rc_tdf_pk = _tdf[_pk_std[0]].astype(str)
                    for _c in _pk_std[1:]:
                        _rc_tdf_pk = _rc_tdf_pk + _SEP + _tdf[_c].astype(str)

                    _rc_displayed_pks = set(_rc_tdf_pk)                          # products in table
                    _rc_pog_str       = df_view[_pog_src_col].astype(str).str.strip()
                    _rc_sv_num        = pd.to_numeric(df_view[_sv_col], errors="coerce")
                    _rc_pk_mask       = _rc_dfv_pk.isin(_rc_displayed_pks)        # rows for displayed products
                    _rc_pog_mask      = _rc_pog_str.isin(set(_dyn_pog_cols))      # rows for displayed planograms

                    # ── Check 1: Grand total ────────────────────────────────────
                    st.markdown("#### Check 1 — Grand total")
                    _c1_tdf = float(_tdf[_piv_pog_cols].sum(skipna=True).sum())
                    _c1_raw = float(_rc_sv_num[_rc_pk_mask & _rc_pog_mask].sum())
                    _c1_diff = abs(_c1_tdf - _c1_raw)
                    _c1_ok   = _c1_diff < 0.5
                    st.write({
                        "Pivot _tdf total (all pog cols, all displayed products)":
                            f"{_c1_tdf:,.2f}",
                        "Raw df_view total (same products + planograms)":
                            f"{_c1_raw:,.2f}",
                        "Absolute difference":
                            f"{_c1_diff:.4f}",
                    })
                    if _c1_ok:
                        st.success("✓ Grand totals match")
                    else:
                        st.error(f"✗ Grand total mismatch — diff = {_c1_diff:,.2f}")

                    # ── Check 2: Cell spot-check (8 random non-NaN cells) ───────
                    st.markdown("#### Check 2 — Cell spot-check (8 random cells)")
                    _c2_melt = (
                        _tdf[_pk_std + _piv_pog_cols]
                        .melt(id_vars=_pk_std, var_name="_pog_", value_name="_val_")
                        .dropna(subset=["_val_"])
                        .reset_index(drop=True)
                    )
                    _c2_sample = (
                        _c2_melt.sample(min(8, len(_c2_melt)), random_state=42)
                        if not _c2_melt.empty else _c2_melt
                    )
                    _c2_rows = []
                    for _, _sr in _c2_sample.iterrows():
                        _pog_lbl  = _sr["_pog_"]
                        _tdf_val  = float(_sr["_val_"])
                        _pk_key   = _SEP.join(str(_sr[c]) for c in _pk_std)
                        _c2_filt  = (_rc_dfv_pk == _pk_key) & (_rc_pog_str == _pog_lbl)
                        _raw_sum  = float(_rc_sv_num[_c2_filt].sum())
                        _c2_rows.append({
                            **{c: _sr[c] for c in _pk_std},
                            "planogram":  _pog_lbl,
                            "raw_sum":    f"{_raw_sum:,.0f}",
                            "tdf_value":  f"{_tdf_val:,.0f}",
                            "match?":     "✓" if abs(_raw_sum - _tdf_val) < 0.5 else "✗",
                        })
                    _c2_df = pd.DataFrame(_c2_rows)
                    st.dataframe(_c2_df, hide_index=True, use_container_width=True)
                    if _c2_df.empty or (_c2_df["match?"] == "✓").all():
                        st.success("✓ All sampled cells match")
                    else:
                        _c2_bad = int((~(_c2_df["match?"] == "✓")).sum())
                        st.error(f"✗ {_c2_bad} cell(s) mismatch")

                    # ── Check 3: Marginals (5 products, row-total) ───────────────
                    st.markdown("#### Check 3 — Marginals (5 sample products, row-total)")
                    _c3_sample = _tdf.sample(min(5, len(_tdf)), random_state=7)
                    _c3_rows = []
                    for _, _mr in _c3_sample.iterrows():
                        _pivot_tot = sum(
                            float(_mr[c]) for c in _piv_pog_cols
                            if c in _mr.index and pd.notna(_mr[c])
                        )
                        _pk_key  = _SEP.join(str(_mr[c]) for c in _pk_std)
                        _c3_filt = (_rc_dfv_pk == _pk_key) & _rc_pog_mask
                        _raw_tot = float(_rc_sv_num[_c3_filt].sum())
                        _c3_rows.append({
                            **{c: _mr[c] for c in _pk_std},
                            "tdf row-total":    f"{_pivot_tot:,.0f}",
                            "raw df_view sum":  f"{_raw_tot:,.0f}",
                            "match?": "✓" if abs(_pivot_tot - _raw_tot) < 0.5 else "✗",
                        })
                    _c3_df = pd.DataFrame(_c3_rows)
                    st.dataframe(_c3_df, hide_index=True, use_container_width=True)
                    if _c3_df.empty or (_c3_df["match?"] == "✓").all():
                        st.success("✓ All row-totals match")
                    else:
                        st.error(f"✗ {int((~(_c3_df['match?'] == '✓')).sum())} row-total(s) mismatch")

                    # ── Check 4: aggfunc sanity ─────────────────────────────────
                    st.markdown("#### Check 4 — aggfunc sanity (repeated product×planogram pairs?)")
                    _c4_df = (
                        df_view[_rc_pk_mask & _rc_pog_mask]
                        .assign(_pk_=_rc_dfv_pk, _pog_=_rc_pog_str)
                        .groupby(["_pk_", "_pog_"], sort=False)
                        .size()
                        .reset_index(name="_count_")
                    )
                    _c4_repeated = _c4_df[_c4_df["_count_"] > 1]
                    st.write(
                        f"Total (product, planogram) pairs in displayed subset: "
                        f"**{len(_c4_df):,}** | "
                        f"Pairs appearing >1 time: **{len(_c4_repeated):,}**"
                    )
                    if _c4_repeated.empty:
                        st.success("✓ All pairs unique — aggfunc='sum' ≡ first (no multi-row summing)")
                    else:
                        st.warning(
                            f"⚠️ {len(_c4_repeated):,} pairs repeat. "
                            f"aggfunc='sum' is adding their rows. "
                            f"Confirm this is intended (e.g. multi-store rows)."
                        )
                        for _, _rr in _c4_repeated.head(3).iterrows():
                            _c4_filt = (_rc_dfv_pk == _rr["_pk_"]) & (_rc_pog_str == _rr["_pog_"])
                            _c4_detail = df_view.loc[_c4_filt, _pk_raw + [_pog_src_col, _sv_col]].copy()
                            _c4_detail["→ summed_to"] = (
                                pd.to_numeric(_c4_detail[_sv_col], errors="coerce").sum()
                            )
                            st.markdown(
                                f"**pk=`{_rr['_pk_']}`  pog=`{_rr['_pog_']}`** "
                                f"({int(_rr['_count_'])} rows)"
                            )
                            st.dataframe(_c4_detail, hide_index=True)
            # ── END TEMPORARY reconciliation ──────────────────────────────────

            # ── Edit-mode toggle ──────────────────────────────────────────────────
            _em_state  = f"{p}_tbl_edit_mode"
            _em_widget = f"{p}_tbl_edit_mode_w"
            _em_c1, _em_c2 = st.columns([9, 1])
            with _em_c2:
                _edit_mode = st.toggle("✏️ Edit", key=_em_widget,
                                       value=st.session_state.get(_em_state, False),
                                       help="Toggle on to edit planogram assignments. Click Done to commit.")
            st.session_state[_em_state] = _edit_mode

            _COL_W  = {"DG Code": 56, "ID": 90, "Item Name": 210}

            if not _edit_mode:
                # ── VIEW MODE: original HTML table, unchanged ─────────────────────
                _s_left, _s_lefts = 0, {}
                for _sc in _STICKY:
                    _s_lefts[_sc] = _s_left
                    _s_left += _COL_W.get(_sc, 130)
                _HDR_BG  = "#D9D9D9"
                _STK_BG  = "#F7F5F2"
                _CELL_H  = "padding:5px 10px;border:1px solid #E0D9D2;font-size:11px;white-space:nowrap;"
                _col_list = list(_tdf.columns)
                # Pre-format Avg Units column before bulk .astype(str) stringifies raw floats.
                # Must happen here — .astype(str) on line below would turn 1.23456789 → "1.23456789"
                # before any per-column formatting could run.
                _avg_u_std = "Avg Units 52wk/ Forecast new item sales"
                if _avg_u_std in _tdf.columns:
                    def _fmt_avg_u(v):
                        if pd.isna(v) or str(v).strip() in ("", "nan", "None"):
                            return ""
                        try:
                            return ("%.5f" % round(float(v), 5)).rstrip("0").rstrip(".")
                        except (ValueError, TypeError):
                            return "" if str(v).strip() in ("nan", "None") else str(v)
                    _tdf = _tdf.copy()
                    _tdf[_avg_u_std] = _tdf[_avg_u_std].apply(_fmt_avg_u)
                _arr      = _tdf.fillna("").astype(str).values
                _th_list, _rows = [], []
                for _ci, _col in enumerate(_col_list):
                    if _col in _STICKY:
                        _lx = _s_lefts[_col]
                        _w  = _COL_W.get(_col, 130)
                        if _col == "DG Code":
                            # DG Code is narrow (80px) — let header wrap so the column stays compact
                            _th_list.append(
                                f'<th style="position:sticky;left:{_lx}px;z-index:3;background:{_HDR_BG};'
                                f'border:1px solid #E0D9D2;font-size:10px;font-weight:700;'
                                f'width:{_w}px;min-width:{_w}px;max-width:{_w}px;'
                                f'padding:4px 4px;vertical-align:bottom;'
                                f'white-space:normal;word-break:break-word;overflow-wrap:break-word;">'
                                f'{_col}</th>')
                        else:
                            _th_list.append(
                                f'<th style="position:sticky;left:{_lx}px;z-index:3;background:{_HDR_BG};'
                                f'{_CELL_H}font-weight:700;min-width:{_w}px;">{_col}</th>')
                    elif _col in _dyn_pog_set:
                        _th_list.append(
                            f'<th style="background:{_HDR_BG};border:1px solid #E0D9D2;'
                            f'width:52px;min-width:52px;max-width:52px;height:260px;'
                            f'padding:6px 2px;text-align:center;vertical-align:bottom;'
                            f'writing-mode:vertical-rl;text-orientation:mixed;'
                            f'font-size:11px;font-weight:700;white-space:nowrap;overflow:hidden;">'
                            f'{_col}</th>')
                    else:
                        # _REST data columns: wrapped header, narrow fixed width
                        _th_list.append(
                            f'<th style="background:{_HDR_BG};border:1px solid #E0D9D2;'
                            f'font-size:10px;font-weight:700;'
                            f'width:80px;min-width:40px;max-width:80px;'
                            f'padding:4px 4px;vertical-align:bottom;'
                            f'white-space:normal;word-break:break-word;overflow-wrap:break-word;">'
                            f'{_col}</th>')
                for _ri in range(len(_tdf)):
                    _row_h = []
                    _rb = "#FFFFFF" if _ri % 2 == 0 else "#FAFAF8"
                    for _ci, _col in enumerate(_col_list):
                        _vs = _arr[_ri, _ci]
                        if _vs in ("nan", "None"):
                            _vs = ""
                        if _col in _STICKY:
                            _lx = _s_lefts[_col]
                            _row_h.append(
                                f'<td style="position:sticky;left:{_lx}px;z-index:2;background:{_STK_BG};'
                                f'{_CELL_H}">{_vs}</td>')
                        elif _col in _dyn_pog_set:
                            _raw = _tdf.at[_ri, _col]
                            try:
                                _vf = f"{float(_raw):,.2f}" if pd.notna(_raw) else ""
                            except (ValueError, TypeError):
                                _vf = ""
                            _row_h.append(
                                f'<td style="background:{_rb};border:1px solid #E0D9D2;'
                                f'width:52px;min-width:52px;text-align:right;'
                                f'font-size:11px;padding:2px 3px;">{_vf}</td>')
                        else:
                            # _REST data cells: match header width, clip overflow
                            _row_h.append(
                                f'<td style="background:{_rb};border:1px solid #E0D9D2;'
                                f'font-size:11px;padding:5px 4px;'
                                f'max-width:80px;white-space:nowrap;'
                                f'overflow:hidden;text-overflow:ellipsis;">{_vs}</td>')
                    _rows.append(f'<tr>{"".join(_row_h)}</tr>')
                st.markdown(
                    '<div style="overflow-x:auto;border-radius:10px;border:1px solid #E0D9D2;'
                    'max-height:560px;overflow-y:auto;">'
                    '<table style="border-collapse:collapse;font-size:11px;min-width:100%;">'
                    f'<thead style="position:sticky;top:0;z-index:4;"><tr>{"".join(_th_list)}</tr></thead>'
                    f'<tbody>{"".join(_rows)}</tbody>'
                    '</table></div>',
                    unsafe_allow_html=True,
                )

            else:
                # ── EDIT MODE: st.data_editor ─────────────────────────────────────
                # _tdf already has committed edits from _pog_edits_key applied above,
                # so it is the correct baseline for the editor.
                _de_key = f"{p}_de_{_sel_dg_code}"

                # Planogram cols: editable numeric. Sticky cols: default width, disabled.
                # REST data cols: width="small" to match narrow view-mode headers.
                # (st.data_editor has no multi-line header support — narrowing is the
                # closest equivalent to the wrapped headers in view mode.)
                _de_col_cfg = {}
                for _c in _tdf.columns:
                    if _c in _dyn_pog_cols:
                        _de_col_cfg[_c] = st.column_config.NumberColumn(
                            _c, format="%.2f", step=0.01)
                    elif _c == "DG Code":
                        _de_col_cfg[_c] = st.column_config.Column(_c, disabled=True, width="small")
                    elif _c in _STICKY:
                        _de_col_cfg[_c] = st.column_config.Column(_c, disabled=True)
                    else:
                        _de_col_cfg[_c] = st.column_config.Column(
                            _c, disabled=True, width="small")

                _edited = st.data_editor(
                    _tdf,
                    key=_de_key,
                    column_config=_de_col_cfg,
                    hide_index=True,
                    height=560,
                    use_container_width=True,
                    num_rows="fixed",
                )

                # Debug: show exactly what the editor recorded as changed
                st.write("editor delta:", st.session_state.get(_de_key, {}))

                st.caption("Check boxes to edit planogram assignments. Click **Done** to save, or **Discard** to cancel.")
                _sv_c1, _sv_c2, _sv_c3 = st.columns([6, 1, 1])
                with _sv_c2:
                    if st.button("✓ Done", key=f"{p}_tbl_save", use_container_width=True, type="primary"):
                        # ── Commit only the cells the user actually changed ────────
                        # st.data_editor stores exactly what changed at
                        # st.session_state[_de_key]["edited_rows"]:
                        #   {row_position: {col_name: new_value}}
                        # We iterate only those entries — never touch other pog cols.
                        _delta = (st.session_state.get(_de_key) or {}).get("edited_rows", {})
                        _pog_edits_commit = st.session_state.setdefault(_pog_edits_key, {})
                        _audit_lines = []
                        if _STICKY and _delta:
                            for _ri_str, _cell_changes in _delta.items():
                                _ri = int(_ri_str)
                                if _ri >= len(_tdf):
                                    continue
                                _rk = _make_rk(_ri)   # includes original pog name → unique per row
                                for _pc, _new_val in _cell_changes.items():
                                    if _pc not in _dyn_pog_cols:
                                        continue
                                    _old_raw = _tdf.at[_ri, _pc]
                                    # Normalise both sides to float/NaN for comparison
                                    try:
                                        _old_num = float(_old_raw)
                                    except (TypeError, ValueError):
                                        _old_num = float("nan")
                                    try:
                                        _new_num = float(_new_val) if _new_val not in (None, "", pd.NA) else float("nan")
                                    except (TypeError, ValueError):
                                        _new_num = float("nan")
                                    _old_blank = pd.isna(_old_num)
                                    _new_blank = pd.isna(_new_num)
                                    _changed = (
                                        _old_blank != _new_blank
                                        or (not _old_blank and not _new_blank and _old_num != _new_num)
                                    )
                                    if _changed:
                                        _store = None if _new_blank else _new_num
                                        _pog_edits_commit.setdefault(_rk, {})[_pc] = _store
                                        _old_fmt = "blank" if _old_blank else f"{int(_old_num):,}"
                                        _new_fmt = "blank" if _new_blank else f"{int(_new_num):,}"
                                        _audit_lines.append(f"{_rk}|{_pc}: {_old_fmt}→{_new_fmt}")
                        if _audit_lines:
                            add_audit(
                                "Edit Planogram Assignments",
                                f"tab={p}; {len(_audit_lines)} change(s): "
                                + "; ".join(_audit_lines[:10])
                                + ("…" if len(_audit_lines) > 10 else ""),
                            )
                        st.session_state.pop(_de_key, None)
                        st.session_state[_em_state] = False
                        st.rerun()
                with _sv_c3:
                    if st.button("✕ Discard", key=f"{p}_tbl_cancel", use_container_width=True):
                        st.session_state.pop(_de_key, None)
                        st.session_state[_em_state] = False
                        st.rerun()

            if len(df_view) > _MAX:
                st.caption(f"Showing {_MAX:,} of {len(df_view):,} rows — increase Rows to see more")
            else:
                st.caption(f"{len(df_view):,} rows · {len(_tdf.columns)} columns")

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
    with _tab_cnv:
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
    if st.button("✅ SUBMIT RANGE", key=f"{p}_btn_sub", type="primary", use_container_width=True):
        st.session_state[f"vw_submit_{p}"] = df_view.reset_index(drop=True)
        st.success("Data submitted to Report page!")
    if st.session_state.get(f"vw_submit_{p}") is not None:
        st.caption(f"✅ {len(st.session_state[f'vw_submit_{p}']):,} rows submitted to Report")

# Pre-filter SSPOG rows for tabs 1 & 2
_pog_col_main = next((c for c in all_cols if "pog" in c.lower() and "cluster" in c.lower()), None)
_sspog_df = (
    merged[
        merged[_pog_col_main].astype(str).str.contains("SSPOG", na=False) &
        ~merged[_pog_col_main].astype(str).str.contains("Non", na=False)
    ] if _pog_col_main else merged
)

# ── Sheet tabs (scrollable via st.tabs) ───────────────────────────────────────
_sheet_tabs = st.tabs(RS_SHEETS)


with _sheet_tabs[0]:   # Range Sheet_Non-SSPOG
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

        # Status bar
        _info_parts = [f"HDET: <strong>{_ns_hfname}</strong>"]
        if _ns_src_lbl:
            _info_parts.append(_ns_src_lbl)
        if _ns_a5_name:
            _info_parts.append(f"A5: <strong>{_ns_a5_name}</strong>")
        st.markdown(
            "<div style='font-size:11px;color:#2BBFA4;margin-bottom:6px;'>"
            + " · ".join(_info_parts) + "</div>",
            unsafe_allow_html=True,
        )

        # Step 1: Convert HDET CSV → Parquet once (makes all future ops much faster)
        if "ns_pq_ready" not in st.session_state:
            with st.status("Preparing HDET for fast access…", expanded=True) as _st:
                _pq = ensure_hdet_parquet(_ns_hdet_path, status_cb=_st.write)
                if _pq:
                    _st.write("✅ Parquet ready — future loads will be instant.")
                else:
                    _st.write("⚠️ Parquet conversion unavailable (pyarrow not installed); using CSV.")
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

        # DG Code + DG Name dropdowns
        _dg_c1, _dg_c2, _dg_c3 = st.columns([2, 2, 1])
        _ALL = "— All (preview) —"
        with _dg_c1:
            _ns_sel_code = st.selectbox(
                "DG Code", [_ALL] + _ns_codes,
                key="ns_sel_dg_code", label_visibility="visible")
        with _dg_c2:
            _ns_sel_name = st.selectbox(
                "DG Name", [_ALL] + _ns_names,
                key="ns_sel_dg_name", label_visibility="visible")
        with _dg_c3:
            st.write("")
            _ns_load_btn = st.button("Load", key="ns_load_dg2",
                                     use_container_width=True, type="primary")
            if st.button("↺ Reset", key="ns_reset_dg2", use_container_width=True):
                for _k in ["ns_hdet_df", "_ns_hdet_src", "ns_lf_dg_sig",
                           "_ns_loaded_dg_code", "_ns_loaded_dg_name"]:
                    st.session_state.pop(_k, None)
                st.rerun()

        # Resolve: if DG Name picked, map to code
        _ns_dg_query = None
        if _ns_sel_code != _ALL:
            _ns_dg_query = _ns_sel_code
        elif _ns_sel_name != _ALL:
            _ns_dg_query = _ns_name_to_code.get(_ns_sel_name, _ns_sel_name)

        if _ns_load_btn and _ns_dg_query:
            _sig = f"{_ns_hdet_path}|{_ns_dg_query}"
            if st.session_state.get("ns_lf_dg_sig") != _sig:
                with st.spinner(f"Loading DG '{_ns_dg_query}' from HDET…"):
                    _loaded = load_large_file_by_dg(_ns_hdet_path, _ns_dg_query)
                st.session_state["ns_hdet_df"]   = _loaded
                st.session_state["_ns_hdet_src"] = f"DG = {_ns_dg_query} ({len(_loaded):,} rows)"
                st.session_state["ns_lf_dg_sig"] = _sig
                st.session_state["_ns_loaded_dg_code"] = _ns_dg_query
                st.session_state["_ns_loaded_dg_name"] = _ns_code_to_name.get(_ns_dg_query, "")
                st.rerun()

    # Require a DG to be picked and loaded before showing any data
    _ns_hdet_df = st.session_state.get("ns_hdet_df")
    if _ns_hdet_path and _ns_hdet_df is None:
        st.info("👆 Select a DG Code or DG Name above, then click **Load** to view planogram data.")
    else:
        _ns_combined = _dedup(_ns_hdet_df.copy()) if _ns_hdet_df is not None else merged
        _ns_idx = st.session_state.get("ns_dg_index", {})
        _render_sheet_content(_ns_combined, "ns",
                              dg_col_hint=_ns_idx.get("dg_col"),
                              large_file_path=_ns_hdet_path,
                              dg_options=_ns_idx.get("codes", []),
                              fixed_dg_code=st.session_state.get("_ns_loaded_dg_code"),
                              fixed_dg_name=st.session_state.get("_ns_loaded_dg_name"))

with _sheet_tabs[1]:   # Range Sheet_SSPOG
    _render_sheet_content(_sspog_df, "ss")
with _sheet_tabs[2]:   # StoreApply_SSPOG
    _render_sheet_content(_sspog_df, "sa")
with _sheet_tabs[3]:   # 5.1 ItembyStore
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
with _sheet_tabs[4]:   # 5.2 ItembyStore_SC
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
with _sheet_tabs[5]:   # 5.3 Upload_product_library
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
with _sheet_tabs[6]:   # 5.4 Upload to Citrix
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
                st.button("SUBMIT RANGE", key="btn_submit_range",
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
        # AS IS types = items that existed before (in AS-IS range)
        _ASIS_TYPES = {"MAINTAIN", "NEW DELETE SOME", "DELETE SOME", "DELETE ALL"}
        # TO BE types = items that will be in TO-BE range
        _TOBE_TYPES = {"MAINTAIN", "NEW DELETE SOME", "DELETE SOME", "NEW SOME", "NEWNEW"}

        def _fmt(v):
            if v == 0: return "0"
            try: return f"{int(v):,}" if v == int(v) else f"{v:,.1f}"
            except: return str(v)

        _arch_rows = []
        for _t in TYPES:
            if _type_col and _type_col in _arch_base.columns:
                _msk = _arch_base[_type_col].astype(str).str.strip() == _t
            else:
                _msk = pd.Series([False] * len(_arch_base), index=_arch_base.index)
            _idx = _arch_base.index[_msk]

            _ai = len(_idx) if _t in _ASIS_TYPES else 0
            _tb = len(_idx) if _t in _TOBE_TYPES else 0

            # Sale Impact AS IS / TO BE from Mer Price (AVG Selling Price × stores)
            _prices = pd.to_numeric(_arch_base.loc[_idx, _price_col], errors="coerce").fillna(0) if (_price_col and len(_idx) > 0) else pd.Series([], dtype=float)
            _asis_s = pd.to_numeric(_arch_base.loc[_idx, _asis_stc], errors="coerce").fillna(0) if (_asis_stc and len(_idx) > 0) else pd.Series([1]*len(_idx), dtype=float)
            _tobe_s = pd.to_numeric(_arch_base.loc[_idx, _tobe_stc], errors="coerce").fillna(0) if (_tobe_stc and len(_idx) > 0) else pd.Series([1]*len(_idx), dtype=float)
            _sale_ai = float((_prices * _asis_s).sum()) if len(_prices) else 0.0
            _sale_tb = float((_prices * _tobe_s).sum()) if len(_prices) else 0.0

            # Margin Impact AS IS from EDLP Price × stores
            _edlp_p = pd.to_numeric(_arch_base.loc[_idx, _edlp_col], errors="coerce").fillna(0) if (_edlp_col and len(_idx) > 0) else pd.Series([], dtype=float)
            _marg_ai = float((_edlp_p * _asis_s).sum()) if len(_edlp_p) else 0.0

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
            _ai_bg = "background:#E53935;color:#fff;" if _t == "NEWNEW"     else ""
            _tb_bg = "background:#E53935;color:#fff;" if _t == "DELETE ALL" else ""
            tbody += (
                f'<tr style="border-bottom:1px solid #D8D8D8;">'
                f'<td style="padding:5px 10px;font-size:11px;color:#1A1A1A;{_B}">{_t}</td>'
                f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}{_ai_bg}">{_fmt(_r["as_is"])}</td>'
                f'<td style="padding:5px 8px;text-align:center;font-size:11px;{_B}{_BR}{_tb_bg}">{_fmt(_r["to_be"])}</td>'
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

    # ── Color Note card ───────────────────────────────────────────────────────
    with _leg_c:
        st.markdown("""
<div style="background:#fff;border-radius:14px;border:1px solid #E8E3DC;padding:16px;">
    <div style="font-size:10px;font-weight:700;color:#888;text-transform:uppercase;
                letter-spacing:.08em;margin-bottom:12px;">Color Note</div>
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px;">
        <div style="width:22px;height:14px;background:#FFFDE7;border:1px solid #E0D9D2;
                    border-radius:3px;flex-shrink:0;"></div>
        <span style="font-size:11px;color:#555;">Display fill</span>
    </div>
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px;">
        <div style="width:22px;height:14px;background:#FCE4EC;border:1px solid #E0D9D2;
                    border-radius:3px;flex-shrink:0;"></div>
        <span style="font-size:11px;color:#555;">Merchandiser fill</span>
    </div>
    <div style="display:flex;align-items:center;gap:8px;">
        <div style="width:22px;height:14px;background:#F5F5F5;border:1px solid #E0D9D2;
                    border-radius:3px;flex-shrink:0;"></div>
        <span style="font-size:11px;color:#555;">Formula</span>
    </div>
</div>""", unsafe_allow_html=True)

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

        if len(df_view) > _MAX:
            st.caption(f"Showing {_MAX:,} of {len(df_view):,} rows — increase Rows input to see more")

        _fc1, _fc2, _fc3 = st.columns([3, 1, 1])
        with _fc1:
            st.caption(f"{len(df_view):,} rows · 24 columns · raw: {len(merged):,} rows")
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
