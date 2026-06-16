"""View Data page — sheet tabs, Range Architecture card, data table."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
import pandas as pd
import re as _re
from datetime import datetime
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar, render_page_nav,
    RS_SHEETS, RS_COL_GROUPS, STATUS_COLORS, FILL_COLORS, COLUMN_LABELS,
    get_fill, df_to_xlsx_bytes, df_to_csv_bytes, add_audit,
)

inject_css()
init_session_state()
render_sidebar("viewdata")
render_topbar("Rangesheet Review")

merged = st.session_state.merged_df
if merged is None:
    st.markdown("""
<div style="background:#fff;border-radius:16px;border:1px solid #E8E3DC;
            padding:60px;text-align:center;">
    <div style="font-size:48px;margin-bottom:16px;">📂</div>
    <div style="font-size:16px;font-weight:600;color:#1A1A1A;margin-bottom:8px;">No data yet</div>
    <div style="font-size:13px;color:#999;">Go to My Files and upload a file first.</div>
</div>
""", unsafe_allow_html=True)
    st.stop()

all_cols = list(merged.columns)

# ── Column group matching ──────────────────────────────────────────────────────
def _nc(s):  return str(s).lower().strip().replace('\n', ' ').replace('  ', ' ')
def _nca(s): return _re.sub(r'[^a-z0-9]', '', _nc(s))   # letters+digits only

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
_remaining   = [c for c in all_cols if c not in _all_grouped]
if _remaining:
    active_groups.append({"group": "Other", "color": "#FAFAFA", "matched": _remaining})

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

def _render_sheet_content(df_src, p):
    """Full All Data + New Canvas tab pair, keyed with prefix p."""
    _src_cols = list(df_src.columns)
    _meta = st.session_state.rangesheet_meta

    # column-group helpers scoped to df_src (columns are same as merged)
    _sag = []
    for _grp in RS_COL_GROUPS:
        _gm = [_fc(k) for k in _grp["cols"] if _fc(k) is not None]
        if _gm:
            _sag.append({**_grp, "matched": _gm})
    _sag_flat = [c for g in _sag for c in g["matched"]]
    _sag_rest = [c for c in _src_cols if c not in _sag_flat]
    if _sag_rest:
        _sag.append({"group": "Other", "color": "#FAFAFA", "matched": _sag_rest})

    _subview = st.radio("sub_view",["📋 Table","🏪 Cluster","📊 Status"],
                        horizontal=True, label_visibility="collapsed", key=f"{p}_subview")

    _tab_all, _tab_cnv = st.tabs(["📊  All Data", "🎨  New Canvas"])

    # ── ALL DATA ──────────────────────────────────────────────────────────────
    with _tab_all:
        _dg_code_col = (
            next((c for c in _src_cols if _nc(c) in ("dg code", "dg_code", "dg")), None)
            or next((c for c in _src_cols if "dg" in _nc(c) and "code" in _nc(c)), None)
            or next((c for c in _src_cols if "department" in _nc(c)), None)
        )
        _dg_name_col = (
            next((c for c in _src_cols if _nc(c) in ("dg name", "dg_name")), None)
            or next((c for c in _src_cols if "dg" in _nc(c) and "name" in _nc(c)), None)
            or next((c for c in _src_cols if "section" in _nc(c)), None)
        )
        _dg_code_opts = (
            ["ALL"] + sorted(df_src[_dg_code_col].dropna().astype(str).str.strip().unique().tolist())
            if _dg_code_col else ["ALL"]
        )

        _dg_c, _arch_c, _leg_c = st.columns([1.05, 2.9, 0.75])
        _sel_dg_code = "ALL"
        _sel_dg_name = "ALL"

        with _dg_c:
            st.markdown("""<div style="background:#fff;border-radius:14px;border:1px solid #E8E3DC;
                        padding:14px 14px 0 14px;">
            <div style="font-size:10px;font-weight:700;color:#888;text-transform:uppercase;
                        letter-spacing:.08em;margin-bottom:10px;">Display Group</div></div>""",
                unsafe_allow_html=True)
            with st.container():
                _prev_code = st.session_state.get(f"_{p}_dg_code_prev", "ALL")
                _sel_dg_code = st.selectbox(
                    "DG CODE", _dg_code_opts,
                    index=(_dg_code_opts.index(_prev_code) if _prev_code in _dg_code_opts else 0),
                    key=f"{p}_dg_code", label_visibility="visible")
                if _dg_code_col and _dg_name_col and _sel_dg_code != "ALL":
                    _ns = df_src[df_src[_dg_code_col].astype(str).str.strip() == _sel_dg_code]
                    _dg_name_opts = ["ALL"] + sorted(_ns[_dg_name_col].dropna().astype(str).str.strip().unique().tolist())
                elif _dg_name_col:
                    _dg_name_opts = ["ALL"] + sorted(df_src[_dg_name_col].dropna().astype(str).str.strip().unique().tolist())
                else:
                    _dg_name_opts = ["ALL"]
                _prev_name = st.session_state.get(f"_{p}_dg_name_prev", "ALL")
                if _sel_dg_code != _prev_code and _prev_name not in _dg_name_opts:
                    _prev_name = "ALL"
                _sel_dg_name = st.selectbox(
                    "DG NAME", _dg_name_opts,
                    index=(_dg_name_opts.index(_prev_name) if _prev_name in _dg_name_opts else 0),
                    key=f"{p}_dg_name", label_visibility="visible")
                st.session_state[f"_{p}_dg_code_prev"] = _sel_dg_code
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
            if _sel_dg_code != "ALL" and _dg_code_col and _dg_code_col in _ab.columns:
                _ab = _ab[_ab[_dg_code_col].astype(str).str.strip() == _sel_dg_code]
            if _sel_dg_name != "ALL" and _dg_name_col and _dg_name_col in _ab.columns:
                _ab = _ab[_ab[_dg_name_col].astype(str).str.strip() == _sel_dg_name]
            TYPES = ["MAINTAIN","NEW DELETE SOME","DELETE SOME","DELETE ALL","NEW SOME","NEWNEW"]
            _ASIS_T = {"MAINTAIN","NEW DELETE SOME","DELETE SOME","DELETE ALL"}
            _TOBE_T = {"MAINTAIN","NEW DELETE SOME","DELETE SOME","NEW SOME","NEWNEW"}
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
            _TH="padding:7px 8px;text-align:center;font-size:10px;font-weight:700;border:1px solid #B8B8B8;background:#D9D9D9;color:#333;"
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
            _n_rows = st.number_input("Rows", min_value=1, max_value=10000, value=10, step=10,
                                      help="Rows to display", key=f"{p}_nrows")
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
        disp_cols = ([c for c in _sag_flat if c in df_src.columns and (not _vis_set or c in _vis_set)] or _src_cols[:20])

        # Filter + search
        df_view = df_src.copy()
        _pog_c = next((c for c in _src_cols if "pog" in c.lower() and "cluster" in c.lower()), None)
        if _sel_dg_code != "ALL" and _dg_code_col and _dg_code_col in df_view.columns:
            df_view = df_view[df_view[_dg_code_col].astype(str).str.strip() == _sel_dg_code]
        if _sel_dg_name != "ALL" and _dg_name_col and _dg_name_col in df_view.columns:
            df_view = df_view[df_view[_dg_name_col].astype(str).str.strip() == _sel_dg_name]
        if _search_q:
            df_view = df_view[df_view.apply(lambda r: r.astype(str).str.contains(_search_q, case=False, na=False).any(), axis=1)]
        df_view = df_view.reset_index(drop=True)

        _m_maintain = int((df_view[_stc].astype(str).str.strip()=="MAINTAIN").sum()) if _stc else 0
        _m_sspog    = int(df_view[_pog_c].astype(str).str.contains("SSPOG",na=False).sum() - df_view[_pog_c].astype(str).str.contains("Non-SSPOG",na=False).sum()) if _pog_c else 0
        _m_nonsspog = int(df_view[_pog_c].astype(str).str.contains("Non-SSPOG",na=False).sum()) if _pog_c else 0
        _m_null     = int(df_view[disp_cols].isnull().sum().sum())
        m1,m2,m3,m4,m5 = st.columns(5)
        m1.metric("Total SKUs",f"{len(df_view):,}"); m2.metric("MAINTAIN",f"{_m_maintain:,}")
        m3.metric("SSPOG",f"{_m_sspog:,}"); m4.metric("Non-SSPOG",f"{_m_nonsspog:,}"); m5.metric("Null Values",f"{_m_null:,}")

        # ── Table ─────────────────────────────────────────────────────────────
        if _subview == "📋 Table":
            _MAX = int(_n_rows)
            _VS = [
                ("Department",["Department","Dept","Department Code&Desc","department_code_desc"]),
                ("Section",["Section","section"]),
                ("Subclass",["Subclass","SubClass","Sub Class","subclass"]),
                ("Barcode",["Barcode","barcode","UPC","EAN","ean"]),
                ("TPNA",["TPNA","tpna"]),
                ("ID",["ID","id","Item ID","ItemID"]),
                ("No. of Unit in Case",["No. of Unit in Case","No_of_Unit_in_Case","Units Per Case","Case Units","no of unit in case"]),
                ("No. of Unit in Inner",["No. of Unit in Inner","No_of_Unit_in_Inner","no of unit in inner"]),
                ("Tray total number",["Tray total number","Tray_total_number","Tray Total Number","tray total"]),
                ("Express Picking Type",["Express Picking Type","Express_Picking_Type","express picking type"]),
                ("HDET Picking Type",["HDET Picking Type","HDET_Picking_Type","hdet picking type"]),
                ("EDLP Price by Format Item name",["EDLP Price by Format","EDLP_Price_by_Format","Item Name","Item name","edlp price by format"]),
                ("As IS planograms applied",["AS IS planograms applied","As IS planograms applied","AS-IS planograms applied","ASIS planograms applied","as is planograms applied"]),
                ("To-BE planograms applied",["TO-BE planograms applied","To-BE planograms applied","TOBE planograms applied","to be planograms applied"]),
                ("AS-IS Store applied",["AS-IS Stores Applied","AS IS Stores Applied","ASIS Stores Applied","AS-IS Store applied","as-is stores applied"]),
                ("To-Be store applied",["TO-Be stores applied","To-Be stores applied","TOBE stores applied","to-be stores applied","to be stores applied"]),
                ("Avg unit 52 wk/forecast new item sales",["Avg Units 52wk/ Forecast new item sales","Avg Units 52wk/Forecast new item sales","avg units 52wk/ forecast new item sales","Avg unit 52wk"]),
                ("Supplier pack size",["Supplier Pack Size","Supplier pack size","Supplier_Pack_Size","supplier pack size"]),
                ("Range Tail YYYY",["Range Tail YYYY","Range_Tail_YYYY","range tail yyyy"]),
                ("AVG selling Price by format",["AVG Selling Price by Format","Avg Selling Price by Format","avg selling price by format","AVG_Selling_Price_by_Format"]),
                ("Star Line",["Star Line","Star_Line","starline","star line"]),
                ("Item priority",["Item Priority","Item priority","Item_Priority","item priority"]),
                ("JDA vs Actual",["JDA vs Actual","JDA_vs_Actual","jda vs actual"]),
                ("Actual-Actual",["Actual-Actual","Actual_Actual","actual-actual","actual actual"]),
            ]
            _VG = [("Item Info","#D9D9D9","#333333",12),("Range Info","#E8E3DC","#444444",8),
                   ("Star Line","#000000","#FFFFFF",1),("Priority","#00CC44","#003300",3)]
            def _sv(src_df, specs):
                def _vn(s): return _re.sub(r'[^a-z0-9]','',str(s).lower())
                res={}
                for on,cands in specs:
                    sc=None
                    for cand in cands:
                        sc=next((c for c in src_df.columns if _vn(c)==_vn(cand)),None)
                        if sc: break
                    res[on]=(src_df[sc].reset_index(drop=True) if sc else pd.Series([None]*len(src_df),name=on))
                return pd.DataFrame(res)
            _tdf=_sv(df_view,_VS).head(_MAX); _hdrs=list(_tdf.columns)
            def _cw(lbl): return max(90,min(240,len(lbl)*7+16))
            _h=['<div style="overflow-x:auto;border-radius:12px;border:1px solid #E0D9D2;margin-top:14px;">',
                '<table style="border-collapse:collapse;font-size:11px;min-width:100%;"><thead><tr>']
            for _gn,_gc,_gt,_gcnt in _VG:
                _h.append(f'<th colspan="{_gcnt}" style="padding:6px 8px;text-align:center;background:{_gc};'
                          f'border:1px solid #E0D9D2;font-size:9px;font-weight:700;color:{_gt};'
                          f'letter-spacing:.05em;text-transform:uppercase;">{_gn}</th>')
            _h.append('</tr><tr style="background:#1C1C1E;">')
            for _lbl in _hdrs:
                _w=_cw(_lbl)
                _h.append(f'<th style="padding:9px 10px;text-align:left;font-weight:700;'
                          f'color:rgba(255,255,255,.85);font-size:10px;white-space:nowrap;'
                          f'min-width:{_w}px;max-width:{_w}px;border-right:1px solid rgba(255,255,255,.07);">{_lbl}</th>')
            _h.append('</tr></thead><tbody>')
            for _i,_row in _tdf.iterrows():
                _rb="#FFFFFF" if _i%2==0 else "#F8F4F0"
                _h.append(f'<tr style="background:{_rb};">')
                for _lbl in _hdrs:
                    _val=_row.get(_lbl,""); _sv2="" if pd.isna(_val) or str(_val) in ("nan","None") else str(_val)
                    _w=_cw(_lbl)
                    _inner=(f'<span style="color:#2BBFA4;font-family:monospace;font-weight:600;">{_sv2}</span>'
                            if _lbl.lower() in ("barcode","id") and _sv2 else f'<span>{_sv2}</span>')
                    _h.append(f'<td style="padding:8px 10px;border-bottom:1px solid #E0D9D2;'
                              f'border-right:1px solid #E0D9D2;min-width:{_w}px;max-width:{_w}px;'
                              f'overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">{_inner}</td>')
                _h.append('</tr>')
            _h.append('</tbody></table></div>')
            st.markdown(''.join(_h),unsafe_allow_html=True)

        # ── Cluster ───────────────────────────────────────────────────────────
        elif _subview == "🏪 Cluster":
            _DM=[("MAX Stores applied count","#F4A460","#000000"),
                 ("count of stores in store cluster","#F4A460","#000000"),
                 ("MODs","#DCDCDC","#000000"),("FIXTURE","#DCDCDC","#000000"),
                 ("New Framework","#DCDCDC","#000000"),
                 ("Total NEW SKUs","#808080","#FFFFFF"),("Total DELETE SKUs","#808080","#FFFFFF"),
                 ("TO-BE SKUs count","#808080","#FFFFFF"),
                 ("AS-IS SKUs count","#F0F0F0","#1565C0"),
                 ("%Achieving LRD CASE (As Is)","#FFFFFF","#000000"),
                 ("%Achieving LRD SALES (As Is)","#FFFFFF","#000000")]
            _DRC={r:(bg,tc) for r,bg,tc in _DM}
            _kc=f"{p}_ct_cls"; _kd=f"{p}_ct_dat"; _kr=f"{p}_ct_clr"
            if _kc not in st.session_state: st.session_state[_kc]=[]
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
                    st.session_state[_kc]=[]; st.session_state[_kd]={r:{} for r,_,_ in _DM}
                    st.session_state[_kr]=dict(_DRC); st.rerun()
            with _tb5: _em=st.toggle("✏️ Edit",key=f"{p}_ct_edit")
            _clsx=st.session_state[_kc]; _rowsx=list(st.session_state[_kd].keys())
            if not _em:
                _ph=['<div style="overflow-x:auto;border-radius:10px;border:1px solid #CCC;margin-top:12px;">',
                     '<table style="border-collapse:collapse;font-size:12px;min-width:100%;"><thead><tr>',
                     '<th style="background:#FFD700;color:#000;font-weight:800;padding:14px 18px;'
                     'border:1px solid #BBB;text-align:left;min-width:230px;font-size:13px;">POG Cluster</th>']
                if _clsx:
                    for _c in _clsx:
                        _ph.append(f'<th style="background:#FFD700;color:#000;font-weight:700;padding:10px 14px;'
                                   f'border:1px solid #BBB;text-align:center;min-width:100px;white-space:nowrap;">{_c}</th>')
                else:
                    _ph.append('<th style="background:#FFD700;color:#888;padding:10px 14px;border:1px solid #BBB;'
                               'font-size:11px;font-style:italic;min-width:260px;">Add cluster columns using the input above ↑</th>')
                _ph.append('</tr></thead><tbody>')
                for _rn in _rowsx:
                    _rbg,_rtc=st.session_state[_kr].get(_rn,("#FFFFFF","#000000"))
                    _ph.append(f'<tr><td style="background:{_rbg};color:{_rtc};font-weight:600;'
                               f'padding:8px 18px;border:1px solid #CCC;white-space:nowrap;">{_rn}</td>')
                    if _clsx:
                        for _c in _clsx:
                            _v=st.session_state[_kd].get(_rn,{}).get(_c,"")
                            _v="" if _v is None or str(_v) in ("nan","None") else str(_v)
                            _ph.append(f'<td style="background:{_rbg};color:{_rtc};padding:8px 12px;'
                                       f'border:1px solid #CCC;text-align:center;">{_v}</td>')
                    else:
                        _ph.append(f'<td style="background:{_rbg};padding:8px 12px;border:1px solid #CCC;"></td>')
                    _ph.append('</tr>')
                _ph.append('</tbody></table></div>')
                st.markdown(''.join(_ph),unsafe_allow_html=True)
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

        # ── Status ────────────────────────────────────────────────────────────
        elif _subview == "📊 Status":
            if _stc:
                _sc=df_view[_stc].astype(str).str.strip().value_counts().reset_index()
                _sc.columns=["Status","Count"]
                _sc=_sc[~_sc["Status"].isin(["nan",""])]; _tot=_sc["Count"].sum()
                _sh=['<div style="overflow-x:auto;border-radius:12px;border:1px solid #E0D9D2;margin-top:14px;">',
                     '<table style="border-collapse:collapse;font-size:12px;width:100%;">',
                     '<thead><tr style="background:#1C1C1E;">',
                     '<th style="padding:10px 16px;text-align:left;color:rgba(255,255,255,.75);font-size:11px;font-weight:700;">STATUS</th>',
                     '<th style="padding:10px 16px;text-align:center;color:rgba(255,255,255,.75);font-size:11px;font-weight:700;">COUNT</th>',
                     '<th style="padding:10px 16px;text-align:center;color:rgba(255,255,255,.75);font-size:11px;font-weight:700;">%</th>',
                     '<th style="padding:10px 16px;color:rgba(255,255,255,.75);font-size:11px;font-weight:700;">BAR</th>',
                     '</tr></thead><tbody>']
                for _si,_sr in _sc.iterrows():
                    _sv3=str(_sr["Status"]); _scc=STATUS_COLORS.get(_sv3,{"bg":"#F5F5F5","c":"#888"})
                    _pct=_sr["Count"]/_tot*100 if _tot else 0; _rb2="#FFFFFF" if _si%2==0 else "#F8F4F0"
                    _sh.append(f'<tr style="background:{_rb2};">'
                               f'<td style="padding:10px 16px;border-bottom:1px solid #F0EBE3;">'
                               f'<span style="background:{_scc["bg"]};color:{_scc["c"]};padding:3px 10px;'
                               f'border-radius:4px;font-weight:700;font-size:11px;">{_sv3}</span></td>'
                               f'<td style="padding:10px 16px;text-align:center;border-bottom:1px solid #F0EBE3;'
                               f'font-weight:700;font-size:13px;">{_sr["Count"]:,}</td>'
                               f'<td style="padding:10px 16px;text-align:center;border-bottom:1px solid #F0EBE3;'
                               f'color:#888;font-size:12px;">{_pct:.1f}%</td>'
                               f'<td style="padding:10px 16px;border-bottom:1px solid #F0EBE3;">'
                               f'<div style="background:#EDE8DF;border-radius:99px;height:6px;">'
                               f'<div style="background:{_scc["c"]};border-radius:99px;height:6px;'
                               f'width:{min(_pct,100):.1f}%;"></div></div></td></tr>')
                _sh.append('</tbody></table></div>')
                st.markdown(''.join(_sh),unsafe_allow_html=True)
                st.caption(f"{len(_sc)} statuses · {_tot:,} total SKUs")
                st.download_button("⬇️ Export Status Summary",df_to_xlsx_bytes(_sc),
                    file_name=f"status_{p}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"{p}_st_dl")
            else:
                st.warning("Status column not found in the data.")

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
        st.dataframe(_cdf.reset_index(drop=True),use_container_width=True,height=300,hide_index=True)
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
    if "ib_data" not in st.session_state:
        st.session_state.ib_data = pd.DataFrame(columns=_IB_COLS)

    _ib_c1, _ib_c2, _ib_c3, _ = st.columns([1.1, 1.0, 1.4, 4.5])
    with _ib_c1:
        if st.button("＋ Add Row", key="ib_add_row", use_container_width=True):
            _empty = pd.DataFrame([{c: None for c in _IB_COLS}])
            st.session_state.ib_data = pd.concat(
                [st.session_state.ib_data, _empty], ignore_index=True)
            st.rerun()
    with _ib_c2:
        if st.button("↺ Clear All", key="ib_clear", use_container_width=True):
            st.session_state.ib_data = pd.DataFrame(columns=_IB_COLS)
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
    if "ibs_sc_data" not in st.session_state:
        st.session_state.ibs_sc_data = pd.DataFrame(columns=_IBS_COLS)

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
        if st.button("↺ Clear All", key="ibs_clear", use_container_width=True):
            st.session_state.ibs_sc_data = pd.DataFrame(columns=_IBS_COLS)
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
    if "vw_prodlib_data" not in st.session_state:
        st.session_state.vw_prodlib_data = pd.DataFrame(columns=_PRODLIB_COLS)

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
        if st.button("↺ Clear All", key="pl_clear", use_container_width=True):
            st.session_state.vw_prodlib_data = pd.DataFrame(columns=_PRODLIB_COLS)
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
    if "vw_citrix_data" not in st.session_state:
        st.session_state.vw_citrix_data = pd.DataFrame(columns=_CITRIX_COLS)

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
        if st.button("↺ Clear All", key="cx_clear", use_container_width=True):
            st.session_state.vw_citrix_data = pd.DataFrame(columns=_CITRIX_COLS)
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

with _sheet_tabs[0]:   # Range Sheet_Non-SSPOG
    _render_sheet_content(merged, "ns")

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
            ("Department",                             ["Department", "Dept", "Department Code&Desc", "department_code_desc"]),
            ("Section",                                ["Section", "section"]),
            ("Subclass",                               ["Subclass", "SubClass", "Sub Class", "subclass"]),
            ("Barcode",                                ["Barcode", "barcode", "UPC", "EAN", "ean"]),
            ("TPNA",                                   ["TPNA", "tpna"]),
            ("ID",                                     ["ID", "id", "Item ID", "ItemID"]),
            ("No. of Unit in Case",                    ["No. of Unit in Case", "No_of_Unit_in_Case", "Units Per Case", "Case Units", "no of unit in case"]),
            ("No. of Unit in Inner",                   ["No. of Unit in Inner", "No_of_Unit_in_Inner", "no of unit in inner"]),
            ("Tray total number",                      ["Tray total number", "Tray_total_number", "Tray Total Number", "tray total"]),
            ("Express Picking Type",                   ["Express Picking Type", "Express_Picking_Type", "express picking type"]),
            ("HDET Picking Type",                      ["HDET Picking Type", "HDET_Picking_Type", "hdet picking type"]),
            ("EDLP Price by Format Item name",         ["EDLP Price by Format", "EDLP_Price_by_Format", "Item Name", "Item name", "edlp price by format"]),
            ("As IS planograms applied",               ["AS IS planograms applied", "As IS planograms applied", "AS-IS planograms applied", "ASIS planograms applied", "as is planograms applied"]),
            ("To-BE planograms applied",               ["TO-BE planograms applied", "To-BE planograms applied", "TOBE planograms applied", "to be planograms applied"]),
            ("AS-IS Store applied",                    ["AS-IS Stores Applied", "AS IS Stores Applied", "ASIS Stores Applied", "AS-IS Store applied", "as-is stores applied"]),
            ("To-Be store applied",                    ["TO-Be stores applied", "To-Be stores applied", "TOBE stores applied", "to-be stores applied", "to be stores applied"]),
            ("Avg unit 52 wk/forecast new item sales", ["Avg Units 52wk/ Forecast new item sales", "Avg Units 52wk/Forecast new item sales", "avg units 52wk/ forecast new item sales", "Avg unit 52wk"]),
            ("Supplier pack size",                     ["Supplier Pack Size", "Supplier pack size", "Supplier_Pack_Size", "supplier pack size"]),
            ("Range Tail YYYY",                        ["Range Tail YYYY", "Range_Tail_YYYY", "range tail yyyy"]),
            ("AVG selling Price by format",            ["AVG Selling Price by Format", "Avg Selling Price by Format", "avg selling price by format", "AVG_Selling_Price_by_Format"]),
            ("Star Line",                              ["Star Line", "Star_Line", "starline", "star line"]),
            ("Item priority",                          ["Item Priority", "Item priority", "Item_Priority", "item priority"]),
            ("JDA vs Actual",                          ["JDA vs Actual", "JDA_vs_Actual", "jda vs actual"]),
            ("Actual-Actual",                          ["Actual-Actual", "Actual_Actual", "actual-actual", "actual actual"]),
        ]
        # Group definitions: (label, bg-color, text-color, column-count)
        _VIEW_GROUPS = [
            ("Item Info",  "#D9D9D9", "#333333", 12),
            ("Range Info", "#E8E3DC", "#444444",  8),
            ("Star Line",  "#000000", "#FFFFFF",   1),
            ("Priority",   "#00CC44", "#003300",   3),
        ]

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
                result[out_name] = (
                    src_df[src_col].reset_index(drop=True) if src_col
                    else pd.Series([None] * len(src_df), name=out_name)
                )
            return pd.DataFrame(result)

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

render_page_nav("viewdata")
