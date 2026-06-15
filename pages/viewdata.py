"""View Data page — sheet tabs, Range Architecture card, data table."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
import pandas as pd
from datetime import datetime
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar, render_page_nav,
    RS_SHEETS, RS_COL_GROUPS, STATUS_COLORS, FILL_COLORS, COLUMN_LABELS,
    get_fill, df_to_xlsx_bytes, df_to_csv_bytes, add_audit,
)

inject_css()
init_session_state()
render_sidebar("viewdata")
render_topbar("View Data")

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
def _nc(s): return str(s).lower().strip().replace('\n', ' ').replace('  ', ' ')
_dcm = {_nc(c): c for c in all_cols}
def _fc(key): return _dcm.get(_nc(key))
def _col_label(c): return COLUMN_LABELS.get(_nc(c), c)

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

# ── Sheet tabs (scrollable via st.tabs) ───────────────────────────────────────
_sheet_tabs = st.tabs(RS_SHEETS)

# Secondary tabs — placeholders / upload widgets
with _sheet_tabs[1]:   # Range Sheet_SSPOG
    st.info("Range Sheet_SSPOG — use the SSPOG filter in the All Data tab on the first sheet.")
with _sheet_tabs[2]:   # StoreApply_SSPOG
    st.info("StoreApply_SSPOG — data will appear here when uploaded.")
with _sheet_tabs[3]:   # 5.1 ItembyStore
    st.info("5.1 ItembyStore — data will appear here when uploaded.")
with _sheet_tabs[4]:   # 5.2 ItembyStore_SC
    st.info("5.2 ItembyStore_SC — data will appear here when uploaded.")
with _sheet_tabs[5]:   # 5.3 Upload_product_library
    _PRODLIB_COLS = ["ID", "Product Description", "Mod_structure_fixture"]
    if "vw_prodlib_data" not in st.session_state:
        st.session_state.vw_prodlib_data = pd.DataFrame(columns=_PRODLIB_COLS)

    _pf = st.file_uploader(
        "Upload product library file (.xlsx / .csv)",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=False,
        key="vw_upload_prodlib",
    )
    if _pf is not None:
        try:
            _raw_pl = (pd.read_csv(_pf) if _pf.name.lower().endswith(".csv")
                       else pd.read_excel(_pf))
            for _pc in _PRODLIB_COLS:
                if _pc not in _raw_pl.columns:
                    _raw_pl[_pc] = ""
            st.session_state.vw_prodlib_data = _raw_pl[_PRODLIB_COLS]
        except Exception as _pe:
            st.error(f"Could not read file: {_pe}")

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
        height=500,
    )
    st.session_state.vw_prodlib_data = _prodlib_edited
    _pl1, _pl2 = st.columns(2)
    _pl1.caption(f"{len(_prodlib_edited):,} rows")
    with _pl2:
        if not _prodlib_edited.empty:
            st.download_button(
                "⬇️ Export .xlsx", df_to_xlsx_bytes(_prodlib_edited),
                file_name=f"product_library_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
with _sheet_tabs[6]:   # 5.4 Upload to Citrix
    st.file_uploader(
        "Upload file to Citrix",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=False,
        key="vw_upload_citrix",
    )

# Sub-tabs live inside sheet tab 0 (Range Sheet_Non-SSPOG); content is written
# into tab_all / tab_canvas below — Streamlit's DeltaGenerator remembers the parent.
with _sheet_tabs[0]:
    tab_all, tab_canvas = st.tabs(["📊  All Data", "🎨  New Canvas"])

# ══════════════════════════════════════════════════════════════════════════════
# Tab: All Data
# ══════════════════════════════════════════════════════════════════════════════
with tab_all:
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
        _tdf = df_view[disp_cols].head(_MAX)
        def _cw(col): return max(80, min(200, len(_col_label(col)) * 7 + 16))

        _h = [
            '<div style="overflow-x:auto;border-radius:12px;border:1px solid #E0D9D2;margin-top:14px;">',
            '<table style="border-collapse:collapse;font-size:11px;min-width:100%;">',
            '<thead><tr>',
        ]
        for _g in _disp_groups:
            _gc = _g.get("hdr_color", "#555555")
            _h.append(
                f'<th colspan="{len(_g["disp_cols"])}" style="padding:6px 8px;text-align:center;'
                f'background:{_g["color"]};border:1px solid #E0D9D2;font-size:9px;font-weight:700;'
                f'color:{_gc};letter-spacing:.05em;text-transform:uppercase;">{_g["group"]}</th>')
        _h.append('</tr><tr style="background:#1C1C1E;">')
        for _g in _disp_groups:
            for _c in _g["disp_cols"]:
                _fill = get_fill(_c)
                _dot  = FILL_COLORS.get(_fill, {}).get("bg", "transparent")
                _w    = _cw(_c)
                _h.append(
                    f'<th style="padding:9px 10px;text-align:left;font-weight:700;'
                    f'color:rgba(255,255,255,.75);font-size:10px;white-space:nowrap;'
                    f'min-width:{_w}px;max-width:{_w}px;border-right:1px solid rgba(255,255,255,.07);">'
                    f'<div style="display:flex;align-items:center;gap:4px;">'
                    f'<span style="width:5px;height:5px;border-radius:50%;background:{_dot};'
                    f'border:1px solid rgba(255,255,255,.3);flex-shrink:0;"></span>'
                    f'<span style="overflow:hidden;text-overflow:ellipsis;">{_col_label(_c)}</span></div></th>')
        _h.append('</tr></thead><tbody>')

        for _i, _row in _tdf.iterrows():
            _rb = "#FFFFFF" if _i % 2 == 0 else "#F8F4F0"
            _h.append(f'<tr style="background:{_rb};">')
            for _g in _disp_groups:
                for _c in _g["disp_cols"]:
                    _val  = _row.get(_c, "")
                    _sv   = "" if pd.isna(_val) or str(_val) == "nan" else str(_val)
                    _fill = get_fill(_c)
                    _cbg  = ("background:#FFFDE740;" if _fill == "display" else
                             "background:#FCE4EC25;" if _fill == "mer" else "")
                    _w    = _cw(_c)
                    _cl   = _c.strip().lower()
                    if _cl == "status" and _sv:
                        _sc2   = STATUS_COLORS.get(_sv.strip(), {"bg": "#F5F5F5", "c": "#888"})
                        _inner = (f'<span style="background:{_sc2["bg"]};color:{_sc2["c"]};'
                                  f'padding:2px 8px;border-radius:4px;font-size:10px;'
                                  f'font-weight:700;white-space:nowrap;">{_sv}</span>')
                    elif _cl in ("barcode", "id") and _sv:
                        _inner = (f'<span style="color:#2BBFA4;font-family:monospace;'
                                  f'font-weight:600;">{_sv}</span>')
                    else:
                        _inner = f'<span>{_sv}</span>'
                    _h.append(
                        f'<td style="padding:8px 10px;border-bottom:1px solid #E0D9D2;'
                        f'border-right:1px solid #E0D9D2;{_cbg}min-width:{_w}px;max-width:{_w}px;'
                        f'overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">{_inner}</td>')
            _h.append('</tr>')
        _h.append('</tbody></table></div>')
        st.markdown(''.join(_h), unsafe_allow_html=True)

        if len(df_view) > _MAX:
            st.caption(f"Showing {_MAX:,} of {len(df_view):,} rows — increase Rows input to see more")

        _fc1, _fc2, _fc3 = st.columns([3, 1, 1])
        with _fc1:
            st.caption(f"{len(df_view):,} rows · {len(disp_cols)} columns · raw: {len(merged):,} rows")
        with _fc2:
            st.download_button("⬇️ Export .xlsx",
                df_to_xlsx_bytes(df_view[disp_cols]),
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
        # Detect cluster name column
        _clust_c = next(
            (c for c in all_cols if "cluster" in c.lower() and "planogram" in c.lower()), None
        ) or next((c for c in all_cols if "cluster" in c.lower()), None)

        # Detect "extra" columns from a second file (not matched by any RS_COL_GROUPS)
        _known_flat = set()
        for _rg in RS_COL_GROUPS:
            for _k in _rg["cols"]:
                _mc = _fc(_k)
                if _mc:
                    _known_flat.add(_mc)
        _extra_cols = [c for c in all_cols if c not in _known_flat]

        # Metric columns that drive the pivot rows (fuzzy-matched)
        _PIVOT_METRICS = [
            ("TO-BE Stores applied count", "to be stores applied count"),
            ("AS-IS Stores applied count", "as is"),
            ("MODs",                       "mods"),
            ("FIXTURE",                    "fixtures"),
            ("RANGE CLASS",                "range class"),
            ("Total NEW SKUs",             "total new skus"),
            ("Total DELETE SKUs",          "total delete skus"),
            ("%Achieving LRD CASE (As Is)","%achieving lrd (as is)"),
            ("%Achieving CRD CASE (As Is)","%achieving crd case (as is)"),
        ]
        _met_cols = {}
        for _label, _key in _PIVOT_METRICS:
            _hit = next((c for c in all_cols if _key in c.lower()), None)
            if _hit:
                _met_cols[_label] = _hit

        # ── If extra columns exist (from cluster file), show them as pivot ──
        if _extra_cols and _clust_c:
            st.markdown(
                "<div style='font-size:13px;font-weight:700;color:#3D0070;margin:12px 0 8px;'>"
                "Cluster columns detected from uploaded files</div>",
                unsafe_allow_html=True,
            )
            _cdf = df_view[[_clust_c] + [c for c in _extra_cols if c in df_view.columns]].copy()
            st.dataframe(_cdf.reset_index(drop=True), use_container_width=True,
                         height=320, hide_index=True)

        # ── Pivot from Cluster (Planogram name) column ────────────────────────
        if _clust_c and _met_cols:
            _uniq_clusters = (
                df_view[_clust_c].dropna().astype(str)
                .str.strip().unique().tolist()
            )
            _uniq_clusters = [c for c in _uniq_clusters if c not in ("", "nan")][:40]

            if _uniq_clusters:
                # Build pivot rows
                _pivot_data = {}
                for _label, _src in _met_cols.items():
                    _row_vals = {}
                    for _cl_name in _uniq_clusters:
                        _mask = df_view[_clust_c].astype(str).str.strip() == _cl_name
                        _sub  = df_view.loc[_mask, _src]
                        _num  = pd.to_numeric(_sub, errors="coerce")
                        if _num.notna().any():
                            _row_vals[_cl_name] = f"{_num.sum():.0f}" if "count" in _label.lower() or "sku" in _label.lower() else f"{_num.mean():.2f}"
                        else:
                            _mode = _sub.dropna().mode()
                            _row_vals[_cl_name] = str(_mode.iloc[0]) if not _mode.empty else ""
                    _pivot_data[_label] = _row_vals

                # Render pivot HTML table
                _ph = [
                    '<div style="overflow-x:auto;border-radius:12px;border:1px solid #E0D9D2;margin-top:14px;">',
                    '<table style="border-collapse:collapse;font-size:11px;min-width:100%;">',
                    '<thead><tr>',
                    '<th style="padding:8px 14px;background:#FFD700;color:#333;font-weight:800;'
                    'font-size:12px;border:1px solid #E0D9D2;min-width:200px;">Cluster</th>',
                ]
                for _cn in _uniq_clusters:
                    _short = _cn if len(_cn) <= 18 else _cn[:16] + "…"
                    _ph.append(
                        f'<th style="padding:6px 10px;background:#C9A0DC;color:#3D0070;'
                        f'font-weight:700;font-size:10px;border:1px solid #E0D9D2;'
                        f'min-width:90px;text-align:center;white-space:nowrap;" title="{_cn}">'
                        f'{_short}</th>'
                    )
                _ph.append('</tr></thead><tbody>')

                for _ri, (_label, _row_vals) in enumerate(_pivot_data.items()):
                    _rb2 = "#FFFFFF" if _ri % 2 == 0 else "#F8F4F0"
                    _ph.append(
                        f'<tr style="background:{_rb2};">'
                        f'<td style="padding:8px 14px;border:1px solid #E0D9D2;font-weight:600;'
                        f'font-size:12px;color:#1A1A1A;white-space:nowrap;">{_label}</td>'
                    )
                    for _cn in _uniq_clusters:
                        _v = _row_vals.get(_cn, "")
                        _ph.append(
                            f'<td style="padding:7px 10px;border:1px solid #E0D9D2;'
                            f'text-align:center;font-size:12px;color:#333;">{_v}</td>'
                        )
                    _ph.append('</tr>')

                _ph.append('</tbody></table></div>')
                st.markdown(''.join(_ph), unsafe_allow_html=True)
                st.caption(f"{len(_uniq_clusters)} clusters · {len(_met_cols)} metrics")
            else:
                st.info("No cluster values found in the data.")
        else:
            st.warning(
                "Cluster column not found. Upload a file that contains a 'Cluster (Planogram name)' column."
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

# ══════════════════════════════════════════════════════════════════════════════
# Tab: New Canvas
# ══════════════════════════════════════════════════════════════════════════════
with tab_canvas:
    st.markdown("""
<div style="font-size:16px;font-weight:700;color:#1A1A1A;margin-bottom:4px;">New Canvas</div>
<div style="font-size:13px;color:#888;margin-bottom:16px;">
    Select columns, filter rows, save as a named canvas.
</div>""", unsafe_allow_html=True)

    _cv1, _cv2 = st.columns(2)
    with _cv1:
        _cnv_srch = st.text_input("Filter", placeholder="🔎  Filter rows...",
                                  label_visibility="collapsed")
    with _cv2:
        _cnv_name = st.text_input("Canvas name", placeholder="e.g. SSPOG-A Items",
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
            lambda r: r.astype(str).str.contains(_cnv_srch, case=False, na=False).any(), axis=1)]

    st.dataframe(_cdf.reset_index(drop=True), use_container_width=True,
                 height=300, hide_index=True)
    st.caption(f"{len(_cdf):,} rows · {len(_cnv_cols)} columns")

    _ccb1, _ccb2 = st.columns(2)
    with _ccb1:
        if st.button("💾 Save Canvas", use_container_width=True):
            _nm = _cnv_name.strip() or f"Canvas {len(st.session_state.canvases)+1}"
            st.session_state.canvases[_nm] = _cdf.copy()
            add_audit("Save Canvas", _nm)
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

render_page_nav("viewdata")
