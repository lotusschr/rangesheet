"""View Data page — sheet tabs, Range Architecture card, data table."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
import pandas as pd
from datetime import datetime
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar, render_page_nav,
    RS_SHEETS, RS_COL_GROUPS, STATUS_COLORS, FILL_COLORS,
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

active_groups = []
for grp in RS_COL_GROUPS:
    matched = [_fc(k) for k in grp["cols"] if _fc(k) is not None]
    if matched:
        active_groups.append({**grp, "matched": matched})
_all_grouped = [c for g in active_groups for c in g["matched"]]
_remaining   = [c for c in all_cols if c not in _all_grouped]
if _remaining:
    active_groups.append({"group": "Other", "color": "#FAFAFA", "matched": _remaining})

if st.session_state.view_vis_cols is not None:
    if not any(c in all_cols for c in st.session_state.view_vis_cols):
        st.session_state.view_vis_cols = None
if st.session_state.view_vis_cols is None:
    st.session_state.view_vis_cols = _all_grouped[:] or all_cols[:]

# ── Sheet tabs (clickable radio styled as tabs) ───────────────────────────────
_ns = st.radio(
    "Sheet",
    RS_SHEETS,
    index=RS_SHEETS.index(st.session_state.view_sheet)
          if st.session_state.view_sheet in RS_SHEETS else 1,
    horizontal=True,
    label_visibility="collapsed",
    key="vw_sheet_radio",
)
if _ns != st.session_state.view_sheet:
    st.session_state.view_sheet = _ns
    st.rerun()

tab_all, tab_canvas = st.tabs(["📊  All Data", "🎨  New Canvas"])

# ══════════════════════════════════════════════════════════════════════════════
# Tab: All Data
# ══════════════════════════════════════════════════════════════════════════════
with tab_all:
    _meta = st.session_state.rangesheet_meta

    _dg_c, _arch_c, _leg_c = st.columns([1.05, 2.9, 0.75])

    # ── Display Group card ───────────────────────────────────────────────────
    with _dg_c:
        _dg_rows = ""
        for _k, _v, _hl in [
            ("DG CODE",         _meta.get("dg_code", "—"),         True),
            ("DG NAME",         _meta.get("dg_name", "—"),         True),
            ("MINOR LIVE WEEK", _meta.get("minor_live_week", "—"), False),
            ("MAJOR LIVE WEEK", _meta.get("major_live_week", "—"), False),
            ("Event Live Date", _meta.get("event_live_date", "—"), False),
            ("Event Des",       _meta.get("event_desc", "—"),      False),
        ]:
            _vbg = "#FFF9C4" if _hl and _v != "—" else "transparent"
            _dg_rows += (
                f'<div style="display:flex;justify-content:space-between;'
                f'align-items:center;margin-bottom:6px;font-size:12px;gap:6px;">'
                f'<span style="color:#888;white-space:nowrap;">{_k}</span>'
                f'<span style="font-weight:600;background:{_vbg};padding:1px 6px;'
                f'border-radius:4px;text-align:right;">{_v}</span></div>')
        st.markdown(f"""
<div style="background:#fff;border-radius:14px;border:1px solid #E8E3DC;padding:16px;">
    <div style="font-size:10px;font-weight:700;color:#888;text-transform:uppercase;
                letter-spacing:.08em;margin-bottom:12px;">Display Group</div>
    {_dg_rows}
    <div style="display:flex;gap:8px;margin-top:14px;">
        <div style="flex:1;background:#2BBFA4;color:#fff;border-radius:8px;padding:9px 6px;
                    text-align:center;font-size:11px;font-weight:700;cursor:pointer;">
            1. SELECT DG
        </div>
        <div style="flex:1;background:#E8A020;color:#fff;border-radius:8px;padding:9px 6px;
                    text-align:center;font-size:11px;font-weight:700;cursor:pointer;">
            2. SUBMIT RANGE
        </div>
    </div>
</div>""", unsafe_allow_html=True)

    # ── Range Architecture card (7-column) ────────────────────────────────────
    with _arch_c:
        _stc         = next((c for c in all_cols if c.strip().lower() == "status"), None)
        _sale_asis_c = next((c for c in all_cols if "as-is" in c.lower() and "sale" in c.lower()), None)
        _sale_tobe_c = next((c for c in all_cols if "to-be" in c.lower() and "sale" in c.lower()), None)
        _marg_asis_c = next((c for c in all_cols if "as-is" in c.lower() and "margin" in c.lower()), None)

        def _safe_num(series):
            return pd.to_numeric(series, errors="coerce").fillna(0)

        def _safe_sum(col, mask=None):
            if not col or col not in merged.columns:
                return 0
            s = merged[mask][col] if mask is not None else merged[col]
            return float(_safe_num(s).sum())

        _arch_rows_meta = list(_meta.get("range_arch", []))
        TYPES = ["MAINTAIN", "NEW DELETE SOME", "DELETE SOME", "DELETE ALL",
                 "NEW SOME", "NEW", "NEWNEW"]

        if not _arch_rows_meta and _stc:
            _arch_rows_meta = []
            for _t in TYPES:
                _msk  = merged[_stc].astype(str).str.strip() == _t
                _cnt  = int(_msk.sum())
                _ai   = _cnt if ("DELETE" in _t or _t == "MAINTAIN") else 0
                _tb   = _cnt if ("MAINTAIN" in _t or "NEW" in _t)    else 0
                _sale_ai = _safe_sum(_sale_asis_c, _msk)
                _sale_tb = _safe_sum(_sale_tobe_c, _msk)
                _marg_ai = _safe_sum(_marg_asis_c, _msk)
                _arch_rows_meta.append({
                    "type": _t, "as_is": _ai, "to_be": _tb,
                    "sale_ai": _sale_ai, "sale_tb": _sale_tb,
                    "sale_diff": _sale_tb - _sale_ai, "marg_ai": _marg_ai,
                })

        # Build totals
        _t_ai    = sum(r["as_is"]     for r in _arch_rows_meta)
        _t_tb    = sum(r["to_be"]     for r in _arch_rows_meta)
        _t_sai   = sum(r.get("sale_ai",   0) for r in _arch_rows_meta)
        _t_stb   = sum(r.get("sale_tb",   0) for r in _arch_rows_meta)
        _t_sdiff = sum(r.get("sale_diff", 0) for r in _arch_rows_meta)
        _t_mai   = sum(r.get("marg_ai",   0) for r in _arch_rows_meta)
        _pct_sku  = f"{(_t_tb/_t_ai-1)*100:.1f}%" if _t_ai else "0.0%"
        _pct_sale = f"{(_t_sdiff/_t_sai*100):.1f}%" if _t_sai else "0.0%"

        # AVG LRD CASE %
        _lrd_pct = "100.00%"

        def _fmt(v):
            if v == 0: return "0"
            try: return f"{int(v):,}" if v == int(v) else f"{v:,.1f}"
            except: return str(v)

        tbody = ""
        for _r in _arch_rows_meta:
            _t      = _r.get("type", "")
            _ai_bg  = "#FCE4EC" if _t in ("NEWNEW",)     else "transparent"
            _tb_bg  = "#FCE4EC" if _t in ("DELETE ALL",) else "transparent"
            tbody += (
                f'<tr style="border-bottom:1px solid #F0EBE3;">'
                f'<td style="padding:7px 12px;font-weight:500;color:#1A1A1A;'
                f'font-size:12px;">{_t}</td>'
                f'<td style="padding:7px 8px;text-align:center;background:{_ai_bg};'
                f'font-size:12px;">{_fmt(_r.get("as_is",0))}</td>'
                f'<td style="padding:7px 8px;text-align:center;background:{_tb_bg};'
                f'border-right:1px solid #E8E3DC;font-size:12px;">{_fmt(_r.get("to_be",0))}</td>'
                f'<td style="padding:7px 8px;text-align:center;color:#888;font-size:12px;">'
                f'{_fmt(_r.get("sale_ai",0))}</td>'
                f'<td style="padding:7px 8px;text-align:center;color:#888;font-size:12px;">'
                f'{_fmt(_r.get("sale_tb",0))}</td>'
                f'<td style="padding:7px 8px;text-align:center;color:#888;'
                f'border-right:1px solid #E8E3DC;font-size:12px;">'
                f'{_fmt(_r.get("sale_diff",0))}</td>'
                f'<td style="padding:7px 8px;text-align:center;color:#888;font-size:12px;">'
                f'{_fmt(_r.get("marg_ai",0))}</td>'
                f'</tr>')

        tbody += (
            f'<tr style="background:#1C1C1E;font-weight:700;">'
            f'<td style="padding:8px 12px;color:#fff;font-size:12px;">TOTAL SKU</td>'
            f'<td style="padding:8px;text-align:center;color:#fff;font-size:12px;">{_fmt(_t_ai)}</td>'
            f'<td style="padding:8px;text-align:center;color:#fff;border-right:1px solid #444;font-size:12px;">{_fmt(_t_tb)}</td>'
            f'<td style="padding:8px;text-align:center;color:#fff;font-size:12px;">{_fmt(_t_sai)}</td>'
            f'<td style="padding:8px;text-align:center;color:#fff;font-size:12px;">{_fmt(_t_stb)}</td>'
            f'<td style="padding:8px;text-align:center;color:#fff;border-right:1px solid #444;font-size:12px;">{_fmt(_t_sdiff)}</td>'
            f'<td style="padding:8px;text-align:center;color:#fff;font-size:12px;">{_fmt(_t_mai)}</td>'
            f'</tr>'
            f'<tr style="background:#F5F0EA;">'
            f'<td style="padding:6px 12px;color:#888;font-size:11px;">% Impact</td>'
            f'<td></td>'
            f'<td style="text-align:center;color:#2BBFA4;font-weight:700;font-size:11px;'
            f'border-right:1px solid #E8E3DC;">{_pct_sku}</td>'
            f'<td></td><td></td>'
            f'<td style="text-align:center;color:#2BBFA4;font-weight:700;font-size:11px;'
            f'border-right:1px solid #E8E3DC;">{_pct_sale}</td>'
            f'<td></td></tr>')

        st.markdown(f"""
<div style="background:#fff;border-radius:14px;border:1px solid #E8E3DC;overflow:hidden;">
    <div style="padding:11px 14px;border-bottom:1px solid #E8E3DC;
                display:flex;justify-content:space-between;align-items:center;">
        <span style="font-weight:700;font-size:13px;color:#1A1A1A;">Range Architecture</span>
        <span style="font-size:11px;color:#2BBFA4;font-weight:700;">
            AVG LRD CASE % : {_lrd_pct}
        </span>
    </div>
    <div style="overflow-x:auto;">
        <table style="width:100%;border-collapse:collapse;font-size:11px;">
            <thead>
                <tr style="background:#F5F0EA;">
                    <th style="padding:8px 12px;text-align:left;font-weight:700;color:#888;
                               border-bottom:1px solid #E8E3DC;min-width:130px;font-size:11px;">TYPE</th>
                    <th style="padding:8px;text-align:center;font-weight:700;color:#888;
                               border-bottom:1px solid #E8E3DC;min-width:55px;font-size:11px;">AS IS</th>
                    <th style="padding:8px;text-align:center;font-weight:700;color:#888;
                               border-bottom:1px solid #E8E3DC;min-width:55px;
                               border-right:1px solid #E8E3DC;font-size:11px;">TO BE</th>
                    <th style="padding:8px;text-align:center;font-weight:700;color:#888;
                               border-bottom:1px solid #E8E3DC;min-width:65px;font-size:11px;">AS IS (val)</th>
                    <th style="padding:8px;text-align:center;font-weight:700;color:#888;
                               border-bottom:1px solid #E8E3DC;min-width:65px;font-size:11px;">TO BE (val)</th>
                    <th style="padding:8px;text-align:center;font-weight:700;color:#888;
                               border-bottom:1px solid #E8E3DC;min-width:55px;
                               border-right:1px solid #E8E3DC;font-size:11px;">DIFF</th>
                    <th style="padding:8px;text-align:center;font-weight:700;color:#888;
                               border-bottom:1px solid #E8E3DC;min-width:55px;font-size:11px;">AS IS</th>
                </tr>
            </thead>
            <tbody>{tbody}</tbody>
        </table>
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

    _vis      = st.session_state.view_vis_cols or _all_grouped or all_cols
    disp_cols = [c for c in _vis if c in merged.columns] or all_cols[:20]

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

    # ── HTML data table ────────────────────────────────────────────────────────
    _MAX = int(_n_rows)
    _tdf = df_view[disp_cols].head(_MAX)
    def _cw(col): return max(80, min(200, len(str(col)) * 7 + 16))

    _h = [
        '<div style="overflow-x:auto;border-radius:12px;border:1px solid #E0D9D2;margin-top:14px;">',
        '<table style="border-collapse:collapse;font-size:11px;min-width:100%;">',
        '<thead><tr>',
    ]
    for _g in _disp_groups:
        _h.append(
            f'<th colspan="{len(_g["disp_cols"])}" style="padding:6px 8px;text-align:center;'
            f'background:{_g["color"]};border:1px solid #E0D9D2;font-size:9px;font-weight:700;'
            f'color:#555;letter-spacing:.05em;text-transform:uppercase;">{_g["group"]}</th>')
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
                f'<span style="overflow:hidden;text-overflow:ellipsis;">{_c}</span></div></th>')
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
