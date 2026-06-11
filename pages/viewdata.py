"""View Data page — Excel-like table viewer with column groups."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
import pandas as pd
from datetime import datetime
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar,
    RS_SHEETS, RS_COL_GROUPS, STATUS_COLORS, FILL_COLORS,
    get_fill, find_col, df_to_xlsx_bytes, df_to_csv_bytes, add_audit,
)

inject_css()
init_session_state()
render_sidebar("viewdata")
render_topbar("View Data")

merged = st.session_state.merged_df
if merged is None:
    st.warning("No data yet — go to **My Files** and upload first.")
    st.stop()

all_cols = list(merged.columns)

# ── Match column groups ────────────────────────────────────────────────────────
def _nc(s): return str(s).lower().strip().replace('\n',' ').replace('  ',' ')
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

# Reset stale vis_cols when df columns change
if st.session_state.view_vis_cols is not None:
    if not any(c in all_cols for c in st.session_state.view_vis_cols):
        st.session_state.view_vis_cols = None

if st.session_state.view_vis_cols is None:
    st.session_state.view_vis_cols = _all_grouped[:] or all_cols[:]

# ── Sheet tabs (visual) ────────────────────────────────────────────────────────
_stabs = ""
for s in RS_SHEETS:
    _a = (s == st.session_state.view_sheet)
    _stabs += (
        f'<div style="padding:5px 13px;border-radius:6px 6px 0 0;'
        f'border-top:{"2px solid #2BBFA4" if _a else "1px solid #E0D9D2"};'
        f'border-left:1px solid #E0D9D2;border-right:1px solid #E0D9D2;'
        f'border-bottom:{"2px solid #fff" if _a else "1px solid #E0D9D2"};'
        f'background:{"#fff" if _a else "#F8F4F0"};'
        f'color:{"#1A1A1A" if _a else "#888"};font-weight:{"700" if _a else "400"};'
        f'font-size:12px;white-space:nowrap;display:inline-block;">{s}</div>'
    )
st.markdown(
    f'<div style="display:flex;gap:3px;flex-wrap:wrap;">{_stabs}</div>'
    f'<div style="border-top:1px solid #E0D9D2;margin-bottom:14px;"></div>',
    unsafe_allow_html=True)
_ns = st.selectbox("Sheet", RS_SHEETS,
                   index=RS_SHEETS.index(st.session_state.view_sheet)
                         if st.session_state.view_sheet in RS_SHEETS else 1,
                   label_visibility="collapsed", key="vw_sheet_sel")
if _ns != st.session_state.view_sheet:
    st.session_state.view_sheet = _ns
    st.rerun()

tab_all, tab_canvas = st.tabs(["📊 All Data", "🎨 New Canvas"])

# ── Tab: All Data ─────────────────────────────────────────────────────────────
with tab_all:
    _meta = st.session_state.rangesheet_meta

    _dg_c, _arch_c, _leg_c = st.columns([1.1, 2.8, 0.75])

    with _dg_c:
        _dg_rows = ""
        for _k, _v, _hl in [
            ("DG CODE",         _meta.get("dg_code","—"),        True),
            ("DG NAME",         _meta.get("dg_name","—"),        True),
            ("MINOR LIVE WEEK", _meta.get("minor_live_week","—"),False),
            ("MAJOR LIVE WEEK", _meta.get("major_live_week","—"),False),
            ("Event Live Date", _meta.get("event_live_date","—"),False),
            ("Event Des",       _meta.get("event_desc","—"),     False),
        ]:
            _bg = "#FFF9C4" if _hl else "transparent"
            _dg_rows += (
                f'<div style="display:flex;justify-content:space-between;'
                f'margin-bottom:5px;font-size:12px;">'
                f'<span style="color:#888;">{_k}</span>'
                f'<span style="font-weight:600;background:{_bg};padding:0 4px;'
                f'border-radius:3px;">{_v}</span></div>')
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
                for _t in ["MAINTAIN","DELETE SOME","DELETE ALL","NEW SOME","NEWNEW"]:
                    _cnt = int((merged[_sc].astype(str).str.strip() == _t).sum())
                    _ai  = _cnt if _t == "MAINTAIN" or "DELETE" in _t else 0
                    _tb  = _cnt if _t == "MAINTAIN" or "NEW" in _t    else 0
                    _arch_rows.append({"type": _t, "as_is": _ai, "to_be": _tb, "diff": _tb - _ai})
                _tai = sum(r["as_is"] for r in _arch_rows)
                _ttb = sum(r["to_be"] for r in _arch_rows)
                _arch_rows.append({"type": "TOTAL SKU", "as_is": _tai, "to_be": _ttb,
                                   "diff": _ttb - _tai, "_total": True})
        _atbody = ""
        for _r in _arch_rows:
            _tot = _r.get("_total") or "TOTAL" in str(_r.get("type","")).upper()
            _rbg = "#1C1C1E" if _tot else "transparent"
            _rc  = "#fff"    if _tot else "#1A1A1A"
            _fw  = "700"     if _tot else "400"
            _atbody += (
                f'<tr style="background:{_rbg};">'
                f'<td style="padding:6px 10px;font-weight:{_fw};color:{_rc};">{_r.get("type","")}</td>'
                f'<td style="padding:6px 8px;text-align:center;color:{_rc};">{_r.get("as_is","")}</td>'
                f'<td style="padding:6px 8px;text-align:center;color:{_rc};">{_r.get("to_be","")}</td>'
                f'<td style="padding:6px 8px;text-align:center;color:{"#fff" if _tot else "#888"};">{_r.get("diff","")}</td>'
                f'</tr>')
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
                     border-bottom:1px solid #E0D9D2;min-width:55px;">AS IS</th>
                <th style="padding:7px 8px;text-align:center;font-weight:700;color:#888;
                     border-bottom:1px solid #E0D9D2;min-width:55px;">TO BE</th>
                <th style="padding:7px 8px;text-align:center;font-weight:700;color:#888;
                     border-bottom:1px solid #E0D9D2;min-width:55px;">DIFF</th>
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
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;">
            <div style="width:24px;height:14px;background:#FFFDE7;border:1px solid #ddd;
                 border-radius:3px;"></div><span style="font-size:11px;">Display fill</span>
          </div>
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;">
            <div style="width:24px;height:14px;background:#FCE4EC;border:1px solid #ddd;
                 border-radius:3px;"></div><span style="font-size:11px;">Merchandiser fill</span>
          </div>
          <div style="display:flex;align-items:center;gap:8px;">
            <div style="width:24px;height:14px;background:#F5F5F5;border:1px solid #ddd;
                 border-radius:3px;"></div><span style="font-size:11px;">Formula</span>
          </div>
        </div>''', unsafe_allow_html=True)

    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

    # Controls
    _ctrl1, _ctrl2, _ctrl3, _ctrl4 = st.columns([3, 3, 1, 1.2])
    with _ctrl1:
        _filter_mode = st.radio("FILTER:", ["ALL","SSPOG","NON-SSPOG"],
                                horizontal=True, label_visibility="visible")
    with _ctrl2:
        _search_q = st.text_input("Search", placeholder="🔎  Search item, barcode, status...",
                                  label_visibility="collapsed")
    with _ctrl3:
        _n_rows = st.number_input("Rows", min_value=1, max_value=10000, value=10, step=10,
                                  help="Number of rows to display")
    with _ctrl4:
        _edit_on = st.toggle("⚙️ Edit Columns", key="vw_edit_cols")

    if _edit_on:
        st.markdown('''<div style="background:#F8F4F0;border-radius:10px;padding:10px 16px 6px;
            border:1px solid #E0D9D2;margin-bottom:10px;">
          <div style="font-size:12px;font-weight:700;margin-bottom:6px;">
            ⚙️ Column visibility — <span style="color:#2BBFA4;">display only</span>
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

    _vis      = st.session_state.view_vis_cols or _all_grouped or all_cols
    disp_cols = [c for c in _vis if c in merged.columns] or all_cols[:20]

    _disp_groups = []
    for _g in active_groups:
        _gc = [c for c in _g["matched"] if c in disp_cols]
        if _gc:
            _disp_groups.append({**_g, "disp_cols": _gc})

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

    # Metrics
    _stc         = next((c for c in all_cols if c.strip().lower() == "status"), None)
    _m_maintain  = int((df_view[_stc].astype(str).str.strip() == "MAINTAIN").sum()) if _stc else 0
    _m_sspog     = int(df_view[_pog_col].astype(str).str.contains("SSPOG", na=False).sum() -
                       df_view[_pog_col].astype(str).str.contains("Non-SSPOG", na=False).sum()
                       ) if _pog_col else 0
    _m_nonsspog  = int(df_view[_pog_col].astype(str).str.contains("Non-SSPOG", na=False).sum()
                       ) if _pog_col else 0
    _m_null      = int(df_view[disp_cols].isnull().sum().sum())
    _tm1,_tm2,_tm3,_tm4,_tm5 = st.columns(5)
    _tm1.metric("Total SKUs",  f"{len(df_view):,}")
    _tm2.metric("MAINTAIN",    f"{_m_maintain:,}")
    _tm3.metric("SSPOG",       f"{_m_sspog:,}")
    _tm4.metric("Non-SSPOG",   f"{_m_nonsspog:,}")
    _tm5.metric("Null Values", f"{_m_null:,}")

    # HTML data table
    _MAX = int(_n_rows)
    _tdf = df_view[disp_cols].head(_MAX)

    def _cw(col): return max(80, min(200, len(str(col))*7+16))

    _h = ['<div style="overflow-x:auto;border-radius:10px;border:1px solid #E0D9D2;margin-top:12px;">',
          '<table style="border-collapse:collapse;font-size:11px;min-width:100%;">',
          '<thead><tr>']
    for _g in _disp_groups:
        _h.append(f'<th colspan="{len(_g["disp_cols"])}" style="padding:5px 8px;text-align:center;'
                  f'background:{_g["color"]};border:1px solid #E0D9D2;font-size:9px;font-weight:700;'
                  f'color:#555;letter-spacing:.05em;text-transform:uppercase;">{_g["group"]}</th>')
    _h.append('</tr><tr style="background:#1C1C1E;">')
    for _g in _disp_groups:
        for _c in _g["disp_cols"]:
            _fill = get_fill(_c)
            _dot  = FILL_COLORS.get(_fill,{}).get("bg","transparent")
            _w    = _cw(_c)
            _h.append(
                f'<th style="padding:8px 10px;text-align:left;font-weight:700;'
                f'color:rgba(255,255,255,.75);font-size:10px;white-space:nowrap;'
                f'min-width:{_w}px;max-width:{_w}px;border-right:1px solid rgba(255,255,255,.08);">'
                f'<div style="display:flex;align-items:center;gap:4px;">'
                f'<span style="width:5px;height:5px;border-radius:50%;background:{_dot};'
                f'border:1px solid rgba(255,255,255,.3);flex-shrink:0;display:inline-block;"></span>'
                f'<span style="overflow:hidden;text-overflow:ellipsis;">{_c}</span></div></th>')
    _h.append('</tr></thead><tbody>')
    for _i, _row in _tdf.iterrows():
        _rb = "#FFFFFF" if _i%2==0 else "#F8F4F0"
        _h.append(f'<tr style="background:{_rb};">')
        for _g in _disp_groups:
            for _c in _g["disp_cols"]:
                _val  = _row.get(_c,"")
                _sv   = "" if pd.isna(_val) or str(_val)=="nan" else str(_val)
                _fill = get_fill(_c)
                _cbg  = ("background:#FFFDE740;" if _fill=="display" else
                         "background:#FCE4EC25;" if _fill=="mer" else "")
                _w    = _cw(_c)
                _cl   = _c.strip().lower()
                if _cl=="status" and _sv:
                    _sc2  = STATUS_COLORS.get(_sv.strip(), {"bg":"#F5F5F5","c":"#888"})
                    _inner = (f'<span style="background:{_sc2["bg"]};color:{_sc2["c"]};'
                              f'padding:2px 8px;border-radius:4px;font-size:10px;'
                              f'font-weight:700;white-space:nowrap;">{_sv}</span>')
                elif _cl in ("barcode","id") and _sv:
                    _inner = f'<span style="color:#2BBFA4;font-family:monospace;font-weight:600;">{_sv}</span>'
                else:
                    _inner = f'<span>{_sv}</span>'
                _h.append(f'<td style="padding:7px 10px;border-bottom:1px solid #E0D9D2;'
                          f'border-right:1px solid #E0D9D2;{_cbg}min-width:{_w}px;max-width:{_w}px;'
                          f'overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">{_inner}</td>')
        _h.append('</tr>')
    _h.append('</tbody></table></div>')
    st.markdown(''.join(_h), unsafe_allow_html=True)

    if len(df_view) > _MAX:
        st.caption(f"Showing {_MAX} of {len(df_view):,} rows — increase the Rows input or export to see all")

    _fc1, _fc2, _fc3 = st.columns([3,1,1])
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

# ── Tab: Canvas ────────────────────────────────────────────────────────────────
with tab_canvas:
    st.markdown('<div style="font-size:16px;font-weight:700;margin-bottom:4px;">New Canvas</div>'
                '<div style="font-size:13px;color:#888;margin-bottom:16px;">'
                'Select columns, filter rows, save as a named canvas.</div>',
                unsafe_allow_html=True)
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
            lambda r: r.astype(str).str.contains(_cnv_srch, case=False, na=False).any(), axis=1)]
    st.dataframe(_cdf.reset_index(drop=True), use_container_width=True, height=300, hide_index=True)
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
            _sc1, _sc2 = st.columns([4,1])
            with _sc1:
                st.write(f"🎨 **{_cnm}** — {len(_cs):,} rows · {len(_cs.columns)} cols")
            with _sc2:
                st.download_button("⬇️", df_to_xlsx_bytes(_cs),
                    file_name=f"{_cnm}.xlsx", key=f"cs_dl_{_cnm}")
