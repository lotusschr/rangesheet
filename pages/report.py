"""Report page — Range Execution & Summary Analytics."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
import pandas as pd
import zipfile
import io
from datetime import datetime
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar, render_page_nav,
    find_col, df_to_xlsx_bytes, df_to_csv_bytes, add_audit,
)

inject_css()
init_session_state()
render_sidebar("report")
render_topbar("Report")

merged = st.session_state.merged_df

rt1, rt2 = st.tabs(["📦  Range Execution", "📈  Summary & Analytics"])

# ══════════════════════════════════════════════════════════════════════════════
# Tab 1: Range Execution — 4 output slots sourced from View Data submissions
# ══════════════════════════════════════════════════════════════════════════════
with rt1:
    st.markdown("""
<div style="margin-bottom:24px;">
    <div style="font-size:18px;font-weight:700;color:#1A1A1A;margin-bottom:4px;">
        Range Execution
    </div>
    <div style="font-size:13px;color:#888;">
        Go to View Data → tables 5.1 – 5.4 → fill in rows → click ✅ Submit to Report,
        then preview and export here.
    </div>
</div>""", unsafe_allow_html=True)

    OUTPUTS = [
        ("01", "5.1 Item by Store",       "vw_submit_51", "#2BBFA4"),
        ("02", "5.2 Item by Store (SC)",  "vw_submit_52", "#3B82F6"),
        ("03", "5.3 Product Library",     "vw_submit_53", "#8B5CF6"),
        ("04", "5.4 Citrix Upload",       "vw_submit_54", "#F59E0B"),
        ("05", "Range Sheet Non-SSPOG",   "vw_submit_ns", "#E11D48"),
        ("06", "Range Sheet SSPOG",       "vw_submit_ss", "#7C3AED"),
        ("07", "StoreApply POG",          "vw_submit_sa", "#0891B2"),
    ]

    _with_data = [
        (n, t, k, c) for n, t, k, c in OUTPUTS
        if st.session_state.get(k) is not None
        and not st.session_state[k].empty
    ]

    # ── Select All toggle ─────────────────────────────────────────────────────
    st.markdown("""
<style>
div[data-testid="stMarkdown"]:has(.sa-all-selected) ~ div[data-testid="stButton"] button {
    background-color: #2BBFA4 !important; color: #ffffff !important;
    border-color: #2BBFA4 !important; font-weight: 600 !important;
}
div[data-testid="stMarkdown"]:has(.sa-all-selected) ~ div[data-testid="stButton"] button:hover {
    background-color: #23A892 !important; border-color: #23A892 !important;
}
div[data-testid="stMarkdown"]:has(.sa-none-selected) ~ div[data-testid="stButton"] button {
    background-color: #F5F0EA !important; color: #555555 !important;
    border-color: #D0CAC2 !important;
}
div[data-testid="stMarkdown"]:has(.sa-none-selected) ~ div[data-testid="stButton"] button:hover {
    background-color: #EDE8E2 !important; border-color: #B0A8A0 !important;
}
</style>""", unsafe_allow_html=True)

    _all_selected = bool(_with_data) and all(
        st.session_state.get(f"rpt_chk_{n}", False) for (n, _, _, _) in _with_data
    )
    _sa_col, _ = st.columns([1.3, 6.7])
    with _sa_col:
        st.markdown(
            f'<div class="{"sa-all-selected" if _all_selected else "sa-none-selected"}"></div>',
            unsafe_allow_html=True,
        )
        if st.button("☑ Select All", key="rpt_select_all", use_container_width=True):
            _new_val = not _all_selected
            for (n, _, _, _) in _with_data:
                st.session_state[f"rpt_chk_{n}"] = _new_val
            st.rerun()

    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

    _sel_nums    = []
    _sel_outputs = []   # (num, title, df_out) for selected outputs

    for num, title, key, color in OUTPUTS:
        df_out   = st.session_state.get(key)
        has_data = df_out is not None and not df_out.empty
        _prev_key = f"prev_{num}_show"
        _chk_key  = f"rpt_chk_{num}"
        if _prev_key not in st.session_state:
            st.session_state[_prev_key] = False
        _is_open = st.session_state[_prev_key]

        if has_data:
            _rows_txt = f"{len(df_out):,} rows · {len(df_out.columns)} cols"
            _radius   = "14px 14px 0 0" if _is_open else "14px"

            _chk_col, _card_col = st.columns([1, 14])
            with _chk_col:
                st.markdown("<div style='padding-top:18px;'>", unsafe_allow_html=True)
                _is_selected = st.checkbox("", key=_chk_key, label_visibility="collapsed")
                st.markdown("</div>", unsafe_allow_html=True)
            if _is_selected:
                _sel_nums.append(num)
                _sel_outputs.append((num, title, df_out))

            with _card_col:
                # ── Card header ───────────────────────────────────────────────
                st.markdown(f"""
<div style="background:#fff;border-radius:{_radius};border:1px solid #E8E3DC;
            border-bottom:{'none' if _is_open else '1px solid #E8E3DC'};
            border-left:4px solid {color};padding:16px 20px;">
    <div style="display:flex;align-items:center;gap:14px;">
        <div style="width:34px;height:34px;background:{color}22;border-radius:8px;
                    display:flex;align-items:center;justify-content:center;
                    font-size:12px;font-weight:800;color:{color};flex-shrink:0;">{num}</div>
        <div>
            <div style="font-size:14px;font-weight:700;color:#1A1A1A;">{title}</div>
            <div style="font-size:11px;color:#999;margin-top:3px;">
                Submitted from View Data &nbsp;·&nbsp; {_rows_txt}
            </div>
        </div>
    </div>
</div>""", unsafe_allow_html=True)

                # ── Buttons ───────────────────────────────────────────────────
                _b1, _b2, _b3, _ = st.columns([1.1, 1, 1, 2.9])
                with _b1:
                    if st.button(
                        "✕ Close Preview" if _is_open else "👁 Preview",
                        key=f"prev_btn_{num}",
                        use_container_width=True,
                        type="primary" if _is_open else "secondary",
                    ):
                        st.session_state[_prev_key] = not _is_open
                        st.rerun()
                with _b2:
                    st.download_button(
                        "⬇️ .xlsx", df_to_xlsx_bytes(df_out),
                        file_name=f"output_{num}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{num}x", use_container_width=True,
                    )
                with _b3:
                    st.download_button(
                        "⬇️ .csv", df_to_csv_bytes(df_out),
                        file_name=f"output_{num}.csv",
                        mime="text/csv",
                        key=f"dl_{num}c", use_container_width=True,
                    )

                # ── Preview panel — READ-ONLY HTML table ──────────────────────
                if _is_open:
                    _all_cols = list(df_out.columns)
                    _vis_key  = f"prev_vis_{num}"
                    _zoom_key = f"prev_zoom_{num}"
                    _full_key = f"prev_full_{num}"
                    if _vis_key  not in st.session_state: st.session_state[_vis_key]  = _all_cols[:]
                    if _zoom_key not in st.session_state: st.session_state[_zoom_key] = 100
                    if _full_key not in st.session_state: st.session_state[_full_key] = False

                    _zoom    = st.session_state[_zoom_key]
                    _is_full = st.session_state[_full_key]
                    _tbl_h   = 580 if _is_full else 320

                    _tc1, _tc2, _tc3, _tc4, _tc5 = st.columns([4.6, 0.55, 0.75, 0.55, 0.65])
                    with _tc1:
                        _vis_sel = st.multiselect(
                            "_vis", options=_all_cols,
                            default=[c for c in st.session_state[_vis_key] if c in _all_cols] or _all_cols,
                            key=f"vis_{num}", placeholder="👁  Show / hide columns…",
                            label_visibility="collapsed",
                        )
                        st.session_state[_vis_key] = _vis_sel if _vis_sel else _all_cols[:]
                    with _tc2:
                        if st.button("A−", key=f"zo_{num}", use_container_width=True, help="Zoom out"):
                            st.session_state[_zoom_key] = max(70, _zoom - 15); st.rerun()
                    with _tc3:
                        st.markdown(
                            f"<p style='text-align:center;font-size:11px;color:#888;"
                            f"margin:0;padding-top:9px;'>{_zoom}%</p>",
                            unsafe_allow_html=True,
                        )
                    with _tc4:
                        if st.button("A+", key=f"zi_{num}", use_container_width=True, help="Zoom in"):
                            st.session_state[_zoom_key] = min(160, _zoom + 15); st.rerun()
                    with _tc5:
                        if st.button("⛶" if not _is_full else "⊠", key=f"fs_{num}",
                                     use_container_width=True):
                            st.session_state[_full_key] = not _is_full; st.rerun()

                    _vis_cols = st.session_state[_vis_key]
                    _s   = _zoom / 100
                    _fz  = round(12 * _s, 1)
                    _pad = f"{round(8*_s)}px {round(14*_s)}px"
                    _hpad= f"{round(10*_s)}px {round(16*_s)}px"
                    _prev_n = min(50, len(df_out))

                    _th = "".join(
                        f'<th style="background:#2BBFA4;color:#fff;font-weight:700;'
                        f'padding:{_hpad};white-space:nowrap;text-align:left;'
                        f'font-size:{_fz}px;letter-spacing:.03em;'
                        f'border-right:1px solid rgba(255,255,255,0.25);'
                        f'position:sticky;top:0;z-index:2;">{c}</th>'
                        for c in _vis_cols
                    )
                    _tbody = ""
                    for _ri, (_, _row) in enumerate(df_out.head(_prev_n).iterrows()):
                        _bg = "#fff" if _ri % 2 == 0 else "#F4FBF9"
                        _tds = "".join(
                            f'<td style="padding:{_pad};font-size:{_fz}px;color:#1A1A1A;'
                            f'white-space:nowrap;border-right:1px solid #EEE;'
                            f'border-bottom:1px solid #F0EBE3;">'
                            f'{str(_row[c]) if _row[c] is not None and str(_row[c]) not in ("nan","None","<NA>") else ""}</td>'
                            for c in _vis_cols
                        )
                        _tbody += f'<tr style="background:{_bg};">{_tds}</tr>'

                    st.markdown(f"""
<div style="background:#F8FFFE;border:1px solid #E8E3DC;border-top:none;
            border-radius:0 0 14px 14px;padding:10px 14px 14px;">
    <div style="display:flex;justify-content:space-between;align-items:center;
                margin-bottom:8px;">
        <span style="font-size:11px;font-weight:700;color:#2BBFA4;
                     text-transform:uppercase;letter-spacing:.06em;">
            Preview — {title}
        </span>
        <span style="font-size:11px;color:#888;">
            First {_prev_n:,} of {len(df_out):,} rows &nbsp;·&nbsp;
            {len(_vis_cols)}/{len(_all_cols)} cols &nbsp;·&nbsp;
            <em>read-only</em>
        </span>
    </div>
    <div style="overflow-x:auto;overflow-y:auto;max-height:{_tbl_h}px;
                border-radius:8px;border:1px solid #E8E3DC;">
        <table style="border-collapse:collapse;width:100%;min-width:400px;">
            <thead><tr>{_th}</tr></thead>
            <tbody>{_tbody}</tbody>
        </table>
    </div>
</div>""", unsafe_allow_html=True)

            add_audit(f"View Output {num}", title)

        else:
            # ── Placeholder card (not yet submitted) ──────────────────────────
            st.markdown(f"""
<div style="background:#FAFAFA;border-radius:14px;border:1px solid #EAEAEA;
            border-left:4px solid {color}55;padding:16px 20px;">
    <div style="display:flex;align-items:center;gap:14px;">
        <div style="width:34px;height:34px;background:{color}11;border-radius:8px;
                    display:flex;align-items:center;justify-content:center;
                    font-size:12px;font-weight:800;color:{color}66;flex-shrink:0;">{num}</div>
        <div>
            <div style="font-size:14px;font-weight:600;color:#BBBBBB;">{title}</div>
            <div style="font-size:11px;color:#CCCCCC;margin-top:3px;">
                Not submitted yet — go to View Data and click ✅ Submit to Report
            </div>
        </div>
    </div>
</div>""", unsafe_allow_html=True)

        st.markdown("<div style='margin-bottom:14px;'></div>", unsafe_allow_html=True)

    # ── Bulk export ───────────────────────────────────────────────────────────
    if _sel_outputs:
        st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
        _n  = len(_sel_outputs)
        _ts = datetime.now().strftime("%Y%m%d_%H%M")
        st.markdown(
            f"<div style='font-size:12px;font-weight:700;color:#555;margin-bottom:8px;'>"
            f"Export selected ({_n} file{'s' if _n > 1 else ''}) as:</div>",
            unsafe_allow_html=True,
        )
        _ec1, _ec2, _ec3 = st.columns(3)

        def _zip_bytes(fmt: str, byte_fn) -> bytes:
            _buf = io.BytesIO()
            with zipfile.ZipFile(_buf, "w", zipfile.ZIP_DEFLATED) as _zf:
                for _sn, _, _dfo in _sel_outputs:
                    _zf.writestr(f"output_{_sn}.{fmt}", byte_fn(_dfo))
            return _buf.getvalue()

        with _ec1:
            if _n == 1:
                _sn, _, _dfo = _sel_outputs[0]
                st.download_button("⬇️ .csv", df_to_csv_bytes(_dfo),
                    file_name=f"output_{_sn}_{_ts}.csv", mime="text/csv",
                    use_container_width=True, key="dl_sel_csv")
            else:
                st.download_button("⬇️ .csv (zip)", _zip_bytes("csv", df_to_csv_bytes),
                    file_name=f"range_outputs_{_ts}.zip", mime="application/zip",
                    use_container_width=True, key="dl_sel_csv")
        with _ec2:
            if _n == 1:
                _sn, _, _dfo = _sel_outputs[0]
                st.download_button("⬇️ .xlsx", df_to_xlsx_bytes(_dfo),
                    file_name=f"output_{_sn}_{_ts}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True, key="dl_sel_xlsx")
            else:
                st.download_button("⬇️ .xlsx (zip)", _zip_bytes("xlsx", df_to_xlsx_bytes),
                    file_name=f"range_outputs_{_ts}.zip", mime="application/zip",
                    use_container_width=True, key="dl_sel_xlsx")
        with _ec3:
            if _n == 1:
                _sn, _, _dfo = _sel_outputs[0]
                st.download_button("⬇️ .xls", df_to_xlsx_bytes(_dfo),
                    file_name=f"output_{_sn}_{_ts}.xls", mime="application/vnd.ms-excel",
                    use_container_width=True, key="dl_sel_xls")
            else:
                st.download_button("⬇️ .xls (zip)", _zip_bytes("xls", df_to_xlsx_bytes),
                    file_name=f"range_outputs_{_ts}.zip", mime="application/zip",
                    use_container_width=True, key="dl_sel_xls")

# ══════════════════════════════════════════════════════════════════════════════
# Tab 2: Summary & Analytics (requires uploaded range file)
# ══════════════════════════════════════════════════════════════════════════════
with rt2:
    st.markdown("""
<div style="font-size:18px;font-weight:700;color:#1A1A1A;margin-bottom:20px;">
    Summary & Analytics
</div>""", unsafe_allow_html=True)

    if merged is None:
        st.markdown("""
<div style="background:#fff;border-radius:16px;border:1px solid #E8E3DC;
            padding:40px;text-align:center;">
    <div style="font-size:36px;margin-bottom:12px;">📈</div>
    <div style="font-size:15px;font-weight:600;color:#1A1A1A;margin-bottom:6px;">No range file uploaded</div>
    <div style="font-size:12px;color:#999;">Go to My Files and upload a range file to see analytics.</div>
</div>""", unsafe_allow_html=True)
    else:
        SUMMARY_VIEWS = [
            "Summary by Status", "Summary by Cluster POG",
            "Summary by POG",    "Summary by Store",
            "Financial by Status", "Financial by POG",
        ]
        sel = st.selectbox("Select Summary View", SUMMARY_VIEWS)

        col_map = {
            "Status":  find_col(merged, ["status"]),
            "Cluster": find_col(merged, ["cluster"]),
            "POG":     find_col(merged, ["pog"]),
            "Store":   find_col(merged, ["store"]),
        }
        group_key = next((k for k in col_map if k.lower() in sel.lower()), None)
        group_col = col_map.get(group_key) if group_key else None
        val_cols  = [c for c in merged.columns
                     if pd.api.types.is_numeric_dtype(merged[c])][:3]

        if group_col:
            try:
                if val_cols:
                    summary_df = merged.groupby(group_col)[val_cols].sum().reset_index()
                    summary_df["Count"] = merged.groupby(group_col).size().values
                else:
                    summary_df = merged[group_col].value_counts().reset_index()
                    summary_df.columns = [group_col, "Count"]

                st.markdown("""
<div style="background:#fff;border-radius:14px;border:1px solid #E8E3DC;
            padding:16px;overflow:hidden;">""", unsafe_allow_html=True)
                st.dataframe(summary_df, use_container_width=True, hide_index=True, height=280)
                st.markdown("</div>", unsafe_allow_html=True)

                num_cols = [c for c in summary_df.columns
                            if c != group_col and pd.api.types.is_numeric_dtype(summary_df[c])][:2]
                if num_cols:
                    st.bar_chart(summary_df.set_index(group_col)[num_cols], height=220)

                st.download_button(
                    f"⬇️ Export: {sel}", df_to_xlsx_bytes(summary_df),
                    f"{sel.replace(' ', '_')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Cannot create summary: {e}")
        else:
            st.markdown(f"""
<div style="background:#FFF9C4;border-radius:12px;padding:16px 20px;border:1px solid #F0E68C;">
    <span style="font-size:13px;color:#7A6000;">
        ⚠️ Column matching <strong>{sel}</strong> not found in the uploaded data.
    </span>
</div>""", unsafe_allow_html=True)

render_page_nav("report")
