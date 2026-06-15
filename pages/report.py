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
if merged is None:
    st.markdown("""
<div style="background:#fff;border-radius:16px;border:1px solid #E8E3DC;
            padding:60px;text-align:center;">
    <div style="font-size:48px;margin-bottom:16px;">📋</div>
    <div style="font-size:16px;font-weight:600;color:#1A1A1A;margin-bottom:8px;">No data yet</div>
    <div style="font-size:13px;color:#999;">Go to My Files and upload a file first.</div>
</div>
""", unsafe_allow_html=True)
    st.stop()

rt1, rt2 = st.tabs(["📦  Range Execution", "📈  Summary & Analytics"])

# ══════════════════════════════════════════════════════════════════════════════
# Tab 1: Range Execution — Output Files only
# ══════════════════════════════════════════════════════════════════════════════
with rt1:
    st.markdown("""
<div style="margin-bottom:24px;">
    <div style="font-size:18px;font-weight:700;color:#1A1A1A;margin-bottom:4px;">
        Range Execution
    </div>
    <div style="font-size:13px;color:#888;">
        Preview and export output files for range execution.
    </div>
</div>""", unsafe_allow_html=True)

    # ── Column selection helper ───────────────────────────────────────────────
    def _select_cols(df: pd.DataFrame, specs: list) -> pd.DataFrame:
        def _norm(s):
            return str(s).lower().replace("_", "").replace(" ", "").strip()

        result = {}
        for out_name, candidates in specs:
            src = None
            for cand in candidates:
                src = next(
                    (c for c in df.columns if c.lower().strip() == cand.lower().strip()),
                    None,
                )
                if not src:
                    src = next(
                        (c for c in df.columns if _norm(c) == _norm(cand)),
                        None,
                    )
                if src:
                    break
            result[out_name] = (
                df[src].reset_index(drop=True) if src
                else pd.Series([None] * len(df), name=out_name)
            )
        return pd.DataFrame(result)

    _COLS_01 = [
        ("store_no",           ["store_no"]),
        ("store_name",         ["store_name"]),
        ("store_Format",       ["store_Format", "store_format", "storeformat"]),
        ("ID",                 ["ID"]),
        ("ProductDescription", ["ProductDescription", "product_description"]),
        ("UPC",                ["UPC"]),
        ("POGName",            ["POGName", "pog_name", "pog"]),
        ("ClusterName",        ["ClusterName", "Range_Class", "range_class", "cluster"]),
    ]
    _COLS_02 = [
        ("ID",                       ["ID"]),
        ("ProductDescription",       ["ProductDescription", "product_description"]),
        ("store_no",                 ["store_no"]),
        ("store_name",               ["store_name"]),
        ("POG_STATUS",               ["POG_STATUS", "pog_status", "status"]),
        ("Display Group",            ["Display Group", "display_group", "DisplayGroup"]),
        ("Display group desc",       ["Display group desc", "display_group_desc", "DisplayGroupDesc"]),
        ("ForecastSales",            ["ForecastSales", "forecast_sales"]),
        ("TH_Tot_Sales_Value_52WK",  ["TH_Tot_Sales_Value_52WK", "th_tot_sales_value_52wk"]),
        ("TH_Tot_Sales_Volume_52WK", ["TH_Tot_Sales_Volume_52WK", "th_tot_sales_volume_52wk"]),
        ("DaysSupply",               ["DaysSupply", "days_supply"]),
        ("MaxDOS",                   ["MaxDOS", "max_dos"]),
    ]
    _COLS_03 = [
        ("ID",                     ["ID"]),
        ("ProductDescription",     ["ProductDescription", "product_description"]),
        ("ACTUAL_MOD",             ["ACTUAL_MOD", "actual_mod"]),
        ("no_of_mod",              ["no_of_mod", "NoOfMods", "noofmods"]),
        ("Fixture_Code",           ["Fixture_Code", "fixture_code"]),
        ("Store_Cluster",          ["Store_Cluster", "store_cluster"]),
        ("POG_Cluster",            ["POG_Cluster", "pog_cluster"]),
        ("Property_Store_Cluster", ["Property_Store_Cluster", "property_store_cluster"]),
        ("store_no",               ["store_no"]),
        ("store_name",             ["store_name"]),
    ]
    _COLS_04 = [
        ("POGName",               ["POGName", "pog_name", "pog"]),
        ("POG_Store.POGName",     ["POG_Store.POGName", "pog_store_pogname"]),
        ("POG_WIDTH",             ["POG_WIDTH", "pog_width"]),
        ("POG_HEIGHT",            ["POG_HEIGHT", "pog_height"]),
        ("POG_DEPTH",             ["POG_DEPTH", "pog_depth"]),
        ("Total_linear",          ["Total_linear", "total_linear"]),
        ("SQM",                   ["SQM", "sqm"]),
        ("Department Code&Desc",  ["Department Code&Desc", "department_code_desc", "Department", "department"]),
        ("Division",              ["Division", "division"]),
        ("Subcategory",           ["Subcategory", "subcategory"]),
        ("store_no",              ["store_no"]),
        ("store_name",            ["store_name"]),
        ("FP_Fixturename",        ["FP_Fixturename", "fp_fixturename"]),
        ("FP_status",             ["FP_status", "fp_status"]),
        ("number_of_product",     ["number_of_product"]),
        ("Capacity",              ["Capacity", "capacity"]),
        ("CaseTotalNumber",       ["CaseTotalNumber", "casetotalnumber"]),
    ]

    # Silent auto-detect SSPOG split — used for output 04
    _pog_col = find_col(merged, ["pog", "ppog", "sspog", "cluster"])
    if _pog_col:
        _pog_vals   = sorted(merged[_pog_col].dropna().astype(str).unique().tolist())
        _sspog_vals = [v for v in _pog_vals
                       if "ss" in v.lower() and "non" not in v.lower()]
        sspog_df    = merged[merged[_pog_col].astype(str).isin(_sspog_vals)]
    else:
        sspog_df = merged.iloc[0:0]

    CARD_COLORS = ["#2BBFA4", "#3B82F6", "#8B5CF6", "#F59E0B"]
    OUTPUTS = [
        ("01", "Range by Item by Store",
         "Full range allocation — one row per item per store",
         _select_cols(merged, _COLS_01)),
        ("02", "Item Store · Item Status · Forecast → Supply Chain",
         "Status and forecast formatted for supply chain",
         _select_cols(merged, _COLS_02)),
        ("03", "Item No. · MOD-Store · Cluster",
         "Item number mapped to MOD-Store and cluster",
         _select_cols(merged, _COLS_03)),
        ("04", "SSPOG File — Citrix Upload",
         "SSPOG items formatted for Citrix upload",
         _select_cols(merged, _COLS_04)),
    ]

    # ── Select All / Deselect All controls ───────────────────────────────────
    _sa_col, _da_col, _ = st.columns([1.2, 1.4, 5])
    with _sa_col:
        if st.button("☑ Select All", key="rpt_select_all", use_container_width=True):
            for (_n, _, _, _) in OUTPUTS:
                st.session_state[f"rpt_chk_{_n}"] = True
            st.rerun()
    with _da_col:
        if st.button("□ Deselect All", key="rpt_desel_all", use_container_width=True):
            for (_n, _, _, _) in OUTPUTS:
                st.session_state[f"rpt_chk_{_n}"] = False
            st.rerun()

    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

    _sel_nums = []

    for (num, title, desc, df_out), color in zip(OUTPUTS, CARD_COLORS):
        _prev_key = f"prev_{num}_show"
        _chk_key  = f"rpt_chk_{num}"
        if _prev_key not in st.session_state:
            st.session_state[_prev_key] = False

        _is_open  = st.session_state[_prev_key]
        _rows_txt = f"{len(df_out):,} rows · {len(df_out.columns)} cols"
        _radius   = "14px 14px 0 0" if _is_open else "14px"

        # ── Checkbox + card row ───────────────────────────────────────────────
        _chk_col, _card_col = st.columns([1, 14])
        with _chk_col:
            st.markdown("<div style='padding-top:18px;'>", unsafe_allow_html=True)
            _is_selected = st.checkbox("", key=_chk_key, label_visibility="collapsed")
            st.markdown("</div>", unsafe_allow_html=True)
        if _is_selected:
            _sel_nums.append(num)

        with _card_col:
            # ── Card header ──────────────────────────────────────────────────
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
            <div style="font-size:11px;color:#999;margin-top:3px;">{desc} &nbsp;·&nbsp; {_rows_txt}</div>
        </div>
    </div>
</div>""", unsafe_allow_html=True)

            # ── Action buttons ────────────────────────────────────────────────
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

            # ── Preview panel ─────────────────────────────────────────────────
            if _is_open:
                _prev_n = min(50, len(df_out))
                _cols = list(df_out.columns)
                _th = "".join(
                    f'<th style="background:#2BBFA4;color:#fff;font-weight:700;'
                    f'padding:10px 16px;white-space:nowrap;text-align:left;'
                    f'font-size:12px;letter-spacing:.03em;'
                    f'border-right:1px solid rgba(255,255,255,0.25);'
                    f'position:sticky;top:0;z-index:2;">{c}</th>'
                    for c in _cols
                )
                _tbody = ""
                for _ri, (_, _row) in enumerate(df_out.head(_prev_n).iterrows()):
                    _bg = "#fff" if _ri % 2 == 0 else "#F4FBF9"
                    _tds = "".join(
                        f'<td style="padding:8px 16px;font-size:12px;color:#1A1A1A;'
                        f'white-space:nowrap;border-right:1px solid #EEE;'
                        f'border-bottom:1px solid #F0EBE3;">'
                        f'{str(_v) if _v is not None and str(_v) not in ("nan","None") else ""}</td>'
                        for _v in _row
                    )
                    _tbody += f'<tr style="background:{_bg};">{_tds}</tr>'
                st.markdown(f"""
<div style="background:#F8FFFE;border:1px solid #E8E3DC;border-top:none;
            border-radius:0 0 14px 14px;padding:10px 14px 14px;">
    <div style="display:flex;justify-content:space-between;align-items:center;
                margin-bottom:8px;">
        <span style="font-size:11px;font-weight:700;color:#2BBFA4;
                     text-transform:uppercase;letter-spacing:.06em;">Preview — {title}</span>
        <span style="font-size:11px;color:#888;">
            First {_prev_n:,} of {len(df_out):,} rows
        </span>
    </div>
    <div style="overflow-x:auto;overflow-y:auto;max-height:320px;
                border-radius:8px;border:1px solid #E8E3DC;">
        <table style="border-collapse:collapse;width:100%;min-width:400px;">
            <thead><tr>{_th}</tr></thead>
            <tbody>{_tbody}</tbody>
        </table>
    </div>
</div>""", unsafe_allow_html=True)

        add_audit(f"View Output {num}", title)
        st.markdown("<div style='margin-bottom:14px;'></div>", unsafe_allow_html=True)

    # ── Bulk export — csv / xlsx / xls for all selected files ────────────────
    if _sel_nums:
        st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
        _sel_outputs = [
            (num, title, desc, df_out)
            for (num, title, desc, df_out), _ in zip(OUTPUTS, CARD_COLORS)
            if num in _sel_nums
        ]
        _n = len(_sel_outputs)
        _ts = datetime.now().strftime("%Y%m%d_%H%M")

        st.markdown(
            f"<div style='font-size:12px;font-weight:700;color:#555;margin-bottom:8px;'>"
            f"Export selected ({_n} file{'s' if _n > 1 else ''}) as:</div>",
            unsafe_allow_html=True,
        )

        _ec1, _ec2, _ec3 = st.columns(3)

        def _zip_format(fmt: str, byte_fn) -> bytes:
            _buf = io.BytesIO()
            with zipfile.ZipFile(_buf, "w", zipfile.ZIP_DEFLATED) as _zf:
                for _num, _, _, _dfo in _sel_outputs:
                    _zf.writestr(f"output_{_num}.{fmt}", byte_fn(_dfo))
            return _buf.getvalue()

        with _ec1:
            if _n == 1:
                _num, _, _, _dfo = _sel_outputs[0]
                _data_csv = df_to_csv_bytes(_dfo)
                _fname_csv = f"output_{_num}_{_ts}.csv"
            else:
                _data_csv  = _zip_format("csv", df_to_csv_bytes)
                _fname_csv = f"range_outputs_{_ts}_csv.zip"
            st.download_button(
                "⬇️ .csv",
                data=_data_csv,
                file_name=_fname_csv,
                mime="text/csv" if _n == 1 else "application/zip",
                use_container_width=True,
                key="dl_sel_csv",
            )

        with _ec2:
            if _n == 1:
                _num, _, _, _dfo = _sel_outputs[0]
                _data_xlsx = df_to_xlsx_bytes(_dfo)
                _fname_xlsx = f"output_{_num}_{_ts}.xlsx"
            else:
                _data_xlsx  = _zip_format("xlsx", df_to_xlsx_bytes)
                _fname_xlsx = f"range_outputs_{_ts}_xlsx.zip"
            st.download_button(
                "⬇️ .xlsx",
                data=_data_xlsx,
                file_name=_fname_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if _n == 1 else "application/zip",
                use_container_width=True,
                key="dl_sel_xlsx",
            )

        with _ec3:
            if _n == 1:
                _num, _, _, _dfo = _sel_outputs[0]
                _data_xls = df_to_xlsx_bytes(_dfo)
                _fname_xls = f"output_{_num}_{_ts}.xls"
            else:
                _data_xls  = _zip_format("xls", df_to_xlsx_bytes)
                _fname_xls = f"range_outputs_{_ts}_xls.zip"
            st.download_button(
                "⬇️ .xls",
                data=_data_xls,
                file_name=_fname_xls,
                mime="application/vnd.ms-excel" if _n == 1 else "application/zip",
                use_container_width=True,
                key="dl_sel_xls",
            )

# ══════════════════════════════════════════════════════════════════════════════
# Tab 2: Summary & Analytics
# ══════════════════════════════════════════════════════════════════════════════
with rt2:
    st.markdown("""
<div style="font-size:18px;font-weight:700;color:#1A1A1A;margin-bottom:20px;">
    Summary & Analytics
</div>""", unsafe_allow_html=True)

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
