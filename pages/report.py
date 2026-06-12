"""Report page — Range Execution & Summary Analytics."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
import pandas as pd
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
# Tab 1: Range Execution
# ══════════════════════════════════════════════════════════════════════════════
with rt1:
    st.markdown("""
<div style="margin-bottom:4px;">
    <span style="font-size:18px;font-weight:700;color:#1A1A1A;">Range Execution</span>
</div>
<div style="font-size:13px;color:#888;margin-bottom:20px;">
    Split PPOG → SSPOG / Non-SSPOG, then export 4 output files.
</div>
""", unsafe_allow_html=True)

    # ── PPOG → SSPOG Split card ───────────────────────────────────────────────
    pog_col = find_col(merged, ["pog", "ppog", "sspog", "cluster"])

    with st.container():
        st.markdown("""
<div style="background:#fff;border-radius:16px;border:1px solid #E8E3DC;
            padding:22px 24px;">
    <div style="font-size:15px;font-weight:700;color:#1A1A1A;margin-bottom:16px;">
        PPOG → SSPOG / Non-SSPOG Split
    </div>""", unsafe_allow_html=True)

        if pog_col:
            st.markdown(f"""
<div style="background:#E8F8F5;border-radius:10px;padding:11px 16px;margin-bottom:16px;">
    <span style="font-size:13px;color:#555;">Detected column: </span>
    <span style="font-size:13px;font-weight:700;color:#2BBFA4;">{pog_col}</span>
</div>""", unsafe_allow_html=True)
        else:
            pog_col = st.selectbox("Select POG / Cluster column",
                                   ["— none —"] + list(merged.columns),
                                   label_visibility="visible")
            if pog_col == "— none —":
                pog_col = None

        if pog_col:
            pog_vals   = sorted(merged[pog_col].dropna().astype(str).unique().tolist())
            sspog_vals = [v for v in pog_vals
                         if "ss" in v.lower() and "non" not in v.lower()]

            if "sspog_sel" not in st.session_state:
                st.session_state.sspog_sel = set(sspog_vals)

            st.markdown('<div style="font-size:13px;color:#555;margin-bottom:10px;">Select values = SSPOG:</div>',
                        unsafe_allow_html=True)

            # Pill toggle buttons
            pill_cols = st.columns(min(len(pog_vals), 6))
            for idx, val in enumerate(pog_vals):
                col_idx = idx % len(pill_cols)
                with pill_cols[col_idx]:
                    is_sel = val in st.session_state.sspog_sel
                    label  = f"✓ {val}" if is_sel else val
                    btn_style = (
                        "background:#2BBFA4;color:#fff;border:none;"
                        if is_sel else
                        "background:#fff;color:#555;border:1.5px solid #D0CAC2;"
                    )
                    if st.button(label, key=f"pill_{val}",
                                 use_container_width=True):
                        if is_sel:
                            st.session_state.sspog_sel.discard(val)
                        else:
                            st.session_state.sspog_sel.add(val)
                        st.rerun()

            sspog_sel = list(st.session_state.sspog_sel)
            sspog_df  = merged[merged[pog_col].astype(str).isin(sspog_sel)]
            non_df    = merged[~merged[pog_col].astype(str).isin(sspog_sel)]
        else:
            sspog_df = merged.iloc[0:0]
            non_df   = merged

        st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)
        _mc1, _mc2 = st.columns(2)
        with _mc1:
            st.markdown(f"""
<div style="background:#E8F8F5;border-radius:14px;padding:20px;text-align:center;">
    <div style="font-size:11px;font-weight:700;color:#2BBFA4;text-transform:uppercase;
                letter-spacing:0.06em;margin-bottom:8px;">SSPOG</div>
    <div style="font-size:36px;font-weight:800;color:#1A1A1A;">{len(sspog_df):,}</div>
</div>""", unsafe_allow_html=True)
        with _mc2:
            st.markdown(f"""
<div style="background:#F5F0EA;border-radius:14px;padding:20px;text-align:center;">
    <div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;
                letter-spacing:0.06em;margin-bottom:8px;">NON-SSPOG</div>
    <div style="font-size:36px;font-weight:800;color:#1A1A1A;">{len(non_df):,}</div>
</div>""", unsafe_allow_html=True)

        st.markdown("</div>", unsafe_allow_html=True)

    # ── Output Files ──────────────────────────────────────────────────────────
    st.markdown("""
<div style="font-size:16px;font-weight:700;color:#1A1A1A;margin:28px 0 16px;">
    Output Files
</div>""", unsafe_allow_html=True)

    CARD_COLORS = ["#2BBFA4", "#3B82F6", "#8B5CF6", "#F59E0B"]
    OUTPUTS = [
        ("01", "Range by Item by Store",
         "Full range allocation — one row per item per store", merged),
        ("02", "Item Store · Item Status · Forecast → Supply Chain",
         "Status and forecast formatted for supply chain", merged),
        ("03", "Item No. · MOD-Store · Cluster",
         "Item number mapped to MOD-Store and cluster", merged),
        ("04", "SSPOG File — Citrix Upload",
         "SSPOG items formatted for Citrix upload", sspog_df),
    ]

    for (num, title, desc, df_out), color in zip(OUTPUTS, CARD_COLORS):
        rows_txt = f"· {len(df_out):,} rows"
        st.markdown(f"""
<div style="background:#fff;border-radius:14px;border:1px solid #E8E3DC;
            border-left:4px solid {color};padding:16px 20px;
            display:flex;align-items:center;justify-content:space-between;
            margin-bottom:10px;flex-wrap:wrap;gap:12px;">
    <div style="display:flex;align-items:center;gap:14px;">
        <div style="width:32px;height:32px;background:{color}20;border-radius:8px;
                    display:flex;align-items:center;justify-content:center;
                    font-size:12px;font-weight:800;color:{color};flex-shrink:0;">{num}</div>
        <div>
            <div style="font-size:14px;font-weight:600;color:#1A1A1A;">{title}</div>
            <div style="font-size:12px;color:#999;margin-top:2px;">{desc} {rows_txt}</div>
        </div>
    </div>
</div>""", unsafe_allow_html=True)

        _dcols = st.columns([1, 1, 4])
        with _dcols[0]:
            st.download_button(
                "⬇️ .xlsx", df_to_xlsx_bytes(df_out),
                file_name=f"output_{num}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{num}x", use_container_width=True,
            )
        with _dcols[1]:
            st.download_button(
                "⬇️ .csv", df_to_csv_bytes(df_out),
                file_name=f"output_{num}.csv",
                mime="text/csv",
                key=f"dl_{num}c", use_container_width=True,
            )
        add_audit(f"Export Output {num}", title)

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
