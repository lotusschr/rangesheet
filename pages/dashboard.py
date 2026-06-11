"""Dashboard page — Range Architecture summary table + charts."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
import pandas as pd
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar,
    STATUS_COLORS, find_col,
)

inject_css()
init_session_state()
render_sidebar("dashboard")
render_topbar("Dashboard")

merged = st.session_state.merged_df
if merged is None:
    st.warning("No data — upload files on **My Files** first.")
    st.stop()

try:
    import plotly.graph_objects as go
    import plotly.express as px
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

# ── Helpers ────────────────────────────────────────────────────────────────────
def _find(df, *kw_sets):
    for kws in kw_sets:
        for c in df.columns:
            if all(k.lower() in c.lower() for k in kws):
                return c
    return None

def _safe_sum(df, col, mask=None):
    if not col or col not in df.columns:
        return 0.0
    sub = df[mask] if mask is not None else df
    return float(pd.to_numeric(sub[col], errors="coerce").fillna(0).sum())

def _fmt(v):
    try:
        return f"{int(v):,}" if v == int(v) else f"{v:,.1f}"
    except Exception:
        return str(v)

TYPE_COLOR = {
    "MAINTAIN":    "#2BBFA4", "DELETE SOME": "#E05555", "DELETE ALL": "#C0392B",
    "NEW SOME":    "#3B82F6", "NEWNEW":      "#8B5CF6", "NEW":        "#3B82F6",
    "DELETE":      "#E05555", "REVIEW":      "#E08A20",
}

# ── Detect type column ─────────────────────────────────────────────────────────
type_col = find_col(merged, ["type"])
if not type_col:
    for c in merged.columns:
        vals = merged[c].dropna().astype(str).str.upper().unique()
        if any(v in ["MAINTAIN","DELETE SOME","DELETE ALL","NEW SOME","NEWNEW"] for v in vals):
            type_col = c
            break

# ── Detect AS IS / TO BE columns ──────────────────────────────────────────────
as_is_cols = [c for c in merged.columns if any(k in str(c).upper() for k in ["AS IS","ASIS","AS_IS"])]
to_be_cols = [c for c in merged.columns if any(k in str(c).upper() for k in ["TO BE","TOBE","TO_BE"])]
asis_col   = as_is_cols[0] if as_is_cols else None
tobe_col   = to_be_cols[0] if to_be_cols else None

# ── Specific impact columns ────────────────────────────────────────────────────
_sale_asis_c    = _find(merged, ["as-is","sale"],    ["as is","sale"])
_sale_tobe_c    = _find(merged, ["to-be","sale"],    ["to be","sale"])
_margin_asis_c  = _find(merged, ["as-is","margin"],  ["as is","margin"])

# ══════════════════════════════════════════════════════════════════════════════
# RANGE ARCHITECTURE SUMMARY TABLE
# ══════════════════════════════════════════════════════════════════════════════
_arch_types = ["MAINTAIN", "DELETE SOME", "DELETE ALL", "NEW SOME", "NEWNEW"]
_arch_rows  = []

for _t in _arch_types:
    _msk = (merged[type_col].astype(str).str.strip().str.upper() == _t
            if type_col else pd.Series([False]*len(merged), index=merged.index))
    _cnt = int(_msk.sum())

    # SKU counts
    if asis_col:
        _asis_sku = int(_safe_sum(merged, asis_col, _msk))
    else:
        _asis_sku = _cnt if ("DELETE" in _t or _t == "MAINTAIN") else 0
    if tobe_col:
        _tobe_sku = int(_safe_sum(merged, tobe_col, _msk))
    else:
        _tobe_sku = _cnt if ("MAINTAIN" in _t or "NEW" in _t) else 0

    _sale_asis_v  = _safe_sum(merged, _sale_asis_c,   _msk)
    _sale_tobe_v  = _safe_sum(merged, _sale_tobe_c,   _msk)
    _margin_v     = _safe_sum(merged, _margin_asis_c,  _msk)
    _arch_rows.append({
        "type": _t, "asis_sku": _asis_sku, "tobe_sku": _tobe_sku,
        "sale_asis": _sale_asis_v, "sale_tobe": _sale_tobe_v,
        "sale_diff": _sale_tobe_v - _sale_asis_v, "margin_asis": _margin_v,
    })

_tot_asis_sku  = sum(r["asis_sku"]   for r in _arch_rows)
_tot_tobe_sku  = sum(r["tobe_sku"]   for r in _arch_rows)
_tot_sale_asis = sum(r["sale_asis"]  for r in _arch_rows)
_tot_sale_tobe = sum(r["sale_tobe"]  for r in _arch_rows)
_tot_sale_diff = _tot_sale_tobe - _tot_sale_asis
_tot_margin    = sum(r["margin_asis"] for r in _arch_rows)

_pct_sku  = f"{(_tot_tobe_sku / _tot_asis_sku - 1) * 100:.1f}%"  if _tot_asis_sku  else "0.0%"
_pct_sale = f"{(_tot_sale_diff / _tot_sale_asis * 100):.1f}%"    if _tot_sale_asis else "0.0%"

_tbody = ""
for _r in _arch_rows:
    _tc       = STATUS_COLORS.get(_r["type"], {"bg":"#F5F5F5","c":"#888"})
    _asis_bg  = "#E53935" if _r["type"] == "NEWNEW"     else "transparent"
    _tobe_bg  = "#E53935" if _r["type"] == "DELETE ALL" else "transparent"
    _asis_fc  = "#fff"    if _asis_bg != "transparent"  else "#1A1A1A"
    _tobe_fc  = "#fff"    if _tobe_bg != "transparent"  else "#1A1A1A"
    _tbody += (
        f'<tr style="border-bottom:1px solid #E0D9D2;">'
        f'<td style="padding:8px 14px;font-weight:600;color:{_tc["c"]};">{_r["type"]}</td>'
        f'<td style="padding:8px;text-align:center;background:{_asis_bg};color:{_asis_fc};">'
        f'{_fmt(_r["asis_sku"])}</td>'
        f'<td style="padding:8px;text-align:center;background:{_tobe_bg};color:{_tobe_fc};'
        f'border-right:2px solid #C8C0B8;">{_fmt(_r["tobe_sku"])}</td>'
        f'<td style="padding:8px;text-align:center;">{_fmt(_r["sale_asis"])}</td>'
        f'<td style="padding:8px;text-align:center;">{_fmt(_r["sale_tobe"])}</td>'
        f'<td style="padding:8px;text-align:center;border-right:2px solid #C8C0B8;">'
        f'{_fmt(_r["sale_diff"])}</td>'
        f'<td style="padding:8px;text-align:center;">{_fmt(_r["margin_asis"])}</td>'
        f'</tr>')

_tbody += (
    f'<tr style="background:#1C1C1E;color:#fff;font-weight:700;border-top:2px solid #555;">'
    f'<td style="padding:8px 14px;">TOTAL SKU</td>'
    f'<td style="padding:8px;text-align:center;">{_fmt(_tot_asis_sku)}</td>'
    f'<td style="padding:8px;text-align:center;border-right:2px solid #555;">{_fmt(_tot_tobe_sku)}</td>'
    f'<td style="padding:8px;text-align:center;">{_fmt(_tot_sale_asis)}</td>'
    f'<td style="padding:8px;text-align:center;">{_fmt(_tot_sale_tobe)}</td>'
    f'<td style="padding:8px;text-align:center;border-right:2px solid #555;">{_fmt(_tot_sale_diff)}</td>'
    f'<td style="padding:8px;text-align:center;">{_fmt(_tot_margin)}</td>'
    f'</tr>'
    f'<tr style="background:#F5F5F5;">'
    f'<td style="padding:6px 14px;color:#888;font-size:12px;">% Impact</td>'
    f'<td style="padding:6px 8px;"></td>'
    f'<td style="padding:6px 8px;text-align:center;color:#2BBFA4;font-weight:700;'
    f'font-size:12px;border-right:2px solid #C8C0B8;">{_pct_sku}</td>'
    f'<td style="padding:6px 8px;"></td><td style="padding:6px 8px;"></td>'
    f'<td style="padding:6px 8px;text-align:center;color:#2BBFA4;font-weight:700;'
    f'font-size:12px;border-right:2px solid #C8C0B8;">{_pct_sale}</td>'
    f'<td style="padding:6px 8px;"></td></tr>')

st.markdown(f"""
<div style="background:#fff;border-radius:10px;border:1px solid #E0D9D2;
     overflow:hidden;margin-bottom:24px;">
  <table style="width:100%;border-collapse:collapse;font-size:13px;">
    <thead>
      <tr style="background:#F2EDE8;">
        <th colspan="3" style="padding:10px 14px;text-align:center;border:1px solid #D0C8C0;
             font-weight:700;">Range architecture</th>
        <th colspan="3" style="padding:10px 14px;text-align:center;border:1px solid #D0C8C0;
             font-weight:700;">Sale Impact ( ex.vat) / Week<br>
          <span style="font-weight:400;font-size:11px;">calcualte from Mer Price</span></th>
        <th style="padding:10px 14px;text-align:center;border:1px solid #D0C8C0;
             font-weight:700;font-size:12px;">Margin Impact ( ex.vat) / Week<br>
          <span style="font-weight:400;font-size:10px;">calcualte from EDLP Price</span></th>
      </tr>
      <tr style="background:#E8E4DF;">
        <th style="padding:8px 14px;text-align:left;border:1px solid #D0C8C0;
             font-weight:700;min-width:140px;">TYPE</th>
        <th style="padding:8px;text-align:center;border:1px solid #D0C8C0;
             font-weight:700;min-width:70px;">AS IS</th>
        <th style="padding:8px;text-align:center;border:1px solid #D0C8C0;
             font-weight:700;min-width:70px;border-right:2px solid #C8C0B8;">TO BE</th>
        <th style="padding:8px;text-align:center;border:1px solid #D0C8C0;
             font-weight:700;min-width:80px;">AS IS</th>
        <th style="padding:8px;text-align:center;border:1px solid #D0C8C0;
             font-weight:700;min-width:80px;">TO BE</th>
        <th style="padding:8px;text-align:center;border:1px solid #D0C8C0;
             font-weight:700;min-width:80px;border-right:2px solid #C8C0B8;">DIFF</th>
        <th style="padding:8px;text-align:center;border:1px solid #D0C8C0;
             font-weight:700;min-width:80px;">AS IS</th>
      </tr>
    </thead>
    <tbody>{_tbody}</tbody>
  </table>
</div>""", unsafe_allow_html=True)

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# CHARTS  (only when we have type + SKU data)
# ══════════════════════════════════════════════════════════════════════════════
if type_col and (as_is_cols or to_be_cols):
    grp       = merged.copy()
    grp[type_col] = grp[type_col].astype(str).str.strip().str.upper()
    data_rows = grp[~grp[type_col].isin(["TOTAL SKU","% IMPACT","TOTAL","NAN",""])]

    def safe_num(s): return pd.to_numeric(s, errors="coerce").fillna(0)

    data_rows = data_rows.copy()
    data_rows["_asis"] = safe_num(data_rows[asis_col]) if asis_col else 0
    data_rows["_tobe"] = safe_num(data_rows[tobe_col]) if tobe_col else 0

    summary = data_rows.groupby(type_col)[["_asis","_tobe"]].sum().reset_index()
    summary.columns = ["TYPE","AS IS","TO BE"]
    summary["CHANGE"] = summary["TO BE"] - summary["AS IS"]
    summary["COLOR"]  = summary["TYPE"].map(lambda t: TYPE_COLOR.get(t,"#888"))

    total_asis   = int(summary["AS IS"].sum())
    total_tobe   = int(summary["TO BE"].sum())
    total_change = total_tobe - total_asis
    pct_change   = (total_change / total_asis * 100) if total_asis else 0

    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Total SKU (AS IS)", f"{total_asis:,}")
    k2.metric("Total SKU (TO BE)", f"{total_tobe:,}", delta=f"{total_change:+,}")
    k3.metric("Net Change",        f"{total_change:+,}")
    k4.metric("% Impact",          f"{pct_change:.1f}%")

    st.divider()
    left, right = st.columns([1.6,1], gap="large")

    with left:
        st.markdown("**📊 AS IS vs TO BE by Type**")
        if HAS_PLOTLY:
            fig = go.Figure()
            fig.add_trace(go.Bar(name="AS IS", x=summary["TYPE"], y=summary["AS IS"],
                                 marker_color="#B0BEC5", text=summary["AS IS"], textposition="outside"))
            fig.add_trace(go.Bar(name="TO BE", x=summary["TYPE"], y=summary["TO BE"],
                                 marker_color=summary["COLOR"].tolist(),
                                 text=summary["TO BE"], textposition="outside"))
            fig.update_layout(barmode="group", height=320, plot_bgcolor="#F2EDE8",
                              paper_bgcolor="#F2EDE8", margin=dict(t=20,b=20,l=10,r=10),
                              font=dict(family="Inter,sans-serif",size=12),
                              legend=dict(orientation="h",y=1.08))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.bar_chart(summary.set_index("TYPE")[["AS IS","TO BE"]], height=300)

    with right:
        st.markdown("**🍩 TO BE Distribution**")
        pie_data = summary[summary["TO BE"] > 0]
        if len(pie_data) and HAS_PLOTLY:
            fig2 = px.pie(pie_data, names="TYPE", values="TO BE", color="TYPE",
                          color_discrete_map={t:c for t,c in zip(pie_data["TYPE"],pie_data["COLOR"])},
                          hole=0.45)
            fig2.update_layout(height=320, plot_bgcolor="#F2EDE8", paper_bgcolor="#F2EDE8",
                               margin=dict(t=20,b=20,l=10,r=10),
                               font=dict(family="Inter,sans-serif",size=11))
            fig2.update_traces(textposition="inside", textinfo="percent+label")
            st.plotly_chart(fig2, use_container_width=True)
        elif len(pie_data):
            st.bar_chart(pie_data.set_index("TYPE")["TO BE"], height=300)
        else:
            st.info("All TO BE values are 0")

    st.divider()
    st.markdown("**📈 Net Change per Type (TO BE − AS IS)**")
    if HAS_PLOTLY:
        fig3 = go.Figure(go.Bar(
            x=summary["TYPE"], y=summary["CHANGE"],
            marker_color=["#2BBFA4" if v>=0 else "#E05555" for v in summary["CHANGE"]],
            text=[f"{v:+}" for v in summary["CHANGE"]], textposition="outside"))
        fig3.add_hline(y=0, line_color="#888", line_width=1)
        fig3.update_layout(height=240, plot_bgcolor="#F2EDE8", paper_bgcolor="#F2EDE8",
                           margin=dict(t=20,b=20,l=10,r=10),
                           font=dict(family="Inter,sans-serif",size=12), yaxis_title="Change")
        st.plotly_chart(fig3, use_container_width=True)
    else:
        st.bar_chart(summary.set_index("TYPE")["CHANGE"], height=220)

else:
    # Generic stats when no type column detected
    st.metric("Total Rows",  f"{len(merged):,}")
    nc = [c for c in merged.columns if pd.api.types.is_numeric_dtype(merged[c])]
    if nc and HAS_PLOTLY:
        totals = merged[nc[:8]].sum().reset_index()
        totals.columns = ["Column","Total"]
        st.bar_chart(totals.set_index("Column"), height=280)
