"""Dashboard page — metric cards, bar chart, range architecture pie."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
import pandas as pd
import datetime as _dt
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar, render_page_nav,
    find_col, add_audit,
)

inject_css()
init_session_state()
render_sidebar("dashboard")
render_topbar("Dashboard")

add_audit("View Dashboard")

merged = st.session_state.merged_df
if merged is None:
    st.markdown("""
<div style="background:#fff;border-radius:16px;border:1px solid #E8E3DC;
            padding:60px;text-align:center;">
    <div style="font-size:48px;margin-bottom:16px;">📊</div>
    <div style="font-size:16px;font-weight:600;color:#1A1A1A;margin-bottom:8px;">No data yet</div>
    <div style="font-size:13px;color:#999;">Go to My Files and upload a file first.</div>
</div>
""", unsafe_allow_html=True)
    st.stop()

try:
    import plotly.graph_objects as go
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

# ── Header: Day/Week/Month filter right, no big title ────────────────────────
st.markdown("""
<style>
/* Period filter pill bar */
div[data-testid="stRadio"][data-key="dash_period"] > label { display: none !important; }
div[data-testid="stRadio"][data-key="dash_period"] > div[role="radiogroup"] {
    display: flex !important; flex-direction: row !important; gap: 0 !important;
    background: #EDE8DF; border-radius: 10px; padding: 4px;
    width: 100% !important;
}
div[data-testid="stRadio"][data-key="dash_period"] > div[role="radiogroup"] > label {
    flex: 1 !important;
    display: flex !important; align-items: center; justify-content: center;
    padding: 7px 0 !important; border-radius: 7px !important;
    font-size: 12px !important; font-weight: 600 !important;
    cursor: pointer; color: #888 !important; transition: all .15s;
    text-align: center !important;
}
div[data-testid="stRadio"][data-key="dash_period"] > div[role="radiogroup"] > label:has(input:checked) {
    background: #fff !important; color: #1A1A1A !important;
    box-shadow: 0 1px 4px rgba(0,0,0,.14) !important;
}
div[data-testid="stRadio"][data-key="dash_period"] > div[role="radiogroup"] > label > div:first-child {
    display: none !important;
}
</style>
""", unsafe_allow_html=True)

_, _h_filter = st.columns([2.5, 1])
with _h_filter:
    _period = st.radio(
        "Period", ["Week", "Month"],
        horizontal=True, key="dash_period", label_visibility="collapsed",
    )

st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

# ── Apply date filter ─────────────────────────────────────────────────────────
_date_col = find_col(merged, ["date", "week", "month", "period"])
_today    = pd.Timestamp.now().normalize()
_df = merged.copy()

_date_filtered = False
if _date_col:
    try:
        _dt_series = pd.to_datetime(_df[_date_col], errors="coerce")
        if _period == "Day":
            _mask = _dt_series >= _today
        elif _period == "Week":
            _mask = _dt_series >= _today - pd.Timedelta(days=7)
        else:
            _mask = _dt_series >= _today - pd.Timedelta(days=30)
        if _mask.sum() > 0:
            _df = _df[_mask]
            _date_filtered = True
    except Exception:
        pass

if _date_filtered and len(_df) < len(merged):
    st.markdown(
        f"<div style='font-size:11px;color:#2BBFA4;margin-bottom:12px;font-weight:600;'>"
        f"Showing {len(_df):,} rows for selected period ({_period})</div>",
        unsafe_allow_html=True,
    )

# ── Derive KPIs ───────────────────────────────────────────────────────────────
status_col = find_col(_df, ["status"])

total_skus = len(_df)

_pog_col   = find_col(_df, ["pogname", "pog name", "planogramname", "planogram name", "pog"])
total_pog  = int(_df[_pog_col].dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique()) if _pog_col else 0

_store_col  = find_col(_df, ["store", "store number", "store no", "storeno", "storenumber"])
total_store = int(_df[_store_col].dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique()) if _store_col else 0

if status_col:
    _statuses    = _df[status_col].astype(str).str.strip().str.upper()
    maintain_cnt = int((_statuses == "MAINTAIN").sum())
    new_new_cnt  = int((_statuses.isin(["NEWNEW", "NEW NEW"])).sum())
    delete_all_cnt = int((_statuses == "DELETE ALL").sum())
else:
    maintain_cnt   = 0
    new_new_cnt    = 0
    delete_all_cnt = 0

pct_maintain   = round(maintain_cnt   / total_skus * 100, 1) if total_skus else 0
pct_new_new    = round(new_new_cnt    / total_skus * 100, 1) if total_skus else 0
pct_delete_all = round(delete_all_cnt / total_skus * 100, 1) if total_skus else 0

# ── Metric cards ──────────────────────────────────────────────────────────────
def _metric_card(label, value, delta_text, delta_positive, bar_color, bar_pct):
    delta_color = "#2BBFA4" if delta_positive else "#E05555"
    val_color   = "#1A1A1A"
    return f"""
<div style="background:#fff;border-radius:16px;border:1px solid #E8E3DC;
            padding:20px 20px 14px;box-shadow:0 1px 4px rgba(0,0,0,.05);">
    <div style="font-size:10px;font-weight:700;color:#888;text-transform:uppercase;
                letter-spacing:0.07em;margin-bottom:8px;">{label}</div>
    <div style="font-size:30px;font-weight:800;color:{val_color};margin-bottom:6px;
                line-height:1.1;">{value}</div>
    <div style="font-size:12px;color:{delta_color};font-weight:500;margin-bottom:14px;">
        {delta_text}
    </div>
    <div style="background:#EDE8DF;border-radius:99px;height:4px;">
        <div style="background:{bar_color};border-radius:99px;height:4px;
                    width:{min(bar_pct,100):.0f}%;"></div>
    </div>
</div>"""

c1, c2, c3 = st.columns(3, gap="medium")
with c1:
    st.markdown(_metric_card(
        "TOTAL SKUS", f"{total_skus:,}",
        f"{total_skus:,} items in range", True, "#2BBFA4", 100,
    ), unsafe_allow_html=True)
with c2:
    st.markdown(_metric_card(
        "TOTAL POG", f"{total_pog:,}",
        f"{total_pog:,} planograms", True, "#3B82F6", min(total_pog, 100),
    ), unsafe_allow_html=True)
with c3:
    st.markdown(_metric_card(
        "TOTAL STORE", f"{total_store:,}",
        f"{total_store:,} stores", True, "#F59E0B", min(total_store, 100),
    ), unsafe_allow_html=True)

st.markdown("<div style='height:24px;'></div>", unsafe_allow_html=True)

# ── Bar chart (monthly) + Range Status ───────────────────────────────────────
left, right = st.columns([1.7, 1], gap="large")

with left:
    st.markdown("""
<div style="background:#fff;border-radius:16px;border:1px solid #E8E3DC;padding:20px 20px 8px;">
    <div style="font-size:15px;font-weight:700;color:#1A1A1A;margin-bottom:16px;">
        Monthly Active SKUs
    </div>""", unsafe_allow_html=True)

    date_col = find_col(_df, ["date", "week", "month", "period"])
    months   = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    cur_month = _dt.datetime.now().month - 1

    if date_col:
        try:
            _dates = pd.to_datetime(_df[date_col], errors="coerce")
            _mc    = _dates.dt.month.value_counts().reindex(range(1, 13), fill_value=0)
            values = [int(_mc.get(i, 0)) for i in range(1, 13)]
        except Exception:
            values = [max(1, int(total_skus * 0.07 + (i - 6) ** 2 * 2)) for i in range(12)]
    else:
        base   = max(total_skus // 12, 1)
        values = [max(1, int(base * 0.6 + (i % 4) * base * 0.1)) for i in range(12)]
        values[cur_month] = max(values)

    if HAS_PLOTLY:
        colors = ["#2BBFA4" if i == cur_month else "#C8EDE7" for i in range(12)]
        fig = go.Figure(go.Bar(
            x=months, y=values,
            marker_color=colors,
            marker_line_width=0,
        ))
        fig.update_layout(
            height=260, plot_bgcolor="#fff", paper_bgcolor="#fff",
            margin=dict(t=4, b=10, l=4, r=4),
            font=dict(family="Inter,sans-serif", size=11, color="#888"),
            xaxis=dict(showgrid=False, tickfont=dict(size=11, color="#888")),
            yaxis=dict(showgrid=True, gridcolor="#F0EBE3",
                       tickfont=dict(size=10, color="#ccc"), zeroline=False),
            showlegend=False,
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    else:
        chart_df = pd.DataFrame({"Month": months, "SKUs": values}).set_index("Month")
        st.bar_chart(chart_df, height=240, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

with right:
    st.markdown("""
<div style="background:#fff;border-radius:16px;border:1px solid #E8E3DC;
            padding:20px 20px 8px;box-sizing:border-box;">
    <div style="font-size:15px;font-weight:700;color:#1A1A1A;margin-bottom:4px;">
        Range Status
    </div>
</div>""", unsafe_allow_html=True)

    # ── DEMO data ────────────────────────────────────────────────────────
    _rs_labels = ["Maintain", "New New", "Delete All"]
    _rs_vals   = [850,         320,       180        ]
    _rs_colors = ["#2BBFA4",  "#3B82F6", "#E05555"  ]
    _rs_total  = sum(_rs_vals)

    if HAS_PLOTLY:
        fig_rs = go.Figure(go.Pie(
            labels=_rs_labels,
            values=_rs_vals,
            marker=dict(colors=_rs_colors, line=dict(color="#fff", width=2)),
            hole=0.45,
            textinfo="value",
            textfont=dict(size=11, color="#fff"),
            hovertemplate="<b>%{label}</b><br>%{value:,} SKUs (%{percent})<extra></extra>",
            sort=False,
        ))
        fig_rs.add_annotation(
            text=f"<b>{_rs_total:,}</b><br><span style='font-size:9px;'>SKUs</span>",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=13, color="#1A1A1A", family="Inter,sans-serif"),
            align="center",
        )
        fig_rs.update_layout(
            height=260,
            margin=dict(t=8, b=8, l=8, r=8),
            showlegend=True,
            legend=dict(
                orientation="v", x=1.02, y=0.5,
                font=dict(size=11, color="#555"),
            ),
            paper_bgcolor="#fff",
            font=dict(family="Inter,sans-serif"),
        )
        st.markdown('<div style="background:#fff;border-radius:0 0 16px 16px;border:1px solid #E8E3DC;border-top:none;padding:0 12px 12px;">', unsafe_allow_html=True)
        st.plotly_chart(fig_rs, use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        _rs_rows = ""
        for _lbl, _pct, _clr in zip(_rs_labels, _rs_vals, _rs_colors):
            _rs_rows += f'<div style="display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid #F5F0EA;"><span style="font-size:13px;color:#1A1A1A;">{_lbl}</span><span style="font-size:13px;font-weight:700;color:{_clr};">{_pct}%</span></div>'
        st.markdown(f'<div style="background:#fff;border-radius:0 0 16px 16px;border:1px solid #E8E3DC;border-top:none;padding:12px 20px;">{_rs_rows}</div>', unsafe_allow_html=True)

# ── Range Architecture Pie Chart ──────────────────────────────────────────────
st.markdown("<div style='height:24px;'></div>", unsafe_allow_html=True)

_nc = lambda s: str(s).lower().strip()
_all_cols = list(_df.columns)
_arch_col = (
    next((c for c in _all_cols if _nc(c) in ("item priority", "itempriority")), None)
    or next((c for c in _all_cols if "item" in _nc(c) and "priority" in _nc(c)), None)
    or next((c for c in _all_cols if _nc(c) == "status"), None)
)

ARCH_TYPES  = ["MAINTAIN", "NEW DELETE SOME", "DELETE SOME", "DELETE ALL", "NEW SOME", "NEWNEW"]
ARCH_COLORS = ["#2BBFA4",  "#F59E0B",         "#E05555",     "#C0392B",    "#3B82F6",  "#8B5CF6"]
ARCH_LABELS = {
    "MAINTAIN":        "Maintain",
    "NEW DELETE SOME": "New Delete Some",
    "DELETE SOME":     "Delete Some",
    "DELETE ALL":      "Delete All",
    "NEW SOME":        "New Some",
    "NEWNEW":          "New New",
}

if _arch_col and HAS_PLOTLY:
    _type_series = _df[_arch_col].astype(str).str.strip().str.upper()
    _counts = {t: int((_type_series == t).sum()) for t in ARCH_TYPES}
    _total_arch = sum(_counts.values())

    _pie_types  = [t for t in ARCH_TYPES if _counts[t] > 0]
    _pie_vals   = [_counts[t] for t in _pie_types]
    _pie_cols   = [ARCH_COLORS[ARCH_TYPES.index(t)] for t in _pie_types]
    _pie_labels = [ARCH_LABELS[t] for t in _pie_types]

    _pie_card, _pie_legend = st.columns([1.4, 1], gap="large")

    with _pie_card:
        st.markdown("""
<div style="background:#fff;border-radius:16px;border:1px solid #E8E3DC;
            padding:20px 20px 8px;">
    <div style="font-size:15px;font-weight:700;color:#1A1A1A;margin-bottom:4px;">
        Range Architecture — SKU Distribution
    </div>
    <div style="font-size:11px;color:#999;margin-bottom:8px;">
        Total SKUs by range action type
    </div>""", unsafe_allow_html=True)

        if _pie_vals:
            fig_pie = go.Figure(go.Pie(
                labels=_pie_labels,
                values=_pie_vals,
                marker=dict(
                    colors=_pie_cols,
                    line=dict(color="#fff", width=2),
                ),
                hole=0.42,
                textinfo="percent",
                textfont=dict(size=11, color="#fff"),
                hovertemplate=(
                    "<b>%{label}</b><br>"
                    "%{value:,} SKUs<br>"
                    "%{percent}<extra></extra>"
                ),
                sort=False,
            ))
            fig_pie.add_annotation(
                text=f"<b>{_total_arch:,}</b><br><span style='font-size:10px;'>Total SKUs</span>",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=13, color="#1A1A1A", family="Inter,sans-serif"),
                align="center",
            )
            fig_pie.update_layout(
                height=300,
                margin=dict(t=10, b=10, l=10, r=10),
                showlegend=False,
                paper_bgcolor="#fff",
                font=dict(family="Inter,sans-serif"),
            )
            st.plotly_chart(fig_pie, use_container_width=True, config={"displayModeBar": False})
        else:
            st.markdown(
                "<p style='color:#999;font-size:13px;padding:40px 0;text-align:center;'>"
                "No matching type data found for this period.</p>",
                unsafe_allow_html=True,
            )

        st.markdown("</div>", unsafe_allow_html=True)

    with _pie_legend:
        _legend_rows = ""
        for t, color in zip(ARCH_TYPES, ARCH_COLORS):
            cnt = _counts.get(t, 0)
            pct = round(cnt / _total_arch * 100, 1) if _total_arch else 0
            _legend_rows += f"""
<div style="display:flex;align-items:center;justify-content:space-between;
            padding:12px 0;border-bottom:1px solid #F5F0EA;">
    <div style="display:flex;align-items:center;gap:10px;">
        <div style="width:12px;height:12px;border-radius:3px;background:{color};flex-shrink:0;"></div>
        <span style="font-size:13px;color:#1A1A1A;font-weight:500;">{ARCH_LABELS[t]}</span>
    </div>
    <div style="text-align:right;">
        <span style="font-size:13px;font-weight:700;color:#1A1A1A;">{cnt:,}</span>
        <span style="font-size:11px;color:#999;margin-left:6px;">{pct:.1f}%</span>
    </div>
</div>"""

        st.markdown(f"""
<div style="background:#fff;border-radius:16px;border:1px solid #E8E3DC;
            padding:20px;height:100%;box-sizing:border-box;">
    <div style="font-size:15px;font-weight:700;color:#1A1A1A;margin-bottom:4px;">
        SKU Breakdown
    </div>
    <div style="font-size:11px;color:#999;margin-bottom:16px;">
        Count &amp; share per action type
    </div>
    {_legend_rows}
    <div style="display:flex;justify-content:space-between;padding-top:12px;">
        <span style="font-size:12px;font-weight:700;color:#555;">Total</span>
        <span style="font-size:13px;font-weight:800;color:#1A1A1A;">{_total_arch:,}</span>
    </div>
</div>
""", unsafe_allow_html=True)

elif not HAS_PLOTLY:
    st.info("Install plotly (`pip install plotly`) to enable the pie chart.")
else:
    st.markdown("""
<div style="background:#FFF9C4;border-radius:12px;padding:16px 20px;border:1px solid #F0E68C;">
    <span style="font-size:13px;color:#7A6000;">
        ⚠️ No <strong>Status</strong> or <strong>Type</strong> column found in the uploaded data
        — Range Architecture chart unavailable.
    </span>
</div>""", unsafe_allow_html=True)

render_page_nav("dashboard")
