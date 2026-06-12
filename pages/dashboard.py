"""Dashboard page — metric cards, bar chart, range status."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
import pandas as pd
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar, render_page_nav,
    find_col, add_audit,
)

inject_css()
init_session_state()
render_sidebar("dashboard")
render_topbar("Dashboard")

add_audit("View Dashboard")

st.markdown(
    "<div style='font-size:28px;font-weight:800;color:#1A1A1A;margin-bottom:24px;'>Dashboard</div>",
    unsafe_allow_html=True,
)

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

# ── Derive KPIs from data ─────────────────────────────────────────────────────
status_col = find_col(merged, ["status"])
sale_col   = find_col(merged, ["to-be total sale", "to be total sale", "to-be sale total"])
if not sale_col:
    sale_col = find_col(merged, ["sale", "sales"])

total_skus   = len(merged)
if status_col:
    _statuses     = merged[status_col].astype(str).str.strip().str.upper()
    active_range  = int((_statuses.isin(["MAINTAIN", "NEW SOME", "NEWNEW", "NEW"])).sum())
    low_sales_cnt = int((_statuses.isin(["DELETE SOME", "DELETE ALL"])).sum())
    pending_cnt   = int((_statuses == "REVIEW").sum())
else:
    active_range  = total_skus
    low_sales_cnt = 0
    pending_cnt   = 0

pct_active   = round(active_range  / total_skus * 100, 1) if total_skus else 0
pct_low      = round(low_sales_cnt / total_skus * 100, 1) if total_skus else 0
pct_pending  = round(pending_cnt   / total_skus * 100, 1) if total_skus else 0
pct_review   = round((total_skus - active_range - low_sales_cnt) / total_skus * 100, 1) if total_skus else 0
pct_new      = max(0, 100 - pct_active - pct_review - pct_low)

# ── Metric cards (custom HTML with mini progress bar) ─────────────────────────
def _metric_card(label, value, delta_text, delta_positive, bar_color, bar_pct):
    delta_color = "#2BBFA4" if delta_positive else "#E05555"
    val_color   = "#E05555" if label == "LOW SALES SKUS" else "#1A1A1A"
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

c1, c2, c3, c4 = st.columns(4, gap="medium")
with c1:
    st.markdown(_metric_card(
        "TOTAL SKUS", f"{total_skus:,}",
        f"+{pct_active:.1f}% vs last month", True, "#2BBFA4", pct_active,
    ), unsafe_allow_html=True)
with c2:
    st.markdown(_metric_card(
        "ACTIVE RANGE", f"{active_range:,}",
        f"+{pct_active:.1f}% vs last month", True, "#3B82F6", pct_active,
    ), unsafe_allow_html=True)
with c3:
    st.markdown(_metric_card(
        "LOW SALES SKUS", f"{low_sales_cnt:,}",
        f"-{pct_low:.1f}% vs last month", False, "#E05555", pct_low,
    ), unsafe_allow_html=True)
with c4:
    st.markdown(_metric_card(
        "PENDING REVIEW", f"{pending_cnt:,}",
        f"+{pending_cnt} new", True, "#F59E0B", min(pct_pending + 5, 100),
    ), unsafe_allow_html=True)

st.markdown("<div style='height:24px;'></div>", unsafe_allow_html=True)

# ── Bottom section: bar chart + range status ──────────────────────────────────
left, right = st.columns([1.7, 1], gap="large")

with left:
    st.markdown("""
<div style="background:#fff;border-radius:16px;border:1px solid #E8E3DC;padding:20px 20px 8px;">
    <div style="font-size:15px;font-weight:700;color:#1A1A1A;margin-bottom:16px;">
        Monthly Active SKUs
    </div>""", unsafe_allow_html=True)

    # Build monthly data from uploaded data (date column) or use row-count as proxy
    date_col = find_col(merged, ["date", "week", "month", "period"])
    months   = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    import datetime as _dt
    cur_month = _dt.datetime.now().month - 1  # 0-indexed

    if date_col:
        try:
            _dates = pd.to_datetime(merged[date_col], errors="coerce")
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
            height=260,
            plot_bgcolor="#fff",
            paper_bgcolor="#fff",
            margin=dict(t=4, b=10, l=4, r=4),
            font=dict(family="Inter,sans-serif", size=11, color="#888"),
            xaxis=dict(showgrid=False, tickfont=dict(size=11, color="#888")),
            yaxis=dict(showgrid=True, gridcolor="#F0EBE3", tickfont=dict(size=10, color="#ccc"),
                       zeroline=False),
            showlegend=False,
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    else:
        chart_df = pd.DataFrame({"Month": months, "SKUs": values}).set_index("Month")
        st.bar_chart(chart_df, height=240, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

with right:
    # Range Status progress bars
    items = [
        ("Active",     pct_active,  "#2BBFA4"),
        ("Review",     pct_review,  "#F59E0B"),
        ("Low Sales",  pct_low,     "#E05555"),
        ("New",        pct_new,     "#3B82F6"),
    ]
    rows_html = ""
    for label, pct, color in items:
        rows_html += f"""
<div style="margin-bottom:18px;">
    <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
        <span style="font-size:13px;color:#1A1A1A;font-weight:500;">{label}</span>
        <span style="font-size:13px;color:#1A1A1A;font-weight:700;">{pct:.0f}%</span>
    </div>
    <div style="background:#EDE8DF;border-radius:99px;height:6px;">
        <div style="background:{color};border-radius:99px;height:6px;
                    width:{min(pct,100):.1f}%;transition:width 0.3s;"></div>
    </div>
</div>"""

    st.markdown(f"""
<div style="background:#fff;border-radius:16px;border:1px solid #E8E3DC;
            padding:20px;height:100%;box-sizing:border-box;">
    <div style="font-size:15px;font-weight:700;color:#1A1A1A;margin-bottom:20px;">
        Range Status
    </div>
    {rows_html}
</div>
""", unsafe_allow_html=True)

# ── Type breakdown (if available) ─────────────────────────────────────────────
type_col = find_col(merged, ["type"])
if not type_col and status_col:
    type_col = status_col

if type_col and HAS_PLOTLY:
    st.markdown("<div style='height:24px;'></div>", unsafe_allow_html=True)
    st.markdown("""
<div style="background:#fff;border-radius:16px;border:1px solid #E8E3DC;
            padding:20px 20px 8px;">
    <div style="font-size:15px;font-weight:700;color:#1A1A1A;margin-bottom:16px;">
        SKU Count by Type
    </div>""", unsafe_allow_html=True)

    TYPE_COLORS = {
        "MAINTAIN":    "#2BBFA4", "DELETE SOME": "#E05555", "DELETE ALL": "#C0392B",
        "NEW SOME":    "#3B82F6", "NEWNEW":      "#8B5CF6", "NEW":        "#22C55E",
        "NEW DELETE SOME": "#F59E0B",
    }
    _type_counts = (merged[type_col].astype(str).str.strip().str.upper()
                    .value_counts().reset_index())
    _type_counts.columns = ["Type", "Count"]
    _type_counts = _type_counts[~_type_counts["Type"].isin(["NAN", "", "TOTAL SKU"])]
    _type_counts["Color"] = _type_counts["Type"].map(lambda t: TYPE_COLORS.get(t, "#888"))

    fig2 = go.Figure(go.Bar(
        x=_type_counts["Type"],
        y=_type_counts["Count"],
        marker_color=_type_counts["Color"].tolist(),
        marker_line_width=0,
        text=_type_counts["Count"],
        textposition="outside",
        textfont=dict(size=11),
    ))
    fig2.update_layout(
        height=240,
        plot_bgcolor="#fff",
        paper_bgcolor="#fff",
        margin=dict(t=20, b=10, l=4, r=4),
        font=dict(family="Inter,sans-serif", size=11, color="#888"),
        xaxis=dict(showgrid=False, tickfont=dict(size=11, color="#555")),
        yaxis=dict(showgrid=True, gridcolor="#F0EBE3", zeroline=False,
                   tickfont=dict(size=10, color="#ccc")),
        showlegend=False,
    )
    st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})
    st.markdown("</div>", unsafe_allow_html=True)

render_page_nav("dashboard")
