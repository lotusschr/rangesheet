"""Report page — Range Execution & Summary Analytics."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
import pandas as pd
import zipfile
import io
import base64
import html as _html
import re as _re
from datetime import datetime
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar, render_page_nav,
    find_col, df_to_xlsx_bytes, df_to_csv_bytes, add_audit,
    BASE_DIR,
)

def _report_partner_logo_svg() -> str:
    return """
<svg viewBox="0 0 160 48" role="img" xmlns="http://www.w3.org/2000/svg">
  <title>Lotus's</title>
  <g fill="none" fill-rule="evenodd">
    <text x="0" y="37" fill="#72D4CD" font-family="Arial, Helvetica, sans-serif" font-size="36" font-weight="800" letter-spacing="-1.4">Lotus</text>
    <path d="M113 4 C120 3 124 9 122 16 C120 22 115 27 113 36 C110 27 105 22 104 16 C103 9 107 5 113 4Z" fill="#F6D975"/>
    <text x="124" y="37" fill="#F6D975" font-family="Arial, Helvetica, sans-serif" font-size="36" font-weight="800" letter-spacing="-1.4">s</text>
  </g>
</svg>
""".strip()

def _report_partner_logo_data_uri() -> str:
    return "data:image/svg+xml;base64," + base64.b64encode(_report_partner_logo_svg().encode("utf-8")).decode("ascii")

def _report_partner_logo_png_bytes(width: int = 160, height: int = 48) -> bytes:
    from PIL import Image, ImageDraw, ImageFont
    scale = width / 160
    img = Image.new("RGBA", (width, height), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    try:
        f_lotus = ImageFont.truetype("arialbd.ttf", max(10, int(36 * scale)))
    except Exception:
        f_lotus = ImageFont.load_default()
    yellow, teal = (246, 217, 117, 255), (114, 212, 205, 255)
    draw.text((0, 4 * scale), "Lotus", fill=teal, font=f_lotus)
    draw.polygon([(113 * scale, 4 * scale), (120 * scale, 3 * scale), (124 * scale, 9 * scale), (122 * scale, 16 * scale), (115 * scale, 27 * scale), (113 * scale, 36 * scale), (104 * scale, 16 * scale), (107 * scale, 5 * scale)], fill=yellow)
    draw.text((124 * scale, 4 * scale), "s", fill=yellow, font=f_lotus)
    out = io.BytesIO()
    img.save(out, format="PNG")
    return out.getvalue()

inject_css()
init_session_state()
render_sidebar("report")
render_topbar("Report")

merged = st.session_state.merged_df

rt0, rt1 = st.tabs(["Execution Report", "Output from Rangesheet"])

_REPORT_PACKET_DIR = os.path.join(BASE_DIR, "rangesheet_data", ".autosave", "rangesheetreview")


def _report_packets_path(p: str) -> str:
    return os.path.join(_REPORT_PACKET_DIR, f"report_packets_{p}.pkl")


def _load_report_packets(p: str) -> dict:
    path = _report_packets_path(p)
    if not os.path.exists(path):
        return {}
    try:
        data = pd.read_pickle(path)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


for _p_load in ("ss", "ns", "sa"):
    _key_load = f"vw_submit_reports_{_p_load}"
    if not st.session_state.get(_key_load):
        _loaded_packets = _load_report_packets(_p_load)
        if _loaded_packets:
            st.session_state[_key_load] = _loaded_packets


def _latest_output_from_packets(p: str):
    packets = st.session_state.get(f"vw_submit_reports_{p}", {})
    if not isinstance(packets, dict) or not packets:
        return None
    packet_items = sorted(
        packets.items(),
        key=lambda kv: str((kv[1] or {}).get("submitted_order") or (kv[1] or {}).get("submitted_at") or kv[0]),
    )
    latest_by_dg = {}
    for packet_key, packet in packet_items:
        if not isinstance(packet, dict):
            continue
        df_packet = packet.get("df")
        if df_packet is None or getattr(df_packet, "empty", True):
            continue
        dg_key = str(packet.get("dg_key") or packet.get("label") or packet_key)
        if dg_key in latest_by_dg:
            latest_by_dg.pop(dg_key, None)
        latest_by_dg[dg_key] = df_packet
    if not latest_by_dg:
        return None
    try:
        return pd.concat(list(latest_by_dg.values()), ignore_index=True, sort=False)
    except Exception:
        return next(iter(latest_by_dg.values()), None)


for _p_out in ("ss", "ns", "sa"):
    _key_out = f"vw_submit_{_p_out}"
    _df_out = st.session_state.get(_key_out)
    if _df_out is None or getattr(_df_out, "empty", True):
        _latest_output = _latest_output_from_packets(_p_out)
        if _latest_output is not None and not getattr(_latest_output, "empty", True):
            st.session_state[_key_out] = _latest_output


def _h(v) -> str:
    return _html.escape("" if v is None or str(v) in ("nan", "None", "<NA>") else str(v))


def _nca(v) -> str:
    return "".join(ch for ch in str(v or "").lower() if ch.isalnum())


def _cluster_from_planogram_name(pog: str) -> str:
    s = str(pog or "").strip()
    if not s:
        return ""
    m = _re.search(r"\b([A-Z]+_G(?:_[A-Z0-9]+)+)\b", s.upper())
    if m:
        return m.group(1)
    m = _re.search(r"\b([A-Z]_G_\d{4,})\b", s.upper())
    if m:
        return m.group(1)
    target = _re.search(r"(?:target[_\-\s]*|_)(\d{4,})", s, flags=_re.I)
    prefix = _re.search(r"\b([A-Z]+_G)\b", s.upper())
    if target and prefix:
        return f"{prefix.group(1)}_{target.group(1)}"
    return ""


def _find_col_any(df: pd.DataFrame, names: list[str]):
    if df is None or df.empty:
        return None
    wanted = {_nca(x) for x in names}
    for c in df.columns:
        if _nca(c) in wanted:
            return c
    for c in df.columns:
        cn = _nca(c)
        if any(w in cn for w in wanted):
            return c
    return None


def _fmt_int(v) -> str:
    try:
        return f"{int(v):,}"
    except Exception:
        return str(v)


def _fmt_num(v) -> str:
    try:
        f = float(v)
        return f"{f:,.0f}" if abs(f - int(f)) < 0.00001 else f"{f:,.1f}"
    except Exception:
        return str(v)


def _total_sku_all_dg(df: pd.DataFrame) -> int:
    if df is None or getattr(df, "empty", True):
        return 0
    id_col = _find_col_any(df, ["ID", "Item ID", "Article"])
    if not id_col or id_col not in df.columns:
        return int(len(df))
    dg_col = _find_col_any(df, ["DG Code", "DG", "Department"])
    _ids = df[id_col].astype(str).str.strip()
    _valid = _ids.ne("") & ~_ids.str.lower().isin(("nan", "none", "<na>"))
    if dg_col and dg_col in df.columns:
        _pairs = pd.DataFrame({
            "_dg": df.loc[_valid, dg_col].astype(str).str.strip(),
            "_id": _ids.loc[_valid],
        })
        return int(_pairs.drop_duplicates().shape[0])
    return int(_ids.loc[_valid].drop_duplicates().shape[0])


def _merge_action_stores(p: str) -> dict:
    out = {}
    for key in (f"{p}_pog_actions", f"{p}_pog_edits"):
        store = st.session_state.get(key, {})
        if isinstance(store, dict):
            for rk, acts in store.items():
                if isinstance(acts, dict):
                    out.setdefault(rk, {}).update(acts)
    return out


def _rk_parts(rk) -> tuple[str, str, str]:
    vals = [str(x) for x in rk] if isinstance(rk, (tuple, list)) else [str(rk)]
    return (
        vals[0] if len(vals) > 0 else "",
        vals[1] if len(vals) > 1 else "",
        vals[2] if len(vals) > 2 else "",
    )


def _dg_code_from_packet_key(v) -> str:
    s = str(v or "").strip()
    if not s:
        return ""
    first = s.split("|", 1)[0].strip()
    if first and first.upper() != "ALL":
        return first
    m = _re.search(r"\b([A-Z]\d[A-Z0-9])\b", s.upper())
    return m.group(1) if m else ""


def _build_management_report(
    p: str,
    label: str,
    df_override=None,
    submitted_at: str = "",
    pog_cluster_map: dict | None = None,
    dg_filter: str = "",
) -> dict:
    df = df_override if df_override is not None else st.session_state.get(f"vw_submit_{p}")
    if df is None:
        df = pd.DataFrame()
    actions = _merge_action_stores(p)
    statuses = st.session_state.get(f"{p}_status_overrides", {})
    statuses = statuses if isinstance(statuses, dict) else {}
    if df_override is not None:
        actions = {}
        statuses = {}

    id_col = _find_col_any(df, ["ID", "Item ID", "Article"])
    pog_col = _find_col_any(df, ["Planogram Name", "POG Name", "Planogram", "Name"])
    cluster_col = _find_col_any(df, ["POG_Cluster", "Cluster", "Cluster Name"])
    value_col = _find_col_any(df, ["TH_Tot_Sales_Value_52WK", "Tot Sales Value 52WK", "Sales Value"])
    volume_col = _find_col_any(df, ["TH_Tot_Sales_Volume_52WK", "Tot Sales Volume 52WK", "Sales Volume"])
    forecast_col = _find_col_any(df, ["ForecastSales", "Forecast Sales", "Forecast new item sales"])
    metric_col = volume_col or value_col or forecast_col
    status_col = _find_col_any(df, ["Status"])
    dg_col = _find_col_any(df, ["DG Code", "DG", "Department"])
    item_name_col = _find_col_any(df, ["Item Name", "ItemName", "Product Name", "Description"])
    _dg_filter = str(dg_filter or "").strip()
    if _dg_filter and dg_col and dg_col in df.columns and isinstance(df, pd.DataFrame) and not df.empty:
        _mask = df[dg_col].astype(str).str.strip().str.upper().eq(_dg_filter.upper())
        df = df.loc[_mask].copy()

    # Prefer the submitted grid snapshot when it contains edits. This keeps the
    # report correct after page refreshes or when widget edit state is not loaded.
    snapshot_statuses = {}
    snapshot_actions = {}
    if isinstance(df, pd.DataFrame) and not df.empty:
        if status_col and status_col in df.columns:
            _status_series = df[status_col].astype(str).str.strip().str.upper()
            _status_mask = _status_series.ne("") & ~_status_series.isin(("MAINTAIN", "NAN", "NONE"))
        else:
            _status_series = pd.Series([""] * len(df), index=df.index)
            _status_mask = pd.Series(False, index=df.index)
        _action_cols = []
        for _col in df.columns:
            _col_s = df[_col].astype(str).str.strip()
            if _col_s.isin(("Delete", "New")).any():
                _action_cols.append(_col)
        _action_mask = (
            df[_action_cols].astype(str).apply(lambda _col: _col.str.strip().isin(("Delete", "New"))).any(axis=1)
            if _action_cols else pd.Series(False, index=df.index)
        )
        _scan_idx = df.index[_status_mask | _action_mask]
        for _idx in _scan_idx:
            row = df.loc[_idx]
            dg_val = str(row.get(dg_col, "") if dg_col else "").strip()
            id_val = str(row.get(id_col, "") if id_col else "").strip()
            item_val = str(row.get(item_name_col, "") if item_name_col else "").strip()
            rk = (dg_val, id_val, item_val)
            status_val = str(_status_series.loc[_idx]).strip().upper()
            if status_val and status_val not in ("MAINTAIN", "NAN", "NONE"):
                snapshot_statuses[rk] = status_val
            acts = {}
            for col in _action_cols:
                val = row.get(col, "")
                action = str(val).strip()
                if action in ("Delete", "New"):
                    acts[str(col)] = action
            if acts:
                snapshot_actions[rk] = acts
        if snapshot_statuses or snapshot_actions:
            statuses = snapshot_statuses
            actions = snapshot_actions
    if _dg_filter:
        statuses = {
            rk: val for rk, val in (statuses or {}).items()
            if _rk_parts(rk)[0].strip().upper() == _dg_filter.upper()
        }
        actions = {
            rk: val for rk, val in (actions or {}).items()
            if _rk_parts(rk)[0].strip().upper() == _dg_filter.upper()
        }

    if id_col and id_col in df.columns:
        total_sku = int(df[id_col].astype(str).replace("", pd.NA).dropna().nunique())
    else:
        total_sku = int(len(df))

    affected_keys = set(statuses) | set(actions)
    status_counts = {
        "MAINTAIN": max(0, total_sku - len(affected_keys)),
        "DELETE SOME": 0,
        "DELETE ALL": 0,
        "NEW SOME": 0,
        "NEWNEW": 0,
    }
    for stv in statuses.values():
        s = str(stv or "").strip().upper()
        if s in status_counts:
            status_counts[s] += 1

    if metric_col and id_col and metric_col in df.columns:
        metric_by_id = (
            df[[id_col, metric_col]]
            .assign(_metric=pd.to_numeric(df[metric_col], errors="coerce"))
            .groupby(id_col)["_metric"].sum(min_count=1)
        )
        vals = metric_by_id.dropna().sort_values()
        top_cut = vals.quantile(0.90) if len(vals) else None
        low_cut = vals.quantile(0.10) if len(vals) else None
    else:
        metric_by_id = pd.Series(dtype=float)
        top_cut = low_cut = None

    id_to_pogs = {}
    id_to_cluster = {}
    if id_col and pog_col and id_col in df.columns and pog_col in df.columns:
        for item_id, grp in df.groupby(df[id_col].astype(str)):
            id_to_pogs[str(item_id)] = sorted(
                set(x for x in grp[pog_col].astype(str).str.strip() if x and x not in ("nan", "None"))
            )
            if cluster_col and cluster_col in grp.columns:
                clusters = [x for x in grp[cluster_col].astype(str).str.strip() if x and x not in ("nan", "None")]
                if clusters:
                    id_to_cluster[str(item_id)] = clusters[0]

    if (not pog_cluster_map) and pog_col and cluster_col and pog_col in df.columns and cluster_col in df.columns:
        _pc_map = (
            df[[pog_col, cluster_col]]
            .dropna()
            .astype(str)
            .drop_duplicates()
        )
        pog_cluster_map = {
            str(_r[pog_col]).strip(): str(_r[cluster_col]).strip()
            for _, _r in _pc_map.iterrows()
            if str(_r.get(pog_col, "")).strip()
            and str(_r.get(cluster_col, "")).strip()
            and str(_r.get(cluster_col, "")).strip() not in ("nan", "None")
        }

    item_changes = []
    planogram_impact = {}
    cluster_impact = {}
    pog_cluster_map = {
        str(k): str(v).strip()
        for k, v in (pog_cluster_map or {}).items()
        if str(v).strip()
    }
    for _pog_name in list(planogram_impact.keys()) + [
        str(c) for c in df.columns
        if isinstance(df, pd.DataFrame) and str(c).strip()
    ]:
        _fallback_cluster = _cluster_from_planogram_name(_pog_name)
        if _fallback_cluster:
            pog_cluster_map.setdefault(str(_pog_name), _fallback_cluster)
            pog_cluster_map.setdefault(_nca(_pog_name), _fallback_cluster)

    def _cluster_for_pog(item_id: str, pog: str) -> str:
        pog_s = str(pog or "").strip()
        return (
            pog_cluster_map.get(pog_s)
            or pog_cluster_map.get(_nca(pog_s))
            or _cluster_from_planogram_name(pog_s)
            or id_to_cluster.get(str(item_id), "")
        )

    def add_impact(bucket, name, change_type):
        name = str(name or "").strip() or "Unspecified"
        rec = bucket.setdefault(name, {"new": 0, "delete": 0, "net": 0})
        if change_type == "new":
            rec["new"] += 1
            rec["net"] += 1
        elif change_type == "delete":
            rec["delete"] += 1
            rec["net"] -= 1

    def sales_flag(item_id):
        metric = metric_by_id.get(item_id, None) if item_id in metric_by_id.index else None
        flag = "Normal"
        if metric is not None and pd.notna(metric):
            if top_cut is not None and metric >= top_cut:
                flag = "Top 10% best seller"
            elif low_cut is not None and metric <= low_cut:
                flag = "Top 10% lowest seller"
        return flag, metric

    for rk, status in statuses.items():
        dg, item_id, item = _rk_parts(rk)
        status_s = str(status or "").strip().upper()
        if status_s == "MAINTAIN":
            continue
        related_pogs = sorted(actions.get(rk, {}).keys())
        if status_s == "DELETE ALL" and not related_pogs:
            related_pogs = id_to_pogs.get(item_id, [])
        related_pogs = related_pogs or ["Unspecified"]
        change_type = "new" if "NEW" in status_s else "delete" if "DELETE" in status_s else "other"
        for pog in related_pogs:
            if change_type in ("new", "delete"):
                add_impact(planogram_impact, pog, change_type)
                add_impact(cluster_impact, _cluster_for_pog(item_id, pog), change_type)
        flag, metric = sales_flag(item_id)
        item_changes.append({
            "status": status_s, "dg": dg, "id": item_id, "item": item,
            "planograms": related_pogs, "sales_flag": flag, "metric": metric,
        })

    for rk, acts in actions.items():
        if rk in statuses:
            continue
        dg, item_id, item = _rk_parts(rk)
        del_pogs = [p for p, a in acts.items() if str(a).strip().lower() == "delete"]
        new_pogs = [p for p, a in acts.items() if str(a).strip().lower() == "new"]
        if not del_pogs and not new_pogs:
            continue
        for pog in del_pogs:
            add_impact(planogram_impact, pog, "delete")
            add_impact(cluster_impact, _cluster_for_pog(item_id, pog), "delete")
        for pog in new_pogs:
            add_impact(planogram_impact, pog, "new")
            add_impact(cluster_impact, _cluster_for_pog(item_id, pog), "new")
        flag, metric = sales_flag(item_id)
        item_changes.append({
            "status": "DELETE SOME" if del_pogs else "NEW SOME",
            "dg": dg, "id": item_id, "item": item,
            "planograms": del_pogs + new_pogs, "sales_flag": flag, "metric": metric,
        })

    to_be = max(0, total_sku - status_counts.get("DELETE ALL", 0) + status_counts.get("NEWNEW", 0))
    total_deleted_units = 0.0
    total_added_units = 0.0
    for _chg in item_changes:
        try:
            _metric_val = _chg.get("metric")
            _metric_val = float(_metric_val) if _metric_val is not None and pd.notna(_metric_val) else 0.0
        except Exception:
            _metric_val = 0.0
        _status_val = str(_chg.get("status", "")).upper()
        if "DELETE" in _status_val:
            total_deleted_units += _metric_val
        if "NEW" in _status_val:
            total_added_units += _metric_val
    net_productivity_units = total_added_units - total_deleted_units
    risks = []
    if any(c["sales_flag"] == "Top 10% best seller" and "DELETE" in c["status"] for c in item_changes):
        risks.append("Top 10% best seller item removed")
    if any(c["sales_flag"] == "Top 10% lowest seller" and "NEW" in c["status"] for c in item_changes):
        risks.append("Top 10% lowest seller item added")
    if any(c["status"] == "DELETE ALL" for c in item_changes):
        risks.append("At least one item removed from all planograms")
    risk_level = "Need Review" if risks else ("Ready to Review" if item_changes else "No Change")
    return {
        "label": label, "submitted_at": submitted_at, "total_sku": total_sku, "to_be": to_be, "net": to_be - total_sku,
        "status_counts": status_counts, "item_changes": item_changes,
        "planogram_impact": planogram_impact, "cluster_impact": cluster_impact,
        "risks": risks, "risk_level": risk_level, "metric_col": metric_col or "",
        "sales_productivity": {
            "total_deleted_units": total_deleted_units,
            "total_added_units": total_added_units,
            "net_productivity_units": net_productivity_units,
        },
    }


def _rows_html(headers, rows):
    th = "".join(f"<th>{_h(x)}</th>" for x in headers)
    body = ""
    for row in rows:
        body += "<tr>" + "".join(f"<td>{_h(x)}</td>" for x in row) + "</tr>"
    if not rows:
        body = f"<tr><td colspan='{len(headers)}' class='muted'>No data</td></tr>"
    return f"<table class='a4-table'><thead><tr>{th}</tr></thead><tbody>{body}</tbody></table>"


def _render_a4_report(report: dict) -> str:
    sc = report["status_counts"]
    logo_uri = _report_partner_logo_data_uri()
    top_pogs = sorted(report["planogram_impact"].items(), key=lambda kv: abs(kv[1]["net"]) + kv[1]["new"] + kv[1]["delete"], reverse=True)[:10]
    top_clusters = sorted(report["cluster_impact"].items(), key=lambda kv: abs(kv[1]["net"]) + kv[1]["new"] + kv[1]["delete"], reverse=True)[:10]
    def _sales_risk_reason(c: dict) -> str:
        _status = str(c.get("status", "")).upper()
        _flag = str(c.get("sales_flag", "")).strip()
        if "DELETE" in _status and _flag == "Top 10% best seller":
            return "Deleting best seller"
        if "NEW" in _status and _flag == "Top 10% lowest seller":
            return "Adding lowest seller"
        return ""
    risk_items = [
        {**c, "risk_reason": _sales_risk_reason(c)}
        for c in report["item_changes"]
    ]
    risk_items = [c for c in risk_items if c.get("risk_reason")][:30]
    _best_delete_items = [c for c in risk_items if c.get("risk_reason") == "Deleting best seller"]
    _lowest_add_items = [c for c in risk_items if c.get("risk_reason") == "Adding lowest seller"]
    def _metric_num(c: dict) -> float:
        try:
            _v = c.get("metric")
            return float(_v) if _v is not None and pd.notna(_v) else 0.0
        except Exception:
            return 0.0
    _pog_risk = {}
    for _c in risk_items:
        _sign = -1 if _c.get("risk_reason") == "Deleting best seller" else 1
        for _pog in _c.get("planograms", []) or ["Unspecified"]:
            _pog_s = str(_pog or "Unspecified")
            _pog_risk[_pog_s] = _pog_risk.get(_pog_s, 0.0) + (_sign * _metric_num(_c))
    _highest_risk_pog = ""
    if _pog_risk:
        _highest_risk_pog = min(_pog_risk.items(), key=lambda kv: kv[1])[0]
    _sales_prod = report.get("sales_productivity", {}) or {}
    _total_deleted_units = float(_sales_prod.get("total_deleted_units", 0.0) or 0.0)
    _total_added_units = float(_sales_prod.get("total_added_units", 0.0) or 0.0)
    _net_productivity_units = float(_sales_prod.get("net_productivity_units", _total_added_units - _total_deleted_units) or 0.0)
    _productivity_direction = "Negative" if _net_productivity_units < 0 else "Positive" if _net_productivity_units > 0 else "Neutral"
    _productivity_rows = [
        ["Total deleted item volume", _fmt_num(_total_deleted_units)],
        ["Total added item forecast / volume", _fmt_num(_total_added_units)],
        ["Net productivity impact", _fmt_num(_net_productivity_units)],
        ["Productivity direction", _productivity_direction],
    ]
    _risk_summary_rows = [
        ["Best-seller delete count", len(_best_delete_items)],
        ["Lowest-seller add count", len(_lowest_add_items)],
        ["Best-seller delete volume", _fmt_num(sum(_metric_num(c) for c in _best_delete_items))],
        ["Lowest-seller add volume", _fmt_num(sum(_metric_num(c) for c in _lowest_add_items))],
        ["Highest risk planogram", _highest_risk_pog or "N/A"],
    ]
    def _impact_status(v: dict) -> str:
        _new = int(v.get("new", 0) or 0)
        _delete = int(v.get("delete", 0) or 0)
        if _new and _delete:
            return "Mixed"
        if _new:
            return "New"
        if _delete:
            return "Delete"
        return "No Change"
    pog_rows = [[pog, _impact_status(v), v["new"], v["delete"], v["net"]] for pog, v in top_pogs]
    cluster_rows = [[cl, v["new"], v["delete"], v["net"]] for cl, v in top_clusters]
    item_rows = [[c["risk_reason"], c["status"], f"{c['dg']}-{c['id']}-{c['item']}", ", ".join(map(str, c["planograms"][:5])), c["sales_flag"], _fmt_num(c["metric"]) if c["metric"] is not None and pd.notna(c["metric"]) else ""] for c in risk_items]
    risks = "; ".join(report["risks"]) if report["risks"] else "No high-risk movement detected."
    recommendation = (
        f"This DG has a {_productivity_direction.lower()} overall sales productivity signal. "
        f"Across all changed items, deleted volume is {_fmt_num(_total_deleted_units)} and added forecast/volume is {_fmt_num(_total_added_units)}, "
        f"giving a net productivity impact of {_fmt_num(_net_productivity_units)} units."
        if item_rows else
        "No major sales-risk movement was detected from the submitted changes."
    )
    generated = datetime.now().strftime("%d %b %Y %H:%M")
    risk_page_html = ""
    if item_rows:
        risk_page_html = f"""
<div class="a4-page">
<img class="a4-logo" src="{logo_uri}" alt="Lotus's">
<div class="a4-page-no">Page 2</div>
<div class="a4-kicker">Product Movement</div>
<div class="a4-title">{_h(report['label'])}</div>
<div class="a4-section"><h3>Sales Productivity Impact</h3><div class="a4-callout risk"><b>Overall productivity view: {_h(_productivity_direction)}</b><br>{_h(recommendation)}</div>{_rows_html(["Metric","Value"], _productivity_rows)}</div>
<div class="a4-section"><h3>Sales Risk Review</h3>{_rows_html(["Risk reason","Action","DG-ID-Item","Planogram","Sales flag", report.get("metric_col") or "Metric"], item_rows)}</div>
<div class="a4-section"><h3>Business Insight</h3><div class="a4-callout risk"><b>Risk view: {_h(report['risk_level'])}</b><br>The table above highlights only movements that are commercially sensitive: deleting a Top 10% best seller or adding a Top 10% lowest seller. These risk items should be reviewed separately from the overall productivity impact because they are a concern subset, not the full net calculation.</div>{_rows_html(["Metric","Value"], _risk_summary_rows)}</div>
</div>"""
    return f"""
<style>
.a4-shell{{display:block;margin-top:12px}}
.a4-page{{position:relative;width:794px;min-height:1123px;background:#fff;color:#1A1A1A;border:1px solid #DDD;box-shadow:0 10px 28px rgba(0,0,0,.10);padding:34px 38px;box-sizing:border-box;font-family:Arial,sans-serif;margin:0 0 24px 0;page-break-after:always}}
.a4-page:last-child{{page-break-after:auto}}
.a4-kicker{{font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:#2BBFA4;font-weight:800}}
.a4-title{{font-size:25px;font-weight:800;margin:6px 0 2px}}
.a4-sub{{font-size:12px;color:#777;margin-bottom:14px}}
.a4-section{{margin-top:14px;padding-top:10px;border-top:1px solid #ECE7E0}}
.a4-section h3{{margin:0 0 10px;font-size:15px}}
.a4-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:10px 0 6px}}
.a4-metric{{border:1px solid #E8E3DC;border-radius:8px;padding:12px;background:#FAFAF8}}
.a4-metric b{{display:block;font-size:20px}}
.a4-metric span{{font-size:10px;color:#777;text-transform:uppercase;letter-spacing:.06em}}
.a4-callout{{background:#F7FBFA;border-left:4px solid #2BBFA4;padding:12px 14px;border-radius:6px;font-size:12px;line-height:1.55}}
.risk{{border-left-color:#F59E0B;background:#FFF8E8}}
.a4-table{{width:100%;border-collapse:collapse;font-size:11px;margin-top:8px}}
.a4-table th{{background:#F0F2F2;text-align:left;padding:6px;border:1px solid #D8DDE0}}
.a4-table td{{padding:5px 6px;border:1px solid #E1E4E6;vertical-align:top}}
.muted{{color:#999;font-style:italic}}
.badge{{display:inline-block;padding:5px 8px;border-radius:999px;background:#EAF8F4;color:#178D7A;font-weight:700;font-size:11px}}
.a4-logo{{position:absolute;right:38px;top:18px;width:78px;height:auto}}
.a4-page-no{{position:absolute;right:38px;top:54px;color:#999;font-size:11px}}
@media print{{.a4-page{{box-shadow:none;border:none;width:210mm;min-height:297mm;margin:0;page-break-after:always}}}}
</style>
<div class="a4-shell">
<div class="a4-page">
<img class="a4-logo" src="{logo_uri}" alt="Lotus's">
<div class="a4-page-no">Page 1</div>
<div class="a4-kicker">Range Change Management Report</div>
<div class="a4-title">{_h(report['label'])}</div>
<div class="a4-sub">Generated {generated} | Submitted { _h(report.get('submitted_at') or '-') } | Source: Rangesheet Review submitted data and saved actions</div>
<div class="a4-callout"><b>Executive Summary</b><br>This range change has a net SKU impact of <b>{_h(report['net'])}</b>. Current decision view: <span class="badge">{_h(report['risk_level'])}</span>. {_h(risks)}</div>
<div class="a4-grid">
<div class="a4-metric"><span>AS-IS SKU</span><b>{_fmt_int(report['total_sku'])}</b></div>
<div class="a4-metric"><span>TO-BE SKU</span><b>{_fmt_int(report['to_be'])}</b></div>
<div class="a4-metric"><span>Net Change</span><b>{_h(report['net'])}</b></div>
<div class="a4-metric"><span>Risk Level</span><b style="font-size:15px">{_h(report['risk_level'])}</b></div>
</div>
<div class="a4-section"><h3>AS-IS vs TO-BE Status</h3>{_rows_html(["Status","SKU Count"], [["Maintain",sc.get("MAINTAIN",0)],["Delete Some",sc.get("DELETE SOME",0)],["Delete All",sc.get("DELETE ALL",0)],["New Some",sc.get("NEW SOME",0)],["NewNew",sc.get("NEWNEW",0)]])}</div>
<div class="a4-section"><h3>Planogram Impact</h3>{_rows_html(["Planogram","Status","New SKU","Delete SKU","Net"], pog_rows)}</div>
<div class="a4-section"><h3>Cluster Impact</h3>{_rows_html(["Cluster","New SKU","Delete SKU","Net"], cluster_rows)}</div>
</div>
{risk_page_html}</div>
"""


def _render_portfolio_summary(reports: list[dict], full_df: pd.DataFrame | None = None) -> str:
    if not reports:
        return ""
    logo_uri = _report_partner_logo_data_uri()
    submitted_as_is = sum(int(r.get("total_sku", 0) or 0) for r in reports)
    total_as_is = _total_sku_all_dg(full_df)
    if total_as_is <= 0:
        total_as_is = submitted_as_is
    total_net = sum(int(r.get("net", 0) or 0) for r in reports)
    total_to_be = max(0, total_as_is + total_net)
    total_changes = sum(len(r.get("item_changes", [])) for r in reports)
    need_review = sum(1 for r in reports if r.get("risk_level") == "Need Review")
    rows = [
        [
            r.get("label", ""),
            _fmt_int(r.get("total_sku", 0)),
            _fmt_int(r.get("to_be", 0)),
            r.get("net", 0),
            len(r.get("item_changes", [])),
            r.get("risk_level", ""),
        ]
        for r in reports
    ]
    generated = datetime.now().strftime("%d %b %Y %H:%M")
    return f"""
<style>
.a4-shell{{display:block;margin-top:12px}}
.a4-page{{position:relative;width:794px;min-height:1123px;background:#fff;color:#1A1A1A;border:1px solid #DDD;box-shadow:0 10px 28px rgba(0,0,0,.10);padding:34px 38px;box-sizing:border-box;font-family:Arial,sans-serif;margin:0 0 24px 0;page-break-after:always}}
.a4-kicker{{font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:#2BBFA4;font-weight:800}}
.a4-title{{font-size:25px;font-weight:800;margin:6px 0 2px}}
.a4-sub{{font-size:12px;color:#777;margin-bottom:20px}}
.a4-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:12px 0 8px}}
.a4-metric{{border:1px solid #E8E3DC;border-radius:8px;padding:12px;background:#FAFAF8}}
.a4-metric b{{display:block;font-size:20px}}
.a4-metric span{{font-size:10px;color:#777;text-transform:uppercase;letter-spacing:.06em}}
.a4-callout{{background:#F7FBFA;border-left:4px solid #2BBFA4;padding:12px 14px;border-radius:6px;font-size:12px;line-height:1.55}}
.a4-table{{width:100%;border-collapse:collapse;font-size:11px;margin-top:8px}}
.a4-table th{{background:#F0F2F2;text-align:left;padding:8px;border:1px solid #D8DDE0}}
.a4-table td{{padding:7px 8px;border:1px solid #E1E4E6;vertical-align:top}}
.a4-logo{{position:absolute;right:38px;top:18px;width:78px;height:auto}}
.a4-page-no{{position:absolute;right:38px;top:54px;color:#999;font-size:11px}}
@media print{{.a4-page{{box-shadow:none;border:none;width:210mm;min-height:297mm;margin:0;page-break-after:always}}}}
</style>
<div class="a4-shell">
<div class="a4-page">
<img class="a4-logo" src="{logo_uri}" alt="Lotus's">
<div class="a4-page-no">Portfolio Summary</div>
<div class="a4-kicker">Range Change Portfolio</div>
<div class="a4-title">Submitted DG Summary</div>
<div class="a4-sub">Generated {generated} | {len(reports)} submitted DG report(s)</div>
<div class="a4-callout"><b>Overall summary</b><br>AS-IS SKU is counted from all DGs in the source data. TO-BE SKU applies only the submitted DG changes to that full AS-IS base. Use the following DG pages for item-level review and team action.</div>
<div class="a4-grid">
<div class="a4-metric"><span>Total AS-IS SKU</span><b>{_fmt_int(total_as_is)}</b></div>
<div class="a4-metric"><span>Total TO-BE SKU</span><b>{_fmt_int(total_to_be)}</b></div>
<div class="a4-metric"><span>Net Change</span><b>{total_to_be - total_as_is}</b></div>
<div class="a4-metric"><span>Need Review DG</span><b>{need_review}</b></div>
</div>
<div class="a4-section"><h3>DG Rollup</h3>{_rows_html(["DG / Source","AS-IS SKU","TO-BE SKU","Net","Changed items","Risk"], rows)}</div>
</div>
</div>
"""


def _wrap_pdf_text(draw, text: str, font, max_width: int) -> list[str]:
    words = str(text or "").replace("\n", " ").split()
    if not words:
        return [""]
    lines, cur = [], ""
    for word in words:
        test = word if not cur else f"{cur} {word}"
        if draw.textlength(test, font=font) <= max_width:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = word
    if cur:
        lines.append(cur)
    return lines


def _reports_to_pdf_bytes(reports: list[dict], full_df: pd.DataFrame | None = None) -> bytes:
    from PIL import Image, ImageDraw, ImageFont

    page_w, page_h = 1240, 1754
    margin = 72
    line_h = 28
    try:
        font = ImageFont.truetype("tahoma.ttf", 22)
        font_b = ImageFont.truetype("tahomabd.ttf", 24)
        font_h = ImageFont.truetype("tahomabd.ttf", 34)
        font_s = ImageFont.truetype("tahoma.ttf", 18)
    except Exception:
        font = ImageFont.load_default()
        font_b = font_h = font_s = font

    pages = []
    try:
        logo_img = Image.open(io.BytesIO(_report_partner_logo_png_bytes(118, 36))).convert("RGBA")
    except Exception:
        logo_img = None

    def new_page():
        img = Image.new("RGB", (page_w, page_h), "white")
        if logo_img is not None:
            img.paste(logo_img, (page_w - margin - logo_img.width, 28), logo_img)
        return img, ImageDraw.Draw(img), margin

    def put(draw, y, text, use_font=None, fill=(25, 25, 25), max_width=None):
        use_font = use_font or font
        max_width = max_width or (page_w - margin * 2)
        for ln in _wrap_pdf_text(draw, text, use_font, max_width):
            draw.text((margin, y), ln, font=use_font, fill=fill)
            y += line_h
        return y

    total_as_is = _total_sku_all_dg(full_df)
    if total_as_is <= 0:
        total_as_is = sum(int(r.get("total_sku", 0) or 0) for r in reports)
    total_net = sum(int(r.get("net", 0) or 0) for r in reports)
    total_to_be = max(0, total_as_is + total_net)
    need_review = sum(1 for r in reports if r.get("risk_level") == "Need Review")

    img, draw, y = new_page()
    y = put(draw, y, "Range Change Portfolio Summary", font_h)
    y += 18
    y = put(draw, y, f"Generated {datetime.now().strftime('%d %b %Y %H:%M')}", font_s, (90, 90, 90))
    y += 18
    for label, val in [
        ("Total AS-IS SKU", _fmt_int(total_as_is)),
        ("Total TO-BE SKU", _fmt_int(total_to_be)),
        ("Net Change", total_to_be - total_as_is),
        ("Need Review DG", need_review),
    ]:
        y = put(draw, y, f"{label}: {val}", font_b)
    y += 18
    y = put(draw, y, "DG Rollup", font_b)
    for r in reports:
        line = (
            f"{r.get('label','')} | AS-IS {_fmt_int(r.get('total_sku',0))} | "
            f"TO-BE {_fmt_int(r.get('to_be',0))} | Net {r.get('net',0)} | "
            f"Changed {len(r.get('item_changes', []))} | {r.get('risk_level','')}"
        )
        y = put(draw, y, line, font_s)
        if y > page_h - margin - 80:
            pages.append(img)
            img, draw, y = new_page()
    pages.append(img)

    for r in reports:
        _pdf_risk_items = []
        for c in (r.get("item_changes", []) or []):
            _status = str(c.get("status", "")).upper()
            _flag = str(c.get("sales_flag", "")).strip()
            _reason = ""
            if "DELETE" in _status and _flag == "Top 10% best seller":
                _reason = "Deleting best seller"
            elif "NEW" in _status and _flag == "Top 10% lowest seller":
                _reason = "Adding lowest seller"
            if _reason:
                _pdf_risk_items.append((c, _reason))
        if not _pdf_risk_items:
            continue
        _pdf_sales_prod = r.get("sales_productivity", {}) or {}
        _pdf_total_deleted_units = float(_pdf_sales_prod.get("total_deleted_units", 0.0) or 0.0)
        _pdf_total_added_units = float(_pdf_sales_prod.get("total_added_units", 0.0) or 0.0)
        _pdf_net_productivity_units = float(_pdf_sales_prod.get("net_productivity_units", _pdf_total_added_units - _pdf_total_deleted_units) or 0.0)
        _pdf_productivity_direction = "Negative" if _pdf_net_productivity_units < 0 else "Positive" if _pdf_net_productivity_units > 0 else "Neutral"
        _pdf_pog_risk = {}
        _pdf_best_delete_volume = 0.0
        _pdf_lowest_add_volume = 0.0
        for c, _reason in _pdf_risk_items:
            try:
                _metric = c.get("metric")
                _metric = float(_metric) if _metric is not None and pd.notna(_metric) else 0.0
            except Exception:
                _metric = 0.0
            if _reason == "Deleting best seller":
                _pdf_best_delete_volume += _metric
                _sign = -1
            else:
                _pdf_lowest_add_volume += _metric
                _sign = 1
            for _pog in c.get("planograms", []) or ["Unspecified"]:
                _pog_s = str(_pog or "Unspecified")
                _pdf_pog_risk[_pog_s] = _pdf_pog_risk.get(_pog_s, 0.0) + (_sign * _metric)
        _pdf_highest_risk_pog = min(_pdf_pog_risk.items(), key=lambda kv: kv[1])[0] if _pdf_pog_risk else "N/A"
        img, draw, y = new_page()
        y = put(draw, y, str(r.get("label", "")), font_h)
        y += 10
        y = put(draw, y, f"Submitted: {r.get('submitted_at') or '-'}", font_s, (90, 90, 90))
        y += 14
        y = put(draw, y, f"AS-IS SKU: {_fmt_int(r.get('total_sku', 0))}", font_b)
        y = put(draw, y, f"TO-BE SKU: {_fmt_int(r.get('to_be', 0))}", font_b)
        y = put(draw, y, f"Net Change: {r.get('net', 0)}", font_b)
        y = put(draw, y, f"Risk Level: {r.get('risk_level', '')}", font_b)
        risks = "; ".join(r.get("risks", [])) if r.get("risks") else "No high-risk movement detected."
        y += 10
        y = put(draw, y, f"Risks: {risks}", font)
        y += 18
        y = put(draw, y, "Sales Productivity Impact", font_b)
        for label, val in [
            ("Total deleted item volume", _fmt_num(_pdf_total_deleted_units)),
            ("Total added item forecast / volume", _fmt_num(_pdf_total_added_units)),
            ("Net productivity impact", _fmt_num(_pdf_net_productivity_units)),
            ("Productivity direction", _pdf_productivity_direction),
        ]:
            y = put(draw, y, f"{label}: {val}", font_s)
            if y > page_h - margin - 80:
                pages.append(img)
                img, draw, y = new_page()
        y += 18
        y = put(draw, y, "Sales Risk Review", font_b)
        for c, _reason in _pdf_risk_items[:25]:
            line = (
                f"{_reason} | {c.get('status','')} | {c.get('dg','')}-{c.get('id','')}-{c.get('item','')} | "
                f"{', '.join(map(str, c.get('planograms', [])[:3]))} | {c.get('sales_flag','')}"
            )
            y = put(draw, y, line, font_s)
            if y > page_h - margin - 80:
                pages.append(img)
                img, draw, y = new_page()
        y += 18
        y = put(draw, y, "Business Insight", font_b)
        for label, val in [
            ("Best-seller delete count", sum(1 for _, reason in _pdf_risk_items if reason == "Deleting best seller")),
            ("Lowest-seller add count", sum(1 for _, reason in _pdf_risk_items if reason == "Adding lowest seller")),
            ("Best-seller delete volume", _fmt_num(_pdf_best_delete_volume)),
            ("Lowest-seller add volume", _fmt_num(_pdf_lowest_add_volume)),
            ("Highest risk planogram", _pdf_highest_risk_pog),
        ]:
            y = put(draw, y, f"{label}: {val}", font_s)
            if y > page_h - margin - 80:
                pages.append(img)
                img, draw, y = new_page()
        pages.append(img)

    buf = io.BytesIO()
    pages[0].save(buf, format="PDF", save_all=True, append_images=pages[1:], resolution=150.0)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# Tab 1: Range Execution — 4 output slots sourced from View Data submissions
# ══════════════════════════════════════════════════════════════════════════════
with rt0:
    st.markdown("""
<div style="margin-bottom:18px;">
  <div style="font-size:18px;font-weight:800;color:#1A1A1A;">Execution Report</div>
  <div style="font-size:13px;color:#777;margin-top:4px;">
    A4 business summary for managers and cross-functional review. Submit data from Rangesheet Review first, then generate the report here.
  </div>
</div>""", unsafe_allow_html=True)

    _sources = [
        ("ss", "Range Sheet SSPOG"),
        ("ns", "Range Sheet Non-SSPOG"),
        ("sa", "StoreApply POG"),
    ]
    _source_labels = ["All submitted DGs"]
    _source_options = [("all", "All submitted DGs")]
    for _p_src, _label_src in _sources:
        _df_src = st.session_state.get(f"vw_submit_{_p_src}")
        _packets_src = st.session_state.get(f"vw_submit_reports_{_p_src}", {})
        _has_packets = isinstance(_packets_src, dict) and bool(_packets_src)
        _has_df = _df_src is not None and not _df_src.empty
        _has_edits = bool(st.session_state.get(f"{_p_src}_status_overrides") or st.session_state.get(f"{_p_src}_pog_actions"))
        _source_options.append((_p_src, _label_src))
        _source_labels.append(f"{_label_src}{' (ready)' if (_has_packets or _has_df or _has_edits) else ' (no submitted data)'}")

    _preferred_source = st.session_state.get("management_report_source")
    _preferred_idx = next(
        (
            i for i, (_src_key, _) in enumerate(_source_options)
            if _src_key == _preferred_source and (_src_key == "all" or "(ready)" in _source_labels[i])
        ),
        None,
    )
    _ready_idx = _preferred_idx if _preferred_idx is not None else 0
    _sel_label = st.selectbox(
        "Report source",
        _source_labels,
        index=_ready_idx,
        label_visibility="collapsed",
    )
    _sel_idx = _source_labels.index(_sel_label)
    _selected_report_source, _selected_report_label = _source_options[_sel_idx]

    def _reports_for_source(_p_src: str, _label_src: str) -> list[dict]:
        _reports = []
        _packets = st.session_state.get(f"vw_submit_reports_{_p_src}", {})
        if isinstance(_packets, dict) and _packets:
            _packet_items = sorted(
                _packets.items(),
                key=lambda kv: str((kv[1] or {}).get("submitted_order") or (kv[1] or {}).get("submitted_at") or kv[0]),
            )
            _latest_by_dg = {}
            for _packet_key, _packet in _packet_items:
                if not isinstance(_packet, dict):
                    continue
                _df_packet = _packet.get("df")
                if _df_packet is None or getattr(_df_packet, "empty", True):
                    continue
                _dg_key = str(_packet.get("dg_key") or _packet.get("label") or _packet_key)
                if _dg_key in _latest_by_dg:
                    _latest_by_dg.pop(_dg_key, None)
                _latest_by_dg[_dg_key] = (_packet_key, _packet)
            for _packet_key, _packet in _latest_by_dg.values():
                _df_packet = _packet.get("df")
                if _df_packet is None or getattr(_df_packet, "empty", True):
                    continue
                _packet_label = f"{_label_src} | {_packet.get('label') or _packet_key}"
                _reports.append(_build_management_report(
                    _p_src,
                    _packet_label,
                    df_override=_df_packet,
                    submitted_at=str(_packet.get("submitted_at", "")),
                    pog_cluster_map=_packet.get("pog_cluster_map", {}),
                    dg_filter=_dg_code_from_packet_key(_packet.get("dg_key") or _packet.get("label") or _packet_key),
                ))
        else:
            _df_src = st.session_state.get(f"vw_submit_{_p_src}")
            if _df_src is not None and not _df_src.empty:
                _reports.append(_build_management_report(_p_src, _label_src))
        return _reports

    if _selected_report_source == "all":
        _reports = []
        for _p_src, _label_src in _sources:
            _reports.extend(_reports_for_source(_p_src, _label_src))
    else:
        _reports = _reports_for_source(_selected_report_source, _selected_report_label)
    _report_html = _render_portfolio_summary(_reports, merged) + "".join(_render_a4_report(_r) for _r in _reports)

    _r1, _r2, _r3 = st.columns([1.2, 1.2, 5])
    with _r1:
        st.download_button(
            "Download HTML",
            data=_report_html.encode("utf-8"),
            file_name=f"management_report_{_selected_report_source}_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
            mime="text/html",
            use_container_width=True,
        )
    with _r2:
        if _reports:
            st.download_button(
                "Export PDF",
                data=_reports_to_pdf_bytes(_reports, merged),
                file_name=f"management_report_{_selected_report_source}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        else:
            st.button("Export PDF", disabled=True, use_container_width=True)
    with _r3:
        st.caption("PDF exports the latest submitted version per DG.")

    if not _reports:
        st.info("No submitted range data or saved range changes found yet. Go to Rangesheet Review, apply a DG, make changes, then click Submit Range.")

    st.markdown(_report_html, unsafe_allow_html=True)


with rt1:
    st.markdown("""
<div style="margin-bottom:24px;">
    <div style="font-size:18px;font-weight:700;color:#1A1A1A;margin-bottom:4px;">
        Output from Rangesheet
    </div>
    <div style="font-size:13px;color:#888;">
        Submit from View Data or Rangesheet Review, then preview and export Excel here.
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

        if has_data:
            _rows_txt = f"{len(df_out):,} rows · {len(df_out.columns)} cols"
            _radius   = "14px"

            _is_selected = bool(st.session_state.get(_chk_key, False))
            if _is_selected:
                _sel_nums.append(num)
                _sel_outputs.append((num, title, df_out))

            # ── Card header ───────────────────────────────────────────────────
            st.markdown(f"""
<div style="background:#fff;border-radius:{_radius};border:1px solid #E8E3DC;
            border-bottom:1px solid #E8E3DC;
            border-left:4px solid {color};padding:16px 20px;">
    <div style="display:flex;align-items:center;gap:14px;justify-content:space-between;">
      <div style="display:flex;align-items:center;gap:14px;">
        <div style="width:34px;height:34px;background:{color}22;border-radius:8px;
                    display:flex;align-items:center;justify-content:center;
                    font-size:12px;font-weight:800;color:{color};flex-shrink:0;">{num}</div>
        <div>
            <div style="font-size:14px;font-weight:700;color:#1A1A1A;">{title}</div>
            <div style="font-size:11px;color:#999;margin-top:3px;">
                Submitted to Report &nbsp;·&nbsp; {_rows_txt}
            </div>
        </div>
      </div>
      <div style="font-size:11px;color:{color};font-weight:700;text-transform:uppercase;">
        Ready
      </div>
    </div>
</div>""", unsafe_allow_html=True)

            # ── Compact controls: keep submitted cards aligned with placeholders.
            _b0, _b1, _b2, _b3, _ = st.columns([0.35, 0.8, 0.75, 0.75, 6])
            with _b0:
                _picked = st.checkbox("", key=_chk_key, label_visibility="collapsed")
                if _picked and not _is_selected:
                    _sel_nums.append(num)
                    _sel_outputs.append((num, title, df_out))
            with _b1:
                if st.button(
                    "Preview",
                    key=f"prev_btn_{num}",
                    use_container_width=True,
                    type="secondary",
                ):
                    _all_cols = list(df_out.columns)
                    _vis_cols = _all_cols
                    _tbl_h = "72vh"
                    _fz = 12
                    _pad = "8px 14px"
                    _hpad = "10px 16px"
                    _prev_n = min(200, len(df_out))
                    _th = "".join(
                        f'<th style="background:#2BBFA4;color:#fff;font-weight:700;'
                        f'padding:{_hpad};white-space:nowrap;text-align:left;'
                        f'font-size:{_fz}px;letter-spacing:.03em;'
                        f'border-right:1px solid rgba(255,255,255,0.25);'
                        f'position:sticky;top:0;z-index:2;">{_h(c)}</th>'
                        for c in _vis_cols
                    )
                    _tbody = ""
                    for _ri, (_, _row) in enumerate(df_out.head(_prev_n).iterrows()):
                        _bg = "#fff" if _ri % 2 == 0 else "#F4FBF9"
                        _tds = "".join(
                            f'<td style="padding:{_pad};font-size:{_fz}px;color:#1A1A1A;'
                            f'white-space:nowrap;border-right:1px solid #EEE;'
                            f'border-bottom:1px solid #F0EBE3;">'
                            f'{_h(_row[c])}</td>'
                            for c in _vis_cols
                        )
                        _tbody += f'<tr style="background:{_bg};">{_tds}</tr>'
                    _preview_html = f"""
<div style="font-size:12px;color:#666;margin-bottom:10px;">
    First {_prev_n:,} of {len(df_out):,} rows &nbsp;·&nbsp; {len(_vis_cols)} cols
</div>
<div style="overflow:auto;max-height:{_tbl_h};border-radius:8px;border:1px solid #E8E3DC;background:white;">
    <table style="border-collapse:collapse;width:100%;min-width:900px;">
        <thead><tr>{_th}</tr></thead>
        <tbody>{_tbody}</tbody>
    </table>
</div>"""

                    if hasattr(st, "dialog"):
                        try:
                            _dialog_decorator = st.dialog(f"Preview — {title}", width="large")
                        except TypeError:
                            _dialog_decorator = st.dialog(f"Preview — {title}")

                        @_dialog_decorator
                        def _preview_dialog():
                            st.markdown(_preview_html, unsafe_allow_html=True)
                        _preview_dialog()
                    else:
                        st.markdown(_preview_html, unsafe_allow_html=True)
            with _b2:
                st.download_button(
                    ".xlsx", df_to_xlsx_bytes(df_out),
                    file_name=f"output_{num}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{num}x", use_container_width=True,
                )
            with _b3:
                st.download_button(
                    ".csv", df_to_csv_bytes(df_out),
                    file_name=f"output_{num}.csv",
                    mime="text/csv",
                    key=f"dl_{num}c", use_container_width=True,
                )

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
                Not submitted yet — submit data to Report first
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
# Hidden legacy Summary & Analytics block (tab removed from UI)
# ══════════════════════════════════════════════════════════════════════════════
if False:
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
