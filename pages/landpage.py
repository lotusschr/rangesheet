"""My Files page — upload, merge, preview."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
import pandas as pd
from datetime import datetime
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar,
    APP_CONFIG, read_uploaded_file, auto_merge, save_merged_snapshot,
    save_file, add_audit, df_to_xlsx_bytes, df_to_csv_bytes,
)

inject_css()
init_session_state()
render_sidebar("landpage")
render_topbar("My Files")


# ── Preview dialog ────────────────────────────────────────────────────────────
@st.dialog("📄 File Preview", width="large")
def _preview_dialog():
    info = st.session_state.get("last_uploaded_preview")
    if info is None:
        st.info("No file to preview.")
        return
    df        = info["df"]
    filename  = info["name"]
    prev_cols = list(df.columns[:10])
    st.markdown(
        f"**{filename}** — {info['rows']:,} rows · {info['cols']} columns  "
        f"*(showing first 50 rows & {len(prev_cols)} columns)*"
    )
    st.dataframe(df[prev_cols].head(50), use_container_width=True,
                 hide_index=True, height=420)
    if st.button("✅ Close", use_container_width=True):
        st.session_state["_show_preview"] = False
        st.rerun()


# Trigger dialog if flagged
if st.session_state.get("_show_preview"):
    _preview_dialog()


# ── Page layout ───────────────────────────────────────────────────────────────
st.markdown("<div style='font-size:22px;font-weight:700;margin-bottom:2px;'>My Files</div>",
            unsafe_allow_html=True)
st.markdown("<div style='color:#888;font-size:13px;margin-bottom:18px;'>"
            "Upload .csv / .xlsx → auto-merge → save locally</div>",
            unsafe_allow_html=True)

col_left, col_right = st.columns([1, 1.8], gap="large")

# ── Left: upload ──────────────────────────────────────────────────────────────
with col_left:
    st.subheader("Upload Files")
    uploaded = st.file_uploader(
        "Drag & drop or browse",
        type=APP_CONFIG["allowed_extensions"],
        accept_multiple_files=True,
    )
    if uploaded:
        new_files, failed = [], []
        for f in uploaded:
            if any(x["name"] == f.name for x in st.session_state.raw_files):
                continue
            df = read_uploaded_file(f)
            if df is None:
                failed.append(f.name)
                continue
            f.seek(0)
            path = save_file(f.read(), f.name)
            new_files.append({
                "name": f.name, "df": df,
                "size": f"{round(f.size/1024, 1)} KB",
                "date": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "rows": len(df), "cols": len(df.columns),
                "saved": bool(path),
            })
        if failed:
            st.error(f"Could not read: {', '.join(failed)}")
        if new_files:
            st.session_state.raw_files.extend(new_files)
            merged, log = auto_merge(st.session_state.raw_files)
            st.session_state.merged_df  = merged
            st.session_state.merge_log  = log
            st.session_state.display_cols = None
            save_merged_snapshot(merged)
            add_audit("UPLOAD", f"{len(new_files)} files → {len(merged):,} rows")
            st.session_state["last_uploaded_preview"] = new_files[-1]
            st.session_state["_show_preview"] = True
            st.success(f"✅ Uploaded {len(new_files)} file(s) — {len(merged):,} rows merged")
            st.rerun()

    if st.session_state.merge_log:
        st.subheader("Merge Log")
        for entry in st.session_state.merge_log:
            st.write(entry)

    if st.session_state.raw_files:
        if st.button("🗑️ Clear all files", key="clear_all"):
            st.session_state.raw_files        = []
            st.session_state.merged_df        = None
            st.session_state.merge_log        = []
            st.session_state.display_cols     = None
            st.session_state["last_uploaded_preview"] = None
            add_audit("CLEAR FILES")
            st.rerun()

# ── Right: file list + preview ────────────────────────────────────────────────
with col_right:
    st.subheader("Uploaded Files")
    if not st.session_state.raw_files:
        st.info("No files yet — drag & drop on the left.")
    else:
        for f in st.session_state.raw_files:
            c_name, c_btn = st.columns([5, 1])
            with c_name:
                st.markdown(
                    f"📄 **{f['name']}** — {f['rows']:,} rows · "
                    f"{f['cols']} cols · {f['size']}")
            with c_btn:
                if st.button("👁️", key=f"prev_{f['name']}",
                             help="Preview this file"):
                    st.session_state["last_uploaded_preview"] = f
                    st.session_state["_show_preview"] = True
                    st.rerun()

    if st.session_state.merged_df is not None:
        m = st.session_state.merged_df
        st.divider()
        c1, c2, c3 = st.columns(3)
        c1.metric("Files",       len(st.session_state.raw_files))
        c2.metric("Merged Rows", f"{len(m):,}")
        c3.metric("Columns",     len(m.columns))

        st.divider()
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "⬇️ Export .xlsx", df_to_xlsx_bytes(m),
                file_name=f"merged_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        with d2:
            st.download_button(
                "⬇️ Export .csv", df_to_csv_bytes(m),
                file_name=f"merged_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv", use_container_width=True)
