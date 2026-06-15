"""My Files page — upload, merge, file management."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
from datetime import datetime
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar, render_page_nav,
    APP_CONFIG, read_uploaded_file, auto_merge, save_merged_snapshot,
    save_file, add_audit, df_to_xlsx_bytes, df_to_csv_bytes, current_user,
)

inject_css()
init_session_state()
render_sidebar("landpage")
render_topbar("My Files")

st.markdown(
    "<div style='font-size:28px;font-weight:800;color:#1A1A1A;margin-bottom:24px;'>Files</div>",
    unsafe_allow_html=True,
)

col_upload, col_files = st.columns([1, 1.6], gap="large")

# ── Left: Upload zone ─────────────────────────────────────────────────────────
with col_upload:
    uploaded = st.file_uploader(
        "Drop files here or browse\n\nUp to 5 files · .xlsx · .xls · .csv · Max 1 GB per file",
        type=APP_CONFIG["allowed_extensions"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    st.markdown("<div style='height:4px;'></div>", unsafe_allow_html=True)

    if uploaded:
        if len(uploaded) > 5:
            st.warning("⚠️ Maximum 5 files per upload. Only the first 5 will be processed.")
            uploaded = uploaded[:5]
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
                "name":  f.name,
                "df":    df,
                "size":  f"{round(f.size / 1024, 1)} KB",
                "date":  datetime.now().strftime("%d/%m/%Y %H:%M"),
                "rows":  len(df),
                "cols":  len(df.columns),
                "saved": bool(path),
                "owner": current_user()["employee_id"],
            })
        if failed:
            st.error(f"Could not read: {', '.join(failed)}")
        if new_files:
            st.session_state.raw_files.extend(new_files)
            merged, log = auto_merge(st.session_state.raw_files)
            st.session_state.merged_df   = merged
            st.session_state.merge_log   = log
            st.session_state.display_cols = None
            save_merged_snapshot(merged)
            add_audit("Upload & Auto-Merge",
                      f"{len(new_files)} files → {len(merged):,} rows")
            st.success(f"✅ Uploaded {len(new_files)} file(s) — {len(merged):,} rows merged")
            st.rerun()

    if st.session_state.raw_files:
        st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
        if st.button("🗑️ Clear All Files", key="clear_all", use_container_width=True):
            st.session_state.raw_files    = []
            st.session_state.merged_df    = None
            st.session_state.merge_log    = []
            st.session_state.display_cols = None
            add_audit("Clear Files")
            st.rerun()

# ── Right: File table ─────────────────────────────────────────────────────────
with col_files:
    search_q = st.text_input(
        "search", placeholder="🔍  Search files and folders...",
        label_visibility="collapsed",
    )

    files = st.session_state.raw_files
    if search_q:
        files = [f for f in files if search_q.lower() in f["name"].lower()]

    header_html = """
<div style="background:#fff;border-radius:16px;border:1px solid #E8E3DC;overflow:hidden;">
    <div style="display:grid;grid-template-columns:1fr 110px 90px 145px;
                padding:10px 18px;background:#F5F0EA;border-bottom:1px solid #E8E3DC;">
        <span style="font-size:11px;font-weight:700;color:#2BBFA4;
                     text-transform:uppercase;letter-spacing:0.06em;">File Name</span>
        <span style="font-size:11px;font-weight:700;color:#2BBFA4;
                     text-transform:uppercase;letter-spacing:0.06em;">User</span>
        <span style="font-size:11px;font-weight:700;color:#2BBFA4;
                     text-transform:uppercase;letter-spacing:0.06em;">Size</span>
        <span style="font-size:11px;font-weight:700;color:#2BBFA4;
                     text-transform:uppercase;letter-spacing:0.06em;">Last Modified</span>
    </div>"""

    if not files:
        table_html = header_html + """
    <div style="padding:60px 20px;text-align:center;">
        <div style="font-size:56px;margin-bottom:12px;">📁</div>
        <div style="color:#999;font-size:13px;">
            No files yet — drop files above or click <strong>New Upload</strong>
        </div>
    </div>
</div>"""
    else:
        rows = ""
        for i, f in enumerate(files):
            bg = "#fff" if i % 2 == 0 else "#FAFAF8"
            rows += f"""
<div style="display:grid;grid-template-columns:1fr 110px 90px 145px;
            padding:13px 18px;background:{bg};border-bottom:1px solid #F0EBE3;
            align-items:center;">
    <div style="display:flex;align-items:center;gap:10px;min-width:0;">
        <span style="font-size:20px;flex-shrink:0;">📄</span>
        <div style="min-width:0;">
            <div style="font-size:13px;font-weight:600;color:#1A1A1A;
                        white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
                {f['name']}
            </div>
            <div style="font-size:11px;color:#999;">{f['rows']:,} rows · {f['cols']} cols</div>
        </div>
    </div>
    <span style="font-size:12px;color:#555;">{f.get('owner','—')}</span>
    <span style="font-size:12px;color:#555;">{f['size']}</span>
    <span style="font-size:12px;color:#555;">{f['date']}</span>
</div>"""
        table_html = header_html + rows + "</div>"

    st.markdown(table_html, unsafe_allow_html=True)

    # Metrics + export when merged data exists
    if st.session_state.merged_df is not None:
        m = st.session_state.merged_df
        st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        c1.metric("Files",       len(st.session_state.raw_files))
        c2.metric("Merged Rows", f"{len(m):,}")
        c3.metric("Columns",     len(m.columns))

        st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "⬇️ Export .xlsx", df_to_xlsx_bytes(m),
                file_name=f"merged_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with d2:
            st.download_button(
                "⬇️ Export .csv", df_to_csv_bytes(m),
                file_name=f"merged_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True,
            )

render_page_nav("landpage")
