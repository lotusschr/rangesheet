"""My Files page — upload, merge, file management."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
from datetime import datetime
from utils.shared import (
    inject_css, init_session_state, render_sidebar, render_topbar, render_page_nav,
    APP_CONFIG, read_uploaded_file, auto_merge, save_merged_snapshot, clear_merged_snapshot,
    save_file, add_audit, current_user,
)

inject_css()
init_session_state()

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

# ── Per-session state for upload deduplication & duplicate handling ───────────
if "dup_pending"      not in st.session_state: st.session_state.dup_pending      = []
if "_upload_sig_last" not in st.session_state: st.session_state._upload_sig_last = set()

render_sidebar("landpage")
render_topbar("My Files")

# ── Duplicate file dialog ──────────────────────────────────────────────────────
@st.dialog("Duplicate File Found")
def _dup_dialog():
    if not st.session_state.dup_pending:
        st.rerun()
        return
    _dup  = st.session_state.dup_pending[0]
    _name = _dup["name"]
    st.markdown(f"""
<div style="text-align:center;padding:4px 0 20px;">
    <div style="font-size:36px;margin-bottom:12px;">📄</div>
    <div style="font-size:15px;font-weight:700;color:#1A1A1A;margin-bottom:8px;line-height:1.4;">
        <strong style="color:#2BBFA4;">{_name}</strong><br>already exists.
    </div>
    <div style="font-size:13px;color:#888;">
        Do you want to <strong>Add</strong> it alongside the existing file,
        or <strong>Replace</strong> the existing one?
    </div>
</div>""", unsafe_allow_html=True)
    _al, _rr = st.columns(2)
    with _al:
        st.markdown('<div class="dup-add-btn">', unsafe_allow_html=True)
        if st.button("Add", key="dup_add", use_container_width=True):
            _base, _ext = os.path.splitext(_name)
            _cnt = sum(1 for f in st.session_state.raw_files
                       if f["name"].startswith(_base) and f["name"].endswith(_ext))
            _new_entry       = dict(_dup)
            _new_entry["name"] = f"{_base} ({_cnt + 1}){_ext}"
            st.session_state.raw_files.append(_new_entry)
            st.session_state.dup_pending.pop(0)
            add_audit("Upload (Add duplicate)", _new_entry["name"])
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with _rr:
        st.markdown('<div class="dup-replace-btn">', unsafe_allow_html=True)
        if st.button("Replace", key="dup_replace", use_container_width=True, type="primary"):
            for _i, _f in enumerate(st.session_state.raw_files):
                if _f["name"] == _name:
                    st.session_state.raw_files[_i] = dict(_dup)
                    break
            st.session_state.dup_pending.pop(0)
            add_audit("Upload (Replace)", _name)
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)


# ── Custom upload zone style ──────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stFileUploadDropzone"] {
    border: 2.5px dashed #C0B8FF !important;
    border-radius: 20px !important;
    background: #fff !important;
    padding: 44px 24px 36px !important;
    text-align: center !important;
    cursor: pointer !important;
    transition: border-color .2s, background .2s !important;
}
[data-testid="stFileUploadDropzone"]:hover {
    border-color: #7C6FF7 !important;
    background: #F9F8FF !important;
}
[data-testid="stFileUploadDropzone"] svg {
    width: 52px !important;
    height: 52px !important;
    color: #7C6FF7 !important;
}
[data-testid="stFileUploadDropzone"] svg path {
    fill: #7C6FF7 !important;
    stroke: #7C6FF7 !important;
}
[data-testid="stFileUploadDropzone"] > div > span {
    font-size: 16px !important;
    font-weight: 600 !important;
    color: #1A1A1A !important;
    display: block !important;
    margin-bottom: 4px !important;
}
[data-testid="stFileUploadDropzone"] small {
    font-size: 12px !important;
    color: #999 !important;
}
[data-testid="stFileUploadDropzone"] button {
    background: #7C6FF7 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 12px !important;
    padding: 9px 36px !important;
    font-size: 14px !important;
    font-weight: 500 !important;
    text-transform: lowercase !important;
    margin-top: 14px !important;
    box-shadow: 0 2px 8px rgba(124,111,247,.3) !important;
}
[data-testid="stFileUploadDropzone"] button:hover {
    background: #6B5FF5 !important;
}
[data-testid="stFileUploader"] > label { display: none !important; }
</style>
""", unsafe_allow_html=True)

st.markdown(
    "<div style='font-size:28px;font-weight:800;color:#1A1A1A;margin-bottom:24px;'>Files</div>",
    unsafe_allow_html=True,
)

col_upload, col_files = st.columns([1, 1.6], gap="large")

# ── Left: Upload zone ─────────────────────────────────────────────────────────
with col_upload:
    st.markdown("""
<div style="font-size:12px;color:#999;margin-bottom:8px;text-align:center;">
    Max <strong style="color:#555;">5 files</strong> per upload &nbsp;·&nbsp;
    Each file max <strong style="color:#555;">1 GB</strong>
    &nbsp;·&nbsp; csv, xls, xlsx, xlsb accepted
</div>""", unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "upload",
        type=APP_CONFIG["allowed_extensions"],
        accept_multiple_files=True,
        label_visibility="collapsed",
        key=f"uploader_{st.session_state.uploader_key}",
    )

    st.markdown("<div style='height:4px;'></div>", unsafe_allow_html=True)

    # ── Process new uploads only (signature-based to avoid rerun re-processing)
    _current_sig = {(f.name, f.size) for f in (uploaded or [])}
    _new_sigs    = _current_sig - st.session_state._upload_sig_last

    if uploaded and _new_sigs:
        st.session_state._upload_sig_last = _current_sig

        _to_process = [f for f in uploaded if (f.name, f.size) in _new_sigs]

        # Enforce 5-file total limit
        _slots = max(0, 5 - len(st.session_state.raw_files))
        if len(_to_process) > _slots:
            st.warning(f"⚠️ Max 5 files total. Only {_slots} file(s) can be added.")
            _to_process = _to_process[:_slots]

        new_files, failed = [], []
        _existing_names = {x["name"] for x in st.session_state.raw_files}

        for f in _to_process:
            df = read_uploaded_file(f)
            if df is None:
                failed.append(f.name)
                continue
            f.seek(0)
            raw_bytes = f.read()
            path = save_file(raw_bytes, f.name)
            _entry = {
                "name":  f.name,
                "df":    df,
                "size":  f"{round(f.size / 1024, 1)} KB",
                "date":  datetime.now().strftime("%d/%m/%Y %H:%M"),
                "rows":  len(df),
                "cols":  len(df.columns),
                "saved": bool(path),
                "owner": current_user()["employee_id"],
            }
            if f.name in _existing_names:
                # Queue for duplicate dialog (avoid re-queuing same name)
                if not any(p["name"] == f.name for p in st.session_state.dup_pending):
                    st.session_state.dup_pending.append(_entry)
            else:
                new_files.append(_entry)

        if failed:
            st.error(f"Could not read: {', '.join(failed)}")
        if new_files:
            st.session_state.raw_files.extend(new_files)
            add_audit("Upload", f"{len(new_files)} new file(s) added")
            if not st.session_state.dup_pending:
                st.success(f"✅ Uploaded {len(new_files)} file(s)")
                st.rerun()
            else:
                st.success(f"✅ {len(new_files)} file(s) added. Handling duplicate(s)...")

    # Show duplicate dialog for first pending duplicate
    if st.session_state.dup_pending:
        _dup_dialog()

    # Clear All button
    # if st.session_state.raw_files:
    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
    if st.button("🗑️ Clear All Files", key="clear_all", use_container_width=True):

        # st.session_state.raw_files         = []
        # st.session_state.merged_df         = None

        clear_merged_snapshot()

        st.session_state.merged_df = None
        st.session_state.raw_files = []

        st.session_state.merge_log         = []
        st.session_state.display_cols      = None

        st.session_state._upload_sig_last  = set()
        st.session_state.dup_pending       = []

        st.session_state.pop("view_file_selection", None)

        # force reset uploader
        st.session_state.uploader_key += 1

        add_audit("Clear Files")
        st.rerun()

# ── Right: File table + selection ─────────────────────────────────────────────
with col_files:
    search_q = st.text_input(
        "search", placeholder="🔍  Search files...",
        label_visibility="collapsed",
    )

    files = st.session_state.raw_files
    if search_q:
        files = [f for f in files if search_q.lower() in f["name"].lower()]

    # ── File table ────────────────────────────────────────────────────────────
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
            No files yet — drop files on the left or click <strong>browse</strong>
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

    # ── File selection for View Data ──────────────────────────────────────────
    if st.session_state.raw_files:
        st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)

        _all_file_names = [f["name"] for f in st.session_state.raw_files]
        _selected = st.multiselect(
            "select_files_label",
            options=_all_file_names,
            key="view_file_selection",
            placeholder="📂  Select files to show in View Data...",
            label_visibility="collapsed",
        )

        if _selected:
            if st.button(
                "View Data →",
                key="go_to_viewdata",
                use_container_width=True,
                type="primary",
            ):
                _sel_raw = [f for f in st.session_state.raw_files if f["name"] in _selected]

                st.session_state.selected_files = _selected

                merged, log = auto_merge(_sel_raw)
                # st.session_state.merged_df    = merged
                st.session_state.upload_df = merged
                st.session_state.merge_log    = log
                st.session_state.display_cols = None
                save_merged_snapshot(merged)
                add_audit("View Data", f"{len(_selected)} file(s) selected")
                
                st.switch_page("pages/viewdata.py")
        else:
            # No files selected → clear any previously loaded data
            if st.session_state.merged_df is not None:
                st.session_state.merged_df = None
            st.markdown(
                "<div style='font-size:12px;color:#AAA;text-align:center;margin-top:6px;'>"
                "Select files above to enable View Data</div>",
                unsafe_allow_html=True,
            )

render_page_nav("landpage")
