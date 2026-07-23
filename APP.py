"""RangeSheet — navigation entry point.
Run: streamlit run APP.py
"""
import streamlit as st
import importlib
import utils.shared as shared

shared = importlib.reload(shared)

st.set_page_config(
    page_title="RangeSheet",
    page_icon="🗂️",
    layout="wide",
    initial_sidebar_state="expanded",
)

shared.inject_css()
_render_login_page = getattr(shared, "render_login_page", None)
if _render_login_page is None:
    st.error("Login module is still loading. Please refresh the page once.")
    st.stop()
if not _render_login_page():
    st.stop()
shared.init_session_state()

pg = st.navigation(
    [
        st.Page("pages/landpage.py",  title="My Files",          icon="🗂️", url_path="landpage",  default=True),
        st.Page("pages/viewdata.py",  title="View Data",         icon="📁", url_path="viewdata"),
        st.Page("pages/rangesheetreview.py",  title="Rangesheet Review", icon="🔍", url_path="rangesheetreview"),
        # st.Page("pages/dashboard.py", title="Dashboard",         icon="📊", url_path="dashboard"),
        st.Page("pages/report.py",    title="Report",            icon="📋", url_path="report"),
        st.Page("pages/audit.py",     title="Audit Log",         icon="🛡️", url_path="audit"),
    ],
    position="hidden",   # hide default sidebar nav — we render our own
)
# Button callbacks run before the next script pass. Resolve queued navigation
# after the page registry exists, but before rendering the current heavy page.
_page_nav_target = st.session_state.pop("_page_nav_target", None)
if _page_nav_target:
    st.switch_page(_page_nav_target)

pg.run()

if __name__ == "__main__":
    try:
        from streamlit.runtime import exists as _st_exists
        if not _st_exists():
            import subprocess, sys
            subprocess.run([sys.executable, "-m", "streamlit", "run", __file__])
    except Exception:
        import subprocess, sys
        subprocess.run([sys.executable, "-m", "streamlit", "run", __file__])
