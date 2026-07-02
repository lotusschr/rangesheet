"""RangeSheet — navigation entry point.
Run: streamlit run APP.py
"""
import streamlit as st

st.set_page_config(
    page_title="RangeSheet",
    page_icon="🗂️",
    layout="wide",
    initial_sidebar_state="expanded",
)

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
