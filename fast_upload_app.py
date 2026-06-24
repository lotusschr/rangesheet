"""
Fast Streamlit app for handling large uploaded files.

Optimizations included:
1. @st.cache_data so the file is only parsed once (not on every rerun)
2. Reads file as bytes first, so caching works reliably
3. Lets you choose pandas (engine="pyarrow") or Polars for faster parsing
4. Only loads needed columns where possible
5. Stores processed result in st.session_state to avoid recompute on reruns
6. Shows a preview (head) instead of rendering the full huge dataframe
7. Has basic timing so you can see where time is actually being spent

Run with:
    streamlit run fast_upload_app.py
"""

import time
import io

import streamlit as st
import pandas as pd

# Polars is optional - app still works without it, just skips that option
try:
    import polars as pl
    POLARS_AVAILABLE = True
except ImportError:
    POLARS_AVAILABLE = False


st.set_page_config(page_title="Fast Large File Upload", layout="wide")
st.title("📂 Fast Large File Upload & Processing")

st.caption(
    "Upload a large CSV/Parquet file. The file is parsed once and cached — "
    "clicking other widgets afterward won't re-trigger the slow parsing step."
)

# ---------------------------------------------------------------------------
# Sidebar: options
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Settings")

    engine = st.radio(
        "Parsing engine",
        options=["pandas (pyarrow)", "pandas (default)"] + (["polars"] if POLARS_AVAILABLE else []),
        index=0,
        help="pyarrow / polars are generally much faster for large CSVs.",
    )

    preview_rows = st.slider("Preview rows to display", 10, 2000, 200, step=10)

    usecols_raw = st.text_input(
        "Only load these columns (comma-separated, optional)",
        value="",
        help="Leave blank to load all columns. Loading fewer columns is faster.",
    )

    if st.button("🗑️ Clear cache & reset"):
        st.cache_data.clear()
        st.session_state.clear()
        st.rerun()


# ---------------------------------------------------------------------------
# Cached loading functions
# IMPORTANT: these take raw bytes (not the UploadedFile object) as input,
# since bytes hash reliably for st.cache_data.
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def load_with_pandas(file_bytes: bytes, use_pyarrow: bool, usecols):
    kwargs = {}
    if usecols:
        kwargs["usecols"] = usecols
    if use_pyarrow:
        kwargs["engine"] = "pyarrow"
    return pd.read_csv(io.BytesIO(file_bytes), **kwargs)


@st.cache_data(show_spinner=False)
def load_with_polars(file_bytes: bytes, usecols):
    if usecols:
        return pl.read_csv(io.BytesIO(file_bytes), columns=usecols)
    return pl.read_csv(io.BytesIO(file_bytes))


@st.cache_data(show_spinner=False)
def load_parquet_pandas(file_bytes: bytes, usecols):
    kwargs = {}
    if usecols:
        kwargs["columns"] = usecols
    return pd.read_parquet(io.BytesIO(file_bytes), **kwargs)


# ---------------------------------------------------------------------------
# File upload
# ---------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Upload a CSV or Parquet file",
    type=["csv", "parquet"],
)

if uploaded_file is not None:
    # Build a cache key so we know if this is a *new* file vs. the one
    # already processed and stored in session_state.
    file_id = f"{uploaded_file.name}-{uploaded_file.size}"

    usecols = [c.strip() for c in usecols_raw.split(",") if c.strip()] or None

    needs_processing = (
        "file_id" not in st.session_state
        or st.session_state["file_id"] != file_id
        or st.session_state.get("engine") != engine
        or st.session_state.get("usecols") != usecols
    )

    if needs_processing:
        with st.spinner("Reading and parsing file..."):
            t0 = time.time()

            # Read bytes ONCE
            file_bytes = uploaded_file.getvalue()
            is_parquet = uploaded_file.name.lower().endswith(".parquet")

            if is_parquet:
                df = load_parquet_pandas(file_bytes, usecols)
            elif engine == "pandas (pyarrow)":
                df = load_with_pandas(file_bytes, use_pyarrow=True, usecols=usecols)
            elif engine == "pandas (default)":
                df = load_with_pandas(file_bytes, use_pyarrow=False, usecols=usecols)
            elif engine == "polars":
                pl_df = load_with_polars(file_bytes, usecols)
                df = pl_df.to_pandas()  # convert once for downstream display/use
            else:
                df = load_with_pandas(file_bytes, use_pyarrow=False, usecols=usecols)

            load_time = time.time() - t0

        # Store everything needed in session_state so future reruns
        # (e.g. moving the preview_rows slider) skip re-parsing entirely.
        st.session_state["file_id"] = file_id
        st.session_state["engine"] = engine
        st.session_state["usecols"] = usecols
        st.session_state["df"] = df
        st.session_state["load_time"] = load_time
    else:
        df = st.session_state["df"]
        load_time = st.session_state["load_time"]

    # -----------------------------------------------------------------
    # Display
    # -----------------------------------------------------------------
    col1, col2, col3 = st.columns(3)
    col1.metric("Rows", f"{len(df):,}")
    col2.metric("Columns", f"{df.shape[1]:,}")
    col3.metric("Load time", f"{load_time:.2f}s")

    st.success(f"File loaded and cached as `{file_id}` using **{engine}**.")

    st.subheader("Preview")
    st.dataframe(df.head(preview_rows), use_container_width=True)

    with st.expander("Show column dtypes"):
        st.write(df.dtypes.astype(str))

    # -----------------------------------------------------------------
    # Example: a widget interaction that does NOT trigger re-parsing
    # -----------------------------------------------------------------
    st.subheader("Quick filter (no reprocessing of the original file)")
    if len(df.columns) > 0:
        filter_col = st.selectbox("Filter by column", options=df.columns)
        if pd.api.types.is_numeric_dtype(df[filter_col]):
            min_val, max_val = float(df[filter_col].min()), float(df[filter_col].max())
            if min_val < max_val:
                low, high = st.slider(
                    "Range", min_val, max_val, (min_val, max_val)
                )
                filtered = df[(df[filter_col] >= low) & (df[filter_col] <= high)]
            else:
                filtered = df
        else:
            options = df[filter_col].dropna().unique().tolist()
            chosen = st.multiselect("Values", options=options, default=options[:5])
            filtered = df[df[filter_col].isin(chosen)] if chosen else df

        st.write(f"Showing {len(filtered):,} of {len(df):,} rows")
        st.dataframe(filtered.head(preview_rows), use_container_width=True)

else:
    st.info("👆 Upload a file to get started.")
    # Clear stale state if file is removed
    for key in ["file_id", "df", "load_time", "engine", "usecols"]:
        st.session_state.pop(key, None)
