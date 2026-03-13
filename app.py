"""
ACFR Sensitivity Extraction — Streamlit App
Wraps sensitivity_extractor.py pipeline with a browser-based UI.
Place this file in the same directory as sensitivity_extractor.py.
Run with: streamlit run app.py
"""

import io
import logging
import os
import sys
import tempfile
import time
from pathlib import Path

import streamlit as st

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="ACFR Sensitivity Extractor",
    page_icon="📊",
    layout="wide",
)

# ── Minimal custom CSS ────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #0D1B2A; color: #E8E8E8; }
    section[data-testid="stSidebar"] { background-color: #0a1520; }
    h1, h2, h3 { color: #4A90C4; }
    .log-box {
        background: #0a1520;
        border: 1px solid #1e3a5f;
        border-radius: 6px;
        padding: 12px 16px;
        font-family: 'Courier New', monospace;
        font-size: 12px;
        height: 300px;
        overflow-y: auto;
        color: #a8c8e8;
    }
    .metric-card {
        background: #0a1520;
        border: 1px solid #1e3a5f;
        border-radius: 8px;
        padding: 16px;
        text-align: center;
    }
    .stButton>button {
        background-color: #4A90C4;
        color: white;
        border: none;
        border-radius: 6px;
        font-weight: 600;
    }
    .stButton>button:hover { background-color: #2AA89A; }
    .warning-text { color: #E8A838; }
    .error-text { color: #e05a5a; }
    .success-text { color: #2AA89A; }
</style>
""", unsafe_allow_html=True)


# ── Logging handler that writes to a Streamlit-readable list ─────────────────
class StreamlitLogHandler(logging.Handler):
    def __init__(self, log_store: list):
        super().__init__()
        self.log_store = log_store

    def emit(self, record):
        msg = self.format(record)
        level = record.levelname
        if level == "WARNING":
            self.log_store.append(f"⚠️  {msg}")
        elif level in ("ERROR", "CRITICAL"):
            self.log_store.append(f"❌ {msg}")
        else:
            self.log_store.append(f"   {msg}")


# ── Import pipeline (must be in same directory) ───────────────────────────────
@st.cache_resource(show_spinner=False)
def import_pipeline():
    try:
        import sensitivity_extractor as se
        return se, None
    except ImportError as e:
        return None, str(e)


# ── Main UI ───────────────────────────────────────────────────────────────────
st.title("📊 ACFR Sensitivity Extractor")
st.caption("Automated discount-rate sensitivity extraction from public pension ACFRs")

se_module, import_error = import_pipeline()
if import_error:
    st.error(f"Could not import `sensitivity_extractor.py` — make sure it's in the same directory as this app.\n\n`{import_error}`")
    st.stop()

# ── Sidebar: configuration ────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuration")

    api_key = st.text_input(
        "Gemini API Key",
        type="password",
        placeholder="AIza...",
        help="Your Google Gemini API key",
    )

    st.divider()
    st.subheader("📁 Upload Files")

    uploaded_pdfs = st.file_uploader(
        "ACFR PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        help="Upload one or more ACFR PDF files",
    )

    uploaded_plan_list = st.file_uploader(
        "Master Plan List (optional)",
        type=["csv", "xlsx", "xls"],
        help="CSV or Excel with YR / State / Plan Name columns",
    )

    st.divider()
    resume_mode = st.checkbox(
        "Resume from cache",
        value=False,
        help="Skip PDFs already processed in a previous run",
    )

    run_btn = st.button(
        "▶ Run Extraction",
        disabled=not (api_key and uploaded_pdfs),
        use_container_width=True,
    )

# ── Main area: status + results ───────────────────────────────────────────────
if not uploaded_pdfs:
    st.info("Upload ACFR PDFs in the sidebar to get started.")
    st.stop()

st.markdown(f"**{len(uploaded_pdfs)} PDF(s) ready** — {', '.join(f.name for f in uploaded_pdfs[:5])}{'…' if len(uploaded_pdfs) > 5 else ''}")

if not run_btn:
    st.stop()

if not api_key:
    st.error("Please enter your Gemini API key in the sidebar.")
    st.stop()

# ── Run pipeline ──────────────────────────────────────────────────────────────
log_messages = []

# Attach Streamlit log handler to the pipeline's logger
pipeline_logger = logging.getLogger("sensitivity")
pipeline_logger.setLevel(logging.INFO)
# Remove existing handlers to avoid duplicate output
pipeline_logger.handlers = []
sl_handler = StreamlitLogHandler(log_messages)
sl_handler.setFormatter(logging.Formatter("%(asctime)s  %(message)s", datefmt="%H:%M:%S"))
pipeline_logger.addHandler(sl_handler)

with tempfile.TemporaryDirectory() as tmpdir:
    # Write uploaded PDFs to temp folder
    pdf_folder = os.path.join(tmpdir, "pdfs")
    os.makedirs(pdf_folder)
    for f in uploaded_pdfs:
        dest = os.path.join(pdf_folder, f.name)
        with open(dest, "wb") as out:
            out.write(f.read())

    # Write plan list if provided
    plan_list_path = None
    if uploaded_plan_list:
        plan_list_path = os.path.join(tmpdir, uploaded_plan_list.name)
        with open(plan_list_path, "wb") as out:
            out.write(uploaded_plan_list.read())

    output_xlsx = os.path.join(tmpdir, "sensitivity_results.xlsx")

    # Progress display
    status_area = st.empty()
    log_area = st.empty()
    progress_bar = st.progress(0, text="Starting pipeline…")

    # We patch time.sleep to update the progress bar as each PDF processes
    processed = [0]
    total_pdfs = len(uploaded_pdfs)
    original_sleep = time.sleep

    def patched_sleep(secs):
        processed[0] += 1
        pct = min(processed[0] / total_pdfs, 1.0)
        progress_bar.progress(pct, text=f"Processing PDF {processed[0]} of {total_pdfs}…")
        # Refresh log display
        log_area.markdown(
            "<div class='log-box'>" +
            "<br>".join(log_messages[-40:]) +
            "</div>",
            unsafe_allow_html=True,
        )
        original_sleep(secs)

    time.sleep = patched_sleep

    try:
        se_module.run_pipeline(
            pdf_folder=pdf_folder,
            output_xlsx=output_xlsx,
            api_key=api_key,
            resume=resume_mode,
            plan_list_path=plan_list_path,
        )
        progress_bar.progress(1.0, text="✅ Extraction complete")
        status_area.success("Pipeline finished successfully.")
    except Exception as e:
        status_area.error(f"Pipeline error: {e}")
        log_messages.append(f"❌ Fatal: {e}")
    finally:
        time.sleep = original_sleep

    # Final log flush
    log_area.markdown(
        "<div class='log-box'>" +
        "<br>".join(log_messages) +
        "</div>",
        unsafe_allow_html=True,
    )

    # ── Results ───────────────────────────────────────────────────────────────
    if os.path.exists(output_xlsx):
        import openpyxl
        import pandas as pd

        df = pd.read_excel(output_xlsx)

        st.divider()
        st.subheader("📋 Results")

        # Summary metrics
        total = len(df)
        success = df["Plan Name (Extracted)"].notna().sum()
        errors = df["Validation Warnings"].notna().sum()

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("PDFs Processed", total_pdfs)
        col2.metric("Plan Rows Extracted", total)
        col3.metric("With Matched Plan Name", int(success))
        col4.metric("With Warnings", int(errors))

        # Dataframe with warning highlighting
        st.dataframe(df, use_container_width=True, height=400)

        # Download button
        with open(output_xlsx, "rb") as f:
            excel_bytes = f.read()

        st.download_button(
            label="⬇️ Download Excel Results",
            data=excel_bytes,
            file_name="sensitivity_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.warning("No output file was produced — check the log above for errors.")
