"""
Streamlit front-end for image_extract.py

Run with:
    streamlit run streamlit_app.py

Requires image_extract.py to sit in the SAME folder as this file.
"""

import io
import shutil
import sys
import tempfile
import traceback
import zipfile
from pathlib import Path

import streamlit as st

# ------------------------------------------------------------------
# Make sure image_extract.py (next to this file) can be imported
# ------------------------------------------------------------------
sys.path.insert(0, str(Path(__file__).parent))
import image_extract as extractor  # noqa: E402  (the uploaded script, unmodified)

st.set_page_config(page_title="Control Narrative Extractor", layout="wide")
st.title("Control Narrative Extractor")
st.caption(
    "Upload a control narrative .docx, click Run, then view/download every "
    "generated image, Excel workbook, and JSON file."
)

# ------------------------------------------------------------------
# Inputs
# ------------------------------------------------------------------
uploaded_file = st.file_uploader("Upload .docx", type=["docx"])

with st.expander("Options", expanded=False):
    target_section = st.text_input(
        "Target section (contains the block diagram)",
        value=getattr(extractor, "TARGET_SECTION", "4.1.1"),
    )
    col1, col2 = st.columns(2)
    with col1:
        extract_diagram_image = st.checkbox(
            "Extract diagram image",
            value=getattr(extractor, "EXTRACT_DIAGRAM_IMAGE", False),
            help="EMF->PNG conversion needs Windows + PowerPoint installed "
                 "on the machine running this app. If unavailable, the "
                 "raw EMF is kept instead and a warning is logged.",
        )
    with col2:
        extract_reference_images = st.checkbox(
            "Extract per-pin reference images",
            value=getattr(extractor, "EXTRACT_REFERENCE_IMAGES", True),
        )

run_clicked = st.button("Run Extraction", type="primary", disabled=uploaded_file is None)

if "last_output_dir" not in st.session_state:
    st.session_state.last_output_dir = None
if "last_work_dir" not in st.session_state:
    st.session_state.last_work_dir = None

# ------------------------------------------------------------------
# Run the extraction
# ------------------------------------------------------------------
if run_clicked and uploaded_file is not None:
    # Clean up the previous run's temp folder before starting a new one
    if st.session_state.last_work_dir:
        shutil.rmtree(st.session_state.last_work_dir, ignore_errors=True)

    work_dir = Path(tempfile.mkdtemp(prefix="narrative_extract_"))
    input_docx_path = work_dir / uploaded_file.name
    input_docx_path.write_bytes(uploaded_file.getbuffer())

    output_dir = work_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    log_placeholder = st.empty()
    log_lines = []

    class _StreamlitLogWriter:
        """Redirects the script's print() calls into a live log box."""
        def write(self, msg):
            if msg.strip():
                log_lines.append(msg.rstrip("\n"))
                log_placeholder.code("\n".join(log_lines))

        def flush(self):
            pass

    old_stdout = sys.stdout
    sys.stdout = _StreamlitLogWriter()

    error_message = None
    try:
        with st.spinner("Running extraction — this can take a while for large documents..."):
            extractor.main(
                doc_path=str(input_docx_path),
                output_dir=str(output_dir),
                target_section=target_section,
                extract_diagram_image=extract_diagram_image,
                extract_reference_images=extract_reference_images,
            )
    except SystemExit as e:
        # main() raises SystemExit on expected failures (e.g. "section
        # not found", "document not found"). SystemExit does NOT inherit
        # from Exception, so it needs its own except clause here or it
        # would propagate uncaught and crash this Streamlit run.
        error_message = f"Stopped: {e}"
    except Exception:
        error_message = traceback.format_exc()
    finally:
        sys.stdout = old_stdout

    if error_message:
        st.error(error_message)
        st.session_state.last_output_dir = None
    else:
        st.success("Extraction complete.")
        st.session_state.last_output_dir = str(output_dir)

    st.session_state.last_work_dir = str(work_dir)

# ------------------------------------------------------------------
# Show + download results
# ------------------------------------------------------------------
if st.session_state.last_output_dir:
    output_dir = Path(st.session_state.last_output_dir)
    all_files = sorted(p for p in output_dir.rglob("*") if p.is_file())

    if not all_files:
        st.warning("No output files were generated for this document/section.")
    else:
        st.subheader("Results")

        # One-click "download everything" as a single zip
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in all_files:
                zf.write(f, f.relative_to(output_dir))
        st.download_button(
            "Download all outputs (.zip)",
            data=zip_buffer.getvalue(),
            file_name="extraction_outputs.zip",
            mime="application/zip",
        )

        # Group files by their containing folder: "overall" and each
        # per-mnemonic folder (see image_extract.py's OUTPUT_DIR layout)
        by_folder = {}
        for f in all_files:
            folder_name = f.relative_to(output_dir).parts[0]
            by_folder.setdefault(folder_name, []).append(f)

        for folder_name in sorted(by_folder):
            files = sorted(by_folder[folder_name])
            with st.expander(f"{folder_name}  ({len(files)} file{'s' if len(files) != 1 else ''})",
                              expanded=(folder_name == "overall")):
                for f in files:
                    suffix = f.suffix.lower()
                    left, right = st.columns([4, 1])

                    with left:
                        st.write(f"**{f.name}**")
                        try:
                            if suffix in (".png", ".jpg", ".jpeg"):
                                st.image(str(f), width=450)
                            elif suffix == ".txt":
                                st.text(f.read_text(encoding="utf-8", errors="replace")[:3000])
                            elif suffix == ".json":
                                st.json(f.read_text(encoding="utf-8"))
                            elif suffix == ".xlsx":
                                try:
                                    import pandas as pd
                                    xls = pd.ExcelFile(f)
                                    sheet = st.selectbox(
                                        f"Sheet ({f.name})", xls.sheet_names,
                                        key=f"sheet_{f}"
                                    )
                                    st.dataframe(xls.parse(sheet), height=250)
                                except Exception:
                                    st.caption("(preview unavailable — download to view)")
                            elif suffix in (".emf", ".wmf"):
                                st.caption("(binary vector image — no browser preview; download to view)")
                        except Exception as e:
                            st.caption(f"(preview failed: {e})")

                    with right:
                        with open(f, "rb") as fh:
                            st.download_button(
                                "Download", data=fh.read(), file_name=f.name,
                                key=f"dl_{f}"
                            )
