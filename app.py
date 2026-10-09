"""Streamlit user interface for the Audio-to-Text Transcription System."""
from __future__ import annotations

from datetime import datetime
import time

import pandas as pd
import streamlit as st

import config
import excel_generator
import transcription

st.set_page_config(page_title="Audio-to-Text Transcription System", page_icon="🎙️", layout="wide")

# Common Whisper languages (ISO-639-1). "Auto-detect" sends no language parameter.
LANGUAGES = {
    "Auto-detect (recommended)": None,
    "English": "en", "Hindi": "hi", "Tamil": "ta", "Telugu": "te", "Malayalam": "ml",
    "Kannada": "kn", "Bengali": "bn", "Marathi": "mr", "Urdu": "ur",
    "Spanish": "es", "French": "fr", "German": "de", "Italian": "it", "Portuguese": "pt",
    "Dutch": "nl", "Russian": "ru", "Arabic": "ar", "Turkish": "tr",
    "Chinese": "zh", "Japanese": "ja", "Korean": "ko", "Indonesian": "id",
}

# --------------------------------------------------------------------------- #
# Session state
# --------------------------------------------------------------------------- #
for key, default in {
    "processing": False,
    "results": None,
    "excel_bytes": None,
    "excel_name": None,
    "excel_error": None,
    "oversized": [],
    "merge_result": None,
    "merge_bytes": None,
    "merge_name": None,
    "merge_error": None,
    "groq_api_key": "",
    "api_key_input": "",
    "api_key_feedback": None,
    "api_key_saved_notice": False,
}.items():
    st.session_state.setdefault(key, default)
st.session_state.setdefault("api_key_config_open", not bool(st.session_state.groq_api_key))


def _clear_api_key_feedback() -> None:
    st.session_state.api_key_feedback = None


def _reset_api_key() -> None:
    st.session_state.groq_api_key = ""
    st.session_state.api_key_input = ""
    st.session_state.api_key_config_open = True
    st.session_state.api_key_feedback = ("info", "API key cleared. Enter your own key to transcribe audio.")


def _open_api_key_config() -> None:
    st.session_state.api_key_config_open = True
    st.session_state.api_key_feedback = None


def _close_api_key_config() -> None:
    st.session_state.api_key_config_open = False
    st.session_state.api_key_input = ""
    st.session_state.api_key_feedback = None


def _validate_and_save_api_key() -> None:
    candidate = st.session_state.api_key_input.strip()
    if not candidate:
        st.session_state.api_key_feedback = ("error", "Enter your Groq API key first.")
        return
    valid, message = transcription.validate_api_key(candidate)
    if not valid:
        st.session_state.api_key_feedback = ("error", message)
        return
    st.session_state.groq_api_key = candidate
    st.session_state.api_key_input = ""
    st.session_state.api_key_config_open = False
    st.session_state.api_key_feedback = None
    st.session_state.api_key_saved_notice = True


def _start_processing() -> None:
    """Button callback: runs before the rerun, so the button is disabled during processing."""
    st.session_state.processing = True
    st.session_state.results = None
    st.session_state.excel_bytes = None
    st.session_state.excel_error = None
    st.session_state.oversized = []


# --------------------------------------------------------------------------- #
# Header + configuration check
# --------------------------------------------------------------------------- #
st.title("Audio-to-Text Transcription System")
st.write(
    "Upload one or more audio recordings. Each file is transcribed with Groq's speech-to-text "
    "service, and you can download the results as an Excel file with two columns: "
    "**Audio version** (file name without extension) and **Audio to text** (the transcription). "
    "You can also merge existing transcription workbooks without running transcription."
)

audio_config_errors = []
model = config.get_model()
try:
    max_mb = config.get_max_file_size_mb()
except config.ConfigError as exc:
    audio_config_errors.append(str(exc))
    max_mb = float(config.DEFAULT_MAX_FILE_SIZE_MB)

with st.sidebar:
    st.subheader("Settings in use")
    st.write(f"**Model:** `{model}`")
    st.write(f"**Max file size:** {max_mb:g} MB")
    st.write(f"**API key:** {'configured for this session' if st.session_state.groq_api_key else 'not configured'}")
    st.caption("Supported formats: " + ", ".join(config.SUPPORTED_EXTENSIONS))

for error in audio_config_errors:
    st.error(error)

if st.session_state.api_key_saved_notice:
    st.success("API key validated and saved successfully!")
    time.sleep(1.0)
    st.session_state.api_key_saved_notice = False

if st.session_state.api_key_config_open:
    st.subheader("Groq API Key Configuration")
    st.markdown("Create a key in the [Groq console](https://console.groq.com/keys).")
    st.text_input(
        "Enter your Groq API Key",
        type="password",
        key="api_key_input",
        on_change=_clear_api_key_feedback,
        disabled=st.session_state.processing,
        help="Your key is used only by this Streamlit session and is never shown in full.",
    )
    actions = st.columns(3)
    actions[0].button(
        "Validate and Save",
        type="primary",
        on_click=_validate_and_save_api_key,
        disabled=st.session_state.processing,
    )
    actions[1].button("Reset API Key", on_click=_reset_api_key, disabled=st.session_state.processing)
    if st.session_state.groq_api_key:
        actions[2].button("Cancel", on_click=_close_api_key_config, disabled=st.session_state.processing)
    if st.session_state.api_key_feedback:
        feedback_kind, feedback_message = st.session_state.api_key_feedback
        getattr(st, feedback_kind)(feedback_message)
    st.stop()

st.button("Change API Key", on_click=_open_api_key_config, disabled=st.session_state.processing)

# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
uploaded_files = st.file_uploader(
    "Upload audio files",
    type=list(config.SUPPORTED_EXTENSIONS),
    accept_multiple_files=True,
    disabled=st.session_state.processing or bool(audio_config_errors) or not st.session_state.groq_api_key,
    help=f"Up to {max_mb:g} MB per file.",
)

language_label = st.selectbox(
    "Audio language (optional)",
    list(LANGUAGES),
    disabled=st.session_state.processing,
    help="Leave on Auto-detect unless you know the language. The transcript is never translated.",
)
language_code = LANGUAGES[language_label]

st.button(
    "Transcribe and Generate Excel",
    type="primary",
    on_click=_start_processing,
    disabled=st.session_state.processing or not uploaded_files or bool(audio_config_errors) or not st.session_state.groq_api_key,
)

st.divider()
st.subheader("Merge Existing Excel Files")
excel_uploads = st.file_uploader(
    "Upload Excel workbooks",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    key="excel_merge_uploads",
)
if st.button("Merge and Standardize", disabled=not excel_uploads, key="merge_excel_button"):
    st.session_state.merge_result = None
    st.session_state.merge_bytes = None
    st.session_state.merge_error = None
    try:
        merge_result = excel_generator.merge_excel_files(
            [(uploaded.name, uploaded.getvalue()) for uploaded in excel_uploads]
        )
        st.session_state.merge_result = merge_result
        if merge_result.dataframe.empty:
            st.session_state.merge_error = "No valid audio records were found in the uploaded workbooks."
        else:
            st.session_state.merge_bytes = excel_generator.generate_merged_excel_bytes(
                merge_result.dataframe
            )
            st.session_state.merge_name = f"merged_transcriptions_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    except excel_generator.ExcelGenerationError as exc:
        st.session_state.merge_error = str(exc)
    except Exception as exc:  # noqa: BLE001 - keep upload failures readable in the UI
        st.session_state.merge_error = f"Could not merge the uploaded workbooks: {exc}"

merge_result = st.session_state.merge_result
if merge_result is not None:
    st.subheader("Merge Summary")
    summary_cols = st.columns(4)
    summary_cols[0].metric("Files received", merge_result.files_received)
    summary_cols[1].metric("Workbooks read", merge_result.files_processed)
    summary_cols[2].metric("Worksheets read", merge_result.worksheets_processed)
    summary_cols[3].metric("Records merged", len(merge_result.dataframe))

    if merge_result.issues:
        st.warning("Some worksheets or rows were skipped:")
        st.dataframe(pd.DataFrame({"Issue": merge_result.issues}), hide_index=True)
    if not merge_result.duplicates.empty:
        st.warning("Possible duplicate slide/audio-version records are included below for review.")
        st.dataframe(merge_result.duplicates, hide_index=True)
    if st.session_state.merge_error:
        st.error(st.session_state.merge_error)
    if not merge_result.dataframe.empty:
        st.subheader("Merged Preview")
        st.dataframe(
            excel_generator.merged_preview_dataframe(merge_result.dataframe), hide_index=True
        )
    if st.session_state.merge_bytes:
        st.download_button(
            "Download Standardized Excel File",
            data=st.session_state.merge_bytes,
            file_name=st.session_state.merge_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_merged_excel",
        )

# --------------------------------------------------------------------------- #
# Processing
# --------------------------------------------------------------------------- #
def _run(files: list[tuple[str, bytes]]) -> None:
    api_key = st.session_state.groq_api_key
    if not api_key:
        raise RuntimeError("Enter and save your Groq API key before transcribing audio.")
    client = transcription.create_client(api_key)
    progress = st.progress(0.0, text="Starting…")
    status = st.empty()
    total = len(files)

    def on_progress(index: int, total_: int, name: str, result) -> None:
        if result is None:
            status.info(f"Transcribing file {index + 1} of {total_}: **{name}**")
        else:
            progress.progress((index + 1) / total_, text=f"{index + 1} of {total_} files processed")

    results = transcription.transcribe_batch(
        client, files, language=language_code, model=model, on_progress=on_progress
    )
    status.empty()
    progress.empty()
    st.session_state.results = results

    df = excel_generator.build_dataframe(results)
    st.session_state.oversized = excel_generator.oversized_files(results)
    if df.empty:
        st.session_state.excel_error = "No file could be transcribed, so there is nothing to export."
        return
    try:
        excel_bytes = excel_generator.generate_excel_bytes(df)
        excel_generator.save_excel_file(excel_bytes)  # keep a copy in the output/ folder
    except excel_generator.ExcelGenerationError as exc:
        st.session_state.excel_error = str(exc)
        return
    st.session_state.excel_bytes = excel_bytes
    st.session_state.excel_name = f"transcriptions_{datetime.now():%Y%m%d_%H%M%S}.xlsx"


if st.session_state.processing:
    try:
        _run([(f.name, f.getvalue()) for f in uploaded_files or []])
    except Exception:  # noqa: BLE001 - do not expose SDK exceptions or credentials
        st.session_state.excel_error = "Processing stopped unexpectedly. Check your key, connection, and file settings, then try again."
    finally:
        st.session_state.processing = False
    st.rerun()  # refresh so the button is enabled again and results are shown

# --------------------------------------------------------------------------- #
# Results
# --------------------------------------------------------------------------- #
results = st.session_state.results
if results is not None:
    ok = [r for r in results if r.success]
    failed = [r for r in results if not r.success]

    st.divider()
    st.subheader("Summary")
    col1, col2 = st.columns(2)
    col1.metric("Transcribed successfully", len(ok))
    col2.metric("Failed", len(failed))

    if failed:
        st.error("Some files could not be transcribed. They are NOT included in the Excel file.")
        st.dataframe(
            pd.DataFrame({"File": [r.filename for r in failed], "Reason": [r.error for r in failed]}),
            hide_index=True,
        )

    if st.session_state.oversized:
        st.warning(
            "These transcripts are longer than Excel allows in one cell, so they continue on following "
            "rows with the same name: " + ", ".join(st.session_state.oversized)
        )

    if st.session_state.excel_error:
        st.error(st.session_state.excel_error)

    if ok:
        st.subheader("Preview")
        st.dataframe(excel_generator.build_dataframe(ok), hide_index=True)

    if st.session_state.excel_bytes:
        st.download_button(
            "Download Excel File",
            data=st.session_state.excel_bytes,
            file_name=st.session_state.excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
