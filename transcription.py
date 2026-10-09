"""Groq speech-to-text calls and transcription logic."""
from __future__ import annotations

import logging
import os
import re
import time
from dataclasses import dataclass
from typing import Callable, Optional, Sequence, Tuple

import groq
from groq import Groq

import config

logger = logging.getLogger(__name__)

_MB = 1024 * 1024


@dataclass
class TranscriptionResult:
    """Outcome for exactly one uploaded file."""

    filename: str          # original filename, with extension
    audio_version: str     # filename without its final extension
    text: str = ""
    success: bool = False
    error: str = ""
    # "", "validation", "auth", "rate_limit", "network", "api", "empty", "skipped"
    error_kind: str = ""


# --------------------------------------------------------------------------- #
# Filename / file validation
# --------------------------------------------------------------------------- #
def strip_extension(filename: str) -> str:
    """Remove only the final extension: 'meeting.final.mp3' -> 'meeting.final'."""
    base = os.path.basename(filename.replace("\\", "/"))
    stem, _ext = os.path.splitext(base)
    return stem or base


# Trailing audio-version suffix such as "- A1", "-A2", " – a 3" (any spacing/case, hyphen or dash).
_VERSION_SUFFIX_RE = re.compile(r"\s*[-\u2013\u2014]\s*A\s*(\d+)\s*$", re.IGNORECASE)


def parse_slide_name(filename: str) -> Tuple[str, Optional[int]]:
    """Split '2.12 Customer Arrival 5 - A2.mp3' into ('2.12 Customer Arrival 5', 2).

    Removes the final extension and a trailing A<number> suffix only; every other part of the
    name (including meaningful numbers) is kept. Returns (name, None) when there is no suffix.
    """
    stem = strip_extension(filename).strip()
    match = _VERSION_SUFFIX_RE.search(stem)
    if not match:
        return stem, None
    return stem[: match.start()].rstrip(), int(match.group(1))


def get_extension(filename: str) -> str:
    return os.path.splitext(filename)[1].lstrip(".").lower()


def validate_file(filename: str, size_bytes: int, max_size_mb: Optional[float] = None) -> Optional[str]:
    """Return an error message if the file cannot be sent to Groq, else None."""
    max_mb = max_size_mb if max_size_mb is not None else config.get_max_file_size_mb()
    ext = get_extension(filename)
    if ext not in config.SUPPORTED_EXTENSIONS:
        shown = f".{ext}" if ext else "(no extension)"
        return (
            f"Unsupported file type {shown}. Supported types: "
            + ", ".join(f".{e}" for e in config.SUPPORTED_EXTENSIONS)
            + "."
        )
    if size_bytes <= 0:
        return "The file is empty (0 bytes)."
    if size_bytes > max_mb * _MB:
        return (
            f"The file is {size_bytes / _MB:.1f} MB, which is over the {max_mb:g} MB limit. "
            "Compress or split the recording, or raise MAX_FILE_SIZE_MB if your Groq plan allows larger files."
        )
    return None


# --------------------------------------------------------------------------- #
# Groq client + error classification
# --------------------------------------------------------------------------- #
def create_client(api_key: str) -> Groq:
    """Build a Groq client for an explicitly supplied user's key."""
    key = api_key.strip() if isinstance(api_key, str) else ""
    if not key:
        raise ValueError("Enter your Groq API key first.")
    return Groq(api_key=key, timeout=config.REQUEST_TIMEOUT_SECONDS, max_retries=0)


def validate_api_key(api_key: str) -> Tuple[bool, str]:
    """Check credentials with the lightweight model-list endpoint, not a transcription."""
    if not isinstance(api_key, str) or not api_key.strip():
        return False, "Enter your Groq API key first."
    try:
        create_client(api_key).models.list()
    except Exception as exc:  # noqa: BLE001 - never surface raw SDK exception text
        _kind, message, _retryable = _classify(exc)
        return False, message
    return True, "The API key is valid. Select Save and Continue to use it for this session."


def _retry_after_seconds(exc: Exception) -> Optional[float]:
    response = getattr(exc, "response", None)
    headers = getattr(response, "headers", None)
    if not headers:
        return None
    try:
        return float(headers.get("retry-after"))
    except (TypeError, ValueError):
        return None


def _classify(exc: Exception) -> Tuple[str, str, bool]:
    """Return (error_kind, user_message, retryable)."""
    if isinstance(exc, groq.AuthenticationError):
        return ("auth", "Groq rejected this API key. Check that it is active and copied correctly.", False)
    if isinstance(exc, groq.PermissionDeniedError):
        return ("auth", "Access denied by Groq. Check your API key permissions and model access.", False)
    if isinstance(exc, groq.RateLimitError):
        return ("rate_limit", "Groq rate or usage limit reached. Check your limits and try again later.", True)
    if isinstance(exc, groq.APITimeoutError):
        return ("network", "The request to Groq timed out.", True)
    if isinstance(exc, groq.APIConnectionError):
        return ("network", "Could not connect to Groq. Check your internet connection.", True)
    if isinstance(exc, groq.APIStatusError):
        status = exc.status_code
        if status == 402:
            return ("api", "Groq account quota or billing limit reached. Check your usage and billing settings.", False)
        if status == 413:
            return ("validation", "Groq rejected the file as too large.", False)
        if status in (400, 404, 422):
            return ("validation", f"Groq could not process this request (HTTP {status}). Check the file format and request settings.", False)
        if status >= 500:
            return ("api", f"Groq server error ({status}).", True)
        return ("api", f"Groq API request failed (HTTP {status}).", False)
    return ("api", "The Groq request failed unexpectedly. Please try again.", False)


def _extract_text(response) -> str:
    if isinstance(response, str):
        return response
    if isinstance(response, dict):
        return str(response.get("text") or "")
    return str(getattr(response, "text", "") or "")


# --------------------------------------------------------------------------- #
# Transcription
# --------------------------------------------------------------------------- #
def transcribe_file(
    client,
    filename: str,
    data: bytes,
    language: Optional[str] = None,
    model: Optional[str] = None,
    max_retries: Optional[int] = None,
    max_size_mb: Optional[float] = None,
    sleep: Callable[[float], None] = time.sleep,
) -> TranscriptionResult:
    """Transcribe one file. Never raises for per-file problems; returns a result instead."""
    result = TranscriptionResult(filename=filename, audio_version=strip_extension(filename))

    problem = validate_file(filename, len(data), max_size_mb)
    if problem:
        result.error, result.error_kind = problem, "validation"
        return result

    model = model or config.get_model()
    retries = config.MAX_RETRIES if max_retries is None else max_retries

    for attempt in range(retries + 1):
        try:
            kwargs = dict(
                file=(filename, data),
                model=model,
                response_format="json",
                temperature=0.0,
            )
            if language:
                kwargs["language"] = language  # omitted => Whisper auto-detects
            response = client.audio.transcriptions.create(**kwargs)
        except Exception as exc:  # noqa: BLE001 - classified below
            kind, message, retryable = _classify(exc)
            wait = _retry_after_seconds(exc)
            if retryable and attempt < retries:
                if wait is not None and wait > config.MAX_RETRY_WAIT_SECONDS:
                    message = (
                        f"{message} Groq asked to wait {wait:.0f} s, which is too long to wait here. "
                        "You may have hit the hourly/daily audio limit; try again later."
                    )
                    result.error, result.error_kind = message, kind
                    return result
                delay = wait if wait is not None else config.RETRY_BASE_DELAY_SECONDS * (2 ** attempt)
                logger.warning("%s (%s): retry %d/%d in %.1fs", filename, kind, attempt + 1, retries, delay)
                sleep(delay)
                continue
            if retryable:
                message = f"{message} Gave up after {retries} retries."
            result.error, result.error_kind = message, kind
            return result

        text = _extract_text(response).strip()
        if not text:
            result.error = "Groq returned an empty transcription (no speech detected?)."
            result.error_kind = "empty"
            return result
        result.text, result.success = text, True
        return result

    return result  # pragma: no cover - loop always returns


ProgressCallback = Callable[[int, int, str, Optional[TranscriptionResult]], None]


def transcribe_batch(
    client,
    files: Sequence[Tuple[str, bytes]],
    language: Optional[str] = None,
    model: Optional[str] = None,
    on_progress: Optional[ProgressCallback] = None,
    **kwargs,
) -> list[TranscriptionResult]:
    """Transcribe files one by one; exactly one result per input file, in order.

    on_progress(index, total, filename, result) is called before each file with
    result=None, and after each file with its result.
    """
    total = len(files)
    results: list[TranscriptionResult] = []
    fatal_error: Optional[str] = None

    for index, (name, data) in enumerate(files):
        if fatal_error:
            results.append(
                TranscriptionResult(
                    filename=name,
                    audio_version=strip_extension(name),
                    error=f"Skipped because of an earlier problem: {fatal_error}",
                    error_kind="skipped",
                )
            )
            if on_progress:
                on_progress(index, total, name, results[-1])
            continue

        if on_progress:
            on_progress(index, total, name, None)
        result = transcribe_file(client, name, data, language=language, model=model, **kwargs)
        results.append(result)
        if result.error_kind == "auth":
            fatal_error = result.error  # the key/permissions will fail for every file
        if on_progress:
            on_progress(index, total, name, result)

    return results
