"""Configuration and environment-variable validation."""
from __future__ import annotations

import os
from pathlib import Path

from dotenv import dotenv_values

BASE_DIR = Path(__file__).resolve().parent
_dotenv_values = dotenv_values(BASE_DIR / ".env")
_LOCAL_SETTINGS = {
    name: _dotenv_values.get(name)
    for name in ("GROQ_MODEL", "MAX_FILE_SIZE_MB")
}
del _dotenv_values

OUTPUT_DIR = BASE_DIR / "output"

# Verified against https://console.groq.com/docs/speech-to-text
DEFAULT_MODEL = "whisper-large-v3"
SUPPORTED_EXTENSIONS = ("flac", "mp3", "mp4", "mpeg", "mpga", "m4a", "ogg", "wav", "webm")
DEFAULT_MAX_FILE_SIZE_MB = 25  # free tier; the developer tier allows 100 MB

# Retry behaviour for temporary errors (rate limits, timeouts, 5xx)
MAX_RETRIES = 3
RETRY_BASE_DELAY_SECONDS = 2.0
MAX_RETRY_WAIT_SECONDS = 60  # if Groq asks us to wait longer, give up on that file
REQUEST_TIMEOUT_SECONDS = 300

class ConfigError(Exception):
    """Raised when required configuration is missing or invalid."""


def _get_setting(name: str) -> str:
    value = os.getenv(name)
    if value is None:
        value = _LOCAL_SETTINGS.get(name) or ""
    return value.strip()


def get_model() -> str:
    return _get_setting("GROQ_MODEL") or DEFAULT_MODEL


def get_max_file_size_mb() -> float:
    raw = _get_setting("MAX_FILE_SIZE_MB")
    if not raw:
        return float(DEFAULT_MAX_FILE_SIZE_MB)
    try:
        value = float(raw)
        if value <= 0:
            raise ValueError
        return value
    except ValueError:
        raise ConfigError(f"MAX_FILE_SIZE_MB must be a positive number, got '{raw}'.") from None
