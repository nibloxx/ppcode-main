"""Load API keys from environment variables (optional .env file)."""
import os
from pathlib import Path

_ENV_LOADED = False


def _load_dotenv() -> None:
    global _ENV_LOADED
    if _ENV_LOADED:
        return
    try:
        from dotenv import load_dotenv
        load_dotenv(Path(__file__).resolve().parent / ".env")
    except ImportError:
        pass
    _ENV_LOADED = True


def get_openai_api_key() -> str:
    _load_dotenv()
    key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not key:
        raise ValueError("OPENAI_API_KEY environment variable is not set")
    return key


def get_google_api_key() -> str:
    _load_dotenv()
    key = os.environ.get("GOOGLE_API_KEY", "").strip()
    if not key:
        raise ValueError("GOOGLE_API_KEY environment variable is not set")
    return key
