"""Load API keys from environment variables (optional .env file)."""
import os
from pathlib import Path
from typing import Optional

_ENV_LOADED = False


def _load_dotenv() -> None:
    global _ENV_LOADED
    if _ENV_LOADED:
        return
    try:
        from dotenv import load_dotenv

        project_root = Path(__file__).resolve().parent
        for env_path in (
            project_root / ".env",
            project_root / "vascreDev" / ".env",
        ):
            if env_path.exists():
                load_dotenv(env_path, override=False)
    except ImportError:
        pass
    _ENV_LOADED = True


def get_openai_api_key() -> str:
    _load_dotenv()
    key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not key:
        raise ValueError("OPENAI_API_KEY environment variable is not set")
    return key


def _clean_key(value: str) -> Optional[str]:
    key = value.strip()
    if not key or key.startswith("your-"):
        return None
    return key


def get_google_api_key() -> Optional[str]:
    _load_dotenv()
    key = _clean_key(os.environ.get("GOOGLE_API_KEY", ""))
    if key:
        return key
    return _clean_key(os.environ.get("NEXT_PUBLIC_APIKEY", ""))


def get_esri_api_key() -> Optional[str]:
    _load_dotenv()
    return _clean_key(os.environ.get("ESRI_API_KEY", ""))
