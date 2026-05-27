import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()


def _required(key: str) -> str:
    val = os.environ.get(key)
    if not val:
        raise RuntimeError(f"Missing required environment variable: {key}")
    return val


CBC_URL = os.environ.get("CBC_URL", "https://cbcinsider.com")
CBC_USERNAME = _required("CBC_USERNAME")
CBC_PASSWORD = _required("CBC_PASSWORD")

ANTHROPIC_API_KEY = _required("ANTHROPIC_API_KEY")
CLAUDE_MODEL = "claude-opus-4-7"

OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", "./output"))
SNAPSHOT_DIR = Path(os.environ.get("SNAPSHOT_DIR", OUTPUT_DIR / "snapshots"))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)

SMTP_HOST = _required("SMTP_HOST")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = _required("SMTP_USER")
SMTP_PASSWORD = _required("SMTP_PASSWORD")
EMAIL_FROM = os.environ.get("EMAIL_FROM", SMTP_USER)
EMAIL_TO = _required("EMAIL_TO")
EMAIL_ALERT_TO = os.environ.get("EMAIL_ALERT_TO", EMAIL_TO)
