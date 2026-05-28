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
# Optional: go straight to the queue page instead of clicking through from the landing page.
CBC_QUEUE_URL = os.environ.get("CBC_QUEUE_URL", "")
# Saved browser session (cookies) from login.py — no password is ever stored.
STORAGE_STATE_PATH = Path(os.environ.get("STORAGE_STATE_PATH", "./cbc_session.json"))

ANTHROPIC_API_KEY = _required("ANTHROPIC_API_KEY")
CLAUDE_MODEL = "claude-opus-4-7"

OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", "./output"))
SNAPSHOT_DIR = Path(os.environ.get("SNAPSHOT_DIR", OUTPUT_DIR / "snapshots"))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)

# Email is sent through your local Outlook desktop app — no password needed.
# These are just the destination addresses (an address alone isn't sensitive).
EMAIL_TO = _required("EMAIL_TO")
EMAIL_ALERT_TO = os.environ.get("EMAIL_ALERT_TO", EMAIL_TO)
