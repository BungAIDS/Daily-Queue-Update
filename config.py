import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

CBC_URL = os.environ.get("CBC_URL", "https://cbcinsider.com")
# Optional: go straight to the queue page instead of clicking through from the landing page.
CBC_QUEUE_URL = os.environ.get("CBC_QUEUE_URL", "")
# Optional: the Work Center your queue is filtered to (e.g. "ENGGL"). If set,
# the scraper refuses to run when the page is showing a different Work Center,
# so you never silently diff the wrong queue.
CBC_WORK_CENTER = os.environ.get("CBC_WORK_CENTER", "")
# Saved browser session (cookies) from login.py — no password is ever stored.
STORAGE_STATE_PATH = Path(os.path.expandvars(os.path.expanduser(
    os.environ.get("STORAGE_STATE_PATH", "./cbc_session.json")
)))

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
CLAUDE_MODEL = "claude-opus-4-7"


def _expand_path(raw: str) -> Path:
    """Expand ~ and Windows env vars like %USERPROFILE% in a configured path."""
    return Path(os.path.expandvars(os.path.expanduser(raw)))


OUTPUT_DIR = _expand_path(os.environ.get("OUTPUT_DIR", "./output"))
SNAPSHOT_DIR = _expand_path(os.environ.get("SNAPSHOT_DIR") or str(OUTPUT_DIR / "snapshots"))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)

# Email is sent through your local Outlook desktop app — no password needed.
# These are just the destination addresses (an address alone isn't sensitive).
EMAIL_TO = os.environ.get("EMAIL_TO", "")
EMAIL_ALERT_TO = os.environ.get("EMAIL_ALERT_TO") or EMAIL_TO


def validate_runtime_config() -> None:
    """Fail fast for the daily run if required settings are missing.

    Kept out of module import on purpose: helper scripts like login.py import
    this module only for CBC_URL / STORAGE_STATE_PATH and shouldn't need an
    Anthropic key or email address just to save a browser session.
    """
    missing = [k for k in ("ANTHROPIC_API_KEY", "EMAIL_TO") if not os.environ.get(k)]
    if missing:
        raise RuntimeError(
            "Missing required environment variable(s): "
            + ", ".join(missing)
            + ". Copy .env.example to .env and fill them in."
        )
