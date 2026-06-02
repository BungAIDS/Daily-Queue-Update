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

_raw_anthropic = (os.environ.get("ANTHROPIC_API_KEY") or "").strip()
# Treat the literal placeholder from .env.example as unset, so a fresh
# install (key not yet added) doesn't try the call and trigger a daily alert.
if _raw_anthropic in ("", "sk-ant-...", "sk-ant-..."):
    _raw_anthropic = ""
ANTHROPIC_API_KEY = _raw_anthropic
# Swap to "claude-opus-4-7" if you want the bigger model — costs ~3-5x more
# per run but is barely better on this small structured-output task.
CLAUDE_MODEL = "claude-sonnet-4-6"


def _expand_path(raw: str) -> Path:
    """Expand ~ and Windows env vars like %USERPROFILE% in a configured path."""
    return Path(os.path.expandvars(os.path.expanduser(raw)))


def _output_path(env_key: str, default: str) -> Path:
    raw = (os.environ.get(env_key) or "").strip()
    # Treat the legacy ".env.example" placeholder ("C:\Users\you\...") as unset,
    # so anyone who copies the template verbatim doesn't blow up on first run.
    if not raw or "\\you\\" in raw or "/you/" in raw:
        raw = default
    return _expand_path(raw)


_DEFAULT_OUTPUT = r"%USERPROFILE%\Documents\DailyQueue"
OUTPUT_DIR = _output_path("OUTPUT_DIR", _DEFAULT_OUTPUT)
SNAPSHOT_DIR = _output_path("SNAPSHOT_DIR", str(OUTPUT_DIR / "snapshots"))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)

# Email is sent through your local Outlook desktop app — no password needed.
# These are just the destination addresses (an address alone isn't sensitive).
EMAIL_TO = os.environ.get("EMAIL_TO", "")
EMAIL_ALERT_TO = os.environ.get("EMAIL_ALERT_TO") or EMAIL_TO


def validate_runtime_config() -> None:
    """Fail fast for the daily run if required settings are missing.

    Only EMAIL_TO is hard-required so the run can email its results. The
    Anthropic API key is optional: if it's empty the daily run still scrapes,
    diffs, builds the Excel, and emails it — just without the AI briefing.
    Helper scripts (login.py, check_access.py, make_report.py) don't call
    this at all.
    """
    if not os.environ.get("EMAIL_TO"):
        raise RuntimeError(
            "Missing required environment variable EMAIL_TO. "
            "Open .env in Notepad and set EMAIL_TO to the address that should "
            "receive the daily briefing, then try again."
        )
