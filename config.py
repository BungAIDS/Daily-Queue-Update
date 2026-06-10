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
# Backfill (backfill_orders.py) looks up old orders through the queue page's
# "search order" / "find order" box. The box is normally auto-detected; if that
# misses on your layout, set the exact CSS selector here (run
# `discover_documents.py --probe <job#>`, which prints it). CBC_SEARCH_BUTTON is
# the optional selector of a search button to click when Enter doesn't submit.
CBC_SEARCH_SELECTOR = os.environ.get("CBC_SEARCH_SELECTOR", "").strip()
CBC_SEARCH_BUTTON = os.environ.get("CBC_SEARCH_BUTTON", "").strip()
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
# Model used by analyzer.py for the daily AI briefing.
#  - "claude-haiku-4-5"   ~$0.02-0.04/run, fastest, still solid quality for this
#  - "claude-sonnet-4-6"  ~$0.10/run, sharper prose
#  - "claude-opus-4-7"    ~$0.13/run, the heaviest; analyzer auto-enables
#                          adaptive thinking only for opus-4.x
CLAUDE_MODEL = "claude-haiku-4-5"


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

# Where the long-running backlog tools (autocad_scan.py, backfill_orders.py)
# keep their resumable progress stores and master workbooks. Kept separate from
# the daily queue report, which stays a one-day snapshot. Defaults under
# OUTPUT_DIR; override in .env if the backlog should live elsewhere.
BACKLOG_DIR = _output_path("BACKLOG_DIR", str(OUTPUT_DIR / "backlog"))
BACKLOG_DIR.mkdir(parents=True, exist_ok=True)

# Where downloaded sales-order PDFs are archived (one subfolder per job).
# Defaults to the Z: drive share; override in .env. A UNC path
# (\\server\share\...) is recommended for the scheduled task, since a mapped
# drive letter may not exist for a non-interactive session. NOT created at
# import — it may be an unmounted network drive when other scripts run, so the
# downloader creates it on use.
SALES_ORDER_DIR = _expand_path(
    (os.environ.get("SALES_ORDER_DIR") or r"Z:\DAG\SALES ORDERS FOR DAILY QUEUE").strip()
)
# Where downloaded construction/drive-run PDFs are archived (one subfolder per
# job). Only the highly-custom orders that actually have a CBC_DriveRun document
# get one. Same conventions as SALES_ORDER_DIR; created on use, not at import.
DRIVE_RUN_DIR = _expand_path(
    (os.environ.get("DRIVE_RUN_DIR") or r"Z:\DAG\DRIVE RUNS FOR DAILY QUEUE").strip()
)
# How many order-detail modals to open in parallel when fetching sales orders.
# Start modest (one login = one server session, which ASP.NET may serialize);
# raise it once you see how the server handles the load.
try:
    SO_CONCURRENCY = max(1, int(os.environ.get("SO_CONCURRENCY", "8")))
except ValueError:
    SO_CONCURRENCY = 8

# Root of the AutoCAD job folders on the Z: drive. A job appears under exactly
# one <type>\<intermediate>\<job> path, so finding its folder both gives the
# Excel link target and tells us the job type. Read-only lookup (not created).
AUTOCAD_JOBS_DIR = _expand_path(
    (os.environ.get("AUTOCAD_JOBS_DIR") or r"Z:\AUTOCAD\CURRENT\JOBS").strip()
)

# Email is sent through your local Outlook desktop app — no password needed.
# These are just the destination addresses (an address alone isn't sensitive).
EMAIL_TO = os.environ.get("EMAIL_TO", "")
EMAIL_ALERT_TO = os.environ.get("EMAIL_ALERT_TO") or EMAIL_TO


def validate_runtime_config() -> None:
    """Fail fast for the daily run if required settings are missing.

    Only EMAIL_TO is hard-required so the run can email its results. The
    Anthropic API key is optional: if it's empty the daily run still scrapes,
    diffs, builds the Excel, and emails it — just without the AI briefing.
    Helper scripts (login.py, check_access.py, dump_report.py) don't call
    this at all.
    """
    if not os.environ.get("EMAIL_TO"):
        raise RuntimeError(
            "Missing required environment variable EMAIL_TO. "
            "Open .env in Notepad and set EMAIL_TO to the address that should "
            "receive the daily briefing, then try again."
        )
