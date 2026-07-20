import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

CBC_URL = os.environ.get("CBC_URL", "https://cbcinsider.com")
# Go straight to the queue page instead of clicking through from the landing page.
# Override in .env if CBC changes this route.
CBC_QUEUE_URL = (os.environ.get("CBC_QUEUE_URL")
                 or "https://www.cbcinsider.com/intranet/offsyte/dispatch.aspx").strip()
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

# --------------------------------------------------------------------------- #
# Drawing-transmittal send (fill_transmittal_insider.py) — DISABLED by design  #
# --------------------------------------------------------------------------- #
# The transmittal tooling fills the Word doc and can pre-fill the CBC Insider
# "Email Drawings" form, but the actual SEND is hard-disabled in code (the submit
# click is commented out) so a transmittal can never be mailed to a customer by
# accident. TRANSMITTAL_MODE only ever reaches "preview"/"review"; "send" is
# intentionally inert. Leave this alone unless you are deliberately enabling
# sends and have re-instated the submit code in fill_transmittal_insider.py.
TRANSMITTAL_MODE = (os.environ.get("TRANSMITTAL_MODE", "preview").strip().lower() or "preview")
# The CBC Insider "Email Drawings" page URL (Engineering -> Email Drawings).
# Defaults to the known transmittal page so the tooling navigates straight there
# without having to hunt for the nav link (which is in a hover menu and flaky to
# click). Override in .env if your instance differs.
EMAIL_DRAWINGS_URL = (os.environ.get("EMAIL_DRAWINGS_URL")
                      or "https://www.cbcinsider.com/intranet/engineering/transmittal.aspx").strip()
# Optional CSS-selector overrides for the Email Drawings flow, filled in after
# running `python fill_transmittal_insider.py --probe`. Blank => unknown / not set.
# The flow is TWO pages:
#   Page 1 (order lookup): type the order # and submit to advance.
#     EMAIL_DRAWINGS_ORDER_SELECTOR        the order # input on page 1
#     EMAIL_DRAWINGS_ORDER_SUBMIT_SELECTOR the button/link that advances to page 2
#                                          (leave blank to submit with Enter)
#   Page 2 (the email form, has the Send button we DON'T press):
#     EMAIL_DRAWINGS_EMAILS_SELECTOR       the recipients field
#     EMAIL_DRAWINGS_ATTACH_SELECTOR       the file <input type=file>
#     EMAIL_DRAWINGS_SUBMIT_SELECTOR       the SEND button — recorded so we know
#                                          what NOT to click; sending is disabled.
EMAIL_DRAWINGS_ORDER_SELECTOR = (os.environ.get("EMAIL_DRAWINGS_ORDER_SELECTOR") or "").strip()
EMAIL_DRAWINGS_ORDER_SUBMIT_SELECTOR = (os.environ.get("EMAIL_DRAWINGS_ORDER_SUBMIT_SELECTOR") or "").strip()
EMAIL_DRAWINGS_EMAILS_SELECTOR = (os.environ.get("EMAIL_DRAWINGS_EMAILS_SELECTOR") or "").strip()
EMAIL_DRAWINGS_ATTACH_SELECTOR = (os.environ.get("EMAIL_DRAWINGS_ATTACH_SELECTOR") or "").strip()
EMAIL_DRAWINGS_SUBMIT_SELECTOR = (os.environ.get("EMAIL_DRAWINGS_SUBMIT_SELECTOR") or "").strip()
# Saved browser session (cookies) from login.py — no password is ever stored.
STORAGE_STATE_PATH = Path(os.path.expandvars(os.path.expanduser(
    os.environ.get("STORAGE_STATE_PATH", "./cbc_session.json")
)))

_raw_anthropic = (os.environ.get("ANTHROPIC_API_KEY") or "").strip()
# Treat the literal placeholder from .env.example as unset, so a fresh
# install (key not yet added) doesn't try the call and trigger a daily alert.
if _raw_anthropic in ("", "sk-ant-..."):
    _raw_anthropic = ""
ANTHROPIC_API_KEY = _raw_anthropic
# Model used by analyzer.py for the daily AI briefing.
#  - "claude-haiku-4-5"   ~$0.02-0.04/run, fastest, still solid quality for this
#  - "claude-sonnet-4-6"  ~$0.10/run, sharper prose
#  - "claude-opus-4-7"    ~$0.13/run, the heaviest; analyzer auto-enables
#                          adaptive thinking only for opus-4.x
CLAUDE_MODEL = "claude-haiku-4-5"
# Model used by pdf_vision.py to read scanned (image-only) quote-run PDFs.
# Haiku reads a form page for a fraction of a cent; bump to a bigger model
# only if the worst scans come back sloppy (it's a one-string change here).
PDF_VISION_MODEL = (os.environ.get("PDF_VISION_MODEL") or "").strip() or CLAUDE_MODEL
# Pages sent per PDF (the run data is on the first page or two; later pages
# are usually part tables). Each page costs roughly 1-2k input tokens.
try:
    PDF_VISION_MAX_PAGES = max(1, int(os.environ.get("PDF_VISION_MAX_PAGES", "2")))
except ValueError:
    PDF_VISION_MAX_PAGES = 2


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


# %USERPROFILE% only expands on Windows; elsewhere (dev sandbox, Mac) fall back
# to ~ so importing this module doesn't create a literal "%USERPROFILE%\..."
# directory in the working folder.
if os.name == "nt":
    _DEFAULT_OUTPUT = r"%USERPROFILE%\Documents\DailyQueue"
else:
    _DEFAULT_OUTPUT = os.path.join("~", "Documents", "DailyQueue")
OUTPUT_DIR = _output_path("OUTPUT_DIR", _DEFAULT_OUTPUT)
SNAPSHOT_DIR = _output_path("SNAPSHOT_DIR", str(OUTPUT_DIR / "snapshots"))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)

# Where the watcher tees its console output (rotated daily, ~a week kept). Defaults
# to a `logs/` folder right next to the code so it's easy to find and share when a
# bug needs chasing; override with LOG_DIR in .env.
LOG_DIR = _output_path("LOG_DIR", str(Path(__file__).resolve().parent / "logs"))
LOG_DIR.mkdir(parents=True, exist_ok=True)

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
# The pid type(s) that identify the construction "drive run" / "quote run"
# document (pids look like <type>-<id>-<rev>-<tag>). Only the HDX fans file
# the run under a dedicated pid type; the run log prints every pid type it
# saw, so if yours is named differently just add it here (comma-separated,
# matching is case-insensitive and ignores a CBC_ prefix).
DRIVE_RUN_TYPES = [
    t.strip() for t in
    (os.environ.get("DRIVE_RUN_TYPES") or "CBC_DriveRun,CBC_QuoteRun").split(",")
    if t.strip()
]
# Everything that isn't an HDX files its quote run under a generic pid type
# (usually CBC_Inquiry), so those are recognized by FILE NAME instead — e.g.
# "421473_909-26-1604 Qt Run.txt", "420410 qt  run.txt",
# "Cascades Wheel Construction REV 2.docx", "... D64 Wheel Construction ...".
# Case-insensitive regexes, comma-separated; extend as new namings turn up.
# The "wheel"/"construction" tokens are deliberately BROAD — better to grab a
# construction doc and flag it (an over-grab that parses to nothing shows as a
# NO FIELDS row you can eyeball) than to silently miss a real run whose name we
# didn't anticipate. A stray over-grab can't corrupt an order's data: run_rank
# keeps a run that actually parsed ahead of an empty match at the same revision.
DRIVE_RUN_NAME_PATTERNS = [
    p.strip() for p in
    (os.environ.get("DRIVE_RUN_NAME_PATTERNS")
     or r"qt\s*run,quote\s*run,wheel,construction").split(",")
    if p.strip()
]
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

# Root of the per-job SolidWorks folders (solidworks_scan.py — feeds the
# explorer's "Has 3D" filter). Layout: <type>\<intermediate>\<job>, see
# solidworks_scan.py for the per-type intermediate rules. Read-only.
SOLIDWORKS_JOBS_DIR = _expand_path(
    (os.environ.get("SOLIDWORKS_JOBS_DIR") or r"Z:\Solidworks\Current\JOBS").strip()
)

# Where the per-job sales-order line items live (one JSON store, fed by the
# daily run, the backfill, and line_items_scan.py; searched by find_orders.py).
# Defaults under BACKLOG_DIR; created on first save.
_li_store_raw = (os.environ.get("LINE_ITEMS_STORE") or "").strip()
LINE_ITEMS_STORE = _expand_path(_li_store_raw) if _li_store_raw else None
# Optional JSON file that EXTENDS the built-in line-item rules (abbreviations,
# skip patterns, tag patterns) with site-specific wording — see line_items.py.
LINE_ITEM_RULES = (os.environ.get("LINE_ITEM_RULES") or "").strip()


def _env_float(name: str, default: float) -> float:
    try:
        return float((os.environ.get(name) or "").strip() or default)
    except ValueError:
        return default


# Similar-order suggester (find_orders.similar_to_items, wired into enrichment):
# each enriched order gets a shortlist of backlog orders that share its rare SO
# features AND already have custom DWGs on file ("DWG Reuse" column + hover note
# + new-order notification). Score = rarity-weighted overlap; on the ~6K-order
# corpus >= 0.5 means "genuinely the same fan" while common-feature noise sits
# well below it. Raise/lower REUSE_MIN_SCORE in .env if it's too chatty/quiet;
# REUSE_MIN_SCORE=99 effectively turns the suggester off.
REUSE_MIN_SCORE = _env_float("REUSE_MIN_SCORE", 0.5)
REUSE_TOP = int(_env_float("REUSE_TOP", 3))

# Email is sent through your local Outlook desktop app — no password needed.
# These are just the destination addresses (an address alone isn't sensitive).
EMAIL_TO = os.environ.get("EMAIL_TO", "")
EMAIL_ALERT_TO = os.environ.get("EMAIL_ALERT_TO") or EMAIL_TO


# --------------------------------------------------------------------------- #
# Live intraday watcher (watch.py)                                            #
# --------------------------------------------------------------------------- #
# The watcher polls the board every couple of minutes through the day and only
# does the slow per-order enrichment for orders that are NEW since the last
# poll, writing the result into a co-authored Excel workbook in real time.

# The Microsoft 365 co-authored workbook the live queue is written into. Use the
# LOCAL OneDrive/SharePoint-synced path (e.g.
# C:\Users\you\OneDrive - Company\Daily Queue\Live Queue.xlsx) so the desktop
# Excel app — which the watcher drives via COM — keeps it synced for coworkers.
_live_wb_raw = (os.environ.get("LIVE_WORKBOOK_PATH") or "").strip()
LIVE_WORKBOOK_PATH = _expand_path(_live_wb_raw) if _live_wb_raw else None

# Where the GL Queue Explorer page (order_explorer.py — the clickable HTML
# search over the line-items store) is written. The canonical default is the
# shared DAG folder coworkers open as Z:\DAG\GL QUEUE LIVE. Use the UNC path
# here so the watcher still reaches it when a scheduled/background session has
# no Z: mapping. EXPLORER_PATH may override this with a folder or .html file.
DEFAULT_EXPLORER_PATH = r"\\gdh-fs02\engineering\DAG\GL QUEUE LIVE"
_explorer_raw = (os.environ.get("EXPLORER_PATH") or DEFAULT_EXPLORER_PATH).strip()
EXPLORER_PATH = _expand_path(_explorer_raw)

# How often to poll the board, in seconds (default 120 = every 2 minutes).
try:
    POLL_INTERVAL_SECONDS = max(15, int(os.environ.get("POLL_INTERVAL_SECONDS", "120")))
except ValueError:
    POLL_INTERVAL_SECONDS = 120

# Backfill can change the similarity index after every scanned order. Repainting
# the grouped Similar Data tab for each tiny change monopolizes desktop Excel,
# while a 10-15 minute lag in ranking candidates is harmless. Queue membership
# changes still force an immediate refresh. Set 0 to refresh on every store write.
try:
    SIMILAR_REFRESH_INTERVAL_SECONDS = max(
        0, int(os.environ.get("SIMILAR_REFRESH_INTERVAL_SECONDS", "900"))
    )
except ValueError:
    SIMILAR_REFRESH_INTERVAL_SECONDS = 900

# The watcher drives the DESKTOP Excel app over COM all day and never quits it, so
# Excel keeps accumulating memory it doesn't fully reclaim (fragmented conditional-
# format rules, a growing calc chain, undo/redraw caches) — left unchecked it
# climbs into the multi-GB range and a rebuild stops finishing inside one poll.
# Every this-many polls the watcher recycles the live workbook: it closes it
# (AutoSave/co-authoring has already synced every edit, and only the bot's own
# Excel is touched — coworkers are unaffected) and the next poll reopens it fresh,
# which frees the accumulated memory. At the 120s default, 30 polls ≈ once an hour.
# Set 0 to disable.
try:
    EXCEL_RECYCLE_EVERY_POLLS = max(0, int(os.environ.get("EXCEL_RECYCLE_EVERY_POLLS", "30")))
except ValueError:
    EXCEL_RECYCLE_EVERY_POLLS = 30

# Background Sales-Order re-verification: each poll, re-check this many on-board
# orders we've gone longest without re-checking (round-robin), but only ones not
# re-checked within the last SO_REVERIFY_MIN_AGE_MIN minutes. This is what lets a
# silently-stale SO (e.g. an order left at an old revision by an earlier failed
# fetch) self-correct within the hour instead of waiting for the next daily run.
# Set SO_REVERIFY_PER_POLL=0 to disable. Costs ~ (per_poll) modal opens per poll.
try:
    SO_REVERIFY_PER_POLL = max(0, int(os.environ.get("SO_REVERIFY_PER_POLL", "2")))
except ValueError:
    SO_REVERIFY_PER_POLL = 2
try:
    SO_REVERIFY_MIN_AGE_MIN = max(1, int(os.environ.get("SO_REVERIFY_MIN_AGE_MIN", "45")))
except ValueError:
    SO_REVERIFY_MIN_AGE_MIN = 45

# Publish the watcher's log to a throwaway branch so it can be read remotely
# without copying files off the machine. The log is pushed at startup and again
# on shutdown (Ctrl+C, or the end of the watch window) — each push force-replaces
# the branch with the current log as a single orphan commit (no history, so no
# repo bloat). Set LOG_PUSH_BRANCH empty to disable. Needs an 'origin' you can
# push to. NOTE: the log (job #s, customers, file paths) goes to that repo.
LOG_PUSH_BRANCH = (os.environ.get("LOG_PUSH_BRANCH", "debug-logs") or "").strip()

# Publish the order-data files (live_master.json + the backlog/quote-run/
# line-item stores and their xlsx sheets) to a throwaway branch the same way as
# the log, so they can be read remotely without copying them off the machine.
# Each push force-replaces the branch with a single orphan commit (no history/
# bloat). Set DATA_PUSH_BRANCH empty to disable. Needs an 'origin' you can push
# to. NOTE: this data (job #s, customers, prices, file paths) goes to that repo
# — only point it at a PRIVATE repo.
DATA_PUSH_BRANCH = (os.environ.get("DATA_PUSH_BRANCH", "order-data") or "").strip()

# When true, the data is republished automatically after any flow that gathers
# order data updates the master (each scan/backfill via master_sync, and the
# live watch on new orders / at session end) — so a remote reader always tracks
# the latest. Off by default; set DATA_PUSH_ON_CHANGE=1 in .env to enable. Needs
# DATA_PUSH_BRANCH set. The manual "Publish Order Data" launcher task works
# regardless of this flag.
DATA_PUSH_ON_CHANGE = (os.environ.get("DATA_PUSH_ON_CHANGE", "") or "").strip().lower() in (
    "1", "true", "yes", "on")


def _parse_hhmm(raw: str, default: str) -> "tuple[int, int]":
    """Parse a 'HH:MM' watch-window bound into (hour, minute)."""
    raw = (raw or "").strip() or default
    try:
        h, m = raw.split(":", 1)
        return max(0, min(23, int(h))), max(0, min(59, int(m)))
    except ValueError:
        h, m = default.split(":")
        return int(h), int(m)


# Daily watch window (local time). Defaults to 05:00–17:00 (5am–5pm).
WATCH_START = _parse_hhmm(os.environ.get("WATCH_START", ""), "05:00")
WATCH_END = _parse_hhmm(os.environ.get("WATCH_END", ""), "17:00")

# Microsoft Teams Incoming Webhook URL. When set, each new order is posted to
# that channel so coworkers (and their phones) get notified — nothing to install
# on their machines. Leave blank to disable Teams notifications.
TEAMS_WEBHOOK_URL = (os.environ.get("TEAMS_WEBHOOK_URL") or "").strip()

# Pop a Windows toast on the watcher PC for each new order. On by default; set
# LIVE_TOAST=0/false/no to silence the local pop-ups (Teams still fires).
LIVE_TOAST_ENABLED = (os.environ.get("LIVE_TOAST", "1").strip().lower()
                      not in ("0", "false", "no", "off", ""))

# Save a dated, frozen copy of the workbook at the first poll each morning (the
# "what it looked like at the start of the day" snapshot). On by default.
LIVE_MORNING_SNAPSHOT = (os.environ.get("LIVE_MORNING_SNAPSHOT", "1").strip().lower()
                         not in ("0", "false", "no", "off", ""))

# The shareable WEB link to the live co-authored workbook (OneDrive/SharePoint
# "Copy link"). The daily 5 AM email sends this as an active link so the team
# opens the live sheet rather than a stale attachment. Blank = the email falls
# back to attaching the dated report.
LIVE_WORKBOOK_LINK = (os.environ.get("LIVE_WORKBOOK_LINK") or "").strip()

# Whether the daily email still attaches the dated .xlsx report. Default off when
# a live link is set (the link is the point); on otherwise so you still get the
# file. Set EMAIL_ATTACH_REPORT=1/0 to force it either way.
_attach_raw = (os.environ.get("EMAIL_ATTACH_REPORT") or "").strip().lower()
if _attach_raw in ("1", "true", "yes", "on"):
    EMAIL_ATTACH_REPORT = True
elif _attach_raw in ("0", "false", "no", "off"):
    EMAIL_ATTACH_REPORT = False
else:
    EMAIL_ATTACH_REPORT = not LIVE_WORKBOOK_LINK


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
