"""Desktop launcher for the Daily Queue helper scripts.

This is intentionally standard-library only. It gives the Windows workstation a
single friendly place to start/stop the daily queue tools, see live output, and
remember the last options used without adding another package to install.
"""
from __future__ import annotations

import json
import os
import queue
import shlex
import signal
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import git_update
import procscan
import stop_signal


ROOT = Path(__file__).resolve().parent
STATE_PATH = ROOT / ".launcher_state.json"
LOG_DIR = ROOT / "launcher_logs"
DEBUG_LOG = LOG_DIR / "launcher_debug.log"
# Local (git-ignored) copy of the debug report for eyeballing; the shared copy
# is pushed to the debug branch, so we keep this out of the tracked working tree.
LOCAL_REPORT = LOG_DIR / "launcher_report.txt"
# Dedicated branch that collects published debug reports, kept off feature
# branches. The launcher pushes here without disturbing the working checkout.
DEBUG_BRANCH = "debug/launcher"
# How long to let a graceful Stop finish (its current poll can run a while) before
# the backstop force-kill steps in. A second Stop click forces immediately.
GRACEFUL_STOP_SECONDS = 150


def hidden_console_kwargs() -> dict[str, Any]:
    if os.name != "nt":
        return {}
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = 0
    return {
        "creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0),
        "startupinfo": startupinfo,
    }


def windows_has_console() -> bool:
    """True if this process owns a console (False when launched via pythonw).

    When the launcher has no console of its own, a console-subsystem child
    (python.exe running a script) would otherwise pop its own black window.
    """
    if os.name != "nt":
        return False
    try:
        import ctypes

        return bool(ctypes.windll.kernel32.GetConsoleWindow())
    except Exception:
        return False


@dataclass(frozen=True)
class OptionSpec:
    key: str
    label: str
    help: str
    kind: str = "text"       # text, args, check, file, save_file, folder
    arg: str | None = None
    default: str | bool = ""
    positional: bool = False
    split: bool = False
    confirm: str = ""


@dataclass(frozen=True)
class LauncherAction:
    id: str
    category: str
    title: str
    description: str
    script: str | None = None
    default_args: tuple[str, ...] = ()
    options: tuple[OptionSpec, ...] = ()
    long_running: bool = False
    email_risk: bool = False
    script_option: str | None = None
    graceful_stop: bool = False   # Stop asks it to finish cleanly (via stop_signal) first


@dataclass
class ProcessInfo:
    action_id: str
    process: subprocess.Popen
    log_file: Any
    log_path: Path
    command: list[str]
    started_at: datetime = field(default_factory=datetime.now)


class ToolTip:
    def __init__(self, widget: tk.Widget, text: str) -> None:
        self.widget = widget
        self.text = text
        self.tip: tk.Toplevel | None = None
        widget.bind("<Enter>", self._show, add="+")
        widget.bind("<Leave>", self._hide, add="+")

    def _show(self, _event: tk.Event | None = None) -> None:
        if self.tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            self.tip,
            text=self.text,
            justify="left",
            wraplength=420,
            background="#ffffe0",
            relief="solid",
            borderwidth=1,
            padx=8,
            pady=5,
        )
        label.pack()

    def _hide(self, _event: tk.Event | None = None) -> None:
        if self.tip:
            self.tip.destroy()
            self.tip = None


def option(
    key: str,
    label: str,
    help_text: str,
    *,
    kind: str = "text",
    arg: str | None = None,
    default: str | bool = "",
    positional: bool = False,
    split: bool = False,
    confirm: str = "",
) -> OptionSpec:
    return OptionSpec(key, label, help_text, kind, arg, default, positional, split, confirm)


JOB_LIST = option(
    "jobs",
    "Job number(s)",
    "Space-separated job/order numbers. Example: 421473 421492 420410.",
    positional=True,
    split=True,
)
ARGUMENTS = option(
    "args",
    "Extra arguments",
    "Optional raw command-line arguments. Quotes are honored.",
    kind="args",
    positional=True,
    split=True,
)
LIMIT = option("limit", "Limit", "Stop after this many records/orders. Blank means no limit.", arg="--limit")
RESCAN = option("rescan", "Ignore saved progress / rescan", "Start over instead of using saved progress.", kind="check", arg="--rescan")
MIN_JOB = option("min_job", "Minimum job", "Skip folders below this job number.", arg="--min-job")
MAX_JOB = option("max_job", "Maximum job", "Skip folders above this job number. Blank or 0 means no cap.", arg="--max-job")
ROOT_FOLDER = option("root", "Root folder", "Override the AutoCAD jobs root from .env.", kind="folder", arg="--root")
OUT_FILE = option("out", "Output file", "Optional output workbook/file path.", kind="save_file", arg="--out")
LIST_FILE = option("list", "Job list file", "A text file with one job number per line.", kind="file", arg="--list")
RANGE = option("range", "Range", "Two job numbers: FIRST LAST. Example: 420000 421000.", kind="args", arg="--range", split=True)


def base_actions() -> list[LauncherAction]:
    actions = [
        LauncherAction(
            "login",
            "Daily Run",
            "Refresh CBC Login Session",
            "Opens a browser so you can log into cbcinsider manually. Saves cbc_session.json for the other tools.",
            "login.py",
            long_running=True,
        ),
        LauncherAction(
            "main",
            "Daily Run",
            "Run Full 5 AM Job",
            "Runs scrape -> AI briefing -> Excel report -> email. This can send email through Outlook.",
            "main.py",
            email_risk=True,
        ),
        LauncherAction(
            "scrape",
            "Daily Run",
            "Scrape + Diff + Excel",
            "Stage 1 of the daily run. Scrapes the queue, enriches it, diffs it, and writes Excel. No AI and no email.",
            "scrape.py",
        ),
        LauncherAction(
            "brief",
            "Daily Run",
            "Generate AI Briefing",
            "Stage 2. Reuses today's scrape and adds the Claude briefing. No email.",
            "brief.py",
        ),
        LauncherAction(
            "send",
            "Daily Run",
            "Send Latest Briefing",
            "Stage 3. Emails the newest report that has an AI overview. Use --dry-run to inspect without sending.",
            "send.py",
            options=(
                option("dry_run", "Dry run", "Do not send email; only show which report would be sent.", kind="check", arg="--dry-run"),
                option("path", "Report path", "Optional specific report file to send or inspect.", kind="file", positional=True),
            ),
            email_risk=True,
        ),
        LauncherAction(
            "dump_report",
            "Daily Run",
            "Inspect Latest Report",
            "Prints a readable summary of the latest Excel report, or a selected report.",
            "dump_report.py",
            options=(option("path", "Report path", "Optional report to inspect. Blank means latest.", kind="file", positional=True),),
        ),
        LauncherAction(
            "watch",
            "Live Watch",
            "Run Live Watcher All Day",
            "The main all-day script. Polls the board, enriches new orders, updates the live workbook, and sends notifications.",
            "watch.py",
            options=(option("now", "Start now / ignore time window", "Start immediately instead of waiting for the configured watch window.", kind="check", arg="--now"),),
            long_running=True,
            graceful_stop=True,
        ),
        LauncherAction(
            "watch_once",
            "Live Watch",
            "Run One Watch Poll",
            "Runs one watcher poll cycle now, then exits. Useful for testing setup.",
            "watch.py",
            default_args=("--once",),
        ),
        LauncherAction(
            "check_access",
            "Live Watch",
            "Check Queue Access",
            "Scrapes a small view of the queue to confirm your saved session and CBC_QUEUE_URL are working.",
            "check_access.py",
        ),
        LauncherAction(
            "autocad_scan",
            "Scans / Backfill",
            "Scan AutoCAD Custom DWGs",
            "Filesystem scan of AutoCAD job folders. Builds the custom drawing matrix and progress store.",
            "autocad_scan.py",
            options=(JOB_LIST, ROOT_FOLDER, OUT_FILE, option("recursive", "Scan subfolders", "Also scan below each job folder.", kind="check", arg="--recursive"), RESCAN, LIMIT, MIN_JOB, MAX_JOB),
            long_running=True,
        ),
        LauncherAction(
            "quote_run_scan",
            "Scans / Backfill",
            "Scan Quote Runs",
            "Filesystem sweep for quote/construction runs in job folders. Parses them through templates and writes an inventory.",
            "quote_run_scan.py",
            options=(JOB_LIST, LIST_FILE, RANGE, ROOT_FOLDER, OUT_FILE, RESCAN, option("reparse_attention", "Reparse attention rows", "Retry rows that were previously flagged for attention.", kind="check", arg="--reparse-attention"), LIMIT, MIN_JOB, MAX_JOB),
            long_running=True,
        ),
        LauncherAction(
            "line_items_scan",
            "Scans / Backfill",
            "Scan Sales Order Line Items",
            "Reads archived Sales Order PDFs into the searchable line-items store.",
            "line_items_scan.py",
            options=(
                option("jobs", "Job(s) or PDF path(s)", "Space-separated job numbers or PDF paths. Blank scans the whole archive.", kind="args", positional=True, split=True),
                option("dump", "Dump parse details", "Show exactly what the extractor captures/skips for selected jobs.", kind="check", arg="--dump"),
                option("renorm", "Re-normalize existing store", "Reapply current rules to already captured raw text.", kind="check", arg="--renorm"),
                option("ai", "Use AI tagging", "Classify still-untagged unique items with the configured AI API.", kind="check", arg="--ai", confirm="This may call the AI API and update cached tags."),
                RESCAN,
                LIMIT,
            ),
            long_running=True,
        ),
        LauncherAction(
            "backfill_orders",
            "Scans / Backfill",
            "Backfill Historical Orders",
            "Long-running resumable browser job for old orders. Downloads/parses Sales Orders and quote runs, then writes backlog.xlsx.",
            "backfill_orders.py",
            options=(JOB_LIST, LIST_FILE, RANGE, ROOT_FOLDER, OUT_FILE, option("delay", "Delay seconds", "Pause between orders. Default is the script's built-in value.", arg="--delay"), LIMIT, MIN_JOB, MAX_JOB, RESCAN),
            long_running=True,
        ),
        LauncherAction(
            "master_sync",
            "Scans / Backfill",
            "Merge Helper Stores Into Master",
            "Consolidates AutoCAD, quote-run, line-item, and backfill stores into live_master.json.",
            "master_sync.py",
            options=(option("sources", "Sources", "Optional source names: autocad quote_runs line_items backfill. Blank merges all.", kind="args", positional=True, split=True),),
        ),
        LauncherAction(
            "find_orders",
            "Search / Inspect",
            "Find Orders By Line Items",
            "Search the line-items store by terms, tag, job number, or write an Excel inventory.",
            "find_orders.py",
            options=(
                option("terms", "Search terms", "Space-separated terms. By default all terms must match.", kind="args", positional=True, split=True),
                option("any", "Match any term", "Match any search term instead of all terms.", kind="check", arg="--any"),
                option("tag", "Feature tag", "Filter by canonical feature tag, such as SHAFT SEAL.", arg="--tag"),
                option("fuzzy", "Fuzzy ratio", "Optional typo-tolerant ratio. Example: 0.84.", arg="--fuzzy"),
                option("job", "Single job", "Show stored items for one job number.", arg="--job"),
                option("list_tags", "List tags", "Show the live tag vocabulary and counts.", kind="check", arg="--list-tags"),
                option("xlsx", "Write Excel workbook", "Write an Excel inventory/matrix. If output path is blank, the script chooses the default.", kind="check", arg="--xlsx"),
            ),
        ),
        LauncherAction(
            "check_orders",
            "Search / Inspect",
            "Check Order Quote Runs",
            "Given order numbers, opens each order and reports the matched quote-run template and fields.",
            "check_orders.py",
            options=(JOB_LIST, option("show", "Show browser", "Run with a visible browser instead of headless.", kind="check", arg="--show")),
            long_running=True,
        ),
        LauncherAction(
            "discover_documents",
            "Search / Inspect",
            "Discover Order Documents",
            "Lists a job's online documents and can probe the old-order search box.",
            "discover_documents.py",
            options=(option("args", "Job / arguments", "Example: 421473 or --probe 421473.", kind="args", positional=True, split=True),),
            long_running=True,
        ),
        LauncherAction(
            "templates",
            "Search / Inspect",
            "Inspect Quote Run Template",
            "Reads one quote/construction run file and shows which parser template matched it.",
            "templates.py",
            options=(
                option("path", "Quote run file", "The quote/construction run file to inspect.", kind="file", positional=True),
                option("design", "Design number", "Optional design number to help template matching.", positional=True),
            ),
        ),
        LauncherAction(
            "dump_pdf",
            "Search / Inspect",
            "Dump PDF Text/Tables",
            "Extracts text and tables from a selected PDF to help tune parsers.",
            "dump_pdf.py",
            options=(option("path", "PDF path", "PDF file to inspect.", kind="file", positional=True),),
        ),
        LauncherAction(
            "find_job_folder",
            "Search / Inspect",
            "Find AutoCAD Job Folder",
            "Searches the configured AutoCAD root for one or more job folders.",
            "find_job_folder.py",
            options=(JOB_LIST,),
        ),
        LauncherAction(
            "drive_run",
            "Search / Inspect",
            "Inspect Drive/Quote Run PDF",
            "Parses one drive-run PDF and prints the detected fields and reconstructed lines.",
            "drive_run.py",
            options=(option("path", "PDF path", "Drive/quote-run PDF file.", kind="file", positional=True),),
        ),
        LauncherAction(
            "extract_so",
            "Search / Inspect",
            "Inspect Sales Order PDF",
            "Dumps key Sales Order fields from one or more PDFs.",
            "extract_so.py",
            options=(option("paths", "PDF path(s)", "One or more Sales Order PDFs.", kind="args", positional=True, split=True),),
        ),
        LauncherAction(
            "email_drawings",
            "Transmittals",
            "Prepare Email Drawings Form",
            "Builds/fills transmittal data and pre-fills CBC Insider's Email Drawings form. The code intentionally does not click Send. Runs once: it pauses with the form open for you to review, then click 'Send Enter' to close it.",
            "fill_transmittal_insider.py",
            options=(
                option("order", "Order", "Order number to prepare.", positional=True),
                option("probe", "Probe form fields", "Read-only selector discovery mode.", kind="check", arg="--probe"),
                option("headless", "Headless browser", "Run browser hidden.", kind="check", arg="--headless"),
                option("initials", "Initials", "Override signature initials.", arg="--initials"),
                option("no_doc", "Skip Word transmittal", "Do not generate/fill the Word transmittal document.", kind="check", arg="--no-doc"),
            ),
        ),
        LauncherAction(
            "transmittal_data",
            "Transmittals",
            "Show Transmittal Data",
            "Prints the data that would feed a drawing transmittal for one or more orders.",
            "transmittal_data.py",
            options=(
                option("orders", "Order number(s)", "One or more order numbers.", kind="args", positional=True, split=True),
                option("backtest", "Backtest", "Use historical/backtest behavior from the script.", kind="check", arg="--backtest"),
                option("no_refresh", "Do not refresh", "Skip live refresh work.", kind="check", arg="--no-refresh"),
                option("customer", "Customer override", "Override customer name.", arg="--customer"),
            ),
        ),
        LauncherAction(
            "transmittal_doc",
            "Transmittals",
            "Fill Drawing Transmittal Doc",
            "Creates/fills a Word Drawing Transmittal document for an order. Use plan-only to preview without Word COM.",
            "transmittal_doc.py",
            options=(
                option("order", "Order", "Order number.", positional=True),
                option("initials", "Initials", "Override signature initials.", arg="--initials"),
                option("customer", "Customer override", "Override customer name.", arg="--customer"),
                option("out", "Output .doc path", "Optional output path for the generated document.", kind="save_file", arg="--out"),
                option("plan_only", "Plan only", "Print the fill plan without opening Word.", kind="check", arg="--plan-only"),
            ),
        ),
        LauncherAction(
            "discover_sales_order",
            "Tools",
            "Discover Sales Order Page",
            "Experimental helper for discovering Sales Order links from the site.",
            "discover_sales_order.py",
            options=(option("job", "Job", "Optional job/order number.", positional=True),),
            long_running=True,
        ),
        LauncherAction(
            "fetch_sales_orders",
            "Tools",
            "Fetch Sales Orders",
            "Standalone fetch helper. Usually the daily run/watch/backfill call this for you.",
            "fetch_sales_orders.py",
            options=(option("limit", "Limit", "Optional count limit.", positional=True),),
            long_running=True,
        ),
        LauncherAction(
            "seed_yesterday",
            "Tools",
            "Seed Yesterday Snapshot",
            "Copies the included baseline seed into yesterday's snapshot. Mostly setup/testing.",
            "seed_yesterday.py",
        ),
        LauncherAction(
            "scrub_baseline",
            "Tools",
            "Scrub Baseline Change Log",
            "Removes the first baseline batch from a change log date.",
            "scrub_baseline.py",
            options=(option("date", "Date", "Date in YYYY-MM-DD. Blank means script default.", positional=True),),
        ),
        LauncherAction(
            "custom_script",
            "Tools",
            "Run Any Python Script",
            "Advanced escape hatch. Pick any .py file in this project and pass optional arguments.",
            options=(
                option("script", "Script", "Python script to run.", kind="file"),
                ARGUMENTS,
            ),
            script_option="script",
            long_running=True,
        ),
    ]

    for path in sorted(ROOT.glob("test_*.py")):
        actions.append(
            LauncherAction(
                f"dev_{path.stem}",
                "Developer",
                path.name,
                "Direct-script test file. These are normally run by CI; use this tab when changing code.",
                path.name,
            )
        )
    return actions


class LauncherApp(tk.Tk):
    categories = [
        "Daily Run",
        "Live Watch",
        "Scans / Backfill",
        "Search / Inspect",
        "Transmittals",
        "Tools",
        "Developer",
    ]

    def __init__(self) -> None:
        super().__init__()
        self.title("Daily Queue Script Launcher")
        self.geometry("1180x760")
        self.minsize(980, 620)

        self.actions = base_actions()
        self.actions_by_id = {a.id: a for a in self.actions}
        self._exclusion_args = self._compute_exclusion_args()
        self.processes: dict[str, ProcessInfo] = {}
        self.external_running: set[str] = set()
        self.external_pids: dict[str, list[int]] = {}
        self.last_exit: dict[str, int] = {}
        self._stop_requested: set[str] = set()   # ids we asked to stop (don't flag as FAIL)
        self.stopped_actions: set[str] = set()    # ids that exited because we stopped them
        self._scan_in_progress = False            # a background process scan is running
        self.logs: dict[str, list[str]] = {a.id: [] for a in self.actions}
        self.output_queue: queue.Queue[tuple[str, str, Any]] = queue.Queue()
        self.option_vars: dict[str, dict[str, tk.Variable]] = {}
        self.option_widgets: list[tk.Widget] = []
        self.listboxes: dict[str, tk.Listbox] = {}
        self.listbox_items: dict[str, list[str]] = {}
        self.current_action_id: str | None = None
        self._git_dialog: GitUpdateDialog | None = None
        self._last_status_key: Any = None
        self.state = self._load_state()
        # Programs stopped for a git update, persisted so they can be restarted
        # the next time the launcher opens (after the new code is in place).
        self.pending_restart: list[str] = [
            aid for aid in (self.state.get("pending_restart") or []) if aid in self.actions_by_id
        ]

        self.python_path = self._find_python()
        self.allow_send_var = tk.BooleanVar(value=False)

        self._debug(f"launcher started (os={os.name}, python={self.python_path})")
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._select_initial_action()
        self._drain_output()
        self._refresh_external_status()
        self._tick_status()
        if self.pending_restart:
            self.after(800, self._restart_pending)

    def _load_state(self) -> dict[str, Any]:
        if not STATE_PATH.exists():
            return {"options": {}, "last_action": ""}
        try:
            return json.loads(STATE_PATH.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return {"options": {}, "last_action": ""}

    def _save_state(self) -> None:
        # Start from previously saved options so actions we never opened this
        # session keep their remembered values instead of being dropped.
        options = dict(self.state.get("options") or {})
        for action_id, vars_by_key in self.option_vars.items():
            options[action_id] = {k: v.get() for k, v in vars_by_key.items()}
        data = {
            "options": options,
            "last_action": self.current_action_id or "",
            "pending_restart": self.pending_restart,
        }
        self.state = data
        try:
            STATE_PATH.write_text(json.dumps(data, indent=2, default=str), encoding="utf-8")
        except OSError as exc:
            messagebox.showwarning("Could not save launcher settings", str(exc))

    def _find_python(self) -> Path:
        candidates = [
            ROOT / "venv" / "Scripts" / "python.exe",
            ROOT / ".venv" / "Scripts" / "python.exe",
        ]
        for path in candidates:
            if path.exists():
                return path
        return Path(sys.executable)

    def _build_ui(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Action.TButton", padding=(12, 6))
        style.configure("Danger.TCheckbutton", foreground="#7a1f1f")

        header = ttk.Frame(self, padding=(12, 10, 12, 4))
        header.pack(fill="x")
        ttk.Label(header, text="Daily Queue Script Launcher", style="Title.TLabel").pack(side="left")
        ttk.Label(header, text=f"Python: {self.python_path}", foreground="#555").pack(side="left", padx=(18, 0))
        send_check = ttk.Checkbutton(
            header,
            text="Allow email / send actions",
            variable=self.allow_send_var,
            style="Danger.TCheckbutton",
        )
        send_check.pack(side="right")
        ToolTip(send_check, "Required before running commands that can send email or prefill email forms.")

        git_button = ttk.Button(header, text="Git Update…", command=self._open_git_update)
        git_button.pack(side="right", padx=(0, 16))
        ToolTip(git_button, "Pull the latest code from a chosen Git branch.")

        main = ttk.PanedWindow(self, orient="horizontal")
        main.pack(fill="both", expand=True, padx=12, pady=(4, 12))

        left = ttk.Frame(main)
        main.add(left, weight=1)
        right = ttk.Frame(main)
        main.add(right, weight=3)

        search_frame = ttk.Frame(left)
        search_frame.pack(fill="x", pady=(0, 8))
        ttk.Label(search_frame, text="Search").pack(side="left")
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        search_entry.pack(side="left", fill="x", expand=True, padx=(8, 0))
        search_entry.bind("<KeyRelease>", lambda _e: self._populate_lists())
        ToolTip(search_entry, "Filter commands by name or description.")

        self.notebook = ttk.Notebook(left)
        self.notebook.pack(fill="both", expand=True)
        for category in self.categories:
            frame = ttk.Frame(self.notebook, padding=6)
            self.notebook.add(frame, text=category)
            listbox = tk.Listbox(frame, activestyle="dotbox", exportselection=False)
            listbox.pack(side="left", fill="both", expand=True)
            scroll = ttk.Scrollbar(frame, orient="vertical", command=listbox.yview)
            scroll.pack(side="right", fill="y")
            listbox.configure(yscrollcommand=scroll.set)
            listbox.bind("<<ListboxSelect>>", lambda _e, cat=category: self._on_list_select(cat))
            self.listboxes[category] = listbox
        self.notebook.bind("<<NotebookTabChanged>>", lambda _e: self._select_first_in_current_tab())

        detail_top = ttk.Frame(right)
        detail_top.pack(fill="x")
        self.status_light = tk.Canvas(detail_top, width=18, height=18, highlightthickness=0)
        self.status_light.pack(side="left", padx=(0, 8))
        self.action_title_var = tk.StringVar(value="")
        ttk.Label(detail_top, textvariable=self.action_title_var, style="Title.TLabel").pack(side="left")
        self.status_var = tk.StringVar(value="")
        ttk.Label(detail_top, textvariable=self.status_var, foreground="#555").pack(side="right")

        self.description = tk.Text(right, height=5, wrap="word", relief="solid", borderwidth=1)
        self.description.pack(fill="x", pady=(8, 8))
        self.description.configure(state="disabled")

        preview_frame = ttk.LabelFrame(right, text="Command Preview", padding=8)
        preview_frame.pack(fill="x")
        self.command_preview = tk.Text(preview_frame, height=2, wrap="word", relief="flat", background="#f6f6f6")
        self.command_preview.pack(fill="x")
        self.command_preview.configure(state="disabled")

        self.options_frame = ttk.LabelFrame(right, text="Options", padding=8)
        self.options_frame.pack(fill="x", pady=(8, 8))

        button_row = ttk.Frame(right)
        button_row.pack(fill="x", pady=(0, 8))
        self.run_button = ttk.Button(button_row, text="Run", style="Action.TButton", command=self._run_selected)
        self.run_button.pack(side="left")
        self.stop_button = ttk.Button(button_row, text="Stop", command=self._stop_selected)
        self.stop_button.pack(side="left", padx=(8, 0))
        self.enter_button = ttk.Button(button_row, text="Send Enter", command=self._send_enter)
        self.enter_button.pack(side="left", padx=(8, 0))
        self.clear_button = ttk.Button(button_row, text="Clear Output", command=self._clear_output)
        self.clear_button.pack(side="left", padx=(8, 0))
        self.logs_button = ttk.Button(button_row, text="Open Logs Folder", command=self._open_logs_folder)
        self.logs_button.pack(side="left", padx=(8, 0))
        self.report_button = ttk.Button(button_row, text="Publish Debug Report", command=self._publish_debug_report)
        self.report_button.pack(side="left", padx=(8, 0))
        ToolTip(self.report_button, f"Save diagnostics/launcher_report.txt and push it to the '{DEBUG_BRANCH}' branch for review.")
        self.refresh_button = ttk.Button(button_row, text="Refresh Status", command=lambda: self._refresh_external_status(schedule=False, verbose=True))
        self.refresh_button.pack(side="right")
        ToolTip(self.refresh_button, "Re-scan for tools running outside the launcher and show a diagnostic (also written to launcher_debug.log).")

        output_frame = ttk.LabelFrame(right, text="Output", padding=8)
        output_frame.pack(fill="both", expand=True)
        self.output = tk.Text(output_frame, wrap="word", relief="solid", borderwidth=1)
        self.output.pack(side="left", fill="both", expand=True)
        output_scroll = ttk.Scrollbar(output_frame, orient="vertical", command=self.output.yview)
        output_scroll.pack(side="right", fill="y")
        self.output.configure(yscrollcommand=output_scroll.set, state="disabled")

        self._populate_lists()

    def _populate_lists(self) -> None:
        query = self.search_var.get().strip().lower()
        for category, listbox in self.listboxes.items():
            listbox.delete(0, tk.END)
            items = []
            for action in self.actions:
                if action.category != category:
                    continue
                haystack = f"{action.title} {action.description} {action.script or ''}".lower()
                if query and query not in haystack:
                    continue
                items.append(action.id)
                listbox.insert(tk.END, self._display_title(action))
            self.listbox_items[category] = items

    def _display_title(self, action: LauncherAction) -> str:
        if self._is_running(action.id):
            return f"[RUNNING] {action.title}"
        if action.id in self.external_running:
            return f"[EXTERNAL] {action.title}"
        if action.id in self.stopped_actions:
            return f"[STOPPED] {action.title}"
        if action.id in self.last_exit:
            return f"[OK] {action.title}" if self.last_exit[action.id] == 0 else f"[FAIL] {action.title}"
        return action.title

    def _select_initial_action(self) -> None:
        preferred = self.state.get("last_action") or "watch"
        if preferred in self.actions_by_id:
            self._select_action(preferred)
            return
        self._select_first_in_current_tab()

    def _select_first_in_current_tab(self) -> None:
        category = self.notebook.tab(self.notebook.select(), "text")
        items = self.listbox_items.get(category) or []
        # Selecting an action programmatically switches tabs, which fires this
        # handler. Don't override a selection that already lives in this tab
        # (e.g. the remembered last action when it isn't first in its list).
        if self.current_action_id in items:
            return
        if items:
            self._select_action(items[0])

    def _on_list_select(self, category: str) -> None:
        listbox = self.listboxes[category]
        selection = listbox.curselection()
        if not selection:
            return
        action_id = self.listbox_items[category][selection[0]]
        self._select_action(action_id)

    def _select_action(self, action_id: str) -> None:
        action = self.actions_by_id[action_id]
        self.current_action_id = action_id
        for category, items in self.listbox_items.items():
            if action_id in items:
                self.notebook.select(list(self.listboxes).index(category))
                listbox = self.listboxes[category]
                idx = items.index(action_id)
                listbox.selection_clear(0, tk.END)
                listbox.selection_set(idx)
                listbox.see(idx)
                break
        self.action_title_var.set(action.title)
        self._set_text(self.description, action.description)
        self._build_options(action)
        self._render_output(action_id)
        self._update_detail_status()
        self._update_command_preview()
        self._save_state()

    def _build_options(self, action: LauncherAction) -> None:
        for child in self.options_frame.winfo_children():
            child.destroy()
        self.option_widgets = []
        vars_by_key: dict[str, tk.Variable] = {}
        saved = (self.state.get("options") or {}).get(action.id, {})
        self.option_vars[action.id] = vars_by_key
        if not action.options:
            ttk.Label(self.options_frame, text="No options for this command.").grid(row=0, column=0, sticky="w")
            return
        self.options_frame.columnconfigure(1, weight=1)
        for row, spec in enumerate(action.options):
            label = ttk.Label(self.options_frame, text=spec.label)
            label.grid(row=row, column=0, sticky="w", padx=(0, 8), pady=3)
            ToolTip(label, spec.help)
            if spec.kind == "check":
                var = tk.BooleanVar(value=bool(saved.get(spec.key, spec.default)))
                widget = ttk.Checkbutton(self.options_frame, variable=var)
                widget.grid(row=row, column=1, sticky="w", pady=3)
            else:
                var = tk.StringVar(value=str(saved.get(spec.key, spec.default or "")))
                widget = ttk.Entry(self.options_frame, textvariable=var)
                widget.grid(row=row, column=1, sticky="ew", pady=3)
                if spec.kind in ("file", "save_file", "folder"):
                    button = ttk.Button(
                        self.options_frame,
                        text="Browse",
                        command=lambda s=spec, v=var: self._browse(s, v),
                    )
                    button.grid(row=row, column=2, padx=(8, 0), pady=3)
                    ToolTip(button, f"Choose {spec.label.lower()}.")
            vars_by_key[spec.key] = var
            self.option_widgets.append(widget)
            ToolTip(widget, spec.help)
            var.trace_add("write", lambda *_args: self._update_command_preview())

    def _browse(self, spec: OptionSpec, var: tk.StringVar) -> None:
        if spec.kind == "folder":
            value = filedialog.askdirectory(initialdir=str(ROOT))
        elif spec.kind == "save_file":
            value = filedialog.asksaveasfilename(initialdir=str(ROOT))
        else:
            value = filedialog.askopenfilename(initialdir=str(ROOT))
        if value:
            var.set(value)

    def _build_command(self, action: LauncherAction, values: dict[str, Any] | None = None) -> list[str]:
        # Default to the live UI option values; callers (e.g. restart) can pass a
        # saved-state mapping when the action's widgets aren't currently built.
        if values is None:
            values = {k: v.get() for k, v in self.option_vars.get(action.id, {}).items()}

        script = action.script
        if action.script_option:
            raw_script = str(values.get(action.script_option, "")).strip()
            if not raw_script:
                raise ValueError("Choose a script first.")
            path = Path(raw_script)
            if not path.is_absolute():
                path = ROOT / path
            if path.suffix.lower() != ".py":
                raise ValueError("The custom runner only accepts .py files.")
            script = str(path)
        if not script:
            raise ValueError("This command has no script configured.")

        command = [str(self.python_path), str(script)]
        command.extend(action.default_args)
        for spec in action.options:
            if spec.key == action.script_option or spec.key not in values:
                continue
            value = values[spec.key]
            if spec.kind == "check":
                if bool(value) and spec.arg:
                    command.append(spec.arg)
                continue
            raw = str(value).strip()
            if not raw:
                continue
            pieces = shlex.split(raw, posix=False) if spec.split else [raw]
            if spec.arg:
                command.append(spec.arg)
            command.extend(pieces)
        return command

    def _command_text(self, command: Iterable[str]) -> str:
        # Render the way subprocess actually builds the command line on
        # Windows, so the preview, confirm dialog, and log header all show a
        # runnable command rather than posix-style single-quoting.
        return subprocess.list2cmdline([str(part) for part in command])

    def _update_command_preview(self) -> None:
        if not self.current_action_id:
            return
        action = self.actions_by_id[self.current_action_id]
        try:
            text = self._command_text(self._build_command(action))
        except Exception as exc:
            text = f"(not ready: {exc})"
        self._set_text(self.command_preview, text)

    def _run_selected(self) -> None:
        if not self.current_action_id:
            return
        action = self.actions_by_id[self.current_action_id]
        if self._is_running(action.id):
            messagebox.showinfo("Already running", f"{action.title} is already running from this launcher.")
            return
        if action.long_running:
            self._scan_now_sync()
            if action.id in self.external_running:
                messagebox.showwarning(
                    "Already running",
                    f"{action.title} already appears to be running outside this launcher.\n\n"
                    "Stop that copy before starting another one.",
                )
                return
        if action.email_risk and not self.allow_send_var.get():
            messagebox.showwarning(
                "Email action locked",
                "Check 'Allow email / send actions' before running this command.",
            )
            return
        try:
            command = self._build_command(action)
        except ValueError as exc:
            messagebox.showwarning("Missing option", str(exc))
            return

        for spec in action.options:
            if spec.confirm and bool(self.option_vars[action.id][spec.key].get()):
                if not messagebox.askyesno("Extra confirmation", spec.confirm):
                    return

        message = f"Run this command?\n\n{self._command_text(command)}"
        if not messagebox.askyesno("Confirm run", message):
            return

        self._save_state()
        try:
            self._start_process(action, command)
        except OSError as exc:
            messagebox.showerror("Could not start command", str(exc))

    def _start_process(self, action: LauncherAction, command: list[str]) -> None:
        """Launch ``command`` for ``action`` and wire up logging/output threads.

        This is the shared core behind the Run button and the Git Update
        auto-restart. It performs no confirmation or gating — callers do that.
        Raises ``OSError`` if the process cannot be started.
        """
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = LOG_DIR / f"{stamp}_{action.id}.log"
        log_file = log_path.open("w", encoding="utf-8", errors="replace")
        log_file.write(f"$ {self._command_text(command)}\n\n")
        log_file.flush()

        creationflags = 0
        if os.name == "nt":
            creationflags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
            # Under pythonw (no console of our own) a console child pops its own
            # window — hide it. When we DO have a console, the child shares it
            # (no new window) and Ctrl+Break can still reach it for a graceful Stop.
            if not windows_has_console():
                creationflags |= getattr(subprocess, "CREATE_NO_WINDOW", 0)
        try:
            process = subprocess.Popen(
                command,
                cwd=str(ROOT),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                creationflags=creationflags,
            )
        except OSError:
            log_file.close()
            raise

        # Fresh run clears any prior stopped/failed status for this action.
        self.stopped_actions.discard(action.id)
        self._stop_requested.discard(action.id)
        self.last_exit.pop(action.id, None)
        self.processes[action.id] = ProcessInfo(action.id, process, log_file, log_path, command)
        self.logs[action.id] = [f"$ {self._command_text(command)}\n", f"[launcher] Log: {log_path}\n"]
        if self.current_action_id == action.id:
            self._render_output(action.id)
        self._update_all_statuses()
        threading.Thread(target=self._reader_thread, args=(action.id, process), daemon=True).start()
        threading.Thread(target=self._waiter_thread, args=(action.id, process), daemon=True).start()

    def restart_action(self, action_id: str) -> None:
        """Rebuild an action's command and start it.

        Uses the live option widgets when they exist, otherwise the saved
        options from the state file (so a restart can happen at startup before
        the action has been selected). Raises on failure so the caller reports it.
        """
        action = self.actions_by_id[action_id]
        if action.id in self.option_vars:
            command = self._build_command(action)
        else:
            saved = (self.state.get("options") or {}).get(action.id, {})
            command = self._build_command(action, values=saved)
        self._start_process(action, command)

    def running_launcher_processes(self, exclude: Iterable[str] = ()) -> list[tuple[str, str]]:
        """(id, title) for every tool this launcher started that is still alive."""
        excluded = set(exclude)
        live: list[tuple[str, str]] = []
        for action_id, info in list(self.processes.items()):
            if action_id in excluded:
                continue
            action = self.actions_by_id.get(action_id)
            if action and info.process.poll() is None:
                live.append((action_id, action.title))
        return live

    def stoppable_external_processes(self) -> list[tuple[str, str]]:
        """(id, title) for external tools we have a PID for and can force-stop."""
        return [
            (aid, self.actions_by_id[aid].title)
            for aid in sorted(self.external_running)
            if aid in self.actions_by_id and self.external_pids.get(aid)
        ]

    def unstoppable_external_processes(self) -> list[str]:
        """Titles of external tools we detected but have no PID to stop."""
        return [
            self.actions_by_id[aid].title
            for aid in sorted(self.external_running)
            if aid in self.actions_by_id and not self.external_pids.get(aid)
        ]

    def relaunch_self(self) -> bool:
        """Start a fresh copy of the launcher (e.g. to run freshly pulled code).

        Spawns a new, detached process using the same interpreter that is
        running this launcher, so it keeps the windowed/console behaviour of the
        original. Returns True if the new process started.
        """
        cmd = [sys.executable, str(Path(__file__).resolve())]
        kwargs: dict[str, Any] = {"cwd": str(ROOT), "close_fds": True}
        if os.name == "nt":
            kwargs["creationflags"] = (
                getattr(subprocess, "DETACHED_PROCESS", 0)
                | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
            )
        else:
            kwargs["start_new_session"] = True
        try:
            subprocess.Popen(cmd, **kwargs)
            return True
        except OSError as exc:
            self._debug(f"relaunch failed: {exc!r}")
            return False

    def relaunch_self_and_exit(self) -> None:
        """Spawn a fresh launcher on the new code and close this one."""
        if not self.relaunch_self():
            messagebox.showwarning(
                "Could not relaunch",
                "The update was applied, but the launcher could not relaunch "
                "itself. Please close and reopen it manually.",
            )
            return
        self._debug("relaunching launcher; closing this instance")
        self._save_state()
        self.destroy()

    def set_pending_restart(self, action_ids: list[str]) -> None:
        """Remember (persistently) which programs to restart when the launcher reopens."""
        self.pending_restart = list(action_ids)
        self._save_state()

    def _restart_pending(self) -> None:
        """Restart the programs that were stopped for a git update last session."""
        pending = list(self.pending_restart)
        # Clear and persist first so a failure can't make us loop-restart forever.
        self.pending_restart = []
        self._save_state()

        restarted: list[str] = []
        skipped: list[str] = []
        failed: list[str] = []
        for action_id in pending:
            action = self.actions_by_id.get(action_id)
            if action is None or self._is_running(action_id):
                continue
            if action.email_risk:  # never auto-launch something that can send email
                skipped.append(action.title)
                continue
            try:
                self.restart_action(action_id)
                restarted.append(action.title)
            except Exception as exc:
                failed.append(f"{action.title}: {exc}")

        self._debug(f"pending restart: restarted={restarted} skipped={skipped} failed={failed}")
        parts = []
        if restarted:
            parts.append("Restarted after the update:\n  " + "\n  ".join(restarted))
        if skipped:
            parts.append("Not auto-restarted (can send email — start manually):\n  " + "\n  ".join(skipped))
        if failed:
            parts.append("Could not restart:\n  " + "\n  ".join(failed))
        if parts:
            messagebox.showinfo("Resumed programs", "\n\n".join(parts))

    def _reader_thread(self, action_id: str, process: subprocess.Popen) -> None:
        assert process.stdout is not None
        for line in process.stdout:
            self.output_queue.put((action_id, "line", line))

    def _waiter_thread(self, action_id: str, process: subprocess.Popen) -> None:
        code = process.wait()
        self.output_queue.put((action_id, "exit", code))

    def _stop_selected(self) -> None:
        if not self.current_action_id:
            return
        info = self.processes.get(self.current_action_id)
        action = self.actions_by_id[self.current_action_id]
        if not info:
            if action.id in self.external_running:
                self._stop_external_interactive(action)
            return
        if not messagebox.askyesno("Confirm stop", f"Stop {action.title}?"):
            return
        self.stop_process(action.id)

    def _stop_external_interactive(self, action: LauncherAction) -> None:
        """Offer to force-stop a tool that is running outside the launcher."""
        pids = self.external_pids.get(action.id, [])
        if not pids:
            messagebox.showinfo(
                "Running outside launcher",
                f"{action.title} is running outside this launcher, but I couldn't read "
                "its process id to stop it.\n\nClick Refresh Status, or end it in Task "
                "Manager.",
            )
            return
        pid_text = ", ".join(str(p) for p in pids)
        if not messagebox.askyesno(
            "Stop external process",
            f"{action.title} is running outside this launcher (PID {pid_text}).\n\n"
            "Force-stop it now? It wasn't started by the launcher, so this is a hard "
            "stop (like ending it in Task Manager).",
        ):
            return
        killed = self.stop_external(action.id)
        if killed:
            self._debug(f"force-stopped external {action.id} ({killed} pid(s))")
            messagebox.showinfo("Stopped", f"Stopped {action.title} (was running outside the launcher).")
        else:
            messagebox.showwarning(
                "Could not stop",
                f"Could not stop {action.title}. It may need administrator rights, or it "
                "already exited. Try Refresh Status.",
            )
        self._refresh_external_status(schedule=False)

    def _kill_pid(self, pid: int) -> bool:
        """Force-stop a process by PID (taskkill on Windows, SIGTERM elsewhere)."""
        try:
            if os.name == "nt":
                result = subprocess.run(
                    ["taskkill", "/PID", str(pid), "/T", "/F"],
                    capture_output=True,
                    text=True,
                    timeout=10,
                    **hidden_console_kwargs(),
                )
                return result.returncode == 0
            os.kill(pid, signal.SIGTERM)
            return True
        except Exception as exc:
            self._debug(f"kill pid {pid} failed: {exc!r}")
            return False

    def stop_external(self, action_id: str) -> int:
        """Force-stop the external process(es) detected for an action. Returns count killed."""
        killed = 0
        for pid in self.external_pids.get(action_id, []):
            if self._kill_pid(pid):
                killed += 1
        if killed:
            self.external_pids.pop(action_id, None)
            self.external_running.discard(action_id)
        return killed

    def stop_any(self, action_id: str, *, force: bool = False) -> bool:
        """Stop a tool whether the launcher started it or it is running externally."""
        if self.stop_process(action_id, force=force):
            return True
        return self.stop_external(action_id) > 0

    def stop_process(self, action_id: str, *, force: bool = False) -> bool:
        """Stop a launcher-started process. Returns True if a stop was issued.

        For an action that supports a clean shutdown (``graceful_stop``, e.g.
        watch.py) the FIRST stop drops a stop-flag file the script watches for —
        so it finishes the current poll, saves state and publishes logs before
        exiting — with a generous backstop force-kill. A second Stop (or
        ``force=True``, used by the Git Update flow) kills it now. Other tools
        are terminated immediately as before.
        """
        info = self.processes.get(action_id)
        if not info or info.process.poll() is not None:
            return False
        proc = info.process
        action = self.actions_by_id.get(action_id)
        graceful = bool(action and action.graceful_stop)
        second_press = action_id in self._stop_requested
        self._stop_requested.add(action_id)  # so the exit isn't flagged as a failure

        if graceful and not force and not second_press:
            stop_signal.request_stop(proc.pid)
            try:  # also nudge via console signal in case we do have a console
                proc.send_signal(signal.CTRL_BREAK_EVENT if os.name == "nt" else signal.SIGINT)
            except Exception:
                pass
            self.logs.setdefault(action_id, []).append(
                "[launcher] Stop requested — finishing the current poll, then saving "
                "state and publishing logs before exiting. Click Stop again to force-quit.\n"
            )
            self._schedule_force_kill(proc, action_id, GRACEFUL_STOP_SECONDS)
            return True

        # Force / non-graceful / second press: stop now.
        try:
            proc.terminate()  # TerminateProcess on Windows — reliable without a console
            self.logs.setdefault(action_id, []).append("[launcher] Stopping now.\n")
        except Exception:
            proc.terminate()
        self._schedule_force_kill(proc, action_id, 8)
        return True

    def _schedule_force_kill(self, proc: subprocess.Popen, action_id: str, delay: float) -> None:
        def force_kill() -> None:
            time.sleep(delay)
            if proc.poll() is None:
                try:
                    proc.kill()
                    self.output_queue.put((action_id, "line", "[launcher] Force killed after stop timeout.\n"))
                except OSError:
                    pass

        threading.Thread(target=force_kill, daemon=True).start()

    def _send_enter(self) -> None:
        if not self.current_action_id:
            return
        info = self.processes.get(self.current_action_id)
        if not info or not info.process.stdin:
            return
        try:
            info.process.stdin.write("\n")
            info.process.stdin.flush()
            self.output_queue.put((self.current_action_id, "line", "[launcher] Sent Enter.\n"))
        except OSError as exc:
            messagebox.showwarning("Could not send input", str(exc))

    def _clear_output(self) -> None:
        if self.current_action_id:
            self.logs[self.current_action_id] = []
            self._render_output(self.current_action_id)

    def _open_git_update(self) -> None:
        if self._git_dialog is not None and self._git_dialog.winfo_exists():
            self._git_dialog.deiconify()
            self._git_dialog.lift()
            self._git_dialog.focus_set()
            return
        self._git_dialog = GitUpdateDialog(self)

    def _open_logs_folder(self) -> None:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        try:
            if os.name == "nt":
                os.startfile(LOG_DIR)  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", str(LOG_DIR)])
        except OSError as exc:
            messagebox.showwarning("Could not open logs folder", str(exc))

    def _drain_output(self) -> None:
        while True:
            try:
                action_id, kind, payload = self.output_queue.get_nowait()
            except queue.Empty:
                break
            if kind == "line":
                line = str(payload)
                self.logs.setdefault(action_id, []).append(line)
                info = self.processes.get(action_id)
                if info:
                    info.log_file.write(line)
                    info.log_file.flush()
                if self.current_action_id == action_id:
                    self._append_output(line)
            elif kind == "exit":
                code = int(payload)
                self.external_running.discard(action_id)
                info = self.processes.pop(action_id, None)
                if info is not None:
                    stop_signal.clear_stop(info.process.pid)  # tidy up any stop-flag
                requested_stop = action_id in self._stop_requested
                self._stop_requested.discard(action_id)
                if requested_stop:
                    # We asked it to stop, so a non-zero code is expected — show
                    # it as stopped, not a failure.
                    self.stopped_actions.add(action_id)
                    self.last_exit.pop(action_id, None)
                    line = f"\n[launcher] Stopped (exit code {code}).\n"
                else:
                    self.stopped_actions.discard(action_id)
                    self.last_exit[action_id] = code
                    line = f"\n[launcher] Exit code: {code}\n"
                if info:
                    info.log_file.write(line)
                    info.log_file.close()
                self.logs.setdefault(action_id, []).append(line)
                if self.current_action_id == action_id:
                    self._append_output(line)
                self._update_all_statuses()
            elif kind == "scan_result":
                self._apply_scan_result(*payload)
            elif kind == "publish_result":
                self._handle_publish_result(*payload)
        self.after(100, self._drain_output)

    def _handle_publish_result(self, ok: bool, detail: str) -> None:
        self.report_button.configure(state="normal", text="Publish Debug Report")
        if ok:
            self._debug(f"publish ok: {detail}")
            messagebox.showinfo(
                "Debug report published",
                f"Pushed the debug report to the '{DEBUG_BRANCH}' branch.\n\n{detail}\n\n"
                "It is ready to be reviewed.",
            )
        else:
            self._debug(f"publish failed: {detail}")
            messagebox.showwarning(
                "Could not publish report",
                f"The report was saved locally to:\n{LOCAL_REPORT}\n\n"
                f"but pushing to '{DEBUG_BRANCH}' failed:\n{detail}\n\n"
                "Check your network/Git sign-in and click Publish Debug Report again.",
            )

    def _render_output(self, action_id: str) -> None:
        self.output.configure(state="normal")
        self.output.delete("1.0", tk.END)
        self.output.insert(tk.END, "".join(self.logs.get(action_id, [])))
        self.output.configure(state="disabled")
        self.output.see(tk.END)

    def _append_output(self, text: str) -> None:
        self.output.configure(state="normal")
        self.output.insert(tk.END, text)
        self.output.configure(state="disabled")
        self.output.see(tk.END)

    def _set_text(self, widget: tk.Text, text: str) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", tk.END)
        widget.insert(tk.END, text)
        widget.configure(state="disabled")

    def _is_running(self, action_id: str) -> bool:
        info = self.processes.get(action_id)
        return bool(info and info.process.poll() is None)

    def _debug(self, message: str) -> None:
        """Append one timestamped line to launcher_logs/launcher_debug.log (best effort)."""
        try:
            LOG_DIR.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with DEBUG_LOG.open("a", encoding="utf-8", errors="replace") as fh:
                fh.write(f"{stamp} {message}\n")
        except OSError:
            pass

    def _scan_processes(self) -> tuple[list[tuple[int | None, str]], str | None, str | None]:
        """List running ``(pid, command_line)`` pairs for external-status detection.

        Returns (pairs, method_used, error). On Windows it tries ``wmic`` first
        (legacy, removed on newer builds) then a PowerShell CIM query (the
        supported modern replacement); off-Windows it uses ``ps``. The PID is
        captured so external copies can be stopped. The error is returned (not
        swallowed) so the diagnostic can show why a scan came back empty.
        """
        if os.name == "nt":
            ps_query = (
                "Get-CimInstance Win32_Process | "
                "ForEach-Object { \"$($_.ProcessId) $($_.CommandLine)\" }"
            )
            candidates = [
                ("wmic", ["wmic", "path", "win32_process", "get", "ProcessId,CommandLine"]),
                ("powershell-cim", ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_query]),
            ]
        else:
            candidates = [("ps", ["ps", "-eo", "pid=,args="])]

        last_error: str | None = None
        for name, argv in candidates:
            try:
                output = subprocess.check_output(
                    argv,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    stderr=subprocess.DEVNULL,
                    timeout=8,
                    **hidden_console_kwargs(),
                )
            except Exception as exc:
                last_error = f"{name}: {exc.__class__.__name__}: {exc}"
                continue
            return procscan.parse_scan_output(output.splitlines(), name), name, None
        return [], None, last_error or "no process-listing tool available"

    def _compute_exclusion_args(self) -> dict[str, set[str]]:
        """Args that disqualify a command line from matching an action.

        Some actions share a script and differ only by default args (e.g.
        the all-day ``watch`` vs. the one-shot ``watch.py --once``). When
        scanning external processes we match by script name, so without this
        a ``--once`` run would light up the long-running ``watch`` action.
        For each action we collect the default args used by *other* actions
        on the same script but not by this one.
        """
        by_script: dict[str, list[LauncherAction]] = {}
        for action in self.actions:
            if action.script:
                by_script.setdefault(Path(action.script).name.lower(), []).append(action)
        exclusion: dict[str, set[str]] = {}
        for action in self.actions:
            if not action.script:
                continue
            own = set(action.default_args)
            siblings = by_script.get(Path(action.script).name.lower(), [])
            others = {arg for sib in siblings if sib.id != action.id for arg in sib.default_args}
            exclusion[action.id] = {a.lower() for a in others - own}
        return exclusion

    def _refresh_external_status(self, schedule: bool = True, verbose: bool = False) -> None:
        # The process scan shells out to PowerShell/wmic/ps, which can be slow or
        # even hang. Run it on a worker thread so it can NEVER freeze the UI; the
        # result is applied back on the main thread via the output queue. Skip a
        # background poll while one is already in flight (a manual/verbose request
        # always runs).
        if verbose or not self._scan_in_progress:
            if not verbose:
                self._scan_in_progress = True
            threading.Thread(target=self._scan_worker, args=(verbose,), daemon=True).start()
        if schedule:
            # PowerShell starts slower than wmic did, so poll a little less
            # often; per-second elapsed time is handled by _tick_status.
            self.after(8000, self._refresh_external_status)

    def _scan_worker(self, verbose: bool) -> None:
        try:
            pairs, method, error = self._scan_processes()
        except Exception as exc:  # pragma: no cover - defensive
            pairs, method, error = [], None, f"scan crashed: {exc!r}"
        self.output_queue.put(("__scan__", "scan_result", (pairs, method, error, verbose)))

    def _scan_now_sync(self) -> None:
        """Scan and apply synchronously, for the rare one-off checks (Run/Pull)
        that must see fresh external state before deciding. Infrequent and
        user-initiated, so a brief pause here is acceptable — unlike the
        recurring background poll, which must never block the UI."""
        try:
            pairs, method, error = self._scan_processes()
        except Exception as exc:  # pragma: no cover - defensive
            pairs, method, error = [], None, f"scan crashed: {exc!r}"
        self._apply_scan_result(pairs, method, error, verbose=False)

    def _apply_scan_result(
        self,
        pairs: list[tuple[int | None, str]],
        method: str | None,
        error: str | None,
        verbose: bool,
    ) -> None:
        self._scan_in_progress = False
        cmd_lines = [cmd for _pid, cmd in pairs]
        running: set[str] = set()
        pids_by_action: dict[str, list[int]] = {}
        for action in self.actions:
            if not action.long_running or not action.script or self._is_running(action.id):
                continue
            # Skip shared-script siblings (e.g. watch.py --once vs the long-running watch).
            exclude = self._exclusion_args.get(action.id, set())
            matched = [
                (pid, cmd)
                for (pid, cmd) in pairs
                if procscan.process_line_matches_script(cmd, action.script)
                and not any(token in cmd.lower() for token in exclude)
            ]
            if matched:
                running.add(action.id)
                pids = [pid for pid, _cmd in matched if pid is not None]
                if pids:
                    pids_by_action[action.id] = pids

        # Log on the manual diagnostic, or whenever the outcome changes — so the
        # background poll does not spam the debug log.
        key = (method, error, tuple(sorted(running)))
        if verbose or key != self._last_status_key:
            self._debug(
                f"status scan: method={method or 'none'} error={error or 'none'} "
                f"python_lines={len(pairs)} detected={sorted(running) or 'none'}"
            )
            self._last_status_key = key
        if verbose:
            for pid, cmd in pairs[:25]:
                self._debug(f"  proc: pid={pid} {cmd[:300]}")

        self.external_running = running
        self.external_pids = pids_by_action
        self._update_all_statuses()
        if verbose:
            self._show_status_diagnostic(method, error, cmd_lines, running)

    def _show_status_diagnostic(
        self, method: str | None, error: str | None, process_lines: list[str], running: set[str]
    ) -> None:
        if method:
            head = f"Process scan used: {method}\nPython processes seen: {len(process_lines)}"
        else:
            head = (
                "Process scan FAILED — no method worked.\n"
                "On newer Windows 'wmic' is removed; the PowerShell fallback also "
                "did not run."
            )
        if error:
            head += f"\nLast error: {error}"
        if running:
            names = ", ".join(self.actions_by_id[a].title for a in sorted(running))
            body = f"\n\nDetected running outside the launcher:\n{names}"
        else:
            body = "\n\nNo long-running tools detected running outside the launcher."
        body += f"\n\nFull details were written to:\n{DEBUG_LOG}"
        messagebox.showinfo("Status diagnostic", head + body)

    def _tail_debug_log(self, count: int) -> list[str]:
        try:
            with DEBUG_LOG.open("r", encoding="utf-8", errors="replace") as fh:
                tail = fh.readlines()[-count:]
        except OSError:
            return ["  (no debug log yet)"]
        return [f"  {line.rstrip()}" for line in tail] or ["  (debug log empty)"]

    def _build_debug_report_text(self) -> str:
        """Build the shareable debug snapshot string (no I/O)."""
        pairs, method, error = self._scan_processes()
        process_lines = [f"pid={pid} {cmd}" for pid, cmd in pairs]
        detected = sorted(self.external_running)

        branch = git_update.current_branch() or "?"
        commit = git_update.head_rev()[:10] or "?"

        out: list[str] = []
        out.append("# Launcher debug report")
        out.append(f"generated: {datetime.now().isoformat(timespec='seconds')}")
        out.append(f"code version: branch {branch} @ commit {commit}")
        out.append(f"os.name: {os.name}   platform: {sys.platform}")
        out.append(f"python (launcher): {sys.executable}")
        out.append(f"python (for scripts): {self.python_path}")
        out.append("")
        out.append("## Process scan (external-status detection)")
        out.append(f"method used: {method or 'NONE - all scanners failed'}")
        out.append(f"error: {error or 'none'}")
        out.append(f"python process lines seen: {len(process_lines)}")
        out.append(f"detected running outside launcher: {detected or 'none'}")
        out.append("")
        out.append("### python process command lines (up to 40)")
        out.extend(f"  {line.strip()[:400]}" for line in process_lines[:40])
        if not process_lines:
            out.append("  (none seen — scanner returned nothing)")
        out.append("")
        out.append("## Launcher-started processes")
        started_any = False
        for action_id, info in self.processes.items():
            started_any = True
            alive = info.process.poll() is None
            out.append(
                f"  {action_id}: pid={info.process.pid} alive={alive} "
                f"started={info.started_at:%Y-%m-%d %H:%M:%S}"
            )
            out.append(f"      cmd: {self._command_text(info.command)}")
        if not started_any:
            out.append("  (none)")
        out.append("")
        out.append("## Last exit codes")
        if self.last_exit:
            out.extend(f"  {action_id}: {code}" for action_id, code in sorted(self.last_exit.items()))
        else:
            out.append("  (none yet)")
        out.append("")
        out.append("## Tail of launcher_debug.log (last 120 lines)")
        out.extend(self._tail_debug_log(120))
        out.append("")
        return "\n".join(out) + "\n"

    def _publish_debug_report(self) -> None:
        """Save the debug snapshot locally and push it to the debug branch.

        The per-run launcher_logs/ are git-ignored and never leave this PC, so
        the report is published to the dedicated ``debug/launcher`` branch where
        it can be reviewed. Pushing runs on a worker thread (git plumbing that
        does not touch the working checkout).
        """
        text = self._build_debug_report_text()
        try:
            LOG_DIR.mkdir(parents=True, exist_ok=True)
            LOCAL_REPORT.write_text(text, encoding="utf-8")
        except OSError as exc:
            messagebox.showerror("Could not write report", str(exc))
            return
        self._debug(f"publishing debug report to branch {DEBUG_BRANCH}")
        self.report_button.configure(state="disabled", text="Publishing…")
        threading.Thread(target=self._publish_worker, args=(text,), daemon=True).start()

    def _publish_worker(self, text: str) -> None:
        try:
            ok, detail = git_update.publish_report(text, branch=DEBUG_BRANCH)
        except Exception as exc:  # pragma: no cover - defensive
            ok, detail = False, f"{exc.__class__.__name__}: {exc}"
        self.output_queue.put(("__publish__", "publish_result", (ok, detail)))

    def _tick_status(self) -> None:
        self._update_detail_status()
        self.after(1000, self._tick_status)

    def _update_all_statuses(self) -> None:
        self._populate_lists()
        self._update_detail_status()

    def _update_detail_status(self) -> None:
        action_id = self.current_action_id
        self.status_light.delete("all")
        color = "#9ca3af"
        status = "Idle"
        self.run_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        self.enter_button.configure(state="disabled")
        if action_id:
            if self._is_running(action_id):
                color = "#16a34a"
                info = self.processes[action_id]
                elapsed = int((datetime.now() - info.started_at).total_seconds())
                if action_id in self._stop_requested:
                    color = "#d97706"  # amber: stop requested, finishing up
                    status = f"Stopping… ({elapsed}s)"
                else:
                    status = f"Running ({elapsed}s)"
                self.run_button.configure(state="disabled")
                self.stop_button.configure(state="normal")
                self.enter_button.configure(state="normal")
            elif action_id in self.external_running:
                color = "#16a34a"
                status = "Running outside launcher"
                self.run_button.configure(state="disabled")
            elif action_id in self.stopped_actions:
                color = "#9ca3af"
                status = "Stopped"
            elif action_id in self.last_exit:
                code = self.last_exit[action_id]
                color = "#2563eb" if code == 0 else "#dc2626"
                status = f"Last exit: {code}"
        self.status_light.create_oval(2, 2, 16, 16, fill=color, outline=color)
        self.status_var.set(status)

    def _on_close(self) -> None:
        if self.processes:
            names = ", ".join(self.actions_by_id[aid].title for aid in self.processes)
            if not messagebox.askyesno(
                "Processes still running",
                f"These launcher-started processes are still running:\n\n{names}\n\nClose the launcher anyway?",
            ):
                return
        self._save_state()
        self._publish_on_close()
        self.destroy()

    def _publish_on_close(self) -> None:
        """Publish a final debug report as the launcher shuts down.

        Runs on a NON-daemon thread so the window can close immediately while
        the push still finishes — the interpreter waits for this thread before
        exiting, and publish_report has its own timeouts so it cannot hang
        shutdown indefinitely. Building the report is done in the worker too so
        the close feels instant. Best effort: failures only go to the debug log.
        """
        self._debug(f"close: publishing debug report to {DEBUG_BRANCH}")
        threading.Thread(target=self._publish_on_close_worker, daemon=False).start()

    def _publish_on_close_worker(self) -> None:
        try:
            text = self._build_debug_report_text()
        except Exception as exc:  # pragma: no cover - defensive
            self._debug(f"close publish: could not build report: {exc!r}")
            return
        try:
            LOG_DIR.mkdir(parents=True, exist_ok=True)
            LOCAL_REPORT.write_text(text, encoding="utf-8")
        except OSError:
            pass
        try:
            ok, detail = git_update.publish_report(
                text,
                branch=DEBUG_BRANCH,
                message=f"Launcher debug report (on close) {datetime.now():%Y-%m-%d %H:%M:%S}",
                timeout=20,
            )
            self._debug(f"close publish {'ok' if ok else 'failed'}: {detail}")
        except Exception as exc:  # pragma: no cover - defensive
            self._debug(f"close publish error: {exc!r}")


class GitUpdateDialog(tk.Toplevel):
    """A small window to pull the latest code from a chosen Git branch.

    Branch discovery and the pull itself run on worker threads; results come
    back through a queue so the Tk UI stays responsive. The heavy lifting lives
    in the import-light ``git_update`` module.
    """

    def __init__(self, parent: "LauncherApp") -> None:
        super().__init__(parent)
        self.app = parent
        self.title("Git Update")
        self.geometry("760x540")
        self.minsize(620, 420)
        self.transient(parent)

        self._queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self._busy = False
        self._current_branch = ""
        self._stop_for_update: list[str] = []
        self.current_var = tk.StringVar(value="Current branch: …")
        self.branch_var = tk.StringVar()
        self.switch_var = tk.BooleanVar(value=True)

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._poll_queue)
        self.refresh_branches()

    def _build_ui(self) -> None:
        body = ttk.Frame(self, padding=12)
        body.pack(fill="both", expand=True)

        ttk.Label(body, textvariable=self.current_var, font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(
            body,
            text="Choose a branch and click Pull to update this folder's code from GitHub.",
            foreground="#555",
        ).pack(anchor="w", pady=(2, 8))

        pick = ttk.Frame(body)
        pick.pack(fill="x")
        ttk.Label(pick, text="Pull from branch:").pack(side="left")
        self.branch_combo = ttk.Combobox(pick, textvariable=self.branch_var, state="readonly", width=42)
        self.branch_combo.pack(side="left", padx=(8, 8))
        self.refresh_button = ttk.Button(pick, text="Refresh", command=lambda: self.refresh_branches(fetch=True))
        self.refresh_button.pack(side="left")
        ToolTip(self.refresh_button, "Fetch from origin and reload the branch list.")

        self.switch_check = ttk.Checkbutton(
            body,
            text="Switch to this branch (checkout) before pulling",
            variable=self.switch_var,
        )
        self.switch_check.pack(anchor="w", pady=(8, 8))
        ToolTip(
            self.switch_check,
            "On: end up on the chosen branch (checkout, then pull).\n"
            "Off: merge the chosen branch into the branch you are already on.",
        )

        actions = ttk.Frame(body)
        actions.pack(fill="x", pady=(0, 8))
        self.pull_button = ttk.Button(actions, text="Pull", style="Action.TButton", command=self.start_pull)
        self.pull_button.pack(side="left")
        self.close_button = ttk.Button(actions, text="Close", command=self._on_close)
        self.close_button.pack(side="right")

        out_frame = ttk.LabelFrame(body, text="Git output", padding=6)
        out_frame.pack(fill="both", expand=True)
        self.output = tk.Text(out_frame, wrap="word", relief="solid", borderwidth=1, height=14)
        self.output.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(out_frame, orient="vertical", command=self.output.yview)
        scroll.pack(side="right", fill="y")
        self.output.configure(yscrollcommand=scroll.set, state="disabled")

    def _log(self, text: str) -> None:
        self.output.configure(state="normal")
        self.output.insert(tk.END, text)
        self.output.configure(state="disabled")
        self.output.see(tk.END)

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        state = "disabled" if busy else "normal"
        self.pull_button.configure(state=state)
        self.refresh_button.configure(state=state)
        self.switch_check.configure(state=state)
        self.branch_combo.configure(state="disabled" if busy else "readonly")

    def refresh_branches(self, *, fetch: bool = False) -> None:
        if self._busy:
            return
        self._set_busy(True)
        if fetch:
            self._log("[git] Fetching and reloading branches…\n")
        threading.Thread(target=self._load_branches_worker, args=(fetch,), daemon=True).start()

    def _load_branches_worker(self, fetch: bool) -> None:
        try:
            if not git_update.git_available():
                self._queue.put(("error", "Git is not installed or not on PATH."))
                return
            if not git_update.is_git_repo():
                self._queue.put(("error", "This folder is not a Git repository."))
                return
            if fetch:
                result = git_update.run_git(["fetch", "--all", "--prune"], timeout=120)
                for stream in (result.stdout, result.stderr):
                    if stream.strip():
                        self._queue.put(("line", stream if stream.endswith("\n") else stream + "\n"))
            current = git_update.current_branch()
            branches = git_update.list_branches()
            self._queue.put(("branches", (current, branches)))
        except Exception as exc:  # pragma: no cover - defensive UI guard
            self._queue.put(("error", f"Could not read branches: {exc}"))

    def _apply_branches(self, current: str, branches: list[str]) -> None:
        self._current_branch = current
        self.current_var.set(f"Current branch: {current or '(unknown)'}")
        self.branch_combo.configure(values=branches)
        chosen = self.branch_var.get().strip()
        if chosen not in branches:
            if current in branches:
                self.branch_var.set(current)
            elif branches:
                self.branch_var.set(branches[0])
            else:
                self.branch_var.set("")

    def start_pull(self) -> None:
        if self._busy:
            return
        branch = self.branch_var.get().strip()
        if not branch:
            messagebox.showwarning("No branch", "Choose a branch to pull.", parent=self)
            return
        try:
            steps = git_update.build_pull_steps(branch, self._current_branch, switch=bool(self.switch_var.get()))
        except ValueError as exc:
            messagebox.showwarning("Cannot pull", str(exc), parent=self)
            return

        # Running programs keep executing the OLD code until restarted, and the
        # launcher itself only picks up changes when reopened — so offer to stop
        # everything now and have it restart when the launcher next opens.
        self._stop_for_update = []
        if not self._handle_live_processes():
            return

        preview = "\n".join("git " + " ".join(step) for step in steps)
        if not messagebox.askyesno("Confirm pull", f"Run these git commands?\n\n{preview}", parent=self):
            self._stop_for_update = []
            return

        if self._stop_for_update:
            # Remember (persistently) what to bring back, then stop them — whether
            # launcher-started or running externally.
            self.app.set_pending_restart(self._stop_for_update)
            stopped = ", ".join(self.app.actions_by_id[aid].title for aid in self._stop_for_update)
            self._log(f"[git] Stopping programs before update: {stopped}\n")
            for action_id in self._stop_for_update:
                self.app.stop_any(action_id, force=True)  # fast: they're restarted fresh after the pull

        self._set_busy(True)
        self._log(f"\n=== Pulling '{branch}' ===\n")
        threading.Thread(target=self._pull_worker, args=(steps,), daemon=True).start()

    def _handle_live_processes(self) -> bool:
        """Prompt about running programs. Returns False if the user cancels the pull.

        Every running program we can stop — launcher-started or external (we have
        a PID) — is offered for stopping; the chosen ones are recorded in
        ``self._stop_for_update`` (and persisted) to be stopped now and restarted
        the next time the launcher opens. External tools with no PID are
        warn-only.
        """
        self.app._scan_now_sync()
        stoppable = self.app.running_launcher_processes() + self.app.stoppable_external_processes()
        unstoppable = self.app.unstoppable_external_processes()

        if stoppable:
            names = ", ".join(title for _, title in stoppable)
            answer = messagebox.askyesnocancel(
                "Programs are running",
                f"These programs are running:\n\n{names}\n\n"
                "Stop them for the update? They will restart automatically when the "
                "launcher reopens.\n\n"
                "Yes — stop them now (they restart on reopen)\n"
                "No — update without stopping them\n"
                "Cancel — don't update",
                parent=self,
            )
            if answer is None:
                return False
            if answer:
                self._stop_for_update = [action_id for action_id, _ in stoppable]

        if unstoppable:
            ext_names = ", ".join(unstoppable)
            if not messagebox.askyesno(
                "Tools running outside the launcher",
                "These are running outside this launcher and I couldn't read a PID to "
                f"stop them:\n\n{ext_names}\n\n"
                "Update anyway? Stop and restart them yourself to pick up the new code.",
                parent=self,
            ):
                return False

        return True

    def _pull_worker(self, steps: list[list[str]]) -> None:
        before = git_update.head_rev()
        code = git_update.run_pull_steps(steps, lambda line: self._queue.put(("line", line)))
        after = git_update.head_rev()
        changed = git_update.changed_files(before, after)
        self._queue.put(("done", (code, changed)))

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == "line":
                    self._log(str(payload))
                elif kind == "branches":
                    current, branches = payload
                    self._apply_branches(current, branches)
                    self._set_busy(False)
                elif kind == "error":
                    self._log(f"[git] {payload}\n")
                    self._set_busy(False)
                    messagebox.showerror("Git", str(payload), parent=self)
                elif kind == "done":
                    code, changed = payload
                    if code == 0:
                        self._log("\n[git] Done. Your code is up to date.\n")
                    else:
                        self._log(f"\n[git] Finished with exit code {code}.\n")
                    self._set_busy(False)
                    self.refresh_branches()  # reflect any branch switch in the label
                    self._after_pull(code, changed)
        except queue.Empty:
            pass
        if self.winfo_exists():
            self.after(100, self._poll_queue)

    def _after_pull(self, code: int, changed: list[str]) -> None:
        """After a successful pull, relaunch the launcher so the new code takes effect.

        Programs that were stopped are persisted and resume in the fresh
        launcher. We only auto-relaunch when it is safe: a clean pull that
        actually changed something, with nothing the user chose to leave
        running (relaunching would orphan those).
        """
        if code != 0:
            if self._stop_for_update:
                messagebox.showwarning(
                    "Update did not complete",
                    "The update did not finish cleanly (see the output above).\n\n"
                    "The programs you stopped will restart when you reopen the launcher.",
                    parent=self,
                )
            return

        # Relaunch only when it actually matters: the launcher's own code changed,
        # or programs were stopped and need to come back. Other script changes are
        # picked up the next time those scripts run, so no relaunch is needed.
        if not (git_update.launcher_needs_restart(changed) or self._stop_for_update):
            if changed:
                self._log("[git] Update applied; scripts will use the new code next run.\n")
            else:
                self._log("[git] Already up to date; nothing to apply.\n")
            return

        # Programs the user chose to leave running would be orphaned by a relaunch.
        leftover = self.app.running_launcher_processes(exclude=self._stop_for_update)
        if leftover:
            names = ", ".join(title for _, title in leftover)
            self._log(f"[git] Not auto-relaunching — still running on old code: {names}\n")
            messagebox.showinfo(
                "Reopen when ready",
                "The update was applied, but these programs are still running on the "
                f"old code:\n\n{names}\n\nReopen the launcher yourself once they finish "
                "to pick up the new version.",
                parent=self,
            )
            return

        self._log("[git] Update complete — relaunching the launcher to apply it…\n")
        self._relaunch_when_clear(attempts=0)

    def _relaunch_when_clear(self, attempts: int) -> None:
        """Wait for the stopped programs to fully exit, then relaunch the launcher."""
        still_running = [aid for aid in self._stop_for_update if self.app._is_running(aid)]
        if still_running and attempts < 40:  # up to ~10s at 250ms
            self.after(250, lambda: self._relaunch_when_clear(attempts + 1))
            return
        self.app.relaunch_self_and_exit()

    def _on_close(self) -> None:
        if self._busy and not messagebox.askyesno(
            "Git running",
            "A git operation is still running. Close this window anyway?",
            parent=self,
        ):
            return
        self.destroy()


if __name__ == "__main__":
    app = LauncherApp()
    app.mainloop()
