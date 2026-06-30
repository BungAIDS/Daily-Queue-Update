"""Desktop launcher for the Daily Queue helper scripts.

This is intentionally standard-library only. It gives the Windows workstation a
single friendly place to start/stop the daily queue tools, see live output, and
remember the last options used without adding another package to install.
"""
from __future__ import annotations

import json
import os
import queue
import re
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


ROOT = Path(__file__).resolve().parent
STATE_PATH = ROOT / ".launcher_state.json"
LOG_DIR = ROOT / "launcher_logs"
DEBUG_LOG = LOG_DIR / "launcher_debug.log"
# Tracked (committed) folder for shareable debug snapshots. launcher_logs/ is
# git-ignored and lives only on the workstation, so "Export Debug Report" writes
# here instead, where it can be pushed and reviewed later.
DIAG_DIR = ROOT / "diagnostics"
DIAG_REPORT = DIAG_DIR / "launcher_report.txt"


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


def process_line_matches_script(line: str, script: str) -> bool:
    script_name = re.escape(Path(script).name.lower())
    pattern = rf'(?:^|[\s"\'/\\]){script_name}(?:$|[\s"\'])'
    return bool(re.search(pattern, line.lower()))


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
        self.processes: dict[str, ProcessInfo] = {}
        self.external_running: set[str] = set()
        self.last_exit: dict[str, int] = {}
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

        self.python_path = self._find_python()
        self.allow_send_var = tk.BooleanVar(value=False)

        self._debug(f"launcher started (os={os.name}, python={self.python_path})")
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._select_initial_action()
        self._drain_output()
        self._refresh_external_status()
        self._tick_status()

    def _load_state(self) -> dict[str, Any]:
        if not STATE_PATH.exists():
            return {"options": {}, "last_action": ""}
        try:
            return json.loads(STATE_PATH.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return {"options": {}, "last_action": ""}

    def _save_state(self) -> None:
        data = {"options": {}, "last_action": self.current_action_id or ""}
        for action_id, vars_by_key in self.option_vars.items():
            data["options"][action_id] = {k: v.get() for k, v in vars_by_key.items()}
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
        self.report_button = ttk.Button(button_row, text="Export Debug Report", command=self._export_debug_report)
        self.report_button.pack(side="left", padx=(8, 0))
        ToolTip(self.report_button, "Write diagnostics/launcher_report.txt — a shareable snapshot to commit/push when debugging the launcher.")
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

    def _build_command(self, action: LauncherAction) -> list[str]:
        script = action.script
        if action.script_option:
            raw_script = str(self.option_vars[action.id][action.script_option].get()).strip()
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
        vars_by_key = self.option_vars.get(action.id, {})
        for spec in action.options:
            if spec.key == action.script_option:
                continue
            var = vars_by_key.get(spec.key)
            if var is None:
                continue
            if spec.kind == "check":
                if bool(var.get()) and spec.arg:
                    command.append(spec.arg)
                continue
            raw = str(var.get()).strip()
            if not raw:
                continue
            pieces = shlex.split(raw, posix=False) if spec.split else [raw]
            if spec.arg:
                command.append(spec.arg)
            command.extend(pieces if spec.positional or spec.arg else pieces)
        return command

    def _command_text(self, command: Iterable[str]) -> str:
        return " ".join(shlex.quote(str(part)) for part in command)

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
            self._refresh_external_status(schedule=False)
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

        self.processes[action.id] = ProcessInfo(action.id, process, log_file, log_path, command)
        self.logs[action.id] = [f"$ {self._command_text(command)}\n", f"[launcher] Log: {log_path}\n"]
        if self.current_action_id == action.id:
            self._render_output(action.id)
        self._update_all_statuses()
        threading.Thread(target=self._reader_thread, args=(action.id, process), daemon=True).start()
        threading.Thread(target=self._waiter_thread, args=(action.id, process), daemon=True).start()

    def restart_action(self, action_id: str) -> None:
        """Rebuild an action's command from its saved options and start it.

        Used by the Git Update window to bring a live tool (e.g. watch.py)
        back up on the freshly pulled code. Raises on failure so the caller
        can report it.
        """
        action = self.actions_by_id[action_id]
        command = self._build_command(action)
        self._start_process(action, command)

    def live_launcher_processes(self) -> list[tuple[str, str]]:
        """(id, title) for long-running tools this launcher started and that are still alive."""
        live: list[tuple[str, str]] = []
        for action_id, info in list(self.processes.items()):
            action = self.actions_by_id.get(action_id)
            if action and action.long_running and info.process.poll() is None:
                live.append((action_id, action.title))
        return live

    def live_external_processes(self) -> list[str]:
        """Titles of long-running tools running outside the launcher (cannot be managed here)."""
        return [self.actions_by_id[aid].title for aid in self.external_running if aid in self.actions_by_id]

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
                messagebox.showinfo(
                    "Running outside launcher",
                    "This appears to be running outside the launcher, so I can show it as running but cannot safely stop it here.",
                )
            return
        if not messagebox.askyesno("Confirm stop", f"Stop {action.title}?"):
            return
        self.stop_process(action.id)

    def stop_process(self, action_id: str) -> bool:
        """Ask a launcher-started process to stop (no dialogs). Returns True if a stop was issued.

        Sends CTRL_BREAK on Windows / terminate elsewhere, then force-kills
        after a timeout if it is still alive. Shared by the Stop button and the
        Git Update auto-stop.
        """
        info = self.processes.get(action_id)
        if not info or info.process.poll() is not None:
            return False
        proc = info.process
        try:
            if os.name == "nt":
                proc.send_signal(signal.CTRL_BREAK_EVENT)
            else:
                proc.terminate()
            self.logs.setdefault(action_id, []).append("[launcher] Stop requested.\n")
        except Exception:
            proc.terminate()

        def force_kill() -> None:
            time.sleep(8)
            if proc.poll() is None:
                try:
                    proc.kill()
                    self.output_queue.put((action_id, "line", "[launcher] Force killed after stop timeout.\n"))
                except OSError:
                    pass

        threading.Thread(target=force_kill, daemon=True).start()
        return True

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
                self.last_exit[action_id] = code
                self.external_running.discard(action_id)
                info = self.processes.pop(action_id, None)
                if info:
                    info.log_file.write(f"\n[launcher] Exit code: {code}\n")
                    info.log_file.close()
                line = f"\n[launcher] Exit code: {code}\n"
                self.logs.setdefault(action_id, []).append(line)
                if self.current_action_id == action_id:
                    self._append_output(line)
                self._update_all_statuses()
        self.after(100, self._drain_output)

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

    def _scan_processes(self) -> tuple[list[str], str | None, str | None]:
        """List running command lines for external-status detection.

        Returns (python_command_lines, method_used, error). On Windows it tries
        ``wmic`` first (legacy, removed on newer builds) and falls back to a
        PowerShell CIM query, which is the supported modern replacement. The
        error is returned (not swallowed) so the diagnostic can show why a scan
        came back empty.
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
            candidates = [("ps", ["ps", "-eo", "args"])]

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
            lines = [line for line in output.splitlines() if "python" in line.lower()]
            return lines, name, None
        return [], None, last_error or "no process-listing tool available"

    def _refresh_external_status(self, schedule: bool = True, verbose: bool = False) -> None:
        lines, method, error = self._scan_processes()
        process_lines = [line.lower() for line in lines]
        running: set[str] = set()
        for action in self.actions:
            if not action.long_running or not action.script or self._is_running(action.id):
                continue
            if any(process_line_matches_script(line, action.script) for line in process_lines):
                running.add(action.id)

        # Log on the manual diagnostic, or whenever the outcome changes — so the
        # background poll every 5s does not spam the debug log.
        key = (method, error, tuple(sorted(running)))
        if verbose or key != self._last_status_key:
            self._debug(
                f"status scan: method={method or 'none'} error={error or 'none'} "
                f"python_lines={len(process_lines)} detected={sorted(running) or 'none'}"
            )
            self._last_status_key = key
        if verbose:
            for line in process_lines[:25]:
                self._debug(f"  proc: {line.strip()[:300]}")

        self.external_running = running
        self._update_all_statuses()
        if verbose:
            self._show_status_diagnostic(method, error, process_lines, running)
        if schedule:
            self.after(5000, self._refresh_external_status)

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

    def _export_debug_report(self) -> None:
        """Write a shareable debug snapshot to diagnostics/launcher_report.txt.

        The diagnostics folder is tracked in git, so this report can be
        committed and pushed, then read later when diagnosing launcher issues
        (the per-run launcher_logs/ are git-ignored and never leave the PC).
        """
        lines, method, error = self._scan_processes()
        process_lines = [line for line in lines]
        detected = sorted(self.external_running)

        out: list[str] = []
        out.append("# Launcher debug report")
        out.append(f"generated: {datetime.now().isoformat(timespec='seconds')}")
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

        try:
            DIAG_DIR.mkdir(parents=True, exist_ok=True)
            DIAG_REPORT.write_text("\n".join(out) + "\n", encoding="utf-8")
        except OSError as exc:
            messagebox.showerror("Could not write report", str(exc))
            return
        self._debug(f"exported debug report to {DIAG_REPORT}")
        messagebox.showinfo(
            "Debug report exported",
            f"Wrote:\n{DIAG_REPORT}\n\n"
            "This file is tracked in the repo. Commit and push it (to the branch "
            "we are working on) so it can be reviewed when debugging the launcher.\n\n"
            "It includes process command lines and file paths — glance over it "
            "before sharing.",
        )

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
                status = f"Running ({elapsed}s)"
                self.run_button.configure(state="disabled")
                self.stop_button.configure(state="normal")
                self.enter_button.configure(state="normal")
            elif action_id in self.external_running:
                color = "#16a34a"
                status = "Running outside launcher"
                self.run_button.configure(state="disabled")
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
        self.destroy()


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
        self._restart_after_pull: list[str] = []
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

        # A running tool keeps executing the *old* code until it is restarted,
        # so offer to stop live tools before the pull and bring them back after.
        self._restart_after_pull = []
        if not self._handle_live_processes():
            return

        preview = "\n".join("git " + " ".join(step) for step in steps)
        if not messagebox.askyesno("Confirm pull", f"Run these git commands?\n\n{preview}", parent=self):
            self._restart_after_pull = []
            return

        if self._restart_after_pull:
            stopped = ", ".join(self.app.actions_by_id[aid].title for aid in self._restart_after_pull)
            self._log(f"[git] Stopping live tools before update: {stopped}\n")
            for action_id in self._restart_after_pull:
                self.app.stop_process(action_id)

        self._set_busy(True)
        self._log(f"\n=== Pulling '{branch}' ===\n")
        threading.Thread(target=self._pull_worker, args=(steps,), daemon=True).start()

    def _handle_live_processes(self) -> bool:
        """Prompt about running live tools. Returns False if the user cancels the pull.

        When the user agrees, the launcher-started tools are recorded in
        ``self._restart_after_pull`` to be stopped now and restarted after the
        pull succeeds.
        """
        self.app._refresh_external_status(schedule=False)
        managed = self.app.live_launcher_processes()
        external = self.app.live_external_processes()

        if managed:
            names = ", ".join(title for _, title in managed)
            answer = messagebox.askyesnocancel(
                "Live tools are running",
                "These live tools are running and will keep using the OLD code until "
                f"they are restarted:\n\n{names}\n\n"
                "Stop them now, run the update, then restart them automatically?\n\n"
                "Yes — stop, update, then restart them\n"
                "No — update without touching them (restart them yourself later)\n"
                "Cancel — don't update",
                parent=self,
            )
            if answer is None:
                return False
            if answer:
                self._restart_after_pull = [action_id for action_id, _ in managed]

        if external:
            ext_names = ", ".join(external)
            if not messagebox.askyesno(
                "Tools running outside the launcher",
                "These are running outside this launcher, so it cannot stop or restart "
                f"them for you:\n\n{ext_names}\n\n"
                "Update anyway? Restart them yourself afterward to pick up the new code.",
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
                    self._after_pull(code)
                    if code == 0 and git_update.launcher_needs_restart(changed):
                        self._log(
                            "\n[git] NOTE: this update changed the launcher itself. "
                            "Close and reopen the launcher to run the new version.\n"
                        )
                        messagebox.showinfo(
                            "Restart the launcher",
                            "This update changed the launcher program itself.\n\n"
                            "Close and reopen the launcher (RunLauncher) so it runs the "
                            "new version — the running window is still on the old code.",
                            parent=self,
                        )
        except queue.Empty:
            pass
        if self.winfo_exists():
            self.after(100, self._poll_queue)

    def _after_pull(self, code: int) -> None:
        """Restart any tools we stopped for the update (only if the pull succeeded)."""
        pending = self._restart_after_pull
        self._restart_after_pull = []
        if not pending:
            return
        if code != 0:
            titles = ", ".join(self.app.actions_by_id[aid].title for aid in pending)
            self._log(
                f"[git] Update failed (exit {code}); leaving stopped tools off. "
                f"Restart when ready: {titles}\n"
            )
            return
        self._wait_then_restart(pending, attempts=0)

    def _wait_then_restart(self, pending: list[str], attempts: int) -> None:
        """Wait (without freezing the UI) for stopped tools to exit, then relaunch them."""
        still_running = [aid for aid in pending if self.app._is_running(aid)]
        if still_running and attempts < 40:  # up to ~10s at 250ms
            self.after(250, lambda: self._wait_then_restart(pending, attempts + 1))
            return
        for action_id in pending:
            title = self.app.actions_by_id[action_id].title
            if self.app._is_running(action_id):
                self._log(f"[git] {title} did not stop in time; not restarting it.\n")
                continue
            try:
                self.app.restart_action(action_id)
                self._log(f"[git] Restarted: {title}\n")
            except Exception as exc:
                self._log(f"[git] Could not restart {title}: {exc}\n")

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
