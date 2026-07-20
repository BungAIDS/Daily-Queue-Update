# AGENTS.md

This file provides guidance to AI coding agents when working with code in this repository.

<!-- AUTO-MANAGED: project-description -->
## Overview

Daily Queue Update — tooling around a daily sales-order/queue workflow for the cbcinsider.com engineering dispatch queue:

- `main.py` — the 5 AM daily job: scrape the work queue, diff against yesterday's snapshot, ask Claude for a briefing/anomalies/ranked action items, build a two-tab Excel report, email it via desktop Outlook.
- `watch.py` — all-day live intraday watcher, the companion to the daily run; maintains live Excel sheets and an all-time master log.
- `launcher.py` — standard-library Tkinter desktop launcher that runs/stops the scripts on the user's Windows PC.
- Transmittal/AutoCAD helpers — `fill_transmittal_insider.py` (Email Drawings: probe + pre-fill only, SEND disabled), `transmittal_data.py`, `transmittal_doc.py`, `autocad_scan.py`.

The live scrape, Z: drive access, and Outlook email are Windows-only and unreachable from CI or a cloud container; the plain-Python logic modules are testable anywhere.

<!-- END AUTO-MANAGED -->

<!-- AUTO-MANAGED: build-commands -->
## Build & Development Commands

- Setup (Windows target machine): `python -m venv venv` → `venv\Scripts\activate` → `pip install -r requirements.txt` → `playwright install chromium`
- Run one test suite: `python test_<name>.py` — tests are plain scripts with `main()` runners, NOT pytest
- CI (`.github/workflows/tests.yml`) runs the pure-logic suites with only `playwright openpyxl python-dotenv pdfplumber` installed (`pywin32` has no Linux build)
- Read the latest launcher debug report (do this FIRST when debugging the launcher): `git fetch origin debug/launcher && git show origin/debug/launcher:diagnostics/launcher_report.txt`
- Runtime configuration lives in `.env` (see README) and `config.py`

<!-- END AUTO-MANAGED -->

<!-- AUTO-MANAGED: architecture -->
## Architecture

Flat single-directory layout — all modules live at the repo root.

Daily 5 AM pipeline (`main.py`, stages in `pipeline.py`):
- `login.py` / `scraper.py` — Playwright login and scrape of the dispatch queue
- `compare.py`, `analyzer.py`, `brief.py` — diff vs yesterday's snapshot, Claude briefing and anomaly analysis
- `excel_writer.py` — two-tab Excel report; `emailer.py` / `send.py` — desktop Outlook email
- `snapshots/` JSON drives tomorrow's diff; outputs older than 60 days are archived, never deleted

Live watcher (`watch.py`):
- `live_state.py`, `live_sheets.py`, `live_excel.py`, `live_master.py` — intraday state, live workbook, all-time master log
- `so_review.py`, `sales_order_validation.py`, `so_hierarchy.py`, `line_items.py` — sales-order review workbook, validation, and parsing
- `stop_signal.py` — per-PID stop-flag files (`.launcher_stops/`) polled by the script and treated like Ctrl+C, because the launcher runs under `pythonw` and cannot deliver Ctrl+Break; 150s backstop force-kill

Launcher (`launcher.py`, Tkinter):
- `procscan.py` — import-light parsing for external-process detection (`wmic` → PowerShell CIM query fallback → `ps` off-Windows); scans run on a worker thread so a hung PowerShell can never freeze the UI
- `git_update.py` — pure, import-light (no tkinter) git pull/branch logic; the Tk `GitUpdateDialog` is a thin GUI on top
- Per-run logs go to `launcher_logs/` (git-ignored, exists only on the user's PC); debug reports are published to the dedicated `debug/launcher` branch on every launcher close and via the Publish Debug Report button
- Single-instance via `.launcher.lock`; Git Update persists stopped programs in `.launcher_state.json` under `pending_restart` and auto-relaunches the launcher when its own files changed

<!-- END AUTO-MANAGED -->

<!-- AUTO-MANAGED: conventions -->
## Code Conventions

- `snake_case` module and function names; flat module layout at the repo root
- Keep `launcher.py` standard-library only — no new pip dependencies for the GUI
- `tkinter` is not always installed (absent in cloud containers, may be absent in CI) — tests must never `import launcher`; test the import-light modules instead (e.g. `git_update.py`, `procscan.py`)
- New pull/branch logic goes in `git_update.py`, keeping the GUI dialog thin
- Tests are plain scripts: `test_*` functions plus a `main()` runner, run with `python test_<name>.py`
- When renaming a script, update the launcher action's `script=` path and any doc-comment references; the launcher action `id` is a separate stable key and does not need to match the filename
- Windows-only dependencies are guarded with `sys_platform == "win32"` markers in `requirements.txt`

<!-- END AUTO-MANAGED -->

<!-- AUTO-MANAGED: patterns -->
## Detected Patterns

- Import-light "logic module + thin GUI/driver" separation: `git_update.py` vs `GitUpdateDialog`, `procscan.py` vs the launcher's scan workers
- A running Python process keeps the code it loaded at startup — a `git pull` only takes effect after restart; the launcher tracks this (`pending_restart`, `relaunch_self_and_exit`) rather than pretending otherwise
- Anything that can hang runs on a worker thread (`_scan_worker` → `scan_result` → `_apply_scan_result`); synchronous scans only for rare one-off Run/Pull pre-checks
- Graceful stop over hard kill: stop-flag files the script polls like Ctrl+C (finish the poll, save state, publish logs), force-kill only as backstop or on second click
- Stops requested via the launcher are recorded in `_stop_requested` so the exit shows `[STOPPED]` (neutral) instead of `[FAIL]` (red)
- Runtime state at repo root (`.launcher_state.json`, `.launcher.lock`, `so_review_handled.json`), runtime dirs git-ignored

<!-- END AUTO-MANAGED -->

<!-- AUTO-MANAGED: git-insights -->
## Git Insights

- `debug/launcher` is a dedicated machine-published branch for launcher debug reports, kept off feature branches; each publish is a new commit, so `git log origin/debug/launcher` is the history of the user's sessions
- The tip of `debug/launcher` normally reflects the end of the user's last launcher session; for a snapshot of a current problem, ask the user to click Publish Debug Report (or close/reopen the launcher)
- The report contains OS/Python info, the external-status process scan, launcher-started processes, last exit codes, the tail of the last Email Drawings per-run log (with the transmittal decision summary), and the tail of `launcher_debug.log`
- A green CI check means "the pure logic didn't regress", NOT "tomorrow's 5 AM run will work" — the Windows-only paths are unreachable from hosted runners
- The launcher header shows `version: branch@commit` (also in the debug report and watch.py's startup log) so the running commit is always visible

<!-- END AUTO-MANAGED -->

<!-- MANUAL -->
## Custom Notes

Add project-specific notes here. This section is never auto-modified.

### Endpoint-security ground rules (agents running on DG's Windows PC)

Background: the PC's antivirus quarantined agent-related files twice
(2026-07-10 13:37, 2026-07-15 09:44). The quarantine lists were analyzed
2026-07-20. Every flagged file was one of: the Codex desktop app itself
(`ChatGPT.exe`, two different auto-updated builds), Codex's bundled plugin
payloads (the latex plugin's `tectonic.exe`, npm packages `abstract-level` /
`classic-level` / `node-gyp-build` under the chrome/browser plugins, a temp
`applypatch.bat` shim), third-party DLLs in the machine-wide NuGet cache
(PDFsharp/MigraDoc — nothing in this repo is .NET), or a stale
`__pycache__/line_items.cpython-313.pyc` byte-compiled inside an agent's
working copy of this repo. All ~20 files in the second event were quarantined
in the same second — a scheduled heuristic scan sweeping low-reputation /
newly-written binaries in one pass, not a detection of something an agent
executed. Two of the "threats" were localization-only `.resources.dll`
satellites (no executable code), the signature of a false-positive sweep, not
malware. The binding lesson: every new binary an agent drops on this machine
is antivirus bait, and a mid-run quarantine can silently break the live
automation. Therefore, for any agent working on the local PC:

- Stay inside the pinned Python toolchain: `pip install -r requirements.txt`
  plus `playwright install chromium`, nothing else. Never install or download
  any other toolchain or standalone binary onto this machine — no .NET/NuGet
  packages (PDFsharp, MigraDoc), no LaTeX engines (tectonic), no npm/
  `node_modules`, no portable `.exe`/`.dll` tools, no Codex marketplace
  plugins (latex/chrome/browser). Every document task already has an in-repo
  path: pdfplumber + `pdf_vision.py` read PDFs, Word/Outlook/Excel COM via
  pywin32 write documents and email, openpyxl builds workbooks. A task that
  seems to need a new toolchain is a task to stop and ask DG about first.
- Do not write executable files (`.bat`, `.ps1`, `.vbs`, `.exe`) anywhere
  outside the repo working tree, and add new ones inside it only when the task
  is explicitly about the launcher/run scripts.
- In disposable agent clones (connector work dirs), set
  `PYTHONDONTWRITEBYTECODE=1` before running code or tests so no `__pycache__`
  bytecode is left on disk for heuristic scanners to chew on — a stale
  agent-clone `.pyc` of `line_items.py` was among the quarantined files.
- NEVER touch the antivirus: no exclusions, no pausing/disabling, no restoring
  files from quarantine, no "fixing" a detection. If the AV flags something,
  stop and report it to DG verbatim.
- If a local run fails inexplicably (module suddenly missing, file vanishing
  mid-run), check the antivirus protection history for a fresh quarantine
  before debugging the code.

Note: flags on the Codex app binary and its bundled plugin files cannot be
prevented from inside this repo — those files exist the moment the app
installs or auto-updates. That recurring false positive is resolved in the AV
console (restore + report false positive / exclusion decision by DG), not by
agent behavior.

<!-- END MANUAL -->
