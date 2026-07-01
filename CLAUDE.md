# CLAUDE.md

Guidance for working in this repo. Read this before debugging the launcher or
touching GUI/Git code.

## What this project is

Tooling around a daily sales-order/queue workflow: scrapers, Excel reporting,
AutoCAD/transmittal helpers, an all-day live watcher (`watch.py`), and a
standard-library Tkinter desktop launcher (`launcher.py`) that runs/stops the
scripts. Most modules are plain-Python logic; the live scrape, Z: drive, and
Outlook email are Windows-only and unreachable from CI / a cloud container.

## Debugging the launcher — CHECK HERE FIRST

The launcher writes per-run logs to `launcher_logs/`, which is **git-ignored**
and only exists on the user's Windows PC — you will not see it in this repo.

Launcher debug reports are published to a dedicated **`debug/launcher`** branch
(kept off feature branches). When debugging the launcher, **read the latest
report from that branch first**:

```
git fetch origin debug/launcher
git show origin/debug/launcher:diagnostics/launcher_report.txt
```

A fresh `diagnostics/launcher_report.txt` is pushed to `debug/launcher`
automatically **every time the launcher is closed**, and on demand via the
**Publish Debug Report** button (git plumbing — it does not touch their
checkout). So the tip of `debug/launcher` normally reflects the end of the
user's last session. The report contains: OS/Python info, the external-status
process scan (method used / error / what it saw / what it detected),
launcher-started processes, last exit codes, and the tail of
`launcher_debug.log`. If you need a snapshot of a *current* problem, ask the user
to click **Publish Debug Report** (or just close/reopen the launcher). Each
publish is a new commit, so `git log origin/debug/launcher` shows the history.

Notable launcher internals:
- External "running outside the launcher" detection scans process command lines
  via `wmic`, falling back to a PowerShell CIM query (`wmic` is removed on newer
  Windows) and `ps` off-Windows. Parsing lives in the import-light `procscan.py`
  (tested by `test_procscan.py`); the scan also captures PIDs so an external copy
  can be force-stopped (Stop button → `taskkill`, or `stop_any` during a Git
  Update). A grey/idle dot for something that IS running usually means the scan
  found nothing — check the report's "method used". The recurring poll runs the
  scan on a worker thread (`_scan_worker` → `scan_result` → `_apply_scan_result`)
  so a slow/hung PowerShell call can never freeze the UI; only the rare one-off
  Run/Pull pre-checks scan synchronously (`_scan_now_sync`).
- Stopping a tool from the launcher records it in `_stop_requested`, so its
  non-zero exit shows as `[STOPPED]` (neutral) rather than `[FAIL]` (red).
- Single-instance: `_acquire_single_instance` holds `.launcher.lock` (its PID);
  a second launch waits ~5s (relaunch handoff) then warns. Closing the launcher
  stops *all* launcher-started processes (`_on_close`, force) and releases the
  lock. The header shows `version: branch@commit` (and it's in the debug report
  + watch.py's startup log) so the running commit is always visible.
- The launcher runs under `pythonw` (no console) so it can't deliver Ctrl+Break.
  For a `graceful_stop` action (watch.py) the Stop button drops a per-PID flag
  file (`stop_signal.py`, `.launcher_stops/`) that the script polls and treats
  like Ctrl+C — finishing the poll, saving state, publishing logs — with a
  150s backstop force-kill; a second Stop click (or the Git Update flow, which
  passes `force=True`) kills immediately.
- A running Python process (watcher OR the launcher itself) keeps the code it
  loaded at startup; a `git pull` only takes effect after a restart. The Git
  Update window offers to stop *all* running programs for the update and
  persists them in `.launcher_state.json` under `pending_restart`. When the pull
  changes the launcher's own files (or programs were stopped) it auto-relaunches
  itself (`relaunch_self_and_exit`, after the stopped programs fully exit) and
  the fresh launcher resumes them on startup (`_restart_pending`, skipping
  `email_risk` actions). If the user chose to leave programs running, it does
  NOT auto-relaunch (that would orphan them) and asks them to reopen manually.

## Git pull logic lives in `git_update.py`

Pure, import-light (no tkinter) so it is unit-testable without a display. The Tk
`GitUpdateDialog` in `launcher.py` drives it. Keep new pull/branch logic in
`git_update.py` with the GUI thin on top.

## Tests

Plain scripts, not pytest. Each `test_*.py` has `test_*` functions and a
`main()` runner; run with `python test_<name>.py`. CI (`.github/workflows/
tests.yml`) runs the pure-logic suites. **`tkinter` is not always installed**
(it is absent in this container and may be in CI), so do not write tests that
`import launcher`; test the import-light modules (e.g. `git_update.py`,
`test_git_update.py`) instead.

## Conventions

- Keep `launcher.py` standard-library only (no new pip deps for the GUI).
- When you rename a script, update the launcher action's `script=` path and any
  doc-comment references; the launcher action `id` is a separate stable key and
  does not need to match the filename.
