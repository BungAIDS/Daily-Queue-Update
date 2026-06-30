# diagnostics/

Shareable debug snapshots from the desktop launcher (`launcher.py`).

The launcher's per-run logs live under `launcher_logs/`, which is **git-ignored**
and never leaves the workstation. When something needs debugging remotely, that
output has to reach the repo somehow — that is what this folder is for.

## How to use it

1. In the launcher, click **Export Debug Report**.
2. That writes/overwrites **`launcher_report.txt`** here with a snapshot:
   OS/Python info, the external-status process scan (which method worked, what it
   saw, what it detected), launcher-started processes, last exit codes, and the
   tail of `launcher_debug.log`.
3. Commit and push `launcher_report.txt` so it can be reviewed.

It may contain process command lines and local file paths — glance over it
before sharing.

> This folder is tracked on purpose. `launcher_report.txt` is meant to be
> overwritten and re-committed; git history keeps the previous snapshots.
