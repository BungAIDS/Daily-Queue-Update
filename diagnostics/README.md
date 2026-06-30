# diagnostics/

Shareable debug snapshots from the desktop launcher (`launcher.py`).

The launcher's per-run logs live under `launcher_logs/`, which is **git-ignored**
and never leaves the workstation. When something needs debugging remotely, that
output has to reach the repo somehow — that is what this folder is for.

## How to use it

The launcher publishes a report to the `debug/launcher` branch **automatically
every time it is closed**, and on demand:

1. In the launcher, click **Publish Debug Report** (or just close the launcher).
2. That builds a snapshot — OS/Python info, the external-status process scan
   (which method worked, what it saw, what it detected), launcher-started
   processes, last exit codes, and the tail of `launcher_debug.log` — and
   **pushes it to the `debug/launcher` branch** without touching your checkout.
3. If the push fails (e.g. no network), a local copy is still saved under
   `launcher_logs/` (git-ignored).

It may contain process command lines and local file paths — glance over it
before sharing.

> The published copy lives on the **`debug/launcher`** branch, kept separate from
> feature branches. Read the latest with:
> `git show origin/debug/launcher:diagnostics/launcher_report.txt`. Each publish
> is a new commit, so history keeps every snapshot.
