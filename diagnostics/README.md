# diagnostics/

Shareable debug snapshots from the desktop launcher (`launcher.py`).

The launcher's per-run logs live under `launcher_logs/`, which is **git-ignored**
and never leaves the workstation. When something needs debugging remotely, that
output has to reach the repo somehow — that is what this folder is for.

## How to use it

1. In the launcher, click **Publish Debug Report**.
2. That writes **`launcher_report.txt`** here (a snapshot: OS/Python info, the
   external-status process scan — which method worked, what it saw, what it
   detected — launcher-started processes, last exit codes, and the tail of
   `launcher_debug.log`) and **pushes it to the `debug/launcher` branch**
   automatically, without touching your current checkout.
3. If the push fails (e.g. no network), the file is still saved here and can be
   pushed by hand.

It may contain process command lines and local file paths — glance over it
before sharing.

> The published copy lives on the **`debug/launcher`** branch, kept separate from
> feature branches. Read the latest with:
> `git show origin/debug/launcher:diagnostics/launcher_report.txt`. Each publish
> is a new commit, so history keeps every snapshot.
