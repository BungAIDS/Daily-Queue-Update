"""Publish the watcher's log to a throwaway git branch so it can be read remotely.

The watcher runs on a Windows desktop and its console log only lives there
(LOG_DIR/watch.log). To make a bug inspectable without asking anyone to copy a
file off the machine, this force-pushes the current log file(s) as a single
*orphan* commit onto a dedicated branch (default 'debug-logs'). Each push replaces
the branch (force, no parent), so the repo never accumulates history/bloat and the
branch is always just the latest snapshot.

It's pure git plumbing — hash-object -> mktree -> commit-tree -> push — so it never
touches the working tree, the index, or the branch you're developing on. Anyone
can then read the log from that branch (e.g. GitHub's file view) without pulling.

Best-effort: any failure (offline, no credentials, not a git repo) is logged and
swallowed so it can never disturb the watch loop.
"""
from __future__ import annotations

import logging
import os
import subprocess
from datetime import datetime
from pathlib import Path
from typing import List

from config import LOG_DIR, LOG_PUSH_BRANCH

log = logging.getLogger(__name__)

_REPO = Path(__file__).resolve().parent

# A fixed identity so `git commit-tree` works even on a machine without
# user.name/user.email configured.
_IDENTITY = {
    "GIT_AUTHOR_NAME": "queue-watch-logger", "GIT_AUTHOR_EMAIL": "queue-watch@local",
    "GIT_COMMITTER_NAME": "queue-watch-logger", "GIT_COMMITTER_EMAIL": "queue-watch@local",
}


def _git(args: List[str], input_text: str | None = None) -> subprocess.CompletedProcess:
    env = {**os.environ, **_IDENTITY}
    return subprocess.run(["git", *args], cwd=_REPO, env=env, input=input_text,
                          capture_output=True, text=True, timeout=90)


def push_logs(branch: str | None = None) -> bool:
    """Force-push the current log file(s) as one orphan commit onto `branch`.
    Returns True on success. Never raises."""
    branch = (branch if branch is not None else LOG_PUSH_BRANCH) or ""
    if not branch:
        return False
    files = sorted(p for p in LOG_DIR.glob("watch.log*") if p.is_file())
    if not files:
        return False
    try:
        # 1. A blob per log file (written into the local object DB, unreferenced).
        entries = []
        for f in files:
            r = _git(["hash-object", "-w", str(f)])
            if r.returncode != 0:
                log.debug("log push: hash-object failed (%s)", r.stderr.strip())
                return False
            entries.append(f"100644 blob {r.stdout.strip()}\t{f.name}")
        # 2. A tree holding just those files.
        r = _git(["mktree"], input_text="\n".join(entries) + "\n")
        if r.returncode != 0:
            log.debug("log push: mktree failed (%s)", r.stderr.strip())
            return False
        tree = r.stdout.strip()
        # 3. An orphan commit (no -p parent) -> the branch is always one snapshot.
        msg = f"watch logs @ {datetime.now().isoformat(timespec='seconds')}"
        r = _git(["commit-tree", tree, "-m", msg])
        if r.returncode != 0:
            log.debug("log push: commit-tree failed (%s)", r.stderr.strip())
            return False
        commit = r.stdout.strip()
        # 4. Replace the remote branch with this snapshot.
        r = _git(["push", "--force", "origin", f"{commit}:refs/heads/{branch}"])
        if r.returncode != 0:
            log.warning("Log push to '%s' failed: %s", branch,
                        (r.stderr.strip() or r.stdout.strip())[:300])
            return False
        return True
    except (subprocess.SubprocessError, OSError) as e:  # noqa: BLE001
        log.debug("log push skipped (%s)", e)
        return False
