"""Publish the order-data files to a throwaway git branch so they can be read
remotely (by a teammate or an assistant) without copying them off the machine.

Same mechanism as `log_push.py`: force-push the current data files as one
*orphan* commit onto a dedicated branch (default 'order-data'). Each push
replaces the branch (force, no parent), so the repo never accumulates
history/bloat and the branch is always just the latest snapshot.

It's pure git plumbing — hash-object -> mktree -> commit-tree -> push — so it
never touches the working tree, the index, or the branch you're developing on.
A clone can then read the data from that branch with
`git fetch origin <branch> && git show <branch>:live_master.json` (or check the
files out), without disturbing anyone's work.

Files published (whichever exist): live_master.json, the resumable JSON stores
(quote-run / line-item / backfill / autocad), and the human xlsx sheets
(quote_runs, backlog, line_items, autocad_dwgs). Missing files are skipped.

NOTE: this data (job #s, customers, prices, file paths) goes to that repo — the
same caveat as the log push. Only point DATA_PUSH_BRANCH at a PRIVATE repo.

Best-effort: any failure (offline, no credentials, not a git repo) is logged and
swallowed so it can never disturb a scan/daily run.
"""
from __future__ import annotations

import logging
import os
import subprocess
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import List, Optional

from config import BACKLOG_DIR, SNAPSHOT_DIR, LINE_ITEMS_STORE, DATA_PUSH_BRANCH

log = logging.getLogger(__name__)

_REPO = Path(__file__).resolve().parent

# A fixed identity so `git commit-tree` works even on a machine without
# user.name/user.email configured.
_IDENTITY = {
    "GIT_AUTHOR_NAME": "queue-data-publisher", "GIT_AUTHOR_EMAIL": "queue-data@local",
    "GIT_COMMITTER_NAME": "queue-data-publisher", "GIT_COMMITTER_EMAIL": "queue-data@local",
}


def data_files() -> List[Path]:
    """The order-data files we publish, in a stable order. Only the ones that
    actually exist on disk are returned (a fresh machine may lack some)."""
    line_items = LINE_ITEMS_STORE if LINE_ITEMS_STORE else BACKLOG_DIR / "line_items.json"
    today = date.today()
    candidates = [
        SNAPSHOT_DIR / "live_master.json",           # the master store (richest)
        # Today's + yesterday's field-change logs: the Changes tab's source, so
        # a page/report built from this branch shows the day's activity too.
        SNAPSHOT_DIR / f"change_log_{today.isoformat()}.json",
        SNAPSHOT_DIR / f"change_log_{(today - timedelta(days=1)).isoformat()}.json",
        BACKLOG_DIR / "quote_run_scan_progress.json",  # every parsed quote run
        line_items,                                  # line-item store
        BACKLOG_DIR / "backfill_line_items.json",   # watcher-safe backfill overlay
        BACKLOG_DIR / "backfill_progress.json",
        BACKLOG_DIR / "autocad_scan_progress.json",
        BACKLOG_DIR / "so_review_notes.json",        # the Sales-Order note queue
        BACKLOG_DIR / "so_review_parser_metrics.json",  # review-count history
        BACKLOG_DIR / "quote_run_review_notes.json",  # the Quote-Run note queue

        BACKLOG_DIR / "quote_runs.xlsx",             # the human sheets
        BACKLOG_DIR / "backlog.xlsx",
        BACKLOG_DIR / "line_items.xlsx",
        BACKLOG_DIR / "autocad_dwgs.xlsx",
    ]
    cleanup_audit = BACKLOG_DIR / "order_verification_cleanup.json"
    try:
        cleanup_mtime = cleanup_audit.stat().st_mtime
    except OSError:
        cleanup_mtime = 0.0

    seen, out = set(), []
    for p in candidates:
        rp = Path(p)
        if rp in seen:
            continue
        seen.add(rp)
        if rp.is_file():
            # A cleanup can invalidate rows in the JSON stores before these
            # derived spreadsheets are regenerated. Do not publish a stale
            # workbook containing report-derived Sales Order data.
            if rp.name in {"backlog.xlsx", "line_items.xlsx"} and cleanup_mtime:
                try:
                    if rp.stat().st_mtime < cleanup_mtime:
                        continue
                except OSError:
                    continue
            out.append(rp)
    return out


def _git(args: List[str]) -> subprocess.CompletedProcess:
    env = {**os.environ, **_IDENTITY}
    return subprocess.run(["git", *args], cwd=_REPO, env=env,
                          capture_output=True, text=True, timeout=120)


def _git_stdin(args: List[str], data: bytes) -> subprocess.CompletedProcess:
    """Feed `data` on stdin as raw BYTES so Windows can't turn the '\\n' tree
    separators into '\\r\\n' (which would leave a stray '\\r' on each filename)."""
    env = {**os.environ, **_IDENTITY}
    return subprocess.run(["git", *args], cwd=_REPO, env=env, input=data,
                          capture_output=True, timeout=120)


def build_snapshot_commit(files: List[Path], message: str) -> Optional[str]:
    """Write `files` into the object DB as one orphan commit (no parent, no
    working-tree/index changes) and return its sha. None on any failure. Two
    files that share a basename collide in the flat tree — callers pass distinct
    names (our data files all differ)."""
    if not files:
        return None
    try:
        entries = []
        for f in files:
            r = _git(["hash-object", "-w", str(f)])
            if r.returncode != 0:
                log.debug("data push: hash-object failed for %s (%s)", f, r.stderr.strip())
                return None
            name = Path(f).name.replace("\r", "").replace("\n", "")   # defensive
            entries.append(f"100644 blob {r.stdout.strip()}\t{name}")
        r = _git_stdin(["mktree"], ("\n".join(entries) + "\n").encode("utf-8"))
        if r.returncode != 0:
            log.debug("data push: mktree failed (%s)", r.stderr.decode(errors="replace").strip())
            return None
        tree = r.stdout.decode().strip()
        r = _git(["commit-tree", tree, "-m", message])
        if r.returncode != 0:
            log.debug("data push: commit-tree failed (%s)", r.stderr.strip())
            return None
        return r.stdout.strip()
    except (subprocess.SubprocessError, OSError) as e:  # noqa: BLE001
        log.debug("data push: snapshot build skipped (%s)", e)
        return None


def push_data(branch: Optional[str] = None) -> bool:
    """Force-push the current order-data files as one orphan commit onto `branch`.
    Returns True on success. Never raises."""
    branch = (branch if branch is not None else DATA_PUSH_BRANCH) or ""
    if not branch:
        log.info("DATA_PUSH_BRANCH is empty — order-data publishing is disabled.")
        return False
    files = data_files()
    if not files:
        log.warning("No order-data files found to publish (looked under %s / %s).",
                    SNAPSHOT_DIR, BACKLOG_DIR)
        return False
    msg = f"order data @ {datetime.now().isoformat(timespec='seconds')} ({len(files)} files)"
    commit = build_snapshot_commit(files, msg)
    if not commit:
        log.warning("Could not build the order-data snapshot commit.")
        return False
    r = _git(["push", "--force", "origin", f"{commit}:refs/heads/{branch}"])
    if r.returncode != 0:
        log.warning("Order-data push to '%s' failed: %s", branch,
                    (r.stderr.strip() or r.stdout.strip())[:300])
        return False
    log.info("Published %d order-data file(s) to '%s': %s",
             len(files), branch, ", ".join(p.name for p in files))
    return True


def main(argv: Optional[List[str]] = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    args = sys.argv[1:] if argv is None else argv
    branch = args[0].strip() if args and args[0].strip() else None
    return 0 if push_data(branch) else 1


if __name__ == "__main__":
    raise SystemExit(main())
