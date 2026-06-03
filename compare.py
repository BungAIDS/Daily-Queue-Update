"""Diff today's queue against yesterday's JSON snapshot.

Also keeps a long-term `history.json` of every job that has ever dropped off
the queue, so a job reappearing later can be reported as "returning" instead
of being mistakenly labeled "new".
"""
from __future__ import annotations

import json
import logging
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Dict, List, Tuple

from config import SNAPSHOT_DIR

log = logging.getLogger(__name__)

# Fields we compare for changes (per row). Skip status because it churns a lot,
# and skip has_notes because a note appearing/disappearing is low signal.
COMPARE_FIELDS = [
    "oper", "item", "assigned_to", "checker",
    "start_date", "end_date", "plan_hrs", "fannet_date", "total_price",
    "customer", "primary_rep", "ship_with",
    "status_note", "unapproved", "credit_hold",
]


def _norm(v: Any) -> str:
    """Normalize a field value to a comparable string (handles bools/None)."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "yes" if v else "no"
    return str(v).strip()


HISTORY_PATH = SNAPSHOT_DIR / "history.json"


def load_history() -> Dict[str, Any]:
    """Job# -> {last_seen: ISO date, snapshot: {...}}. Empty if no file yet."""
    if not HISTORY_PATH.exists():
        return {}
    try:
        return json.loads(HISTORY_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as e:
        log.warning("Could not read %s (%s); treating as empty", HISTORY_PATH, e)
        return {}


def save_history(history: Dict[str, Any]) -> None:
    HISTORY_PATH.write_text(json.dumps(history, indent=2), encoding="utf-8")
    log.info("Updated history (%d archived jobs)", len(history))


def snapshot_path(d: date) -> Path:
    return SNAPSHOT_DIR / f"queue_{d.isoformat()}.json"


def save_snapshot(jobs: List[Dict[str, Any]], d: date) -> Path:
    path = snapshot_path(d)
    path.write_text(json.dumps(jobs, indent=2), encoding="utf-8")
    log.info("Saved snapshot: %s (%d jobs)", path, len(jobs))
    return path


def load_snapshot(d: date) -> List[Dict[str, Any]] | None:
    path = snapshot_path(d)
    if not path.exists():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def _index(jobs: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    return {j["job"]: j for j in jobs if j.get("job")}


def _load_lookback_jobsets(today: date, max_lookback: int = 14) -> List[set]:
    """Sets of job numbers present on today-1, today-2, ... stopping at the
    first day that has no snapshot (a missing day breaks the streak).

    Loaded once per run so persistence is O(jobs) instead of re-reading and
    re-parsing every snapshot file once per job.
    """
    day_sets: List[set] = []
    for delta in range(1, max_lookback):
        prev = load_snapshot(today - timedelta(days=delta))
        if prev is None:
            break
        day_sets.append({j.get("job") for j in prev if j.get("job")})
    return day_sets


def _persistence_count(job_num: str, lookback: List[set]) -> int:
    """Consecutive days (including today) this job has been in the queue."""
    count = 1
    for day_jobs in lookback:
        if job_num in day_jobs:
            count += 1
        else:
            break
    return count


def diff_queues(
    today_jobs: List[Dict[str, Any]],
    yesterday_jobs: List[Dict[str, Any]] | None,
    today: date,
    persist_history: bool = True,
) -> Dict[str, Any]:
    """Return new / returning / removed / changed / persistent jobs.

    Also updates the long-term history store as a side effect: jobs that left
    the queue get archived to it, and jobs that come back get popped off.

    Set persist_history=False for ad-hoc/manual reports (make_report.py) so the
    official tracking state is only ever advanced by the once-a-day run. With
    it False, history is still read to label returning orders, but not written.
    """
    today_idx = _index(today_jobs)
    yesterday_idx = _index(yesterday_jobs or [])

    history = load_history()
    yesterday_date = (today - timedelta(days=1)).isoformat()

    # New today = in today, not in yesterday. Split by whether we've ever seen
    # them before: anything in history is "returning"; otherwise truly new.
    new_jobs: List[Dict[str, Any]] = []
    returning_jobs: List[Dict[str, Any]] = []
    for jn, j in today_idx.items():
        if jn in yesterday_idx:
            continue
        prev = history.pop(jn, None)
        if prev is not None:
            entry = dict(j)
            entry["_last_seen"] = prev.get("last_seen", "")
            # If it came back at a higher CO# than when it left, flag *why* it
            # returned: a change order. Only compare when the archived snapshot
            # actually carried a co_number (older snapshots predate the field).
            snap = prev.get("snapshot") or {}
            if "co_number" in snap:
                old_co, new_co = int(snap.get("co_number") or 0), int(j.get("co_number") or 0)
                if new_co > old_co:
                    entry["_co_returned"] = {"old_co": old_co, "new_co": new_co}
            returning_jobs.append(entry)
        else:
            new_jobs.append(j)

    # Removed today = in yesterday, not in today. Archive them so we can
    # recognize them if they ever come back.
    removed_jobs: List[Dict[str, Any]] = []
    for jn, j in yesterday_idx.items():
        if jn not in today_idx:
            removed_jobs.append(j)
            history[jn] = {"last_seen": yesterday_date, "snapshot": j}

    changed: List[Dict[str, Any]] = []
    for jn, today_job in today_idx.items():
        if jn not in yesterday_idx:
            continue
        prev = yesterday_idx[jn]
        changes: List[Tuple[str, str, str]] = []
        for field in COMPARE_FIELDS:
            old = _norm(prev.get(field))
            new = _norm(today_job.get(field))
            if old != new:
                changes.append((field, old, new))
        if changes:
            changed.append({"job": jn, "customer": today_job.get("customer", ""), "changes": changes})

    # Change orders that landed today: a job present both days whose CO# rose.
    # Only compare when yesterday's snapshot carried co_number (older snapshots
    # predate it), so the first enriched run doesn't flag the whole board.
    co_changed: List[Dict[str, Any]] = []
    for jn, today_job in today_idx.items():
        prev = yesterday_idx.get(jn)
        if not prev or "co_number" not in prev:
            continue
        old_co, new_co = int(prev.get("co_number") or 0), int(today_job.get("co_number") or 0)
        if new_co > old_co:
            co_changed.append({
                "job": jn,
                "customer": today_job.get("customer", ""),
                "old_co": old_co,
                "new_co": new_co,
                "co_history": today_job.get("co_history", []),
            })

    lookback = _load_lookback_jobsets(today)
    persistent: List[Dict[str, Any]] = []
    for jn, job in today_idx.items():
        days = _persistence_count(jn, lookback)
        if days >= 3:
            persistent.append({"job": jn, "customer": job.get("customer", ""), "days": days, "snapshot": job})

    if persist_history:
        save_history(history)

    return {
        "new": new_jobs,
        "returning": returning_jobs,
        "removed": removed_jobs,
        "changed": changed,
        "co_changed": co_changed,
        "persistent": persistent,
        "today_count": len(today_jobs),
        "yesterday_count": len(yesterday_jobs or []),
    }
