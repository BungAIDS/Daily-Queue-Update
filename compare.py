"""Diff today's queue against yesterday's JSON snapshot."""
from __future__ import annotations

import json
import logging
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Dict, List, Tuple

from config import SNAPSHOT_DIR

log = logging.getLogger(__name__)

# Fields we compare for changes (per row). Skip status because it churns a lot.
COMPARE_FIELDS = [
    "oper", "item_rep", "assigned_to", "checker",
    "start_date", "end_date", "plan_hrs", "fannet_date", "total_price",
    "customer", "ship_with",
]


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
) -> Dict[str, Any]:
    """Return new / removed / changed / persistent jobs."""
    today_idx = _index(today_jobs)
    yesterday_idx = _index(yesterday_jobs or [])

    new_jobs = [j for jn, j in today_idx.items() if jn not in yesterday_idx]
    removed_jobs = [j for jn, j in yesterday_idx.items() if jn not in today_idx]

    changed: List[Dict[str, Any]] = []
    for jn, today_job in today_idx.items():
        if jn not in yesterday_idx:
            continue
        prev = yesterday_idx[jn]
        changes: List[Tuple[str, str, str]] = []
        for field in COMPARE_FIELDS:
            old = (prev.get(field) or "").strip()
            new = (today_job.get(field) or "").strip()
            if old != new:
                changes.append((field, old, new))
        if changes:
            changed.append({"job": jn, "customer": today_job.get("customer", ""), "changes": changes})

    lookback = _load_lookback_jobsets(today)
    persistent: List[Dict[str, Any]] = []
    for jn, job in today_idx.items():
        days = _persistence_count(jn, lookback)
        if days >= 3:
            persistent.append({"job": jn, "customer": job.get("customer", ""), "days": days, "snapshot": job})

    return {
        "new": new_jobs,
        "removed": removed_jobs,
        "changed": changed,
        "persistent": persistent,
        "today_count": len(today_jobs),
        "yesterday_count": len(yesterday_jobs or []),
    }
