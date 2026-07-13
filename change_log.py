"""Per-day log of field-level changes to orders on the board.

Every poll, live_master.update returns the fields that were modified since the
last scan; those events are appended here, one file per day
(SNAPSHOT_DIR/change_log_<date>.json). The same field changing several times in a
day is several events, so the Changes tab can show each modification as its own
line. Each event:

    {"time": ISO-8601, "job": "421000", "customer": "ACME",
     "field": "End Date", "old": "06/20/2026", "new": "06/25/2026"}

Pure JSON list logic; unit-tested in test_change_log.
"""
from __future__ import annotations

import json
import logging
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List

from config import SNAPSHOT_DIR

log = logging.getLogger(__name__)


def log_path(d: date) -> Path:
    return SNAPSHOT_DIR / f"change_log_{d.isoformat()}.json"


def load(d: date) -> List[Dict[str, Any]]:
    p = log_path(d)
    if not p.exists():
        return []
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except (json.JSONDecodeError, OSError) as e:
        log.warning("Could not read %s (%s); treating as empty", p, e)
        return []


def save(d: date, events: List[Dict[str, Any]]) -> None:
    p = log_path(d)
    tmp = p.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(events, indent=2, default=str), encoding="utf-8")
    tmp.replace(p)


def append(d: date, events: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Append `events` to day `d`'s log and return the full day's log."""
    if not events:
        return load(d)
    full = load(d)
    full.extend(events)
    save(d, full)
    return full


def purge_day_once(d: date) -> int:
    """ONE-SHOT cleanup (2026-07-13): the day's log accumulated flip artifacts
    from the pre-fix backfill-merge era, tangled in with real changes; rather
    than guess which is which, the first startup on this code archives the
    day's log to <name>.json.bak and lets the Changes tab start clean. A
    permanent marker file stops this ever running again — on any later day —
    so real changes are never dropped after the one shot. Returns how many
    events were archived (0 once the marker exists)."""
    marker = SNAPSHOT_DIR / ".change_log_purged_once"
    if marker.exists():
        return 0
    events = load(d)
    p = log_path(d)
    try:
        if p.exists():
            p.replace(p.parent / (p.name + ".bak"))
        marker.write_text(datetime.now().isoformat(timespec="seconds"), encoding="utf-8")
    except OSError as e:
        log.warning("One-time change-log purge failed (%s)", e)
        return 0
    return len(events)


def scrub_phantom_blanks(d: date, master: Dict[str, Any]) -> int:
    """Drop phantom '-> (blank)' events from day `d`'s log: ones claiming a field
    went blank while the master still holds the exact non-empty value the event
    said it changed FROM — i.e. the blanking never really happened. These were
    mass-logged by the update()/save_master flip-flop (a board-only job dict
    stripping enrichment fields that the save then revived from disk), repeating
    every poll. Real blankings survive: for those the master's current value is
    "" (or was later re-set to something other than the old value).

    Returns how many events were removed; rewrites the log only when it changes.
    """
    events = load(d)
    if not events:
        return 0
    import live_master   # local import: keep this module free of heavier deps

    current: Dict[str, Dict[str, str]] = {}
    for jn, entry in (master.get("orders") or {}).items():
        current[jn] = live_master.tracked_values(entry.get("job") or {})

    kept = []
    for ev in events:
        vals = current.get(str(ev.get("job") or ""))
        phantom = (not ev.get("new") and ev.get("old") and vals is not None
                   and vals.get(str(ev.get("field") or "")) == ev.get("old"))
        if not phantom:
            kept.append(ev)
    dropped = len(events) - len(kept)
    if dropped:
        save(d, kept)
    return dropped
