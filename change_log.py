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
from datetime import date
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
