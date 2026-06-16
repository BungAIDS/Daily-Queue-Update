"""The all-time master log of every order the watcher has ever seen.

This is the persistent store behind the workbook's Order History tab — one entry
per order number, ever, kept in its most recent state and never wiped. It backs
the chronological "master list" the live workbook upserts into: each poll checks
every board order against this log, appends the ones it hasn't seen, and updates
the ones whose data changed. Unlike compare.py's history.json (which pops an
order when it returns), this log is append-only, so it's a faithful running
record.

Store shape:

    {
      "orders": {
        "<job#>": {
          "added":    ISO-8601,        # first time we ever saw it
          "left":     ISO-8601 | null, # when it last dropped off the board
          "on_queue": bool,            # on the board as of the last poll?
          "job":      { ...latest enriched job dict... },
        },
        ...
      }
    }

Pure dict/JSON logic (no Excel), so it's unit-tested directly (test_live_master).
"""
from __future__ import annotations

import json
import logging
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple

from config import SNAPSHOT_DIR

log = logging.getLogger(__name__)

MASTER_PATH = SNAPSHOT_DIR / "live_master.json"


def load_master() -> Dict[str, Any]:
    if not MASTER_PATH.exists():
        return {"orders": {}}
    try:
        data = json.loads(MASTER_PATH.read_text(encoding="utf-8"))
        if isinstance(data, dict) and isinstance(data.get("orders"), dict):
            return data
    except (json.JSONDecodeError, OSError) as e:
        log.warning("Could not read %s (%s); starting master log fresh", MASTER_PATH, e)
    return {"orders": {}}


def save_master(master: Dict[str, Any]) -> None:
    tmp = MASTER_PATH.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(master, indent=2, default=str), encoding="utf-8")
    tmp.replace(MASTER_PATH)


def _jobnum(j: Dict[str, Any]) -> str:
    return str(j.get("job") or "").strip()


def update(master: Dict[str, Any], present: List[Dict[str, Any]], now: datetime) -> None:
    """Fold the current board into the log: upsert every present order (append if
    new, refresh its data + mark on_queue), then mark anything that's no longer
    present as off the board. `added` is set once (prefers the order's first-seen
    marker), `left` is stamped when an order transitions off the board and cleared
    when it returns."""
    now_iso = now.isoformat(timespec="seconds")
    orders = master.setdefault("orders", {})
    present_nums = set()

    for j in present:
        jn = _jobnum(j)
        if not jn:
            continue
        present_nums.add(jn)
        entry = orders.get(jn)
        added = j.get("_first_seen") or now_iso
        if entry is None:
            orders[jn] = {"added": added, "left": None, "on_queue": True, "job": dict(j)}
        else:
            entry["job"] = dict(j)
            entry["on_queue"] = True
            entry["left"] = None
            entry.setdefault("added", added)

    for jn, entry in orders.items():
        if jn not in present_nums and entry.get("on_queue"):
            entry["on_queue"] = False
            if not entry.get("left"):
                entry["left"] = now_iso


def ordered(master: Dict[str, Any]) -> List[Tuple[str, Dict[str, Any]]]:
    """All logged orders as (job#, entry), oldest-added first — the chronological
    order the Order History tab grows in (newest appended at the bottom)."""
    items = list(master.get("orders", {}).items())
    items.sort(key=lambda kv: (kv[1].get("added") or "", kv[0]))
    return items


def on_queue(master: Dict[str, Any]) -> List[Dict[str, Any]]:
    """The job dicts currently on the board, oldest-added first."""
    return [e["job"] for _, e in ordered(master) if e.get("on_queue") and e.get("job")]
