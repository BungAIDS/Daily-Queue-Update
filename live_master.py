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


def _norm(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "yes" if v else "no"
    return str(v).strip()


# The per-order fields we hold + watch for changes (label shown in the change
# log). Board fields + the Sales-Order spec; CO#, Drawings and Features are
# derived below. This is "all the info we have on each order" in one place.
_TRACKED = [
    ("oper", "Oper"), ("item", "Item"), ("assigned_to", "Assigned To"),
    ("checker", "Checker"), ("start_date", "Start Date"), ("end_date", "End Date"),
    ("plan_hrs", "Plan Hrs"), ("fannet_date", "FanNet Date"), ("total_price", "Total Price"),
    ("customer", "Customer"), ("primary_rep", "Primary Rep"), ("ship_with", "Ship With"),
    ("status_note", "Note"), ("unapproved", "Unapproved"), ("credit_hold", "Credit Hold"),
    ("so_design_desc", "Description"), ("so_size", "Size"), ("so_arrangement", "Arrangement"),
    ("so_motor_pos", "Motor Pos"), ("so_class", "Class"), ("so_rotation", "Rotation"),
    ("so_discharge", "Discharge"), ("so_pct_width", "% Width"), ("so_wheel_type", "Wheel Type"),
    ("so_design_temp", "Design Temp"), ("so_max_temp", "Max Temp"), ("so_special_temp", "Special Temp"),
]


def _suffix_sort(s: str):
    return (int(s), s) if str(s).isdigit() else (10 ** 9, s)


def tracked_values(job: Dict[str, Any]) -> Dict[str, str]:
    """The watched fields of an order as comparable strings, including the derived
    CO#, the custom-Drawings set, and the line-item Features set."""
    vals = {label: _norm(job.get(key)) for key, label in _TRACKED}
    co = job.get("co_number")
    vals["CO#"] = str(int(co)) if co else "0"
    de = job.get("dwg_extras") or {}
    vals["Drawings"] = ", ".join("-" + s for s in sorted(de, key=_suffix_sort))
    tags = sorted({t for it in (job.get("line_items") or []) for t in (it.get("tags") or [])})
    vals["Features"] = ", ".join(tags)
    return vals


def _diffs(old_job: Dict[str, Any], new_job: Dict[str, Any]) -> List[Tuple[str, str, str]]:
    """(field, old, new) for every watched field that was MODIFIED — i.e. it had a
    prior value and now differs. Skips initial population (''->value), which is
    just enrichment filling in, so the change log stays meaningful."""
    ov, nv = tracked_values(old_job), tracked_values(new_job)
    out = []
    for label, new in nv.items():
        old = ov.get(label, "")
        if old != new and old != "":
            out.append((label, old, new))
    return out


def update(master: Dict[str, Any], present: List[Dict[str, Any]],
           now: datetime) -> List[Dict[str, Any]]:
    """Fold the current board into the master log: upsert every present order
    (append if new, refresh its data + mark on_queue), then mark anything no
    longer present as off the board. `added` is set once; `left` is stamped on
    departure and cleared on return.

    Returns the list of field-change events detected this poll (one per modified
    field), so the caller can append them to the day's change log:
        {time, job, customer, field, old, new}
    """
    now_iso = now.isoformat(timespec="seconds")
    orders = master.setdefault("orders", {})
    present_nums = set()
    events: List[Dict[str, Any]] = []

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
            for field, old, new in _diffs(entry.get("job") or {}, j):
                events.append({"time": now_iso, "job": jn, "customer": _norm(j.get("customer")),
                               "field": field, "old": old, "new": new})
            entry["job"] = dict(j)
            entry["on_queue"] = True
            entry["left"] = None
            entry.setdefault("added", added)

    for jn, entry in orders.items():
        if jn not in present_nums and entry.get("on_queue"):
            entry["on_queue"] = False
            if not entry.get("left"):
                entry["left"] = now_iso

    return events


def merge_order(master: Dict[str, Any], job_num: str, fields: Dict[str, Any],
                when: datetime | None = None) -> bool:
    """Merge `fields` into one order's record in the master — the single place
    every helper (the live enrichment, line_items_scan, autocad_scan,
    quote_run_scan, backfill_orders, …) folds what it learned about an order.

    Only non-empty incoming values are written, and an existing non-empty value
    is never regressed to empty, so a sparse source (e.g. the archive scan, which
    has no board context) never wipes richer data. An order we've never seen is
    created off-queue (it's not on the board, just known). Returns True if
    anything actually changed."""
    jn = str(job_num or "").strip()
    if not jn:
        return False
    orders = master.setdefault("orders", {})
    now_iso = (when or datetime.now()).isoformat(timespec="seconds")
    entry = orders.get(jn)
    if entry is None:
        entry = orders[jn] = {"added": now_iso, "left": now_iso, "on_queue": False, "job": {"job": jn}}
    job = entry.setdefault("job", {})
    job.setdefault("job", jn)
    changed = False
    for k, v in (fields or {}).items():
        if v in (None, "", [], {}):
            continue                      # don't write/clobber with an empty value
        if job.get(k) != v:
            job[k] = v
            changed = True
    return changed


def ordered(master: Dict[str, Any]) -> List[Tuple[str, Dict[str, Any]]]:
    """All logged orders as (job#, entry), oldest-added first — the chronological
    order the Order History tab grows in (newest appended at the bottom)."""
    items = list(master.get("orders", {}).items())
    items.sort(key=lambda kv: (kv[1].get("added") or "", kv[0]))
    return items


def on_queue(master: Dict[str, Any]) -> List[Dict[str, Any]]:
    """The job dicts currently on the board, oldest-added first."""
    return [e["job"] for _, e in ordered(master) if e.get("on_queue") and e.get("job")]
