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


# Enrichment (Sales-Order / quote-run / folder) fields. When a re-fetch or the
# morning snapshot comes back having LOST a change order we already had — a CO#
# never really drops, so that means the fetch failed or returned a stale doc — we
# keep these from what we knew rather than wiping them to the failed fetch's
# blanks. Board fields (status, dates, price, assignee, …) still refresh.
_ENRICHMENT_KEEP = (
    "co_number", "so_pdf", "co_history",
    "so_design_desc", "so_size", "so_arrangement", "so_motor_pos", "so_class",
    "so_rotation", "so_discharge", "so_pct_width", "so_wheel_type",
    "so_design_temp", "so_max_temp", "so_special_temp",
    "line_items", "line_item_tags",
    "has_drive_run", "drive_run_pdf", "drive_run_count", "drive_run_rev",
    "drive_run", "drive_run_summary", "drive_run_template",
    "job_type", "job_folder", "dwg_extras", "dwg_missing_std",
)


def _keep_better_enrichment(stored: Dict[str, Any], incoming: Dict[str, Any]) -> Dict[str, Any]:
    """Guard against a regression: if `incoming` has a LOWER change order than
    `stored` (or the same CO# but a now-blank SO link), a re-fetch must have failed
    or returned a stale/original Sales Order — keep the prior enrichment and take
    only the fresh board fields. Otherwise use `incoming` unchanged."""
    s_co = int(stored.get("co_number") or 0)
    i_co = int(incoming.get("co_number") or 0)
    s_pdf = (stored.get("so_pdf") or "").strip()
    i_pdf = (incoming.get("so_pdf") or "").strip()
    regressed = i_co < s_co or (i_co == s_co and s_pdf and not i_pdf)
    if not regressed:
        return incoming
    merged = dict(incoming)
    for f in _ENRICHMENT_KEEP:
        if stored.get(f) not in (None, "", [], {}):
            merged[f] = stored[f]
    return merged


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
        # We only truly KNOW the add time if we watched it arrive (a genuine new
        # arrival, not one already on the board when the watch began — those are
        # carried over with an approximate time).
        known = not bool(j.get("_carried_over"))
        if entry is None:
            orders[jn] = {"added": added, "added_known": known,
                          "last_in": now_iso, "last_out": None,
                          "left": None, "on_queue": True, "seen_on_queue": True, "job": dict(j)}
        else:
            stored_job = entry.get("job") or {}
            # Never let a failed/stale re-fetch wipe a change order we already had.
            merged = _keep_better_enrichment(stored_job, j)
            for field, old, new in _diffs(stored_job, merged):
                events.append({"time": now_iso, "job": jn, "customer": _norm(j.get("customer")),
                               "field": field, "old": old, "new": new})
            if not entry.get("on_queue"):
                # Off-board -> on-board: it just (re)entered the queue, so this is
                # the moment that matters for "Added" — stamp last_in and treat the
                # add time as known (we watched this arrival ourselves).
                entry["last_in"] = now_iso
                entry["added_known"] = known
            entry["job"] = dict(merged)
            entry["on_queue"] = True
            entry["seen_on_queue"] = True   # the watcher has seen it on the board
            entry["left"] = None
            entry.setdefault("added", added)
            entry.setdefault("last_in", entry.get("added") or now_iso)  # migrate old entries
            entry.setdefault("last_out", None)
            entry.setdefault("added_known", known)

    for jn, entry in orders.items():
        if jn not in present_nums and entry.get("on_queue"):
            entry["on_queue"] = False
            entry["last_out"] = now_iso     # most recent departure (kept across returns)
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
        # A merged order has never been on the board (the watcher didn't see it
        # arrive or leave), so left=None — it must NOT look "removed from the
        # queue". Only live_master.update marks an order on/off the board.
        entry = orders[jn] = {"added": now_iso, "left": None, "on_queue": False, "job": {"job": jn}}
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
    """The job dicts currently on the board, oldest-added first. Each carries when
    it most recently came onto the board (`_added_iso` = last_in, so a returning
    order shows today's entry rather than its all-time first sighting), whether
    that time is known (`_added_known`), and its last departure (`_last_out`), so
    the Live Queue 'Added' column can show a real time/date or 'NO DATA'."""
    out = []
    for _, e in ordered(master):
        if e.get("on_queue") and e.get("job"):
            j = dict(e["job"])
            j["_added_iso"] = e.get("last_in") or e.get("added")
            j["_added_known"] = e.get("added_known", False)
            j["_last_out"] = e.get("last_out")
            out.append(j)
    return out
