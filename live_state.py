"""Persistent intraday state for the live queue watcher (watch.py).

The daily run (main.py) is a once-a-day photo: scrape the whole board, enrich
every order, diff, email. The live watcher is the opposite shape — it polls the
board often (every couple of minutes, all day) but must stay cheap, so it does
the expensive per-order enrichment (open the detail modal, download + parse the
Sales Order / quote run, scan the AutoCAD folder) ONCE per order, the first time
that order number appears. This module is the memory that makes that possible.

State shape (one entry per order number we've seen today):

    {
      "<job#>": {
        "first_seen":  ISO-8601 timestamp,   # when WE first saw it on the board
        "enriched":    bool,                  # has the slow enrichment run yet?
        "present":     bool,                  # on the board as of the last poll?
        "last_seen":   ISO-8601 timestamp,    # most recent poll it was present
        "job":         { ...full job dict },  # board fields + (once enriched) SO fields
      },
      ...
    }

"first_seen" is the time the order was *added* (to the precision of the poll
interval — it's the first poll we observed it, not a server-side field, since the
board doesn't expose one). It is set once and never moved, so an order that
leaves and comes back the same day keeps its original add time.

Everything here is pure dict/JSON logic with no browser or Office dependency, so
it is unit-tested directly (test_live_state.py) and safe to import anywhere.
"""
from __future__ import annotations

import json
import logging
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List

from config import SNAPSHOT_DIR
from process_lock import data_file_lock
from scraper import _design_from_item  # only a pure string helper; no browser
from sales_order_validation import (
    SALES_ORDER_DERIVED_FIELDS,
    clear_sales_order_data,
    effective_sales_order_invalidation,
    is_order_verification_record,
    is_true_sales_order_record,
)

log = logging.getLogger(__name__)

# Board-level fields produced by scraper.scrape_queue. These can change through
# the day (status, dates, price, flags), so each poll refreshes them on the
# stored job dict while leaving the enrichment fields (so_*, drive_run_*, …)
# untouched. Kept explicit rather than "everything in the fresh dict" so a future
# scraper field can't silently clobber an enrichment field of the same name.
BOARD_FIELDS = (
    "status", "job", "oper", "item", "design", "assigned_to", "checker",
    "start_date", "end_date", "plan_hrs", "fannet_date", "total_price",
    "status_note", "customer", "primary_rep", "ship_with",
    "unapproved", "credit_hold", "has_notes",
)


def state_path(d: date) -> Path:
    """Per-day state file, so a watcher restart mid-day resumes where it left
    off rather than re-enriching (and re-toasting) the whole board."""
    return SNAPSHOT_DIR / f"live_state_{d.isoformat()}.json"


def load_state(d: date) -> Dict[str, Any]:
    """Today's live state, or an empty dict on first run / unreadable file."""
    path = state_path(d)
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (json.JSONDecodeError, OSError) as e:
        log.warning("Could not read %s (%s); starting live state fresh", path, e)
        return {}


def save_state(state: Dict[str, Any], d: date) -> None:
    """Write the state atomically (temp file + replace) so a crash mid-write
    can't leave a half-written file that the next poll would discard."""
    path = state_path(d)
    with data_file_lock(path, label="live-state data update"):
        external = load_state(d)
        for job, external_entry in external.items():
            current_entry = state.get(job)
            if not isinstance(current_entry, dict) or not isinstance(external_entry, dict):
                continue
            current_job = current_entry.get("job") or {}
            external_job = external_entry.get("job") or {}
            invalidated_at = effective_sales_order_invalidation(
                external_job, current_job
            )
            if (
                is_order_verification_record(current_job)
                or (
                    is_order_verification_record(external_job)
                    and not is_true_sales_order_record(current_job)
                )
            ):
                invalidated_at = max(
                    invalidated_at,
                    str(external_job.get("so_invalidated_at") or ""),
                    str(current_job.get("so_invalidated_at") or ""),
                    datetime.now().isoformat(timespec="seconds"),
                )
            if invalidated_at:
                removed_stale_data = clear_sales_order_data(
                    current_job, invalidated_at
                )
                current_entry["job"] = current_job
                if removed_stale_data:
                    current_entry["enriched"] = False
        tmp = path.with_suffix(".json.tmp")
        tmp.write_text(json.dumps(state, indent=2, default=str), encoding="utf-8")
        tmp.replace(path)


def _job_number(j: Dict[str, Any]) -> str:
    return str(j.get("job") or "").strip()


def new_job_numbers(state: Dict[str, Any], board: List[Dict[str, Any]]) -> List[str]:
    """Order numbers on the board that we have never seen before today — the ones
    that still need the slow enrichment. Order is preserved (board order)."""
    out: List[str] = []
    for j in board:
        jn = _job_number(j)
        if jn and jn not in state and jn not in out:
            out.append(jn)
    return out


def seed_from_snapshot(
    state: Dict[str, Any],
    snapshot_jobs: List[Dict[str, Any]],
    seen_at: str,
) -> int:
    """Pre-load orders already on the board at start-of-day (from the morning
    snapshot the 5 AM run wrote) so they are NOT treated as new arrivals — only
    orders that show up *after* the watcher starts get enriched and announced.

    These come from the daily run already fully enriched, so they go in with
    enriched=True. Their first_seen is the morning baseline timestamp, flagged
    `carried_over` so the report can show "before today's watch" rather than a
    misleadingly precise time. Returns how many were seeded. Never overwrites an
    entry already in the state (a restart re-seeds harmlessly)."""
    seeded = 0
    for j in snapshot_jobs:
        jn = _job_number(j)
        if not jn or jn in state:
            continue
        job = dict(j)
        job.setdefault("design", _design_from_item(job.get("item", "")))
        invalidated = effective_sales_order_invalidation(job)
        state[jn] = {
            "first_seen": seen_at,
            "carried_over": True,   # was already in the queue when the watch began
            "enriched": not bool(invalidated),
            "present": True,
            "last_seen": seen_at,
            "job": job,
        }
        seeded += 1
    return seeded


def record_poll(
    state: Dict[str, Any],
    board: List[Dict[str, Any]],
    now: datetime,
    baseline: bool = False,
) -> Dict[str, List[str]]:
    """Fold one board scrape into the state. Returns the deltas this poll:

        {"new": [...], "returning": [...], "removed": [...]}

    - new:       order numbers seen for the very first time today.
    - returning: previously-seen orders that had dropped off and are back.
    - removed:   orders that were present last poll and are now gone.

    New/returning orders are inserted with their board fields and enriched=False
    (watch.py runs the slow enrichment for them, then calls mark_enriched). For
    orders still present, the volatile board fields are refreshed in place so
    intraday changes (a new End Date, a price bump, a credit hold) show up
    without re-enriching. first_seen is never moved once set.

    `baseline=True` is for the watcher's very first poll, which establishes the
    start-of-day picture: orders found then were already in the queue, so they're
    marked carried_over (no precise add time, sorted to the bottom, not
    announced). They're still enriched=False so the watcher fills them in once.
    """
    now_iso = now.isoformat(timespec="seconds")
    board_nums = {_job_number(j) for j in board if _job_number(j)}

    new: List[str] = []
    returning: List[str] = []
    for j in board:
        jn = _job_number(j)
        if not jn:
            continue
        entry = state.get(jn)
        if entry is None:
            state[jn] = {
                "first_seen": now_iso,
                "carried_over": baseline,
                "enriched": False,
                "present": True,
                "last_seen": now_iso,
                "job": dict(j),
            }
            new.append(jn)
        else:
            was_gone = not entry.get("present", True)
            # Refresh only the board-level fields; keep enrichment fields intact.
            stored = entry.setdefault("job", {})
            for f in BOARD_FIELDS:
                if f in j:
                    stored[f] = j[f]
            entry["present"] = True
            entry["last_seen"] = now_iso
            if was_gone:
                returning.append(jn)

    removed: List[str] = []
    for jn, entry in state.items():
        if jn not in board_nums and entry.get("present", False):
            entry["present"] = False
            removed.append(jn)

    return {"new": new, "returning": returning, "removed": removed}


def mark_enriched(state: Dict[str, Any], enriched_jobs: List[Dict[str, Any]]) -> None:
    """Store the fully-enriched job dicts (from enrich_with_sales_orders) back on
    their state entries and flag them enriched, so later polls don't redo the
    slow work. Board fields keep refreshing on top of these each poll."""
    for j in enriched_jobs:
        jn = _job_number(j)
        entry = state.get(jn)
        if entry is None:
            continue
        stored = entry.get("job") or {}
        merged = dict(j)
        recovering_true_so = (
            is_order_verification_record(stored)
            and is_true_sales_order_record(merged)
            and not effective_sales_order_invalidation(stored, merged)
        )
        for field, stored_value in stored.items():
            protects_enrichment = (
                field.startswith("so_") or field in ("line_items", "line_item_tags")
            )
            if recovering_true_so and field in SALES_ORDER_DERIVED_FIELDS:
                continue
            if (
                protects_enrichment
                and stored_value not in (None, "", [], {})
                and merged.get(field) in (None, "", [], {})
            ):
                merged[field] = stored_value
        if (
            stored.get("so_special_temp") not in (None, "", "0")
            and str(merged.get("so_special_temp") or "") == "0"
            and not merged.get("so_design_temp")
            and not merged.get("so_max_temp")
        ):
            merged["so_special_temp"] = stored["so_special_temp"]
        invalidated_at = effective_sales_order_invalidation(stored, merged)
        if (
            is_order_verification_record(merged)
            or (
                is_order_verification_record(stored)
                and not is_true_sales_order_record(merged)
            )
        ):
            invalidated_at = max(
                invalidated_at,
                str(stored.get("so_invalidated_at") or ""),
                str(merged.get("so_invalidated_at") or ""),
                datetime.now().isoformat(timespec="seconds"),
            )
        if invalidated_at:
            clear_sales_order_data(merged, invalidated_at)
        entry["job"] = merged
        entry["enriched"] = not is_order_verification_record(j)


def present_jobs(state: Dict[str, Any]) -> List[Dict[str, Any]]:
    """The job dicts currently on the board, each carrying its _first_seen /
    _carried_over markers so the live workbook can show the add time and sort by
    it. Newest arrivals first; carried-over (start-of-day) orders sort last."""
    rows: List[Dict[str, Any]] = []
    for entry in state.values():
        if not entry.get("present", False):
            continue
        job = dict(entry.get("job", {}))
        job["_first_seen"] = entry.get("first_seen", "")
        job["_carried_over"] = entry.get("carried_over", False)
        rows.append(job)
    # Sort: real arrivals by first_seen descending (newest on top); carried-over
    # orders (no precise add time) fall to the bottom in job-number order.
    def _key(j: Dict[str, Any]):
        carried = j.get("_carried_over", False)
        return (0 if not carried else 1,
                _neg_iso(j.get("_first_seen", "")) if not carried else "",
                j.get("job", ""))
    rows.sort(key=_key)
    return rows


def _neg_iso(iso: str) -> str:
    """A sort key that orders ISO timestamps newest-first as plain strings."""
    # Invert each digit so lexical ascending == chronological descending, which
    # lets present_jobs sort newest-first without parsing every timestamp.
    table = str.maketrans("0123456789", "9876543210")
    return (iso or "").translate(table)
