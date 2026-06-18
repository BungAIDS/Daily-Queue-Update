"""Live intraday queue watcher — the all-day companion to the 5 AM daily run.

    python watch.py            # run the daytime watch loop (5am-5pm by default)
    python watch.py --once     # do a single poll cycle now, then exit (testing)
    python watch.py --now      # ignore the time window; start polling immediately

The daily run (main.py) is a once-a-day photo. This keeps the queue *live*: every
POLL_INTERVAL_SECONDS it does the cheap board scrape (order numbers + row data,
no clicking into anything), and runs the slow enrichment (open the detail, pull
the Sales Order + quote run, scan the AutoCAD folder) ONLY for orders that are
new since it last looked. Each order's first-seen time (the time it was added) is
recorded, the current board is written into your co-authored Excel workbook in
real time (so coworkers see it update live), and new orders fire a Windows toast
plus a Microsoft Teams card.

How a day runs:
  - First poll establishes the start-of-day baseline silently (seeded from the
    morning daily-run snapshot when available, so the whole board isn't
    re-enriched), and saves a dated "morning snapshot" copy of the workbook.
  - Every poll after that announces only genuinely new arrivals.
  - State lives in SNAPSHOT_DIR/live_state_<date>.json, so a restart mid-day
    resumes without re-enriching or re-announcing what it already saw.

Best-effort throughout: a failed scrape/Excel/notify cycle is logged and the
loop carries on — a watcher that dies on one hiccup is worse than one that skips
a beat. Run it from your logged-in Windows session (Excel + Outlook/Teams), the
same place the daily run lives. Configure it in .env (see .env.example).
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
import time
from datetime import date, datetime, time as dtime, timedelta

import change_log
import line_items
import live_excel
import live_master
import live_sheets
import live_state
import notify
from compare import load_latest_snapshot, load_snapshot
from config import (LIVE_MORNING_SNAPSHOT, LIVE_WORKBOOK_PATH, OUTPUT_DIR,
                    POLL_INTERVAL_SECONDS, WATCH_END, WATCH_START,
                    validate_runtime_config)
from live_excel import save_morning_copy, update_master_workbook
from live_sheets import added_label
from runstate import load_diff
from sales_orders import enrich_with_sales_orders, refresh_autocad_folders
from scraper import scrape_queue

log = logging.getLogger("queue-watch")

# Bump to force a one-time clean rebuild of the Order History tab (e.g. after a
# layout change). In normal operation the tab is built once and only appended to.
OH_BUILD_VERSION = 3


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )


def _window_today(today: date) -> "tuple[datetime, datetime]":
    start = datetime.combine(today, dtime(*WATCH_START))
    end = datetime.combine(today, dtime(*WATCH_END))
    return start, end


def _seed_baseline(state: dict, today: datetime) -> None:
    """Pre-load the start-of-day board from the morning daily-run snapshot so the
    watcher doesn't re-enrich orders the 5 AM run already covered. Today's
    snapshot is preferred; a recent prior one is a reasonable fallback."""
    snap = load_snapshot(today.date())
    label = "today's"
    if snap is None:
        snap, prev = load_latest_snapshot(today.date())
        label = f"the {prev}" if prev else "no"
    if snap:
        n = live_state.seed_from_snapshot(state, snap, today.isoformat(timespec="seconds"))
        log.info("Seeded %d order(s) from %s daily snapshot as the start-of-day baseline.", n, label)
    else:
        log.info("No prior daily snapshot to seed from — the first poll's board "
                 "becomes the baseline (enriched once, not announced).")


def _enrich_pending(state: dict) -> list:
    """Run the slow Sales-Order enrichment for every present order not yet
    enriched, then fold the results back into the state. Returns the enriched
    job dicts."""
    pending = [e["job"] for e in state.values()
               if e.get("present") and not e.get("enriched") and e.get("job")]
    if not pending:
        return []
    log.info("Enriching %d new order(s) (Sales Order + quote run + AutoCAD folder)...", len(pending))
    try:
        enrich_with_sales_orders(pending)
    except Exception:  # noqa: BLE001 - never let enrichment sink the loop
        log.exception("Enrichment failed this cycle; will retry these next poll")
        return []
    live_state.mark_enriched(state, pending)
    return pending


def _force_rebuild(master: dict) -> None:
    """Live Queue is small and is rebuilt fresh each process start, so drop its
    row signatures. Order History is NOT touched here — it's built once and then
    only appended to (its oh_sigs persist across runs), so a restart never wipes
    the ~12K-row tab."""
    master["lq_sigs"] = {}
    master.pop("below_sig", None)   # force the 'removed' block to redraw on (re)start
    # One-time cleanup: a backlog order the watcher never saw on the board must
    # have no 'left' time (earlier merges wrongly stamped one, which made the
    # whole backlog look "removed today").
    for e in master.get("orders", {}).values():
        if not e.get("seen_on_queue"):
            e["left"] = None


def _plan(records: list, sig_store: dict, allow_delete: bool) -> list:
    """Turn (key, cells) records into upsert ops vs what we last wrote (the sigs
    in `sig_store`, a flat {order#: sig} persisted in the master), and update the
    store. Flat (not per-order) so backlog orders that live only in the line-items
    store are tracked too. Change detection survives restarts."""
    desired = [(key, live_sheets.row_sig(cells), cells) for key, cells in records]
    ops = live_sheets.plan_upsert(desired, dict(sig_store), allow_delete=allow_delete)
    sig_by = {key: sig for key, sig, _ in desired}
    for kind, key, _ in ops:
        if kind in ("append", "update"):
            sig_store[key] = sig_by[key]
        elif kind == "delete":
            sig_store.pop(key, None)
    return ops


def _oh_orders(master: dict, store: dict) -> list:
    """Every order for the Order History log: the live master entries, plus the
    backlog orders that exist only in the line-items store (~12K once
    line_items_scan.py has run). Chronological by when first seen."""
    merged = dict(master.get("orders", {}))
    for jn, rec in (store.get("jobs") or {}).items():
        if jn in merged:
            continue
        merged[jn] = {
            "on_queue": False, "added": rec.get("scanned_at"), "left": rec.get("scanned_at"),
            "job": {"job": jn, "customer": rec.get("customer", ""),
                    "co_number": rec.get("co_number") or 0, "so_pdf": rec.get("so_pdf", ""),
                    "line_items": rec.get("items") or []},
        }
    return sorted(merged.items(), key=lambda kv: (str(kv[1].get("added") or ""), kv[0]))


def _changes_sheet(master: dict, lq_jobs: list, new_today: set, today: date,
                   now: datetime) -> "live_sheets.Sheet":
    """Today's activity log: new orders today, change orders (CO#), the
    field-modification log, and orders removed today. `now` stamps a 'last
    updated' line at the top so users can see the tab is live."""
    new_today_jobs = [j for j in lq_jobs if str(j.get("job") or "") in new_today]
    events = change_log.load(today)
    removed_today = [e.get("job", {}) for e in master.get("orders", {}).values()
                     if e.get("seen_on_queue") and not e.get("on_queue")
                     and str(e.get("left") or "")[:10] == today.isoformat()]
    updated_at = live_sheets.fmt_datetime(now)
    return live_sheets.changes_sheet(new_today_jobs, events, removed_today,
                                     today.isoformat(), updated_at=updated_at)


def _new_today_ids(lq_jobs: list, today: date) -> set:
    """Order numbers that are 'new as of today'. Uses the 5 AM main.py run's own
    output — today's snapshot + diff — so it works whenever you start the watcher
    later in the day: an order is new today if main.py flagged it new/returning
    this morning, OR it has appeared on the board since the morning snapshot.
    Falls back to the most recent prior snapshot if main.py hasn't run today."""
    board_ids = {str(j.get("job")) for j in lq_jobs if j.get("job")}
    today_snap = load_snapshot(today)
    if today_snap is None:
        prev, _ = load_latest_snapshot(today)
        if prev is None:
            return set()
        prev_ids = {str(j.get("job")) for j in prev if j.get("job")}
        return {jn for jn in board_ids if jn not in prev_ids}

    ids: set = set()
    d = load_diff(today)
    if d:
        for grp in ("new", "returning"):
            ids |= {str(j.get("job")) for j in (d.get(grp) or []) if j.get("job")}
    snap_ids = {str(j.get("job")) for j in today_snap if j.get("job")}
    ids |= {jn for jn in board_ids if jn not in snap_ids}   # arrived since the morning snapshot
    return ids & board_ids


def _render_master(master: dict, now: datetime, board_order: list | None = None) -> None:
    """Build Live Queue (incremental upsert) + Order History (matrix log) +
    the Changes snapshot from the master log + line-items store, and push them in.
    `board_order` is the order numbers as they appear on cbcinsider, so the Live
    Queue can match that order."""
    today = now.date()
    lq_jobs = live_master.on_queue(master)

    # Tag each on-board order with its cbcinsider position and order the rows by
    # it (the sheet is also sorted by the "#" column to keep that order).
    pos = {str(jn): i + 1 for i, jn in enumerate(board_order or [])}
    for j in lq_jobs:
        j["_cbc_pos"] = pos.get(str(j.get("job")))
    lq_jobs.sort(key=lambda j: (j.get("_cbc_pos") is None, j.get("_cbc_pos") or 0))

    new_today = _new_today_ids(lq_jobs, today)
    # Orders that had a change order (CO#) land today -> their text goes red.
    co_changed = {str(e.get("job")) for e in change_log.load(today)
                  if e.get("field") == "CO#" and e.get("job")}

    lq_sigs = master.setdefault("lq_sigs", {})
    lq_ops = _plan(live_sheets.live_queue_records(lq_jobs, today, new_ids=new_today,
                                                  co_changed_ids=co_changed, ref=now),
                   lq_sigs, allow_delete=True)
    lq_payload = {"name": "Live Queue", "headers": live_sheets.LIVE_QUEUE_HEADERS, "ops": lq_ops,
                  "key_col": live_sheets.LIVE_QUEUE_KEY_COL, "allow_delete": True,
                  "sort_col": live_sheets.LIVE_QUEUE_CBC_COL, "text_cols": [1],  # Added (col 1) -> AM/PM text
                  "freeze": "C2"}   # header on row 1, frozen with the Added + Job # columns

    # "Removed since this morning" block below the Live Queue. Render it only when
    # the set actually changes (a removal/return), not every cycle. Only orders
    # the watcher actually saw on the board count — never the merged backlog.
    removed = [(e.get("job", {}), e.get("left")) for e in master.get("orders", {}).values()
               if e.get("seen_on_queue") and not e.get("on_queue")
               and str(e.get("left") or "")[:10] == today.isoformat()]
    removed.sort(key=lambda jl: str(jl[1] or ""), reverse=True)   # most recently removed first
    # Keep the key column (LIVE_QUEUE_KEY_COL = col 2, Job #) blank in this block
    # so the key-based "last live data row" lookup in _render_below — and the next
    # poll's keymap — skip these rows. The order numbers go in col 1.
    below = {"title": "Removed from the queue since this morning",
             "headers": ["Job #", "", "Customer", "Design", "Removed"],
             "rows": [[j.get("job", ""), "", j.get("customer", ""),
                       j.get("so_design_desc") or j.get("design", ""),
                       live_sheets.fmt_time(left)] for j, left in removed]}
    below_sig = json.dumps(below, default=str, sort_keys=True)
    if below_sig != master.get("below_sig"):
        lq_payload["below"] = below
        master["below_sig"] = below_sig

    # Order History: live master + the whole line-items backlog, as a matrix log.
    # Built ONCE then only appended to; we rebuild only on a build-format bump
    # (one-time migration) or when the matrix columns actually grow.
    spec = live_sheets.order_history_build(_oh_orders(master, line_items.load_store()), today,
                                           prev_columns=master.get("oh_columns"))
    rebuild = (master.get("oh_build_version") != OH_BUILD_VERSION
               or master.get("oh_columns") != spec["columns"])
    if rebuild:
        master["oh_sigs"] = {}
        master["oh_columns"] = spec["columns"]
        master["oh_headers"] = spec["headers"]
        master["oh_build_version"] = OH_BUILD_VERSION
    oh_sigs = master.setdefault("oh_sigs", {})
    oh_ops = _plan(spec["records"], oh_sigs, allow_delete=False)
    oh_payload = {"name": "Order History", "spec": spec, "ops": oh_ops, "rebuild": rebuild,
                  "key_col": live_sheets.ORDER_HISTORY_KEY_COL, "freeze": "B2"}  # pin Job # only

    update_master_workbook(LIVE_WORKBOOK_PATH, lq_payload, oh_payload,
                           changes_sheet=_changes_sheet(master, lq_jobs, new_today, today, now))


def poll_once(state: dict, master: dict, now: datetime, baseline: bool, announce: bool) -> dict:
    """One cycle: scrape -> record -> enrich new -> upsert the master tabs -> notify.
    Mutates `state` and `master`; returns the deltas dict from record_poll."""
    board = scrape_queue(headless=True)
    if not board:
        log.warning("Board scrape returned 0 orders — skipping this cycle (site/session issue?).")
        return {"new": [], "returning": [], "removed": []}

    deltas = live_state.record_poll(state, board, now, baseline=baseline)
    log.info("Poll @ %s: %d on board | new=%d returning=%d removed=%d",
             now.strftime("%H:%M:%S"), len(board),
             len(deltas["new"]), len(deltas["returning"]), len(deltas["removed"]))

    _enrich_pending(state)
    # Re-resolve AutoCAD folders for any on-board order still without one (cheap
    # now), so orders an earlier lookup missed ("Open") get filled in without a
    # full re-enrich. Mutates the state entries' job dicts in place so it sticks.
    unresolved = [e["job"] for e in state.values()
                  if e.get("present") and e.get("job") and not e["job"].get("job_type")]
    if unresolved:
        resolved = refresh_autocad_folders(unresolved)
        if resolved:
            log.info("Resolved %d AutoCAD folder(s) for order(s) that showed 'Open'.", resolved)

    present = live_state.present_jobs(state)
    # Fold the board into the master log; log any field modifications it found.
    events = live_master.update(master, present, now)
    change_log.append(now.date(), events)
    if events:
        log.info("Logged %d field change(s) this poll.", len(events))

    if LIVE_WORKBOOK_PATH:
        _render_master(master, now, board_order=[str(j.get("job")) for j in board if j.get("job")])
    else:
        log.warning("LIVE_WORKBOOK_PATH not set — live workbook not updated. "
                    "Set it in .env to the co-authored workbook's local path.")

    if announce and deltas["new"]:
        fresh = [j for j in present if j.get("job") in set(deltas["new"])]
        for j in fresh:
            j["_added_label"] = added_label(j, ref=now)
        notify.notify_new_orders(fresh)

    live_state.save_state(state, now.date())
    live_master.save_master(master)
    return deltas


def _morning_snapshot(now: datetime) -> None:
    if not (LIVE_MORNING_SNAPSHOT and LIVE_WORKBOOK_PATH):
        return
    dest = OUTPUT_DIR / "live_snapshots" / f"live_queue_{now.date().isoformat()}.xlsx"
    if dest.exists():
        return  # already froze this morning's copy (e.g. a restart)
    save_morning_copy(LIVE_WORKBOOK_PATH, dest)


def run_watch(ignore_window: bool = False) -> int:
    setup_logging()
    validate_runtime_config()
    if not LIVE_WORKBOOK_PATH:
        log.warning("LIVE_WORKBOOK_PATH is not set in .env — the watcher will run "
                    "and notify, but won't write a live workbook until you set it.")

    today = date.today()
    start, end = _window_today(today)
    now = datetime.now()
    if not ignore_window:
        if now < start:
            wait = (start - now).total_seconds()
            log.info("Before the watch window; sleeping %.0f min until %s.",
                     wait / 60, start.strftime("%H:%M"))
            time.sleep(wait)
        elif now >= end:
            log.info("Past today's watch window (%s). Nothing to do; exiting.",
                     end.strftime("%H:%M"))
            return 0

    state = live_state.load_state(today)
    master = live_master.load_master()
    _force_rebuild(master)            # process start -> rebuild the upsert tabs once
    first = not state
    if first:
        _seed_baseline(state, datetime.now())

    log.info("=== Live watch start: %s  window %s-%s  every %ds  (%d orders in master log) ===",
             today.isoformat(), start.strftime("%H:%M"), end.strftime("%H:%M"),
             POLL_INTERVAL_SECONDS, len(master.get("orders", {})))

    cycle = 0
    try:
        while True:
            now = datetime.now()
            if not ignore_window and now >= end:
                log.info("Reached end of watch window (%s); stopping.", end.strftime("%H:%M"))
                break
            baseline = first and cycle == 0
            t0 = time.monotonic()
            try:
                poll_once(state, master, now, baseline=baseline, announce=not baseline)
            except Exception:  # noqa: BLE001 - one bad cycle must not end the watch
                log.exception("Poll cycle failed; continuing to the next one")
            if baseline:
                _morning_snapshot(now)
            cycle += 1
            elapsed = time.monotonic() - t0
            time.sleep(max(0.0, POLL_INTERVAL_SECONDS - elapsed))
    except KeyboardInterrupt:
        log.info("Interrupted — saving state and exiting.")
        live_state.save_state(state, today)
        live_master.save_master(master)
    log.info("=== Live watch done (%d cycles) ===", cycle)
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="Live intraday queue watcher.")
    ap.add_argument("--once", action="store_true",
                    help="Run a single poll cycle now and exit (testing).")
    ap.add_argument("--now", action="store_true",
                    help="Ignore the daily time window; start polling immediately.")
    args = ap.parse_args()

    if args.once:
        setup_logging()
        validate_runtime_config()
        today = date.today()
        state = live_state.load_state(today)
        master = live_master.load_master()
        _force_rebuild(master)
        first = not state
        if first:
            _seed_baseline(state, datetime.now())
        poll_once(state, master, datetime.now(), baseline=first, announce=not first)
        if first and LIVE_MORNING_SNAPSHOT:
            _morning_snapshot(datetime.now())
        return 0

    return run_watch(ignore_window=args.now)


if __name__ == "__main__":
    sys.exit(main())
