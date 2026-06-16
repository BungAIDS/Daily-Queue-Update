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
import logging
import sys
import time
from datetime import date, datetime, time as dtime, timedelta

import line_items
import live_sheets
import live_state
import notify
from compare import diff_queues, load_history, load_latest_snapshot, load_snapshot
from config import (LIVE_MORNING_SNAPSHOT, LIVE_WORKBOOK_PATH, OUTPUT_DIR,
                    POLL_INTERVAL_SECONDS, WATCH_END, WATCH_START,
                    validate_runtime_config)
from live_excel import save_morning_copy, update_workbook
from live_sheets import added_label
from sales_orders import enrich_with_sales_orders
from scraper import scrape_queue

log = logging.getLogger("queue-watch")


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


def _build_sheets(state: dict, present: list, today: date, now: datetime) -> list:
    """Assemble the four master tabs from the current board + on-disk baselines.
    All present orders are already enriched, so the full Full-Queue data is in
    hand — these builders just shape it (live_sheets is pure/tested)."""
    # Highlight every order that arrived during today's watch (not carried over).
    new_ids = {jn for jn, e in state.items()
               if e.get("present") and not e.get("carried_over")}

    # Changes since this morning's frozen baseline, and vs the previous run.
    baseline = live_state.load_baseline(today)
    intraday = diff_queues(present, baseline, today, persist_history=False, prev_date=today)
    yest, yest_date = load_latest_snapshot(today)
    yesterday = diff_queues(present, yest, today, persist_history=False, prev_date=yest_date)
    co_changed_ids = ({c["job"] for c in intraday.get("co_changed", [])}
                      | {c["job"] for c in yesterday.get("co_changed", [])})

    store = line_items.load_store()
    return [
        live_sheets.full_queue_sheet(present, today, new_ids=new_ids,
                                     co_changed_ids=co_changed_ids, ref=now),
        live_sheets.changes_sheet(
            intraday, f"{today.isoformat()} (start of day)",
            yesterday, yest_date.isoformat() if yest_date else "no prior run"),
        live_sheets.history_sheet(load_history()),
        live_sheets.line_items_sheet(store, order_nums=[j.get("job") for j in present]),
    ]


def poll_once(state: dict, now: datetime, baseline: bool, announce: bool) -> dict:
    """One cycle: scrape -> record -> enrich new -> write the master tabs -> notify.
    Mutates `state`; returns the deltas dict from record_poll."""
    board = scrape_queue(headless=True)
    if not board:
        log.warning("Board scrape returned 0 orders — skipping this cycle (site/session issue?).")
        return {"new": [], "returning": [], "removed": []}

    deltas = live_state.record_poll(state, board, now, baseline=baseline)
    log.info("Poll @ %s: %d on board | new=%d returning=%d removed=%d",
             now.strftime("%H:%M:%S"), len(board),
             len(deltas["new"]), len(deltas["returning"]), len(deltas["removed"]))

    _enrich_pending(state)
    present = live_state.present_jobs(state)

    # Freeze the start-of-day baseline (enriched) on the first poll, so the
    # 'changes since this morning' view has stable morning values to diff against.
    if baseline:
        live_state.save_baseline(present, today=now.date())

    if LIVE_WORKBOOK_PATH:
        update_workbook(_build_sheets(state, present, now.date(), now), LIVE_WORKBOOK_PATH)
    else:
        log.warning("LIVE_WORKBOOK_PATH not set — live workbook not updated. "
                    "Set it in .env to the co-authored workbook's local path.")

    if announce and deltas["new"]:
        fresh = [j for j in present if j.get("job") in set(deltas["new"])]
        for j in fresh:
            j["_added_label"] = added_label(j, ref=now)
        notify.notify_new_orders(fresh)

    live_state.save_state(state, now.date())
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
    first = not state
    if first:
        _seed_baseline(state, datetime.now())

    log.info("=== Live watch start: %s  window %s-%s  every %ds ===",
             today.isoformat(), start.strftime("%H:%M"), end.strftime("%H:%M"),
             POLL_INTERVAL_SECONDS)

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
                poll_once(state, now, baseline=baseline, announce=not baseline)
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
        first = not state
        if first:
            _seed_baseline(state, datetime.now())
        poll_once(state, datetime.now(), baseline=first, announce=not first)
        if first and LIVE_MORNING_SNAPSHOT:
            _morning_snapshot(datetime.now())
        return 0

    return run_watch(ignore_window=args.now)


if __name__ == "__main__":
    sys.exit(main())
