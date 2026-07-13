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
import os
import signal
import sys
import threading
import time
from datetime import date, datetime, time as dtime, timedelta
from logging.handlers import TimedRotatingFileHandler

import change_log
import engineers
import git_update
import line_items
import live_master
import live_sheets
import live_state
import notify
import stop_signal
from compare import load_latest_snapshot, load_snapshot, save_snapshot
from config import (DATA_PUSH_BRANCH, DATA_PUSH_ON_CHANGE, EXCEL_RECYCLE_EVERY_POLLS,
                    LIVE_MORNING_SNAPSHOT, LIVE_WORKBOOK_PATH, LOG_DIR,
                    LOG_PUSH_BRANCH, OUTPUT_DIR, POLL_INTERVAL_SECONDS,
                    SO_REVERIFY_MIN_AGE_MIN, SO_REVERIFY_PER_POLL, WATCH_END,
                    WATCH_START, validate_runtime_config)
from live_excel import (recycle_workbook, save_morning_copy,
                        update_master_workbook)
from live_sheets import added_label
from log_push import push_logs
from runstate import load_diff
from sales_orders import (enrich_with_sales_orders, repair_missing_sales_order_summaries,
                          refresh_autocad_folders, refresh_sales_orders)
from scraper import scrape_queue

log = logging.getLogger("queue-watch")

# Bump to force a one-time clean rebuild of the Order History tab (e.g. after a
# layout change). In normal operation the tab is built once and only appended to.
OH_BUILD_VERSION = 5   # rebuild once so the fixed ✓/red matrix conditional formatting lands


def setup_logging() -> None:
    """Log to the console AND to a daily-rotated file under LOG_DIR (keeping ~a
    week), so the output survives the window closing and there's a record to chase
    a bug through after the fact."""
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
    handlers: list[logging.Handler] = [logging.StreamHandler(sys.stdout)]
    try:
        fileh = TimedRotatingFileHandler(
            LOG_DIR / "watch.log", when="midnight", backupCount=7, encoding="utf-8")
        handlers.append(fileh)
    except OSError as e:  # a bad/locked log dir must never stop the watcher
        print(f"(could not open log file in {LOG_DIR}: {e}; console only)", file=sys.stderr)
    for h in handlers:
        h.setFormatter(fmt)
    logging.basicConfig(level=logging.INFO, handlers=handlers)
    log.info("Logging to console and %s (rotated daily, ~7 days kept).", LOG_DIR / "watch.log")


def _publish_logs() -> None:
    """Flush the log to disk and publish it to the debug branch (best-effort), so
    it can be read remotely without copying files off the machine."""
    for h in logging.getLogger().handlers:
        try:
            h.flush()
        except Exception:  # noqa: BLE001
            pass
    if push_logs():
        log.info("Published logs to the debug branch.")


def _publish_data() -> None:
    """Publish the order data to its branch (best-effort) so a remote reader
    tracks new orders as the watch gathers them. Only when auto-publish is on."""
    if not (DATA_PUSH_ON_CHANGE and DATA_PUSH_BRANCH):
        return
    try:
        from data_push import push_data
        if push_data():
            log.info("Published order data to the '%s' branch.", DATA_PUSH_BRANCH)
    except Exception as e:  # noqa: BLE001 - publishing must never disturb the watch
        log.debug("data publish skipped (%s)", e)


def _merge_backlog_sources() -> None:
    """Absorb durable backfill stores before loading the watcher's master copy."""
    try:
        import master_sync

        counts = master_sync.run("backfill", "line_items")
        changed = sum(counts.values())
        if changed:
            log.info("Merged %d backlog update(s) into the live master at startup.", changed)
    except Exception as e:  # noqa: BLE001 - startup repair must not stop the watch
        log.warning("Could not merge backlog stores at watcher startup (%s)", e)


def _window_today(today: date) -> "tuple[datetime, datetime]":
    start = datetime.combine(today, dtime(*WATCH_START))
    end = datetime.combine(today, dtime(*WATCH_END))
    return start, end


def _is_repair_order(job: str) -> bool:
    """A repair order carries a trailing letter (e.g. 352366A, 354721D). These
    don't live under the standard AutoCAD job-folder layout, so we table them
    rather than re-sweeping the whole tree for them on every poll."""
    jn = str(job or "").strip()
    return bool(jn) and jn[-1].isalpha()


def _queue_stale_so_rechecks(state: dict, now: datetime) -> int:
    """Round-robin background re-verification: flag the few on-board orders we've
    gone longest without re-checking for a fresh Sales-Order fetch this poll, so a
    silently-stale SO (e.g. one left at an old revision by an earlier failed
    fetch) self-corrects within the hour instead of waiting for the next daily
    run. Bounded by SO_REVERIFY_PER_POLL; skips repair orders and ones re-checked
    within SO_REVERIFY_MIN_AGE_MIN. Returns how many were queued."""
    if SO_REVERIFY_PER_POLL <= 0:
        return 0
    cutoff = (now - timedelta(minutes=SO_REVERIFY_MIN_AGE_MIN)).isoformat(timespec="seconds")
    cands = []
    for jn, e in state.items():
        if not (e.get("present") and e.get("enriched") and e.get("job")):
            continue
        if _is_repair_order(jn):
            continue
        va = e.get("verified_at") or ""
        if va and va > cutoff:
            continue                       # re-checked recently enough
        cands.append((va, jn))             # "" (never verified) sorts first -> oldest
    cands.sort()
    queued = [jn for _, jn in cands[:SO_REVERIFY_PER_POLL]]
    for jn in queued:
        state[jn]["enriched"] = False      # _enrich_pending re-fetches it this poll
    return len(queued)


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


def _daily_snapshot(board: list, now: datetime) -> None:
    """Freeze the start-of-day board as today's queue snapshot
    (snapshots/queue_<date>.json) the first time the watcher sees the board each
    day — the baseline that 'new today' and the start-of-day seed compare against.
    This was the 5 AM main.py run's job; the watcher now owns it. Guarded on the
    snapshot's absence, so it writes once per day (the first run Windows Scheduler
    fires) and a later restart never clobbers the morning baseline."""
    today = now.date()
    if load_snapshot(today) is not None:
        return
    save_snapshot(board, today)   # logs "Saved snapshot: … (N jobs)"


def _enrich_pending(state: dict) -> list:
    """Run the slow Sales-Order enrichment for every present order not yet
    enriched, then fold the results back into the state. Returns the enriched
    job dicts."""
    # Enrichment mutates each job dict. Work on copies so mark_enriched can
    # compare the result with the last trusted values and preserve parser gaps.
    pending = [dict(e["job"]) for e in state.values()
               if e.get("present") and not e.get("enriched") and e.get("job")]
    if not pending:
        return []
    log.info("Enriching %d order(s) — new, repriced or re-verified "
             "(Sales Order + quote run + AutoCAD folder)...", len(pending))
    try:
        # deep_folders=False: skip the full-tree folder sweep intraday — a
        # standard folder is still found by direct lookup, the daily run catches
        # the rest — so a folderless job doesn't cost a ~12K-folder walk each poll.
        enrich_with_sales_orders(pending, deep_folders=False)
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
    master.pop("below_sig", None)   # legacy: the 'removed' block is now passed every cycle
    # One-time cleanup: a backlog order the watcher never saw on the board must
    # have no 'left' time (earlier merges wrongly stamped one, which made the
    # whole backlog look "removed today").
    for e in master.get("orders", {}).values():
        if not e.get("seen_on_queue"):
            e["left"] = None


def _plan(records: list, sig_store: dict, allow_delete: bool):
    """Turn (key, cells) records into upsert ops vs what we last wrote (the sigs
    in `sig_store`, a flat {order#: sig} persisted in the master). Flat (not
    per-order) so backlog orders that live only in the line-items store are tracked
    too. Change detection survives restarts.

    Returns (ops, commit). `commit()` folds the new signatures into `sig_store`;
    call it ONLY after the tab's Excel write actually succeeds. Committing eagerly
    was the bug behind 'all the orders vanished': if the write then failed (Excel
    busy / OLE error) the store believed those rows were on the sheet, so the next
    poll planned no op for them and they stayed missing until a restart. Deferring
    the commit means a failed write simply re-plans the same ops next cycle."""
    desired = [(key, live_sheets.row_sig(cells), cells) for key, cells in records]
    ops = live_sheets.plan_upsert(desired, dict(sig_store), allow_delete=allow_delete)
    sig_by = {key: sig for key, sig, _ in desired}

    def commit() -> None:
        for kind, key, _ in ops:
            if kind in ("append", "update"):
                sig_store[key] = sig_by[key]
            elif kind == "delete":
                sig_store.pop(key, None)

    return ops, commit


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
    # STABLE ordering (newest job first), NOT board order: lq_jobs follows the
    # board position, which jitters between scrapes — feeding that order into
    # the tables changed the tab's fingerprint every poll and full-repainted
    # this (big) tab each cycle, freezing Excel for the whole render.
    new_today_jobs = sorted((j for j in lq_jobs if str(j.get("job") or "") in new_today),
                            key=lambda j: _sim_sort_key(str(j.get("job") or "")), reverse=True)
    events = change_log.load(today)
    # Copies with the entry's departure time attached, for the table's leading
    # Time column (the timestamp lives on the entry, not the job dict).
    removed_today = sorted(({**e.get("job", {}), "_left_iso": e.get("left") or ""}
                            for e in master.get("orders", {}).values()
                            if e.get("seen_on_queue") and not e.get("on_queue")
                            and str(e.get("left") or "")[:10] == today.isoformat()),
                           key=lambda j: _sim_sort_key(str(j.get("job") or "")), reverse=True)
    updated_at = live_sheets.fmt_datetime(now)
    # job# -> its latest job dict, for the change-order table's Design / Arrangement
    # / change-description columns.
    order_lookup = {jn: e.get("job", {}) for jn, e in master.get("orders", {}).items()}
    return live_sheets.changes_sheet(new_today_jobs, events, removed_today,
                                     today.isoformat(), updated_at=updated_at,
                                     order_lookup=order_lookup)


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


# Similar-order rows are recomputed only when the board or the stores actually
# change — the index build over the whole line-items store is ~1s, too much to
# repeat every poll for identical output.
_SIM_CACHE: dict = {"key": None, "rows": []}


def _sim_sort_key(job: str) -> tuple:
    return (0, int(job), job) if job.isdigit() else (1, 0, job)


def _similar_orders_rows(lq_jobs: list) -> list:
    """One row per (on-board order, similar past order) for the Similar Orders
    tab: each order's top lookalikes from the whole backlog, best score first,
    with custom DWGs and the shared SO lines spelled out. Groups are ordered by
    JOB NUMBER, not board position — board order reshuffles every poll, and
    following it would repaint the Similar Data tab (and shift every Live Queue
    'Similar' anchor) each cycle for nothing. Best-effort: any failure returns
    the last good rows."""
    import autocad_scan
    import find_orders

    def _mtime(p) -> int:
        try:
            return p.stat().st_mtime_ns
        except OSError:
            return 0

    key = (tuple(sorted(str(j.get("job") or "") for j in lq_jobs)),
           _mtime(line_items.store_path()), _mtime(line_items.backfill_store_path()),
           _mtime(autocad_scan.PROGRESS_PATH))
    if _SIM_CACHE["key"] == key:
        return _SIM_CACHE["rows"]
    try:
        idx = find_orders.build_index(line_items.load_store(),
                                      dwg=autocad_scan.load_progress())
        rows = []
        for j in sorted(lq_jobs, key=lambda x: _sim_sort_key(str(x.get("job") or ""))):
            jn = str(j.get("job") or "")
            for r in find_orders.similar_to_items(idx, j.get("line_items") or [],
                                                  exclude_job=jn, top=8):
                rows.append({
                    "job": jn, "similar": r["job"], "customer": r["customer"],
                    "score": round(r["score"], 2),
                    "dwg": find_orders._dwg_label(r["dwg_extras"]) or "—",
                    "shared": "; ".join(r["shared_lines"][:3] or r["shared_tags"][:4]),
                    "folder": r["dwg_folder"],
                })
        _SIM_CACHE.update(key=key, rows=rows)
    except Exception as e:  # noqa: BLE001 - the tab is a nicety, never sink a poll
        log.warning("Similar Orders rows not refreshed (%s)", e)
    return _SIM_CACHE["rows"]


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

    new_today = _new_today_ids(lq_jobs, today)   # for the Changes tab / notifications
    # Orders that had a change order (CO#) land today -> their text goes red.
    co_changed = {str(e.get("job")) for e in change_log.load(today)
                  if e.get("field") == "CO#" and e.get("job")}
    # Shade rows whose Added date is today or the most recent business day, so the
    # highlight matches the 'Added' column (when each order last came onto the
    # board) — not the morning snapshot diff, which misses intraday returns and
    # over-marks when the daily run's snapshot is stale.
    recent_days = {today, live_sheets.prev_business_day(today)}
    shade_ids = {str(j.get("job")) for j in lq_jobs
                 if live_sheets.added_date(j) in recent_days}

    # Similar-order data first: each on-board order gets its lookalike count and
    # the deep link to its group on the Similar Data tab ('Similar' column), so
    # the Live Queue rows are planned with the click-through already in place.
    sim_rows = _similar_orders_rows(lq_jobs)
    for j in lq_jobs:
        jn = str(j.get("job") or "")
        anchor = live_sheets.similar_anchor(sim_rows, jn)
        j["_sim_anchor"] = anchor
        j["_sim_count"] = sum(1 for r in sim_rows if r["job"] == jn) if anchor else 0

    lq_sigs = master.setdefault("lq_sigs", {})
    lq_ops, lq_commit = _plan(live_sheets.live_queue_records(lq_jobs, today, new_ids=shade_ids,
                                                             co_changed_ids=co_changed, ref=now),
                              lq_sigs, allow_delete=True)
    lq_payload = {"name": "Live Queue", "headers": live_sheets.LIVE_QUEUE_HEADERS, "ops": lq_ops,
                  "key_col": live_sheets.LIVE_QUEUE_KEY_COL, "allow_delete": True,
                  "sort_col": live_sheets.LIVE_QUEUE_CBC_COL,
                  "text_cols": [1, live_sheets.LIVE_QUEUE_LAST_OUT_COL],  # Added + Last Out -> AM/PM text
                  "header_row": 2, "search": True,   # search bar on row 1, headers on row 2
                  "positions": pos,  # '#' column refreshed in one bulk write (volatile in row sigs)
                  "freeze": "C3"}   # keep the search bar + header + Added/Job # columns visible

    # "Removed since this morning" block below the Live Queue. Pass it EVERY cycle
    # so the renderer can reposition it when the queue grows (otherwise the new
    # rows append right over it). Only orders the watcher actually saw on the
    # board count — never the merged backlog.
    removed = []
    for e in master.get("orders", {}).values():
        if (e.get("seen_on_queue") and not e.get("on_queue")
                and str(e.get("left") or "")[:10] == today.isoformat()):
            j = dict(e.get("job", {}))
            # Mirror on_queue()'s injection so the 'Added' column and new-today
            # shading match exactly what the order showed on the board before it left.
            j["_added_iso"] = e.get("last_in") or e.get("added")
            j["_added_known"] = e.get("added_known", False)
            removed.append((j, e.get("left")))
    removed.sort(key=lambda jl: str(jl[1] or ""), reverse=True)   # most recently removed first
    # Render each removed order like its Live Queue row — same columns, same
    # urgency/new fill and CO#-red text it had on the board. They're off the board
    # so not in shade_ids; recompute new-today shading from their own Added date.
    removed_new = {str(j.get("job")) for j, _ in removed
                   if live_sheets.added_date(j) in recent_days}
    lq_payload["below"] = live_sheets.removed_block(removed, today, new_ids=removed_new,
                                                    co_changed_ids=co_changed, ref=now)

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
    oh_ops, oh_commit = _plan(spec["records"], oh_sigs, allow_delete=False)
    oh_payload = {"name": "Order History", "spec": spec, "ops": oh_ops, "rebuild": rebuild,
                  "key_col": live_sheets.ORDER_HISTORY_KEY_COL, "freeze": "B2"}  # pin Job # only

    # Similar Orders: the picker tab + the grouped data tab it filters over
    # (sim_rows computed above, before the Live Queue rows were planned).
    # Same stable job-number order as the rows, so the sheet doesn't repaint
    # every time the board position order reshuffles.
    queue_ids = sorted((str(j.get("job")) for j in lq_jobs if j.get("job")),
                       key=_sim_sort_key)
    extra_sheets = [live_sheets.similar_data_sheet(sim_rows, queue_ids),
                    live_sheets.similar_orders_sheet(len(sim_rows), len(queue_ids))]

    # Sales Order: pick an on-board order -> its parsed SO spec + every captured
    # line item spill instantly from the flat SO Data tab (same mechanics as
    # Similar Orders — the data sheet only repaints when an order's SO content
    # actually changes).
    so_data = live_sheets.sales_order_data_sheet(lq_jobs)
    extra_sheets += [so_data,
                     live_sheets.sales_order_sheet(so_data.nrows - 1, len(queue_ids))]

    rendered = update_master_workbook(LIVE_WORKBOOK_PATH, lq_payload, oh_payload,
                                      changes_sheet=_changes_sheet(master, lq_jobs, new_today, today, now),
                                      extra_sheets=extra_sheets)
    # Commit each tab's row signatures only if its write actually landed. A tab
    # whose write failed (Excel busy / OLE error) is left out of `rendered`, so its
    # ops are re-planned and re-drawn next poll instead of being lost.
    if lq_payload["name"] in rendered:
        lq_commit()
    if oh_payload["name"] in rendered:
        oh_commit()


def poll_once(state: dict, master: dict, now: datetime, baseline: bool, announce: bool) -> dict:
    """One cycle: scrape -> record -> enrich new -> upsert the master tabs -> notify.
    Mutates `state` and `master`; returns the deltas dict from record_poll."""
    board = scrape_queue(headless=True)
    if not board:
        log.warning("Board scrape returned 0 orders — skipping this cycle (site/session issue?).")
        return {"new": [], "returning": [], "removed": []}

    # First time we see the board today, freeze it as the day's queue snapshot —
    # the 'new today' baseline that the 5 AM main.py run used to write.
    _daily_snapshot(board, now)

    # Board prices before this poll overwrites them — a change order almost always
    # moves the order's price, so a shift is our (free) signal to re-fetch the SO.
    prev_price = {jn: str((e.get("job") or {}).get("total_price") or "")
                  for jn, e in state.items()}

    deltas = live_state.record_poll(state, board, now, baseline=baseline)
    log.info("Poll @ %s: %d on board | new=%d returning=%d removed=%d",
             now.strftime("%H:%M:%S"), len(board),
             len(deltas["new"]), len(deltas["returning"]), len(deltas["removed"]))

    # Re-fetch the Sales Order for any on-board order whose price just changed: a
    # likely change order, whose new CO sales order needs downloading so the CO#,
    # line items and the Job# link all follow the newest revision. Clearing
    # `enriched` makes _enrich_pending pick it up alongside the genuinely new ones.
    if not baseline:
        repriced = 0
        for jn, oldp in prev_price.items():
            e = state.get(jn)
            if (e and e.get("present") and e.get("enriched")
                    and str((e.get("job") or {}).get("total_price") or "") != oldp):
                e["enriched"] = False
                repriced += 1
        if repriced:
            log.info("%d order(s) changed price (possible change order) — re-fetching their SO.", repriced)
        # Background round-robin: re-verify the few longest-unchecked orders so a
        # silently-stale Sales Order self-corrects within the hour.
        n_recheck = _queue_stale_so_rechecks(state, now)
        if n_recheck:
            log.info("Re-verifying %d on-board order(s) we hadn't re-checked recently.", n_recheck)

    enriched_now = _enrich_pending(state)
    # Stamp when each order was last verified, so the round-robin re-check cycles
    # through the board instead of re-doing the same orders.
    if enriched_now:
        now_iso = now.isoformat(timespec="seconds")
        for jb in enriched_now:
            e = state.get(str(jb.get("job") or ""))
            if e:
                e["verified_at"] = now_iso
    # Re-resolve AutoCAD folders for any on-board order still without one (cheap
    # now), so orders an earlier lookup missed ("Open") get filled in without a
    # full re-enrich. Mutates the state entries' job dicts in place so it sticks.
    # Repair orders (a trailing letter, e.g. 352366A) are tabled: they have no
    # standard AutoCAD folder, so re-searching triggers a full-tree sweep every
    # poll for nothing — skip them here (they're still looked up once on enrich).
    unresolved = [e["job"] for e in state.values()
                  if e.get("present") and e.get("job") and not e["job"].get("job_type")
                  and not _is_repair_order(e["job"].get("job"))]
    if unresolved:
        resolved = refresh_autocad_folders(unresolved)
        if resolved:
            log.info("Resolved %d AutoCAD folder(s) for order(s) that showed 'Open'.", resolved)

    # Repoint any Job# link whose SO file was superseded/renamed by a change order
    # at the latest revision on disk (also syncs co_number). Cheap; self-heals
    # links that broke when a CO renamed the file out from under a stored path.
    on_board = [e["job"] for e in state.values() if e.get("present") and e.get("job")]
    repointed = refresh_sales_orders(on_board)
    if repointed:
        log.info("Repointed %d Sales Order link(s) to the latest revision on disk.", repointed)

    present = live_state.present_jobs(state)
    # Fold the board into the master log; log any field modifications it found.
    events = live_master.update(master, present, now)
    # The baseline poll establishes the start-of-day picture silently: its deltas
    # are differences vs yesterday's saved master (overnight moves, or fields the
    # raw start-of-day seed hadn't re-enriched yet), NOT changes that happened
    # during today's watch — so they must NOT land in today's change log. Logging
    # them put a grey "changed today" row under (nearly) every order at 5 AM. The
    # master is still updated above; we only skip recording the deltas.
    if not baseline:
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
    # New orders (or field changes) this poll -> refresh the published snapshot so
    # a remote reader tracks them. No-ops unless auto-publish is enabled.
    if deltas["new"] or events:
        _publish_data()
    return deltas


def _morning_snapshot(now: datetime) -> None:
    if not (LIVE_MORNING_SNAPSHOT and LIVE_WORKBOOK_PATH):
        return
    dest = OUTPUT_DIR / "live_snapshots" / f"live_queue_{now.date().isoformat()}.xlsx"
    if dest.exists():
        return  # already froze this morning's copy (e.g. a restart)
    save_morning_copy(LIVE_WORKBOOK_PATH, dest)


_CONSOLE_CTRL_REF: list = []   # keep the console handler alive (and register once)


def _install_stop_handler(stop: "threading.Event") -> None:
    """Make Ctrl+C reliable. A poll blocks the main thread inside native browser
    (Playwright) and Excel (COM) calls, and that machinery can also leave Python's
    own SIGINT handler swallowed — the symptom being a watcher you can't stop
    without closing the window.

    On Windows we register an OS console-control handler (pywin32): the OS runs it
    on its *own* thread, so it fires even while the main thread is stuck in a
    native call a Python signal handler can't preempt. First Ctrl+C / Ctrl+Break
    asks the loop to stop cleanly after the current step (state is still saved); a
    second one hard-quits immediately (works mid-call, from that other thread). A
    plain SIGINT handler is also installed as a fallback (and for non-Windows)."""
    def _request_stop() -> None:
        if stop.is_set():                       # already stopping -> hard quit
            log.warning("Second interrupt — force quitting now.")
            os._exit(130)
        stop.set()
        log.info("Interrupt received — finishing the current poll, then saving and "
                 "exiting. Press Ctrl+C again to force-quit.")

    # Windows console-control handler — the robust path. Registered once.
    if os.name == "nt" and not _CONSOLE_CTRL_REF:
        try:
            import win32api  # type: ignore  (part of pywin32, already used for COM)

            def _console_handler(ctrl_type):     # runs on an OS-created thread
                if ctrl_type in (0, 1):          # CTRL_C_EVENT, CTRL_BREAK_EVENT
                    _request_stop()
                    return True                  # handled — suppress the default
                return False
            win32api.SetConsoleCtrlHandler(_console_handler, True)
            _CONSOLE_CTRL_REF.append(_console_handler)   # prevent GC
        except Exception as e:  # noqa: BLE001 - fall back to the signal handler
            log.warning("Windows console-control handler unavailable (%s) — falling "
                        "back to SIGINT, which a browser poll can occasionally cut "
                        "short. Install pywin32 for a rock-solid Ctrl+C.", e)

    # Signal handler — the only option off Windows, and a backup. Re-installed
    # each call since a poll's async work can reset it.
    def _sig_handler(signum, frame):
        _request_stop()
    for signame in ("SIGINT", "SIGBREAK"):      # SIGBREAK = Ctrl+Break on Windows
        sig = getattr(signal, signame, None)
        if sig is None:
            continue
        try:
            signal.signal(sig, _sig_handler)
        except (ValueError, OSError):           # not main thread / unsupported
            pass


def _install_stopfile_watcher(stop: "threading.Event") -> None:
    """Stop on a flag file from the desktop launcher.

    The launcher runs windowless (pythonw) and can't send Ctrl+Break, so it drops
    a per-PID flag file instead (see stop_signal.py). A daemon thread polls for it
    and sets the same stop event Ctrl+C uses — so the launcher's Stop button runs
    this same clean exit (finish the poll, save state, publish logs) rather than a
    hard kill. Harmless when run outside the launcher: the flag never appears."""
    pid = os.getpid()
    stop_signal.clear_stop(pid)                 # drop any stale flag for a reused PID

    def _watch() -> None:
        while not stop.wait(1.0):               # returns True the instant stop is set
            if stop_signal.stop_requested(pid):
                log.info("Stop requested by the launcher — finishing the current poll, "
                         "then saving state and publishing logs before exiting.")
                stop.set()
                return

    threading.Thread(target=_watch, name="stopfile-watch", daemon=True).start()


def run_watch(ignore_window: bool = False) -> int:
    setup_logging()
    try:   # record the running code version so the published log confirms it
        log.info("watch.py version: branch %s @ commit %s  (pid %d)",
                 git_update.current_branch() or "?", git_update.head_rev()[:10] or "?",
                 os.getpid())
    except Exception:  # noqa: BLE001
        pass
    validate_runtime_config()
    try:
        import order_verification_cleanup

        cleanup_counts = order_verification_cleanup.run()
        if order_verification_cleanup.changed(cleanup_counts):
            _publish_data()
    except Exception:  # noqa: BLE001 - cleanup failure must remain visible without ending watch
        log.exception("Order Verification cleanup failed at watcher startup")
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
    _merge_backlog_sources()
    master = live_master.load_master()
    purged = change_log.purge_day_once(today)
    if purged:
        log.info("One-time purge: archived %d pre-fix event(s); today's change log "
                 "(and the Changes tab) restarts clean.", purged)
    state_jobs = [entry["job"] for entry in state.values()
                  if entry.get("present") and entry.get("job")]
    repaired = repair_missing_sales_order_summaries(state_jobs)
    if repaired:
        log.info("Repaired missing Sales Order summary fields for %d on-board order(s) "
                 "from the archived PDFs.", repaired)
    # Reclaim on-board orders whose master copy a helper merge overwrote while
    # we were down: reset them to the live state's values silently, so the
    # difference doesn't surface as bogus 'changes' when polls/re-verifications
    # fold the live values back in.
    realigned = live_master.realign_orders(master, live_state.present_jobs(state))
    if realigned:
        log.info("Re-aligned %d on-board order(s) in the master to the live state "
                 "(no change events).", realigned)
    # Persist repairs before using the restored values to clean the event log.
    if repaired or realigned:
        live_state.save_state(state, today)
        live_master.save_master(master)
    scrubbed = change_log.scrub_phantom_blanks(today, master)
    if scrubbed:
        log.info("Scrubbed %d phantom '-> blank' event(s) from today's change log.", scrubbed)
    deduped = change_log.dedupe_repeated_transitions(today)
    if deduped:
        log.info("Removed %d repeated field-transition event(s) from today's change log.",
                 deduped)
    tagged = engineers.backfill(master)   # tag historical orders by engineer (roster edits too)
    if tagged:
        log.info("Engineer roster: tagged/updated %d existing order(s)", tagged)
    _force_rebuild(master)            # process start -> rebuild the upsert tabs once
    first = not state
    if first:
        _seed_baseline(state, datetime.now())

    log.info("=== Live watch start: %s  window %s-%s  every %ds  (%d orders in master log) ===",
             today.isoformat(), start.strftime("%H:%M"), end.strftime("%H:%M"),
             POLL_INTERVAL_SECONDS, len(master.get("orders", {})))

    stop = threading.Event()
    _install_stop_handler(stop)
    _install_stopfile_watcher(stop)             # launcher Stop button -> clean exit
    if _CONSOLE_CTRL_REF:
        log.info("Ctrl+C: the poll in progress finishes first, then state is saved and the "
                 "watch exits (press Ctrl+C again to force-quit).")
    else:
        log.info("Ctrl+C: requests a clean stop and the run finishes the current poll where it "
                 "can; press Ctrl+C again to force-quit.")
    if LOG_PUSH_BRANCH:
        _publish_logs()               # snapshot the prior session at startup so the branch exists
    cycle = 0
    try:
        while not stop.is_set():
            now = datetime.now()
            if not ignore_window and now >= end:
                log.info("Reached end of watch window (%s); stopping.", end.strftime("%H:%M"))
                break
            baseline = first and cycle == 0
            t0 = time.monotonic()
            try:
                poll_once(state, master, now, baseline=baseline, announce=not baseline)
            except KeyboardInterrupt:   # handler was bypassed and the poll was cut short
                log.info("Ctrl+C received mid-poll — saving and exiting now.")
                stop.set()
            except Exception:  # noqa: BLE001 - one bad cycle must not end the watch
                log.exception("Poll cycle failed; continuing to the next one")
            if baseline:
                _morning_snapshot(now)
            cycle += 1
            # Periodically recycle the live workbook so Excel's all-day memory
            # accumulation (fragmented CF rules, calc chain, undo/redraw caches)
            # can't climb unbounded: close it now (AutoSave already synced; only the
            # bot's Excel is touched) and the next poll reopens it fresh. Done during
            # the idle wait so the reopen cost lands on the next poll, not this one.
            if (EXCEL_RECYCLE_EVERY_POLLS and LIVE_WORKBOOK_PATH
                    and cycle % EXCEL_RECYCLE_EVERY_POLLS == 0):
                try:
                    recycle_workbook(LIVE_WORKBOOK_PATH)
                except Exception:  # noqa: BLE001 - recycling is a nicety, never fatal
                    log.exception("Workbook recycle failed; continuing")
            # A poll's browser/async work can reset our SIGINT handler — re-assert
            # it, then wait out the rest of the interval interruptibly: the wait
            # returns the instant a Ctrl+C sets the stop event.
            _install_stop_handler(stop)
            elapsed = time.monotonic() - t0
            if stop.wait(max(0.0, POLL_INTERVAL_SECONDS - elapsed)):
                break
    except KeyboardInterrupt:           # belt and suspenders if the handler is bypassed
        pass
    stop_signal.clear_stop(os.getpid())         # consume the launcher's stop flag
    log.info("Stopping — saving state.")
    live_state.save_state(state, today)
    live_master.save_master(master)
    log.info("=== Live watch done (%d cycles) ===", cycle)
    if LOG_PUSH_BRANCH:
        _publish_logs()               # publish on Ctrl+C / window end — captures the full session
    _publish_data()                   # final snapshot of the day's gathered order data
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
        try:
            import order_verification_cleanup

            cleanup_counts = order_verification_cleanup.run()
            if order_verification_cleanup.changed(cleanup_counts):
                _publish_data()
        except Exception:  # noqa: BLE001
            log.exception("Order Verification cleanup failed at watcher startup")
        today = date.today()
        state = live_state.load_state(today)
        _merge_backlog_sources()
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
