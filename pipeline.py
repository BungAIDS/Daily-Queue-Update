"""Shared pipeline stages for the daily queue briefing.

The work is split into three stages, each exposed as its own runnable script:
  scrape.py -> stage_scrape : scrape + diff + Excel (no AI, no email)
  brief.py  -> stage_brief  : add the AI overview to today's run (no email)
  send.py   -> emails the most recent run

main.py runs all three in one shot — that's the 5 AM job. These functions are
the single source of truth, so the scripts stay thin and never drift apart.

History is advanced exactly once, by whichever run does the scrape+diff
(scrape.py or main.py). brief.py only ever reads/recomputes it read-only.
"""
from __future__ import annotations

import logging
import sys
import traceback
from datetime import date
from pathlib import Path
from typing import Callable

from analyzer import analyze
from compare import (diff_queues, load_history, load_latest_snapshot,
                     load_snapshot, save_snapshot)
from config import ANTHROPIC_API_KEY, validate_runtime_config
from emailer import send_alert, send_daily_briefing
from excel_writer import build_workbook
from runstate import (archive_old_runs, load_diff, save_briefing, save_diff,
                      save_excel_path)
from sales_orders import enrich_with_sales_orders
from scraper import scrape_queue

log = logging.getLogger("daily-queue")

# Placeholder AI section written by scrape.py until brief.py fills it in.
_AI_PENDING = {
    "briefing": "(AI briefing not generated yet — run `python brief.py`.)",
    "anomalies": [],
    "action_items": [],
}


class StageNotReady(Exception):
    """A stage was run before the one it depends on. Operator error, not a system
    failure, so run_stage reports it without emailing an alert."""


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )


# --- primitives ------------------------------------------------------------

def scrape_and_diff(today: date) -> tuple[list, dict]:
    """Scrape, enrich, diff vs the most recent prior snapshot, and persist
    today's snapshot + diff. Advances the long-term history exactly once."""
    jobs = scrape_queue(headless=True)
    if not jobs:
        raise RuntimeError("Scraper returned 0 jobs — site may be down or layout changed.")

    # Enrich every job with its sales order: CO# / change-order history, fan
    # design/size/arrangement, and the AutoCAD folder link. The slow step
    # (~minutes); never let it sink the run — on failure alert and continue.
    try:
        enrich_with_sales_orders(jobs)
    except Exception:
        log.exception("Sales-order enrichment failed — continuing without it")
        send_alert(today.isoformat(), "Sales-order enrichment failed:\n\n" + traceback.format_exc())

    # Diff against the most recent prior snapshot, not a fixed today-1: the run
    # only fires on business days, so on a Monday "yesterday" is Sunday (no
    # snapshot) and a today-1 lookback would flag the whole queue as new.
    yesterday, prev_date = load_latest_snapshot(today)
    if yesterday is None:
        log.warning("No prior snapshot within lookback window — every job will be "
                    "reported as new (expected only on a first run).")
    else:
        log.info("Diffing against previous snapshot from %s", prev_date)
    diff = diff_queues(jobs, yesterday, today, prev_date=prev_date)
    save_snapshot(jobs, today)
    save_diff(diff, today)
    # Sweep per-run files older than ~2 months into archive/ subfolders. Moved,
    # never deleted — the full record of every order stays on disk.
    archive_old_runs(today)
    return jobs, diff


def run_ai(diff: dict, today: date, jobs: list) -> dict:
    """Generate the AI briefing, or a placeholder if no key / on API failure."""
    if not ANTHROPIC_API_KEY:
        log.info("ANTHROPIC_API_KEY not set — skipping AI briefing this run.")
        return {**_AI_PENDING, "briefing": "(AI briefing skipped — no Anthropic API key set in .env.)"}
    try:
        return analyze(diff, today, all_jobs=jobs)
    except Exception as e:
        log.exception("Claude analysis failed — continuing with empty briefing")
        send_alert("Claude API error", traceback.format_exc())
        return {**_AI_PENDING, "briefing": f"(AI analysis unavailable: {e})"}


def build_excel(jobs: list, diff: dict, briefing: dict, today: date) -> Path:
    # diff_queues already updated history on disk; read it back so the report's
    # History tab reflects today's archived/returned jobs.
    path = build_workbook(jobs, diff, briefing, today, history=load_history())
    save_excel_path(path, today)
    return path


# --- stages (one per script) -----------------------------------------------

def stage_scrape(today: date) -> None:
    """scrape.py: scrape + diff + Excel, with a placeholder AI section."""
    jobs, diff = scrape_and_diff(today)
    briefing = dict(_AI_PENDING)
    save_briefing(briefing, today)
    build_excel(jobs, diff, briefing, today)
    log.info("Scrape stage complete. Next: python brief.py")


def stage_brief(today: date) -> None:
    """brief.py: add the AI overview to today's scraped run and rebuild the Excel."""
    jobs = load_snapshot(today)
    if jobs is None:
        raise StageNotReady("No snapshot for today — run `python scrape.py` first "
                            "(or `python main.py` for the full run).")
    diff = load_diff(today)
    if diff is None:
        # Snapshot exists but no staged diff (e.g. an older run). Recompute it
        # read-only from the snapshot so we can add AI without re-scraping.
        yesterday, prev_date = load_latest_snapshot(today)
        diff = diff_queues(jobs, yesterday, today, persist_history=False, prev_date=prev_date)
        log.info("No saved diff for today; recomputed read-only from the snapshot (vs %s).",
                 prev_date or "no prior snapshot")
    briefing = run_ai(diff, today, jobs)
    save_briefing(briefing, today)
    save_diff(diff, today)
    build_excel(jobs, diff, briefing, today)
    log.info("Brief stage complete. Next: python send.py")


def stage_full(today: date) -> None:
    """main.py: scrape -> brief -> send, in one shot (the 5 AM job)."""
    jobs, diff = scrape_and_diff(today)
    briefing = run_ai(diff, today, jobs)
    save_briefing(briefing, today)
    excel_path = build_excel(jobs, diff, briefing, today)
    send_daily_briefing(briefing, diff, excel_path, today.isoformat())


# --- shared entry-point wrapper --------------------------------------------

def run_stage(label: str, fn: Callable[[date], None]) -> int:
    """Run a stage with shared logging, config validation, and error handling:
    StageNotReady exits cleanly (operator error); anything else alerts by email."""
    setup_logging()
    today = date.today()
    try:
        log.info("=== Daily queue (%s): %s ===", label, today.isoformat())
        validate_runtime_config()
        fn(today)
        log.info("=== Done ===")
        return 0
    except StageNotReady as e:
        # Ran a stage before its prerequisite — not a real failure, so report it
        # plainly and skip the alert email (don't spam the team).
        log.error("%s", e)
        return 1
    except Exception:
        log.exception("Daily queue run failed")
        send_alert(today.isoformat(), traceback.format_exc())
        return 1
