"""Daily queue briefing — entrypoint.

Run order:
  1. Scrape today's queue from cbcinsider
  2. Load yesterday's snapshot (if any)
  3. Diff -> save today's snapshot
  4. Send the diff to Claude for analysis
  5. Build the Excel report
  6. Email the plain-text briefing

Failures at any step send an alert email and exit non-zero.
"""
from __future__ import annotations

import logging
import sys
import traceback
from datetime import date

from analyzer import analyze
from compare import diff_queues, load_history, load_latest_snapshot, save_snapshot
from config import ANTHROPIC_API_KEY, validate_runtime_config
from emailer import send_alert, send_daily_briefing
from excel_writer import build_workbook
from sales_orders import enrich_with_sales_orders
from scraper import scrape_queue


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )


def main() -> int:
    setup_logging()
    log = logging.getLogger("daily-queue")
    today = date.today()

    try:
        log.info("=== Daily queue run: %s ===", today.isoformat())
        validate_runtime_config()

        jobs = scrape_queue(headless=True)
        if not jobs:
            raise RuntimeError("Scraper returned 0 jobs — site may be down or layout changed.")

        # Enrich every job with its sales order: CO# / change-order history,
        # fan design/size/arrangement, and the AutoCAD folder link. This opens
        # each order's detail in parallel and is the slow step (~minutes). Never
        # let it sink the run — on failure we alert and continue without it.
        try:
            enrich_with_sales_orders(jobs)
        except Exception:
            log.exception("Sales-order enrichment failed — continuing without it")
            send_alert(today.isoformat(), "Sales-order enrichment failed:\n\n" + traceback.format_exc())

        # Diff against the most recent prior snapshot, not a fixed today-1: the
        # run only fires on business days, so on a Monday "yesterday" is Sunday
        # (no snapshot) and a today-1 lookback would flag the whole queue as new.
        yesterday, prev_date = load_latest_snapshot(today)
        if yesterday is None:
            log.warning(
                "No prior snapshot within lookback window — every job will be "
                "reported as new (expected only on a first run)."
            )
        else:
            log.info("Diffing against previous snapshot from %s", prev_date)
        diff = diff_queues(jobs, yesterday, today, prev_date=prev_date)
        save_snapshot(jobs, today)

        if not ANTHROPIC_API_KEY:
            log.info("ANTHROPIC_API_KEY not set — skipping AI briefing this run.")
            briefing = {
                "briefing": "(AI briefing skipped — no Anthropic API key set in .env.)",
                "anomalies": [],
                "action_items": [],
            }
        else:
            try:
                briefing = analyze(diff, today, all_jobs=jobs)
            except Exception as e:
                log.exception("Claude analysis failed — continuing with empty briefing")
                send_alert("Claude API error", traceback.format_exc())
                briefing = {
                    "briefing": f"(AI analysis unavailable: {e})",
                    "anomalies": [],
                    "action_items": [],
                }

        # diff_queues has already updated history on disk; read it back so the
        # report's History tab reflects today's archived/returned jobs.
        excel_path = build_workbook(jobs, diff, briefing, today, history=load_history())
        send_daily_briefing(briefing, diff, excel_path, today.isoformat())

        log.info("=== Done ===")
        return 0

    except Exception as e:
        log.exception("Daily queue run failed")
        send_alert(today.isoformat(), traceback.format_exc())
        return 1


if __name__ == "__main__":
    sys.exit(main())
