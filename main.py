"""Daily queue briefing — entrypoint.

Full run (default):
  1. Scrape today's queue from cbcinsider
  2. Load the most recent prior snapshot
  3. Diff -> save today's snapshot + diff
  4. Send the diff to Claude for analysis
  5. Build the Excel report
  6. Email the plain-text briefing

The run can also be driven one stage at a time, so a botched 5 AM run can be
recovered without redoing the slow steps:

    python main.py --no-ai      # scrape + diff + Excel; no AI, no email
    python main.py --ai-only    # add the AI briefing to today's saved run; no email
    python main.py --mail-only  # email today's finished briefing + Excel

--ai-only and --mail-only reuse what the earlier stage wrote to disk (snapshot,
diff, briefing, Excel path), so neither re-scrapes the site. History is advanced
exactly once — by whichever stage does the scrape+diff (--no-ai or a full run).

Failures at any step send an alert email and exit non-zero.
"""
from __future__ import annotations

import argparse
import logging
import sys
import traceback
from datetime import date
from pathlib import Path

from analyzer import analyze
from compare import (diff_queues, load_history, load_latest_snapshot,
                     load_snapshot, save_snapshot)
from config import ANTHROPIC_API_KEY, validate_runtime_config
from emailer import send_alert, send_daily_briefing
from excel_writer import build_workbook
from runstate import (load_briefing, load_diff, load_excel_path, save_briefing,
                      save_diff, save_excel_path)
from sales_orders import enrich_with_sales_orders
from scraper import scrape_queue

log = logging.getLogger("daily-queue")

# Placeholder briefing used when the AI stage hasn't run yet (e.g. after --no-ai).
_AI_PENDING = {
    "briefing": "(AI briefing not generated yet — run `python main.py --ai-only`.)",
    "anomalies": [],
    "action_items": [],
}


class StageNotReady(Exception):
    """A staged run was invoked before the stage it depends on. This is operator
    error, not a system failure, so main() reports it without emailing an alert."""


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )


def _scrape_and_diff(today: date) -> tuple[list, dict]:
    """Scrape, enrich, diff against the most recent prior snapshot, and persist
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
        log.warning(
            "No prior snapshot within lookback window — every job will be "
            "reported as new (expected only on a first run)."
        )
    else:
        log.info("Diffing against previous snapshot from %s", prev_date)
    diff = diff_queues(jobs, yesterday, today, prev_date=prev_date)
    save_snapshot(jobs, today)
    save_diff(diff, today)
    return jobs, diff


def _run_ai(diff: dict, today: date, jobs: list) -> dict:
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


def _build_excel(jobs: list, diff: dict, briefing: dict, today: date) -> Path:
    # diff_queues already updated history on disk; read it back so the report's
    # History tab reflects today's archived/returned jobs.
    path = build_workbook(jobs, diff, briefing, today, history=load_history())
    save_excel_path(path, today)
    return path


def _require(value, what: str, hint: str):
    if value is None:
        raise StageNotReady(f"No saved {what} for today — {hint}")
    return value


def run_full(today: date) -> None:
    jobs, diff = _scrape_and_diff(today)
    briefing = _run_ai(diff, today, jobs)
    save_briefing(briefing, today)
    excel_path = _build_excel(jobs, diff, briefing, today)
    send_daily_briefing(briefing, diff, excel_path, today.isoformat())


def run_no_ai(today: date) -> None:
    jobs, diff = _scrape_and_diff(today)
    briefing = dict(_AI_PENDING)
    save_briefing(briefing, today)
    _build_excel(jobs, diff, briefing, today)
    log.info("No-AI stage complete. Next: python main.py --ai-only")


def run_ai_only(today: date) -> None:
    jobs = _require(load_snapshot(today), "snapshot", "run `python main.py --no-ai` first.")
    diff = _require(load_diff(today), "diff", "run `python main.py --no-ai` first.")
    briefing = _run_ai(diff, today, jobs)
    save_briefing(briefing, today)
    _build_excel(jobs, diff, briefing, today)
    log.info("AI stage complete. Next: python main.py --mail-only")


def run_mail_only(today: date) -> None:
    diff = _require(load_diff(today), "diff", "run `python main.py --no-ai` first.")
    briefing = _require(load_briefing(today), "briefing", "run `python main.py --ai-only` first.")
    excel_path = _require(load_excel_path(today), "Excel report", "run `python main.py --no-ai` first.")
    if not excel_path.exists():
        raise StageNotReady(f"Saved Excel path no longer exists: {excel_path} — re-run --no-ai.")
    send_daily_briefing(briefing, diff, excel_path, today.isoformat())
    log.info("Email sent.")


def main() -> int:
    setup_logging()
    parser = argparse.ArgumentParser(description="Daily queue briefing.")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--no-ai", action="store_true",
                       help="Scrape, diff, and build the Excel; skip AI and email.")
    group.add_argument("--ai-only", action="store_true",
                       help="Add the AI briefing to today's saved no-AI run; skip email.")
    group.add_argument("--mail-only", action="store_true",
                       help="Email today's already-built briefing + Excel.")
    args = parser.parse_args()

    today = date.today()
    stage = ("no-AI stage" if args.no_ai else "AI-only stage" if args.ai_only
             else "mail-only stage" if args.mail_only else "full run")
    try:
        log.info("=== Daily queue run (%s): %s ===", stage, today.isoformat())
        validate_runtime_config()

        if args.no_ai:
            run_no_ai(today)
        elif args.ai_only:
            run_ai_only(today)
        elif args.mail_only:
            run_mail_only(today)
        else:
            run_full(today)

        log.info("=== Done ===")
        return 0

    except StageNotReady as e:
        # Operator ran a stage before its prerequisite — not a real failure, so
        # report it plainly and skip the alert email (don't spam the team).
        log.error("%s", e)
        return 1

    except Exception:
        log.exception("Daily queue run failed")
        send_alert(today.isoformat(), traceback.format_exc())
        return 1


if __name__ == "__main__":
    sys.exit(main())
