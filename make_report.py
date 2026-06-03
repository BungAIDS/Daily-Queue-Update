"""Build the Excel report from a live scrape WITHOUT sending email or touching
tracking state — handy for a quick demo or a one-off manual report.

    python make_report.py            # includes the AI overview if a key is set
    python make_report.py --no-ai    # skip Claude even when a key is set

Produces the same workbook the daily run does. If ANTHROPIC_API_KEY is set (and
--no-ai isn't passed) it generates the real AI overview too; otherwise that
section is a placeholder. Either way it sends no email and writes no
snapshot/history.

This is READ-ONLY: it does not save today's snapshot or modify history, so you
can run it as many times as you like without disturbing the official once-a-day
tracking state (that's owned solely by main.py).
"""
from __future__ import annotations

import logging
import sys
from datetime import date, timedelta

from analyzer import analyze
from compare import diff_queues, load_history, load_snapshot
from config import ANTHROPIC_API_KEY
from excel_writer import build_workbook
from sales_orders import enrich_with_sales_orders
from scraper import scrape_queue


def _placeholder(msg: str) -> dict:
    return {"briefing": msg, "anomalies": [], "action_items": []}


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    today = date.today()

    # headless=False so you can watch it work (nice for a demo).
    jobs = scrape_queue(headless=False)
    if not jobs:
        print("No jobs scraped — run check_access.py for troubleshooting tips.")
        return 1

    # Enrich with sales orders (CO#, fan size/arrangement, folder links). This
    # is the slow step (opens each order's detail in parallel); skip-friendly.
    try:
        enrich_with_sales_orders(jobs)
    except Exception as e:  # noqa: BLE001
        print(f"(Sales-order enrichment skipped: {e})")

    yesterday = load_snapshot(today - timedelta(days=1))
    # Read-only: classify returning orders from history but don't write state.
    diff = diff_queues(jobs, yesterday, today, persist_history=False)

    # Run the real AI overview if a key is set — still no email, no snapshot.
    # Pass --no-ai to skip Claude even when a key is set.
    no_ai = "--no-ai" in sys.argv
    if ANTHROPIC_API_KEY and not no_ai:
        print("Generating AI overview via Claude...")
        try:
            briefing = analyze(diff, today, all_jobs=jobs)
        except Exception as e:  # noqa: BLE001
            print(f"(AI analysis failed: {e})")
            briefing = _placeholder(f"(AI analysis failed: {e})")
    elif no_ai:
        briefing = _placeholder("(AI overview skipped — ran with --no-ai.)")
    else:
        briefing = _placeholder(
            "(AI overview skipped — set ANTHROPIC_API_KEY in .env to generate it here.)"
        )

    path = build_workbook(jobs, diff, briefing, today, history=load_history())
    print("\n" + "=" * 64)
    print("Excel report written to:")
    print(f"  {path}")
    print("=" * 64)
    print("Open that file to see the Full Queue, Changes, and History tabs.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
