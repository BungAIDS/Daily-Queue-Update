"""Build the Excel report from a live scrape WITHOUT calling Claude or sending
email. No Anthropic API key or email needed — handy for a quick demo or a
one-off manual report.

    python make_report.py

Produces the same two-tab workbook the daily run does (Full Queue + Changes),
just with the AI-briefing section left as a placeholder.

This is READ-ONLY: it does not save today's snapshot or modify history, so you
can run it as many times as you like without disturbing the official once-a-day
tracking state (that's owned solely by main.py).
"""
from __future__ import annotations

import logging
import sys
from datetime import date, timedelta

from compare import diff_queues, load_history, load_snapshot
from excel_writer import build_workbook
from scraper import scrape_queue


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    today = date.today()

    # headless=False so you can watch it work (nice for a demo).
    jobs = scrape_queue(headless=False)
    if not jobs:
        print("No jobs scraped — run check_access.py for troubleshooting tips.")
        return 1

    yesterday = load_snapshot(today - timedelta(days=1))
    # Read-only: classify returning orders from history but don't write state.
    diff = diff_queues(jobs, yesterday, today, persist_history=False)

    briefing = {
        "briefing": "(AI briefing skipped. Run main.py with an Anthropic API key "
                    "to generate the natural-language summary, anomaly flags, and "
                    "ranked action items here.)",
        "anomalies": [],
        "action_items": [],
    }

    path = build_workbook(jobs, diff, briefing, today, history=load_history())
    print("\n" + "=" * 64)
    print("Excel report written to:")
    print(f"  {path}")
    print("=" * 64)
    print("Open that file to see the Full Queue, Changes, and History tabs.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
