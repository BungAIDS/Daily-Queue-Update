"""Quick connectivity check — scrape the queue and print what we found.

Run this AFTER `python login.py`. It does NOT need your Anthropic key or
email; it only confirms the saved session can reach your queue and that the
parser reads the columns correctly.

    python check_access.py
"""
from __future__ import annotations

import logging
import sys

from scraper import scrape_queue


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

    # headless=False so you can watch the browser open and load your queue.
    jobs = scrape_queue(headless=False)

    print("\n" + "=" * 64)
    print(f"Scraped {len(jobs)} jobs from the queue.")
    print("=" * 64)

    if not jobs:
        print("No jobs found. Check that CBC_QUEUE_URL in .env points at your")
        print("dispatch page and that `python login.py` finished on that page.")
        return 1

    print(f"{'Job':<10}{'Status':<13}{'Customer':<30}{'End Date':<12}{'Price':>12}  Flags")
    print("-" * 90)
    for j in jobs[:8]:
        flags = [k for k in ("unapproved", "credit_hold", "has_notes") if j.get(k)]
        print(
            f"{j.get('job',''):<10}{j.get('status',''):<13}"
            f"{j.get('customer','')[:28]:<30}{j.get('end_date',''):<12}"
            f"{j.get('total_price',''):>12}  {','.join(flags)}"
        )
    if len(jobs) > 8:
        print(f"... and {len(jobs) - 8} more")

    print("\nLooks good if these job numbers, customers, and prices match the site.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
