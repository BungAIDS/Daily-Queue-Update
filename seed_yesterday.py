"""One-time helper: drop the bundled 49-job baseline into your snapshots
folder as YESTERDAY's snapshot, so the next make_report.py / main.py run
shows a real day-over-day diff against the live cbcinsider queue.

Run this once. After it succeeds, run `python make_report.py` to scrape the
live site and see what's changed since the baseline.

    python seed_yesterday.py

Safe to re-run: if a snapshot for yesterday already exists, the script asks
before overwriting it.
"""
from __future__ import annotations

import json
import sys
from datetime import date, timedelta
from pathlib import Path

from compare import save_snapshot, snapshot_path

SEED_FILE = Path(__file__).parent / "seed_data" / "queue_baseline.json"


def main() -> int:
    if not SEED_FILE.exists():
        print(f"Seed file missing: {SEED_FILE}")
        return 1

    jobs = json.loads(SEED_FILE.read_text(encoding="utf-8"))
    yesterday = date.today() - timedelta(days=1)
    target = snapshot_path(yesterday)

    if target.exists():
        resp = input(f"{target} already exists. Overwrite? [y/N] ").strip().lower()
        if resp != "y":
            print("Aborted; no changes made.")
            return 0

    save_snapshot(jobs, yesterday)
    print(f"\nSeeded yesterday ({yesterday}) with {len(jobs)} baseline jobs.")
    print("Next: run  python make_report.py  to scrape live and see the real diff.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
