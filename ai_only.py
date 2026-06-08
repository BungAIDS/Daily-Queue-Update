"""Re-run ONLY the AI briefing against the latest already-saved snapshot.

No scrape, no email, no state change. Uses today's snapshot that the daily
run (main.py) already wrote to disk, diffs it against the most recent prior
snapshot (read-only — history is not advanced), and prints the AI briefing.

    python ai_only.py            # re-brief today's saved snapshot
    python ai_only.py 2026-06-08 # re-brief a specific saved day (ISO date)

Use this when the morning run already scraped but you want a fresh briefing
(e.g. after a fix) without paying for the slow live scrape again. Requires
ANTHROPIC_API_KEY in .env; otherwise there's nothing for it to do.
"""
from __future__ import annotations

import json
import sys
from datetime import date

from analyzer import analyze
from compare import diff_queues, load_latest_snapshot, load_snapshot, snapshot_path
from config import ANTHROPIC_API_KEY


def main() -> int:
    if not ANTHROPIC_API_KEY:
        print("ANTHROPIC_API_KEY is not set in .env — nothing to run.")
        return 1

    day = date.fromisoformat(sys.argv[1]) if len(sys.argv) > 1 else date.today()

    jobs = load_snapshot(day)
    if jobs is None:
        print(f"No saved snapshot for {day} at {snapshot_path(day)}.")
        print("Run the daily job (main.py) or make_report.py first.")
        return 1

    prev, prev_date = load_latest_snapshot(day)
    print(f"Briefing {day} ({len(jobs)} jobs) vs {prev_date or 'no prior snapshot'}\n")

    # persist_history=False: read-only, never advances the official tracking
    # state (that's owned solely by main.py).
    diff = diff_queues(jobs, prev, day, persist_history=False, prev_date=prev_date)
    briefing = analyze(diff, day, all_jobs=jobs)

    print("=" * 64)
    print(briefing.get("briefing", ""))
    if briefing.get("anomalies"):
        print("\nANOMALIES")
        for a in briefing["anomalies"]:
            print(f"  - {a}")
    if briefing.get("action_items"):
        print("\nACTION ITEMS")
        for a in briefing["action_items"]:
            print(f"  - {a}")
    print("=" * 64)
    return 0


if __name__ == "__main__":
    sys.exit(main())
