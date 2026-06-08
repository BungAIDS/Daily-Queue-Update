"""Stage 2 of the daily run: add the AI overview to today's already-scraped run
and rebuild the Excel. No email.

    python brief.py

Reuses today's snapshot/diff from `scrape.py` (or any earlier run that scraped
today), so it never re-scrapes — just the one Claude call. Run `scrape.py` first
if today hasn't been scraped yet. main.py runs this as the second step of the
full run.
"""
from __future__ import annotations

import sys

from pipeline import run_stage, stage_brief

if __name__ == "__main__":
    sys.exit(run_stage("brief", stage_brief))
