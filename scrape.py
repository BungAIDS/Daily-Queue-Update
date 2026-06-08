"""Stage 1 of the daily run: scrape the queue, diff it against the most recent
prior run, and build the Excel report. No AI overview, no email.

    python scrape.py

Saves today's snapshot, diff, and Excel so `brief.py` can add the AI overview
and `send.py` can email it. main.py runs this as the first step of the full run.
"""
from __future__ import annotations

import sys

from pipeline import run_stage, stage_scrape

if __name__ == "__main__":
    sys.exit(run_stage("scrape", stage_scrape))
