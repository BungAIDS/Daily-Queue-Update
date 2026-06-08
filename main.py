"""Daily queue briefing — the 5 AM job.

Scrapes today's queue, diffs it against the most recent prior run, generates the
AI overview, builds the Excel report, and emails it — all in one shot. This is
what the scheduled task runs every morning.

The same work is available as three separate scripts, for recovering a botched
run without redoing everything:

    python scrape.py   # scrape + diff + Excel   (no AI, no email)
    python brief.py    # add the AI overview      (no email)
    python send.py     # email the most recent run

On any failure, an alert email is sent and the run exits non-zero.
"""
from __future__ import annotations

import sys

from pipeline import run_stage, stage_full

if __name__ == "__main__":
    sys.exit(run_stage("full run", stage_full))
