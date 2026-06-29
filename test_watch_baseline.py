"""Regression test: the start-of-day baseline poll must not pollute the change log.

    python test_watch_baseline.py

The live watcher's first poll of the day establishes the start-of-day picture
(seeded from the morning snapshot, no notifications). Its `live_master.update`
deltas are differences vs *yesterday's* saved master — overnight moves, or
fields the raw seed hadn't re-enriched yet — not changes that happened during
today's watch. They must NOT be appended to today's change log; otherwise the
Changes tab grows a grey "changed today" row under (nearly) every order at 5 AM.

Later (non-baseline) polls must still log genuine intraday changes, so each real
movement keeps getting its own grey row.
"""
from __future__ import annotations

import sys
import tempfile
from datetime import date, datetime
from pathlib import Path
from unittest import mock

import config

# Redirect every snapshot / change-log / state write to a throwaway dir so the
# test never touches real run data.
config.SNAPSHOT_DIR = Path(tempfile.mkdtemp())

import change_log  # noqa: E402  (after SNAPSHOT_DIR is redirected)
import live_master  # noqa: E402
import live_state  # noqa: E402
import watch  # noqa: E402

TODAY = date(2026, 6, 25)


def _master_with_enriched_order(so_pdf: str = "") -> dict:
    """Yesterday's master: one order fully enriched with Sales-Order fields."""
    job = {
        "job": "412348", "customer": "MOTION", "oper": "53", "co_number": 0,
        "so_pdf": so_pdf, "end_date": "07/01/2026",
        "so_design_desc": "Pressure Blower", "so_size": "P4", "so_arrangement": "4",
    }
    return {"orders": {"412348": {
        "added": "2026-06-24T05:00:00", "last_in": "2026-06-24T05:00:00",
        "last_out": None, "left": None, "on_queue": True, "seen_on_queue": True,
        "history": [{"event": "in", "time": "2026-06-24T05:00:00"}], "job": job,
    }}}


def _poll(master: dict, state: dict, now: datetime, baseline: bool, board: list) -> None:
    """Run one poll_once with all browser / Excel / disk side effects stubbed."""
    with mock.patch.object(watch, "scrape_queue", return_value=board), \
            mock.patch.object(watch, "_enrich_pending", return_value=[]), \
            mock.patch.object(watch, "refresh_autocad_folders", return_value=0), \
            mock.patch.object(watch, "refresh_sales_orders", return_value=0), \
            mock.patch.object(watch, "_render_master"), \
            mock.patch.object(watch.live_state, "save_state"), \
            mock.patch.object(watch.live_master, "save_master"):
        watch.poll_once(state, master, now, baseline=baseline, announce=False)


def test_baseline_poll_does_not_log_changes():
    """A raw start-of-day seed would 'lose' the SO fields vs the enriched master,
    but the baseline poll must record none of that in today's change log."""
    change_log.save(TODAY, [])  # start clean
    master = _master_with_enriched_order(so_pdf="")  # blank link -> wipe is visible
    state: dict = {}
    live_state.seed_from_snapshot(
        state, [{"job": "412348", "customer": "MOTION", "oper": "53"}],
        "2026-06-25T05:00:00")

    _poll(master, state, datetime(2026, 6, 25, 5, 0, 0), baseline=True,
          board=[{"job": "412348", "customer": "MOTION", "oper": "53"}])

    assert change_log.load(TODAY) == [], "baseline poll must not write to the change log"


def test_normal_poll_still_logs_a_real_change():
    """A genuine intraday move on a later poll must still produce a change event,
    so the Changes tab keeps showing real changes."""
    change_log.save(TODAY, [])  # start clean
    master = _master_with_enriched_order(so_pdf="S/O.pdf")
    # An order already enriched and present from earlier today.
    state = {"412348": {
        "first_seen": "2026-06-25T05:00:00", "carried_over": True, "enriched": True,
        "present": True, "last_seen": "2026-06-25T05:00:00",
        "job": {"job": "412348", "customer": "MOTION", "oper": "53",
                "co_number": 0, "so_pdf": "S/O.pdf", "end_date": "07/01/2026",
                "so_design_desc": "Pressure Blower", "so_size": "P4",
                "so_arrangement": "4"},
    }}

    # End Date moves mid-day -> exactly one logged change.
    _poll(master, state, datetime(2026, 6, 25, 9, 0, 0), baseline=False,
          board=[{"job": "412348", "customer": "MOTION", "oper": "53",
                  "end_date": "07/20/2026"}])

    events = change_log.load(TODAY)
    assert len(events) == 1, f"expected one change event, got {events}"
    assert events[0]["field"] == "End Date"
    assert events[0]["old"] == "07/01/2026" and events[0]["new"] == "07/20/2026"


def main() -> int:
    passed = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
            print(f"  ok  {name}")
            passed += 1
    print(f"\n{passed} tests passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
