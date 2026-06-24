"""Tests for the live watcher's intraday state (live_state.py).

No pytest needed — run it directly:

    python test_live_state.py

Covers the pure dict logic the watcher relies on: detecting new orders, stamping
a stable first-seen time, refreshing volatile board fields without clobbering
enrichment, baseline vs announced polls, removed/returning transitions, seeding
from the morning snapshot, and the newest-first present-orders ordering.
"""
from __future__ import annotations

import sys
from datetime import datetime

import live_state


def _job(num, **kw):
    j = {"job": num, "item": "47-0-0000", "customer": "ACME", "end_date": "06/20/2026",
         "total_price": "$1,000.00"}
    j.update(kw)
    return j


T0 = datetime(2026, 6, 16, 9, 0, 0)
T1 = datetime(2026, 6, 16, 9, 2, 0)
T2 = datetime(2026, 6, 16, 9, 4, 0)


def test_new_job_numbers():
    state = {"100": {"job": {"job": "100"}}}
    board = [_job("100"), _job("200"), _job("300"), _job("200")]
    assert live_state.new_job_numbers(state, board) == ["200", "300"]


def test_baseline_poll_marks_carried_over_not_announced():
    state = {}
    deltas = live_state.record_poll(state, [_job("100"), _job("200")], T0, baseline=True)
    # Returned as "new" so the caller enriches them, but flagged carried_over so
    # they aren't announced and have no precise add time.
    assert set(deltas["new"]) == {"100", "200"}
    assert state["100"]["carried_over"] is True
    assert state["100"]["enriched"] is False
    assert state["100"]["present"] is True


def test_new_order_on_later_poll_is_announced():
    state = {}
    live_state.record_poll(state, [_job("100")], T0, baseline=True)
    deltas = live_state.record_poll(state, [_job("100"), _job("200")], T1)
    assert deltas["new"] == ["200"]
    assert state["200"]["carried_over"] is False
    assert state["200"]["first_seen"] == T1.isoformat(timespec="seconds")


def test_first_seen_is_stable_and_board_fields_refresh():
    state = {}
    live_state.record_poll(state, [_job("200", end_date="06/20/2026")], T0, baseline=True)
    state["200"]["enriched"] = True
    state["200"]["job"]["so_design_desc"] = "Vaneaxial"   # enrichment field
    first_seen = state["200"]["first_seen"]
    # Next poll: the End Date changed on the board.
    live_state.record_poll(state, [_job("200", end_date="06/25/2026")], T1)
    assert state["200"]["first_seen"] == first_seen            # never moves
    assert state["200"]["job"]["end_date"] == "06/25/2026"     # volatile field refreshed
    assert state["200"]["job"]["so_design_desc"] == "Vaneaxial"  # enrichment preserved


def test_removed_then_returning():
    state = {}
    live_state.record_poll(state, [_job("100"), _job("200")], T0, baseline=True)
    d1 = live_state.record_poll(state, [_job("100")], T1)       # 200 drops off
    assert d1["removed"] == ["200"]
    assert state["200"]["present"] is False
    d2 = live_state.record_poll(state, [_job("100"), _job("200")], T2)  # 200 is back
    assert d2["returning"] == ["200"]
    assert d2["new"] == []                                      # not a brand-new order
    assert state["200"]["present"] is True


def test_seed_from_snapshot():
    state = {}
    n = live_state.seed_from_snapshot(state, [_job("100"), _job("200")], T0.isoformat())
    assert n == 2
    assert state["100"]["enriched"] is True
    assert state["100"]["carried_over"] is True
    # Already-present entries are never overwritten by a re-seed (restart-safe).
    state["100"]["enriched"] = False
    assert live_state.seed_from_snapshot(state, [_job("100")], T1.isoformat()) == 0
    assert state["100"]["enriched"] is False


def test_present_jobs_orders_newest_first_carried_last():
    state = {}
    live_state.record_poll(state, [_job("100"), _job("200")], T0, baseline=True)  # carried
    live_state.record_poll(state, [_job("100"), _job("200"), _job("300")], T1)    # 300 new
    live_state.record_poll(state, [_job("100"), _job("200"), _job("300"), _job("400")], T2)  # 400 new
    state["300"]["present"] = state["400"]["present"] = True
    rows = live_state.present_jobs(state)
    order = [r["job"] for r in rows]
    # Newest real arrival first (400, then 300), carried-over baseline last (100, 200).
    assert order[0] == "400"
    assert order[1] == "300"
    assert set(order[2:]) == {"100", "200"}
    assert rows[0]["_carried_over"] is False
    assert rows[-1]["_carried_over"] is True


def test_mark_enriched():
    state = {}
    live_state.record_poll(state, [_job("200")], T1)
    enriched = [_job("200", so_design_desc="Plug Fan")]
    live_state.mark_enriched(state, enriched)
    assert state["200"]["enriched"] is True
    assert state["200"]["job"]["so_design_desc"] == "Plug Fan"


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
