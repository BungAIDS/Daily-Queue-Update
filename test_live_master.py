"""Tests for the all-time master log (live_master.py).

    python test_live_master.py

Covers the upsert: first-seen 'added' is set once and stable, returning orders
clear 'left' and flip on_queue back on, departing orders get stamped 'left' and
on_queue=False, and ordered() is chronological by added.
"""
from __future__ import annotations

import sys
from datetime import datetime

import live_master as lm

T0 = datetime(2026, 6, 16, 9, 0, 0)
T1 = datetime(2026, 6, 16, 9, 2, 0)
T2 = datetime(2026, 6, 16, 9, 4, 0)


def _job(num, **kw):
    j = {"job": num, "customer": "ACME", "end_date": "06/20/2026"}
    j.update(kw)
    return j


def test_append_and_added_stable():
    m = {"orders": {}}
    lm.update(m, [_job("100", _first_seen="2026-06-16T09:00:00")], T0)
    assert m["orders"]["100"]["added"] == "2026-06-16T09:00:00"
    assert m["orders"]["100"]["on_queue"] is True and m["orders"]["100"]["left"] is None
    # A later poll with changed data keeps 'added' fixed.
    lm.update(m, [_job("100", end_date="06/25/2026")], T1)
    assert m["orders"]["100"]["added"] == "2026-06-16T09:00:00"
    assert m["orders"]["100"]["job"]["end_date"] == "06/25/2026"


def test_leave_then_return():
    m = {"orders": {}}
    lm.update(m, [_job("100"), _job("200")], T0)
    lm.update(m, [_job("100")], T1)                  # 200 drops off
    assert m["orders"]["200"]["on_queue"] is False
    assert m["orders"]["200"]["left"] == T1.isoformat(timespec="seconds")
    lm.update(m, [_job("100"), _job("200")], T2)     # 200 returns
    assert m["orders"]["200"]["on_queue"] is True
    assert m["orders"]["200"]["left"] is None


def test_ordered_is_chronological_and_on_queue_filter():
    m = {"orders": {}}
    lm.update(m, [_job("100", _first_seen="2026-06-16T09:00:00")], T0)
    lm.update(m, [_job("100"), _job("300", _first_seen="2026-06-16T09:02:00")], T1)
    lm.update(m, [_job("300")], T2)                  # 100 leaves
    keys = [k for k, _ in lm.ordered(m)]
    assert keys == ["100", "300"]                    # oldest-added first
    assert [j["job"] for j in lm.on_queue(m)] == ["300"]


def test_update_logs_field_modifications():
    m = {"orders": {}}
    lm.update(m, [_job("100", end_date="06/20/2026", total_price="$1,000.00")], T0)
    # No events on first sight (it's new, not modified).
    ev = lm.update(m, [_job("100", end_date="06/25/2026", total_price="$1,000.00")], T1)
    assert len(ev) == 1
    e = ev[0]
    assert e["job"] == "100" and e["field"] == "End Date"
    assert e["old"] == "06/20/2026" and e["new"] == "06/25/2026"
    assert e["time"] == T1.isoformat(timespec="seconds")


def test_update_tracks_co_and_skips_initial_population():
    m = {"orders": {}}
    # First sight has no SO size yet; later it's enriched (''-> value): NOT a change.
    lm.update(m, [_job("100", co_number=0)], T0)
    ev = lm.update(m, [_job("100", co_number=0, so_size="M2")], T1)
    assert not any(x["field"] == "Size" for x in ev)        # initial population skipped
    # A real CO# bump (0 -> 1) is logged.
    ev2 = lm.update(m, [_job("100", co_number=1, so_size="M2")], T2)
    co = [x for x in ev2 if x["field"] == "CO#"]
    assert co and co[0]["old"] == "0" and co[0]["new"] == "1"
    # And a real Size modification (M2 -> M3) is logged.
    ev3 = lm.update(m, [_job("100", co_number=1, so_size="M3")], T2)
    assert any(x["field"] == "Size" and x["old"] == "M2" and x["new"] == "M3" for x in ev3)


def test_merge_order_adds_and_never_regresses():
    m = {"orders": {}}
    # A backlog order we've never seen on the board is created off-queue.
    assert lm.merge_order(m, "900", {"so_size": "M2", "dwg_extras": {"51": "x"}}) is True
    e = m["orders"]["900"]
    assert e["on_queue"] is False and e["job"]["so_size"] == "M2"
    assert e["job"]["dwg_extras"] == {"51": "x"}
    # A sparse source must not wipe an existing value with an empty one.
    assert lm.merge_order(m, "900", {"so_size": "", "customer": "ACME"}) is True
    assert e["job"]["so_size"] == "M2" and e["job"]["customer"] == "ACME"
    # No change -> returns False.
    assert lm.merge_order(m, "900", {"so_size": "M2"}) is False
    # Merging onto a live (on-queue) order keeps it on-queue.
    lm.update(m, [_job("900")], T0)
    assert m["orders"]["900"]["on_queue"] is True
    lm.merge_order(m, "900", {"so_arrangement": "9"})
    assert m["orders"]["900"]["on_queue"] is True
    assert m["orders"]["900"]["job"]["so_arrangement"] == "9"


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
