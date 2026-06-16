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
