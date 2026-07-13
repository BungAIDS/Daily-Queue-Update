"""Tests for the per-day change log (change_log.py).

    python test_change_log.py
"""
from __future__ import annotations

import sys
from datetime import date

import change_log


def test_append_and_load():
    d = date(2025, 1, 2)              # a date unlikely to collide with real logs
    try:
        change_log.save(d, [])       # start clean
        assert change_log.load(d) == []
        e1 = {"time": "t1", "job": "100", "field": "End Date", "old": "a", "new": "b"}
        full = change_log.append(d, [e1])
        assert full == [e1]
        e2 = {"time": "t2", "job": "100", "field": "End Date", "old": "b", "new": "c"}
        full = change_log.append(d, [e2])
        assert full == [e1, e2]       # same field, two events -> both kept (two lines)
        assert change_log.append(d, []) == [e1, e2]   # empty append is a no-op read
    finally:
        p = change_log.log_path(d)
        if p.exists():
            p.unlink()


def test_scrub_phantom_blanks():
    d = date(2025, 1, 3)              # a date unlikely to collide with real logs
    # Master still holds Size="245" for job 100 (so the '-> blank' never stuck)
    # and an empty Note (so the Note blanking was real).
    master = {"orders": {"100": {"job": {"job": "100", "so_size": "245",
                                         "status_note": ""}}}}
    phantom = {"time": "t1", "job": "100", "field": "Size", "old": "245", "new": ""}
    legit_blank = {"time": "t2", "job": "100", "field": "Note", "old": "HOLD", "new": ""}
    normal = {"time": "t3", "job": "100", "field": "End Date", "old": "a", "new": "b"}
    unknown_job = {"time": "t4", "job": "999", "field": "Size", "old": "50", "new": ""}
    try:
        change_log.save(d, [phantom, dict(phantom, time="t5"),
                            legit_blank, normal, unknown_job])
        assert change_log.scrub_phantom_blanks(d, master) == 2
        # Real blankings, real changes and unverifiable jobs all survive.
        assert change_log.load(d) == [legit_blank, normal, unknown_job]
        assert change_log.scrub_phantom_blanks(d, master) == 0   # idempotent
    finally:
        p = change_log.log_path(d)
        if p.exists():
            p.unlink()


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
