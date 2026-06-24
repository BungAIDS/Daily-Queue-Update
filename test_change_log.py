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
