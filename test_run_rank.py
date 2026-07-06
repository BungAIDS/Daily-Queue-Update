"""Tests for the multiple-quote-run currency ranking (run_rank.py).

No pytest needed — run it directly:

    python test_run_rank.py
"""
from __future__ import annotations

import os
import sys
import tempfile
import time
from pathlib import Path

from run_rank import rank_paths, rank_runs, revision_key


def test_revision_key():
    assert revision_key("Qt Run CO#1.txt") == (1, 0)
    assert revision_key("QT RUN CO # 3.pdf") == (3, 0)
    assert revision_key("QT RUN REV A.pdf") == (0, 1)
    assert revision_key("qt run rev-b.txt") == (0, 2)
    assert revision_key("Qt Run REV 2.txt") == (0, 2)
    assert revision_key("QT RUN 8-11-21.pdf") == (0, 0)      # a date is not a revision
    assert revision_key("421473_909-26-1604 Qt Run.txt") == (0, 0)  # part # noise
    assert revision_key("CONSTRUCTION RUN.txt") == (0, 0)    # 'CO' inside a word
    assert revision_key("") == (0, 0)


def test_rank_runs_the_400567_case():
    # The real shape that motivated this: an old dated run sorts alphabetically
    # ahead of its CO#1 — ranking must put the change-order run first.
    runs = [
        {"file": "QT RUN 8-11-21.pdf", "fields": {"Size": "20"}, "mtime": 1_600_000_000},
        {"file": "Qt Run CO#1.txt", "fields": {"Size": "20"}, "mtime": 1_650_000_000},
        {"file": "QT RUN REV A.pdf", "fields": {"Size": "20"}, "mtime": 1_620_000_000},
    ]
    ranked = rank_runs(runs)
    assert [r["file"] for r in ranked] == [
        "Qt Run CO#1.txt",        # CO# beats everything
        "QT RUN REV A.pdf",       # then REV
        "QT RUN 8-11-21.pdf",     # base run last
    ]


def test_rank_runs_tiebreakers_and_stability():
    # Same revision signals -> newest mtime wins.
    a = {"file": "quote run.txt", "fields": {}, "mtime": 100}
    b = {"file": "quote run.pdf", "fields": {}, "mtime": 200}
    assert rank_runs([a, b])[0] is b
    # Same mtime -> more extracted fields wins.
    c = {"file": "x.txt", "fields": {"Size": "1"}, "mtime": 100}
    d = {"file": "y.txt", "fields": {"Size": "1", "CFM": "2"}, "mtime": 100}
    assert rank_runs([c, d])[0] is d
    # Full tie -> original order preserved (stable).
    e = {"file": "e.txt", "fields": {}, "mtime": 0}
    f = {"file": "f.txt", "fields": {}, "mtime": 0}
    assert rank_runs([e, f]) == [e, f]
    # Legacy records with no mtime key rank fine.
    g = {"file": "g CO#2.txt", "fields": {}}
    assert rank_runs([e, g])[0] is g


def test_rank_paths_by_name_then_disk_mtime(tmp: Path):
    old = tmp / "quote run old.txt"; old.write_text("x")
    new = tmp / "quote run new.txt"; new.write_text("x")
    co = tmp / "quote run CO#1.txt"; co.write_text("x")
    now = time.time()
    os.utime(old, (now - 1000, now - 1000))
    os.utime(new, (now, now))
    os.utime(co, (now - 5000, now - 5000))          # oldest on disk, but CO#1
    ranked = rank_paths([old, new, co])
    assert ranked[0] == co                          # name signal beats mtime
    assert ranked[1] == new and ranked[2] == old    # then newest first
    missing = tmp / "gone.txt"                      # stat() failure ranks last, no raise
    assert rank_paths([new, missing])[0] == new


def main() -> int:
    passed = 0
    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d)
        for name, fn in sorted(globals().items()):
            if not name.startswith("test_") or not callable(fn):
                continue
            (fn(tmp) if "tmp" in fn.__code__.co_varnames else fn())
            print(f"  ok  {name}")
            passed += 1
    print(f"\n{passed} tests passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
