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

from run_rank import dedupe_runs, rank_paths, rank_runs, revision_key


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


def test_rank_runs_parsed_beats_unparsed_over_grab():
    # A broad name pattern ("wheel"/"construction") can sweep in a stray non-run
    # doc. If that doc was edited more recently than the real run, mtime alone
    # would let it lead the order and blank its fields. A run that parsed must
    # win over an empty match at the same revision level regardless of mtime.
    real = {"file": "Cascades Wheel Construction REV 2.docx",
            "fields": {"Size": "37", "Design": "16A"}, "mtime": 100}
    stray = {"file": "Wheel Balance.xlsx", "fields": {}, "mtime": 999}  # newer, junk
    assert rank_runs([stray, real])[0] is real
    # But a genuine change-order rerun still supersedes even before it parsed:
    # CO#/REV rank ahead of the parsed-over-empty tiebreak.
    co = {"file": "Wheel Construction CO#1.pdf", "fields": {}, "mtime": 50}
    assert rank_runs([real, co])[0] is co


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


def test_dedupe_runs_collapses_dupes_keeps_distinct():
    # Same fan as .txt + .pdf, plus an old CO of the same fan -> one row (the
    # current .txt, ranked by CO# then format/mtime). A different fan stays.
    txt = {"file": "Quote Run.txt", "fields": {"Size": "20", "Design": "16A", "Arrangement": "9H"},
           "mtime": 200}
    pdf = {"file": "Quote Run.pdf", "fields": {"Size": "20", "Design": "16A", "Arrangement": "9H"},
           "mtime": 100}
    co = {"file": "Quote Run CO#1.txt", "fields": {"Size": "20", "Design": "16A", "Arrangement": "9H"},
          "mtime": 150}
    other = {"file": "Damper.txt", "fields": {"Size": "40", "Design": "64", "Arrangement": "4S"},
             "mtime": 50}
    got = dedupe_runs([txt, pdf, co, other])
    files = {r["file"] for r in got}
    assert files == {"Quote Run CO#1.txt", "Damper.txt"}   # CO#1 wins its group; other kept
    # Unparsed runs fall back to file identity (never merged blindly).
    a = {"file": "a.txt", "fields": {}}
    b = {"file": "b.txt", "fields": {}}
    assert len(dedupe_runs([a, b])) == 2


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
