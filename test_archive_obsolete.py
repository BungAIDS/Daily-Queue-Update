"""Tests for the obsolete-file archiver (archive_obsolete.py).

No pytest needed — run it directly:

    python test_archive_obsolete.py

Exercises the plan/move/undo logic on synthetic temp folders (the real tool
runs against Z:, which isn't reachable here — so the safety behaviors are what
we prove: dry run moves nothing, active file stays, moves are reversible, and
name clashes never overwrite).
"""
from __future__ import annotations

import os
import sys
import tempfile
import time
from pathlib import Path

import archive_obsolete as ao


def _mk(d: Path, name: str, mtime: float | None = None) -> Path:
    p = d / name
    p.write_text("x")
    if mtime is not None:
        os.utime(p, (mtime, mtime))
    return p


def test_plan_folder_active_vs_obsolete(tmp: Path):
    job = tmp / "421457"; job.mkdir()
    now = time.time()
    _mk(job, "Quote Run.pdf", now - 1000)
    _mk(job, "Quote Run.txt", now - 500)
    co = _mk(job, "Quote Run CO#1.txt", now - 5000)   # oldest on disk, but CO#1
    active, obsolete = ao.plan_folder(job)
    assert active == co                                # CO# beats mtime
    assert {p.name for p in obsolete} == {"Quote Run.pdf", "Quote Run.txt"}
    # A single-file folder has nothing obsolete.
    solo = tmp / "solo"; solo.mkdir(); _mk(solo, "only.txt")
    assert ao.plan_folder(solo)[1] == []


def test_dry_run_moves_nothing(tmp: Path):
    job = tmp / "j"; job.mkdir()
    _mk(job, "a CO#2.txt"); _mk(job, "a.txt"); _mk(job, "a.pdf")
    moves = ao.archive_folder(job, apply=False)
    assert len(moves) == 2                              # 2 would move
    assert not (job / ao.OBSOLETE).exists()             # but nothing actually moved
    assert len(list(job.iterdir())) == 3


def test_apply_moves_then_undo_restores(tmp: Path):
    job = tmp / "j"; job.mkdir()
    _mk(job, "run CO#2.txt"); _mk(job, "run.txt"); _mk(job, "run.pdf")
    moves = ao.archive_folder(job, apply=True)
    obs = job / ao.OBSOLETE
    assert obs.is_dir()
    assert {p.name for p in obs.iterdir()} == {"run.txt", "run.pdf"}
    assert {p.name for p in job.iterdir() if p.is_file()} == {"run CO#2.txt"}   # active stays
    # Undo puts them back.
    import json
    man = tmp / "manifest.json"; man.write_text(json.dumps(moves))
    assert ao.undo(man) == 0
    assert {p.name for p in job.iterdir() if p.is_file()} == {"run CO#2.txt", "run.txt", "run.pdf"}


def test_name_clash_never_overwrites(tmp: Path):
    job = tmp / "j"; job.mkdir()
    obs = job / ao.OBSOLETE; obs.mkdir()
    _mk(obs, "run.pdf")                                 # a prior obsolete copy exists
    _mk(job, "run CO#1.txt"); _mk(job, "run.pdf")       # new obsolete has the same name
    ao.archive_folder(job, apply=True)
    names = {p.name for p in obs.iterdir()}
    assert "run.pdf" in names and "run (1).pdf" in names   # suffixed, not clobbered


def main() -> int:
    passed = 0
    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d)
        for name, fn in sorted(globals().items()):
            if not name.startswith("test_") or not callable(fn):
                continue
            sub = tmp / name; sub.mkdir()
            fn(sub)
            print(f"  ok  {name}")
            passed += 1
    print(f"\n{passed} tests passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
