"""Tests for the quote-run folder sweep (quote_run_scan.py).

No pytest needed — run it directly:

    python test_quote_run_scan.py

Covers the pure logic (status classification, the per-run row flattening with
the core/Other field split) and an end-to-end scan of a temp folder tree built
to mirror the real layout — a run in an `ENG REF` subfolder, a superseded copy
in `history`, and a non-run file that must be ignored.
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

from quote_run_scan import classify_status, run_rows, scan_one, CORE_FIELDS


def test_classify_status():
    assert classify_status("cbc_qt_run_text", {"Size": "37"}, ["x"], ".txt") == "OK"
    assert classify_status("unknown", {}, [], ".docx") == "UNRECOGNIZED FORMAT"
    assert classify_status("pdf", {}, [], ".pdf") == "PDF (no text layer)"
    # A readable text run that pulled nothing still flags for tuning.
    assert classify_status("generic_text", {}, ["some text"], ".txt") == "NO FIELDS"


def test_run_rows_splits_core_and_other():
    records = {
        "421572": {
            "job": "421572", "type": "GENERAL LINE", "folder": "Z:\\...\\421572",
            "runs": [{
                "file": "421572 QUOTE RUN.txt", "path": "Z:\\...\\421572 QUOTE RUN.txt",
                "template": "cbc_qt_run_text", "summary": "Size=19; ...", "status": "OK",
                "fields": {"Size": "19", "Liner Material": "PLAIN FIRMEX", "Mystery": "42"},
            }],
        }
    }
    rows = run_rows(records)
    assert len(rows) == 1
    row = rows[0]
    assert row["job"] == "421572" and row["status"] == "OK"
    assert row["core"]["Size"] == "19"
    assert row["core"]["Liner Material"] == "PLAIN FIRMEX"
    assert row["core"]["CFM"] == ""            # a core field with no value is blank
    assert "Mystery=42" in row["other"]        # unknown field -> Other, never dropped
    assert "Size" in CORE_FIELDS               # sanity


def test_run_rows_one_row_per_run_sorted():
    records = {
        "b": {"job": "421473", "type": "GL", "folder": "f", "runs": [
            {"file": "quote run.txt", "path": "p1", "template": "cbc_qt_run_text",
             "summary": "", "status": "OK", "fields": {"Size": "37"}},
            {"file": "history/quote run.txt", "path": "p2", "template": "cbc_qt_run_text",
             "summary": "", "status": "OK", "fields": {"Size": "37"}},
        ]},
        "a": {"job": "420990", "type": "GL", "folder": "f", "runs": [
            {"file": "qt run.txt", "path": "p3", "template": "cbc_qt_run_text",
             "summary": "", "status": "OK", "fields": {"Size": "12"}},
        ]},
    }
    rows = run_rows(records)
    assert [r["job"] for r in rows] == ["420990", "421473", "421473"]  # sorted by job


def test_scan_one_finds_runs_recursively(tmp: Path):
    job = tmp / "421473"
    (job / "ENG REF").mkdir(parents=True)
    (job / "history").mkdir(parents=True)
    (job / "ENG REF" / "421473 quote run.txt").write_text(
        "CHICAGO BLOWER CORP.\n 421473\n"
        " SIZE   37 DESIGN 16A LS   ARR 9H  100.0 PCT DISCH TH  ROT CW\n"
        " WHEEL MATERIAL A569 HRS\n BELT DRIVEN.\n")
    (job / "history" / "quote run.txt").write_text(    # superseded copy: must be skipped
        "CHICAGO BLOWER CORP.\n 421473\n SIZE   37 DESIGN 16A LS   ARR 9H\n")
    (job / "421473-01A.dwg").write_text("not a run")     # must be ignored
    (job / "notes.txt").write_text("not a run either")   # must be ignored

    rec = scan_one("421473", "GENERAL LINE", job)
    assert rec["job"] == "421473" and rec["type"] == "GENERAL LINE"
    files = sorted(r["file"] for r in rec["runs"])
    assert files == ["421473 quote run.txt"]   # live run only — history copy, dwg, notes excluded
    eng = next(r for r in rec["runs"] if r["file"] == "421473 quote run.txt")
    assert eng["template"] == "cbc_qt_run_text"
    assert eng["status"] == "OK"
    assert eng["fields"]["Size"] == "37"
    assert eng["fields"]["Wheel Material"] == "A569 HRS"


def test_scan_one_no_runs(tmp: Path):
    job = tmp / "400111"
    job.mkdir()
    (job / "400111-01A.dwg").write_text("drawing only")
    rec = scan_one("400111", "GL", job)
    assert rec["runs"] == []   # job recorded (so resume skips it) but no runs


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
