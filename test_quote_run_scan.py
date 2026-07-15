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

from quote_run_scan import classify_status, is_damper, run_rows, scan_one, CORE_FIELDS


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


def test_run_rows_sorted_and_deduped():
    records = {
        "b": {"job": "421473", "type": "GL", "folder": "f", "runs": [
            # Two genuinely distinct runs (different fans) -> two rows, kept.
            {"file": "quote run.txt", "path": "p1", "template": "cbc_qt_run_text",
             "summary": "", "status": "OK", "fields": {"Size": "37", "Design": "16A"}},
            {"file": "quote run 2.txt", "path": "p2", "template": "cbc_qt_run_text",
             "summary": "", "status": "OK", "fields": {"Size": "40", "Design": "64"}},
            # A format-duplicate of the first (.pdf, same spec) -> collapses away.
            {"file": "quote run.pdf", "path": "p1b", "template": "pdf",
             "summary": "", "status": "OK", "fields": {"Size": "37", "Design": "16A"}},
        ]},
        "a": {"job": "420990", "type": "GL", "folder": "f", "runs": [
            {"file": "qt run.txt", "path": "p3", "template": "cbc_qt_run_text",
             "summary": "", "status": "OK", "fields": {"Size": "12"}},
        ]},
    }
    rows = run_rows(records)
    assert [r["job"] for r in rows] == ["420990", "421473", "421473"]   # sorted; dupe collapsed
    files = {r["file"] for r in rows if r["job"] == "421473"}
    assert files == {"quote run.txt", "quote run 2.txt"}               # both distinct kept


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


def test_run_files_matches_wheel_construction_names(tmp: Path):
    # Broadened name matching (job 421959's "Cascades Wheel Construction REV 2"
    # was missed): the wheel-construction phrase and bare "Wheel"/"Construction"
    # docs must be caught, while a CAD file with a matching name (wrong
    # extension), a plain non-run doc, and a superseded history copy stay out.
    from sales_orders import _run_files_in_folder
    job = tmp / "421959"
    (job / "ENG REF").mkdir(parents=True)
    (job / "history").mkdir(parents=True)
    (job / "ENG REF" / "Cascades Wheel Construction REV 2.docx").write_text("x")
    (job / "Wheel.txt").write_text("x")                  # bare token
    (job / "Construction.doc").write_text("x")           # bare token
    (job / "421959 qt run.txt").write_text("x")          # existing pattern still works
    (job / "Wheel Assembly.dwg").write_text("x")         # name matches, but CAD ext -> out
    (job / "shipping notes.txt").write_text("x")         # not a run
    (job / "history" / "Wheel Construction.txt").write_text("x")  # superseded -> skipped
    got = {p.name for p in _run_files_in_folder(job)}
    assert got == {
        "Cascades Wheel Construction REV 2.docx",
        "Wheel.txt",
        "Construction.doc",
        "421959 qt run.txt",
    }


def test_scan_one_no_runs(tmp: Path):
    job = tmp / "400111"
    job.mkdir()
    (job / "400111-01A.dwg").write_text("drawing only")
    rec = scan_one("400111", "GL", job)
    assert rec["runs"] == []   # job recorded (so resume skips it) but no runs


def test_is_damper():
    assert is_damper("420848 damper quote run.docx")    # in-house damper quote (name)
    assert is_damper("420848 DAMPER RUN.txt")           # any case
    # A fan run that only *mentions* a damper accessory is NOT a damper quote.
    assert not is_damper("421311 qt run.txt")


def test_scan_one_flags_damper(tmp: Path):
    job = tmp / "420848"
    job.mkdir()
    (job / "420848 qt run.txt").write_text("CHICAGO BLOWER\n SIZE 37 ARR 9H\n")  # fan
    (job / "420848 damper quote run.txt").write_text("DAMPER QUOTE\n MODEL CBD\n")  # damper
    rec = scan_one("420848", "GL", job)
    flags = {r["file"]: r["damper"] for r in rec["runs"]}
    assert flags["420848 qt run.txt"] is False
    assert flags["420848 damper quote run.txt"] is True


def test_reparse_stored_applies_new_patterns_offline():
    from quote_run_scan import reparse_stored
    cb_text = ("CHICAGO BLOWER CORP.\n SN#401221\n"
               " SIZE  4014,DESIGN 6195 ,ARR 4S ,100.0 PCT,DISCH UB ,ROT CW\n"
               " WHEEL WEIGHT  267 LB, THRUST  190 LB, WR2  387 LB-FT2\n"
               " BASE , 405T  FR MOTOR    579   4722\n")
    records = {
        # A text run parsed by an OLD parser (fields sparse) — raw_lines stored.
        "401221": {"job": "401221", "runs": [{
            "path": "Z:\\j\\401221 qt run.txt", "template": "cbc_qt_run_text",
            "fields": {"Size": "4014"}, "summary": "", "status": "OK",
            "raw_lines": cb_text.splitlines(),
        }]},
        # A vision run: model fields kept, pattern hits merged over them.
        "406244": {"job": "406244", "runs": [{
            "path": "Z:\\j\\406244 qt run.pdf", "template": "pdf_vision",
            "fields": {"Size": "22", "Oddball": "kept"}, "summary": "", "status": "OK",
            "vision": {"model": "m", "transcript": cb_text},
        }]},
        # A non-CB document: untouched.
        "400111": {"job": "400111", "runs": [{
            "path": "Z:\\j\\notes.txt", "template": "generic_text",
            "fields": {"Coating": "epoxy"}, "summary": "", "status": "OK",
            "raw_lines": ["Coating: epoxy"],
        }]},
    }
    changed = reparse_stored(records)
    assert changed == 2
    txt = records["401221"]["runs"][0]
    assert txt["fields"]["Wheel Weight Lb"] == "267"      # new pattern applied
    assert txt["fields"]["Motor Frame"] == "405T"
    assert txt["fields"]["Serial"] == "401221"
    vis = records["406244"]["runs"][0]
    assert vis["fields"]["Oddball"] == "kept"             # model's extras survive
    assert vis["fields"]["Motor Frame"] == "405T"         # pattern fills the gap
    assert vis["fields"]["Size"] == "22"                  # model wins overlaps (it saw the image)
    assert vis["template"] == "pdf_vision"                # provenance kept
    assert records["400111"]["runs"][0]["fields"] == {"Coating": "epoxy"}


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
