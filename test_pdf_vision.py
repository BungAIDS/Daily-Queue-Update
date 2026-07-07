"""Tests for the Claude-vision scanned-PDF reader (pdf_vision.py).

No pytest needed — run it directly:

    python test_pdf_vision.py

Covers the pure logic only (response parsing, folding a result into a run
record, candidate selection, and the rescan carry-forward in quote_run_scan) —
no network, no API key, no PDF rendering.
"""
from __future__ import annotations

import sys

from pdf_vision import (apply_vision_result, candidate_runs,
                        parse_vision_response, build_prompt,
                        NO_TEXT_STATUS, DRAWING_STATUS)
from quote_run_scan import carry_vision_forward


def test_parse_clean_json():
    p = parse_vision_response(
        '{"doc_type": "quote_run", "fields": {"Size": 3300, "BX": "15  3/8"}, '
        '"note": "Qt Run", "transcript": "CHICAGO BLOWER CORP.\\nSIZE 3300"}')
    assert p["doc_type"] == "quote_run"
    assert p["fields"]["Size"] == "3300"          # numbers normalized to strings
    assert p["fields"]["BX"] == "15 3/8"          # runs of spaces collapsed
    assert p["note"] == "Qt Run"
    assert p["transcript"] == "CHICAGO BLOWER CORP.\nSIZE 3300"
    # A reply without a transcript still parses (older/partial replies).
    assert parse_vision_response('{"doc_type": "other", "fields": {}}')["transcript"] == ""


def test_parse_fenced_and_padded_json():
    fenced = "```json\n{\"doc_type\": \"drawing\", \"fields\": {}, \"note\": \"outline dwg\"}\n```"
    assert parse_vision_response(fenced)["doc_type"] == "drawing"
    padded = 'Sure! Here is the JSON:\n{"doc_type": "other", "fields": {}, "note": "a letter"}\nHope that helps.'
    assert parse_vision_response(padded)["doc_type"] == "other"


def test_parse_garbage_is_error_not_exception():
    assert parse_vision_response("I could not read this document.")["doc_type"] == "error"
    assert parse_vision_response('{"doc_type": "quote_run", "fields": ')["doc_type"] == "error"
    assert parse_vision_response("")["doc_type"] == "error"
    # An unknown doc_type collapses to "other", never an invalid value.
    assert parse_vision_response('{"doc_type": "banana", "fields": {}}')["doc_type"] == "other"


def test_apply_quote_run_result():
    run = {"path": "Z:\\j\\406244 qt run.pdf", "status": NO_TEXT_STATUS,
           "template": "pdf", "fields": {}, "summary": ""}
    parsed = {"doc_type": "quote_run",
              "fields": {"Size": "3300", "CFM": "22500"}, "note": "scanned Qt Run",
              "transcript": "SIZE 3300, ARR 9H\n22500 CFM"}
    assert apply_vision_result(run, parsed, "claude-haiku-4-5") is True
    assert run["status"] == "OK" and run["template"] == "pdf_vision"
    assert run["fields"]["CFM"] == "22500"
    assert "Size=3300" in run["summary"]
    assert run["vision"]["model"] == "claude-haiku-4-5"
    # The full transcription is stored, so re-parsing later is free.
    assert run["vision"]["transcript"] == "SIZE 3300, ARR 9H\n22500 CFM"


def test_apply_drawing_and_error_results():
    run = {"path": "p.pdf", "status": NO_TEXT_STATUS, "fields": {}}
    assert apply_vision_result(run, {"doc_type": "drawing", "fields": {},
                                     "note": "dim sketch"}, "m") is True
    assert run["status"] == DRAWING_STATUS and run["fields"] == {}

    # An error leaves the run untouched, so the next batch retries it for free.
    run2 = {"path": "p2.pdf", "status": NO_TEXT_STATUS, "fields": {}}
    assert apply_vision_result(run2, {"doc_type": "error", "fields": {},
                                      "note": "API error"}, "m") is False
    assert run2["status"] == NO_TEXT_STATUS and "vision" not in run2


def test_candidate_selection():
    records = {
        "1": {"job": "1", "runs": [
            {"path": "a.pdf", "status": NO_TEXT_STATUS},              # wanted
            {"path": "b.txt", "status": NO_TEXT_STATUS},              # not a pdf
            {"path": "c.pdf", "status": "OK"},                        # already parsed
        ]},
        "2": {"job": "2", "runs": [
            {"path": "d.pdf", "status": "OK", "vision": {"model": "m"}},  # done (vision)
        ]},
    }
    got = candidate_runs(records)
    assert [(j, r["path"]) for j, r in got] == [("1", "a.pdf")]
    # --redo also re-reads runs that already have a vision result.
    got_redo = candidate_runs(records, redo=True)
    assert ("2", records["2"]["runs"][0]) in got_redo
    # Explicit jobs filter.
    assert candidate_runs(records, jobs=["2"]) == []


def test_carry_vision_forward_survives_rescan():
    old = {"job": "406244", "runs": [{
        "path": "Z:\\j\\406244 qt run.pdf", "template": "pdf_vision",
        "fields": {"Size": "22"}, "summary": "Size=22", "status": "OK",
        "vision": {"model": "m", "doc_type": "quote_run"},
    }]}
    # What a fresh rescan produces: pdfplumber still gets nothing from the scan.
    new = {"job": "406244", "runs": [{
        "path": "Z:\\j\\406244 qt run.pdf", "template": "pdf",
        "fields": {}, "summary": "", "status": NO_TEXT_STATUS,
    }]}
    merged = carry_vision_forward(old, new)
    r = merged["runs"][0]
    assert r["status"] == "OK" and r["fields"] == {"Size": "22"}
    assert r["template"] == "pdf_vision" and r["vision"]["model"] == "m"

    # A run the text parser CAN now read keeps the fresh (better) parse.
    new2 = {"job": "406244", "runs": [{
        "path": "Z:\\j\\406244 qt run.pdf", "template": "pdf",
        "fields": {"Size": "23"}, "summary": "Size=23", "status": "OK",
    }]}
    assert carry_vision_forward(old, new2)["runs"][0]["fields"] == {"Size": "23"}
    # No prior record -> unchanged (fresh dict: the helper mutates in place).
    new3 = {"job": "406244", "runs": [{
        "path": "Z:\\j\\406244 qt run.pdf", "template": "pdf",
        "fields": {}, "summary": "", "status": NO_TEXT_STATUS,
    }]}
    assert carry_vision_forward(None, new3)["runs"][0]["status"] == NO_TEXT_STATUS


def test_vision_qc_repairs_and_flags():
    from pdf_vision import apply_vision_qc, CHECK_STATUS
    # Garbled model value + clean transcript -> repaired in place, not flagged.
    clean_line = "CHICAGO BLOWER CORP.\n 24100 CFM, 22.00 SP, 198.5 BHP, 1770 RPM,  70 DEG F, DENSITY 0.0750\n"
    run = {"status": "OK", "fields": {"CFM": "4/100", "RPM": "1770"},
           "vision": {"model": "m", "transcript": clean_line}}
    notes = apply_vision_qc(run)
    assert run["fields"]["CFM"] == "24100"           # repaired from transcript
    assert run["status"] == "OK"                     # repair is not a flag
    assert any(n.startswith("repaired CFM") for n in notes)
    assert run["vision"]["suspect"] == []

    # Garbled value with NO clean source -> flagged CHECK VISION.
    run2 = {"status": "OK", "fields": {"CFM": "4/100"},
            "vision": {"model": "m", "transcript": "CHICAGO BLOWER\nno cfm here"}}
    apply_vision_qc(run2)
    assert run2["status"] == CHECK_STATUS
    assert any("implausible CFM" in r for r in run2["vision"]["suspect"])

    # Missing transcript (trial batch) -> flagged for a cheap re-read.
    run3 = {"status": "OK", "fields": {"Size": "22"}, "vision": {"model": "m"}}
    apply_vision_qc(run3)
    assert run3["status"] == CHECK_STATUS

    # Odd arrangement (OCR garble) -> flagged.
    run4 = {"status": "OK", "fields": {"Arrangement": "781"},
            "vision": {"model": "m", "transcript": "CHICAGO BLOWER"}}
    apply_vision_qc(run4)
    assert any("odd Arrangement" in r for r in run4["vision"]["suspect"])

    # Model and a clean transcript disagree on a hard number -> flagged.
    run5 = {"status": "OK", "fields": {"CFM": "21900"},
            "vision": {"model": "m", "transcript": clean_line}}
    apply_vision_qc(run5)
    assert any("model read 21900" in r for r in run5["vision"]["suspect"])

    # A clean run passes and a previously-flagged clean run is unflagged.
    run6 = {"status": CHECK_STATUS, "fields": {"CFM": "24100", "RPM": "1770"},
            "vision": {"model": "m", "transcript": clean_line}}
    apply_vision_qc(run6)
    assert run6["status"] == "OK" and run6["vision"]["suspect"] == []

    # Non-vision runs are untouched.
    run7 = {"status": "OK", "fields": {"CFM": "bad/val"}}
    assert apply_vision_qc(run7) == [] and run7["status"] == "OK"


def test_qc_flagged_runs_are_reread_candidates():
    from pdf_vision import CHECK_STATUS
    records = {"1": {"job": "1", "runs": [
        {"path": "a.pdf", "status": CHECK_STATUS, "vision": {"model": "m"}},
        {"path": "b.pdf", "status": "OK", "vision": {"model": "m"}},
    ]}}
    got = candidate_runs(records)
    assert [r["path"] for _, r in got] == ["a.pdf"]   # flagged re-reads without --redo


def test_prompt_carries_field_names_and_contract():
    p = build_prompt()
    for name in ("Blade Gauge", "BX", "STB", "Housing", "Hub")[:3]:
        assert name in p
    assert "doc_type" in p and "quote_run" in p and "drawing" in p


def main() -> int:
    passed = 0
    for name, fn in sorted(globals().items()):
        if not name.startswith("test_") or not callable(fn):
            continue
        fn()
        print(f"  ok  {name}")
        passed += 1
    print(f"\n{passed} tests passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
