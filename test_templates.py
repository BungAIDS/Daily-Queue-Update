"""Tests for the quote-run TEMPLATE collection (templates.py).

No pytest needed — run it directly:

    python test_templates.py

Covers the pure logic: design-number parsing, which template a run matches
(by design #, extension, and file-name marker), and the field extraction for
each text/xlsx format. The PDF template only delegates to drive_run (covered by
the live PDF path), so here we only assert it WINS the match for .pdf — we never
open a PDF, so pdfplumber isn't required to run this suite.

File names are taken from the real discovery listings (jobs 421473 / 421492).
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

from templates import (
    QuoteRunContext, _design_num, _rtf_to_text, match_template, parse_quote_run,
    kv_from_lines, kv_from_rows, summarize,
    D64WheelConstruction, QtRunText, PdfQuoteRun, GenericTextRun, UnknownRun,
)


def test_design_num():
    assert _design_num("64") == 64
    assert _design_num("36P") == 36          # leading digits only
    assert _design_num(" 95 ") == 95
    assert _design_num("EMSI") is None
    assert _design_num("") is None
    assert _design_num(None) is None
    assert _design_num(64) == 64


def test_match_d64_xlsx_by_design_and_name():
    # The real Design-64 run: the wheel-construction xlsx. Both the design # and
    # the file-name marker point at the D64 template.
    ctx = QuoteRunContext("421492_314-26-1647 D64 Wheel Construction (Inner).xlsx", design="64")
    assert match_template(ctx).key == "d64_wheel_construction"
    # Even without the design # (off-board backfill), the name still wins it.
    ctx2 = QuoteRunContext("421492 D64 Wheel Construction (Inner).xlsx")
    assert match_template(ctx2).key == "d64_wheel_construction"


def test_match_qt_run_text():
    # The real HDX run: a .txt named "Qt Run".
    ctx = QuoteRunContext("421473_909-26-1604 Qt Run.txt")
    assert match_template(ctx).key == "qt_run_text"
    # Quirky spacing from the AutoCAD-folder copies.
    assert match_template(QuoteRunContext("420410 qt  run.txt")).key == "qt_run_text"
    # An .rtf named like a quote run also routes to the text template.
    assert match_template(QuoteRunContext("123 Quote Run.rtf")).key == "qt_run_text"


def test_match_pdf_wins_for_pdf():
    # Any .pdf run is read by the PDF template (delegates to drive_run).
    assert match_template(QuoteRunContext("421473_Quote-34592.pdf")).key == "pdf"
    assert match_template(QuoteRunContext("some construction run.pdf")).key == "pdf"


def test_match_generic_and_unknown_fallbacks():
    # A plain .txt with no quote-run marker still gets read as generic text.
    assert match_template(QuoteRunContext("420410 notes.txt")).key == "generic_text"
    # An extension no template reads -> the Unknown last resort (never crashes).
    assert match_template(QuoteRunContext("run.docx")).key == "unknown"
    assert match_template(QuoteRunContext("run.zip")).key == "unknown"


def test_design_breaks_a_tie_toward_d64():
    # A bare .xlsx with no name marker: only D64 reads .xlsx, so it matches; and
    # a matching design # raises its confidence above the bare-extension score.
    plain = QuoteRunContext("wheel.xlsx")
    d64 = QuoteRunContext("wheel.xlsx", design="64")
    assert match_template(plain).key == "d64_wheel_construction"
    assert D64WheelConstruction().score(d64) > D64WheelConstruction().score(plain)


def test_kv_helpers():
    assert kv_from_lines(["Material: 316 SS", "Gauge = 10", "just prose here"]) == \
        {"Material": "316 SS", "Gauge": "10"}
    # First occurrence of a label wins.
    assert kv_from_lines(["Shaft: A", "Shaft: B"]) == {"Shaft": "A"}
    assert kv_from_rows([["Bearing", "SKF"], ["x", "y", "z"], ["Coating", "epoxy"]]) == \
        {"Bearing": "SKF", "Coating": "epoxy"}


def test_summarize_orders_fields_of_interest_first():
    s = summarize({"Notes": "blah", "Material": "316 SS", "Bearing": "SKF"})
    assert s.index("Material") < s.index("Notes")
    assert s.index("Bearing") < s.index("Notes")


def test_rtf_to_text_strips_markup():
    rtf = r"{\rtf1\ansi\deff0 Material:\tab 316 SS\par Gauge: 10\par}"
    txt = _rtf_to_text(rtf)
    assert "Material:" in txt and "316 SS" in txt and "Gauge: 10" in txt
    assert "\\rtf1" not in txt and "{" not in txt


def test_parse_text_run_end_to_end(tmp: Path):
    p = tmp / "421473_909-26-1604 Qt Run.txt"
    p.write_text("Construction: welded\nWheel Material: 316 SS\nBearing: SKF 22216\n"
                 "Spark Resistant: AMCA Type B\n")
    r = parse_quote_run(p, design="11")
    assert r["template"] == "qt_run_text"
    assert r["fields"]["Wheel Material"] == "316 SS"
    assert r["fields"]["Bearing"] == "SKF 22216"
    assert "Wheel Material=316 SS" in r["summary"]
    assert r["raw_lines"][0] == "Construction: welded"


def test_parse_rtf_run_end_to_end(tmp: Path):
    p = tmp / "123 Quote Run.rtf"
    p.write_text(r"{\rtf1\ansi Coating: epoxy\par Flange: ANSI 150\par}")
    r = parse_quote_run(p)
    assert r["template"] == "qt_run_text"
    assert r["fields"].get("Coating") == "epoxy"
    assert r["fields"].get("Flange") == "ANSI 150"


def test_parse_d64_xlsx_end_to_end(tmp: Path):
    import openpyxl
    p = tmp / "421492_314-26-1647 D64 Wheel Construction (Inner).xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Wheel Type", "Airfoil"])          # adjacent Label | value row
    ws.append(["Blade Material", "316 SS"])
    ws.append(["Backplate", "0.250 in"])
    ws.append(["Notes:", "in-cell label: value"])  # in-cell "Label: value"
    wb.save(str(p))
    r = parse_quote_run(p, design="64")
    assert r["template"] == "d64_wheel_construction"
    assert r["design"] == 64
    assert r["fields"]["Wheel Type"] == "Airfoil"
    assert r["fields"]["Blade Material"] == "316 SS"
    assert r["fields"]["Backplate"] == "0.250 in"
    assert any("Wheel Type" in ln for ln in r["raw_lines"])


def test_parse_unknown_format_is_safe(tmp: Path):
    p = tmp / "run.docx"
    p.write_bytes(b"not really a docx")
    r = parse_quote_run(p)
    assert r["template"] == "unknown"
    assert r["fields"] == {} and r["summary"] == ""


def test_parse_missing_file_never_raises():
    r = parse_quote_run("/no/such/quote run.txt", design="64")
    # Matches a template by name/ext, reads nothing, and returns the empty shape.
    assert r["template"] in {"qt_run_text", "generic_text"}
    assert r["fields"] == {} and r["raw_lines"] == []


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
