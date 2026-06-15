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
    kv_from_lines, kv_from_rows, summarize, _parse_chicago_blower,
    D64WheelConstruction, ChicagoBlowerQtRun, QtRunText, PdfQuoteRun,
    GenericTextRun, UnknownRun,
)


# A real Chicago Blower "Qt Run" text (job 421579, captured 2026-06-15). The
# leading spaces and column spacing are exactly as the selection program emits.
REAL_CBC_QT_RUN = """\
---------------------------------------
CONFIDENTIAL
               FRI JUN  5 15:25:02 CST 2026
               CHICAGO BLOWER CORP.
 SN#421579
 SIZE  3300,DESIGN 6195 ,ARR 9H ,100.0 PCT,DISCH TH ,ROT CCW
 EFFECTIVE WHEEL DIA.  31  3/8
   22500 CFM, 18.00 SP,  78.3 BHP, 2465 RPM,  70 DEG F, DENSITY 0.0714
 MAX HP  100.0, MAX RPM 2585, MAX TEMP   95 F, AMBIENT TEMP  41 F
 ENGINEERING APPROVAL REQUIRED
 TIP SPEED 21216 FPM, EQ. TS  22372 RE SK-9-105         WEIGHT   PRICE
 WHEEL         THICK.(GA)    MATERIAL        WR2 WEIGHT
  BLADES          1/4     ASTM A572 X-TEN     81    54
  SIDEPL,SPUN     1/4     ASTM A572 X-TEN     72    50
  BACKPLATE       3/8     ASTM CQ HRS A36    108   105
  HUB 19-5-16   BORE 2 11/16 CAST IRON         7    62     270    9071
    SHRINK FIT PRICE INCLUDED
    ***NON STD WHEEL MATERIALS, CHECK FOR CORRECT WELD WIRE***
     WHEEL HUB NOMINAL BORE = 2.6875
 BELT DRIVEN.  FAN SHEAVE   ASSUMED PD  6.2 IN, 4000 FPM    25
 SHAFT DIA  2 15/16, BRG CENTERS 12, CRITICAL SPEED  5148 RPM
  ROTOR WR2   267 LB-FT2, ROTOR MAX RPM 2860, MTL. 1045 STEEL
"""


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


# A second real CB run (job 421237): ribbed blades + decimal/gauge thickness, a
# bare order-number header (no SN#), a "PACKAGED FORCED DRAFT FAN" descriptor, a
# coupling (direct drive), and an FEA flag — the variations the first sample lacked.
REAL_CBC_QT_RUN_421237 = """\
---------------------------------------
CONFIDENTIAL
               MON APR 27 07:51:38 CST 2026
               CHICAGO BLOWER CORP.
 421237
 SIZE  2412,DESIGN 1904 ,ARR 8S , 90.0 PCT,DISCH UB ,ROT CW
 PACKAGED FORCED DRAFT FAN
 EFFECTIVE WHEEL DIA.  26 15/16
   14000 CFM, 28.00 SP,  80.8 BHP, 3550 RPM,  70 DEG F, DENSITY 0.0750
 MAX HP  100.0, MAX RPM 3600, MAX TEMP   70 F, AMBIENT TEMP  90 F
 TIP SPEED 25400 FPM, EQ. TS  23091 RE SK-9-105         WEIGHT   PRICE
 WHEEL         THICK.(GA)    MATERIAL        WR2 WEIGHT
  BLADES/2 RIB 0.048 (18) ASTM A1011-HSLAS    13    15
  SIDEPL,SPUN  0.179 ( 7) ASTM A1011-HSLAS    22    24
  BACKPLATE       1/4     ASTM CQ HRS A36     29    43
  HUB 19-5-1056, Q1 BUSHING 1 15/16, C. IRON   3    25     107
    FEA ANALYSIS REQUIRED                                         7044
 COUPLING   FALK T10, SIZE 1060T, BORE  1.938 NOT INCL.     16
"""


def test_chicago_blower_ribbed_blades_and_descriptor():
    f = _parse_chicago_blower(REAL_CBC_QT_RUN_421237)
    assert f["Serial"] == "421237"          # bare-number header (no SN#)
    assert f["Fan Type"] == "PACKAGED FORCED DRAFT FAN"
    assert f["Size"] == "2412" and f["Design"] == "1904"
    assert f["Arrangement"] == "8S" and f["% Width"] == "90.0"
    assert f["Discharge"] == "UB" and f["Rotation"] == "CW"
    # The materials the fraction-only gauge pattern used to miss:
    assert f["Blade Material"] == "ASTM A1011-HSLAS"
    assert f["Sideplate Material"] == "ASTM A1011-HSLAS"
    assert f["Backplate Material"] == "ASTM CQ HRS A36"
    assert f["Hub"] == "19-5-1056"
    assert f["Coupling"] == "FALK T10"
    assert f["Drive"] == "Direct"           # coupling, not belt
    assert f["FEA Analysis"] == "Required"


# A third real CB run (job 421572): a space-delimited spec line (no commas), an
# LS-class wheel with a LINER row whose material is "PLAIN FIRMEX", and a fan
# CLASS — the variations the comma-delimited PFD samples lacked.
REAL_CBC_QT_RUN_421572 = """\
---------------------------------------
CONFIDENTIAL
               MON JUN  8 09:46:44 CST 2026
               CHICAGO BLOWER CORP.
 421572
 SIZE   19 DESIGN 16A LS   ARR 9H  100.0 PCT DISCH UB  ROT CW
 EFFECTIVE WHEEL DIA.  33
   12500 CFM, 12.00 SP,  45.7 BHP, 1673 RPM, 103 DEG F, DENSITY 0.0690
 MAX HP   50.0, MAX RPM 1673, MAX TEMP  103 F, AMBIENT TEMP  90 F
 TIP SPEED 14454 FPM, EQ. TS  14648 RE SK-9-105         WEIGHT   PRICE
 NEW DESIGN LS CLASS 4
 WHEEL         THICK.(GA)    MATERIAL        WR2 WEIGHT
  BLADES          3/8     ASTM A572 X-TEN     91   108
    CHECK AVAILBILITY OF FIRMEX BEFORE QUOTING FAN
  LINER,WELDED    1/4          PLAIN FIRMEX   61    72
  GUSSETS         5/8     ASTM CQ HRS A36     16    37
  HUB BORE 2 11/16, HUB OD   7.00                          255
 BELT DRIVEN.  FAN SHEAVE   ASSUMED PD  9.1 IN, 4000 FPM    26
 SHAFT DIA  2 11/16, BRG CENTERS 26, CRITICAL SPEED  2214 RPM
"""


def test_chicago_blower_space_delimited_and_liner():
    f = _parse_chicago_blower(REAL_CBC_QT_RUN_421572)
    # Space-delimited spec line parses the same as the comma-delimited one.
    assert f["Size"] == "19" and f["Design"] == "16A"   # "16A LS" -> 16A
    assert f["Arrangement"] == "9H" and f["% Width"] == "100.0"
    assert f["Discharge"] == "UB" and f["Rotation"] == "CW"
    assert f["Effective Wheel Dia"] == "33"
    assert f["Blade Material"] == "ASTM A572 X-TEN"
    assert f["Liner Material"] == "PLAIN FIRMEX"          # the notable wear liner
    assert f["Class"] == "4"
    assert f["Drive"] == "Belt"


def test_chicago_blower_wheel_material_fallback():
    # A run with no construction table, just a single wheel-material line.
    txt = ("CHICAGO BLOWER CORP.\n SN#421473\n"
           " SIZE   37 DESIGN 16A LS   ARR 9H  100.0 PCT DISCH TH  ROT CW\n"
           " WHEEL MATERIAL A569 HRS\n CLASS 3   HRS  WHEEL, WR2  1144 LB-FT2\n")
    f = _parse_chicago_blower(txt)
    assert f["Wheel Material"] == "A569 HRS"
    assert f["Class"] == "3"
    assert "Blade Material" not in f   # no table -> no blade row


def test_chicago_blower_fields():
    f = _parse_chicago_blower(REAL_CBC_QT_RUN)
    assert f["Serial"] == "421579"
    assert f.get("Fan Type") is None        # this run has no descriptor line
    assert f["Hub"] == "19-5-16"
    assert f["Size"] == "3300"
    assert f["Design"] == "6195"
    assert f["Arrangement"] == "9H"
    assert f["% Width"] == "100.0"
    assert f["Discharge"] == "TH"
    assert f["Rotation"] == "CCW"
    assert f["Effective Wheel Dia"] == "31 3/8"
    assert f["CFM"] == "22500"
    assert f["SP"] == "18.00"
    assert f["BHP"] == "78.3"
    assert f["RPM"] == "2465"          # operating RPM, not MAX RPM 2585
    assert f["Air Temp F"] == "70"
    assert f["Density"] == "0.0714"
    assert f["Max HP"] == "100.0"
    assert f["Max RPM"] == "2585"
    assert f["Max Temp F"] == "95"
    assert f["Ambient Temp F"] == "41"
    assert f["Tip Speed FPM"] == "21216"
    assert f["Shaft Dia"] == "2 15/16"
    assert f["Brg Centers"] == "12"
    assert f["Critical Speed RPM"] == "5148"
    assert f["Blade Material"] == "ASTM A572 X-TEN"
    assert f["Sideplate Material"] == "ASTM A572 X-TEN"
    assert f["Backplate Material"] == "ASTM CQ HRS A36"
    assert f["Drive"] == "Belt"
    assert f["Engineering Approval"] == "Required"
    assert f["Non-Std Wheel Materials"] == "Yes"
    assert f["Shrink Fit"] == "Yes"
    # The dimension-table noise the generic sweep produced must NOT appear.
    assert "DA" not in f and "DK" not in f and "A" not in f


def test_chicago_blower_matches_and_parses_end_to_end(tmp: Path):
    # Even named just "QT RUN.txt" (the AutoCAD-folder copy), the CB content
    # marker routes it to the CB template over the generic qt_run_text.
    p = tmp / "QT RUN.txt"
    p.write_text(REAL_CBC_QT_RUN)
    ctx = QuoteRunContext(p)
    assert match_template(ctx).key == "cbc_qt_run_text"
    assert ChicagoBlowerQtRun().score(ctx) > QtRunText().score(ctx)
    r = parse_quote_run(p)
    assert r["template"] == "cbc_qt_run_text"
    assert r["fields"]["Size"] == "3300"
    # CB supplies its own engineering-ordered summary (size first, not material).
    assert r["summary"].startswith("Size=3300")
    assert "CFM=22500" in r["summary"]


def test_non_cbc_text_run_falls_to_generic_qt_template(tmp: Path):
    # A quote-run-named .txt WITHOUT the CB marker stays on qt_run_text.
    p = tmp / "421473 Qt Run.txt"
    p.write_text("Construction: welded\nWheel Material: 316 SS\n")
    assert match_template(QuoteRunContext(p)).key == "qt_run_text"
    assert parse_quote_run(p)["fields"]["Wheel Material"] == "316 SS"


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
