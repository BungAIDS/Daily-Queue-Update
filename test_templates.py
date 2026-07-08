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
    select_primary_run_text,
    D64WheelConstruction, ChicagoBlowerQtRun, QtRunText,
)


# A real dual-arrangement document (job 413224): the SAME fan quoted as an
# arrangement-4 (motor-mounted, no bearings) and an arrangement-8 (on bearings),
# two printouts in one file. Per DG the arr-4 is the built unit.
REAL_DUAL_4S_8S = """\
---------------------------------------
CONFIDENTIAL
               MON APR  1 12:56:15 CST 2024
               CHICAGO BLOWER CORP.
 SN#413224
 SIZE 3612,DESIGN 5500 ,ARR 4S ,100.0 PCT,DISCH UB ,ROT CW
 EFFECTIVE WHEEL DIA. 37 1/4
   14400 CFM, 16.80 SP, 50.1 BHP, 1770 RPM, 70 DEG F, DENSITY 0.0739
 WHEEL WEIGHT  190 LB, THRUST  120 LB, WR2  240 LB-FT2
 BASE , 326T  FR MOTOR                                     440    3800
---------------------------------------
               MON APR  1 12:56:15 CST 2024
               CHICAGO BLOWER CORP.
 SN#413224
 SIZE 3612,DESIGN 5500 ,ARR 8S ,100.0 PCT,DISCH UB ,ROT CW
 EFFECTIVE WHEEL DIA. 37 1/4
   14400 CFM, 16.80 SP, 50.1 BHP, 1770 RPM, 70 DEG F, DENSITY 0.0739
 ROTOR WR2  55 LB-FT2, ROTOR MAX RPM 1800, MTL. 1045 STEEL
 SHAFT DIA  2 3/16, BRG CENTERS 20, CRITICAL SPEED  4200 RPM
    SIDE    STATIC  DYN. THRUST  L10 HR   P/C
 DRIVE-FIXED   301     15      0  400000 0.0224
---------------------------------------
               MON APR  1 12:56:15 CST 2024
               CHICAGO BLOWER CORP.
 SN#413224
 SIZE 3612,DESIGN 5500 ,ARR 8S ,100.0 PCT,DISCH UB ,ROT CW
      AXIAL VIEW                               IN.          MM
  N   HSG WIDTH OS                            22  3/4       578
"""


def test_select_primary_run_prefers_arr4():
    sel = select_primary_run_text(REAL_DUAL_4S_8S)
    assert "ARR 4S" in sel
    assert "ARR 8S" not in sel                    # the whole 8S run is dropped
    # And the 8S run's bearing/rotor section must NOT leak into the kept text.
    assert "ROTOR WR2" not in sel and "DRIVE-FIXED" not in sel and "SHAFT DIA" not in sel
    # A single-arrangement document is returned unchanged.
    assert select_primary_run_text(REAL_CBC_QT_RUN) == REAL_CBC_QT_RUN


def test_dual_arr4_parse_has_no_8s_bearing_leak():
    f = _parse_chicago_blower(REAL_DUAL_4S_8S)
    assert f["Arrangement"] == "4S"
    assert f["Wheel Weight Lb"] == "190"          # the arr-4's own block
    assert f["Motor Frame"] == "326T"
    # None of the arr-8 fan's bearing/rotor fields are attributed to the arr-4.
    for leaked in ("Shaft Dia", "Brg Centers", "Critical Speed RPM",
                   "Rotor WR2", "Bearing L10 Hr", "Drive Brg Mount"):
        assert leaked not in f, f"8S field {leaked} leaked onto the 4S fan"


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


def test_pdf_selection_program_routes_to_qt_run_parser(tmp: Path):
    # A selection-program Qt Run saved as PDF: the PDF reader hands us the text,
    # and (because the header is present) we parse it with the full Qt Run field
    # set, not the generic key/value sweep. pdfplumber can't run here, so stub
    # the PDF text extraction and check the routing + parse.
    import drive_run
    p = tmp / "421572_300-25-3241 QT Run.pdf"
    p.write_bytes(b"%PDF-1.7 stub")
    orig = drive_run.parse_drive_run_pdf
    drive_run.parse_drive_run_pdf = lambda path: {
        "fields": {"Stray": "generic"}, "raw_lines": REAL_CBC_QT_RUN_421572.splitlines()[:40],
        "summary": "", "text": REAL_CBC_QT_RUN_421572,
    }
    try:
        r = parse_quote_run(p)
    finally:
        drive_run.parse_drive_run_pdf = orig
    assert r["template"] == "pdf"                      # matched by extension
    assert r["fields"]["Size"] == "19"                 # but parsed as a Qt Run
    assert r["fields"]["Liner Material"] == "PLAIN FIRMEX"
    assert "Stray" not in r["fields"]                  # generic KV not used
    assert r["summary"].startswith("Size=19")


def test_pdf_non_selection_program_keeps_generic(tmp: Path):
    # A PDF that isn't the Qt Run layout (no header) keeps the generic fields.
    import drive_run
    p = tmp / "vendor quote run.pdf"
    p.write_bytes(b"%PDF-1.7 stub")
    orig = drive_run.parse_drive_run_pdf
    drive_run.parse_drive_run_pdf = lambda path: {
        "fields": {"Vendor": "Acme", "Total": "$5"}, "raw_lines": ["Vendor: Acme"],
        "summary": "Vendor=Acme", "text": "VENDOR QUOTE\nVendor: Acme\nTotal: $5\n",
    }
    try:
        r = parse_quote_run(p)
    finally:
        drive_run.parse_drive_run_pdf = orig
    assert r["template"] == "pdf"
    assert r["fields"] == {"Vendor": "Acme", "Total": "$5"}   # generic KV preserved


def test_match_pdf_wins_for_pdf():
    # Any .pdf run is read by the PDF template (matched by extension).
    assert match_template(QuoteRunContext("421473_Quote-34592.pdf")).key == "pdf"
    assert match_template(QuoteRunContext("some construction run.pdf")).key == "pdf"


def test_match_generic_and_unknown_fallbacks():
    # A plain .txt with no quote-run marker still gets read as generic text.
    assert match_template(QuoteRunContext("420410 notes.txt")).key == "generic_text"
    # An extension no template reads -> the Unknown last resort (never crashes).
    assert match_template(QuoteRunContext("run.doc")).key == "unknown"   # old-binary Word: no reader yet
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
    # Decimal/parenthetical gauges carry through verbatim from THICK.(GA).
    assert f["Blade Gauge"] == "0.048 (18)"
    assert f["Sideplate Gauge"] == "0.179 ( 7)"
    assert f["Backplate Gauge"] == "1/4"
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
    assert f["Blade Gauge"] == "3/8"
    assert f["Liner Gauge"] == "1/4"
    # This wheel has GUSSETS, not a sideplate/backplate row -> no such gauges.
    assert "Sideplate Gauge" not in f and "Backplate Gauge" not in f
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
    # Wheel-construction gauges (THICK.(GA) column) paired with the materials.
    assert f["Blade Gauge"] == "1/4"
    assert f["Sideplate Gauge"] == "1/4"
    assert f["Backplate Gauge"] == "3/8"
    assert f["Drive"] == "Belt"
    assert f["Engineering Approval"] == "Required"
    assert f["Non-Std Wheel Materials"] == "Yes"
    assert f["Shrink Fit"] == "Yes"
    # The dimension-table noise the generic sweep produced must NOT appear.
    assert "DA" not in f and "DK" not in f and "A" not in f


# The shaft/bearing + outline section from the full job-421579 run (the part
# below BRG CENTERS the earlier trimmed sample dropped): the comma-delimited
# shaft-geometry line (LENGTH/OH/BX/STB/TG&P), STH, the bearing spec block, and
# the AXIAL/SIDE VIEW outline dimensions (N = housing width, F/2 = base to CL).
REAL_CBC_QT_RUN_421579_TAIL = """\
 SHAFT DIA  2 15/16, BRG CENTERS 12, CRITICAL SPEED  5119 RPM
  ROTOR WR2   267 LB-FT2, ROTOR MAX RPM 3000, MTL. 1045 STEEL
  STRESS RATIO AT HUB 0.17, AT BEARING 0.47
  LENGTH 35  3/8 ,OH  8.64,BX 15  3/8 , STB  12     , TG&P  68     880
  STH  2.12
  CHECK MOTOR WR2 >   516 LB-FT2

 SIZE  2 15/16 BEARINGS, LINK BELT SERIES 6800
    SIDE    STATIC  DYN. THRUST  L10 HR   P/C
 DRIVE-FLOAT  1534    33      0  400000 0.0320
 OTHER-FIXED   457    80    183  400000 0.0224              84    3158

      AXIAL VIEW                               IN.          MM
  A   DISCHARGE HEIGHT OS                     37  9/16      953
  W   BOTTOM OF DISCH TO CL                    0              0
 KK   OUTLET AND INLET FLANGE                  2             51
  E   DISCHARGE FLANGE TO CL                  26  1/8       664
 RB   UNITARY BASE TO CL                      69 13/16     1773
 RM   MOTOR CL TO FAN CL                      49  3/16     1249
     CENTER DISTANCE, IN.: 47.94- 54.65
F/2   BASE TO CL                              23  7/8       606
  H   SHAFT HEIGHT                            35  9/16      903
 TV   TOTAL VERT HEIGHT TO DISCH FL           75  1/8      1908
 RH   MAX DIM RIGHT OF CL TO HSG              32  9/16      827
 LH   MAX DIM LEFT OF CL TO DISCH FL          26  1/8       664

      SIDE VIEW                                IN.          MM
 MA   MOUNTING CHANNEL                         2  1/4        57
  D   BASE FLANGE TO CL                       31  7/8       810
  K   DRIVE END OF SHAFT TO CL                39  3/8      1000
  N   HSG WIDTH OS                            22  3/4       578
 LR   MTG FLANGE TO CL                        13  5/8       346
      BEARINGS, TYPE 6800                      2 15/16

 FAN OUTLET AREA   5.934 FT2,   0.551 M2
"""


def test_chicago_blower_shaft_bearing_and_outline_fields():
    f = _parse_chicago_blower(REAL_CBC_QT_RUN_421579_TAIL)
    # Shaft/rotor geometry line (the "BX STB ... and everything near there").
    assert f["Shaft Length"] == "35 3/8"
    assert f["OH"] == "8.64"
    assert f["BX"] == "15 3/8"
    assert f["STB"] == "12"
    assert f["TG&P"] == "68"          # the price column after it is not captured
    assert f["STH"] == "2.12"
    # Bearing spec block.
    assert f["Bearing Size"] == "2 15/16"
    assert f["Bearing Series"] == "LINK BELT SERIES 6800"
    assert f["Bearing L10 Hr"] == "400000"
    # The whole AXIAL/SIDE VIEW outline table — every coded dim, inches value.
    assert f["Discharge Height (A)"] == "37 9/16"
    assert f["Bottom of Disch to CL (W)"] == "0"
    assert f["Outlet/Inlet Flange (KK)"] == "2"
    assert f["Discharge Flange to CL (E)"] == "26 1/8"
    assert f["Unitary Base to CL (RB)"] == "69 13/16"
    assert f["Motor CL to Fan CL (RM)"] == "49 3/16"
    assert f["Base to CL (F)"] == "23 7/8"
    assert f["Shaft Height (H)"] == "35 9/16"
    assert f["Total Vert Height (TV)"] == "75 1/8"
    assert f["Max Right of CL to Hsg (RH)"] == "32 9/16"
    assert f["Max Left of CL to Disch (LH)"] == "26 1/8"
    assert f["Mounting Channel (MA)"] == "2 1/4"
    assert f["Base Flange to CL (D)"] == "31 7/8"
    assert f["Drive End of Shaft to CL (K)"] == "39 3/8"
    assert f["Housing Width (N)"] == "22 3/4"
    assert f["Mtg Flange to CL (LR)"] == "13 5/8"
    # Exactly the 16 outline codes are captured — the "CENTER DISTANCE" note and
    # the "BEARINGS, TYPE 6800" line (no valid code+inches+mm shape) are skipped.
    outline = [k for k in f if k.endswith(")")]
    assert len(outline) == 16
    assert not any("Center Distance" in k for k in f)


# A real Arr-4S run (job 401221): the motor-mounted family (94+ runs — the
# largest) whose back half has NO shaft/bearing section. Instead: the wheel
# weight/thrust/WR2 line, housing-to-CG dims, motor base/frame, housing
# construction, and quote totals — the sections the parser used to miss 100%.
REAL_CBC_QT_RUN_401221 = """\
---------------------------------------
CONFIDENTIAL
               MON JAN 10 10:13:52 CST 2022
               CHICAGO BLOWER CORP.
 401221
 SIZE  4014,DESIGN 6195 ,ARR 4S ,100.0 PCT,DISCH UB ,ROT CW
 EFFECTIVE WHEEL DIA.  38  1/4
   30240 CFM, 14.00 SP,  82.8 BHP, 1770 RPM,  70 DEG F, DENSITY 0.0750
 MAX HP  100.0, MAX RPM 1770, MAX TEMP   70 F, AMBIENT TEMP  70 F
 TIP SPEED 19491 FPM, EQ. TS  20516 RE SK-9-105         WEIGHT   PRICE
 WHEEL         THICK.(GA)    MATERIAL        WR2 WEIGHT
  BLADES          3/16    ASTM A6  XF-100    134    60
  SIDEPL,SPUN  0.135 (10) ASTM A1011-HSLAS    88    41
  BACKPLATE       1/4     ASTM A572 X-TEN    159   104
  HUB 19-5-16   BORE 2  7/8  CAST IRON         7    62     267    9916
    MAX RPM, WHEEL ONLY  2346.  10 BLADES.       RES  3110 CPM
    ***NON STD WHEEL MATERIALS, CHECK FOR CORRECT WELD WIRE***
  HOUSING TO WHEEL CG   5.33, HOUSING TO HUB INLET FACE  6.38 IN.
  STH  2.50
  WHEEL WEIGHT  267 LB, THRUST  190 LB, WR2  387 LB-FT2
 RUN TEST AT FACTORY
 OUTLET N X A:   27  3/4  IN. X   45 13/16 IN.
 BASE , 405T  FR MOTOR                                     579    4722
 MOTOR MOUNTING INCLUDED
 MOTOR WEIGHT, PRICE NOT INCL.                            1575
 STIFFENERS SK-9-236 HI  PRESSURE, 22 IN. CENTERS
 FLANGED INLET  INCLUDED
 HOUSING  1/4  C.Q. HRS                                   1352   17437
                                  GOOD FOR 60 DAYS        3773   32851
 NOTES
     ONLY ITEMIZED OR STANDARD ACCESSORIES INCLUDED IN PRICE
     FAN OUTLET AREA   8.828 FT2
     SHAFT SEAL NOT INCLUDED
 401221
    2 '6195'   63 'UB  '   64 'CW  '  112 '405T'  113 'TEFC'
"""


def test_chicago_blower_arr4_motor_mounted_sections():
    f = _parse_chicago_blower(REAL_CBC_QT_RUN_401221)
    assert f["Arrangement"] == "4S"
    # The arr-4 wheel dynamics / weight block.
    assert f["Wheel Weight Lb"] == "267"
    assert f["Wheel Thrust Lb"] == "190"
    assert f["Wheel WR2"] == "387"
    assert f["Max RPM Wheel Only"] == "2346"     # no trailing dot
    assert f["Blades"] == "10"
    assert f["Wheel Resonance CPM"] == "3110"
    assert f["Housing to Wheel CG"] == "5.33"
    assert f["Housing to Hub Inlet Face"] == "6.38"
    # Motor / base block.
    assert f["Motor Frame"] == "405T"
    assert f["Motor Enclosure"] == "TEFC"
    assert f["Motor Weight Lb"] == "1575"
    assert f["Drive"] == "Motor mounted"
    # Housing / accessories / totals.
    assert f["Housing Construction"] == "1/4 C.Q. HRS"
    assert f["Stiffeners"] == "SK-9-236 HI PRESSURE, 22 IN. CENTERS"
    assert f["Fan Outlet Area FT2"] == "8.828"
    assert f["Flanged Inlet"] == "Included"
    assert f["Shaft Seal"] == "Not included"
    assert f["Total Weight Lb"] == "3773"
    assert f["Total Price"] == "32851"
    # And no shaft/bearing section on this arrangement.
    assert "Shaft Dia" not in f and "Bearing L10 Hr" not in f


def test_chicago_blower_rotor_and_bearing_loads_variants():
    # DRIVE-FIXED (not -FLOAT), negative static load, and the wider row layout.
    txt = ("CHICAGO BLOWER CORP.\n SN#400567\n"
           " SIZE  1908,DESIGN 1904 ,ARR 8S , 90.0 PCT,DISCH UB ,ROT CW\n"
           " ROTOR WR2  28 LB-FT2, ROTOR MAX RPM 3600, MTL. A240 304 SS\n"
           " STRESS RATIO AT HUB 0.08, AT BEARING 0.21\n"
           "    SIDE    STATIC  DYN. THRUST  L10 HR   P/C\n"
           " DRIVE-FIXED    -51     15      0  400000 0.0037\n"
           " OTHER-FLOAT   301      3     71   47882 0.0494\n"
           " SPLIT HOUSING AND BOX  3/8  C.Q. HRS  6591  75987\n")
    f = _parse_chicago_blower(txt)
    assert f["Rotor WR2"] == "28"
    assert f["Rotor Max RPM"] == "3600"
    assert f["Rotor Material"] == "A240 304 SS"
    assert f["Stress Ratio at Hub"] == "0.08"
    assert f["Stress Ratio at Bearing"] == "0.21"
    assert f["Drive Brg Mount"] == "FIXED"
    assert f["Drive Brg Static Lb"] == "-51"     # negative loads happen
    assert f["Other Brg Mount"] == "FLOAT"
    assert f["Other Brg Static Lb"] == "301"
    assert f["Brg Thrust Lb"] == "71"
    assert f["Bearing L10 Hr"] == "400000"       # anchored before the P/C decimal
    assert f["Housing Construction"] == "3/8 C.Q. HRS"   # arr-7 SPLIT... variant


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


def _write_docx(path: Path, text: str) -> None:
    """Minimal .docx: a zip whose word/document.xml holds one <w:p> per line."""
    import zipfile
    paras = "".join(
        f"<w:p><w:r><w:t xml:space=\"preserve\">{ln}</w:t></w:r></w:p>"
        for ln in text.splitlines())
    doc = ("<?xml version=\"1.0\"?><w:document xmlns:w=\"http://x\"><w:body>"
           f"{paras}</w:body></w:document>")
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", doc)


def test_chicago_blower_in_a_docx(tmp: Path):
    # A CB run saved as Word (e.g. "QT RUN.docx") — text is pulled from the zip
    # and routed to the CB template by its content marker.
    p = tmp / "421237 QT RUN.docx"
    _write_docx(p, REAL_CBC_QT_RUN_421237)
    ctx = QuoteRunContext(p)
    assert "CHICAGO BLOWER" in ctx.text()
    assert match_template(ctx).key == "cbc_qt_run_text"
    r = parse_quote_run(p)
    assert r["template"] == "cbc_qt_run_text"
    assert r["fields"]["Size"] == "2412"
    assert r["fields"]["Blade Material"] == "ASTM A1011-HSLAS"


def test_docx_without_cb_marker_falls_to_generic(tmp: Path):
    p = tmp / "quote run.docx"
    _write_docx(p, "Coating: epoxy\nFlange: ANSI 150\n")
    assert match_template(QuoteRunContext(p)).key == "qt_run_text"
    assert parse_quote_run(p)["fields"]["Coating"] == "epoxy"


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
    p = tmp / "run.xyz"
    p.write_bytes(b"some bytes")
    r = parse_quote_run(p)
    assert r["template"] == "unknown"
    assert r["fields"] == {} and r["summary"] == ""


def test_corrupt_docx_is_safe(tmp: Path):
    # A .docx that isn't a valid zip must not raise — it just yields no fields.
    p = tmp / "421000 qt run.docx"
    p.write_bytes(b"not really a docx")
    r = parse_quote_run(p)
    assert r["fields"] == {} and r["template"] in {"qt_run_text", "generic_text"}


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
