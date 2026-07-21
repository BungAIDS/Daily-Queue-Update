"""Unit tests for the transmittal SO extractor (pure parsing only — no I/O).

The fixtures mirror real CBC Sales-Order text (the 421473 dump used elsewhere in
the suite) and the confirmed filled transmittal for job 421693, so these lock
the field mapping decoded from those documents.
"""
from __future__ import annotations

import json
import sys
import tempfile
from pathlib import Path

import transmittal_data as td

# A faithful slice of a real SO's reconstructed lines (job 421473 dump), with the
# order/PO header, Sold To, design header, and the Additional Features / Notes +
# Fan Drawings checklist that drive the transmittal.
SO_473 = [
    "Chicago Blower Corporation Sales Order",
    "Design 16A SW",
    "Order # Rep Ref. # Customer P.O. # Fan Serial Number:",
    "421473 7074-49840-00-AI26 7074-49840-00-AI26",
    "Sold To: Ship To:",
    "INNO-VENT INDUSTRIAL INC. INNO-VENT INDUSTRIAL INC.",
    "Qty Design Size Arrangement Motor Pos Class Rotation Discharge % Width Wheel Type",
    "1 D16A 37 A/9H Z 3 CW TH 100 LS",
    "Include 3D STEP Drawings L 991.00",
    "List Total Each 81,710.00 0.38 31,050.00 2,484.00",
    "Additional Features / Notes:",
    "E-Mail Prints to: hmallette@inno-vent.ca",
    "Run Test - N/A (Send to Sales for Run Test Availability Check)",
    "Fan Drawings:",
    "Emailed Mailed",
    "Fan Drawings Both",
    "O & M X",
    "Motor Prints X",
    "Motor Data Sheets X",
    "Buyout Prints (e.g. silencer, filter, etc.) X",
    "Other",
    "v1.8.1.5 -2-",
]


def test_emails():
    assert td.parse_emails(SO_473) == ["hmallette@inno-vent.ca"]


def test_emails_multiple_addresses_one_line():
    lines = [
        "Additional Features / Notes:",
        "E-Mail Prints to: andrew.crockett@kes.global, sandra.hamze@kes.global bagwell@intcon.net",
        "Fan Drawings:",
    ]
    assert td.parse_emails(lines) == [
        "andrew.crockett@kes.global", "sandra.hamze@kes.global", "bagwell@intcon.net",
    ]


def test_emails_ignores_addresses_outside_notes_block():
    lines = ["Quote contact internal@chicagoblower.com", "Design 16A SW"] + SO_473
    # internal address sits before the Notes block -> not a recipient.
    assert td.parse_emails(lines) == ["hmallette@inno-vent.ca"]


def test_po():
    assert td.parse_po(SO_473) == "7074-49840-00-AI26"


def test_po_single_token():
    lines = [
        "Order # Rep Ref. # Customer P.O. # Fan Serial Number:",
        "421693 EOR058724-F0034883K",
    ]
    assert td.parse_po(lines) == "EOR058724-F0034883K"


def test_customer_from_sold_to():
    assert td.parse_customer(SO_473) == "INNO-VENT INDUSTRIAL INC."


def test_customer_dedupes_unequal_sold_ship():
    # Real 421693 case: Sold To has 'LLC', Ship To doesn't — keep only Sold To.
    lines = ["Sold To: Ship To:",
             "JOHN ZINK COMPANY LLC JOHN ZINK COMPANY"]
    assert td.parse_customer(lines) == "JOHN ZINK COMPANY LLC"


def test_customer_no_repeat_passes_through():
    lines = ["Sold To: Ship To:", "ACME CORP"]
    assert td.parse_customer(lines) == "ACME CORP"


def test_design():
    assert td.parse_design(SO_473) == ("16A", "SW")
    assert td.parse_design(["Design 1904 PFD"]) == ("1904", "PFD")


def test_approval_unreleased_default():
    box, released = td.parse_approval(SO_473)
    assert box == "approval" and released is False


def test_approval_released_from_so_status():
    lines = SO_473 + ["STATUS: APPROVED - RELEASED FOR PRODUCTION"]
    box, released = td.parse_approval(lines)
    assert box == "record" and released is True


def test_approval_released_from_flags():
    box, released = td.parse_approval(SO_473, flags="Released For Production")
    assert box == "record" and released is True


def test_approval_negated_release_is_not_released():
    # Unapproved orders carry conditional boilerplate that MENTIONS release —
    # it must never tick the 'record / released for fabrication' box.
    for negated in [
        "Drawings will not be released for production until approved",
        "HOLD - NOT RELEASED FOR PRODUCTION",
        "Fan is not to be RELEASED FOR FABRICATION before customer approval",
        "Order held pending release: not released for production",
    ]:
        box, released = td.parse_approval(SO_473 + [negated])
        assert box == "approval" and released is False, negated


def test_approval_positive_status_beats_boilerplate():
    # A real APPROVED/RELEASED stamp wins even when conditional boilerplate
    # elsewhere on the SO mentions (negated) release.
    lines = SO_473 + [
        "Drawings will not be released for production until approved",
        "STATUS: APPROVED - RELEASED FOR PRODUCTION",
    ]
    box, released = td.parse_approval(lines)
    assert box == "record" and released is True


def test_approval_status_does_not_pair_across_lines():
    # 'STATUS' ending one line must not pair with 'APPROVED' opening the next.
    lines = ["Fan Drawings: STATUS", "APPROVED VENDOR LIST ATTACHED"]
    box, released = td.parse_approval(lines)
    assert box == "approval" and released is False


def test_approval_evidence_records_the_line():
    box, released, why = td.parse_approval_evidence(
        SO_473 + ["STATUS: APPROVED - RELEASED FOR PRODUCTION"])
    assert released is True
    assert why == "STATUS: APPROVED - RELEASED FOR PRODUCTION"
    box, released, why = td.parse_approval_evidence(SO_473)
    assert released is False and "no APPROVED" in why


def test_board_unapproved_overrides_released_so():
    d = td.build_transmittal_data(
        SO_473 + ["STATUS: APPROVED - RELEASED FOR PRODUCTION"], order="421473")
    assert d.box == "record"
    td.apply_board_approval(d, True)   # board says UNAPPROVED today
    assert d.box == "approval" and d.released is False
    assert any("UNAPPROVED" in w for w in d.warnings)
    assert "board flag" in d.box_evidence


def test_board_flag_none_or_false_leaves_so_result():
    d = td.build_transmittal_data(
        SO_473 + ["STATUS: APPROVED - RELEASED FOR PRODUCTION"], order="421473")
    td.apply_board_approval(d, None)    # order not in today's live state
    td.apply_board_approval(d, False)   # board explicitly not-unapproved
    assert d.box == "record" and d.released is True
    assert not any("UNAPPROVED" in w for w in d.warnings)


def test_fan_drawings_checklist():
    cl = td.parse_fan_drawings(SO_473)
    assert cl["fan_drawings"].lower() == "both"
    assert cl["om"] == "X"
    assert cl["motor_prints"] == "X"
    assert cl["motor_data"] == "X"
    # The parenthetical is stripped from the Buyout Prints mark.
    assert cl["buyout_prints"] == "X"


def test_has_step():
    assert td.has_step(SO_473) is True
    assert td.has_step(["Base Fan", "Motor"]) is False


def test_has_step_tolerates_variant_spellings():
    assert td.has_step(["Include 3-D STEP Drawings L 991.00"]) is True
    assert td.has_step(["Include 3D-STEP Drawings"]) is True
    assert td.has_step(["INCLUDE 3DSTEP DRAWINGS"]) is True
    # The PDF reconstruction can split the phrase across two lines.
    assert td.has_step(["Include 3D", "STEP Drawings L 991.00"]) is True
    # No false positives from ordinary words.
    assert td.has_step(["Stepped inlet cone", "3 D-rings"]) is False


def test_step_requires_explicit_current_so_mention():
    lines = [ln for ln in SO_473 if "STEP" not in ln.upper()]
    d = td.build_transmittal_data(lines, order="421473")
    assert d.include_step is False
    assert not any("3D STEP" in r.description for r in d.drawing_rows)


def test_step_is_withheld_when_so_was_not_read_today():
    d = td.build_transmittal_data(SO_473, order="421473", so_is_current=False)
    assert d.include_step is False
    assert not any("3D STEP" in r.description for r in d.drawing_rows)


def test_build_drawing_rows_matches_421693_recipe():
    # Reproduce the real 421693 transmittal table from its inputs.
    checklist = {"fan_drawings": "Both", "om": "X", "motor_prints": "X", "motor_data": "X"}
    rows = td.build_drawing_rows(
        order="421693", design_no="1904", design_desc="PFD",
        checklist=checklist, include_step=True, imi_number="IMI-HD_A4", suffix="01",
    )
    descs = [(r.drawing_no, r.description) for r in rows]
    assert descs == [
        ("421693-01", "DESIGN 1904 PFD FAN ASSEMBLY (AUTOCAD/PDF)"),
        ("IMI-HD_A4", "FAN OPERATING AND MAINTENANCE MANUAL"),
        ("421693-01", "3D STEP DRAWING"),
        ("", "MOTOR DOCUMENTS TO FOLLOW"),
    ]
    assert all(r.email for r in rows[:3])  # first three are emailed


def test_build_drawing_rows_no_om_no_step():
    rows = td.build_drawing_rows(
        order="421000", design_no="34", design_desc="Vaneaxial",
        checklist={"fan_drawings": "X"}, include_step=False,
    )
    assert [r.description for r in rows] == ["DESIGN 34 Vaneaxial FAN ASSEMBLY (AUTOCAD/PDF)"]


def test_build_transmittal_data_end_to_end():
    d = td.build_transmittal_data(SO_473, order="421473", imi_number="IMI-GL-2021")
    assert d.po == "7074-49840-00-AI26"
    assert d.emails == ["hmallette@inno-vent.ca"]
    assert d.customer == "INNO-VENT INDUSTRIAL INC."
    assert d.design_no == "16A"
    assert d.box == "approval"
    assert d.include_step is True
    assert [r.description for r in d.drawing_rows][0].startswith("DESIGN 16A SW FAN ASSEMBLY")
    # O&M row present -> IMI used, no IMI warning.
    assert d.drawing_rows[1].drawing_no == "IMI-GL-2021"
    assert not any("IMI" in w for w in d.warnings)


def test_warns_when_om_but_no_imi():
    d = td.build_transmittal_data(SO_473, order="421473", imi_number="")
    assert any("IMI" in w for w in d.warnings)


def test_421968_attachments_use_selected_standard_suffix_and_so_step_decision():
    with tempfile.TemporaryDirectory() as tmp:
        folder = Path(tmp)
        for name in [
            "421968-01.dwg", "421968-01.pdf", "421968-01A.pdf",
            "421968-02.dwg", "421968-02.pdf", "421968-51.pdf",
            "421968-010.pdf", "IMI-GL 2021.pdf", "421968 3D STEP.stp",
            "OTHER 421999.stp", "X4219689.step", "readme.txt",
        ]:
            (folder / name).write_bytes(b"")
        sub = folder / "STEP"
        sub.mkdir()
        (sub / "421968.step").write_bytes(b"")

        got = [p.name for p in td.find_attachments(
            folder, "421968", "IMI-GL 2021", drawing_suffix="01")]
        assert got == [
            "421968-01.dwg", "421968-01.pdf", "421968-01A.pdf",
            "IMI-GL 2021.pdf",
        ]

        got_ccw = [p.name for p in td.find_attachments(
            folder, "421968", "IMI-GL 2021", drawing_suffix="02")]
        assert got_ccw == [
            "421968-02.dwg", "421968-02.pdf", "IMI-GL 2021.pdf",
        ]

        got_with_step = [p.name for p in td.find_attachments(
            folder, "421968", "IMI-GL 2021",
            drawing_suffix="01", include_step=True)]
        assert "421968 3D STEP.stp" in got_with_step
        assert "421968.step" in got_with_step
        assert "421968-51.pdf" not in got_with_step
        assert "421968-02.pdf" not in got_with_step
        assert "X4219689.step" not in got_with_step


def test_board_unapproved_reads_live_state():
    from datetime import date
    old_snapshot_dir = td.SNAPSHOT_DIR
    try:
        with tempfile.TemporaryDirectory() as d:
            tmp_path = Path(d)
            td.SNAPSHOT_DIR = tmp_path
            ref = date(2026, 7, 6)
            # No state file -> unknown.
            assert td.board_unapproved("421395", ref) is None
            (tmp_path / "live_state_2026-07-06.json").write_text(json.dumps({
                "421395": {"job": {"job": "421395", "unapproved": True}},
                "421507": {"job": {"job": "421507", "unapproved": False}},
                "421999": {"first_seen": "x"},   # no job dict -> unknown
            }))
            assert td.board_unapproved("421395", ref) is True
            assert td.board_unapproved("421507", ref) is False
            assert td.board_unapproved("421999", ref) is None
            assert td.board_unapproved("400000", ref) is None
    finally:
        td.SNAPSHOT_DIR = old_snapshot_dir


def test_so_read_today_uses_watcher_state():
    """so_read_today reads the watcher's per-day live_state verified_at."""
    from datetime import date
    old_snapshot_dir = td.SNAPSHOT_DIR
    try:
        with tempfile.TemporaryDirectory() as d:
            tmp_path = Path(d)
            td.SNAPSHOT_DIR = tmp_path
            ref = date(2026, 6, 29)
            # No state file yet -> not read today, no timestamp.
            assert td.so_read_today("421693", ref) is False
            assert td.so_last_verified("421693", ref) is None
            # Watcher verified it earlier today -> read today.
            (tmp_path / "live_state_2026-06-29.json").write_text(json.dumps({
                "421693": {"verified_at": "2026-06-29T08:15:00"},
                "421800": {"verified_at": "2026-06-28T16:00:00"},   # yesterday
            }))
            assert td.so_read_today("421693", ref) is True
            assert td.so_last_verified("421693", ref) == "2026-06-29T08:15:00"
            # Verified yesterday only -> not today.
            assert td.so_read_today("421800", ref) is False
    finally:
        td.SNAPSHOT_DIR = old_snapshot_dir


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
