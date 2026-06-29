"""Unit tests for the transmittal SO extractor (pure parsing only — no I/O).

The fixtures mirror real CBC Sales-Order text (the 421473 dump used elsewhere in
the suite) and the confirmed filled transmittal for job 421693, so these lock
the field mapping decoded from those documents.
"""
from __future__ import annotations

import json

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


def test_so_read_today_uses_watcher_state(tmp_path, monkeypatch):
    """so_read_today reads the watcher's per-day live_state verified_at."""
    from datetime import date
    monkeypatch.setattr(td, "SNAPSHOT_DIR", tmp_path)
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
