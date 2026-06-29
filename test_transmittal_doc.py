"""Unit tests for the transmittal doc fill PLAN (pure) and the signature lookup.

The Word COM applier itself needs desktop Word and is exercised on the Windows
box; here we lock the semantic plan it consumes.
"""
from __future__ import annotations

import sys
from datetime import date

import engineers
import transmittal_data as td
import transmittal_doc as doc
from test_transmittal_data import SO_473


def _data(**over):
    d = td.build_transmittal_data(SO_473, order="421473", imi_number="IMI-GL-2021")
    for k, v in over.items():
        setattr(d, k, v)
    return d


def test_signature_lookup():
    assert engineers.signature_for_user("dgroth") == "DAG"
    assert engineers.signature_for_user("DGroth") == "DAG"   # case-insensitive
    assert engineers.signature_for_user("ddecker") == "DD"
    assert engineers.signature_for_user("nobody") == ""


def test_plan_basic_fields():
    plan = doc.plan_fill(_data(), initials="DAG", today=date(2026, 6, 29))
    assert plan.order == "421473"
    assert plan.subject == "INNO-VENT INDUSTRIAL INC."
    assert plan.po == "7074-49840-00-AI26"
    assert plan.date == "06/29/2026"
    assert plan.initials == "DAG"
    assert plan.to_emails == ["hmallette@inno-vent.ca"]


def test_plan_box_index_approval_vs_record():
    assert doc.plan_fill(_data(box="approval"), "DAG").box_index == 1
    assert doc.plan_fill(_data(box="record"), "DAG").box_index == 2
    assert doc.plan_fill(_data(box="sales"), "DAG").box_index == 0
    # An unknown box defaults to the approval box, never auto-released.
    assert doc.plan_fill(_data(box="???"), "DAG").box_index == 1


def test_plan_rows_marks_and_text():
    plan = doc.plan_fill(_data(), "DAG")
    # First row is the emailed assembly: EMAIL marked, others blank.
    em, pr, mn, no, rev, desc = plan.rows[0]
    assert (em, pr, mn) == ("X", "", "")
    assert no == "421473-01"
    assert desc.startswith("DESIGN 16A SW FAN ASSEMBLY")
    # IMI O&M row present.
    assert any(r[3] == "IMI-GL-2021" for r in plan.rows)


def test_plan_warns_without_initials():
    plan = doc.plan_fill(_data(), initials="")
    assert any("signature" in w.lower() or "initials" in w.lower() for w in plan.warnings)


def test_default_out_path_naming():
    p = doc.default_out_path("421693")
    assert p.name == "421693 DWG TRANSMITTAL-01.doc"
    assert p.parent.name == "421693"


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
