"""Unit tests for the transmittal doc fill PLAN (pure) and the signature lookup.

The Word COM applier itself needs desktop Word and is exercised on the Windows
box; here we lock the semantic plan it consumes.
"""
from __future__ import annotations

import sys
import tempfile
from datetime import date
from pathlib import Path

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


# The three checkbox paragraphs as they read in DWG TRANSMITTAL MASTER.doc.
_BOX_PARAS = [
    "For sales purposes only – and is not specialized.",
    "For approval only, certified – and not released for fabrication.",
    "For record only, certified – and released for fabrication.",
]


def test_pick_checkbox_by_label():
    assert doc.pick_checkbox_index(_BOX_PARAS, "sales", 0) == 0
    assert doc.pick_checkbox_index(_BOX_PARAS, "approval", 1) == 1
    assert doc.pick_checkbox_index(_BOX_PARAS, "record", 2) == 2


def test_pick_checkbox_label_beats_position():
    # If Word enumerates the fields in an unexpected order (or one is missing),
    # the label still finds the right box — position is only the fallback.
    shuffled = [_BOX_PARAS[2], _BOX_PARAS[0], _BOX_PARAS[1]]
    assert doc.pick_checkbox_index(shuffled, "record", 2) == 0
    assert doc.pick_checkbox_index(shuffled, "approval", 1) == 2
    two_only = [_BOX_PARAS[1], _BOX_PARAS[2]]   # first field dropped
    assert doc.pick_checkbox_index(two_only, "approval", 1) == 0
    assert doc.pick_checkbox_index(two_only, "record", 2) == 1


def test_pick_checkbox_falls_back_when_labels_unreadable():
    assert doc.pick_checkbox_index(["", "", ""], "record", 2) == 2
    assert doc.pick_checkbox_index([], "approval", 1) == 1
    # Ambiguous (two 'record only' paragraphs) -> positional fallback.
    dup = [_BOX_PARAS[2], _BOX_PARAS[2], _BOX_PARAS[0]]
    assert doc.pick_checkbox_index(dup, "record", 2) == 2


def test_default_out_path_naming():
    p = doc.default_out_path("421693")
    assert p.name == "421693 DWG TRANSMITTAL-01.doc"
    assert p.parent.name == "421693"


def _touch(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"")


def test_next_suffix_first_when_nothing_sent():
    with tempfile.TemporaryDirectory() as d:
        tdir = Path(d) / "TRANSMITTAL"
        tdir.mkdir()
        # A draft .doc is NOT proof of sending — it must not bump the number.
        _touch(tdir / "AUTO-GENERATED" / "419624 DWG TRANSMITTAL-01.doc")
        assert doc.sent_transmittal_numbers(tdir, "419624") == []
        assert doc.next_transmittal_suffix(tdir, "419624") == "01"


def test_next_suffix_is_one_past_highest_sent_msg():
    with tempfile.TemporaryDirectory() as d:
        tdir = Path(d) / "TRANSMITTAL"
        tdir.mkdir()
        # Two transmittals were emailed (saved as Outlook .msg) -> next is 03,
        # even though a -03 draft .doc already exists (it hasn't been sent).
        _touch(tdir / "419624 DWG TRANSMITTAL-01.msg")
        _touch(tdir / "419624 DWG TRANSMITTAL-02.msg")
        _touch(tdir / "419624 DWG TRANSMITTAL-03.doc")
        assert doc.sent_transmittal_numbers(tdir, "419624") == [1, 2]
        assert doc.next_transmittal_suffix(tdir, "419624") == "03"


def test_next_suffix_ignores_other_orders_and_non_transmittals():
    with tempfile.TemporaryDirectory() as d:
        tdir = Path(d) / "TRANSMITTAL"
        tdir.mkdir()
        _touch(tdir / "419624 DWG TRANSMITTAL-05.msg")
        _touch(tdir / "999999 DWG TRANSMITTAL-09.msg")  # different order
        _touch(tdir / "419624 some other email.msg")    # not a transmittal
        assert doc.sent_transmittal_numbers(tdir, "419624") == [5]
        assert doc.next_transmittal_suffix(tdir, "419624") == "06"


def test_default_out_path_uses_sent_msg_suffix():
    with tempfile.TemporaryDirectory() as d:
        folder = Path(d) / "419624"
        tdir = folder / "TRANSMITTAL"
        tdir.mkdir(parents=True)
        _touch(tdir / "419624 DWG TRANSMITTAL-01.msg")
        _touch(tdir / "419624 DWG TRANSMITTAL-02.msg")
        data = td.TransmittalData(order="419624", folder=str(folder))
        p = doc.default_out_path("419624", data=data)
        assert p.name == "419624 DWG TRANSMITTAL-03.doc"
        assert p.parent.name == doc.GENERATED_SUBDIR


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
