"""Tests for the Sales-Order review tool (so_review.py): the note queue, the
hierarchy rows, and the workbook write/read round-trip.

No pytest — run directly:

    python test_so_review.py
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import so_review as sr


def _line_items_store():
    # One order with the 421966-shaped IVC family (three lines -> one component)
    # plus a standalone line, so tree_rows yields COMPONENT/FACT/SOURCE rows.
    return {"jobs": {"421966": {"items": [
        {"raw": "Base Fan L 16,649.00", "norm": "BASE FAN", "price": "16,649.00",
         "details": [], "tags": []},
        {"raw": "Inlet Volume Control, Low Leak, Automatic L 3,531.00",
         "norm": "INLET VOLUME CONTROL LOW LEAK AUTOMATIC", "price": "3,531.00",
         "details": ["Actuator Manufacturer: By Others"], "tags": ["DAMPER"],
         "attributes": {"used_on": "IVC", "operation": "Automatic"}},
        {"raw": "Inlet, Flanged, Punched (with IVC) L 1,559.00",
         "norm": "INLET FLANGED PUNCHED WITH IVC", "price": "1,559.00",
         "details": [], "tags": ["FLANGE"], "attributes": {"used_on": "IVC"}},
    ]}}}


def test_record_note_appends_and_dedups():
    store = {"notes": []}
    a = sr.record_note(store, "421966", 1, "Base Fan", "keep as-is")
    assert a and a["id"] == 1 and a["status"] == sr.STATUS_OPEN
    # Exact repeat (same order+item+text) is ignored.
    assert sr.record_note(store, "421966", 1, "Base Fan", "keep as-is") is None
    assert len(store["notes"]) == 1
    # Different text -> a new note; ids increment.
    b = sr.record_note(store, "421966", 1, "Base Fan", "actually check the price")
    assert b["id"] == 2 and len(store["notes"]) == 2
    # Blank note or blank order is dropped.
    assert sr.record_note(store, "421966", 1, "Base Fan", "   ") is None
    assert sr.record_note(store, "", 1, "x", "note") is None


def test_mark_handled_and_open_filter():
    store = {"notes": []}
    n = sr.record_note(store, "421966", 2, "IVC", "these 3 should be one component")
    assert len(sr.open_notes(store)) == 1
    assert sr.mark_handled(store, n["id"], "confirmed used_on=IVC grouping")
    assert sr.open_notes(store) == []
    done = store["notes"][0]
    assert done["status"] == sr.STATUS_HANDLED and done["handled_at"]
    assert done["resolution"] == "confirmed used_on=IVC grouping"
    assert not sr.mark_handled(store, 999, "no such id")


def test_review_rows_attach_notes_to_line_items_only():
    li = _line_items_store()
    store = {"notes": []}
    # A note on item #2 (the IVC's primary source line).
    sr.record_note(store, "421966", 2, "Inlet Volume Control…", "grouping looks right")
    rows = sr.review_rows(li, store)
    # Derived rows (component headers, merged attributes) carry no item # and are
    # NOT annotatable — a note there would have no stable line to anchor to.
    comp = [r for r in rows if r["kind"] == sr.so_hierarchy.KIND_COMPONENT]
    assert comp and all(c["order"] == "421966" for c in comp)
    attrs = [r for r in rows if r["kind"] == sr.so_hierarchy.KIND_ATTRIBUTE]
    assert attrs and all(not a["annotatable"] and a["note"] == "" for a in attrs)
    # The SOURCE row for item #2 carries the recorded note and is annotatable.
    src2 = [r for r in rows if r["kind"] == sr.so_hierarchy.KIND_SOURCE
            and str(r["item_no"]) == "2"]
    assert src2 and src2[0]["annotatable"] and src2[0]["note"] == "grouping looks right"
    # Every row names its order, so the sheet is self-identifying.
    assert all(r["order"] == "421966" for r in rows)
    assert rows[0]["group_start"] and not rows[1]["group_start"]


def test_ingest_edits_records_only_line_item_notes():
    store = {"notes": []}
    edits = [
        {"order": "421966", "item_no": "2", "item_text": "IVC", "note": "good"},
        {"order": "421966", "item_no": "", "item_text": "[IVC] — 3 lines", "note": "ignored"},  # not a line item
        {"order": "421966", "item_no": "1", "item_text": "Base Fan", "note": ""},   # empty note
    ]
    added = sr.ingest_edits(store, edits)
    assert added == 1
    assert store["notes"][0]["item_no"] == "2" and store["notes"][0]["note"] == "good"
    # Re-ingesting the same edits adds nothing (dedup).
    assert sr.ingest_edits(store, edits) == 0


def test_store_roundtrip(tmp: Path):
    p = tmp / "so_review_notes.json"
    store = {"notes": []}
    sr.record_note(store, "421966", 2, "IVC", "note one")
    sr.save_store(store, p)
    back = sr.load_store(p)
    assert back["notes"][0]["note"] == "note one"
    # A missing/garbage file loads as an empty queue.
    assert sr.load_store(tmp / "nope.json") == {"notes": []}


def test_workbook_write_read_and_sync_roundtrip(tmp: Path):
    li = _line_items_store()
    store = {"notes": []}
    wb = tmp / "review.xlsx"
    n = sr.write_workbook(wb, li, store)
    assert n > 0 and wb.exists()

    # Simulate the human typing a note by editing the Note cell on item #2's
    # SOURCE row, then reading it back.
    from openpyxl import load_workbook
    book = load_workbook(str(wb))
    ws = book["Line Items"]
    note_col = sr.HEADERS.index("Note") + 1
    item_col = sr.HEADERS.index("Item") + 1
    typed = False
    for i in range(2, ws.max_row + 1):
        if str(ws.cell(row=i, column=item_col).value) == "2":
            ws.cell(row=i, column=note_col).value = "parsed IVC grouping is correct"
            typed = True
            break
    assert typed
    book.save(str(wb))

    edits = sr.read_edits(wb)
    assert {"order": "421966", "item_no": "2"} == {k: edits[0][k] for k in ("order", "item_no")}
    added = sr.ingest_edits(store, edits)
    assert added == 1 and sr.open_notes(store)[0]["note"] == "parsed IVC grouping is correct"

    # Rebuild with the note recorded -> it shows on the sheet, marked open.
    sr.write_workbook(wb, li, store)
    book2 = load_workbook(str(wb))
    ws2 = book2["Line Items"]
    status_col = sr.HEADERS.index("Status") + 1
    found = [i for i in range(2, ws2.max_row + 1)
             if str(ws2.cell(row=i, column=item_col).value) == "2"
             and ws2.cell(row=i, column=note_col).value]
    assert found and str(ws2.cell(row=found[0], column=status_col).value) == sr.STATUS_OPEN


def main() -> int:
    passed = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            code = fn.__code__
            if "tmp" in code.co_varnames[:code.co_argcount]:
                with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as d:
                    fn(Path(d))
            else:
                fn()
            print(f"  ok  {name}")
            passed += 1
    print(f"\n{passed} tests passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
