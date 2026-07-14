"""Tests for the Sales-Order review tool (so_review.py): the note queue, the
hierarchy rows, and the workbook write/read round-trip.

No pytest — run directly:

    python test_so_review.py
"""
from __future__ import annotations

import sys
import tempfile
from contextlib import redirect_stderr
from io import StringIO
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


def test_multiple_notes_on_one_item_remain_visible_and_readable(tmp: Path):
    li = _line_items_store()
    store = {"notes": []}
    sr.record_note(store, "421966", 2, "IVC", "first note for item 2")
    sr.record_note(store, "421966", 2, "IVC", "second note for item 2")

    rows = [
        row for row in sr.review_rows(li, store)
        if str(row["item_no"]) == "2" and row["note"]
    ]
    assert rows
    assert all("first note for item 2" in row["note"] for row in rows)
    assert all("second note for item 2" in row["note"] for row in rows)

    wb = tmp / "review.xlsx"
    sr.write_workbook(wb, li, store)
    edits = sr.read_edits(wb)
    notes = {
        edit["note"] for edit in edits
        if edit["order"] == "421966" and edit["item_no"] == "2"
    }
    assert notes == {"first note for item 2", "second note for item 2"}


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


def test_handled_notes_drop_off_the_sheet_open_ones_stay():
    li = _line_items_store()
    store = {"notes": []}
    sr.record_note(store, "421966", 1, "Base Fan", "handled one")       # id 1
    sr.record_note(store, "421966", 2, "IVC", "still open")             # id 2
    sr.mark_handled(store, 1, "did the thing")
    rows = sr.review_rows(li, store)
    # The open note shows on its line item; the handled note is gone.
    src1 = [r for r in rows if str(r["item_no"]) == "1" and r["kind"] == sr.so_hierarchy.KIND_SOURCE]
    src2 = [r for r in rows if str(r["item_no"]) == "2" and r["kind"] == sr.so_hierarchy.KIND_SOURCE]
    # 421966 item 1 is "Base Fan" (a lone COMPONENT, item_no on the component row).
    comp1 = [r for r in rows if str(r["item_no"]) == "1"]
    assert comp1 and comp1[0]["note"] == ""            # handled -> removed from sheet
    assert src2 and src2[0]["note"] == "still open"    # open -> stays


def test_handled_marks_ledger_roundtrip(tmp: Path):
    import so_review
    ledger = tmp / "so_review_handled.json"
    old = so_review.HANDLED_MARKS_PATH
    so_review.HANDLED_MARKS_PATH = ledger
    try:
        # Claude records two resolutions in the tracked ledger.
        so_review.record_handled_mark(2, "confirmed IVC grouping")
        so_review.record_handled_mark(5, "fixed the flange rule")
        # Re-recording the same id updates in place (no duplicate).
        so_review.record_handled_mark(2, "confirmed IVC grouping (v2)")
        led = so_review._load_ledger()["handled"]
        assert {m["id"] for m in led} == {2, 5} and len(led) == 2

        # The user's local queue has notes 2 and 5 open; applying the ledger
        # closes exactly those, with Claude's resolutions.
        store = {"notes": [
            {"id": 2, "order": "421966", "item_no": "2", "note": "x", "status": "open"},
            {"id": 5, "order": "421966", "item_no": "15", "note": "y", "status": "open"},
            {"id": 9, "order": "421900", "item_no": "3", "note": "z", "status": "open"},
        ]}
        closed = so_review.apply_handled_marks(store)
        assert {c["id"] for c in closed} == {2, 5}
        assert sr.open_notes(store) == [store["notes"][2]]       # note 9 still open
        assert store["notes"][0]["resolution"] == "confirmed IVC grouping (v2)"
        # Idempotent: applying again closes nothing new.
        assert so_review.apply_handled_marks(store) == []
    finally:
        so_review.HANDLED_MARKS_PATH = old


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
    ws = book[sr.NOTES_SHEET]
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
    ws2 = book2[sr.NOTES_SHEET]
    status_col = sr.HEADERS.index("Status") + 1
    found = [i for i in range(2, ws2.max_row + 1)
             if str(ws2.cell(row=i, column=item_col).value) == "2"
             and ws2.cell(row=i, column=note_col).value]
    assert found and str(ws2.cell(row=found[0], column=status_col).value) == sr.STATUS_OPEN


def test_workbook_has_browse_and_notes_tabs(tmp: Path):
    from openpyxl import load_workbook
    wb = tmp / "review.xlsx"
    sr.write_workbook(wb, _line_items_store(), {"notes": []})
    book = load_workbook(str(wb))
    # Two visible tabs like the live Sales Order view, plus a hidden dropdown
    # source; the browse (picker) tab opens first.
    assert book.sheetnames[:2] == [sr.BROWSE_SHEET, sr.NOTES_SHEET]
    assert book[sr.ORDERS_SHEET].sheet_state == "hidden"
    assert book.active.title == sr.BROWSE_SHEET
    browse = book[sr.BROWSE_SHEET]
    # Picker dropdown plus line-item-only Add Note validation, and compatible
    # INDEX formulas that read Notes.
    # A normal FILTER formula makes Excel repair and discard Sales Order!A4
    # because openpyxl cannot write the required dynamic-array metadata.
    assert [browse.cell(sr.BROWSE_HEADER_ROW, c).value
            for c in range(1, len(sr.BROWSE_HEADERS) + 1)] == sr.BROWSE_HEADERS
    validations = browse.data_validations.dataValidation
    assert len(validations) == 2
    assert {dv.type for dv in validations} == {"list", "custom"}
    f = browse["A4"].value
    assert f.startswith('=IF(OR($B$1=""')
    assert f"INDEX('{sr.NOTES_SHEET}'!$B:$B" in f
    assert "FILTER(" not in f and not browse.array_formulae
    assert str(browse["G1"].value).startswith("=IFERROR(INDEX('Orders'!")
    assert str(browse["H1"].value).startswith("=IFERROR(INDEX('Orders'!")
    assert '$B$1&""' in browse["G1"].value
    assert '$B$1&""' in browse["H1"].value
    assert browse.column_dimensions["G"].hidden
    assert browse.column_dimensions["H"].hidden
    assert browse["E4"].value.startswith("=IF(")
    assert browse["F4"].value is None
    assert browse.protection.sheet
    assert not browse["B1"].protection.locked
    assert browse["E4"].protection.locked
    assert not browse["F4"].protection.locked
    assert str(browse["C2"].value).startswith("=IF(COUNTIF(")
    assert any(str(cf.sqref).startswith("F4:F") for cf in browse.conditional_formatting)
    # The hidden Orders list holds each order once plus its Notes row window.
    assert book[sr.ORDERS_SHEET].cell(1, 1).value == "421966"
    assert book[sr.ORDERS_SHEET].cell(1, 2).value == 2
    assert book[sr.ORDERS_SHEET].cell(1, 3).value > 1
    # Colour cues are conditional-format rules (fast), not per-cell styling.
    assert book[sr.NOTES_SHEET].conditional_formatting._cf_rules


def test_sales_order_add_note_is_read_back(tmp: Path):
    from openpyxl import load_workbook

    wb = tmp / "review.xlsx"
    sr.write_workbook(wb, _line_items_store(), {"notes": []})
    book = load_workbook(str(wb), data_only=False)
    browse = book[sr.BROWSE_SHEET]
    notes = book[sr.NOTES_SHEET]
    orders = book[sr.ORDERS_SHEET]
    browse["B1"] = "421966"

    start_row, row_count = None, None
    for order, start, count in orders.iter_rows(min_row=1, max_col=3, values_only=True):
        if str(order) == "421966":
            start_row, row_count = int(start), int(count)
            break
    assert start_row is not None and row_count is not None

    item_col = sr.HEADERS.index("Item") + 1
    item_offset = next(
        offset for offset in range(row_count)
        if str(notes.cell(start_row + offset, item_col).value) == "2"
    )
    add_note_col = sr.BROWSE_HEADERS.index(sr.BROWSE_ADD_NOTE) + 1
    browse.cell(sr.BROWSE_FIRST_ROW + item_offset, add_note_col).value = (
        "parsed IVC grouping is correct"
    )
    book.save(str(wb))
    book.close()

    edits = sr.read_edits(wb)
    added = [e for e in edits if e.get("source") == "sales_order"]
    assert len(added) == 1
    assert added[0]["order"] == "421966"
    assert added[0]["item_no"] == "2"
    assert added[0]["note"] == "parsed IVC grouping is correct"
    assert "Inlet Volume Control" in added[0]["item_text"]


def test_sync_moves_sales_order_input_to_notes_and_clears_it(tmp: Path):
    from types import SimpleNamespace

    from openpyxl import load_workbook

    line_items = _line_items_store()
    wb = tmp / "review.xlsx"
    queue = tmp / "so_review_notes.json"
    sr.write_workbook(wb, line_items, {"notes": []})
    book = load_workbook(str(wb), data_only=False)
    browse = book[sr.BROWSE_SHEET]
    notes = book[sr.NOTES_SHEET]
    orders = book[sr.ORDERS_SHEET]
    browse["B1"] = "421966"
    _, start_row, row_count = next(
        row for row in orders.iter_rows(min_row=1, max_col=3, values_only=True)
        if str(row[0]) == "421966"
    )
    item_col = sr.HEADERS.index("Item") + 1
    item_offset = next(
        offset for offset in range(int(row_count))
        if str(notes.cell(int(start_row) + offset, item_col).value) == "2"
    )
    add_note_col = sr.BROWSE_HEADERS.index(sr.BROWSE_ADD_NOTE) + 1
    browse.cell(sr.BROWSE_FIRST_ROW + item_offset, add_note_col).value = "sync this note"
    book.save(str(wb))
    book.close()

    old_store_path = sr.REVIEW_STORE_PATH
    old_loader = sr._load_line_items
    sr.REVIEW_STORE_PATH = queue
    sr._load_line_items = lambda: line_items
    try:
        assert sr._cmd_sync(SimpleNamespace(out=str(wb))) == 0
    finally:
        sr.REVIEW_STORE_PATH = old_store_path
        sr._load_line_items = old_loader

    stored = sr.load_store(queue)
    assert len(stored["notes"]) == 1
    assert stored["notes"][0]["note"] == "sync this note"
    rebuilt = load_workbook(str(wb), data_only=False)
    rebuilt_browse = rebuilt[sr.BROWSE_SHEET]
    assert all(
        rebuilt_browse.cell(row, add_note_col).value is None
        for row in range(sr.BROWSE_FIRST_ROW, rebuilt_browse.max_row + 1)
    )
    rebuilt_notes = rebuilt[sr.NOTES_SHEET]
    note_col = sr.HEADERS.index("Note") + 1
    assert any(
        str(rebuilt_notes.cell(row, item_col).value) == "2"
        and rebuilt_notes.cell(row, note_col).value == "sync this note"
        for row in range(2, rebuilt_notes.max_row + 1)
    )
    rebuilt.close()


def test_legacy_sales_order_formula_overwrite_is_recovered(tmp: Path):
    from openpyxl import load_workbook

    wb = tmp / "review.xlsx"
    sr.write_workbook(wb, _line_items_store(), {"notes": []})
    book = load_workbook(str(wb), data_only=False)
    browse = book[sr.BROWSE_SHEET]
    notes = book[sr.NOTES_SHEET]
    orders = book[sr.ORDERS_SHEET]
    browse["B1"] = "421966"
    browse.cell(sr.BROWSE_HEADER_ROW, 5).value = "Note"
    browse.delete_cols(6, 1)

    order, start_row, row_count = next(
        row for row in orders.iter_rows(min_row=1, max_col=3, values_only=True)
        if str(row[0]) == "421966"
    )
    assert str(order) == "421966"
    item_col = sr.HEADERS.index("Item") + 1
    item_offset = next(
        offset for offset in range(int(row_count))
        if str(notes.cell(int(start_row) + offset, item_col).value) == "2"
    )
    browse.cell(sr.BROWSE_FIRST_ROW + item_offset, 5).value = "recover this old note"
    book.save(str(wb))
    book.close()

    edits = sr.read_edits(wb)
    recovered = [e for e in edits if e.get("source") == "sales_order"]
    assert len(recovered) == 1
    assert recovered[0]["order"] == "421966"
    assert recovered[0]["item_no"] == "2"
    assert recovered[0]["note"] == "recover this old note"


def test_sync_upgrades_legacy_browse_without_pending_notes(tmp: Path):
    from types import SimpleNamespace

    from openpyxl import load_workbook

    line_items = _line_items_store()
    wb = tmp / "review.xlsx"
    queue = tmp / "so_review_notes.json"
    sr.write_workbook(wb, line_items, {"notes": []})
    book = load_workbook(str(wb), data_only=False)
    browse = book[sr.BROWSE_SHEET]
    browse.cell(sr.BROWSE_HEADER_ROW, 5).value = "Note"
    browse.delete_cols(6, 1)
    book.save(str(wb))
    book.close()
    assert sr._browse_layout_needs_upgrade(wb)

    old_store_path = sr.REVIEW_STORE_PATH
    old_loader = sr._load_line_items
    sr.REVIEW_STORE_PATH = queue
    sr._load_line_items = lambda: line_items
    try:
        assert sr._cmd_sync(SimpleNamespace(out=str(wb))) == 0
    finally:
        sr.REVIEW_STORE_PATH = old_store_path
        sr._load_line_items = old_loader

    rebuilt = load_workbook(str(wb), data_only=False)
    headers = [
        rebuilt[sr.BROWSE_SHEET].cell(sr.BROWSE_HEADER_ROW, col).value
        for col in range(1, len(sr.BROWSE_HEADERS) + 1)
    ]
    rebuilt.close()
    assert headers == sr.BROWSE_HEADERS


def test_refresh_reports_notes_safe_when_workbook_rewrite_fails(tmp: Path):
    from types import SimpleNamespace

    from openpyxl import load_workbook

    line_items = _line_items_store()
    wb = tmp / "review.xlsx"
    queue = tmp / "so_review_notes.json"
    sr.write_workbook(wb, line_items, {"notes": []})
    book = load_workbook(str(wb), data_only=False)
    notes = book[sr.NOTES_SHEET]
    item_col = sr.HEADERS.index("Item") + 1
    note_col = sr.HEADERS.index("Note") + 1
    row = next(
        row_no for row_no in range(2, notes.max_row + 1)
        if str(notes.cell(row_no, item_col).value) == "2"
    )
    notes.cell(row, note_col).value = "keep this even if Excel is open"
    book.save(str(wb))
    book.close()

    old_store_path = sr.REVIEW_STORE_PATH
    old_loader = sr._load_line_items
    old_writer = sr.write_workbook
    sr.REVIEW_STORE_PATH = queue
    sr._load_line_items = lambda: line_items

    def blocked_writer(*_args, **_kwargs):
        raise RuntimeError("Could not write review.xlsx because Excel has it open.")

    sr.write_workbook = blocked_writer
    error = StringIO()
    try:
        with redirect_stderr(error):
            assert sr._run(sr._cmd_refresh, SimpleNamespace(out=str(wb))) == 1
    finally:
        sr.REVIEW_STORE_PATH = old_store_path
        sr._load_line_items = old_loader
        sr.write_workbook = old_writer

    message = error.getvalue()
    assert "All 1 workbook note(s) are safely recorded" in message
    assert "workbook refresh could not finish" in message
    stored = sr.load_store(queue)
    assert len(stored["notes"]) == 1
    assert stored["notes"][0]["note"] == "keep this even if Excel is open"


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
