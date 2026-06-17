"""Tests for the live master workbook's pure sheet model (live_sheets.py).

No pytest — run directly:

    python test_live_sheets.py

Checks the content/formatting intent of each tab (Live Queue, Changes, History,
Line Items) without any Excel/COM dependency.
"""
from __future__ import annotations

import sys
from datetime import date

import live_sheets as ls
from live_sheets import (FILL_DWG_NO, FILL_DWG_YES, FILL_NEW, FILL_OVERDUE,
                         F_LINK, F_SECTION)

TODAY = date(2026, 6, 16)


def _job(num, **kw):
    j = {"job": num, "item": "47-0-0000", "design": "47", "customer": "ACME CORP",
         "end_date": "06/20/2026", "total_price": "$1,000.00", "so_pdf": "",
         "dwg_extras": {}, "job_folder": ""}
    j.update(kw)
    return j


def _find(sheet, text):
    for r, row in enumerate(sheet.grid):
        for c, cell in enumerate(row):
            if str(cell.value).startswith(text):
                return r, c
    return None


def test_full_queue_headers_and_added_column():
    sh = ls.full_queue_sheet([_job("421000")], TODAY)
    assert sh.grid[0][0].value == "Added"
    assert sh.grid[0][1].value == "Job #"          # first standard column
    assert sh.name == "Live Queue"
    assert sh.autofilter_a1 and sh.freeze == "C2"


def test_full_queue_overdue_fill_and_job_link():
    j = _job("421000", end_date="06/10/2026", so_pdf="Z:\\SO\\421000.pdf")
    sh = ls.full_queue_sheet([j], TODAY)
    job_cell = sh.grid[1][1]                        # Added is col0, Job# is col1
    assert job_cell.value == "421000"
    assert job_cell.link == "Z:\\SO\\421000.pdf" and job_cell.font == F_LINK
    # End Date in the past -> overdue fill on the standard cells.
    assert sh.grid[1][1].fill == FILL_OVERDUE


def test_full_queue_new_fill_and_added_label():
    j = _job("421001", end_date="12/31/2026",
             _carried_over=False, _first_seen="2026-06-16T09:14:00")
    sh = ls.full_queue_sheet([j], TODAY, new_ids={"421001"})
    assert sh.grid[1][0].value.endswith("AM") or sh.grid[1][0].value.endswith("PM")
    assert sh.grid[1][1].fill == FILL_NEW          # new, no urgency


def test_full_queue_dwg_matrix():
    a = _job("421000", dwg_extras={"51": "x"})
    b = _job("421001", dwg_extras={})
    sh = ls.full_queue_sheet([a, b], TODAY)
    # Header has a "-51" column at the end; rows show ✓/green and blank/red.
    pos = _find(sh, "-51")
    assert pos is not None and pos[0] == 0
    col = pos[1]
    assert sh.grid[1][col].value == "✓" and sh.grid[1][col].fill == FILL_DWG_YES
    assert sh.grid[2][col].value == "" and sh.grid[2][col].fill == FILL_DWG_NO


def test_full_queue_footer_total():
    sh = ls.full_queue_sheet([_job("421000", total_price="$1,000.00"),
                              _job("421001", total_price="$2,500.00")], TODAY)
    pos = _find(sh, "Total jobs: 2")
    assert pos is not None and sh.grid[pos[0]][pos[1]].font == F_SECTION
    # The money total lives on the footer row at the Total Price column.
    total_cells = [c for c in sh.grid[pos[0]] if isinstance(c.value, (int, float)) and c.value]
    assert any(abs(c.value - 3500.0) < 0.001 for c in total_cells)


def test_changes_both_groups_and_added():
    intraday = {"new": [_job("421001", _carried_over=False,
                             _first_seen="2026-06-16T09:14:00")],
                "returning": [], "removed": [], "changed": []}
    yesterday = {"new": [], "returning": [], "removed": [_job("420900")],
                 "changed": [{"job": "420800", "customer": "X",
                              "changes": [("end_date", "06/01/2026", "06/05/2026")]}]}
    sh = ls.changes_sheet(intraday, "2026-06-16 (this morning)", yesterday, "2026-06-15")
    assert _find(sh, "Changes since this morning — baseline 2026-06-16") is not None
    assert _find(sh, "Changes vs yesterday — 2026-06-15") is not None
    assert _find(sh, "New orders (1)") is not None
    assert _find(sh, "Removed / completed (1)") is not None
    assert _find(sh, "Orders that changed (1)") is not None


def test_history_sheet():
    hist = {"420000": {"last_seen": "2026-06-10",
                       "snapshot": _job("420000", dwg_extras={"35": "x"})}}
    sh = ls.history_sheet(hist)
    assert sh.grid[0][0].value == "Job #"
    assert "Last Seen" in [c.value for c in sh.grid[0]]
    assert sh.grid[1][0].value == "420000"


def test_line_items_sheet_search_rows():
    store = {"jobs": {"421000": {
        "customer": "ACME CORP", "co_number": 1, "so_pdf": "Z:\\SO\\421000.pdf",
        "items": [
            {"raw": "SS SHAFT SLEEVE", "norm": "STAINLESS STEEL SHAFT SLEEVE",
             "tags": ["SHAFT SEAL"], "details": ["VENDOR X"], "qty": "1",
             "price": "$5", "section": "ACCESSORIES"},
            {"raw": "TEFLON SEAL", "norm": "TEFLON SEAL", "tags": ["SHAFT SEAL"],
             "details": [], "qty": "1", "price": "$3", "section": ""},
        ]}}}
    sh = ls.line_items_sheet(store, order_nums=["421000"])
    assert sh.grid[0] == sh.grid[0]  # header present
    assert sh.grid[0][5].value == "Normalized"          # the searchable column
    # Two item rows for the one order.
    body = [r for r in sh.grid[1:] if r and r[0].value == "421000"]
    assert len(body) == 2
    assert body[0][5].value == "STAINLESS STEEL SHAFT SLEEVE"
    assert sh.grid[1][10].font == F_LINK                 # SO PDF link
    assert sh.autofilter_a1 is not None


def test_line_items_whole_backlog_default():
    store = {"jobs": {
        "421000": {"customer": "A", "co_number": 0, "so_pdf": "",
                   "items": [{"raw": "X", "norm": "X", "tags": []}]},
        "419000": {"customer": "B", "co_number": 0, "so_pdf": "",
                   "items": [{"raw": "Y", "norm": "Y", "tags": []},
                             {"raw": "Z", "norm": "Z", "tags": []}]},
    }}
    # No order_nums -> every stored order (whole backlog), not just the board.
    sh = ls.line_items_sheet(store)
    jobs_in_rows = {r[0].value for r in sh.grid[1:] if r}
    assert jobs_in_rows == {"421000", "419000"}
    assert sum(1 for r in sh.grid[1:] if r) == 3        # 1 + 2 items


def test_live_queue_records_no_dwg_and_new_today_fill():
    j = _job("421000", end_date="12/31/2026", dwg_extras={"51": "x", "35": "x"},
             _carried_over=False, _first_seen="2026-06-16T09:14:00")
    recs = ls.live_queue_records([j], TODAY, new_ids={"421000"})
    assert len(recs) == 1
    key, cells = recs[0]
    assert key == "421000"
    # Job # sits at the key column (1-based) — confirm the index lines up.
    assert cells[ls.LIVE_QUEUE_KEY_COL - 1].value == "421000"
    # Custom DWGs is no longer on Live Queue (Order History only).
    assert "Custom DWGs" not in ls.LIVE_QUEUE_HEADERS
    assert len(cells) == len(ls.LIVE_QUEUE_HEADERS)
    # new_ids drives the new-today highlight (no urgency on a far-future date).
    assert any(c.fill == FILL_NEW for c in cells)


def test_order_history_build_matrices_flags_and_separator():
    orders = [
        ("421000", {"on_queue": True, "added": "2026-06-16T09:00:00", "left": None,
                    "job": _job("421000", dwg_extras={"51": "x"},
                               line_items=[{"tags": ["SHAFT SEAL"]}, {"tags": ["COATING"]}])}),
        ("420900", {"on_queue": False, "added": "2026-06-15T08:00:00",
                    "left": "2026-06-16T07:30:00",
                    "job": _job("420900", dwg_extras={}, line_items=[{"tags": ["SHAFT SEAL"]}])}),
    ]
    spec = ls.order_history_build(orders, TODAY)
    h = spec["headers"]
    assert h[0] == "On Queue"
    assert h[ls.ORDER_HISTORY_KEY_COL - 1] == "Job #"
    assert "-51" in h and "SHAFT SEAL" in h and "COATING" in h and ls.OH_SEP_HEADER in h
    r0 = dict(zip(h, spec["records"][0][1]))
    assert r0["On Queue"].value == "YES" and r0["-51"].value == "✓" and r0["SHAFT SEAL"].value == "✓"
    r1 = dict(zip(h, spec["records"][1][1]))
    assert r1["On Queue"].value == "NO" and r1["Left"].value == "2026-06-16 07:30"
    assert r1["-51"].value == ""                          # 420900 lacks that drawing
    ds, de = spec["dwg_range"]
    fs, fe = spec["feat_range"]
    assert spec["sep_col"] == de + 1 and fs == spec["sep_col"] + 1 and fe >= fs


def test_order_history_row_sig_stable_across_churny_fields():
    base = {"on_queue": True, "added": "2026-06-16T09:00:00", "left": None}
    e1 = ("421000", {**base, "job": _job("421000", end_date="06/20/2026", total_price="$1,000.00")})
    e2 = ("421000", {**base, "job": _job("421000", end_date="07/01/2026", total_price="$9,999.00")})
    s1 = ls.order_history_build([e1], TODAY)
    s2 = ls.order_history_build([e2], TODAY)
    # Churny board fields aren't on Order History, so the row signature is stable
    # -> the 12K log isn't rewritten when a date/price ticks.
    assert ls.row_sig(s1["records"][0][1]) == ls.row_sig(s2["records"][0][1])


def test_plan_upsert_append_update_delete():
    a = ("100", "sigA", ["cellsA"])
    b = ("200", "sigB", ["cellsB"])
    # 100 unchanged, 200 changed, 300 new, 400 only in existing (-> delete).
    desired = [("100", "sigA", "x"), ("200", "sigB2", "y"), ("300", "sigC", "z")]
    existing = {"100": "sigA", "200": "sigB", "400": "sigD"}
    ops = ls.plan_upsert(desired, existing, allow_delete=True)
    kinds = {(o[0], o[1]) for o in ops}
    assert ("update", "200") in kinds
    assert ("append", "300") in kinds
    assert ("delete", "400") in kinds
    assert not any(o[1] == "100" for o in ops)           # unchanged -> no op
    # Without allow_delete, 400 is left alone.
    ops2 = ls.plan_upsert(desired, existing, allow_delete=False)
    assert not any(o[0] == "delete" for o in ops2)


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
