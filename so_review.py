"""A throwaway Sales-Order review workbook + a note queue Claude works through.

Purpose: give a human a place to comment on how each Sales-Order line item was
captured and what should be done with it, without touching the live product.
The co-authored master stays read-only; this builds a SEPARATE, disposable
.xlsx you can scribble in. It exists only as a path to a better parser - it is
not part of the final workbook.

Why separate: a picker-driven Sales Order view has no Excel event code to bind
typed text to an order. This workbook therefore keeps Notes as the canonical
STATIC list. Sales Order has a temporary Add Note column for one selected
order; sync it before changing the picker. The red unsynced warning makes that
state visible, and sync transfers the text to Notes before clearing the input.

The loop:
  1. `python so_review.py build`  -> writes sales_order_review.xlsx from the
     line-items store: every order's component hierarchy (so_hierarchy), one
     row per row, with an editable "Note" column and the running status of any
     note already recorded.
  2. You either type directly on Notes, or pick one order and use Sales Order's
     Add Note cells. Save and close, then `python so_review.py sync` folds them
     into the note queue
     (so_review_notes.json), which is published with the other stores so
     Claude can read it.
  3. `python so_review.py list` shows the OPEN notes. Claude acts on each and
     `python so_review.py handle <id> "what I did"` marks it handled with a
     resolution; the next `build` shows it as handled so the list visibly
     burns down.

Notes anchor to a LINE ITEM (order #, item #) — the SOURCE / single-line
COMPONENT rows that carry an item #; a note on such a line covers its facts and
details. The item's raw text is stored alongside for context and to re-attach a
note if item numbers shift on a re-parse.

Pure logic (store + rows) is import-light and unit-tested; the two Excel
functions lazy-import openpyxl so the rest of the module loads without it.
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import so_hierarchy
from config import BACKLOG_DIR

REVIEW_STORE_PATH = BACKLOG_DIR / "so_review_notes.json"
DEFAULT_WORKBOOK = BACKLOG_DIR / "sales_order_review.xlsx"

# Return channel for Claude's handled-marks. The note queue flows UP to Claude
# with the other published stores (data_push), but that push is one-way, so
# Claude records each note it resolves in this small TRACKED ledger at the repo
# root. It rides down to the user's machine on the normal Git Update, and the
# "Update SO Review" action applies it — marking those notes handled locally so
# they drop off the sheet. Append-only {handled: [{id, resolution, handled_at}]}
# so it can never merge-conflict.
HANDLED_MARKS_PATH = Path(__file__).resolve().parent / "so_review_handled.json"

STATUS_OPEN = "open"
STATUS_HANDLED = "handled"

# Workbook columns (also the read-back contract). Order + Item + Note are the
# ones sync reads; the rest are context the human reads.
HEADERS = ["Order", "Item", "Kind", "Hierarchy", "Price", "Note", "Status", "Resolution"]
_COL = {h: i for i, h in enumerate(HEADERS)}

# Two tabs, like the live GL Queue workbook's Sales Order view: BROWSE_SHEET is
# a picker with a dedicated Add Note input; NOTES_SHEET is the canonical flat
# note grid. ORDERS_SHEET maps each picker order to its rows in NOTES_SHEET.
BROWSE_SHEET = "Sales Order"
NOTES_SHEET = "Notes"
ORDERS_SHEET = "Orders"
BROWSE_HEADERS = ["Item", "Kind", "Hierarchy", "Price", "Existing Note", "Add Note"]
BROWSE_ADD_NOTE = "Add Note"
BROWSE_HEADER_ROW = 3
BROWSE_FIRST_ROW = 4


# --------------------------------------------------------------------------- #
# Note queue store                                                             #
# --------------------------------------------------------------------------- #
def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def load_store(path: Optional[Path] = None) -> Dict[str, Any]:
    p = Path(path or REVIEW_STORE_PATH)
    if not p.exists():
        return {"notes": []}
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {"notes": []}
    if not isinstance(data, dict) or not isinstance(data.get("notes"), list):
        return {"notes": []}
    return data


def save_store(store: Dict[str, Any], path: Optional[Path] = None) -> None:
    p = Path(path or REVIEW_STORE_PATH)
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(p.suffix + ".tmp")
    tmp.write_text(json.dumps(store, indent=2, default=str), encoding="utf-8")
    tmp.replace(p)


def _next_id(store: Dict[str, Any]) -> int:
    return max((int(n.get("id", 0)) for n in store["notes"]), default=0) + 1


def record_note(store: Dict[str, Any], order: str, item_no: Any, item_text: str,
                note: str, when: Optional[str] = None) -> Optional[Dict[str, Any]]:
    """Append a note for (order, item). Idempotent on exact text: if a note with
    the same order + item + text already exists (any status), nothing is added
    and None is returned — so re-syncing the same workbook never duplicates."""
    order, note = str(order).strip(), str(note).strip()
    item = str(item_no).strip()
    if not order or not note:
        return None
    for n in store["notes"]:
        if (str(n.get("order")) == order and str(n.get("item_no")) == item
                and str(n.get("note")) == note):
            return None
    entry = {"id": _next_id(store), "order": order, "item_no": item,
             "item_text": str(item_text or ""), "note": note,
             "status": STATUS_OPEN, "created_at": when or _now(),
             "handled_at": None, "resolution": None}
    store["notes"].append(entry)
    return entry


def mark_handled(store: Dict[str, Any], note_id: int, resolution: str,
                 when: Optional[str] = None) -> bool:
    """Mark one note handled with what was done. Claude calls this after acting."""
    for n in store["notes"]:
        if int(n.get("id", -1)) == int(note_id):
            n["status"] = STATUS_HANDLED
            n["resolution"] = str(resolution or "").strip()
            n["handled_at"] = when or _now()
            return True
    return False


def open_notes(store: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [n for n in store["notes"] if n.get("status") != STATUS_HANDLED]


def open_notes_by_item(store: Dict[str, Any]) -> Dict[tuple, Dict[str, Any]]:
    """(order, item#) -> the most recent OPEN note for it, for pre-filling the
    sheet. Handled notes are intentionally excluded so a resolved note drops off
    the sheet on the next build, leaving every still-open note in place."""
    out: Dict[tuple, Dict[str, Any]] = {}
    for n in store["notes"]:
        if n.get("status") != STATUS_HANDLED:
            out[(str(n.get("order")), str(n.get("item_no")))] = n
    return out


# --------------------------------------------------------------------------- #
# Handled-marks ledger (Claude -> user return channel)                         #
# --------------------------------------------------------------------------- #
def _load_ledger() -> Dict[str, Any]:
    if not HANDLED_MARKS_PATH.exists():
        return {"handled": []}
    try:
        data = json.loads(HANDLED_MARKS_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {"handled": []}
    return data if isinstance(data.get("handled"), list) else {"handled": []}


def record_handled_mark(note_id: int, resolution: str, when: Optional[str] = None) -> None:
    """Append (or update) a handled-mark in the tracked ledger, so it travels to
    the user's machine on the next Git Update."""
    led = _load_ledger()
    led["handled"] = [m for m in led["handled"] if int(m.get("id", -1)) != int(note_id)]
    led["handled"].append({"id": int(note_id), "resolution": str(resolution or "").strip(),
                           "handled_at": when or _now()})
    HANDLED_MARKS_PATH.write_text(json.dumps(led, indent=2), encoding="utf-8")


def apply_handled_marks(store: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Fold the ledger's handled-marks into a local queue: any note whose id is
    marked (and isn't already handled) becomes handled with the recorded
    resolution. Returns the notes newly closed this call (for reporting)."""
    marks = {int(m["id"]): m for m in _load_ledger()["handled"] if "id" in m}
    closed = []
    for n in store["notes"]:
        m = marks.get(int(n.get("id", -1)))
        if m and n.get("status") != STATUS_HANDLED:
            n["status"] = STATUS_HANDLED
            n["resolution"] = m.get("resolution", "")
            n["handled_at"] = m.get("handled_at") or _now()
            closed.append(n)
    return closed


# --------------------------------------------------------------------------- #
# Rows: every order's hierarchy + the note recorded for each line item         #
# --------------------------------------------------------------------------- #
def _job_sort_key(job: str) -> tuple:
    return (0, int(job), job) if str(job).isdigit() else (1, 0, str(job))


def review_rows(line_items_store: Dict[str, Any],
                review_store: Dict[str, Any]) -> List[Dict[str, Any]]:
    """One display row per hierarchy row across every order in the line-items
    store, newest job first, with any recorded note/status/resolution attached
    to its line item. `group_start` marks the first row of each order (banded in
    the sheet). Pure — no Excel."""
    by_item = open_notes_by_item(review_store)   # handled notes drop off the sheet
    jobs = (line_items_store.get("jobs") or {})
    rows: List[Dict[str, Any]] = []
    for jn in sorted(jobs, key=_job_sort_key, reverse=True):
        items = (jobs[jn] or {}).get("items") or []
        first = True
        for tr in so_hierarchy.tree_rows(items):
            item_no = tr.get("item_no")
            rec = by_item.get((str(jn), str(item_no))) if item_no != "" else None
            rows.append({
                "order": str(jn),
                "item_no": item_no if item_no != "" else "",
                "kind": tr["kind"],
                "hierarchy": so_hierarchy.indent_text(tr),
                "price": tr.get("price", ""),
                "note": (rec or {}).get("note", "") or "",
                "status": (rec or {}).get("status", "") or "",
                "resolution": (rec or {}).get("resolution", "") or "",
                "annotatable": item_no != "",   # only line-item rows take a note
                "group_start": first,
            })
            first = False
    return rows


def ingest_edits(review_store: Dict[str, Any],
                 edits: List[Dict[str, Any]], when: Optional[str] = None) -> int:
    """Fold rows read back from the workbook into the queue. Each edit is
    {order, item_no, item_text, note}; only rows with an order, an item # and a
    non-empty note are recorded (record_note dedups exact repeats). Returns how
    many NEW notes were added."""
    added = 0
    for e in edits:
        if str(e.get("item_no", "")).strip() == "":
            continue
        if record_note(review_store, e.get("order", ""), e.get("item_no", ""),
                        e.get("item_text", ""), e.get("note", ""), when=when):
            added += 1
    return added


# --------------------------------------------------------------------------- #
# Excel I/O (lazy openpyxl)                                                     #
# --------------------------------------------------------------------------- #
def _unique_orders(rows: List[Dict[str, Any]]) -> List[str]:
    seen, out = set(), []
    for r in rows:
        if r["order"] not in seen:
            seen.add(r["order"])
            out.append(r["order"])
    return out


def write_workbook(path: Path, line_items_store: Dict[str, Any],
                   review_store: Dict[str, Any]) -> int:
    """Write the two-tab review workbook. Returns the data-row count.

    Colour and grouping are applied as a handful of workbook-level conditional-
    format rules (O(1) each), never cell-by-cell — styling every one of the
    ~16K rows was what made the build take ~30s (and the launcher button look
    hung); this stays ~1-2s.

    Tabs (like the live GL Queue Sales Order view):
      - 'Sales Order' — a picker: choose an order, review its hierarchy, and add
        notes in the dedicated input column.
      - 'Notes' — the canonical flat grid; sync reads both tabs back."""
    from openpyxl import Workbook
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Font, PatternFill, Protection
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    header_fill = PatternFill("solid", fgColor="305496")
    header_font = Font(color="FFFFFF", bold=True)
    note_fill = PatternFill("solid", fgColor="FFF2CC")      # invites typing
    band_fill = PatternFill("solid", fgColor="F2F2F2")      # alternating orders
    comp_fill = PatternFill("solid", fgColor="DDEBF7")      # component header rows
    comp_font = Font(bold=True)
    review_font = Font(color="C00000", bold=True)           # parser review rows
    picker_fill = PatternFill("solid", fgColor="FFF2CC")

    rows = review_rows(line_items_store, review_store)
    orders = _unique_orders(rows)
    order_windows: Dict[str, List[int]] = {}
    for sheet_row, row in enumerate(rows, start=2):
        order = row["order"]
        if order not in order_windows:
            order_windows[order] = [sheet_row, 0]
        order_windows[order][1] += 1
    ncols = len(HEADERS)
    wb = Workbook()

    # --- Notes tab: the data + input grid ------------------------------------
    notes = wb.active
    notes.title = NOTES_SHEET
    band_col = ncols + 1                                    # hidden parity helper
    notes.append(HEADERS + ["_band"])
    parity = 0
    for r in rows:
        if r["group_start"]:
            parity ^= 1                                     # flip per order group
        notes.append([r["order"], r["item_no"], r["kind"], r["hierarchy"], r["price"],
                      r["note"], r["status"], r["resolution"], parity])
    last = notes.max_row
    for c in range(1, ncols + 1):
        notes.cell(1, c).fill = header_fill
        notes.cell(1, c).font = header_font
    notes.column_dimensions[get_column_letter(band_col)].hidden = True

    if last >= 2:
        full = f"A2:{get_column_letter(ncols)}{last}"
        bandL, kindL = get_column_letter(band_col), get_column_letter(_COL["Kind"] + 1)
        noteL = get_column_letter(_COL["Note"] + 1)
        # Order banding (fill only) + Kind cues (font only) coexist: each rule
        # sets only the properties it needs, so they layer without fighting.
        notes.conditional_formatting.add(full, FormulaRule(formula=[f"${bandL}2=1"], fill=band_fill))
        notes.conditional_formatting.add(full, FormulaRule(formula=[f'${kindL}2="COMPONENT"'], fill=comp_fill, font=comp_font))
        notes.conditional_formatting.add(full, FormulaRule(formula=[f'${kindL}2="REVIEW"'], font=review_font))
        notes.conditional_formatting.add(f"{noteL}2:{noteL}{last}",
                                         FormulaRule(formula=["TRUE"], fill=note_fill, stopIfTrue=False))
    notes.freeze_panes = "A2"
    notes.auto_filter.ref = f"A1:{get_column_letter(ncols)}{last}"
    for h, w in {"Order": 10, "Item": 6, "Kind": 11, "Hierarchy": 60, "Price": 12,
                 "Note": 45, "Status": 10, "Resolution": 45}.items():
        notes.column_dimensions[get_column_letter(_COL[h] + 1)].width = w

    # --- Orders tab (hidden): dropdown plus source row windows ---------------
    osheet = wb.create_sheet(ORDERS_SHEET)
    for o in orders:
        start_row, row_count = order_windows[o]
        osheet.append([o, start_row, row_count])
    osheet.sheet_state = "hidden"

    # --- Sales Order tab (first): pick an order -> its hierarchy appears ------
    browse = wb.create_sheet(BROWSE_SHEET, 0)
    browse["A1"] = "Order:"
    browse["A1"].font = Font(bold=True)
    browse["B1"].fill = picker_fill
    browse["B1"].font = Font(bold=True)
    browse["B1"].protection = Protection(locked=False)
    browse["C1"] = "Pick an order, then use the yellow Add Note cells"
    browse["C1"].font = Font(italic=True, color="808080")
    if orders:
        dv = DataValidation(type="list", formula1=f"={ORDERS_SHEET}!$A$1:$A${len(orders)}",
                            allow_blank=True)
        dv.showErrorMessage = False                        # free typing allowed too
        browse.add_data_validation(dv)
        dv.add(browse["B1"])
    for i, h in enumerate(BROWSE_HEADERS, start=1):
        browse.cell(BROWSE_HEADER_ROW, i).value = h
        browse.cell(BROWSE_HEADER_ROW, i).fill = header_fill
        browse.cell(BROWSE_HEADER_ROW, i).font = header_font
    add_note_col = BROWSE_HEADERS.index(BROWSE_ADD_NOTE) + 1
    browse.cell(BROWSE_HEADER_ROW, add_note_col).fill = note_fill
    browse.cell(BROWSE_HEADER_ROW, add_note_col).font = Font(color="7F6000", bold=True)
    # Keep the source lookup in hidden helper cells, then use ordinary INDEX
    # formulas for the picked order.  openpyxl cannot emit Excel's required
    # dynamic-array metadata for FILTER; writing FILTER as a normal formula
    # makes Excel repair the workbook and remove Sales Order!A4 on open.
    if orders:
        order_last = len(orders)
        browse["G1"] = (
            f'=IFERROR(INDEX(\'{ORDERS_SHEET}\'!$B$1:$B${order_last},'
            f'MATCH($B$1&"",\'{ORDERS_SHEET}\'!$A$1:$A${order_last},0)),0)'
        )
        browse["H1"] = (
            f'=IFERROR(INDEX(\'{ORDERS_SHEET}\'!$C$1:$C${order_last},'
            f'MATCH($B$1&"",\'{ORDERS_SHEET}\'!$A$1:$A${order_last},0)),0)'
        )
    else:
        browse["G1"], browse["H1"] = 0, 0
    browse.column_dimensions["G"].hidden = True
    browse.column_dimensions["H"].hidden = True

    max_order_rows = max((window[1] for window in order_windows.values()), default=1)
    for output_row in range(BROWSE_FIRST_ROW, BROWSE_FIRST_ROW + max_order_rows):
        position = f"ROWS($A${BROWSE_FIRST_ROW}:A{output_row})"
        for output_col, source_col in enumerate(range(2, 7), start=1):
            source_letter = get_column_letter(source_col)
            lookup = (
                f"INDEX('{NOTES_SHEET}'!${source_letter}:${source_letter},"
                f"$G$1+{position}-1)"
            )
            browse.cell(output_row, output_col).value = (
                f'=IF(OR($B$1="",$H$1<{position}),"",'
                f'IF({lookup}="","",{lookup}))'
            )
    bmax = BROWSE_HEADER_ROW + max_order_rows
    bkindL = get_column_letter(2)                           # Kind is the 2nd browse col (B)
    browse.conditional_formatting.add(f"A4:E{bmax}", FormulaRule(formula=[f'${bkindL}4="COMPONENT"'], fill=comp_fill, font=comp_font))
    browse.conditional_formatting.add(f"A4:E{bmax}", FormulaRule(formula=[f'${bkindL}4="REVIEW"'], font=review_font))
    add_note_range = f"F{BROWSE_FIRST_ROW}:F{bmax}"
    browse.conditional_formatting.add(
        add_note_range,
        FormulaRule(formula=[f'$A{BROWSE_FIRST_ROW}<>""'], fill=note_fill),
    )
    add_note_validation = DataValidation(
        type="custom",
        formula1=f'=$A{BROWSE_FIRST_ROW}<>""',
        allow_blank=True,
    )
    add_note_validation.showErrorMessage = True
    add_note_validation.errorTitle = "Choose a line item"
    add_note_validation.error = "Notes can only be added beside a numbered line item."
    add_note_validation.showInputMessage = True
    add_note_validation.promptTitle = "Add Note"
    add_note_validation.prompt = "Save and close Excel, then run Read SO Notes."
    browse.add_data_validation(add_note_validation)
    add_note_validation.add(add_note_range)
    for output_row in range(BROWSE_FIRST_ROW, bmax + 1):
        browse.cell(output_row, add_note_col).protection = Protection(locked=False)
    browse["C2"] = (
        f'=IF(COUNTIF({add_note_range},"<>")=0,"",'
        f'"UNSYNCED NOTES - save, close, and run Read SO Notes before changing order")'
    )
    browse["C2"].font = Font(color="C00000", bold=True)
    browse.protection.sheet = True
    browse.freeze_panes = f"A{BROWSE_FIRST_ROW}"
    for i, w in enumerate((6, 11, 70, 12, 35, 45), start=1):
        browse.column_dimensions[get_column_letter(i)].width = w
    wb.active = 0                                           # open on the browse tab

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(str(path))
    except PermissionError:
        raise RuntimeError(
            f"Could not write {path.name} — it looks like it's still open in "
            f"Excel. Close it, then run this again.") from None
    return len(rows)


def _notes_sheet(wb):
    """The sheet holding the note grid: NOTES_SHEET, else the legacy 'Line Items'
    name, else the first sheet whose header row carries Order + Note."""
    for name in (NOTES_SHEET, "Line Items"):
        if name in wb.sheetnames:
            return wb[name]
    for ws in wb.worksheets:
        header = {str(c.value or "") for c in ws[1]}
        if {"Order", "Note"} <= header:
            return ws
    return wb.active


def _excel_text(value: Any) -> str:
    """Normalize common Excel scalar values without turning 2 into '2.0'."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def read_edits(path: Path) -> List[Dict[str, Any]]:
    """Read canonical Notes plus temporary Sales Order Add Note inputs."""
    from openpyxl import load_workbook

    wb = load_workbook(str(path), data_only=False)
    try:
        notes = _notes_sheet(wb)
        header = [_excel_text(c.value) for c in notes[1]]
        idx = {h: header.index(h) for h in HEADERS if h in header}

        def notes_value(row: tuple, heading: str) -> str:
            i = idx.get(heading)
            return "" if i is None or i >= len(row) else _excel_text(row[i])

        edits: List[Dict[str, Any]] = []
        for row in notes.iter_rows(min_row=2, values_only=True):
            note = notes_value(row, "Note")
            order = notes_value(row, "Order")
            item_no = notes_value(row, "Item")
            if order and item_no and note:
                edits.append({
                    "order": order,
                    "item_no": item_no,
                    "item_text": notes_value(row, "Hierarchy"),
                    "note": note,
                    "source": "notes",
                })

        if BROWSE_SHEET not in wb.sheetnames:
            return edits
        browse = wb[BROWSE_SHEET]
        browse_header = [_excel_text(c.value) for c in browse[BROWSE_HEADER_ROW]]
        legacy_formula_input = BROWSE_ADD_NOTE not in browse_header and "Note" in browse_header
        if BROWSE_ADD_NOTE in browse_header:
            add_note_col = browse_header.index(BROWSE_ADD_NOTE) + 1
        elif legacy_formula_input:
            # Before Add Note existed, typing in Sales Order replaced the Note
            # formula. A literal in that old column is recoverable; formulas are
            # merely the normal browse display and must not be imported.
            add_note_col = browse_header.index("Note") + 1
        else:
            return edits
        pending = [
            (row_no, _excel_text(browse.cell(row_no, add_note_col).value))
            for row_no in range(BROWSE_FIRST_ROW, browse.max_row + 1)
            if _excel_text(browse.cell(row_no, add_note_col).value)
            and not (
                legacy_formula_input
                and str(browse.cell(row_no, add_note_col).value).startswith("=")
            )
        ]
        if not pending:
            return edits

        selected_order = _excel_text(browse["B1"].value)
        if not selected_order:
            raise RuntimeError(
                "Sales Order has unsynced Add Note entries but no selected order. "
                "Select the order again, save, close Excel, and retry."
            )
        if ORDERS_SHEET not in wb.sheetnames:
            raise RuntimeError("The hidden Sales Order row map is missing; Add Note entries were not imported.")

        window = None
        for order, start_row, row_count in wb[ORDERS_SHEET].iter_rows(
                min_row=1, max_col=3, values_only=True):
            if _excel_text(order) == selected_order:
                window = (int(start_row), int(row_count))
                break
        if window is None:
            raise RuntimeError(
                f"Selected order {selected_order} is not in the workbook row map; "
                "Add Note entries were not imported."
            )

        start_row, row_count = window
        for output_row, note in pending:
            offset = output_row - BROWSE_FIRST_ROW
            if offset < 0 or offset >= row_count:
                raise RuntimeError(
                    f"Sales Order row {output_row} is outside order {selected_order}; "
                    "its Add Note entry was not imported."
                )
            source_row = start_row + offset
            source_order = _excel_text(notes.cell(source_row, idx["Order"] + 1).value)
            item_no = _excel_text(notes.cell(source_row, idx["Item"] + 1).value)
            if source_order != selected_order or not item_no:
                raise RuntimeError(
                    f"Sales Order row {output_row} is not a numbered line item for "
                    f"order {selected_order}; its Add Note entry was not imported."
                )
            edits.append({
                "order": source_order,
                "item_no": item_no,
                "item_text": _excel_text(notes.cell(source_row, idx["Hierarchy"] + 1).value),
                "note": note,
                "source": "sales_order",
            })
        return edits
    finally:
        wb.close()


# --------------------------------------------------------------------------- #
# CLI                                                                          #
# --------------------------------------------------------------------------- #
def _load_line_items() -> Dict[str, Any]:
    import line_items
    return line_items.load_store()


def _cmd_build(args) -> int:
    store = load_store()
    n = write_workbook(Path(args.out), _load_line_items(), store)
    pend = len(open_notes(store))
    print(f"Wrote {args.out} ({n} rows). {pend} open note(s) shown; "
          f"pick an order, type in the yellow Add Note cells, save and close, then "
          f"'Read SO Notes'.")
    return 0


def _cmd_open(args) -> int:
    """Open the review workbook in Excel, building it first if it's missing."""
    import os
    path = Path(args.out)
    if not path.exists():
        write_workbook(path, _load_line_items(), load_store())
        print(f"Built {path}.")
    try:
        os.startfile(str(path))          # Windows: opens in Excel  # type: ignore[attr-defined]
    except AttributeError:               # non-Windows: best-effort
        import subprocess
        subprocess.Popen(["xdg-open", str(path)])
    print(f"Opened {path}.")
    return 0


def _cmd_sync(args) -> int:
    path = Path(args.out)
    if not path.exists():
        print(f"{path} not found — 'Open SO Review' builds it first.", file=sys.stderr)
        return 1
    edits = read_edits(path)
    browse_edits = [e for e in edits if e.get("source") == "sales_order"]
    store = load_store()
    added = ingest_edits(store, edits)
    save_store(store)
    if browse_edits:
        try:
            write_workbook(path, _load_line_items(), store)
        except RuntimeError as exc:
            raise RuntimeError(
                f"Recorded {added} new note(s), but could not clear the Sales Order "
                f"Add Note cells. {exc} Re-run Read SO Notes after closing Excel; "
                f"the notes will not duplicate."
            ) from None
    cleared = f" Cleared {len(browse_edits)} Sales Order input cell(s)." if browse_edits else ""
    print(f"Recorded {added} new note(s).{cleared} {len(open_notes(store))} open in the "
          f"queue ({REVIEW_STORE_PATH}).")
    return 0


def _cmd_refresh(args) -> int:
    """The 'Update SO Review' action: capture anything typed, fold in Claude's
    handled-marks (so resolved notes drop off), and rewrite the sheet — leaving
    every still-open note in place."""
    path = Path(args.out)
    store = load_store()
    added = 0
    if path.exists():
        added = ingest_edits(store, read_edits(path))   # don't lose un-synced typing
    closed = apply_handled_marks(store)                 # Claude's resolutions
    save_store(store)
    n = write_workbook(path, _load_line_items(), store)
    print(f"Updated {path} ({n} rows). Captured {added} new note(s); "
          f"removed {len(closed)} handled note(s); {len(open_notes(store))} still open.")
    for c in closed:
        print(f"  handled #{c['id']} (order {c['order']} item {c['item_no']}): "
              f"{c.get('resolution', '')}")
    return 0


def _cmd_list(args) -> int:
    store = load_store()
    rows = open_notes(store) if not args.all else store["notes"]
    if not rows:
        print("No open notes." if not args.all else "No notes recorded.")
        return 0
    for n in sorted(rows, key=lambda x: int(x.get("id", 0))):
        flag = "✓" if n.get("status") == STATUS_HANDLED else " "
        print(f"[{flag}] #{n['id']}  order {n['order']}  item {n['item_no']}")
        print(f"      line: {n.get('item_text', '')}")
        print(f"      note: {n['note']}")
        if n.get("resolution"):
            print(f"      handled: {n['resolution']}")
    return 0


def _cmd_handle(args) -> int:
    # Always record the mark in the tracked ledger (the return channel to the
    # user's machine); also update a local queue if one is present here.
    resolution = " ".join(args.resolution)
    record_handled_mark(args.id, resolution)
    store = load_store()
    if mark_handled(store, args.id, resolution):
        save_store(store)
    print(f"Marked #{args.id} handled (recorded in {HANDLED_MARKS_PATH.name}; it "
          f"applies on the user's next 'Update SO Review').")
    return 0


def _run(func, args) -> int:
    """Run a subcommand, turning an expected error (e.g. the workbook is open in
    Excel) into a clear one-line message rather than a traceback in the log."""
    try:
        return func(args)
    except RuntimeError as e:
        print(str(e), file=sys.stderr)
        return 1


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = ap.add_subparsers(dest="cmd", required=True)

    b = sub.add_parser("build", help="write the review workbook from the line-items store")
    b.add_argument("--out", default=str(DEFAULT_WORKBOOK))
    b.set_defaults(func=_cmd_build)

    o = sub.add_parser("open", help="open the review workbook (building it if missing)")
    o.add_argument("--out", default=str(DEFAULT_WORKBOOK))
    o.set_defaults(func=_cmd_open)

    s = sub.add_parser("sync", help="read your notes back out of the workbook into the queue")
    s.add_argument("--out", default=str(DEFAULT_WORKBOOK))
    s.set_defaults(func=_cmd_sync)

    r = sub.add_parser("refresh", help="capture notes + apply handled-marks + rewrite the sheet")
    r.add_argument("--out", default=str(DEFAULT_WORKBOOK))
    r.set_defaults(func=_cmd_refresh)

    ls_ = sub.add_parser("list", help="show open notes (--all for handled too)")
    ls_.add_argument("--all", action="store_true")
    ls_.set_defaults(func=_cmd_list)

    h = sub.add_parser("handle", help="mark a note handled with a resolution")
    h.add_argument("id", type=int)
    h.add_argument("resolution", nargs="+")
    h.set_defaults(func=_cmd_handle)

    args = ap.parse_args(argv)
    return _run(args.func, args)


if __name__ == "__main__":
    sys.exit(main())
