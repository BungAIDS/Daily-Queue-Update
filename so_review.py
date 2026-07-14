"""A throwaway Sales-Order review workbook + a note queue Claude works through.

Purpose: give a human a place to comment on how each Sales-Order line item was
captured and what should be done with it, without touching the live product.
The co-authored master and its Sales Order tab stay a read-only browse; this
builds a SEPARATE, disposable .xlsx you can scribble in. It exists only as a
path to a better parser — it is not part of the final workbook.

Why separate: the live Sales Order tab is a picker-driven FILTER spill, so a
note typed beside it belongs to the CELL, not the order — switch orders before
the next scan and the note lands on the wrong row, and the next repaint wipes
it (no macros to catch the edit). A standalone sheet with one STATIC row per
line item sidesteps all of that: each row carries its own order # + item #, so
a note can never mis-associate.

The loop:
  1. `python so_review.py build`  -> writes sales_order_review.xlsx from the
     line-items store: every order's component hierarchy (so_hierarchy), one
     row per row, with an editable "Note" column and the running status of any
     note already recorded.
  2. You filter to an order, type notes on the line-item rows, save, and
     `python so_review.py sync` folds them into the note queue
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
def write_workbook(path: Path, line_items_store: Dict[str, Any],
                   review_store: Dict[str, Any]) -> int:
    """Write the standalone review workbook. Returns the row count.

    Values are bulk-appended and the only visual cues are applied as a handful
    of workbook-level rules (header style + one conditional-format fill on the
    Note column). Styling cell-by-cell across the ~16K rows was the whole cost
    of this build (~30s, so the launcher button looked hung); the rules make it
    ~1s."""
    from openpyxl import Workbook
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", fgColor="305496")
    header_font = Font(color="FFFFFF", bold=True)
    note_fill = PatternFill("solid", fgColor="FFF2CC")      # invites typing

    rows = review_rows(line_items_store, review_store)
    wb = Workbook()
    ws = wb.active
    ws.title = "Line Items"

    # One bulk append of every row — the fast path. No per-cell styling.
    ws.append(HEADERS)
    for r in rows:
        ws.append([r["order"], r["item_no"], r["kind"], r["hierarchy"], r["price"],
                   r["note"], r["status"], r["resolution"]])

    for c in range(1, len(HEADERS) + 1):                    # header row only
        cell = ws.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = header_font

    last = ws.max_row
    note_col = get_column_letter(_COL["Note"] + 1)
    if last >= 2:
        # Tint the whole Note column via a SINGLE conditional-format rule (O(1)
        # to write) instead of filling thousands of cells one at a time. Type on
        # the rows that show an Item # — the others (component/attribute rows)
        # aren't line items and are ignored on read.
        ws.conditional_formatting.add(
            f"{note_col}2:{note_col}{last}",
            FormulaRule(formula=["TRUE"], fill=note_fill, stopIfTrue=False))

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{last}"
    widths = {"Order": 10, "Item": 6, "Kind": 11, "Hierarchy": 60, "Price": 12,
              "Note": 45, "Status": 10, "Resolution": 45}
    for h, w in widths.items():
        ws.column_dimensions[get_column_letter(_COL[h] + 1)].width = w

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(str(path))
    except PermissionError:
        raise RuntimeError(
            f"Could not write {path.name} — it looks like it's still open in "
            f"Excel. Close it, then run this again.") from None
    return len(rows)


def read_edits(path: Path) -> List[Dict[str, Any]]:
    """Read the Note column back out of the workbook. Returns one dict per row
    that has an Order, an Item # and a non-empty Note."""
    from openpyxl import load_workbook

    wb = load_workbook(str(path), data_only=True)
    ws = wb["Line Items"] if "Line Items" in wb.sheetnames else wb.active
    header = [str(c.value or "") for c in ws[1]]
    idx = {h: header.index(h) for h in HEADERS if h in header}
    edits: List[Dict[str, Any]] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        def val(h: str) -> str:
            i = idx.get(h)
            return "" if i is None or i >= len(row) or row[i] is None else str(row[i]).strip()
        note = val("Note")
        if val("Order") and val("Item") and note:
            edits.append({"order": val("Order"), "item_no": val("Item"),
                          "item_text": val("Hierarchy"), "note": note})
    return edits


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
          f"filter the Order column, type in the yellow Note cells, save, then "
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
    store = load_store()
    added = ingest_edits(store, read_edits(path))
    save_store(store)
    print(f"Recorded {added} new note(s). {len(open_notes(store))} open in the queue "
          f"({REVIEW_STORE_PATH}).")
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
