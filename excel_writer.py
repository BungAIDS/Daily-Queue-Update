"""Build the daily Excel report with two tabs: Changes (first) and Full Queue."""
from __future__ import annotations

import logging
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from config import OUTPUT_DIR

log = logging.getLogger(__name__)

RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
# One step darker than each base fill, used when a row is also new today.
RED_FILL_NEW    = PatternFill(start_color="F4A5A8", end_color="F4A5A8", fill_type="solid")
YELLOW_FILL_NEW = PatternFill(start_color="F5D750", end_color="F5D750", fill_type="solid")
NEW_FILL        = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
HEADER_FILL = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
SECTION_FONT = Font(bold=True, size=12)
RED_FONT = Font(color="C00000", bold=True)  # rows whose change order landed this run
LINK_FONT = Font(color="0563C1", underline="single")  # hyperlinks (Job # -> SO pdf, Folder)

# Single source of truth for column order. Each entry is (header, key); the key
# names the job field to print, or a special renderer handled in _write_job_row:
#   "job"         -> job number, hyperlinked to its Sales Order pdf (Z: drive)
#   "folder"      -> AutoCAD job folder (or SO archive folder) hyperlink
#   "co"          -> CO# label
#   "total_price" -> money-formatted cell
#   "flags"       -> flag summary string
# To reorder the report, reorder this list — everything else follows.
COLUMNS = [
    ("Job #", "job"),
    ("Folder", "folder"),
    ("CO#", "co"),
    ("Oper", "oper"),
    ("Design", "design"),
    ("Description", "so_design_desc"),
    ("Size", "so_size"),
    ("Arrangement", "so_arrangement"),
    ("Motor Pos", "so_motor_pos"),
    ("Class", "so_class"),
    ("Rotation", "so_rotation"),
    ("Discharge", "so_discharge"),
    ("% Width", "so_pct_width"),
    ("Wheel Type", "so_wheel_type"),
    ("Design Temp", "so_design_temp"),
    ("Max Temp", "so_max_temp"),
    ("Special Temp", "so_special_temp"),
    ("Customer", "customer"),
    ("Primary Rep", "primary_rep"),
    ("Assigned To", "assigned_to"),
    ("Checker", "checker"),
    ("Start Date", "start_date"),
    ("End Date", "end_date"),
    ("FanNet Date", "fannet_date"),
    ("Item", "item"),
    ("Plan Hrs", "plan_hrs"),
    ("Total Price", "total_price"),
    ("Ship With", "ship_with"),
    ("Note", "status_note"),
    ("Flags", "flags"),
    ("Status", "status"),
]
QUEUE_HEADERS = [h for h, _ in COLUMNS]
_COL_IDX = {key: i for i, (_, key) in enumerate(COLUMNS, start=1)}
TOTAL_PRICE_COL = _COL_IDX["total_price"]  # 1-based; used by the footer total


def _flags_str(j: Dict[str, Any]) -> str:
    flags = []
    if j.get("unapproved"):
        flags.append("UNAPPROVED")
    if j.get("credit_hold"):
        flags.append("CREDIT HOLD")
    if j.get("has_notes"):
        flags.append("NOTES")
    return ", ".join(flags)


def _parse_date(s: str) -> date | None:
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime((s or "").strip(), fmt).date()
        except (ValueError, TypeError):
            continue
    return None


def _parse_money(s: str) -> float:
    try:
        return float((s or "").replace("$", "").replace(",", "").strip() or 0)
    except ValueError:
        return 0.0


MONEY_FMT = '"$"#,##0.00'


def _write_money_cell(ws, row: int, col: int, raw: str):
    """Write a price as a real number (so Excel can sum/sort it), formatted as
    currency. Blank stays blank rather than showing $0.00."""
    raw = (raw or "").strip()
    if not raw:
        return ws.cell(row=row, column=col, value="")
    cell = ws.cell(row=row, column=col, value=_parse_money(raw))
    cell.number_format = MONEY_FMT
    return cell


def _autosize(ws, num_cols: int, skip_cells: set | None = None) -> None:
    """Size each column to its widest cell. Cells listed in skip_cells (a set of
    (row, col) tuples) are ignored — used for values we deliberately let overflow
    into an empty neighbor instead of widening their whole column."""
    skip_cells = skip_cells or set()
    for col in range(1, num_cols + 1):
        letter = get_column_letter(col)
        max_len = 0
        for cell in ws[letter]:
            if (cell.row, cell.column) in skip_cells:
                continue
            val = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(val))
        ws.column_dimensions[letter].width = min(max(max_len + 2, 10), 60)


def _write_changes_tab(ws, briefing: Dict[str, Any], diff: Dict[str, Any],
                       co_changed_ids: set | None = None) -> None:
    co_changed_ids = co_changed_ids or set()
    overflow: set = set()  # (row, col) cells excluded from autosize so they overrun
    row = 1

    # AI Briefing block
    ws.cell(row=row, column=1, value="AI Briefing").font = SECTION_FONT
    row += 1
    ws.cell(row=row, column=1, value=briefing.get("briefing", "(no briefing)"))
    ws.cell(row=row, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.row_dimensions[row].height = 80
    row += 2

    # Anomalies
    anomalies = briefing.get("anomalies", []) or []
    if anomalies:
        ws.cell(row=row, column=1, value="Anomalies").font = SECTION_FONT
        row += 1
        for a in anomalies:
            ws.cell(row=row, column=1, value=f"- {a}")
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
            row += 1
        row += 1

    # Action items
    items = briefing.get("action_items", []) or []
    if items:
        ws.cell(row=row, column=1, value="Top Action Items").font = SECTION_FONT
        row += 1
        ws.cell(row=row, column=1, value="Rank").font = HEADER_FONT
        ws.cell(row=row, column=2, value="Job #").font = HEADER_FONT
        ws.cell(row=row, column=3, value="Reason").font = HEADER_FONT
        for c in range(1, 4):
            ws.cell(row=row, column=c).fill = HEADER_FILL
        row += 1
        for item in items:
            ws.cell(row=row, column=1, value=item.get("rank", ""))
            ws.cell(row=row, column=2, value=str(item.get("job", "")))
            ws.cell(row=row, column=3, value=item.get("reason", ""))
            row += 1
        row += 1

    # New orders
    ws.cell(row=row, column=1, value=f"New orders ({len(diff['new'])})").font = SECTION_FONT
    row += 1
    if diff["new"]:
        for c, h in enumerate(QUEUE_HEADERS, start=1):
            cell = ws.cell(row=row, column=c, value=h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
        row += 1
        for j in diff["new"]:
            _write_job_row(ws, row, j, co_changed=j.get("job") in co_changed_ids)
            row += 1
    else:
        ws.cell(row=row, column=1, value="(none)")
        row += 1
    row += 1

    # Change orders that landed this run (CO# rose since yesterday, including
    # jobs that came back at a higher CO#). Placed right under New orders.
    co_changed = list(diff.get("co_changed", []))
    for j in diff.get("returning", []):
        cr = j.get("_co_returned")
        if cr:
            co_changed.append({"job": j.get("job"), "customer": j.get("customer", ""),
                               "old_co": cr["old_co"], "new_co": cr["new_co"],
                               "co_history": j.get("co_history", []), "_returned": True})
    ws.cell(row=row, column=1, value=f"Change orders this run ({len(co_changed)})").font = SECTION_FONT
    row += 1
    if co_changed:
        for c, h in enumerate(["Job #", "Customer", "Change", "Latest change-order note"], start=1):
            cell = ws.cell(row=row, column=c, value=h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
        row += 1
        for c in co_changed:
            arrow = f"CO#{c['old_co']} -> CO#{c['new_co']}"
            if c.get("_returned"):
                arrow += " (returned)"
            note = c["co_history"][0] if c.get("co_history") else ""
            ws.cell(row=row, column=1, value=c["job"]).font = RED_FONT
            ws.cell(row=row, column=2, value=c["customer"]).font = RED_FONT
            ws.cell(row=row, column=3, value=arrow).font = RED_FONT
            ws.cell(row=row, column=4, value=note).font = RED_FONT
            row += 1
    else:
        ws.cell(row=row, column=1, value="(none)")
        row += 1
    row += 1

    # Returning orders (came back from history)
    returning = diff.get("returning", [])
    ws.cell(row=row, column=1, value=f"Returning orders — back from history ({len(returning)})").font = SECTION_FONT
    row += 1
    if returning:
        headers = QUEUE_HEADERS + ["Last Seen"]
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=row, column=c, value=h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
        row += 1
        for j in returning:
            _write_job_row(ws, row, j, co_changed=j.get("job") in co_changed_ids)
            ws.cell(row=row, column=len(QUEUE_HEADERS) + 1, value=j.get("_last_seen", ""))
            row += 1
    else:
        ws.cell(row=row, column=1, value="(none)")
        row += 1
    row += 1

    # Completed/Removed
    ws.cell(row=row, column=1, value=f"Completed / Removed ({len(diff['removed'])})").font = SECTION_FONT
    row += 1
    if diff["removed"]:
        for c, h in enumerate(QUEUE_HEADERS, start=1):
            cell = ws.cell(row=row, column=c, value=h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
        row += 1
        for j in diff["removed"]:
            _write_job_row(ws, row, j)
            row += 1
    else:
        ws.cell(row=row, column=1, value="(none)")
        row += 1
    row += 1

    # Orders that have changed. Customer spans two columns: it's written in col 2
    # with col 3 left blank so a long name overflows into it, and the customer
    # cells are excluded from autosize so they don't force col 2 (shared with the
    # narrow Folder column in the sections above) wide. Field/Old/New shift right.
    ws.cell(row=row, column=1, value=f"Orders that have changed ({len(diff['changed'])})").font = SECTION_FONT
    row += 1
    if diff["changed"]:
        for c, h in enumerate(["Job #", "Customer", "", "Field", "Old value", "New value"], start=1):
            cell = ws.cell(row=row, column=c, value=h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
        row += 1
        for ch in diff["changed"]:
            for (field, old, new) in ch["changes"]:
                ws.cell(row=row, column=1, value=ch["job"])
                ws.cell(row=row, column=2, value=ch["customer"])
                overflow.add((row, 2))  # let the customer name overrun the blank col 3
                ws.cell(row=row, column=4, value=field)
                ws.cell(row=row, column=5, value=old)
                ws.cell(row=row, column=6, value=new)
                row += 1
    else:
        ws.cell(row=row, column=1, value="(none)")
        row += 1
    row += 1

    # Persistent (3+ days)
    ws.cell(row=row, column=1, value=f"Persistent orders — 3+ days in queue ({len(diff['persistent'])})").font = SECTION_FONT
    row += 1
    if diff["persistent"]:
        for c, h in enumerate(["Job #", "Customer", "Days in queue", "End Date", "Assigned To", "Total Price"], start=1):
            cell = ws.cell(row=row, column=c, value=h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
        row += 1
        for p in diff["persistent"]:
            snap = p["snapshot"]
            ws.cell(row=row, column=1, value=p["job"])
            ws.cell(row=row, column=2, value=p["customer"])
            ws.cell(row=row, column=3, value=p["days"])
            ws.cell(row=row, column=4, value=snap.get("end_date", ""))
            ws.cell(row=row, column=5, value=snap.get("assigned_to", ""))
            _write_money_cell(ws, row, 6, snap.get("total_price", ""))
            row += 1
    else:
        ws.cell(row=row, column=1, value="(none)")
        row += 1

    _autosize(ws, num_cols=len(QUEUE_HEADERS), skip_cells=overflow)


def _co_label(j: Dict[str, Any]) -> str:
    co = j.get("co_number") or 0
    return f"CO#{co}" if co else ""


def _write_job_row(ws, row: int, j: Dict[str, Any], co_changed: bool = False) -> None:
    linked_cols = set()  # hyperlink cells keep their link style, not red
    for col, (_header, key) in enumerate(COLUMNS, start=1):
        if key == "job":
            cell = ws.cell(row=row, column=col, value=j.get("job", ""))
            # Job # links to its Sales Order pdf on the Z: drive (when we have one).
            so_pdf = (j.get("so_pdf") or "").strip()
            if so_pdf and j.get("job"):
                cell.hyperlink = so_pdf
                cell.font = LINK_FONT
                linked_cols.add(col)
        elif key == "folder":
            # AutoCAD job folder, or the SO archive folder as fallback.
            folder = (j.get("job_folder") or "").strip()
            cell = ws.cell(row=row, column=col, value=(j.get("job_type") or "Open") if folder else "")
            if folder:
                cell.hyperlink = folder
                cell.font = LINK_FONT
                linked_cols.add(col)
        elif key == "co":
            ws.cell(row=row, column=col, value=_co_label(j))
        elif key == "total_price":
            _write_money_cell(ws, row, col, j.get("total_price", ""))
        elif key == "flags":
            ws.cell(row=row, column=col, value=_flags_str(j))
        else:
            ws.cell(row=row, column=col, value=j.get(key, ""))

    # A change order that landed this run -> the whole row's text goes red,
    # except the hyperlink cells, which keep their link style.
    if co_changed:
        for col in range(1, len(COLUMNS) + 1):
            if col not in linked_cols:
                ws.cell(row=row, column=col).font = RED_FONT


def _write_full_queue_tab(
    ws,
    jobs: List[Dict[str, Any]],
    today: date,
    new_job_ids: set | None = None,
    co_changed_ids: set | None = None,
) -> None:
    new_job_ids = new_job_ids or set()
    co_changed_ids = co_changed_ids or set()
    for c, h in enumerate(QUEUE_HEADERS, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

    total_price_sum = 0.0
    soon_threshold = today + timedelta(days=3)

    for i, j in enumerate(jobs, start=2):
        _write_job_row(ws, i, j, co_changed=j.get("job") in co_changed_ids)
        # Pick a row fill based on End Date urgency; if the order is also new
        # today, step the chosen color one shade darker (or to light gray if
        # there's no urgency fill yet).
        is_new = j.get("job") in new_job_ids
        end = _parse_date(j.get("end_date", ""))
        if end is not None and end < today:
            fill = RED_FILL_NEW if is_new else RED_FILL
        elif end is not None and end <= soon_threshold:
            fill = YELLOW_FILL_NEW if is_new else YELLOW_FILL
        elif is_new:
            fill = NEW_FILL
        else:
            fill = None
        if fill is not None:
            for c in range(1, len(QUEUE_HEADERS) + 1):
                ws.cell(row=i, column=c).fill = fill
        total_price_sum += _parse_money(j.get("total_price", ""))

    # Summary footer
    footer = len(jobs) + 3
    ws.cell(row=footer, column=1, value=f"Total jobs: {len(jobs)}").font = SECTION_FONT
    ws.cell(row=footer, column=TOTAL_PRICE_COL - 1, value="Total $ in process:").font = SECTION_FONT
    total_cell = ws.cell(row=footer, column=TOTAL_PRICE_COL, value=total_price_sum)
    total_cell.number_format = MONEY_FMT
    total_cell.font = SECTION_FONT

    # AutoFilter across the data rows
    if jobs:
        last_col = get_column_letter(len(QUEUE_HEADERS))
        ws.auto_filter.ref = f"A1:{last_col}{len(jobs) + 1}"

    ws.freeze_panes = "B2"  # keep the header row AND the Job # column visible
    _autosize(ws, num_cols=len(QUEUE_HEADERS))


def _write_history_tab(ws, history: Dict[str, Any]) -> None:
    """Archived jobs (left the queue, not yet returned), newest departure first."""
    headers = QUEUE_HEADERS + ["Last Seen"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

    if not history:
        ws.cell(row=2, column=1,
                value="(no archived orders yet — a job appears here after it drops off the queue)")
        ws.freeze_panes = "B2"  # keep the header row AND the Job # column visible
        _autosize(ws, num_cols=len(headers))
        return

    entries = sorted(history.values(), key=lambda e: e.get("last_seen", ""), reverse=True)
    for i, entry in enumerate(entries, start=2):
        _write_job_row(ws, i, entry.get("snapshot", {}))
        ws.cell(row=i, column=len(QUEUE_HEADERS) + 1, value=entry.get("last_seen", ""))

    last_col = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last_col}{len(entries) + 1}"
    ws.freeze_panes = "B2"  # keep the header row AND the Job # column visible
    _autosize(ws, num_cols=len(headers))


def build_workbook(
    jobs: List[Dict[str, Any]],
    diff: Dict[str, Any],
    briefing: Dict[str, Any],
    today: date,
    history: Dict[str, Any] | None = None,
) -> Path:
    # Jobs whose change order landed this run (CO# rose, or returned higher).
    co_changed_ids = {c.get("job") for c in diff.get("co_changed", [])}
    co_changed_ids |= {j.get("job") for j in diff.get("returning", []) if j.get("_co_returned")}

    wb = Workbook()
    changes_ws = wb.active
    changes_ws.title = "Changes"
    _write_changes_tab(changes_ws, briefing, diff, co_changed_ids=co_changed_ids)

    full_ws = wb.create_sheet("Full Queue")
    # Mark new + returning orders so they pop on the Full Queue tab too.
    new_ids = {j.get("job") for j in diff.get("new", []) if j.get("job")} | \
              {j.get("job") for j in diff.get("returning", []) if j.get("job")}
    _write_full_queue_tab(full_ws, jobs, today, new_job_ids=new_ids, co_changed_ids=co_changed_ids)

    history_ws = wb.create_sheet("History")
    _write_history_tab(history_ws, history or {})

    path = OUTPUT_DIR / f"queue_{today.isoformat()}.xlsx"
    try:
        wb.save(path)
    except PermissionError:
        # The file is almost always locked because it's open in Excel. Fall
        # back to a timestamped name so the run still produces a report.
        alt = OUTPUT_DIR / f"queue_{today.isoformat()}_{datetime.now():%H%M%S}.xlsx"
        log.warning(
            "Could not write %s (is it open in Excel?). Saving to %s instead.",
            path.name, alt.name,
        )
        wb.save(alt)
        path = alt
    log.info("Wrote Excel report: %s", path)
    return path
