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
LINK_FONT = Font(color="0563C1", underline="single")  # job-folder hyperlinks
DRIVE_RUN_FONT = Font(color="C55A11", bold=True)  # highly-custom (has a drive run)

QUEUE_HEADERS = [
    "Status", "Customer", "Primary Rep", "Ship With", "Job #", "Oper", "Item",
    "Design", "Description", "Size", "Arrangement", "Motor Pos", "Class", "Rotation",
    "Discharge", "% Width", "Wheel Type", "Design Temp", "Max Temp", "Special Temp",
    "CO#", "Assigned To", "Checker",
    "Start Date", "End Date", "Plan Hrs", "FanNet Date", "Total Price",
    "Note", "Flags", "Drive Run", "Folder",
]
TOTAL_PRICE_COL = 28  # 1-based column index of Total Price in QUEUE_HEADERS
DRIVE_RUN_COL = 31    # 1-based column index of the Drive Run (highly-custom) flag
FOLDER_COL = 32       # 1-based column index of the Folder hyperlink


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


def _autosize(ws, num_cols: int) -> None:
    for col in range(1, num_cols + 1):
        letter = get_column_letter(col)
        max_len = 0
        for cell in ws[letter]:
            val = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(val))
        ws.column_dimensions[letter].width = min(max(max_len + 2, 10), 60)


def _write_changes_tab(ws, briefing: Dict[str, Any], diff: Dict[str, Any],
                       co_changed_ids: set | None = None) -> None:
    co_changed_ids = co_changed_ids or set()
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

    # Changed
    ws.cell(row=row, column=1, value=f"Changed orders ({len(diff['changed'])})").font = SECTION_FONT
    row += 1
    if diff["changed"]:
        for c, h in enumerate(["Job #", "Customer", "Field", "Old value", "New value"], start=1):
            cell = ws.cell(row=row, column=c, value=h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
        row += 1
        for ch in diff["changed"]:
            for (field, old, new) in ch["changes"]:
                ws.cell(row=row, column=1, value=ch["job"])
                ws.cell(row=row, column=2, value=ch["customer"])
                ws.cell(row=row, column=3, value=field)
                ws.cell(row=row, column=4, value=old)
                ws.cell(row=row, column=5, value=new)
                row += 1
    else:
        ws.cell(row=row, column=1, value="(none)")
        row += 1
    row += 1

    # Change orders that landed this run (CO# rose since yesterday, including
    # jobs that came back at a higher CO#).
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

    _autosize(ws, num_cols=len(QUEUE_HEADERS))


def _co_label(j: Dict[str, Any]) -> str:
    co = j.get("co_number") or 0
    return f"CO#{co}" if co else ""


def _write_job_row(ws, row: int, j: Dict[str, Any], co_changed: bool = False) -> None:
    ws.cell(row=row, column=1, value=j.get("status", ""))
    ws.cell(row=row, column=2, value=j.get("customer", ""))
    ws.cell(row=row, column=3, value=j.get("primary_rep", ""))
    ws.cell(row=row, column=4, value=j.get("ship_with", ""))
    ws.cell(row=row, column=5, value=j.get("job", ""))
    ws.cell(row=row, column=6, value=j.get("oper", ""))
    ws.cell(row=row, column=7, value=j.get("item", ""))
    ws.cell(row=row, column=8, value=j.get("design", ""))
    ws.cell(row=row, column=9, value=j.get("so_design_desc", ""))
    ws.cell(row=row, column=10, value=j.get("so_size", ""))
    ws.cell(row=row, column=11, value=j.get("so_arrangement", ""))
    ws.cell(row=row, column=12, value=j.get("so_motor_pos", ""))
    ws.cell(row=row, column=13, value=j.get("so_class", ""))
    ws.cell(row=row, column=14, value=j.get("so_rotation", ""))
    ws.cell(row=row, column=15, value=j.get("so_discharge", ""))
    ws.cell(row=row, column=16, value=j.get("so_pct_width", ""))
    ws.cell(row=row, column=17, value=j.get("so_wheel_type", ""))
    ws.cell(row=row, column=18, value=j.get("so_design_temp", ""))
    ws.cell(row=row, column=19, value=j.get("so_max_temp", ""))
    ws.cell(row=row, column=20, value=j.get("so_special_temp", ""))
    ws.cell(row=row, column=21, value=_co_label(j))
    ws.cell(row=row, column=22, value=j.get("assigned_to", ""))
    ws.cell(row=row, column=23, value=j.get("checker", ""))
    ws.cell(row=row, column=24, value=j.get("start_date", ""))
    ws.cell(row=row, column=25, value=j.get("end_date", ""))
    ws.cell(row=row, column=26, value=j.get("plan_hrs", ""))
    ws.cell(row=row, column=27, value=j.get("fannet_date", ""))
    _write_money_cell(ws, row, TOTAL_PRICE_COL, j.get("total_price", ""))
    ws.cell(row=row, column=29, value=j.get("status_note", ""))
    ws.cell(row=row, column=30, value=_flags_str(j))

    # Drive Run: YES when a CBC_DriveRun exists -> highly-custom fan.
    dr_cell = ws.cell(row=row, column=DRIVE_RUN_COL, value="YES" if j.get("has_drive_run") else "")
    if j.get("has_drive_run"):
        dr_cell.font = DRIVE_RUN_FONT

    # Folder hyperlink (AutoCAD job folder, or the SO archive folder as fallback).
    folder = (j.get("job_folder") or "").strip()
    fcell = ws.cell(row=row, column=FOLDER_COL, value=(j.get("job_type") or "Open") if folder else "")
    if folder:
        fcell.hyperlink = folder
        fcell.font = LINK_FONT

    # A change order that landed this run -> the whole row's text goes red.
    if co_changed:
        for c in range(1, FOLDER_COL):  # leave the hyperlink cell its link style
            ws.cell(row=row, column=c).font = RED_FONT


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

    ws.freeze_panes = "A2"
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
        ws.freeze_panes = "A2"
        _autosize(ws, num_cols=len(headers))
        return

    entries = sorted(history.values(), key=lambda e: e.get("last_seen", ""), reverse=True)
    for i, entry in enumerate(entries, start=2):
        _write_job_row(ws, i, entry.get("snapshot", {}))
        ws.cell(row=i, column=len(QUEUE_HEADERS) + 1, value=entry.get("last_seen", ""))

    last_col = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last_col}{len(entries) + 1}"
    ws.freeze_panes = "A2"
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
