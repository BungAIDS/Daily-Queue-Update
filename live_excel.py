"""Write the live queue into your co-authored Excel workbook by driving the
desktop Excel app through COM — the same no-password, use-the-signed-in-app
trick emailer.py uses for Outlook.

Why COM and not openpyxl: the daily report (excel_writer.py) is written with
openpyxl, which *replaces the whole file*. That can't touch a Microsoft 365
co-authored workbook without kicking everyone out / conflicting. Driving the
real Excel application means edits flow through Excel itself — it syncs them to
OneDrive/SharePoint, so coworkers see them appear live (cursors and all), and
there's no file-lock fight because we're not writing the file out of band.

Requirements (Windows): Excel installed and signed into the same Microsoft 365
account, with the workbook stored in OneDrive/SharePoint. The watcher PC keeps
Excel running 5am-5pm; this module attaches to that running instance (opening
the workbook if it isn't already).

Everything is lazy-imported and best-effort: a failed Excel update logs and the
poll cycle carries on (the state + notifications still happen).
"""
from __future__ import annotations

import logging
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple

# These are plain helpers (no Excel/openpyxl side effects at call time); reuse
# them so the live sheet labels match the daily report exactly.
from excel_writer import _co_label, _drive_run_label, _flags_str, _parse_money

log = logging.getLogger(__name__)

# Live-sheet columns: (header, key). "Added" leads so the newest arrivals — which
# present_jobs sorts to the top — are read first. A handful of special keys are
# rendered below in _cell_value; the rest print the job field verbatim.
LIVE_COLUMNS: List[Tuple[str, str]] = [
    ("Added", "_added"),
    ("Job #", "job"),
    ("CO#", "_co"),
    ("Quote Run", "_drive_run"),
    ("Oper", "oper"),
    ("Design", "design"),
    ("Description", "so_design_desc"),
    ("Size", "so_size"),
    ("Arrangement", "so_arrangement"),
    ("Features", "line_item_tags"),
    ("Customer", "customer"),
    ("Primary Rep", "primary_rep"),
    ("Assigned To", "assigned_to"),
    ("Checker", "checker"),
    ("Start Date", "start_date"),
    ("End Date", "end_date"),
    ("FanNet Date", "fannet_date"),
    ("Plan Hrs", "plan_hrs"),
    ("Total Price", "_money"),
    ("Ship With", "ship_with"),
    ("Note", "status_note"),
    ("Flags", "_flags"),
    ("Status", "status"),
]

SHEET_NAME = "Live Queue"
_HEADER_BG = 0x965430   # excel_writer's header blue (305496) as Excel BGR
_NEW_BG = 0xCEF0C6      # light green for orders added during today's watch (BGR)
_WHITE = 0xFFFFFF


def added_label(job: Dict[str, Any], ref: datetime | None = None) -> str:
    """Human-friendly 'time it was added'. Carried-over orders (already in the
    queue when the watch began) show a plain marker rather than a fake time."""
    if job.get("_carried_over"):
        return "before watch"
    iso = job.get("_first_seen") or ""
    try:
        dt = datetime.fromisoformat(iso)
    except ValueError:
        return iso
    ref = ref or datetime.now()
    if dt.date() == ref.date():
        return dt.strftime("%-I:%M %p") if os.name != "nt" else dt.strftime("%#I:%M %p")
    return dt.strftime("%b %-d %-I:%M %p") if os.name != "nt" else dt.strftime("%b %#d %#I:%M %p")


def _cell_value(job: Dict[str, Any], key: str) -> Any:
    if key == "_added":
        return added_label(job)
    if key == "_co":
        return _co_label(job)
    if key == "_drive_run":
        return _drive_run_label(job)
    if key == "_flags":
        return _flags_str(job)
    if key == "_money":
        raw = (job.get("total_price") or "").strip()
        return _parse_money(raw) if raw else ""
    return job.get(key, "")


def _values_grid(rows: List[Dict[str, Any]]) -> List[List[Any]]:
    """Header row + one row per order, as a plain 2D grid for a single bulk
    write to the sheet (one Excel edit, not thousands of cell pokes)."""
    grid: List[List[Any]] = [[h for h, _ in LIVE_COLUMNS]]
    for j in rows:
        grid.append([_cell_value(j, key) for _, key in LIVE_COLUMNS])
    return grid


def _get_excel():
    """Attach to a running Excel, or start one. Lazy COM import (Windows-only)."""
    import win32com.client  # type: ignore
    try:
        app = win32com.client.GetActiveObject("Excel.Application")
    except Exception:  # noqa: BLE001 - not currently running; launch it
        app = win32com.client.Dispatch("Excel.Application")
    app.Visible = True
    return app


def _find_workbook(app, path: Path):
    """The already-open workbook matching `path`, or open it. Co-authored files
    live in OneDrive/SharePoint, so FullName may be a URL — match on the file
    name first, full path second."""
    name = path.name
    target = os.path.normcase(str(path))
    for w in app.Workbooks:
        try:
            if w.Name == name or os.path.normcase(w.FullName) == target:
                return w
        except Exception:  # noqa: BLE001 - a workbook can be in a weird state; skip it
            continue
    return app.Workbooks.Open(str(path))


def _get_or_make_sheet(wb, name: str):
    for s in wb.Worksheets:
        if s.Name == name:
            return s, False
    ws = wb.Worksheets.Add()
    ws.Name = name
    return ws, True


def update_live_sheet(rows: List[Dict[str, Any]], workbook_path: str | Path) -> bool:
    """Write the current board into the live workbook's 'Live Queue' sheet via
    Excel COM. Returns True on success. Best-effort: any COM error is logged and
    swallowed so the poll cycle continues."""
    path = Path(workbook_path)
    try:
        app = _get_excel()
    except Exception as e:  # noqa: BLE001
        log.warning("Could not reach Excel via COM (%s); live sheet not updated. "
                    "On Windows, ensure Excel is installed and signed in.", e)
        return False

    grid = _values_grid(rows)
    nrows, ncols = len(grid), len(LIVE_COLUMNS)
    try:
        wb = _find_workbook(app, path)
        ws, created = _get_or_make_sheet(wb, SHEET_NAME)

        # One bulk write of the whole grid — a single co-authoring edit.
        rng = ws.Range(ws.Cells(1, 1), ws.Cells(nrows, ncols))
        rng.Value = grid

        # Clear any rows left over from a previous, longer board.
        try:
            used = ws.UsedRange.Rows.Count
            if used > nrows:
                ws.Range(ws.Cells(nrows + 1, 1), ws.Cells(used, ncols)).ClearContents()
        except Exception:  # noqa: BLE001 - cosmetic cleanup, never fatal
            pass

        _format_sheet(ws, app, nrows, ncols, rows, created)

        # Persist. On a OneDrive AutoSave workbook this is effectively a no-op,
        # but it makes the non-AutoSave case (local share) flush each cycle.
        try:
            wb.Save()
        except Exception as e:  # noqa: BLE001
            log.debug("wb.Save() raised (likely AutoSave-managed): %s", e)
        log.info("Live sheet updated: %d orders -> %s [%s]", len(rows), path.name, SHEET_NAME)
        return True
    except Exception as e:  # noqa: BLE001
        log.warning("Live Excel update failed (%s); state + alerts still recorded.", e)
        return False


def _format_sheet(ws, app, nrows: int, ncols: int, rows: List[Dict[str, Any]], created: bool) -> None:
    """Light, stable formatting only — header styling, a freeze pane, autofit,
    and a green tint on orders added during today's watch. Kept idempotent
    (same result every cycle) so it doesn't churn the co-authored file."""
    try:
        header = ws.Range(ws.Cells(1, 1), ws.Cells(1, ncols))
        header.Font.Bold = True
        header.Font.Color = _WHITE
        header.Interior.Color = _HEADER_BG
        if created:
            ws.Application.ActiveWindow.FreezePanes = False
            ws.Cells(2, 2).Select()
            app.ActiveWindow.FreezePanes = True
    except Exception:  # noqa: BLE001
        pass

    # Tint rows for orders that arrived during the watch (not carried over), so
    # the genuinely-new work stands out. Stable per order -> no co-author churn.
    try:
        for i, j in enumerate(rows, start=2):
            color = _WHITE if j.get("_carried_over") else _NEW_BG
            ws.Range(ws.Cells(i, 1), ws.Cells(i, ncols)).Interior.Color = color
    except Exception:  # noqa: BLE001
        pass

    try:
        ws.Range(ws.Cells(1, 1), ws.Cells(max(nrows, 1), ncols)).Columns.AutoFit()
    except Exception:  # noqa: BLE001
        pass


def save_morning_copy(workbook_path: str | Path, dest: str | Path) -> bool:
    """Freeze a dated copy of the workbook as the morning snapshot, using Excel's
    SaveCopyAs (doesn't disturb the live file or its co-authors)."""
    src, dest = Path(workbook_path), Path(dest)
    try:
        app = _get_excel()
        wb = _find_workbook(app, src)
        dest.parent.mkdir(parents=True, exist_ok=True)
        wb.SaveCopyAs(str(dest))
        log.info("Saved morning snapshot copy: %s", dest)
        return True
    except Exception as e:  # noqa: BLE001
        log.warning("Could not save morning snapshot copy (%s)", e)
        return False
