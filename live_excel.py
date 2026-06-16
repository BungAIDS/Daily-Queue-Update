"""Render the live master workbook by driving the desktop Excel app through COM —
the same no-password, use-the-signed-in-app trick emailer.py uses for Outlook.

Why COM and not openpyxl: the daily report (excel_writer.py) is written with
openpyxl, which *replaces the whole file* — that can't touch a Microsoft 365
co-authored workbook without kicking everyone out / conflicting. Driving the real
Excel application means edits flow through Excel itself, which syncs them to
OneDrive/SharePoint so coworkers see them live (cursors and all).

What it writes: one worksheet per `live_sheets.Sheet` model (Live Queue, Changes,
History, Line Items). The *content* lives in live_sheets.py (pure, tested); this
module is the generic renderer — bulk-write the values, then map each cell's
named fill/font to real Excel colors, add hyperlinks, freeze panes, and
AutoFilter. The named styles mirror excel_writer so the live master and the daily
report look the same.

Everything is lazy-imported and best-effort: a failed Excel update logs and the
poll cycle carries on (state + notifications still happen).
"""
from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Any, Dict, List, Optional

from live_sheets import Sheet

log = logging.getLogger(__name__)

_XL_UNDERLINE_SINGLE = 2  # xlUnderlineStyleSingle


def _bgr(rgb_hex: str) -> int:
    """Excel COM colors are BGR longs; convert an 'RRGGBB' hex string."""
    r, g, b = int(rgb_hex[0:2], 16), int(rgb_hex[2:4], 16), int(rgb_hex[4:6], 16)
    return (b << 16) | (g << 8) | r


# Named fills -> Excel BGR (same RGB values as excel_writer's PatternFills).
_FILL_RGB = {
    "header": "305496",
    "overdue": "FFC7CE", "soon": "FFEB9C", "new": "D9D9D9",
    "overdue_new": "F4A5A8", "soon_new": "F5D750",
    "dwg_yes": "C6EFCE", "dwg_no": "FFC7CE",
}
_FILL = {k: _bgr(v) for k, v in _FILL_RGB.items()}

# Named fonts -> (rgb color or None, bold, underline, size or None).
_FONT = {
    "header": ("FFFFFF", True, False, None),
    "section": (None, True, False, 12),
    "link": ("0563C1", False, True, None),
    "drive_run": ("C55A11", True, False, None),
    "drive_run_link": ("C55A11", True, True, None),
    "red": ("C00000", True, False, None),
}

SHEET_ORDER = ["Live Queue", "Changes", "History", "Line Items"]

# Last-rendered fingerprint per sheet, so a tab is only repainted when its
# content actually changed — otherwise a coworker's active filter/scroll on that
# tab would be reset every poll. Reset on process start (re-renders once).
_RENDER_CACHE: Dict[str, int] = {}


def _fingerprint(sheet: Sheet) -> int:
    parts = [sheet.name, sheet.freeze or "", sheet.autofilter_a1 or ""]
    for row in sheet.grid:
        for cell in row:
            parts.append(f"{cell.value}|{cell.fill}|{cell.font}|{cell.link}|"
                         f"{cell.number_format}|{cell.center}")
        parts.append(";")
    return hash("\n".join(parts))


# --------------------------------------------------------------------------- #
# COM plumbing                                                                 #
# --------------------------------------------------------------------------- #
def _get_excel():
    """Attach to a running Excel, or start one. Lazy COM import (Windows-only)."""
    import win32com.client  # type: ignore
    try:
        app = win32com.client.GetActiveObject("Excel.Application")
    except Exception:  # noqa: BLE001 - not running; launch it
        app = win32com.client.Dispatch("Excel.Application")
    app.Visible = True
    return app


def _find_workbook(app, path: Path):
    """The already-open workbook matching `path`, or open it. Co-authored files
    live in OneDrive/SharePoint, so FullName may be a URL — match on file name
    first, full path second."""
    name = path.name
    target = os.path.normcase(str(path))
    for w in app.Workbooks:
        try:
            if w.Name == name or os.path.normcase(w.FullName) == target:
                return w
        except Exception:  # noqa: BLE001
            continue
    return app.Workbooks.Open(str(path))


def _get_or_make_sheet(wb, name: str):
    for s in wb.Worksheets:
        if s.Name == name:
            return s
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    return ws


# --------------------------------------------------------------------------- #
# Styling                                                                      #
# --------------------------------------------------------------------------- #
def _apply_font(rng, font_name: Optional[str]) -> None:
    spec = _FONT.get(font_name or "")
    if not spec:
        return
    color, bold, underline, size = spec
    f = rng.Font
    f.Bold = bool(bold)
    if color:
        f.Color = _bgr(color)
    f.Underline = _XL_UNDERLINE_SINGLE if underline else False
    if size:
        f.Size = size


def _apply_run(ws, r: int, c1: int, c2: int, fill: Optional[str],
               font: Optional[str], center: bool) -> None:
    rng = ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
    if fill and fill in _FILL:
        rng.Interior.Color = _FILL[fill]
    if font:
        _apply_font(rng, font)
    if center:
        rng.HorizontalAlignment = -4108  # xlCenter


def _style_row(ws, r: int, cells: List) -> None:
    """Style one model row: collapse adjacent cells that share (fill, font,
    center) into a single Range (few COM calls), then add hyperlinks and number
    formats per cell."""
    n = len(cells)
    c = 0
    while c < n:
        key = (cells[c].fill, cells[c].font, cells[c].center)
        if key == (None, None, False):
            c += 1
            continue
        c2 = c
        while c2 + 1 < n and (cells[c2 + 1].fill, cells[c2 + 1].font,
                              cells[c2 + 1].center) == key:
            c2 += 1
        _apply_run(ws, r, c + 1, c2 + 1, *key)
        c = c2 + 1
    for i, cell in enumerate(cells, start=1):
        if cell.link:
            try:
                ws.Hyperlinks.Add(Anchor=ws.Cells(r, i), Address=str(cell.link))
                _apply_font(ws.Cells(r, i), cell.font)  # keep our link style
            except Exception:  # noqa: BLE001 - a bad path shouldn't stop the row
                pass
        if cell.number_format:
            ws.Cells(r, i).NumberFormat = cell.number_format


def _pad(row: List, ncols: int) -> List[Any]:
    vals = [(cell.value if cell.value is not None else "") for cell in row]
    return vals + [""] * (ncols - len(vals))


def render_sheet(app, wb, sheet: Sheet) -> None:
    """Write one Sheet model into its worksheet: clear, bulk-write values, style,
    freeze, AutoFilter, autofit."""
    ws = _get_or_make_sheet(wb, sheet.name)
    nrows, ncols = sheet.nrows, sheet.ncols
    try:
        if ws.AutoFilterMode:
            ws.AutoFilterMode = False
    except Exception:  # noqa: BLE001
        pass
    ws.Cells.Clear()  # bot-owned sheet — full repaint keeps it correct
    if nrows == 0 or ncols == 0:
        return

    ws.Range(ws.Cells(1, 1), ws.Cells(nrows, ncols)).Value = [
        _pad(row, ncols) for row in sheet.grid
    ]
    for r, row in enumerate(sheet.grid, start=1):
        if row:
            _style_row(ws, r, row)

    if sheet.autofilter_a1:
        try:
            ws.Range(sheet.autofilter_a1).AutoFilter()
        except Exception:  # noqa: BLE001
            pass

    try:
        ws.UsedRange.Columns.AutoFit()
        for col in range(1, ncols + 1):
            if ws.Columns(col).ColumnWidth > 60:
                ws.Columns(col).ColumnWidth = 60
    except Exception:  # noqa: BLE001
        pass

    if sheet.freeze:
        try:
            ws.Activate()
            app.ActiveWindow.FreezePanes = False
            ws.Range(sheet.freeze).Select()
            app.ActiveWindow.FreezePanes = True
        except Exception:  # noqa: BLE001
            pass


def update_workbook(sheets: List[Sheet], workbook_path: str | Path) -> bool:
    """Render all sheet models into the live workbook via Excel COM. Returns True
    on success. Best-effort: any COM error is logged and swallowed so the poll
    cycle continues."""
    path = Path(workbook_path)
    try:
        app = _get_excel()
    except Exception as e:  # noqa: BLE001
        log.warning("Could not reach Excel via COM (%s); live workbook not updated. "
                    "On Windows, ensure Excel is installed and signed in.", e)
        return False
    try:
        wb = _find_workbook(app, path)
        order = {n: i for i, n in enumerate(SHEET_ORDER)}
        rendered = []
        for sheet in sorted(sheets, key=lambda s: order.get(s.name, 99)):
            fp = _fingerprint(sheet)
            if _RENDER_CACHE.get(sheet.name) == fp:
                continue  # unchanged — leave it alone so filters/scroll persist
            render_sheet(app, wb, sheet)
            _RENDER_CACHE[sheet.name] = fp
            rendered.append(sheet.name)
        if rendered:
            try:
                wb.Save()
            except Exception as e:  # noqa: BLE001
                log.debug("wb.Save() raised (likely AutoSave-managed): %s", e)
            log.info("Live workbook updated: %s [%s]", path.name, ", ".join(rendered))
        else:
            log.info("Live workbook unchanged this cycle — nothing repainted.")
        return True
    except Exception as e:  # noqa: BLE001
        log.warning("Live workbook update failed (%s); state + alerts still recorded.", e)
        return False


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
