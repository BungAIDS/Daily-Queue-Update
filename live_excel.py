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

from openpyxl.utils import get_column_letter

from live_sheets import Sheet

log = logging.getLogger(__name__)

_XL_UNDERLINE_SINGLE = 2  # xlUnderlineStyleSingle
_XL_EXPRESSION = 2        # xlExpression (conditional formatting)


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
    "sep": "808080",   # the vertical divider column between the two matrices
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

# "History" is a RESERVED worksheet name in Excel (shared-workbook change
# tracking), so the archived-orders tab is "Order History".
SHEET_ORDER = ["Live Queue", "Changes", "Order History", "Line Items"]

# Last-rendered fingerprint per sheet, so a tab is only repainted when its
# content actually changed — otherwise a coworker's active filter/scroll on that
# tab would be reset every poll. Reset on process start (re-renders once).
_RENDER_CACHE: Dict[str, int] = {}

# Sheets whose freeze pane is already set this process. Freeze panes persist in
# the workbook and setting one requires Activating the sheet — which would yank a
# coworker's view to that tab — so we only ever do it once per sheet.
_FROZEN: set = set()

# Upsert tabs (Live Queue, Order History) whose header row + initial autofit are
# already done this process, so we don't rewrite the header every cycle.
_HEADER_DONE: set = set()

_XL_UP = -4162           # xlUp
_XL_NONE = -4142         # xlColorIndexNone
_XL_AUTO = -4105         # xlColorIndexAutomatic


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
    # Best-effort: when Excel is busy (co-authoring sync, a dialog, you typing in
    # a cell) setting Visible is rejected. It's only a nicety, so never let it
    # abort the update — Excel is already visible if it was running.
    try:
        app.Visible = True
    except Exception:  # noqa: BLE001
        pass
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
    try:
        ws.Name = name
    except Exception:  # noqa: BLE001 - invalid/reserved name; don't leave a stray
        try:
            ws.Application.DisplayAlerts = False
            ws.Delete()
            ws.Application.DisplayAlerts = True
        except Exception:  # noqa: BLE001
            pass
        raise
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

    # Freeze panes persist in the workbook and setting one steals focus to this
    # tab, so do it only once per sheet (not every repaint).
    if sheet.freeze and sheet.name not in _FROZEN:
        try:
            ws.Activate()
            app.ActiveWindow.FreezePanes = False
            ws.Range(sheet.freeze).Select()
            app.ActiveWindow.FreezePanes = True
            _FROZEN.add(sheet.name)
        except Exception:  # noqa: BLE001
            pass


# --------------------------------------------------------------------------- #
# Incremental upsert renderer (Live Queue + Order History)                     #
#                                                                             #
# Rows are keyed on the order number: append new ones, update changed ones in  #
# place, delete departed ones (Live Queue). No Cells.Clear(), so a coworker's  #
# filter/sort/scroll survives — only an add/remove shifts an active filter.    #
# --------------------------------------------------------------------------- #
def _norm_key(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def _write_header(ws, headers: List[str]) -> None:
    ncols = len(headers)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, ncols)).Value = [list(headers)]
    _apply_run(ws, 1, 1, ncols, "header", "header", False)


def _read_keymap(ws, key_col: int):
    """{order# -> row} for the data rows, by reading the key column. Robust to a
    coworker sorting the sheet (we find rows by key, not a remembered index)."""
    last = ws.Cells(ws.Rows.Count, key_col).End(_XL_UP).Row
    keymap: Dict[str, int] = {}
    if last >= 2:
        data = ws.Range(ws.Cells(2, key_col), ws.Cells(last, key_col)).Value
        if not isinstance(data, tuple):       # single cell -> scalar
            data = ((data,),)
        for i, row in enumerate(data, start=2):
            v = row[0] if isinstance(row, (list, tuple)) else row
            k = _norm_key(v)
            if k:
                keymap[k] = i
    return keymap, last


def _write_row(ws, r: int, cells: List, ncols: int) -> None:
    """Overwrite one row's values + styling in place. Clears the row's prior
    fill/font/links first so a cleared style doesn't linger."""
    vals = [(c.value if c.value is not None else "") for c in cells]
    if len(vals) < ncols:
        vals += [""] * (ncols - len(vals))
    rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, ncols))
    rng.Value = [vals]
    try:
        rng.Interior.ColorIndex = _XL_NONE
        rng.Font.Bold = False
        rng.Font.Underline = False
        rng.Font.ColorIndex = _XL_AUTO
        rng.Hyperlinks.Delete()
    except Exception:  # noqa: BLE001
        pass
    _style_row(ws, r, cells)


def apply_upserts(app, wb, name: str, headers: List[str], ops: List,
                  key_col: int, allow_delete: bool, freeze: str | None = None,
                  sort_col: int | None = None, text_cols: List[int] | None = None) -> int:
    """Apply append/update/delete ops to a keyed sheet. Returns rows touched.
    When `sort_col` is set, the data is re-sorted ascending by that column after
    any change (used to keep Live Queue ordered by End Date, overdue at the top).
    `text_cols` are pre-formatted as Text before any data is written so Excel
    can't coerce a value (e.g. the AM/PM 'Added' label) into a datetime serial."""
    ws = _get_or_make_sheet(wb, name)
    ncols = len(headers)
    first_time = name not in _HEADER_DONE
    if first_time:
        # Once per process: wipe the sheet so a previous run's (possibly
        # different-schema) content can't collide with the keyed upsert. The
        # watcher resets the stored sigs at startup to match, so this cycle
        # rebuilds the tab; every later cycle is incremental.
        ws.Cells.Clear()
        _write_header(ws, headers)
        for col in (text_cols or []):
            try:
                ws.Columns(col).NumberFormat = "@"   # Text — must precede any write
            except Exception:  # noqa: BLE001
                pass
        _HEADER_DONE.add(name)

    keymap, last_row = _read_keymap(ws, key_col)
    deletes = [k for kind, k, _ in ops if kind == "delete"]
    updates = [(k, c) for kind, k, c in ops if kind == "update"]
    appends = [(k, c) for kind, k, c in ops if kind == "append"]
    rowcount_changed = bool(deletes or appends)

    for r in sorted((keymap[k] for k in deletes if k in keymap), reverse=True):
        try:
            ws.Rows(r).Delete()
        except Exception:  # noqa: BLE001
            pass
    if deletes:
        keymap, last_row = _read_keymap(ws, key_col)

    for k, cells in updates:
        r = keymap.get(k)
        if r:
            _write_row(ws, r, cells, ncols)
        else:
            appends.append((k, cells))   # vanished from the sheet — re-add it
    for k, cells in appends:
        last_row += 1
        _write_row(ws, last_row, cells, ncols)

    # Keep the tab sorted (Live Queue by End Date — overdue/red at the top) and
    # re-extend AutoFilter. Sort the DATA ROWS ONLY (row 2 down) so the header
    # never moves; blanks fall to the bottom under ascending order.
    touched = bool(deletes or updates or appends)
    if touched and last_row >= 2:
        try:
            if ws.AutoFilterMode:
                ws.AutoFilterMode = False
            if sort_col and last_row >= 3:
                ws.Range(ws.Cells(2, 1), ws.Cells(last_row, ncols)).Sort(
                    Key1=ws.Cells(2, sort_col), Order1=1, Header=2)  # xlAscending, xlNo
            ws.Range(ws.Cells(1, 1), ws.Cells(last_row, ncols)).AutoFilter()
        except Exception:  # noqa: BLE001
            pass

    if first_time:
        try:
            ws.UsedRange.Columns.AutoFit()
            for col in range(1, ncols + 1):
                if ws.Columns(col).ColumnWidth > 60:
                    ws.Columns(col).ColumnWidth = 60
        except Exception:  # noqa: BLE001
            pass

    if freeze and name not in _FROZEN:
        try:
            ws.Activate()
            app.ActiveWindow.FreezePanes = False
            ws.Range(freeze).Select()
            app.ActiveWindow.FreezePanes = True
            _FROZEN.add(name)
        except Exception:  # noqa: BLE001
            pass

    return len(updates) + len(appends) + len(deletes)


def reset_sheet(name: str) -> None:
    """Forget that a sheet's header/freeze are done, so the next render rebuilds
    it from scratch — used when its column schema changed (e.g. a new DWG suffix
    or feature tag appeared in the Order History matrices)."""
    _HEADER_DONE.discard(name)
    _FROZEN.discard(name)


def _cell_to_value(cell) -> Any:
    """A bulk-writable value for a cell. A hyperlink becomes a =HYPERLINK()
    formula so thousands of links go in one Range write (Hyperlinks.Add per cell
    would crawl over a 12K-row tab)."""
    if cell.link:
        url = str(cell.link).replace('"', '""')
        disp = str(cell.value if cell.value is not None else "").replace('"', '""')
        return f'=HYPERLINK("{url}","{disp}")'
    return cell.value if cell.value is not None else ""


def _write_oh_row(ws, r: int, cells: List, ncols: int) -> None:
    vals = [_cell_to_value(c) for c in cells]
    if len(vals) < ncols:
        vals += [""] * (ncols - len(vals))
    ws.Range(ws.Cells(r, 1), ws.Cells(r, ncols)).Value = [vals]


def _apply_matrix_cf(ws, key_col: int, c0: int, c1: int) -> None:
    """Color a ✓/blank matrix block by conditional formatting — green for ✓, red
    for blank — but only on rows that actually have an order number (the key
    column guard), so the empty area below the data isn't painted. Applied over
    the whole columns so appended rows are colored automatically."""
    if c1 < c0:
        return
    key = get_column_letter(key_col)
    tl = get_column_letter(c0)            # top-left of the CF range, for the relative formula
    rng = ws.Range(ws.Cells(2, c0), ws.Cells(ws.Rows.Count, c1))
    try:
        rng.FormatConditions.Delete()
        green = rng.FormatConditions.Add(Type=_XL_EXPRESSION,
                                         Formula1=f'=AND(${key}2<>"",{tl}2="✓")')
        green.Interior.Color = _FILL["dwg_yes"]
        red = rng.FormatConditions.Add(Type=_XL_EXPRESSION,
                                       Formula1=f'=AND(${key}2<>"",{tl}2="")')
        red.Interior.Color = _FILL["dwg_no"]
    except Exception as e:  # noqa: BLE001 - values still readable without color
        log.debug("Matrix conditional formatting failed (%s)", e)


def _draw_separator(ws, sep_col: int) -> None:
    try:
        col = ws.Columns(sep_col)
        col.Interior.Color = _FILL["sep"]
        col.ColumnWidth = 2
    except Exception:  # noqa: BLE001
        pass


def apply_order_history(app, wb, name: str, spec: Dict[str, Any], ops: List,
                        key_col: int, freeze: str | None = None) -> int:
    """Render the Order History log. On the first touch this process: clear, write
    the header, BULK-write every row at once, color the two matrices via
    conditional formatting, and draw the divider. After that: append/update only
    the few changed rows (a stable presence log, so this is rare)."""
    ws = _get_or_make_sheet(wb, name)
    headers = spec["headers"]
    ncols = len(headers)
    if name not in _HEADER_DONE:
        ws.Cells.Clear()
        _write_header(ws, headers)
        _HEADER_DONE.add(name)
        records = spec["records"]
        if records:
            grid = [[_cell_to_value(c) for c in cells] + [""] * (ncols - len(cells))
                    for _, cells in records]
            ws.Range(ws.Cells(2, 1), ws.Cells(1 + len(grid), ncols)).Value = grid
        _apply_matrix_cf(ws, key_col, *spec["dwg_range"])
        _apply_matrix_cf(ws, key_col, *spec["feat_range"])
        _draw_separator(ws, spec["sep_col"])
        # AutoFilter so it's sortable/filterable by any column, like Live Queue.
        if records:
            try:
                ws.Range(ws.Cells(1, 1), ws.Cells(1 + len(records), ncols)).AutoFilter()
            except Exception:  # noqa: BLE001
                pass
        try:
            ws.UsedRange.Columns.AutoFit()
            for col in range(1, ncols + 1):
                if col == spec["sep_col"]:
                    continue
                if ws.Columns(col).ColumnWidth > 40:
                    ws.Columns(col).ColumnWidth = 40
        except Exception:  # noqa: BLE001
            pass
        if freeze and name not in _FROZEN:
            try:
                ws.Activate()
                app.ActiveWindow.FreezePanes = False
                ws.Range(freeze).Select()
                app.ActiveWindow.FreezePanes = True
                _FROZEN.add(name)
            except Exception:  # noqa: BLE001
                pass
        return len(spec["records"])

    # Incremental: append new orders / update the few whose flags changed.
    keymap, last_row = _read_keymap(ws, key_col)
    updates = [(k, c) for kind, k, c in ops if kind == "update"]
    appends = [(k, c) for kind, k, c in ops if kind == "append"]
    for k, cells in updates:
        r = keymap.get(k)
        if r:
            _write_oh_row(ws, r, cells, ncols)
        else:
            appends.append((k, cells))
    for k, cells in appends:
        last_row += 1
        _write_oh_row(ws, last_row, cells, ncols)
    if appends:   # re-extend AutoFilter to cover the newly appended rows
        try:
            if ws.AutoFilterMode:
                ws.AutoFilterMode = False
            ws.Range(ws.Cells(1, 1), ws.Cells(last_row, ncols)).AutoFilter()
        except Exception:  # noqa: BLE001
            pass
    return len(updates) + len(appends)


def update_master_workbook(workbook_path: str | Path, lq_payload: Dict[str, Any],
                           oh_payload: Dict[str, Any],
                           changes_sheet: Sheet | None = None) -> bool:
    """Render the master workbook: an incremental upsert for Live Queue, the
    matrix log for Order History, and a full repaint for the Changes snapshot
    (only when changed). Best-effort — any COM error is logged and swallowed."""
    path = Path(workbook_path)
    try:
        app = _get_excel()
    except Exception as e:  # noqa: BLE001
        log.warning("Could not reach Excel via COM (%s); live workbook not updated. "
                    "On Windows, ensure Excel is installed and signed in.", e)
        return False
    try:
        wb = _find_workbook(app, path)
        touched = []
        n = apply_upserts(app, wb, lq_payload["name"], lq_payload["headers"], lq_payload["ops"],
                          lq_payload["key_col"], lq_payload["allow_delete"], lq_payload.get("freeze"),
                          sort_col=lq_payload.get("sort_col"), text_cols=lq_payload.get("text_cols"))
        if n:
            touched.append(f"{lq_payload['name']}(+{n})")
        n = apply_order_history(app, wb, oh_payload["name"], oh_payload["spec"],
                                oh_payload["ops"], oh_payload["key_col"], oh_payload.get("freeze"))
        if n:
            touched.append(f"{oh_payload['name']}(+{n})")
        if changes_sheet is not None:
            fp = _fingerprint(changes_sheet)
            if _RENDER_CACHE.get(changes_sheet.name) != fp:
                render_sheet(app, wb, changes_sheet)
                _RENDER_CACHE[changes_sheet.name] = fp
                touched.append(changes_sheet.name)
        if touched:
            try:
                wb.Save()
            except Exception as e:  # noqa: BLE001
                log.debug("wb.Save() raised (likely AutoSave-managed): %s", e)
            log.info("Live workbook updated: %s [%s]", path.name, ", ".join(touched))
        else:
            log.info("Live workbook unchanged this cycle — nothing written.")
        return True
    except Exception as e:  # noqa: BLE001
        log.warning("Live workbook update failed (%s); state + alerts still recorded.", e)
        return False


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
