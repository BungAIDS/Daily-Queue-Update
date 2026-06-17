"""Pure 'sheet model' for the live master workbook — what each tab contains and
how each cell should look, with NO Excel/COM dependency.

This is the single source of truth for the live workbook's content, kept
separate from how it's pushed into Excel (live_excel.py drives the desktop app
via COM). Keeping it pure means it's unit-tested directly (test_live_sheets.py)
and the daily openpyxl report and the live COM report can't drift — both lean on
excel_writer's column list and label helpers.

A tab is a `Sheet`: a grid of `Cell`s plus a freeze pane and AutoFilter extent.
Each Cell carries a value and *named* style intents (fills/fonts as strings);
the renderer maps those names to concrete Excel colors. The names mirror
excel_writer so the look matches the daily report.
"""
from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Optional

from excel_writer import (COLUMNS, QUEUE_HEADERS, MONEY_FMT, _co_label,
                          _drive_run_label, _flags_str, _parse_date,
                          _parse_money, _dwg_suffixes)

# --- named styles (resolved to real colors by live_excel) ------------------- #
F_HEADER = "header"           # blue header bg + white bold
F_SECTION = "section"         # bold section title
F_LINK = "link"               # blue underline hyperlink
F_DRIVE_RUN = "drive_run"     # orange bold (highly-custom)
F_DRIVE_RUN_LINK = "drive_run_link"
F_RED = "red"                 # red bold (a change order landed)

FILL_HEADER = "header"
FILL_OVERDUE = "overdue"      # End Date today/past
FILL_SOON = "soon"            # End Date within 3 days
FILL_NEW = "new"              # new/arrived today (no urgency)
FILL_OVERDUE_NEW = "overdue_new"
FILL_SOON_NEW = "soon_new"
FILL_DWG_YES = "dwg_yes"      # green ✓
FILL_DWG_NO = "dwg_no"        # red (missing)
FILL_SEP = "sep"              # the vertical divider column between the two matrices


@dataclass
class Cell:
    value: Any = ""
    fill: Optional[str] = None
    font: Optional[str] = None
    link: Optional[str] = None
    number_format: Optional[str] = None
    center: bool = False
    comment: Optional[str] = None   # hover note (e.g. the CO# change-order history)


@dataclass
class Sheet:
    name: str
    grid: List[List[Cell]] = field(default_factory=list)
    freeze: Optional[str] = "B2"
    autofilter_a1: Optional[str] = None   # e.g. "A1:AB57"; None = no filter

    # -- builders the sheet functions below use to assemble the grid --
    def row(self, cells: List[Cell]) -> None:
        self.grid.append(cells)

    def blank(self, n: int = 1) -> None:
        for _ in range(n):
            self.grid.append([])

    @property
    def ncols(self) -> int:
        return max((len(r) for r in self.grid), default=0)

    @property
    def nrows(self) -> int:
        return len(self.grid)


# --------------------------------------------------------------------------- #
# Shared helpers                                                               #
# --------------------------------------------------------------------------- #
def added_label(job: Dict[str, Any], ref: Optional[datetime] = None) -> str:
    """The 'Added' label: the AM/PM time if it was added today, the date + time
    if it was added earlier, or 'NO DATA' when we don't have a real add time (an
    order that was already on the board when the watch began, or an older entry)."""
    known = job.get("_added_known")
    if known is None:                                  # live present-dict fallback
        known = not job.get("_carried_over", False)
    if not known:
        return "NO DATA"
    iso = job.get("_added_iso") or job.get("_first_seen") or ""
    try:
        dt = datetime.fromisoformat(iso)
    except (ValueError, TypeError):
        return "NO DATA"
    ref = ref or datetime.now()
    if dt.date() == ref.date():
        return dt.strftime("%#I:%M %p") if _is_windows() else dt.strftime("%-I:%M %p")
    fmt = "%b %#d, %#I:%M %p" if _is_windows() else "%b %-d, %-I:%M %p"
    return dt.strftime(fmt)


def _is_windows() -> bool:
    import os
    return os.name == "nt"


def _header_cells(headers: List[str]) -> List[Cell]:
    return [Cell(h, fill=FILL_HEADER, font=F_HEADER) for h in headers]


def _money_cell(raw: str) -> Cell:
    raw = (raw or "").strip()
    if not raw:
        return Cell("")
    return Cell(_parse_money(raw), number_format=MONEY_FMT)


def _job_value_cells(j: Dict[str, Any], columns: Optional[List] = None,
                     co_changed: bool = False) -> List[Cell]:
    """The cells for one job across `columns` (defaults to the full COLUMNS) —
    mirrors excel_writer._write_job_row (hyperlinks, CO#, drive-run label, money,
    flags), as style-tagged Cells."""
    columns = COLUMNS if columns is None else columns
    cells: List[Cell] = []
    linked_idx = set()
    for idx, (_h, key) in enumerate(columns):
        if key == "job":
            c = Cell(j.get("job", ""))
            so = (j.get("so_pdf") or "").strip()
            if so and j.get("job"):
                c.link, c.font = so, F_LINK
                linked_idx.add(idx)
        elif key == "folder":
            folder = (j.get("job_folder") or "").strip()
            c = Cell((j.get("job_type") or "Open") if folder else "")
            if folder:
                c.link, c.font = folder, F_LINK
                linked_idx.add(idx)
        elif key == "co":
            c = Cell(_co_label(j))
        elif key == "drive_run":
            c = Cell(_drive_run_label(j))
            dr = (j.get("drive_run_pdf") or "").strip()
            if dr:
                c.link, c.font = dr, F_DRIVE_RUN_LINK
                linked_idx.add(idx)
            elif j.get("has_drive_run"):
                c.font = F_DRIVE_RUN
        elif key == "total_price":
            c = _money_cell(j.get("total_price", ""))
        elif key == "flags":
            c = Cell(_flags_str(j))
        else:
            c = Cell(j.get(key, ""))
        cells.append(c)
    # A change order that landed -> non-link cells go red.
    if co_changed:
        for idx, c in enumerate(cells):
            if idx not in linked_idx:
                c.font = F_RED
    return cells


def _row_fill(j: Dict[str, Any], today: date, is_new: bool) -> Optional[str]:
    end = _parse_date(j.get("end_date", ""))
    soon = today + timedelta(days=3)
    if end is not None and end < today:
        return FILL_OVERDUE_NEW if is_new else FILL_OVERDUE
    if end is not None and end <= soon:
        return FILL_SOON_NEW if is_new else FILL_SOON
    return FILL_NEW if is_new else None


def _dwg_header_cells(suffixes: List[str]) -> List[Cell]:
    return [Cell(f"-{s}", fill=FILL_HEADER, font=F_HEADER) for s in suffixes]


def _dwg_row_cells(j: Dict[str, Any], suffixes: List[str]) -> List[Cell]:
    extras = j.get("dwg_extras") or {}
    out = []
    for s in suffixes:
        if s in extras:
            out.append(Cell("✓", fill=FILL_DWG_YES, center=True))
        else:
            out.append(Cell("", fill=FILL_DWG_NO))
    return out


def _a1(rows: int, cols: int) -> str:
    from openpyxl.utils import get_column_letter
    return f"A1:{get_column_letter(max(cols, 1))}{max(rows, 1)}"


# --------------------------------------------------------------------------- #
# Live Queue (the master board)                                               #
# --------------------------------------------------------------------------- #
def full_queue_sheet(
    jobs: List[Dict[str, Any]],
    today: date,
    new_ids: Optional[set] = None,
    co_changed_ids: Optional[set] = None,
    ref: Optional[datetime] = None,
    name: str = "Live Queue",
) -> Sheet:
    """The full board: an 'Added' column, every Full Queue column, then the
    green-✓/red DWG matrix. Urgency + new-order fills, hyperlinks, a totals
    footer, freeze panes, and AutoFilter — same look as the daily Full Queue."""
    new_ids = new_ids or set()
    co_changed_ids = co_changed_ids or set()
    suffixes = _dwg_suffixes(jobs)
    headers = ["Added"] + list(QUEUE_HEADERS) + [f"-{s}" for s in suffixes]

    sh = Sheet(name)
    sh.row([Cell("Added", fill=FILL_HEADER, font=F_HEADER)]
           + _header_cells(list(QUEUE_HEADERS)) + _dwg_header_cells(suffixes))

    total = 0.0
    for j in jobs:
        is_new = j.get("job") in new_ids
        fill = _row_fill(j, today, is_new)
        added = Cell(added_label(j, ref=ref))
        std = _job_value_cells(j, co_changed=j.get("job") in co_changed_ids)
        # Apply the urgency/new row fill across Added + standard columns (not the
        # DWG cells, which keep their own green/red).
        if fill is not None:
            added.fill = fill
            for c in std:
                c.fill = fill
        sh.row([added] + std + _dwg_row_cells(j, suffixes))
        total += _parse_money(j.get("total_price", ""))

    # Totals footer two rows below the data.
    sh.blank(2)
    footer = [Cell("") for _ in headers]
    footer[0] = Cell(f"Total jobs: {len(jobs)}", font=F_SECTION)
    price_col = 1 + QUEUE_HEADERS.index("Total Price")  # +1 for the Added column
    if price_col - 1 >= 0:
        footer[price_col - 1] = Cell("Total $ in process:", font=F_SECTION)
    footer[price_col] = Cell(total, number_format=MONEY_FMT, font=F_SECTION)
    sh.row(footer)

    sh.freeze = "C2"  # keep header row + the Added & Job# columns visible
    sh.autofilter_a1 = _a1(len(jobs) + 1, len(headers))
    return sh


# --------------------------------------------------------------------------- #
# Changes (both baselines, each date-labeled)                                 #
# --------------------------------------------------------------------------- #
def _job_table(sh: Sheet, title: str, jobs: List[Dict[str, Any]],
               extra_headers: Optional[List[str]] = None,
               extra: Optional[Any] = None) -> None:
    """A titled mini-table of full job rows (used by the Changes sections)."""
    sh.row([Cell(f"{title} ({len(jobs)})", font=F_SECTION)])
    if not jobs:
        sh.row([Cell("(none)")])
        sh.blank()
        return
    sh.row(_header_cells(list(QUEUE_HEADERS) + (extra_headers or [])))
    for j in jobs:
        cells = _job_value_cells(j, co_changed=False)
        if extra is not None:
            cells = cells + [Cell(extra(j))]
        sh.row(cells)
    sh.blank()


def fmt_time(iso: str) -> str:
    """An ISO timestamp as an AM/PM time, e.g. '3:53 PM'."""
    try:
        dt = datetime.fromisoformat(iso)
    except (ValueError, TypeError):
        return iso or ""
    return dt.strftime("%#I:%M %p") if _is_windows() else dt.strftime("%-I:%M %p")


_fmt_time = fmt_time   # internal alias


def _events_table(sh: Sheet, title: str, headers: List[str],
                  rows: List[List[Any]]) -> None:
    sh.row([Cell(f"{title} ({len(rows)})", font=F_SECTION)])
    if not rows:
        sh.row([Cell("(none)")])
        sh.blank()
        return
    sh.row(_header_cells(headers))
    for r in rows:
        sh.row([c if isinstance(c, Cell) else Cell(c) for c in r])
    sh.blank()


def changes_sheet(
    new_today: List[Dict[str, Any]],
    change_events: List[Dict[str, Any]],
    removed_today: List[Dict[str, Any]],
    date_str: str,
    name: str = "Changes",
) -> Sheet:
    """Today's activity log:
      - New orders today (new as of today, with their Added time).
      - Change orders today (CO# increases — restored change-order tracking).
      - Orders that changed today: one line per field modification (a field that
        changes several times in a day is several lines), newest first.
      - Removed / completed today.
    `change_events` is the day's change log (see change_log.py)."""
    sh = Sheet(name, freeze=None)
    sh.row([Cell(f"Changes — {date_str}", font=F_SECTION)])
    sh.blank()

    _job_table(sh, "New orders today", new_today,
               extra_headers=["Added"], extra=lambda j: added_label(j))

    newest_first = sorted(change_events, key=lambda e: e.get("time", ""), reverse=True)
    co_events = [e for e in newest_first if e.get("field") == "CO#"]
    _events_table(sh, "Change orders today", ["Time", "Job #", "Customer", "Change"],
                  [[_fmt_time(e.get("time", "")), e.get("job", ""), e.get("customer", ""),
                    f"CO#{e.get('old', '')} -> CO#{e.get('new', '')}"] for e in co_events])

    field_events = [e for e in newest_first if e.get("field") != "CO#"]
    _events_table(sh, "Orders that changed today",
                  ["Time", "Job #", "Customer", "Field", "Old", "New"],
                  [[_fmt_time(e.get("time", "")), e.get("job", ""), e.get("customer", ""),
                    e.get("field", ""), e.get("old", ""), e.get("new", "")] for e in field_events])

    _job_table(sh, "Removed / completed today", removed_today)
    return sh


# --------------------------------------------------------------------------- #
# History                                                                      #
# --------------------------------------------------------------------------- #
def history_sheet(history: Dict[str, Any], name: str = "Order History") -> Sheet:
    """Archived orders (left the queue, not yet returned), newest departure
    first, with the DWG matrix appended like the Full Queue."""
    entries = sorted((history or {}).values(),
                     key=lambda e: e.get("last_seen", ""), reverse=True)
    snaps = [e.get("snapshot", {}) for e in entries]
    suffixes = _dwg_suffixes(snaps)
    headers = list(QUEUE_HEADERS) + ["Last Seen"]

    sh = Sheet(name)
    sh.row(_header_cells(headers) + _dwg_header_cells(suffixes))
    if not entries:
        sh.row([Cell("(no archived orders yet — a job appears here after it drops "
                     "off the queue)")])
        return sh
    for e in entries:
        snap = e.get("snapshot", {})
        sh.row(_job_value_cells(snap, co_changed=False)
               + [Cell(e.get("last_seen", ""))] + _dwg_row_cells(snap, suffixes))
    sh.autofilter_a1 = _a1(len(entries) + 1, len(headers) + len(suffixes))
    return sh


# --------------------------------------------------------------------------- #
# Line Items (one row per order x normalized item; filter to find orders)     #
# --------------------------------------------------------------------------- #
LINE_ITEM_HEADERS = ["Job #", "Customer", "CO#", "Tags", "Item (as printed)",
                     "Normalized", "Details", "Qty", "Price", "Section", "SO PDF"]


# --------------------------------------------------------------------------- #
# Incremental "master log" tabs (Live Queue + Order History)                   #
#                                                                             #
# These are upserted row-by-row (keyed on the order number) instead of being   #
# repainted, so a coworker's filter/sort/scroll survives. Live Queue is the     #
# on-board board (churny fields, full formatting); Order History is the stable  #
# log of every order with its two ✓/red matrices (built below). A record is     #
# (order#, [Cell, ...]); the renderer writes/append/updates the row whose key   #
# matches.                                                                     #
# --------------------------------------------------------------------------- #
LIVE_QUEUE_HEADERS = ["Added"] + list(QUEUE_HEADERS)          # the DWG/feature matrices live on Order History
LIVE_QUEUE_KEY_COL = 2 + QUEUE_HEADERS.index("Job #")          # 1-based col of Job # (Added is col 1)
LIVE_QUEUE_END_DATE_COL = 2 + QUEUE_HEADERS.index("End Date")  # 1-based col of End Date (for the sort)
_END_DATE_IDX = QUEUE_HEADERS.index("End Date")               # its index within the standard cells
_CO_IDX = [i for i, (_, k) in enumerate(COLUMNS) if k == "co"][0]   # CO# cell index


def _co_comment(job: Dict[str, Any]) -> Optional[str]:
    """A hover note for the CO# cell: the order's change-order history, most
    recent first, so hovering shows what the latest change order was."""
    hist = job.get("co_history") or []
    if not hist:
        return None

    def _co_num(line: str) -> int:
        m = re.search(r"C\s*/?\s*O\s*#?\s*(\d+)", line, re.I)
        return int(m.group(1)) if m else 0

    ordered = sorted(hist, key=_co_num, reverse=True)
    return "Change orders (most recent first):\n" + "\n".join(ordered)


_EXCEL_EPOCH = date(1899, 12, 30)   # Excel's day 0 (1900 date system)


def _excel_serial(d: date) -> int:
    """An Excel date serial (days since 1899-12-30). Written as a plain int with a
    date number-format so the cell sorts chronologically and shows as a date —
    avoids passing a Python date through COM (pywin32 won't marshal a bare date)."""
    return (d - _EXCEL_EPOCH).days


def _fmt_dt(iso: Optional[str]) -> str:
    if not iso:
        return ""
    try:
        return datetime.fromisoformat(iso).strftime("%Y-%m-%d %H:%M")
    except (ValueError, TypeError):
        return iso


def row_sig(cells: List[Cell]) -> str:
    """A stable string signature of a row's content+style, so the renderer can
    tell whether a row actually changed and skip writing it if not. Returned as a
    hex digest (not a tuple) so it survives a JSON round-trip in the master store
    — change detection has to work across watcher restarts."""
    payload = [[c.value, c.fill, c.font, c.link, c.number_format, c.center] for c in cells]
    return hashlib.md5(json.dumps(payload, default=str, sort_keys=True).encode()).hexdigest()


def live_queue_records(jobs: List[Dict[str, Any]], today: date,
                       new_ids: Optional[set] = None,
                       ref: Optional[datetime] = None) -> List:
    """(order#, cells) per on-board order: Added + every Full Queue column, with
    urgency / new-today row fills and hyperlinks. `new_ids` is the set of order
    numbers that are new today (not in the previous snapshot)."""
    new_ids = new_ids or set()
    out = []
    for j in jobs:
        # number_format "@" (Text) keeps the AM/PM label (e.g. "3:53 PM") from
        # being coerced by Excel into a 24h datetime serial.
        added = Cell(added_label(j, ref=ref), number_format="@")
        std = _job_value_cells(j, co_changed=False)
        # Write End Date as a real date so Live Queue can sort by it (overdue
        # first -> red at the top); blanks sort to the bottom.
        ed = _parse_date(j.get("end_date", ""))
        if ed is not None:
            std[_END_DATE_IDX].value = _excel_serial(ed)
            std[_END_DATE_IDX].number_format = "mm/dd/yyyy"
        # Hover the CO# cell to see the change-order history (most recent first).
        std[_CO_IDX].comment = _co_comment(j)
        fill = _row_fill(j, today, is_new=str(j.get("job") or "") in new_ids)
        cells = [added] + std
        if fill:
            for c in cells:
                if c.fill is None:
                    c.fill = fill
        out.append((str(j.get("job") or ""), cells))
    return out


# Order History is a stable *log*: only the order's identity + SO spec + the two
# matrices + the presence flags — NOT the churny board fields (dates, price,
# assignee, status), which live on Live Queue. Keeping only stable columns means
# a row's signature changes solely when the order is added or its On Queue/Left
# flags flip, so the 12K-row log isn't rewritten on every field tick. Layout:
# the data columns first (Job # leads, so it's the pinned column), then the
# On Queue / Added / Left flags right before the two ✓/red matrices.
OH_DATA_COLUMNS = [
    ("Job #", "job"), ("Folder", "folder"), ("Quote Run", "drive_run"), ("CO#", "co"),
    ("Design", "design"), ("Description", "so_design_desc"), ("Size", "so_size"),
    ("Arrangement", "so_arrangement"), ("Motor Pos", "so_motor_pos"), ("Class", "so_class"),
    ("Rotation", "so_rotation"), ("Discharge", "so_discharge"), ("% Width", "so_pct_width"),
    ("Wheel Type", "so_wheel_type"), ("Design Temp", "so_design_temp"),
    ("Max Temp", "so_max_temp"), ("Special Temp", "so_special_temp"),
    ("Customer", "customer"), ("Primary Rep", "primary_rep"), ("Item", "item"),
]
OH_FLAG_HEADERS = ["On Queue", "Added", "Left"]
OH_SEP_HEADER = "│"
ORDER_HISTORY_KEY_COL = 1   # Job # is the first column (pinned)


def _order_tags(j: Dict[str, Any]) -> set:
    """The canonical feature tags on an order, from its captured line items."""
    out = set()
    for it in j.get("line_items") or []:
        for t in it.get("tags") or []:
            out.add(t)
    return out


def order_history_build(orders: List, today: date) -> Dict[str, Any]:
    """Build the Order History tab: one row per order with the stable data
    columns (Job # first), then the On Queue / Added / Left flags, then the
    AutoCAD **DWG matrix** (green ✓ / red), a vertical divider, and the line-item
    **Feature matrix** (green ✓ / red). The matrix columns are the union of DWG
    suffixes / feature tags across `orders`, so they grow only when a brand-new
    suffix or tag appears.

    `orders` is a list of (job#, entry) with entry = {on_queue, added, left, job}.
    Returns a dict: headers, records [(key, cells)], and the 1-based column ranges
    of each matrix + the separator (the renderer colors the matrices via
    conditional formatting and draws the divider)."""
    jobs = [e.get("job", {}) for _, e in orders]
    suffixes = _dwg_suffixes(jobs)
    tag_count: Dict[str, int] = {}
    for j in jobs:
        for t in _order_tags(j):
            tag_count[t] = tag_count.get(t, 0) + 1
    tags = sorted(tag_count, key=lambda t: (-tag_count[t], t))   # most common first

    headers = ([h for h, _ in OH_DATA_COLUMNS] + OH_FLAG_HEADERS
               + [f"-{s}" for s in suffixes] + [OH_SEP_HEADER] + list(tags))
    n_data_flags = len(OH_DATA_COLUMNS) + len(OH_FLAG_HEADERS)
    dwg_range = (n_data_flags + 1, n_data_flags + len(suffixes))
    sep_col = n_data_flags + len(suffixes) + 1
    feat_range = (sep_col + 1, sep_col + len(tags))

    records = []
    for jn, e in orders:
        j = e.get("job", {})
        onq = bool(e.get("on_queue"))
        data = _job_value_cells(j, columns=OH_DATA_COLUMNS)
        flags = [Cell("YES" if onq else "NO"),
                 Cell(_fmt_dt(e.get("added"))), Cell(_fmt_dt(e.get("left")))]
        de = j.get("dwg_extras") or {}
        # Matrix cells carry only the ✓ / blank value; the renderer colors them
        # via conditional formatting (cheap enough for ~12K rows).
        dwg = [Cell("✓" if s in de else "", center=True) for s in suffixes]
        sep = [Cell("", fill=FILL_SEP)]
        otags = _order_tags(j)
        feat = [Cell("✓" if t in otags else "", center=True) for t in tags]
        records.append((str(jn), data + flags + dwg + sep + feat))

    return {"headers": headers, "records": records,
            "dwg_range": dwg_range, "sep_col": sep_col, "feat_range": feat_range}


def plan_upsert(desired: List, existing_sigs: Dict[str, tuple],
                allow_delete: bool = False) -> List:
    """Diff desired rows against what's already on the sheet (by key) and return
    the minimal op list: ('append', key, cells) for keys not present,
    ('update', key, cells) for keys whose signature changed, and (with
    allow_delete) ('delete', key, None) for keys no longer desired. Unchanged
    rows produce no op — so they're never rewritten and filters/scroll persist.

    `desired` is a list of (key, sig, cells); `existing_sigs` is {key: sig}.
    """
    ops, seen = [], set()
    for key, sig, cells in desired:
        if not key:
            continue
        seen.add(key)
        prev = existing_sigs.get(key)
        if prev is None:
            ops.append(("append", key, cells))
        elif prev != sig:
            ops.append(("update", key, cells))
    if allow_delete:
        for key in existing_sigs:
            if key not in seen:
                ops.append(("delete", key, None))
    return ops


def line_items_sheet(
    store: Dict[str, Any],
    order_nums: Optional[List[str]] = None,
    name: str = "Line Items",
) -> Sheet:
    """One row per (order, normalized line item) so you can AutoFilter the
    'Normalized' column by an item name and have the matching orders populate.
    Covers every stored order by default (the whole backlog, so the search spans
    all history); pass `order_nums` to restrict it (e.g. just the board). Mirrors
    find_orders.write_xlsx's Line Items sheet."""
    jobs_store = (store or {}).get("jobs", {})
    keys = [k for k in (order_nums if order_nums is not None else jobs_store)
            if k in jobs_store]

    sh = Sheet(name)
    sh.row(_header_cells(LINE_ITEM_HEADERS))
    n_rows = 0
    for jn in keys:
        rec = jobs_store.get(jn) or {}
        co = f"CO#{rec['co_number']}" if rec.get("co_number") else ""
        for it in rec.get("items") or []:
            so = rec.get("so_pdf") or ""
            link_cell = Cell("Open", link=so, font=F_LINK) if so else Cell("")
            sh.row([
                Cell(jn), Cell(rec.get("customer", "")), Cell(co),
                Cell(", ".join(it.get("tags") or [])),
                Cell(it.get("raw", "")), Cell(it.get("norm", "")),
                Cell(" ; ".join(it.get("details") or [])),
                Cell(it.get("qty", "")), Cell(it.get("price", "")),
                Cell(it.get("section", "")), link_cell,
            ])
            n_rows += 1
    if n_rows == 0:
        sh.row([Cell("(no line items captured yet for the current orders)")])
    else:
        sh.autofilter_a1 = _a1(n_rows + 1, len(LINE_ITEM_HEADERS))
    return sh
