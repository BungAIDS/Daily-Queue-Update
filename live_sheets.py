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

import engineers
from excel_writer import (COLUMNS, QUEUE_HEADERS, MONEY_FMT, _co_label,
                          _drive_run_label, _flags_str, _parse_date,
                          _parse_money, _dwg_suffixes, folder_of, split_arrangement,
                          split_size)

# --- named styles (resolved to real colors by live_excel) ------------------- #
F_HEADER = "header"           # blue header bg + white bold
F_SECTION = "section"         # bold section title
F_LINK = "link"               # blue underline hyperlink
F_DRIVE_RUN = "drive_run"     # orange bold (highly-custom)
F_DRIVE_RUN_LINK = "drive_run_link"
F_RED = "red"                 # red bold (a change order landed)
F_NOTE = "note"               # muted gray (e.g. the Changes 'last updated' stamp)

FILL_HEADER = "header"
FILL_OVERDUE = "overdue"      # End Date in the past (overdue / red)
FILL_DUETODAY = "duetoday"    # End Date today -> overdue (red) tomorrow (orange)
FILL_SOON = "soon"            # End Date within 3 days (gold)
FILL_NEW = "new"              # new/arrived today (no urgency)
FILL_OVERDUE_NEW = "overdue_new"
FILL_DUETODAY_NEW = "duetoday_new"
FILL_SOON_NEW = "soon_new"
FILL_DWG_YES = "dwg_yes"      # green ✓
FILL_DWG_NO = "dwg_no"        # red (missing)
FILL_SEP = "sep"              # the vertical divider column between the two matrices
# Grey shading on the 'after' row of a change instance, getting darker for each
# later instance the same order picks up through the day.
FILL_CHANGE1 = "chg1"
FILL_CHANGE2 = "chg2"
FILL_CHANGE3 = "chg3"
FILL_CHANGE4 = "chg4"
_CHANGE_FILLS = [FILL_CHANGE1, FILL_CHANGE2, FILL_CHANGE3, FILL_CHANGE4]


@dataclass
class Cell:
    value: Any = ""
    fill: Optional[str] = None
    font: Optional[str] = None
    link: Optional[str] = None
    number_format: Optional[str] = None
    center: bool = False
    comment: Optional[str] = None   # hover note (e.g. the CO# change-order history)
    volatile: bool = False          # value changes every cycle (e.g. the 'Last updated'
                                    # stamp): excluded from the render fingerprint so it
                                    # doesn't force a full repaint, refreshed in place


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
    """The 'Added' label: the AM/PM time if it was added today, just the date if
    it was added earlier, or 'NO DATA' only when there's no add timestamp at all."""
    iso = job.get("_added_iso") or job.get("_first_seen") or ""
    try:
        dt = datetime.fromisoformat(iso)
    except (ValueError, TypeError):
        return "NO DATA"
    ref = ref or datetime.now()
    if dt.date() == ref.date():
        return dt.strftime("%#I:%M %p") if _is_windows() else dt.strftime("%-I:%M %p")
    return dt.strftime("%b %#d, %Y") if _is_windows() else dt.strftime("%b %-d, %Y")


def last_out_label(job: Dict[str, Any], ref: Optional[datetime] = None) -> str:
    """The 'Last Out' label: when the order was MOST RECENTLY removed before its
    current stint (the entry's last_out, injected as `_last_out`) — a time if that
    was today, else a short date. Blank if the order has never left the board."""
    iso = job.get("_last_out") or ""
    try:
        dt = datetime.fromisoformat(iso)
    except (ValueError, TypeError):
        return ""
    ref = ref or datetime.now()
    if dt.date() == ref.date():
        return dt.strftime("%#I:%M %p") if _is_windows() else dt.strftime("%-I:%M %p")
    return dt.strftime("%b %#d, %Y") if _is_windows() else dt.strftime("%b %-d, %Y")


def added_date(job: Dict[str, Any]) -> Optional[date]:
    """The date an order most recently came onto the board (from _added_iso /
    _first_seen) — i.e. the date the 'Added' column reflects — or None if unknown."""
    iso = job.get("_added_iso") or job.get("_first_seen") or ""
    try:
        return datetime.fromisoformat(iso).date()
    except (ValueError, TypeError):
        return None


def prev_business_day(d: date) -> date:
    """The most recent business day strictly before `d` (weekends skipped): Mon ->
    Fri, Sun -> Fri, otherwise the day before. Holidays are not accounted for."""
    wd = d.weekday()                  # Mon=0 .. Sun=6
    if wd == 0:
        return d - timedelta(days=3)
    if wd == 6:
        return d - timedelta(days=2)
    return d - timedelta(days=1)


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
                     co_changed: bool = False, arrange_comment: bool = False) -> List[Cell]:
    """The cells for one job across `columns` (defaults to the full COLUMNS) —
    mirrors excel_writer._write_job_row (hyperlinks, CO#, drive-run label, money,
    flags), as style-tagged Cells. `arrange_comment` trims the Arrangement cell to
    its short 'A/X' code and moves any descriptive suffix to a hover note (used on
    the display tabs; Order History keeps the full text to avoid re-writing the
    whole log)."""
    columns = COLUMNS if columns is None else columns
    cells: List[Cell] = []
    linked_idx = set()
    for idx, (_h, key) in enumerate(columns):
        if key == "job":
            c = Cell(j.get("job", ""))
            so = (j.get("so_pdf") or "").strip()
            if so and j.get("job"):
                # Links straight to the latest Sales Order PDF. The watcher keeps
                # so_pdf pointed at the newest revision (sales_orders
                # .refresh_sales_orders / a re-fetch on a change order), so a CO
                # renaming the file (… (original).pdf -> … CO#1.pdf) follows here.
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
                # With more than one run, link to the folder that holds the
                # downloaded runs, not just one of the files.
                c.link = folder_of(dr) if (j.get("drive_run_count") or 0) > 1 else dr
                c.font = F_DRIVE_RUN_LINK
                linked_idx.add(idx)
            elif j.get("has_drive_run"):
                c.font = F_DRIVE_RUN
        elif key == "total_price":
            c = _money_cell(j.get("total_price", ""))
        elif key == "so_arrangement" and arrange_comment:
            # Keep the column to 'A/X'; hover the cell for the descriptive suffix.
            code, note = split_arrangement(j.get("so_arrangement", ""))
            c = Cell(code)
            if note:
                c.comment = note
        elif key == "so_size" and arrange_comment:
            # Keep the column to the main size; hover for any trailing detail.
            main, note = split_size(j.get("so_size", ""))
            c = Cell(main)
            if note:
                c.comment = note
        elif key == "flags":
            c = Cell(_flags_str(j))
        elif key == "engineers":
            c = Cell(engineers.cell_text(j))
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
        return FILL_OVERDUE_NEW if is_new else FILL_OVERDUE        # past due -> red
    if end is not None and end == today:                          # due today -> red tomorrow
        return FILL_DUETODAY_NEW if is_new else FILL_DUETODAY      #             -> orange
    if end is not None and end <= soon:
        return FILL_SOON_NEW if is_new else FILL_SOON             # within 3 days -> gold
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
        std = _job_value_cells(j, co_changed=j.get("job") in co_changed_ids, arrange_comment=True)
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


def fmt_datetime(when: "datetime | str") -> str:
    """A date + AM/PM time, e.g. 'Jun 18, 2026 3:53 PM' — for the Changes tab's
    'last updated' stamp. Accepts a datetime or an ISO string."""
    if isinstance(when, str):
        try:
            when = datetime.fromisoformat(when)
        except (ValueError, TypeError):
            return when or ""
    fmt = "%b %#d, %Y %#I:%M %p" if _is_windows() else "%b %-d, %Y %-I:%M %p"
    return when.strftime(fmt)


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


def _orders_changed_table(sh: Sheet, field_events: List[Dict[str, Any]]) -> None:
    """'Orders that changed today' as a running history: a white 'was' row with
    each changed field's start-of-day value, then ONE row per change instance (a
    poll in which the order changed) showing only the fields that moved in that
    instance — the prior row's value carries forward as the implied 'old', so a
    value is never repeated. Instance rows shade progressively darker; white marks
    a new order. Changed cells are red. Columns are the union of fields any order
    changed today, ordered like the queue."""
    by_job: Dict[str, Dict[str, Any]] = {}
    for e in sorted(field_events, key=lambda x: x.get("time", "")):   # oldest first
        jn, label = str(e.get("job", "") or ""), e.get("field", "") or ""
        old, new = e.get("old", ""), e.get("new", "")
        if not jn or not label or label == "Customer" or str(old) == str(new):
            continue
        rec = by_job.setdefault(jn, {"customer": "", "instances": {}})
        rec["customer"] = e.get("customer", "") or rec["customer"]
        # An instance = one poll (same timestamp); collect every field that moved
        # in it as {field: (old, new)}.
        rec["instances"].setdefault(e.get("time", ""), {})[label] = (old, new)

    orders: Dict[str, Dict[str, Any]] = {}
    for jn, rec in by_job.items():
        instances = [rec["instances"][t] for t in sorted(rec["instances"]) if rec["instances"][t]]
        if not instances:
            continue
        first_old: Dict[str, Any] = {}
        for moved in instances:                       # oldest first -> first old per field
            for label, (old, _new) in moved.items():
                first_old.setdefault(label, old)
        orders[jn] = {"customer": rec["customer"], "time": max(rec["instances"]),
                      "fields": set(first_old),
                      # each instance as {field: new value}
                      "steps": [{l: n for l, (_o, n) in moved.items()} for moved in instances],
                      "baseline": first_old}

    sh.row([Cell(f"Orders that changed today ({len(orders)})", font=F_SECTION)])
    if not orders:
        sh.row([Cell("(none)")])
        sh.blank()
        return
    changed: set = set()
    for o in orders.values():
        changed |= o["fields"]
    order_idx = {h: i for i, h in enumerate(QUEUE_HEADERS)}
    cols = sorted(changed, key=lambda l: (order_idx.get(l, 10 ** 6), l))
    sh.row(_header_cells(["Job #", "Customer"] + cols))
    for jn in sorted(orders, key=lambda j: orders[j]["time"], reverse=True):  # most recent first
        o = orders[jn]
        # 'was' row (white): start-of-day value of each changed field; the order's
        # only white row, carrying Job #/Customer so white reads as a new order.
        was = [Cell(jn), Cell(o["customer"])]
        for label in cols:
            was.append(Cell(o["baseline"][label], font=F_RED) if label in o["fields"] else Cell(""))
        sh.row(was)
        # one row per instance, progressively darker, showing only what it changed.
        for i, step in enumerate(o["steps"]):
            fill = _CHANGE_FILLS[min(i, len(_CHANGE_FILLS) - 1)]
            row = [Cell("", fill=fill), Cell("", fill=fill)]
            for label in cols:
                row.append(Cell(step[label], fill=fill, font=F_RED)
                           if label in step else Cell("", fill=fill))
            sh.row(row)
    sh.blank()


def _co_change_desc(order: Dict[str, Any], co_num: Any) -> str:
    """The change-order description for CO#<co_num> from the order's co_history
    (the part after the 'C/O #N date initials:' prefix), or '' if not found."""
    cn = str(co_num or "").strip()
    for line in (order.get("co_history") or []):
        m = re.match(r"\s*C\s*/?\s*O\s*#?\s*(\d+)", str(line), re.I)
        if m and m.group(1) == cn:
            return line.split(":", 1)[1].strip() if ":" in line else str(line).strip()
    return ""


def changes_sheet(
    new_today: List[Dict[str, Any]],
    change_events: List[Dict[str, Any]],
    removed_today: List[Dict[str, Any]],
    date_str: str,
    updated_at: Optional[str] = None,
    order_lookup: Optional[Dict[str, Dict[str, Any]]] = None,
    name: str = "Changes",
) -> Sheet:
    """Today's activity log:
      - New orders today (new as of today, with their Added time).
      - Change orders today (CO# increases — restored change-order tracking).
      - Orders that changed today: one line per field modification (a field that
        changes several times in a day is several lines), newest first.
      - Removed / completed today.
    `change_events` is the day's change log (see change_log.py); `updated_at` is a
    display string stamped at the top so users can see the tab is live;
    `order_lookup` maps job# -> the order's job dict (for the Design / Arrangement
    / change description columns on the change-order table)."""
    order_lookup = order_lookup or {}
    sh = Sheet(name, freeze=None)
    sh.row([Cell(f"Changes — {date_str}", font=F_SECTION)])
    if updated_at:
        sh.row([Cell(f"Last updated {updated_at}", font=F_NOTE, volatile=True)])
    sh.blank()

    _job_table(sh, "New orders today", new_today,
               extra_headers=["Added"], extra=lambda j: added_label(j))

    newest_first = sorted(change_events, key=lambda e: e.get("time", ""), reverse=True)
    co_events = [e for e in newest_first if e.get("field") == "CO#"]

    def _co_row(e: Dict[str, Any]) -> List[Any]:
        o = order_lookup.get(str(e.get("job", "")), {})
        return [_fmt_time(e.get("time", "")), e.get("job", ""), e.get("customer", ""),
                o.get("design", ""), split_arrangement(o.get("so_arrangement", ""))[0],
                f"CO#{e.get('old', '')} -> CO#{e.get('new', '')}",
                _co_change_desc(o, e.get("new"))]
    _events_table(sh, "Change orders today",
                  ["Time", "Job #", "Customer", "Design", "Arrangement", "Change", "What changed"],
                  [_co_row(e) for e in co_events])

    field_events = [e for e in newest_first if e.get("field") != "CO#"]
    _orders_changed_table(sh, field_events)

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
# "Added" (last time it came onto the board), then the Full Queue columns, then a
# trailing "#" column carrying the cbcinsider board position. The tab is sorted by
# "#" so the order matches the queue on the site (re-sort "#" to restore it).
# "Last Out" (most recent prior departure) sits just before the trailing "#"
# board-position column, so adding it doesn't shift Job #/End Date (still after the
# single leading "Added").
LIVE_QUEUE_HEADERS = ["Added"] + list(QUEUE_HEADERS) + ["Last Out", "#"]
LIVE_QUEUE_CBC_COL = len(LIVE_QUEUE_HEADERS)                   # the trailing "#" board-position col (sort key)
LIVE_QUEUE_LAST_OUT_COL = len(LIVE_QUEUE_HEADERS) - 1         # the "Last Out" col (AM/PM-or-date text)
LIVE_QUEUE_KEY_COL = 2 + QUEUE_HEADERS.index("Job #")          # 1-based col of Job # (Added is col 1)
LIVE_QUEUE_END_DATE_COL = 2 + QUEUE_HEADERS.index("End Date")  # 1-based col of End Date
# The 'Removed since this morning' block lines its data columns up under the board
# above, but LEADS with the removal time (the "Removed" column) instead of "Added"
# — the most relevant fact for an order that just left — and drops the trailing "#".
LIVE_QUEUE_REMOVED_HEADERS = ["Removed"] + list(QUEUE_HEADERS)
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


def _board_row_cells(j: Dict[str, Any], today: date, co_changed: bool, is_new: bool,
                     ref: Optional[datetime], leading: Cell, trailing: List[Cell]) -> List[Cell]:
    """`leading` cell + every Full Queue column + `trailing` cells, styled exactly
    like a Live Queue row: real-date End Date, CO# hover note, CO#-red text, and the
    urgency / new-today row fill. `leading` is the Added time on the board or the
    removal time in the Removed block; `trailing` is [Last Out, '#'] on the board,
    [] in the Removed block — so both render with the same data columns aligned."""
    std = _job_value_cells(j, co_changed=co_changed, arrange_comment=True)
    # Write End Date as a real date so it sorts/filters as a date in Excel.
    ed = _parse_date(j.get("end_date", ""))
    if ed is not None:
        std[_END_DATE_IDX].value = _excel_serial(ed)
        std[_END_DATE_IDX].number_format = "mm/dd/yyyy"
    # Hover the CO# cell to see the change-order history (most recent first).
    std[_CO_IDX].comment = _co_comment(j)
    if co_changed:                       # carry the red over to the non-std cells
        leading.font = F_RED
        for t in trailing:
            t.font = F_RED
    cells = [leading] + std + list(trailing)
    fill = _row_fill(j, today, is_new=is_new)
    if fill:
        for c in cells:
            if c.fill is None:
                c.fill = fill
    return cells


def removed_block(removed: List, today: date, new_ids: Optional[set] = None,
                  co_changed_ids: Optional[set] = None, ref: Optional[datetime] = None,
                  title: str = "Removed from the queue since this morning") -> Dict[str, Any]:
    """The Live Queue's 'Removed since this morning' section as a styled block —
    each removed order drawn exactly like its Live Queue row (same columns, the
    urgency/new-today fill and CO#-red text it had on the board), with the trailing
    board-position '#' slot replaced by the time it left. `removed` is a list of
    (job_dict, left_iso). Returns a payload the COM renderer draws below the board."""
    new_ids = new_ids or set()
    co_changed_ids = co_changed_ids or set()
    rows = []
    for j, left in removed:
        jn = str(j.get("job") or "")
        rem = Cell(fmt_time(left), number_format="@")   # when it left -> the LEADING column
        rows.append(_board_row_cells(j, today, jn in co_changed_ids, jn in new_ids, ref,
                                     leading=rem, trailing=[]))
    return {"title": title,
            "header_cells": _header_cells(LIVE_QUEUE_REMOVED_HEADERS),
            "rows": rows}


def live_queue_records(jobs: List[Dict[str, Any]], today: date,
                       new_ids: Optional[set] = None,
                       co_changed_ids: Optional[set] = None,
                       ref: Optional[datetime] = None) -> List:
    """(order#, cells) per on-board order: Added + every Full Queue column + the
    "#" board position, with urgency / new-today row fills and hyperlinks.
    `new_ids` is the set of order numbers new today (not in the previous
    snapshot); `co_changed_ids` is the set that had a change order (CO#) land
    today — their text goes red."""
    new_ids = new_ids or set()
    co_changed_ids = co_changed_ids or set()
    out = []
    for j in jobs:
        jn = str(j.get("job") or "")
        added = Cell(added_label(j, ref=ref), number_format="@")
        last_out = Cell(last_out_label(j, ref=ref), number_format="@")   # most recent prior departure
        # "#" = the cbcinsider board position (the sort key for board order).
        pos = j.get("_cbc_pos")
        cbc = Cell(pos if isinstance(pos, int) else "", center=True)
        cells = _board_row_cells(j, today, jn in co_changed_ids, jn in new_ids, ref,
                                 leading=added, trailing=[last_out, cbc])
        out.append((jn, cells))
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
    ("Customer", "customer"), ("Engineer", "engineers"),
    ("Primary Rep", "primary_rep"), ("Item", "item"),
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


def _suffix_key(s: str):
    return (int(s), s) if str(s).isdigit() else (10 ** 9, s)


def order_history_build(orders: List, today: date,
                        prev_columns: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Build the Order History tab: one row per order with the stable data
    columns (Job # first), then the On Queue / Added / Left flags, then the
    AutoCAD **DWG matrix** (green ✓ / red), a vertical divider, and the line-item
    **Feature matrix** (green ✓ / red).

    Column order is STABLE so the tab isn't rebuilt in normal operation: it keeps
    `prev_columns`' order and only APPENDS brand-new suffixes/tags at the end (it
    never drops or reorders). `spec["columns"]` is the order used — the caller
    persists it; the headers change only when a new suffix/tag appears, which is
    the only time the tab is rebuilt.

    `orders` is a list of (job#, entry) with entry = {on_queue, added, left, job}.
    Returns headers, records [(key, cells)], the matrix/separator column ranges,
    and the column order."""
    jobs = [e.get("job", {}) for _, e in orders]
    cur_suffixes = set(_dwg_suffixes(jobs))
    cur_tags: set = set()
    for j in jobs:
        cur_tags |= _order_tags(j)

    prev = prev_columns or {}
    prev_suffixes = list(prev.get("suffixes") or [])
    prev_tags = list(prev.get("tags") or [])
    suffixes = prev_suffixes + sorted(cur_suffixes - set(prev_suffixes), key=_suffix_key)
    tags = prev_tags + sorted(cur_tags - set(prev_tags))
    columns = {"suffixes": suffixes, "tags": tags}

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

    return {"headers": headers, "records": records, "columns": columns,
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
