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


@dataclass
class Cell:
    value: Any = ""
    fill: Optional[str] = None
    font: Optional[str] = None
    link: Optional[str] = None
    number_format: Optional[str] = None
    center: bool = False


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
    """Human 'time it was added'. Carried-over orders (already in the queue when
    the watch began) show a marker rather than a fake precise time."""
    if job.get("_carried_over"):
        return "before watch"
    iso = job.get("_first_seen") or ""
    try:
        dt = datetime.fromisoformat(iso)
    except ValueError:
        return iso
    ref = ref or datetime.now()
    fmt_t = "%#I:%M %p" if _is_windows() else "%-I:%M %p"
    if dt.date() == ref.date():
        return dt.strftime(fmt_t)
    fmt_d = "%b %#d %#I:%M %p" if _is_windows() else "%b %-d %-I:%M %p"
    return dt.strftime(fmt_d)


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


def _job_value_cells(j: Dict[str, Any], co_changed: bool) -> List[Cell]:
    """The standard COLUMNS cells for one job — mirrors excel_writer._write_job_row
    (hyperlinks, CO#, drive-run label, money, flags), as style-tagged Cells."""
    cells: List[Cell] = []
    linked_idx = set()
    for idx, (_h, key) in enumerate(COLUMNS):
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


def _changed_table(sh: Sheet, title: str, changed: List[Dict[str, Any]]) -> None:
    sh.row([Cell(f"{title} ({len(changed)})", font=F_SECTION)])
    if not changed:
        sh.row([Cell("(none)")])
        sh.blank()
        return
    sh.row(_header_cells(["Job #", "Customer", "Field", "Old value", "New value"]))
    for ch in changed:
        for (fieldname, old, new) in ch.get("changes", []):
            sh.row([Cell(ch.get("job", "")), Cell(ch.get("customer", "")),
                    Cell(fieldname), Cell(old), Cell(new)])
    sh.blank()


def _group(sh: Sheet, heading: str, diff: Dict[str, Any]) -> None:
    sh.row([Cell(heading, font=F_SECTION)])
    sh.blank()
    _job_table(sh, "New orders", diff.get("new", []),
               extra_headers=["Added"], extra=lambda j: added_label(j))
    _job_table(sh, "Returning orders", diff.get("returning", []),
               extra_headers=["Last seen"], extra=lambda j: j.get("_last_seen", ""))
    _job_table(sh, "Removed / completed", diff.get("removed", []))
    _changed_table(sh, "Orders that changed", diff.get("changed", []))


def changes_sheet(
    intraday: Dict[str, Any],
    intraday_label: str,
    yesterday: Dict[str, Any],
    yesterday_label: str,
    name: str = "Changes",
) -> Sheet:
    """Two stacked, date-labeled groups: changes since this morning's baseline,
    and changes vs yesterday's run. Each group lists new / returning / removed /
    changed orders."""
    sh = Sheet(name, freeze=None)
    _group(sh, f"Changes since this morning — baseline {intraday_label}", intraday)
    sh.blank()
    _group(sh, f"Changes vs yesterday — {yesterday_label}", yesterday)
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
# repainted, so a coworker's filter/sort/scroll survives. That needs a STABLE  #
# column schema, so the variable-width DWG matrix is collapsed to one compact  #
# "Custom DWGs" text column here (the full green/red matrix stays in the daily #
# report). A record is (order#, [Cell, ...]); the renderer writes/append/      #
# updates the row whose key matches.                                          #
# --------------------------------------------------------------------------- #
LIVE_QUEUE_HEADERS = ["Added"] + list(QUEUE_HEADERS)          # Custom DWGs lives on Order History only
LIVE_QUEUE_KEY_COL = 2 + QUEUE_HEADERS.index("Job #")          # 1-based col of Job # (Added is col 1)
ORDER_HISTORY_HEADERS = ["On Queue", "Added", "Left"] + list(QUEUE_HEADERS) + ["Custom DWGs"]
ORDER_HISTORY_KEY_COL = 4 + QUEUE_HEADERS.index("Job #")        # 1-based col of Job # (3 lead cols)


def _dwg_text(j: Dict[str, Any]) -> str:
    """The job's custom-DWG suffixes as one compact cell, e.g. '-35, -51'."""
    extras = j.get("dwg_extras") or {}
    if not extras:
        return ""
    order = sorted(extras, key=lambda s: (int(s), s) if s.isdigit() else (10**9, s))
    return ", ".join(f"-{s}" for s in order)


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
        added = Cell(added_label(j, ref=ref))
        std = _job_value_cells(j, co_changed=False)
        fill = _row_fill(j, today, is_new=str(j.get("job") or "") in new_ids)
        cells = [added] + std
        if fill:
            for c in cells:
                if c.fill is None:
                    c.fill = fill
        out.append((str(j.get("job") or ""), cells))
    return out


def order_history_records(orders: List, today: date) -> List:
    """(order#, cells) per logged order, from live_master.ordered() entries:
    On Queue / Added / Left, then every Full Queue column + Custom DWGs. On-board
    rows get the urgency fill; off-board rows stay plain."""
    out = []
    for jn, entry in orders:
        j = entry.get("job", {})
        onq = bool(entry.get("on_queue"))
        lead = [Cell("YES" if onq else "NO"),
                Cell(_fmt_dt(entry.get("added"))), Cell(_fmt_dt(entry.get("left")))]
        std = _job_value_cells(j, co_changed=False)
        cells = lead + std + [Cell(_dwg_text(j))]
        if onq:
            fill = _row_fill(j, today, is_new=False)
            if fill:
                for c in cells:
                    if c.fill is None:
                        c.fill = fill
        out.append((str(jn), cells))
    return out


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
