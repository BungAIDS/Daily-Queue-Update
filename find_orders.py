"""Find orders by what's on their Sales Order — search the line-items store.

The store (fed by the daily run, backfill_orders.py, and line_items_scan.py)
holds every captured Sales Order line item in three forms: the verbatim text,
a normalized form (abbreviations expanded, qty/prices stripped), and canonical
feature tags. Search hits all three, so "stainless" finds the order whose SO
says "SS SHAFT SLEEVE".

    python find_orders.py shaft seal           # orders matching BOTH terms
    python find_orders.py --any teflon viton   # ...ANY of the terms
    python find_orders.py --tag "SHAFT SEAL"   # by canonical tag
    python find_orders.py cermic felt --fuzzy  # typo-tolerant (ratio 0.84;
                                               # give --fuzzy AFTER the terms,
                                               # or as --fuzzy=0.8)
    python find_orders.py --job 421314         # one job's stored items
    python find_orders.py --list-tags          # live tag vocabulary + counts
    python find_orders.py shaft seal --xlsx    # ...write matches to Excel too
    python find_orders.py --xlsx               # full inventory workbook, one
                                               # row per item, AutoFilter on

Terms are case-insensitive substrings (multi-word terms must appear as a
phrase in the normalized text; separate words are separate ANDed terms)."""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

from config import BACKLOG_DIR
import line_items as li

log = logging.getLogger("find-orders")

XLSX_DEFAULT = BACKLOG_DIR / "line_items.xlsx"


def _store_stats(store: Dict[str, Any]) -> str:
    n_items = sum(len(r.get("items") or []) for r in store["jobs"].values())
    return f"{len(store['jobs'])} order(s), {n_items} item(s) in {li.store_path()}"


def _print_hits(hits: List[Dict[str, Any]], terms: List[str] | None = None,
                all_details: bool = False) -> None:
    """Print matched orders. Detail lines (vendor, motor specs, ...) are shown
    when they're what the search term actually hit — or all of them with
    all_details (the --job view)."""
    terms_u = [t.upper() for t in (terms or [])]
    for h in hits:
        co = f"  CO#{h['co_number']}" if h.get("co_number") else ""
        cust = f"  {h['customer']}" if h.get("customer") else ""
        print(f"\n{h['job']}{cust}{co}   (scanned {h.get('scanned_at', '')[:10]})")
        if h.get("so_pdf"):
            print(f"    SO: {h['so_pdf']}")
        for it in h["matches"]:
            tags = ", ".join(it.get("tags") or []) or "-"
            print(f"    [{tags}]  {it['raw']}")
            for d in it.get("details") or []:
                if all_details or any(t in d.upper() for t in terms_u):
                    print(f"        · {d}")


def write_xlsx(hits: List[Dict[str, Any]], path: Path) -> Path:
    """One row per (job, item) with AutoFilter — the filter-in-Excel view."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", fgColor="305496")
    header_font = Font(color="FFFFFF", bold=True)
    link_font = Font(color="0563C1", underline="single")

    headers = ["Job #", "Customer", "CO#", "Tags", "Item (as printed)",
               "Normalized", "Details", "Qty", "Price", "Section", "SO PDF"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Line Items"
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(1, c, h)
        cell.font = header_font
        cell.fill = header_fill

    row = 2
    for h in hits:
        co = f"CO#{h['co_number']}" if h.get("co_number") else ""
        for it in h["matches"]:
            ws.cell(row, 1, h["job"])
            ws.cell(row, 2, h.get("customer", ""))
            ws.cell(row, 3, co)
            ws.cell(row, 4, ", ".join(it.get("tags") or []))
            ws.cell(row, 5, it.get("raw", ""))
            ws.cell(row, 6, it.get("norm", ""))
            ws.cell(row, 7, " ; ".join(it.get("details") or []))
            ws.cell(row, 8, it.get("qty", ""))
            ws.cell(row, 9, it.get("price", ""))
            ws.cell(row, 10, it.get("section", ""))
            if h.get("so_pdf"):
                cell = ws.cell(row, 11, "Open")
                cell.hyperlink = h["so_pdf"]
                cell.font = link_font
            row += 1

    if row > 2:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{row - 1}"
    ws.freeze_panes = "B2"
    for col in range(1, len(headers) + 1):
        letter = get_column_letter(col)
        width = max((len(str(c.value)) for c in ws[letter] if c.value is not None), default=8)
        ws.column_dimensions[letter].width = min(max(width + 2, 6), 70)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def main(argv: Optional[List[str]] = None) -> int:
    logging.basicConfig(level=logging.WARNING, format="[%(levelname)s] %(message)s")
    ap = argparse.ArgumentParser(description="Find orders by their Sales Order line items.")
    ap.add_argument("terms", nargs="*", help="Search terms (all must match; see --any).")
    ap.add_argument("--any", action="store_true", help="Match ANY term instead of all.")
    ap.add_argument("--tag", default="", help='Filter by canonical tag (e.g. "SHAFT SEAL").')
    ap.add_argument("--fuzzy", nargs="?", const=0.84, default=0.0, type=float, metavar="RATIO",
                    help="Also accept close (typo) matches; optional ratio 0-1, default 0.84.")
    ap.add_argument("--job", default="", help="Show the stored items for one job number.")
    ap.add_argument("--list-tags", action="store_true", help="List the tag vocabulary with counts.")
    ap.add_argument("--xlsx", nargs="?", const=str(XLSX_DEFAULT), default="", metavar="PATH",
                    help=f"Write the result (or, with no filters, the full inventory) to Excel "
                         f"(default {XLSX_DEFAULT}).")
    args = ap.parse_args(sys.argv[1:] if argv is None else argv)

    store = li.load_store()
    if not store["jobs"]:
        print(f"The line-items store is empty ({li.store_path()}).\n"
              "It fills from the daily run as orders are parsed — or build it now from\n"
              "the archived Sales Orders:  python line_items_scan.py")
        return 1

    if args.list_tags:
        counts = li.tag_counts(store)
        if not counts:
            print("No tagged items yet — run `python line_items_scan.py --ai` or extend the rules.")
            return 0
        w = max(len(t) for t, _, _ in counts)
        print(f"{'TAG'.ljust(w)}  ORDERS  ITEMS")
        for t, n_jobs, n_items in counts:
            print(f"{t.ljust(w)}  {n_jobs:6d}  {n_items:5d}")
        print(f"\n({_store_stats(store)})")
        return 0

    if args.job:
        rec = store["jobs"].get(args.job)
        if not rec:
            print(f"Job {args.job} is not in the store ({_store_stats(store)}).")
            return 1
        hits = [{"job": args.job, **{k: rec.get(k) for k in
                 ("customer", "co_number", "so_pdf", "scanned_at")},
                 "matches": rec.get("items") or []}]
        _print_hits(hits, all_details=True)
        return 0

    hits = li.search(store, args.terms, any_mode=args.any, tag=args.tag, fuzzy=args.fuzzy)
    if args.terms or args.tag:
        what = " ".join(args.terms) + (f" [tag={args.tag}]" if args.tag else "")
        print(f"{len(hits)} order(s) match {what!r}   ({_store_stats(store)})")
        _print_hits(hits, terms=args.terms)
    elif not args.xlsx:
        print(f"Nothing to search for — give terms, --tag, --job, --list-tags or --xlsx.\n"
              f"({_store_stats(store)})")
        return 1

    if args.xlsx:
        out = write_xlsx(hits, Path(args.xlsx))
        n = sum(len(h["matches"]) for h in hits)
        print(f"\nWrote {n} item row(s) across {len(hits)} order(s) -> {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
