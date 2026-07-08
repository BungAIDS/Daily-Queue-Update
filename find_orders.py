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
    python find_orders.py --audit-untagged     # names current rules still miss
    python find_orders.py --audit-review       # rows marked for human/template review
    python find_orders.py shaft seal --xlsx    # ...write matches to Excel too
    python find_orders.py --xlsx               # full inventory workbook:
                                               # "Line Items" (row per item,
                                               # AutoFilter) + "Feature Matrix"
                                               # (jobs x tags, green ✓ = has
                                               # it / red = doesn't — like the
                                               # AutoCAD DWG matrix)

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
            review_flags = _review_flags_label(it)
            if review_flags:
                print(f"        review: {review_flags}")
            attrs = _attrs_label(it)
            if attrs:
                print(f"        attrs: {attrs}")
            for d in it.get("details") or []:
                if all_details or any(t in d.upper() for t in terms_u):
                    print(f"        · {d}")


def _attrs_label(item: Dict[str, Any]) -> str:
    attrs = item.get("attributes") or {}
    if not isinstance(attrs, dict):
        return ""
    keys = [
        "inquiry_num", "note_type", "component", "component_review",
        "used_on", "used_on_review", "vendor", "product",
        "split_type",
        "drawing_type", "drawing_scope",
        "motor_enclosure", "motor_mounting", "motor_explosion_class",
        "motor_explosion_groups", "motor_explosion_division", "motor_base",
        "motor_feature", "motor_conduit_box_location", "motor_conduit_box_position",
        "motor_conduit_box_orientation",
        "motor_nameplate", "motor_nameplate_action",
        "nameplate_type", "nameplate_mount_location", "nameplate_mounting",
        "flange_scope", "flange_type", "flange_location",
        "flex_connector_type",
        "coupling_subcategory", "coupling_type", "fit", "cover_type", "set_screws",
        "leakage_class", "duty_rating",
        "temperature_service", "temperature_direction", "temperature_rating", "grease_type",
        "special_construction_type", "special_construction_scope",
        "special_construction_detail", "effective_diameter_percent", "welding_code",
        "wheel_feature", "wheel_effective_diameter_percent", "wheel_hub_construction",
        "wheel_hub_bore", "wheel_bore",
        "unitary_base_type", "unitary_base_size", "unitary_base_detail", "unitary_base_clearance",
        "material", "material_grade", "material_scope", "material_treatment",
        "component_material", "component_material_grade", "component_material_scope",
        "drain_type", "drain_closure", "drain_detail",
        "coating_context", "coating_category", "coating_scope", "coating_process",
        "coating_state", "coating_type", "coating_color", "alternate_coating_color", "coats",
        "guard_type", "guard_material", "tach_hole", "tach_hole_location",
        "screen_subcategory", "screen_feature", "screen_diameter", "screen_size",
        "shaft_cooler", "shaft_cooler_type", "shaft_cooler_construction",
        "shaft_seal_type", "shaft_sleeve", "shaft_sleeve_type",
        "silencer_subcategory", "silencer_used_on", "silencer_model",
        "silencer_noise_target", "pressure_drop", "silencer_feature",
        "spark_resistant", "spark_resistant_type",
        "spare_part_type", "spare_part_component",
        "testing_type", "testing_status", "testing_duration", "testing_measurements",
        "testing_voltage",
        "shipping_state", "shipping_method", "shipping_instruction", "shipping_scope",
        "balance_type", "balance_grade",
        "bearing_type", "bearing_bore",
        "manufacturer", "model", "size", "operation", "supplied_by", "mounting", "fail_power", "fail_signal",
        "drive_subcategory", "belt_qty", "belt", "selected_drive", "drive_sheave", "drive_bushing",
        "driven_sheave", "driven_bushing", "drive_sheave_bushing", "driven_sheave_bushing",
        "actual_sf", "actual_cd", "service_factor", "center_distance_range",
    ]
    return "; ".join(f"{k}={attrs[k]}" for k in keys if attrs.get(k))


def _review_flags_label(item: Dict[str, Any]) -> str:
    return "; ".join(str(x) for x in item.get("review_flags") or [])


def _print_untagged_audit(store: Dict[str, Any], limit: int) -> None:
    rows = li.audit_untagged(store, limit=limit)
    if not rows:
        print(f"Current rules tag every stored normalized item. ({_store_stats(store)})")
        return
    w = max(len(r["norm"]) for r in rows)
    print(f"Top {len(rows)} normalized item name(s) current rules still do not tag")
    print(f"({'after re-deriving names/tags from stored raw text'}; {_store_stats(store)})")
    print(f"{'COUNT':>5}  {'NORMALIZED'.ljust(min(w, 72))}  JOBS  AI")
    for r in rows:
        norm = r["norm"]
        shown = norm if len(norm) <= 72 else norm[:69] + "..."
        jobs = ",".join(r.get("jobs") or [])
        ai = ", ".join(r.get("ai_tags") or [])
        print(f"{r['count']:5d}  {shown.ljust(min(w, 72))}  {jobs}  {ai}")


def _print_review_audit(store: Dict[str, Any], limit: int, tag: str = "") -> None:
    rows = li.audit_review(store, limit=limit, tag=tag)
    if not rows:
        tag_msg = f" for tag {tag!r}" if tag else ""
        print(f"No line-item review candidates{tag_msg}. ({_store_stats(store)})")
        return
    w = max(len(r["norm"]) for r in rows)
    tag_msg = f" tagged {tag!r}" if tag else ""
    print(f"Top {len(rows)} line-item template/review candidate(s){tag_msg}")
    print(f"({'after re-deriving names/tags from stored raw text'}; {_store_stats(store)})")
    print(f"{'COUNT':>5}  {'NORMALIZED'.ljust(min(w, 64))}  TAGS  REVIEW FLAGS  JOBS")
    for r in rows:
        norm = r["norm"]
        shown = norm if len(norm) <= 64 else norm[:61] + "..."
        tags = ", ".join(r.get("tags") or []) or "-"
        flags = "; ".join(r.get("review_flags") or [])
        jobs = ",".join(r.get("jobs") or [])
        print(f"{r['count']:5d}  {shown.ljust(min(w, 64))}  {tags}  {flags}  {jobs}")


def write_xlsx(hits: List[Dict[str, Any]], path: Path,
               store: Dict[str, Any] | None = None) -> Path:
    """Two sheets: "Line Items" — one row per (job, item) with AutoFilter, the
    filter-in-Excel view — and "Feature Matrix" — one row per job, one column
    per canonical tag, green ✓ when the order has that feature and red when it
    doesn't (same look as the AutoCAD DWG matrix). When `store` is given the
    matrix shows each job's FULL feature profile even if the hits were
    filtered by a search."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", fgColor="305496")
    header_font = Font(color="FFFFFF", bold=True)
    link_font = Font(color="0563C1", underline="single")

    headers = ["Job #", "Customer", "CO#", "Tags", "Review Flags", "Attributes",
               "Item (as printed)", "Normalized", "Details", "Qty", "Price", "Section", "SO PDF"]
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
            ws.cell(row, 5, _review_flags_label(it))
            ws.cell(row, 6, _attrs_label(it))
            ws.cell(row, 7, it.get("raw", ""))
            ws.cell(row, 8, it.get("norm", ""))
            ws.cell(row, 9, " ; ".join(it.get("details") or []))
            ws.cell(row, 10, it.get("qty", ""))
            ws.cell(row, 11, it.get("price", ""))
            ws.cell(row, 12, it.get("section", ""))
            if h.get("so_pdf"):
                cell = ws.cell(row, 13, "Open")
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

    # ---- Feature Matrix: jobs x tags, green ✓ / red (like the DWG matrix) --
    has_fill = PatternFill("solid", fgColor="C6EFCE")  # green: order HAS the feature
    no_fill = PatternFill("solid", fgColor="FFC7CE")   # red: it doesn't
    center = Alignment(horizontal="center")

    # Each job's full item set — from the store when given, so a search-
    # filtered workbook still shows the whole feature profile per matched job.
    items_by_job: Dict[str, List[Dict[str, Any]]] = {}
    for h in hits:
        rec = (store or {}).get("jobs", {}).get(h["job"]) or {}
        items_by_job[h["job"]] = rec.get("items") or h["matches"]
    tags_by_job = {job: {t for it in items for t in it.get("tags") or []}
                   for job, items in items_by_job.items()}
    counts: Dict[str, int] = {}
    for tags in tags_by_job.values():
        for t in tags:
            counts[t] = counts.get(t, 0) + 1
    all_tags = sorted(counts, key=lambda t: (-counts[t], t))  # most common left

    mx = wb.create_sheet("Feature Matrix")
    fixed = ["Job #", "Customer", "CO#", "Items"]
    for c, h in enumerate(fixed, start=1):
        cell = mx.cell(1, c, h)
        cell.font = header_font
        cell.fill = header_fill
    for k, t in enumerate(all_tags, start=len(fixed) + 1):
        cell = mx.cell(1, k, t)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="bottom", text_rotation=90)
    if all_tags:  # rotated headers keep the tag columns narrow
        mx.row_dimensions[1].height = min(150, 14 + 5.4 * max(len(t) for t in all_tags))

    for i, h in enumerate(hits, start=2):
        cell = mx.cell(i, 1, h["job"])
        if h.get("so_pdf"):
            cell.hyperlink = h["so_pdf"]
            cell.font = link_font
        mx.cell(i, 2, h.get("customer", ""))
        mx.cell(i, 3, f"CO#{h['co_number']}" if h.get("co_number") else "")
        mx.cell(i, 4, len(items_by_job[h["job"]]))
        job_tags = tags_by_job[h["job"]]
        for k, t in enumerate(all_tags, start=len(fixed) + 1):
            cell = mx.cell(i, k)
            if t in job_tags:
                cell.value, cell.fill, cell.alignment = "✓", has_fill, center
            else:
                cell.fill = no_fill

    if hits:
        mx.auto_filter.ref = f"A1:{get_column_letter(len(fixed) + len(all_tags))}{len(hits) + 1}"
    mx.freeze_panes = "B2"
    for c in range(1, len(fixed) + 1):
        letter = get_column_letter(c)
        width = max((len(str(cell.value)) for cell in mx[letter] if cell.value is not None), default=8)
        mx.column_dimensions[letter].width = min(max(width + 2, 6), 40)
    for k in range(len(fixed) + 1, len(fixed) + len(all_tags) + 1):
        mx.column_dimensions[get_column_letter(k)].width = 4.5

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
    ap.add_argument("--list-inquiries", action="store_true",
                    help="List parsed inquiry numbers with order/item counts.")
    ap.add_argument("--audit-untagged", action="store_true",
                    help="List the most common normalized item names current rules still do not tag.")
    ap.add_argument("--audit-review", action="store_true",
                    help="List line-item templates marked for human/rule review. Use --tag to narrow it.")
    ap.add_argument("--audit-limit", type=int, default=50,
                    help="How many audit rows to print (0 = all; default 50).")
    ap.add_argument("--xlsx", nargs="?", const=str(XLSX_DEFAULT), default="", metavar="PATH",
                    help=f"Write the result (or, with no filters, the full inventory) to Excel: "
                         f"a per-item sheet + a green-✓/red jobs-x-features matrix "
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

    if args.list_inquiries:
        counts = li.inquiry_counts(store)
        if not counts:
            print("No inquiry numbers parsed yet.")
            return 0
        w = max(len(n) for n, _, _, _ in counts)
        print(f"{'INQUIRY #'.ljust(w)}  ORDERS  ITEMS  JOBS")
        for num, n_jobs, n_items, jobs in counts:
            print(f"{num.ljust(w)}  {n_jobs:6d}  {n_items:5d}  {', '.join(jobs[:12])}")
        print(f"\n({_store_stats(store)})")
        return 0

    if args.audit_untagged:
        _print_untagged_audit(store, args.audit_limit)
        return 0

    if args.audit_review:
        _print_review_audit(store, args.audit_limit, tag=args.tag)
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
        out = write_xlsx(hits, Path(args.xlsx), store)
        n = sum(len(h["matches"]) for h in hits)
        print(f"\nWrote {n} item row(s) across {len(hits)} order(s) "
              f"(+ the green-✓/red Feature Matrix tab) -> {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
