"""Print one order's Sales-Order line items as a component hierarchy — the
deep-troubleshooting companion to the live workbook's Sales Order tab.

    python so_tree.py 421966                # the component tree from the store
    python so_tree.py 421966 --flat         # the raw capture, item by item
    python so_tree.py --lines dump.txt      # classify reconstructed SO text
                                            # (from dump_pdf.py): shows how the
                                            # extractor treats EVERY line —
                                            # item / detail / skipped — which
                                            # the store alone can never show

The tree (see so_hierarchy) shows what we KNOW about the job, not what the SO
printed: ONE COMPONENT per real thing — lines the extractors tied together
(shared used_on / component attribute, e.g. the three IVC charges) merge into
it — with every merged FACT, stored DETAIL sub-line and REVIEW flag (including
fact conflicts) beneath it, and the contributing SOURCE lines at the bottom.
A wrong tree therefore means a wrong capture or attribute — use --flat to see
the stored items exactly as captured, and --lines to replay the extractor over
the PDF's text when the problem is a line that was skipped or merged before it
ever reached the store.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import line_items as li
import so_hierarchy as soh


def _print_header(jn: str, rec: dict) -> None:
    co = f"  CO#{rec['co_number']}" if rec.get("co_number") else ""
    cust = f"  {rec['customer']}" if rec.get("customer") else ""
    scanned = f"   (scanned {rec.get('scanned_at', '')[:10]})" if rec.get("scanned_at") else ""
    print(f"{jn}{cust}{co}{scanned}")
    if rec.get("so_pdf"):
        print(f"SO: {rec['so_pdf']}")
    print()


def print_tree(jn: str, rec: dict) -> None:
    _print_header(jn, rec)
    rows = soh.tree_rows(rec.get("items") or [])
    if not rows:
        print("(no line items captured for this order)")
        return
    for r in rows:
        no = f"#{r['item_no']}" if r["item_no"] else ""
        price = f"  {r['price']}" if r["price"] else ""
        print(f"{r['kind']:>9} {no:>4}  {soh.indent_text(r)}{price}")


def print_flat(jn: str, rec: dict) -> None:
    _print_header(jn, rec)
    items = rec.get("items") or []
    if not items:
        print("(no line items captured for this order)")
        return
    for i, it in enumerate(items, start=1):
        print(f"#{i}  {it.get('raw', '')}")
        if it.get("section"):
            print(f"      section: {it['section']}")
        print(f"      norm: {it.get('norm', '')}")
        qty, price, ptype = it.get("qty", ""), it.get("price", ""), it.get("ptype", "")
        if qty or price or ptype:
            print(f"      qty/price/type: {qty} / {price} / {ptype}")
        if it.get("tags"):
            print(f"      tags: {', '.join(it['tags'])}")
        attrs = it.get("attributes") or {}
        if attrs:
            print("      attrs: " + "; ".join(f"{k}={v}" for k, v in sorted(attrs.items()) if v))
        for d in it.get("details") or []:
            print(f"      · {d}")
        if it.get("review_flags"):
            print(f"      REVIEW: {'; '.join(str(f) for f in it['review_flags'])}")
        print()


def print_lines_trace(path: Path) -> int:
    """Replay the extractor's classifier over reconstructed SO text (one line
    per line, e.g. dump_pdf.py output) and print the verdict for each — the
    only view that shows lines the extractor SKIPPED or folded away."""
    try:
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError as e:
        print(f"Could not read {path}: {e}", file=sys.stderr)
        return 2
    labels = {"item-priced": "ITEM", "item-section": "ITEM", "detail": "detail",
              "section-start": "SECTION>", "section-end": "<SECTION",
              "skip": "-skip-", "text": "(text)"}
    for kind, detail, s in li.iter_classified(lines):
        label = labels.get(kind, kind)
        extra = f"  [{detail}]" if detail and kind != "detail" else ""
        print(f"{label:>9}  {s}{extra}")
    return 0


def main(argv: list | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("job", nargs="?", help="order number, e.g. 421966")
    ap.add_argument("--flat", action="store_true",
                    help="print the raw capture per item instead of the tree")
    ap.add_argument("--lines", type=Path, metavar="FILE",
                    help="classify a file of reconstructed SO text lines "
                         "(dump_pdf.py output): item / detail / skipped")
    args = ap.parse_args(argv)

    if args.lines:
        return print_lines_trace(args.lines)
    if not args.job:
        ap.print_help()
        return 2

    jn = str(args.job).strip()
    rec = (li.load_store().get("jobs") or {}).get(jn)
    if not rec:
        print(f"Order {jn} is not in the line-items store ({li.store_path()}).\n"
              f"It appears there once its Sales Order has been parsed (the "
              f"watcher/daily run for board orders, backfill for archived ones).")
        return 1
    (print_flat if args.flat else print_tree)(jn, rec)
    return 0


if __name__ == "__main__":
    sys.exit(main())
