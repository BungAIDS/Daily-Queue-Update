"""Build / maintain the sales-order line-items store from the PDF archive.

The daily run and the backfill record line items as they parse orders; this
tool does the same WITHOUT a browser or login, straight from the PDFs already
archived under SALES_ORDER_DIR — so the whole archive becomes searchable
(find_orders.py) in one local pass.

    python line_items_scan.py                  # every job folder in the archive
    python line_items_scan.py 421314 421388    # just these jobs (or .pdf paths)
    python line_items_scan.py --limit 200      # first 200 not yet in the store

Tuning (the line items are free text and rarely written the same way twice —
these are how the capture/normalize rules get fitted to the real documents):

    python line_items_scan.py --dump 421314    # per-line view: what was
                                               # captured, what was skipped and
                                               # by which rule. Paste a few of
                                               # these back to tune the rules.
    python line_items_scan.py --renorm         # re-derive every stored item's
                                               # normalized form + tags from
                                               # its verbatim raw text with the
                                               # CURRENT rules (after editing
                                               # the LINE_ITEM_RULES file) —
                                               # no PDFs touched.
    python line_items_scan.py --ai             # send still-untagged unique
                                               # items to Claude once, cache
                                               # the tags forever, apply them.

Only the latest revision of each order is scanned (highest CO# in the file
name); already-stored jobs are skipped unless --rescan. The store is saved
every 50 jobs, so an interrupted run resumes."""
from __future__ import annotations

import argparse
import logging
import re
import sys
from pathlib import Path
from typing import List, Optional, Tuple

from config import SALES_ORDER_DIR
import line_items as li

log = logging.getLogger("line-items-scan")

CO_IN_NAME = re.compile(r"CO#(\d+)", re.I)

# How --dump marks each line of the PDF (kind -> short prefix).
_DUMP_MARKS = {
    "item-priced": "ITEM $", "item-section": "ITEM +", "detail": " +det ",
    "skip": "skip  ", "section-start": "sect> ", "section-end": "<sect ",
    "text": ".     ", "blank": "",
}


def _latest_so_pdf(folder: Path) -> Optional[Tuple[Path, int]]:
    """The newest Sales Order revision in a job's archive folder: highest CO#
    in the file name ('... CO#2.pdf' beats '... (original).pdf'), mtime as the
    tiebreak. Returns (pdf, co_number) or None."""
    best = None
    try:
        for p in folder.glob("*.pdf"):
            m = CO_IN_NAME.search(p.name)
            co = int(m.group(1)) if m else 0
            key = (co, p.stat().st_mtime)
            if best is None or key > best[0]:
                best = (key, p, co)
    except OSError as e:
        log.warning("Could not read %s (%s)", folder, e)
    return (best[1], best[2]) if best else None


def _iter_archive(root: Path) -> List[Tuple[str, Path, int]]:
    """(job, latest SO pdf, CO#) for every job-numbered folder in the archive."""
    out: List[Tuple[str, Path, int]] = []
    if not root.is_dir():
        log.error("Sales-order archive not reachable: %s (set SALES_ORDER_DIR in .env)", root)
        return out
    for d in sorted(root.iterdir()):
        if d.is_dir() and re.fullmatch(r"\d{4,}[A-Za-z]?", d.name.strip()):
            hit = _latest_so_pdf(d)
            if hit:
                out.append((d.name.strip(), hit[0], hit[1]))
    return out


def _resolve(args_jobs: List[str]) -> List[Tuple[str, Path, int]]:
    """Explicit targets: each arg is a job number (looked up in the archive)
    or a path to a Sales Order .pdf."""
    out: List[Tuple[str, Path, int]] = []
    for a in args_jobs:
        p = Path(a)
        if p.is_file():
            m = CO_IN_NAME.search(p.name)
            jm = re.match(r"(\d{4,})", p.parent.name) or re.match(r"(\d{4,})", p.name)
            out.append((jm.group(1) if jm else p.stem, p, int(m.group(1)) if m else 0))
            continue
        folder = SALES_ORDER_DIR / a
        hit = _latest_so_pdf(folder) if folder.is_dir() else None
        if hit:
            out.append((a, hit[0], hit[1]))
        else:
            log.warning("No Sales Order pdf for %r (checked %s)", a, folder)
    return out


def dump(targets: List[Tuple[str, Path, int]]) -> None:
    """The tuning view: every reconstructed line of each PDF, marked with how
    the extractor treated it (and which rule skipped it), then the captured
    items with their normalized form and tags. Paste this back to refine the
    rules — classify_line is the same code the real extraction runs."""
    try:
        import pdfplumber
    except ImportError:
        raise SystemExit("pdfplumber isn't installed yet. Run:  pip install pdfplumber")
    from sales_orders import _recon_lines
    rules = li.load_rules()
    for job, pdf, _co in targets:
        print(f"\n{'=' * 72}\n{job}: {pdf}\n{'=' * 72}")
        lines: List[str] = []
        with pdfplumber.open(str(pdf)) as doc:
            for page in doc.pages:
                lines.extend(_recon_lines(page))
        for kind, detail, s in li.iter_classified(lines, rules):
            if kind == "blank":
                continue
            note = f"   [{detail}]" if kind in ("skip", "item-priced") and detail else ""
            print(f"  {_DUMP_MARKS[kind]}  {s}{note}")
        items = li.extract_items(lines, rules)
        print(f"\n  CAPTURED {len(items)} item(s):")
        for it in items:
            tags = ", ".join(it["tags"]) or "(no tag yet)"
            extras = "  ".join(x for x in (f"qty={it['qty']}" if it['qty'] else "",
                                           f"price={it['price']}" if it['price'] else "") if x)
            print(f"    [{tags}]  {it['norm']}" + (f"   ({extras})" if extras else ""))
            for d in it["details"]:
                print(f"        · {d}")


def scan(targets: List[Tuple[str, Path, int]], rescan: bool, limit: int) -> int:
    """Parse each target's SO pdf and record its line items in the store."""
    try:
        import pdfplumber  # noqa: F401 - parse degrades to zero items without it
    except ImportError:
        raise SystemExit("pdfplumber isn't installed yet. Run:  pip install pdfplumber")
    from sales_orders import parse_sales_order_pdf
    store = li.load_store()
    done = 0
    for job, pdf, co in targets:
        if limit and done >= limit:
            break
        if not rescan and job in store["jobs"]:
            continue
        items = parse_sales_order_pdf(pdf).get("line_items") or []
        li.apply_ai_cache(items, store)
        li.record_job(store, job, items, co_number=co, so_pdf=str(pdf))
        done += 1
        log.info("  %d  %s -> %d item(s)", done, job, len(items))
        if done % 50 == 0:
            li.save_store(store)
    li.save_store(store)
    log.info("Scanned %d order(s) this run; store now holds %d job(s) -> %s",
             done, len(store["jobs"]), li.store_path())
    return 0


def main(argv: Optional[List[str]] = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    ap = argparse.ArgumentParser(description="Scan archived Sales Order PDFs into the line-items store.")
    ap.add_argument("jobs", nargs="*", help="Job numbers or SO .pdf paths (default: the whole archive).")
    ap.add_argument("--dump", action="store_true",
                    help="Print the per-line capture view for the given jobs/PDFs (no store writes).")
    ap.add_argument("--renorm", action="store_true",
                    help="Re-derive every stored item's norm/tags from raw with the current rules.")
    ap.add_argument("--ai", action="store_true",
                    help="Classify still-untagged unique items with Claude (cached forever).")
    ap.add_argument("--rescan", action="store_true", help="Re-parse jobs already in the store.")
    ap.add_argument("--limit", type=int, default=0, help="Stop after N new jobs this run.")
    args = ap.parse_args(sys.argv[1:] if argv is None else argv)

    if args.dump:
        targets = _resolve(args.jobs) if args.jobs else _iter_archive(SALES_ORDER_DIR)[:3]
        if not targets:
            return 1
        dump(targets)
        return 0

    if args.renorm:
        store = li.load_store()
        n = li.renormalize_store(store)
        li.save_store(store)
        log.info("Re-normalized %d item(s) across %d job(s) with the current rules.",
                 n, len(store["jobs"]))
        return 0

    rc = 0
    if args.jobs or not args.ai:
        targets = _resolve(args.jobs) if args.jobs else _iter_archive(SALES_ORDER_DIR)
        if not targets and not args.ai:
            return 1
        rc = scan(targets, args.rescan, args.limit)

    if args.ai:
        store = li.load_store()
        before = len(li.unknown_norms(store))
        try:
            n = li.ai_classify_unknowns(store)
        except RuntimeError as e:
            log.error("%s", e)
            return 1
        li.renormalize_store(store)  # fold the fresh AI tags into every job's items
        li.save_store(store)
        log.info("AI pass: %d of %d unknown unique item(s) classified (cached; re-runs are free).",
                 n, before)
    return rc


if __name__ == "__main__":
    sys.exit(main())
