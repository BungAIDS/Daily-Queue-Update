"""Backfill — look up OLD orders on cbcinsider one at a time, all day, resumably.

The daily run only enriches what's on the board today. This walks a much larger
backlog of historical jobs, opens each one's detail, downloads + parses its
Sales Order and (if present) construction/drive run, merges in the AutoCAD DWG
scan, and writes a master backlog workbook.

It is deliberately serial ("1 by 1") and gentle on the server, and it writes its
progress after every job — kill it any time and re-run to pick up where it left
off.

Job source (pick one; default is the AutoCAD folders):
    python backfill_orders.py                     # every job folder under AUTOCAD_JOBS_DIR
    python backfill_orders.py 421314 421388       # explicit job numbers
    python backfill_orders.py --list jobs.txt     # one job number per line
    python backfill_orders.py --range 420000 421000
Options:
    --limit N     stop after N jobs this run        --delay S   seconds between jobs (default 1.0)
    --rescan      ignore saved progress             --out PATH  workbook path

Outputs (under BACKLOG_DIR):
    backfill_progress.json   resumable per-job store (source of truth)
    backlog.xlsx             master sheet: SO/drive-run fields + DWG suffix matrix

Old orders are opened through the queue page's "search order" / "find order"
box: each job number is typed in and the surfaced order's detail is opened. The
box is auto-detected; if that misses on your layout, set CBC_SEARCH_SELECTOR
(and optionally CBC_SEARCH_BUTTON) in .env — `python discover_documents.py
--probe <job#>` prints the exact selector. The run preflights the box and stops
with a clear message rather than grinding the whole list if it can't be found.
"""
from __future__ import annotations

import argparse
import json
import logging
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import urljoin

from config import (
    CBC_URL, CBC_QUEUE_URL, STORAGE_STATE_PATH,
    SALES_ORDER_DIR, DRIVE_RUN_DIR, BACKLOG_DIR, AUTOCAD_JOBS_DIR,
    CBC_SEARCH_SELECTOR, CBC_SEARCH_BUTTON,
)
from templates import parse_quote_run
from sales_orders import (
    _parse_doc, _latest_of_type, _run_docs, _run_filename, _run_files_in_folder,
    _trigger_js, _so_filename, _download_error, SO_TYPE, parse_sales_order_pdf,
)
from scraper import CONTAINER_SELECTOR
import autocad_scan
import line_items

log = logging.getLogger("backfill")

PROGRESS_PATH = BACKLOG_DIR / "backfill_progress.json"
WORKBOOK_PATH = BACKLOG_DIR / "backlog.xlsx"


# --------------------------------------------------------------------------- #
# Job sources                                                                 #
# --------------------------------------------------------------------------- #
def jobs_from_folders(root: Path, min_job: int = autocad_scan.DEFAULT_MIN_JOB, max_job: int = 0) -> List[str]:
    """Every real job number under AUTOCAD_JOBS_DIR — same list the DWG scan walks.
    Folders below min_job (year/template/archive dirs) are skipped."""
    return [job for job, _type, _path in autocad_scan.iter_job_folders(root, min_job, max_job)]


def jobs_from_list(path: Path) -> List[str]:
    return [ln.strip() for ln in path.read_text(encoding="utf-8").splitlines() if ln.strip()]


def jobs_from_range(first: int, last: int) -> List[str]:
    lo, hi = sorted((first, last))
    return [str(n) for n in range(lo, hi + 1)]


# --------------------------------------------------------------------------- #
# Browser helpers (sync — serial by design)                                   #
# --------------------------------------------------------------------------- #
def _jobnum(args_js: str) -> str:
    return args_js.split(",", 1)[0].strip().strip("'\"").split("-", 1)[0]


def _board_args(page) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for c in page.locator(CONTAINER_SELECTOR).all():
        m = re.search(r"loadDetail\((.*?)\)", c.get_attribute("onclick") or "")
        if m:
            out[_jobnum(m.group(1).strip())] = m.group(1).strip()
    return out


def _open_via_loaddetail(page, job: str, args_js: str) -> bool:
    page.evaluate(_trigger_js(args_js))
    try:
        # The modal's docs typically load in ~30s. The backfill is serial
        # (uncontended), so 45s is ample headroom — the daily run's parallel
        # fetch keeps its longer 90s wait. This bounds the per-job worst case
        # on an all-day grind.
        page.locator("#modalDetail a").filter(
            has_text=re.compile(re.escape(job))).first.wait_for(state="attached", timeout=45000)
        return True
    except Exception:  # noqa: BLE001
        return False


def find_search_box(page):
    """Locate the queue page's "search order" / "find order" box, or None.

    Honors CBC_SEARCH_SELECTOR if set; otherwise matches a text input whose
    placeholder / aria-label / title / id / name mentions "search" or "find"
    (preferring one that also mentions "order")."""
    if CBC_SEARCH_SELECTOR:
        loc = page.locator(CBC_SEARCH_SELECTOR)
        return loc.first if loc.count() else None
    best = None
    for el in page.locator("input[type=text], input[type=search], input:not([type])").all():
        blob = " ".join(filter(None, [
            el.get_attribute("placeholder"), el.get_attribute("aria-label"),
            el.get_attribute("title"), el.get_attribute("id"), el.get_attribute("name"),
        ])).lower()
        if "search" in blob or "find" in blob:
            if "order" in blob:
                return el          # best: a "search order" / "find order" box
            best = best or el      # fallback: any search/find input
    return best


def _modal_has_job(page, job: str) -> bool:
    try:
        return page.locator("#modalDetail a").filter(
            has_text=re.compile(re.escape(job))).count() > 0
    except Exception:  # noqa: BLE001
        return False


def open_order_detail(page, job: str) -> bool:
    """Surface `job`'s documents in #modalDetail via the queue's search box.

    Types the job number into the search bar, submits (Enter, plus an optional
    CBC_SEARCH_BUTTON click), then waits for either the board to re-render with
    the order (and opens it via loadDetail) or the detail modal to populate
    directly. Returns True once the documents are present, else False."""
    box = find_search_box(page)
    if box is None:
        return False
    try:
        box.click()
        box.fill("")
        box.fill(job)
        box.press("Enter")
    except Exception:  # noqa: BLE001
        return False
    if CBC_SEARCH_BUTTON:
        try:
            btn = page.locator(CBC_SEARCH_BUTTON)
            if btn.count():
                btn.first.click()
        except Exception:  # noqa: BLE001
            pass
    # Poll: the search may re-render the board (postback/filter) or open the
    # detail directly. Either way, wait for the order to surface.
    deadline = time.monotonic() + 20
    while time.monotonic() < deadline:
        page.wait_for_timeout(500)
        amap = _board_args(page)
        if job in amap and _open_via_loaddetail(page, job, amap[job]):
            return True
        if _modal_has_job(page, job):
            return True
    return False


def _collect_docs(page) -> List[Dict[str, Any]]:
    docs = []
    for a in page.locator("#modalDetail a").all():
        href = a.get_attribute("href") or ""
        if "downloaddoc.aspx" in href.lower():
            d = _parse_doc(href)
            d["href"] = href
            docs.append((href, d))
    return docs


def _download(context, page_url: str, href: str, dest: Path) -> Optional[str]:
    if dest.exists():
        return str(dest)
    try:
        resp = context.request.get(urljoin(page_url, href), timeout=60000)
        body = resp.body()
        # Never archive an error page / expired-session login page as the PDF —
        # the dest.exists() cache would then skip re-downloading it forever.
        err = _download_error(resp.status, body, dest.suffix.lower() == ".pdf")
        if err:
            log.warning("  download failed for %s: %s", dest.name, err)
            return None
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(body)
        return str(dest)
    except Exception as e:  # noqa: BLE001
        log.warning("  download failed for %s: %s", dest.name, e)
        return None


def _close_modal(page) -> None:
    try:
        page.evaluate("() => window.jQuery && jQuery('#modalDetail').modal('hide')")
        page.wait_for_timeout(250)
    except Exception:  # noqa: BLE001
        pass


def process_one(page, context, job: str, folder: str = "",
                li_store: Dict[str, Any] | None = None) -> Dict[str, Any]:
    """Open one order via the search box, download + parse its SO and quote run.
    `folder` (the job's AutoCAD folder, when known from the DWG scan) is checked
    for a quote-run file if the order's documents don't carry one. `li_store`
    (the shared line-items store) gets the order's parsed line items — the
    caller saves it alongside the progress file."""
    rec: Dict[str, Any] = {"job": job, "status": "", "scanned_at": datetime.now().isoformat(timespec="seconds")}
    try:
        if not open_order_detail(page, job):
            rec["status"] = "not-found"  # search returned no order/documents for this job#
            return rec
        docs = _collect_docs(page)

        so_pdf = dr_pdf = None
        so = _latest_of_type(docs, SO_TYPE)
        if so:
            href, doc = so
            rec["co_number"] = (doc["rev"] - 1) if doc["rev"] and doc["rev"] > 1 else 0
            so_pdf = _download(context, page.url, href, SALES_ORDER_DIR / job / _so_filename(job, doc["rev"]))
            if so_pdf:
                p = parse_sales_order_pdf(so_pdf)
                rec.update({
                    "so_design_desc": p.get("design_desc", ""), "so_size": p.get("size", ""),
                    "so_arrangement": p.get("arrangement", ""), "so_motor_pos": p.get("motor_pos", ""),
                    "so_class": p.get("fan_class", ""), "so_rotation": p.get("rotation", ""),
                    "so_discharge": p.get("discharge", ""), "so_pct_width": p.get("pct_width", ""),
                    "so_wheel_type": p.get("wheel_type", ""),
                    "so_design_temp": p.get("design_temp", ""), "so_max_temp": p.get("max_temp", ""),
                    "so_special_temp": p.get("special_temp", ""),
                    "so_pdf": so_pdf,
                })
                items = p.get("line_items") or []
                rec["line_item_count"] = len(items)
                if li_store is not None:
                    line_items.apply_ai_cache(items, li_store)
                    line_items.record_job(li_store, job, items,
                                          co_number=rec.get("co_number"), so_pdf=so_pdf)

        runs = _run_docs(docs)
        dr_count = 0
        if runs:
            rec["has_drive_run"] = True
            dr_count = len(runs)
            for href, doc in runs:
                got = _download(context, page.url, href, DRIVE_RUN_DIR / job / _run_filename(job, doc))
                if got and not dr_pdf:
                    dr_pdf = got
            rec["drive_run_pdf"] = dr_pdf or ""
            if dr_pdf:
                qr = parse_quote_run(dr_pdf, design=rec.get("design"))
                rec["drive_run_summary"] = qr.get("summary", "")
                rec["drive_run_template"] = qr.get("template", "")
        if folder:
            # Always scan the AutoCAD folder — cheap rglob, catches runs that
            # only live there and any folder copies alongside document ones.
            hits = _run_files_in_folder(Path(folder))
            if hits:
                if not rec.get("has_drive_run"):
                    rec["has_drive_run"] = True
                    rec["drive_run_pdf"] = str(hits[0])
                dr_count += len(hits)
        if dr_count:
            rec["drive_run_count"] = dr_count

        # "ok" only when every document we found actually downloaded — a found-
        # but-failed download stays "error" so the resume retries it, instead of
        # permanently recording the job with empty fields.
        if not so:
            rec["status"] = "no-SO"
        elif so_pdf and (not runs or dr_pdf):
            rec["status"] = "ok"
        else:
            rec["status"] = "error"
    except Exception as e:  # noqa: BLE001 - one bad order never stops the backlog
        log.warning("  %s: error (%s)", job, e)
        rec["status"] = "error"
    finally:
        _close_modal(page)
    return rec


# --------------------------------------------------------------------------- #
# Progress store                                                              #
# --------------------------------------------------------------------------- #
def load_progress() -> Dict[str, Dict[str, Any]]:
    if not PROGRESS_PATH.exists():
        return {}
    try:
        return json.loads(PROGRESS_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as e:
        log.warning("Could not read %s (%s); starting fresh", PROGRESS_PATH, e)
        return {}


def save_progress(records: Dict[str, Dict[str, Any]]) -> None:
    PROGRESS_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp = PROGRESS_PATH.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(records, indent=2), encoding="utf-8")
    tmp.replace(PROGRESS_PATH)  # atomic — a crash mid-write never corrupts the store


def _is_done(rec: Dict[str, Any]) -> bool:
    """A job is 'done' (skip on resume) once we have a real answer for it.
    'error' is NOT done — those get retried on the next run."""
    return bool(rec) and rec.get("status") in ("ok", "no-SO", "not-found")


# --------------------------------------------------------------------------- #
# Master workbook (SO/drive-run fields + DWG suffix matrix)                   #
# --------------------------------------------------------------------------- #
def write_workbook(records: Dict[str, Dict[str, Any]], dwg: Dict[str, Dict[str, Any]], path: Path) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    from excel_writer import _drive_run_label

    header_fill = PatternFill("solid", fgColor="305496")
    header_font = Font(color="FFFFFF", bold=True)
    link_font = Font(color="0563C1", underline="single")
    dr_font = Font(color="C55A11", bold=True)
    dr_link_font = Font(color="C55A11", bold=True, underline="single")  # Drive Run -> PDF link
    present_fill = PatternFill("solid", fgColor="C6EFCE")  # green: job HAS this drawing
    absent_fill = PatternFill("solid", fgColor="FFC7CE")   # red: it doesn't
    center = Alignment(horizontal="center")

    suffixes = autocad_scan.all_extra_suffixes(dwg)
    # -01/-02 (CW/CCW) aren't shown — nearly every job has them; only the custom
    # extra suffixes carry signal.
    fixed = ["Job #", "Type", "Description", "Size", "Arrangement", "Motor Pos", "Class",
             "Rotation", "Discharge", "% Width", "Special Temp", "CO#", "Quote Run",
             "Quote Run Summary", "Folder", "Order Status"]
    headers = fixed + [f"-{s}" for s in suffixes]
    folder_col = fixed.index("Folder") + 1

    wb = Workbook()
    ws = wb.active
    ws.title = "Backlog"
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(1, c, h)
        cell.font = header_font
        cell.fill = header_fill

    jobs = sorted(set(records) | set(dwg))
    for i, job in enumerate(jobs, start=2):
        r = records.get(job, {})
        d = dwg.get(job, {})
        vals = [
            job, d.get("type", ""), r.get("so_design_desc", ""), r.get("so_size", ""),
            r.get("so_arrangement", ""), r.get("so_motor_pos", ""), r.get("so_class", ""),
            r.get("so_rotation", ""), r.get("so_discharge", ""), r.get("so_pct_width", ""),
            r.get("so_special_temp", ""), (f"CO#{r['co_number']}" if r.get("co_number") else ""),
            _drive_run_label(r), r.get("drive_run_summary", ""),
            "", r.get("status", ""),  # "" = Folder placeholder (hyperlinked below)
        ]
        for c, v in enumerate(vals, start=1):
            ws.cell(i, c, v)
        # Drive Run cell links to the archived drive-run PDF when we have it.
        if r.get("has_drive_run"):
            dr_cell = ws.cell(i, fixed.index("Quote Run") + 1)
            if r.get("drive_run_pdf"):
                dr_cell.hyperlink = r["drive_run_pdf"]
                dr_cell.font = dr_link_font
            else:
                dr_cell.font = dr_font
        # Link the AutoCAD folder if we scanned one, else the SO archive folder.
        folder = d.get("folder") or (str(Path(r["so_pdf"]).parent) if r.get("so_pdf") else "")
        if folder:
            fcell = ws.cell(i, folder_col, "Open")
            fcell.hyperlink = folder
            fcell.font = link_font
        extras = d.get("extras", {})
        for k, s in enumerate(suffixes, start=len(fixed) + 1):
            cell = ws.cell(i, k)
            if s in extras:  # green + a tiny check; red + blank when absent
                cell.value, cell.fill, cell.alignment = "✓", present_fill, center
            else:
                cell.fill = absent_fill

    if jobs:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(jobs) + 1}"
    ws.freeze_panes = "B2"
    for col in range(1, len(headers) + 1):
        letter = get_column_letter(col)
        width = max((len(str(c.value)) for c in ws[letter] if c.value is not None), default=8)
        ws.column_dimensions[letter].width = min(max(width + 2, 6), 48)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


# --------------------------------------------------------------------------- #
# CLI                                                                         #
# --------------------------------------------------------------------------- #
def _resolve_jobs(args: argparse.Namespace) -> List[str]:
    if args.jobs:
        return args.jobs
    if args.list:
        return jobs_from_list(Path(args.list))
    if args.range:
        return jobs_from_range(args.range[0], args.range[1])
    return jobs_from_folders(Path(args.root), args.min_job, args.max_job)


def main(argv: Optional[List[str]] = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    ap = argparse.ArgumentParser(description="Backfill old orders from cbcinsider, one at a time.")
    ap.add_argument("jobs", nargs="*", help="Explicit job numbers (default: all AutoCAD folders).")
    ap.add_argument("--list", help="File of job numbers, one per line.")
    ap.add_argument("--range", nargs=2, type=int, metavar=("FIRST", "LAST"), help="Numeric job range.")
    ap.add_argument("--root", default=str(AUTOCAD_JOBS_DIR), help="AutoCAD jobs root for folder enumeration.")
    ap.add_argument("--out", default=str(WORKBOOK_PATH), help="Master workbook path.")
    ap.add_argument("--delay", type=float, default=1.0, help="Seconds to pause between orders.")
    ap.add_argument("--limit", type=int, default=0, help="Stop after N orders this run (0 = no limit).")
    ap.add_argument("--min-job", type=int, default=autocad_scan.DEFAULT_MIN_JOB,
                    help=f"On a folder sweep, skip job numbers below this (default {autocad_scan.DEFAULT_MIN_JOB}).")
    ap.add_argument("--max-job", type=int, default=0, help="Skip job numbers above this (0 = no cap).")
    ap.add_argument("--rescan", action="store_true", help="Ignore saved progress.")
    args = ap.parse_args(sys.argv[1:] if argv is None else argv)

    if not STORAGE_STATE_PATH.exists():
        log.error("No saved session at %s. Run `python login.py` first.", STORAGE_STATE_PATH)
        return 1

    targets = _resolve_jobs(args)
    log.info("Backfill target set: %d job(s).", len(targets))
    records = {} if args.rescan else load_progress()
    li_store = line_items.load_store()  # shared lookup store; saved with progress
    dwg = autocad_scan.load_progress()  # read-only merge of the DWG scan, if it's been run
    if dwg:
        log.info("Merging AutoCAD DWG scan for %d job(s).", len(dwg))

    from playwright.sync_api import sync_playwright
    processed = 0
    rc = 0
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(storage_state=str(STORAGE_STATE_PATH), accept_downloads=True)
        page = context.new_page()
        page.goto(CBC_QUEUE_URL or CBC_URL, wait_until="domcontentloaded", timeout=60000)
        if "login" in page.url.lower() or page.locator('input[type="password"]').count() > 0:
            browser.close()
            log.error("Landed on login — session expired. Re-run `python login.py`.")
            return 1
        page.wait_for_selector(CONTAINER_SELECTOR, timeout=45000)

        # Preflight: the backfill drives the queue page's "search order" box.
        # If we can't find it, don't grind through the whole list — tell the
        # operator how to pin it and stop (the DWG merge still writes output).
        if find_search_box(page) is None:
            log.error("Could not find the 'search order' box on %s. Set CBC_SEARCH_SELECTOR in .env "
                      "to its CSS selector (run `python discover_documents.py --probe %s` to print it), "
                      "then re-run.", CBC_QUEUE_URL or CBC_URL, targets[0] if targets else "<job#>")
            rc = 1  # fail loudly so a scripted/scheduled run notices (workbook still written below)
        else:
            for job in targets:
                if args.limit and processed >= args.limit:
                    break
                if not args.rescan and _is_done(records.get(job, {})):
                    continue
                records[job] = process_one(page, context, job, dwg.get(job, {}).get("folder", ""),
                                           li_store=li_store)
                processed += 1
                log.info("  %d  %s -> %s", processed, job, records[job].get("status"))
                if processed % 25 == 0:
                    save_progress(records)
                    line_items.save_store(li_store)
                if args.delay:
                    time.sleep(args.delay)

        browser.close()

    save_progress(records)
    line_items.save_store(li_store)
    try:   # fold the backfilled SO spec + drive runs + line items into the one master store
        import master_sync
        master_sync.run("backfill", "line_items")
    except Exception as e:  # noqa: BLE001
        log.warning("Could not sync backfill to the live master (%s)", e)
    out = write_workbook(records, dwg, Path(args.out))
    by_status: Dict[str, int] = {}
    for r in records.values():
        by_status[r.get("status", "?")] = by_status.get(r.get("status", "?"), 0) + 1
    log.info("Done: processed %d this run; store has %d job(s).", processed, len(records))
    log.info("  status breakdown: %s", ", ".join(f"{k}={v}" for k, v in sorted(by_status.items())))
    if by_status.get("not-found"):
        log.info("  %d job(s) not found via search — if that seems high, the search box may need "
                 "CBC_SEARCH_SELECTOR / CBC_SEARCH_BUTTON set (see `discover_documents.py --probe`).",
                 by_status["not-found"])
    log.info("  Wrote %s", out)
    return rc


if __name__ == "__main__":
    sys.exit(main())
