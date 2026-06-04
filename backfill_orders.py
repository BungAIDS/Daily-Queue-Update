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

>>> THE ONE UNKNOWN is how to open an order that is NOT on the board. Run
    `python discover_documents.py --probe <old-job#>` to find the exact lookup,
    then wire it into open_order_detail() below (marked SEAM). Until then this
    handles any target that happens to still be on the board and records the
    rest as "lookup-unconfirmed" — the AutoCAD/DWG half of the backlog still
    fills in regardless.
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
from typing import Any, Dict, Iterable, List, Optional
from urllib.parse import urljoin

from config import (
    CBC_URL, CBC_QUEUE_URL, STORAGE_STATE_PATH,
    SALES_ORDER_DIR, DRIVE_RUN_DIR, BACKLOG_DIR, AUTOCAD_JOBS_DIR,
)
from drive_run import parse_drive_run_pdf
from sales_orders import (
    _parse_doc, _latest_of_type, _trigger_js, _so_filename, _doc_filename,
    SO_TYPE, DRIVE_RUN_TYPE, parse_sales_order_pdf,
)
from scraper import CONTAINER_SELECTOR
import autocad_scan

log = logging.getLogger("backfill")

PROGRESS_PATH = BACKLOG_DIR / "backfill_progress.json"
WORKBOOK_PATH = BACKLOG_DIR / "backlog.xlsx"


# --------------------------------------------------------------------------- #
# Job sources                                                                 #
# --------------------------------------------------------------------------- #
def jobs_from_folders(root: Path) -> List[str]:
    """Every job number under AUTOCAD_JOBS_DIR — same list the DWG scan walks."""
    return [job for job, _type, _path in autocad_scan.iter_job_folders(root)]


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
        page.locator("#modalDetail a").filter(
            has_text=re.compile(re.escape(job))).first.wait_for(state="attached", timeout=90000)
        return True
    except Exception:  # noqa: BLE001
        return False


def open_order_detail(page, job: str, board_args: Dict[str, str]) -> bool:
    """Surface `job`'s documents in #modalDetail. Returns True if they appear.

    If the job is still on the board we just call its loadDetail(...). Otherwise
    we hit the SEAM below — the off-board lookup that discovery must confirm.
    """
    args = board_args.get(job)
    if args:
        return _open_via_loaddetail(page, job, args)

    # ---------------------------------------------------------------- SEAM ---
    # OFF-BOARD LOOKUP — replace with the mechanism `discover_documents.py
    # --probe` confirms (a search box, a deep-link URL, or loadDetail with a
    # known arg shape). The search-box attempt below is a best guess; if it
    # doesn't apply on your site it simply returns False and the job is recorded
    # as "lookup-unconfirmed" so nothing crashes.
    box = page.locator(
        "input#MainContent_txtSearch, input[id*=earch], input[placeholder*=earch]").first
    if box.count():
        try:
            box.fill(job)
            box.press("Enter")
            page.wait_for_timeout(4000)
            refreshed = _board_args(page)
            if job in refreshed:
                return _open_via_loaddetail(page, job, refreshed[job])
        except Exception:  # noqa: BLE001
            return False
    return False
    # ------------------------------------------------------------------------ #


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
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(resp.body())
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


def process_one(page, context, job: str, board_args: Dict[str, str]) -> Dict[str, Any]:
    """Open one order, download + parse its SO and drive run."""
    rec: Dict[str, Any] = {"job": job, "status": "", "scanned_at": datetime.now().isoformat(timespec="seconds")}
    try:
        if not open_order_detail(page, job, board_args):
            rec["status"] = "no-order" if job in board_args else "lookup-unconfirmed"
            return rec
        docs = _collect_docs(page)

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
                    "so_wheel_type": p.get("wheel_type", ""), "so_special_temp": p.get("special_temp", ""),
                    "so_pdf": so_pdf,
                })

        dr = _latest_of_type(docs, DRIVE_RUN_TYPE)
        if dr:
            href, doc = dr
            dr_pdf = _download(context, page.url, href, DRIVE_RUN_DIR / job / _doc_filename(job, "Drive Run", doc["rev"]))
            rec["has_drive_run"] = True
            rec["drive_run_pdf"] = dr_pdf or ""
            rec["drive_run_summary"] = parse_drive_run_pdf(dr_pdf).get("summary", "") if dr_pdf else ""

        rec["status"] = "ok" if so else "no-SO"
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
    'lookup-unconfirmed' is NOT done — it should be retried once the seam works."""
    return bool(rec) and rec.get("status") in ("ok", "no-SO", "no-order")


# --------------------------------------------------------------------------- #
# Master workbook (SO/drive-run fields + DWG suffix matrix)                   #
# --------------------------------------------------------------------------- #
def write_workbook(records: Dict[str, Dict[str, Any]], dwg: Dict[str, Dict[str, Any]], path: Path) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", fgColor="305496")
    header_font = Font(color="FFFFFF", bold=True)
    link_font = Font(color="0563C1", underline="single")
    dr_font = Font(color="C55A11", bold=True)

    suffixes = autocad_scan.all_extra_suffixes(dwg)
    fixed = ["Job #", "Type", "Description", "Size", "Arrangement", "Motor Pos", "Class",
             "Rotation", "Discharge", "% Width", "Special Temp", "CO#", "Drive Run",
             "Drive Run Summary", "CW (01)", "CCW (02)", "Folder", "Order Status"]
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
            "YES" if r.get("has_drive_run") else "", r.get("drive_run_summary", ""),
            d.get("cw", ""), d.get("ccw", ""), "", r.get("status", ""),
        ]
        for c, v in enumerate(vals, start=1):
            ws.cell(i, c, v)
        if r.get("has_drive_run"):
            ws.cell(i, fixed.index("Drive Run") + 1).font = dr_font
        # Link the AutoCAD folder if we scanned one, else the SO archive folder.
        folder = d.get("folder") or (str(Path(r["so_pdf"]).parent) if r.get("so_pdf") else "")
        if folder:
            fcell = ws.cell(i, folder_col, "Open")
            fcell.hyperlink = folder
            fcell.font = link_font
        extras = d.get("extras", {})
        for k, s in enumerate(suffixes, start=len(fixed) + 1):
            ws.cell(i, k, "yes" if s in extras else "no")

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
    return jobs_from_folders(Path(args.root))


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
    ap.add_argument("--rescan", action="store_true", help="Ignore saved progress.")
    args = ap.parse_args(sys.argv[1:] if argv is None else argv)

    if not STORAGE_STATE_PATH.exists():
        log.error("No saved session at %s. Run `python login.py` first.", STORAGE_STATE_PATH)
        return 1

    targets = _resolve_jobs(args)
    log.info("Backfill target set: %d job(s).", len(targets))
    records = {} if args.rescan else load_progress()
    dwg = autocad_scan.load_progress()  # read-only merge of the DWG scan, if it's been run
    if dwg:
        log.info("Merging AutoCAD DWG scan for %d job(s).", len(dwg))

    from playwright.sync_api import sync_playwright
    processed = 0
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
        board_args = _board_args(page)

        for job in targets:
            if args.limit and processed >= args.limit:
                break
            if not args.rescan and _is_done(records.get(job, {})):
                continue
            records[job] = process_one(page, context, job, board_args)
            processed += 1
            log.info("  %d  %s -> %s", processed, job, records[job].get("status"))
            if processed % 25 == 0:
                save_progress(records)
            if args.delay:
                time.sleep(args.delay)

        browser.close()

    save_progress(records)
    out = write_workbook(records, dwg, Path(args.out))
    by_status: Dict[str, int] = {}
    for r in records.values():
        by_status[r.get("status", "?")] = by_status.get(r.get("status", "?"), 0) + 1
    log.info("Done: processed %d this run; store has %d job(s).", processed, len(records))
    log.info("  status breakdown: %s", ", ".join(f"{k}={v}" for k, v in sorted(by_status.items())))
    if by_status.get("lookup-unconfirmed"):
        log.info("  %d job(s) need the off-board lookup wired in (see the SEAM in this file).",
                 by_status["lookup-unconfirmed"])
    log.info("  Wrote %s", out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
