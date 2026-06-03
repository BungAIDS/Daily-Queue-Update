"""Sales-order enrichment for the daily run.

For every job on the board this:
  1. opens its detail modal (in parallel across SO_CONCURRENCY tabs) and reads
     the CBC_SalesOrder revision  ->  CO# = rev - 1  (CO#1 = rev 2),
  2. downloads the latest Sales Order pdf into SALES_ORDER_DIR/<job>/ if that
     revision isn't already on disk (keeping older revisions),
  3. parses Design / Size / Arrangement + the change-order history out of the
     pdf, and
  4. looks up the job's AutoCAD folder, which also yields its type.

`enrich_with_sales_orders(jobs)` mutates each job dict in place, adding:
    co_number      int   (0 = no change orders)
    co_history     list[str]  (the "CO#N date initials - description" lines)
    so_design_desc str   (e.g. "Vaneaxial Belt Drive")
    so_size        str
    so_arrangement str
    so_pdf         str   (path to the latest SO pdf, or "")
    job_type       str   (e.g. "AXIAL" / "GENERAL LINE", or "")
    job_folder     str   (AutoCAD folder if found, else the SO archive folder)

It is resilient: any job that errors, has no sales order (e.g. HDX), or whose
folder isn't found simply gets blank/zero fields rather than failing the run.
"""
from __future__ import annotations

import asyncio
import logging
import re
from pathlib import Path
from typing import Any, Dict, List
from urllib.parse import urlparse, parse_qs, urljoin

from playwright.async_api import async_playwright, TimeoutError as PWTimeout

from config import (
    CBC_URL, CBC_QUEUE_URL, STORAGE_STATE_PATH,
    SALES_ORDER_DIR, SO_CONCURRENCY, AUTOCAD_JOBS_DIR,
)
from scraper import CONTAINER_SELECTOR

log = logging.getLogger(__name__)

PID_RE = re.compile(r"^(?P<type>.+?)-(?P<id>\d+)-(?P<rev>\d+)-(?P<tag>[A-Za-z0-9]+)$")
CO_START = re.compile(r"^\s*C\s*/?\s*O\s*#?\s*\d", re.I)
DESIGN_HDR = re.compile(r"^\s*Design\s+(\S+)\s*(.*)$")
SPEC_CELL = re.compile(r"^(Design|Size|Arrangement)\b\s*(.*)$", re.I)


# --------------------------------------------------------------------------- #
# AutoCAD folder / job-type lookup                                            #
# --------------------------------------------------------------------------- #
def _build_autocad_index() -> Dict[str, Dict[str, Any]]:
    """Walk AUTOCAD_JOBS_DIR/<type>/<intermediate>/<job> once, mapping each job
    number to its type + folder. Returns {} if the drive isn't reachable."""
    index: Dict[str, Dict[str, Any]] = {}
    root = AUTOCAD_JOBS_DIR
    try:
        if not root.exists():
            log.warning("AutoCAD jobs root not reachable: %s (folder links disabled)", root)
            return index
        for type_dir in root.iterdir():
            if not type_dir.is_dir():
                continue
            for inter_dir in type_dir.iterdir():
                if not inter_dir.is_dir():
                    continue
                for job_dir in inter_dir.iterdir():
                    if job_dir.is_dir():
                        # first 6+ digit token is the job number
                        m = re.match(r"(\d{4,})", job_dir.name)
                        if m:
                            index.setdefault(m.group(1), {"type": type_dir.name, "path": job_dir})
        log.info("Indexed %d AutoCAD job folders under %s", len(index), root)
    except OSError as e:
        log.warning("Could not index AutoCAD folders (%s); folder links disabled", e)
    return index


# --------------------------------------------------------------------------- #
# PDF parsing                                                                 #
# --------------------------------------------------------------------------- #
def _recon_lines(page, x_tol: float = 1.5) -> List[str]:
    """Rebuild text lines from word positions so spaces survive (plain
    extraction glues the Notes text together)."""
    words = page.extract_words(x_tolerance=x_tol, keep_blank_chars=False, use_text_flow=False)
    rows: Dict[int, list] = {}
    for w in words:
        rows.setdefault(round(w["top"]), []).append(w)
    out = []
    for top in sorted(rows):
        ws = sorted(rows[top], key=lambda w: w["x0"])
        out.append(" ".join(w["text"] for w in ws))
    return out


def _spec_from_tables(tables) -> Dict[str, str]:
    fields: Dict[str, str] = {}
    for table in tables or []:
        for row in table:
            for cell in row:
                if not cell:
                    continue
                m = SPEC_CELL.match(cell.replace("\n", " ").strip())
                if m:
                    label, val = m.group(1).title(), m.group(2).strip()
                    if label not in fields and val:  # first wins (vaneaxial repeats "Design")
                        fields[label] = val
    return fields


def parse_sales_order_pdf(path: str | Path) -> Dict[str, Any]:
    """Pull Design/Size/Arrangement + change-order history out of an SO pdf."""
    res = {"design_desc": "", "size": "", "arrangement": "", "header_co": None, "co_history": []}
    try:
        import pdfplumber
    except ImportError:
        log.warning("pdfplumber not installed; cannot parse SO pdfs (pip install pdfplumber)")
        return res
    try:
        with pdfplumber.open(str(path)) as pdf:
            p1 = pdf.pages[0]
            for ln in (p1.extract_text() or "").splitlines()[:8]:
                if res["header_co"] is None:
                    m = re.search(r"CO\s*#\s*(\d+)", ln)
                    if m:
                        res["header_co"] = int(m.group(1))
                d = DESIGN_HDR.match(ln)
                if d and not res["design_desc"]:
                    res["design_desc"] = d.group(2).strip()
            spec = _spec_from_tables(p1.extract_tables())
            res["size"] = spec.get("Size", "")
            res["arrangement"] = spec.get("Arrangement", "") or "N/A"
            for page in pdf.pages:
                for ln in _recon_lines(page):
                    if CO_START.match(ln):
                        res["co_history"].append(ln.strip())
    except Exception as e:  # noqa: BLE001 - never let a bad pdf fail the run
        log.warning("Could not parse SO pdf %s: %s", path, e)
    return res


# --------------------------------------------------------------------------- #
# Parallel fetch of each job's sales order                                    #
# --------------------------------------------------------------------------- #
_STATIC = (".js", ".css", ".png", ".gif", ".jpg", ".jpeg", ".svg", ".woff", ".woff2", ".ico")


def _jobnum(args_js: str) -> str:
    return args_js.split(",", 1)[0].strip().strip("'\"").split("-", 1)[0]


def _parse_doc(href: str) -> Dict[str, Any]:
    q = parse_qs(urlparse(href).query)
    pid, fn = q.get("pid", [""])[0], q.get("fn", [""])[0]
    m = PID_RE.match(pid)
    return {"fn": fn, "type": m["type"] if m else pid, "rev": int(m["rev"]) if m else None}


def _so_filename(job: str, rev: int | None) -> str:
    if rev and rev > 1:
        return f"{job} - Sales Order CO#{rev - 1}.pdf"
    return f"{job} - Sales Order (original).pdf"


def _trigger_js(args_js: str) -> str:
    return f"""() => {{
        if (window.jQuery) {{
            jQuery('#modalDetail').off('show.bs.modal')
                .on('show.bs.modal', function () {{ loadDetail({args_js}); }})
                .modal('show');
        }} else {{ loadDetail({args_js}); }}
    }}"""


async def _open_board(context, url):
    page = await context.new_page()
    await page.goto(url, wait_until="domcontentloaded", timeout=45000)
    try:
        await page.wait_for_load_state("networkidle", timeout=15000)
    except PWTimeout:
        pass
    await page.wait_for_selector(CONTAINER_SELECTOR, timeout=45000)
    return page


async def _args_map(page) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for c in await page.locator(CONTAINER_SELECTOR).all():
        m = re.search(r"loadDetail\((.*?)\)", await c.get_attribute("onclick") or "")
        if m:
            out[_jobnum(m.group(1).strip())] = m.group(1).strip()
    return out


async def _process_job(page, context, job: str, args_js: str) -> Dict[str, Any]:
    res = {"rev": None, "pdf_path": None}
    await page.evaluate(_trigger_js(args_js))
    link = page.locator("#modalDetail a").filter(has_text=re.compile(re.escape(job)))
    try:
        await link.first.wait_for(state="attached", timeout=90000)
    except PWTimeout:
        return res  # no docs for this job (e.g. HDX)

    sos = []
    for a in await page.locator("#modalDetail a").all():
        href = await a.get_attribute("href") or ""
        if "downloaddoc.aspx" in href.lower():
            d = _parse_doc(href)
            if "salesorder" in (d["type"] or "").lower() and "sales order" in (d["fn"] or "").lower():
                sos.append((href, d))
    if not sos:
        return res

    href, doc = max(sos, key=lambda hd: hd[1]["rev"] or 0)
    rev = doc["rev"]
    res["rev"] = rev
    dest = SALES_ORDER_DIR / job / _so_filename(job, rev)
    if dest.exists():
        res["pdf_path"] = str(dest)
        return res
    try:
        dest.parent.mkdir(parents=True, exist_ok=True)
        resp = await context.request.get(urljoin(page.url, href))
        dest.write_bytes(await resp.body())
        res["pdf_path"] = str(dest)
    except Exception as e:  # noqa: BLE001
        log.warning("SO download failed for %s: %s", job, e)
    return res


async def _worker(context, url, queue, results):
    page = await _open_board(context, url)
    amap = await _args_map(page)
    while True:
        try:
            job = queue.get_nowait()
        except asyncio.QueueEmpty:
            break
        try:
            args_js = amap.get(job)
            results[job] = await _process_job(page, context, job, args_js) if args_js else {"rev": None, "pdf_path": None}
        except Exception as e:  # noqa: BLE001
            log.warning("SO fetch error for %s: %s", job, e)
            results[job] = {"rev": None, "pdf_path": None}
        finally:
            await page.evaluate("() => window.jQuery && jQuery('#modalDetail').modal('hide')")
            await page.wait_for_timeout(300)
            queue.task_done()
    await page.close()


async def _afetch_all(job_numbers: List[str]) -> Dict[str, Dict[str, Any]]:
    if not STORAGE_STATE_PATH.exists():
        raise RuntimeError(f"No saved session at {STORAGE_STATE_PATH}. Run `python login.py`.")
    url = CBC_QUEUE_URL or CBC_URL
    results: Dict[str, Dict[str, Any]] = {}
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(storage_state=str(STORAGE_STATE_PATH), accept_downloads=True)
        queue: asyncio.Queue = asyncio.Queue()
        for j in job_numbers:
            queue.put_nowait(j)
        n = min(SO_CONCURRENCY, len(job_numbers)) or 1
        await asyncio.gather(*[asyncio.create_task(_worker(context, url, queue, results)) for _ in range(n)])
        await browser.close()
    return results


# --------------------------------------------------------------------------- #
# Public entry point                                                          #
# --------------------------------------------------------------------------- #
def enrich_with_sales_orders(jobs: List[Dict[str, Any]]) -> None:
    """Mutate `jobs` in place, attaching sales-order + folder fields (see module
    docstring). Opens every job's detail modal in parallel — the slow step."""
    by_job = {j["job"]: j for j in jobs if j.get("job")}
    if not by_job:
        return

    index = _build_autocad_index()
    log.info("Fetching sales orders for %d jobs (%d parallel)...", len(by_job), SO_CONCURRENCY)
    so_results = asyncio.run(_afetch_all(list(by_job.keys())))

    n_co = n_dl = 0
    for jn, j in by_job.items():
        r = so_results.get(jn, {})
        rev = r.get("rev")
        j["co_number"] = (rev - 1) if rev and rev > 1 else 0
        if j["co_number"]:
            n_co += 1

        pdf = r.get("pdf_path")
        parsed = parse_sales_order_pdf(pdf) if pdf else {}
        j["co_history"] = parsed.get("co_history", [])
        j["so_design_desc"] = parsed.get("design_desc", "")
        j["so_size"] = parsed.get("size", "")
        j["so_arrangement"] = parsed.get("arrangement", "")
        j["so_pdf"] = pdf or ""
        if pdf:
            n_dl += 1

        info = index.get(jn)
        if info:
            j["job_type"] = info["type"]
            j["job_folder"] = str(info["path"])
        else:
            j["job_type"] = ""
            # Fall back to the SO archive folder when there's no AutoCAD folder yet.
            j["job_folder"] = str(SALES_ORDER_DIR / jn) if pdf else ""

    log.info("Sales orders: %d jobs have a SO, %d currently at a change order.", n_dl, n_co)
