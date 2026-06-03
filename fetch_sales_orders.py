"""Fetch sales orders for the whole board, in parallel — validation harness.

For every order on the dispatch board this:
  - opens the order's detail modal (the slow ~30s NotesPanel load),
  - reads the CBC_SalesOrder revision  ->  CO# = rev - 1  (CO#1 = rev 2),
  - downloads the latest Sales Order pdf into
        SALES_ORDER_DIR/<job>/<job> - Sales Order CO#N.pdf
    (skipping any revision already on disk), and
  - reports each job's CO# plus total timing.

Concurrency is SO_CONCURRENCY worker tabs sharing the one saved session, so we
can see empirically whether the server lets the ~30s loads overlap or
serializes them.

    python fetch_sales_orders.py          # whole board
    python fetch_sales_orders.py 16       # only the first 16 orders (quick test)

Read-only against cbcinsider — it just views orders and downloads their SO.
"""
from __future__ import annotations

import asyncio
import logging
import re
import sys
import time
from urllib.parse import urlparse, parse_qs, urljoin

from playwright.async_api import async_playwright, TimeoutError as PWTimeout

from config import CBC_URL, CBC_QUEUE_URL, STORAGE_STATE_PATH, SALES_ORDER_DIR, SO_CONCURRENCY
from scraper import CONTAINER_SELECTOR

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger("so-fetch")

# pid looks like "CBC_SalesOrder-32045-3-LATEST" -> type / id / revision / tag.
PID_RE = re.compile(r"^(?P<type>.+?)-(?P<id>\d+)-(?P<rev>\d+)-(?P<tag>[A-Za-z0-9]+)$")


def _jobnum(args_js: str) -> str:
    first = args_js.split(",", 1)[0].strip().strip("'\"")
    return first.split("-", 1)[0]


def _parse_doc(href: str) -> dict:
    q = parse_qs(urlparse(href).query)
    pid = q.get("pid", [""])[0]
    fn = q.get("fn", [""])[0]
    m = PID_RE.match(pid)
    return {"pid": pid, "fn": fn, "type": m["type"] if m else pid, "rev": int(m["rev"]) if m else None}


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


async def _open_board(context, dispatch_url):
    page = await context.new_page()
    await page.goto(dispatch_url, wait_until="domcontentloaded", timeout=45000)
    try:
        await page.wait_for_load_state("networkidle", timeout=15000)
    except PWTimeout:
        pass
    await page.wait_for_selector(CONTAINER_SELECTOR, timeout=45000)
    return page


async def _args_map(page) -> dict:
    """jobnum -> loadDetail args, read from this page's own containers (indices
    are page-specific, so each worker resolves against its own board)."""
    out: dict = {}
    for c in await page.locator(CONTAINER_SELECTOR).all():
        m = re.search(r"loadDetail\((.*?)\)", await c.get_attribute("onclick") or "")
        if m:
            a = m.group(1).strip()
            out[_jobnum(a)] = a
    return out


async def _process_job(page, context, job: str, args_js: str) -> dict:
    res = {"job": job, "rev": None, "co": "", "status": ""}
    await page.evaluate(_trigger_js(args_js))

    # The new job's own documents are named with its job number — waiting on that
    # both confirms the load finished AND distinguishes it from the previous job.
    link = page.locator("#modalDetail a").filter(has_text=re.compile(re.escape(job)))
    try:
        await link.first.wait_for(state="attached", timeout=90000)
    except PWTimeout:
        res["status"] = "no docs (timeout)"
        return res

    docs = []
    for a in await page.locator("#modalDetail a").all():
        href = await a.get_attribute("href") or ""
        if "downloaddoc.aspx" in href.lower():
            docs.append((href, _parse_doc(href)))

    sos = [(h, d) for h, d in docs
           if "salesorder" in (d["type"] or "").lower() and "sales order" in (d["fn"] or "").lower()]
    if not sos:
        res["status"] = "no sales order doc"
        return res

    href, doc = max(sos, key=lambda hd: hd[1]["rev"] or 0)
    rev = doc["rev"]
    res["rev"] = rev
    res["co"] = f"CO#{rev - 1}" if rev and rev > 1 else "(original)"

    dest = SALES_ORDER_DIR / job / _so_filename(job, rev)
    if dest.exists():
        res["status"] = "have it"
        return res
    try:
        dest.parent.mkdir(parents=True, exist_ok=True)
        resp = await context.request.get(urljoin(page.url, href))
        body = await resp.body()
        dest.write_bytes(body)
        res["status"] = f"downloaded {len(body)}b"
    except Exception as e:  # noqa: BLE001
        res["status"] = f"download FAILED: {e}"
    return res


async def _worker(name, context, dispatch_url, queue, results):
    page = await _open_board(context, dispatch_url)
    amap = await _args_map(page)
    while True:
        try:
            job = queue.get_nowait()
        except asyncio.QueueEmpty:
            break
        try:
            args_js = amap.get(job)
            if not args_js:
                results.append({"job": job, "rev": None, "co": "", "status": "not on board", "secs": 0})
            else:
                t0 = time.monotonic()
                r = await _process_job(page, context, job, args_js)
                r["secs"] = round(time.monotonic() - t0, 1)
                results.append(r)
                log.info("[%s] %s  %-11s %-20s %4.1fs", name, job, r["co"], r["status"], r["secs"])
            await page.evaluate("() => window.jQuery && jQuery('#modalDetail').modal('hide')")
            await page.wait_for_timeout(400)
        except Exception as e:  # noqa: BLE001
            results.append({"job": job, "rev": None, "co": "", "status": f"ERROR: {e}", "secs": 0})
            log.warning("[%s] %s ERROR: %s", name, job, e)
        finally:
            queue.task_done()
    await page.close()


async def _amain(limit: int | None):
    if not STORAGE_STATE_PATH.exists():
        raise SystemExit(f"No saved session at {STORAGE_STATE_PATH}. Run `python login.py` first.")
    dispatch_url = CBC_QUEUE_URL or CBC_URL

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(storage_state=str(STORAGE_STATE_PATH), accept_downloads=True)

        board = await _open_board(context, dispatch_url)
        all_jobs = list((await _args_map(board)).keys())
        await board.close()
        total = len(all_jobs)
        jobs = all_jobs[:limit] if limit else all_jobs
        n_workers = min(SO_CONCURRENCY, len(jobs)) or 1
        log.info("Board has %d orders; fetching %d with %d workers.", total, len(jobs), n_workers)
        log.info("Archiving under: %s", SALES_ORDER_DIR)

        queue: asyncio.Queue = asyncio.Queue()
        for j in jobs:
            queue.put_nowait(j)

        results: list = []
        t0 = time.monotonic()
        await asyncio.gather(*[
            asyncio.create_task(_worker(f"w{i + 1}", context, dispatch_url, queue, results))
            for i in range(n_workers)
        ])
        elapsed = time.monotonic() - t0
        await browser.close()

        downloaded = sum(1 for r in results if str(r["status"]).startswith("downloaded"))
        have = sum(1 for r in results if r["status"] == "have it")
        cos = [r for r in results if r.get("rev") and r["rev"] > 1]
        # "no sales order doc" is expected for some order types (e.g. HDX /
        # Michael's) — count it separately, not as a problem.
        no_so = [r for r in results if r["status"] == "no sales order doc"]
        problems = [r for r in results if any(k in str(r["status"]) for k in ("FAIL", "ERROR", "timeout", "not on"))]
        log.info("=" * 64)
        log.info("Done: %d orders in %.0fs (%.1fs/order wall-clock, %d workers).",
                 len(results), elapsed, elapsed / max(1, len(results)), n_workers)
        log.info("  %d downloaded, %d already had, %d with a change order, %d no-SO (e.g. HDX), %d problems.",
                 downloaded, have, len(cos), len(no_so), len(problems))
        if cos:
            log.info("  Change orders: %s", ", ".join(f"{r['job']}={r['co']}" for r in cos))
        if problems:
            for r in problems:
                log.info("    PROBLEM %s: %s", r["job"], r["status"])


def main():
    limit = None
    if len(sys.argv) > 1:
        try:
            limit = int(sys.argv[1])
        except ValueError:
            pass
    asyncio.run(_amain(limit))


if __name__ == "__main__":
    main()
