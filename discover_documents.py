"""Discovery — see every document on a job, and how to reach an OLD order.

Two modes:

  python discover_documents.py [job#]
      Open the job's detail (it must be on the board), list EVERY document with
      its pid type / revision, flag the CBC_SalesOrder and CBC_DriveRun, and
      download both as samples. This confirms how the construction/drive run
      shows up and what its pid type is.

  python discover_documents.py --probe <old-job#>
      For backfill: figure out how to open an order that is NOT on the board.
      Prints a real loadDetail(...) signature from a current row, lists the
      page's search/lookup controls, then tries to search for the old job and
      reports whether its documents surface.

Read-only except for downloading the two sample PDFs. Paste the console output
back so we can pin down the exact commands.
"""
from __future__ import annotations

import logging
import re
import sys
from urllib.parse import urljoin

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

from config import CBC_URL, CBC_QUEUE_URL, STORAGE_STATE_PATH, OUTPUT_DIR
from scraper import CONTAINER_SELECTOR
# Reuse the SAME selection logic the daily run uses, so discovery reflects
# exactly what enrichment will pick.
from sales_orders import (
    _parse_doc, _norm_type, _latest_of_type, _trigger_js,
    SO_TYPE, DRIVE_RUN_TYPE,
)

logging.basicConfig(level=logging.INFO, format="%(message)s")
log = logging.getLogger("discover")

OUT = OUTPUT_DIR / "doc_discovery"


def _jobnum(args_js: str) -> str:
    return args_js.split(",", 1)[0].strip().strip("'\"").split("-", 1)[0]


def _board_args(page) -> list[tuple[str, str]]:
    """(job, loadDetail-args) for every container currently on the board."""
    out = []
    for c in page.locator(CONTAINER_SELECTOR).all():
        m = re.search(r"loadDetail\((.*?)\)", c.get_attribute("onclick") or "")
        if m:
            out.append((_jobnum(m.group(1).strip()), m.group(1).strip()))
    return out


def _open_detail(page, jn: str, args_js: str) -> bool:
    """Trigger the detail modal for a job and wait for its documents."""
    page.evaluate(_trigger_js(args_js))
    try:
        page.locator("#modalDetail a").filter(
            has_text=re.compile(re.escape(jn))).first.wait_for(state="attached", timeout=120000)
        return True
    except PlaywrightTimeout:
        return False


def _collect_docs(page) -> list[dict]:
    docs = []
    for a in page.locator("#modalDetail a").all():
        href = a.get_attribute("href") or ""
        if "downloaddoc.aspx" in href.lower():
            d = _parse_doc(href)
            d["href"] = href
            docs.append(d)
    return docs


def _report_docs(jn: str, docs: list[dict]) -> None:
    log.info("\n--- Documents for job %s (%d) ---", jn, len(docs))
    for d in docs:
        rev = f"rev {d['rev']}" if d["rev"] is not None else "rev ?"
        flag = ""
        if _norm_type(d["type"]) == _norm_type(SO_TYPE):
            flag = "   <== SALES ORDER"
        elif _norm_type(d["type"]) == _norm_type(DRIVE_RUN_TYPE):
            flag = "   <== DRIVE RUN (construction)"
        log.info("  %-26s %-7s  %s%s", d["type"], rev, (d["fn"] or "")[:48], flag)


def _download(context, page, doc: dict, label: str) -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    dest = OUT / f"{label}.pdf"
    try:
        resp = context.request.get(urljoin(page.url, doc["href"]))
        dest.write_bytes(resp.body())
        log.info("  downloaded %-12s -> %s (%d bytes)", label, dest, len(resp.body()))
    except Exception as e:  # noqa: BLE001
        log.info("  %s download failed: %s", label, e)


def run_list(want_job: str | None) -> None:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(storage_state=str(STORAGE_STATE_PATH), accept_downloads=True)
        page = context.new_page()
        page.goto(CBC_QUEUE_URL or CBC_URL, wait_until="domcontentloaded", timeout=30000)
        if "login" in page.url.lower() or page.locator('input[type="password"]').count() > 0:
            raise SystemExit("Landed on login — session expired. Re-run `python login.py`.")
        page.wait_for_selector(CONTAINER_SELECTOR, timeout=30000)

        rows = _board_args(page)
        log.info("Found %d order rows.", len(rows))
        chosen = next(((jn, a) for jn, a in rows if want_job is None or jn == want_job), None)
        if chosen is None:
            raise SystemExit(f"Job {want_job!r} not on the board. Try one that's in the queue, "
                             "or use --probe to look up an old order.")
        jn, args = chosen
        log.info("\nOpening detail for job %s ...", jn)
        if not _open_detail(page, jn, args):
            log.info("Documents didn't appear within 120s.")

        docs = _collect_docs(page)
        _report_docs(jn, docs)

        so = _latest_of_type([(d["href"], d) for d in docs], SO_TYPE)
        dr = _latest_of_type([(d["href"], d) for d in docs], DRIVE_RUN_TYPE)
        log.info("\n--- Summary ---")
        log.info("  Sales Order : %s", f"rev {so[1]['rev']}  {so[1]['fn']}" if so else "NONE FOUND")
        log.info("  Drive Run   : %s", f"rev {dr[1]['rev']}  {dr[1]['fn']}" if dr else "none (not a custom order)")
        if so:
            _download(context, page, so[1], f"{jn}_sales_order")
        if dr:
            _download(context, page, dr[1], f"{jn}_drive_run")
            log.info("\n  Now dump the drive run to see its fields:")
            log.info("    python dump_pdf.py \"%s\"", OUT / f"{jn}_drive_run.pdf")

        input("\nPress Enter to close the browser... ")
        browser.close()


def run_probe(old_job: str) -> None:
    """Try to reach an order that's not on the board — for the backfill tool."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(storage_state=str(STORAGE_STATE_PATH), accept_downloads=True)
        page = context.new_page()
        page.goto(CBC_QUEUE_URL or CBC_URL, wait_until="domcontentloaded", timeout=30000)
        if "login" in page.url.lower() or page.locator('input[type="password"]').count() > 0:
            raise SystemExit("Landed on login — session expired. Re-run `python login.py`.")
        page.wait_for_selector(CONTAINER_SELECTOR, timeout=30000)

        # 1) The exact loadDetail(...) signature from a live row — the off-board
        #    lookup probably calls the same function with different args.
        rows = _board_args(page)
        log.info("=== loadDetail signature (from a live board row) ===")
        if rows:
            sample = page.locator(CONTAINER_SELECTOR).first.get_attribute("onclick") or ""
            log.info("  sample onclick: %s", sample[:240])
            log.info("  parsed job=%s  args=%s", rows[0][0], rows[0][1])

        # 2) Any search / lookup controls on the page.
        log.info("\n=== candidate search/lookup controls ===")
        inputs = page.locator(
            "input[type=text], input[type=search], input:not([type]), select").all()
        for el in inputs:
            ident = el.get_attribute("id") or el.get_attribute("name") or ""
            ph = el.get_attribute("placeholder") or ""
            if any(k in (ident + ph).lower() for k in ("search", "job", "order", "find", "filter", "lookup")):
                log.info("  <%s id/name=%r placeholder=%r>", el.evaluate("e => e.tagName"), ident, ph)
        log.info("  (if nothing useful here, the order detail may live at its own URL — "
                 "open an old order in the browser and copy the address bar)")

        # 3) Best-effort: type the old job into the first search-y box and submit.
        box = page.locator("input#MainContent_txtSearch, input[id*=earch], input[placeholder*=earch]").first
        if box.count():
            log.info("\n=== trying to search for %s ===", old_job)
            try:
                box.fill(old_job)
                box.press("Enter")
                page.wait_for_timeout(4000)
                rows2 = _board_args(page)
                hit = next((a for jn, a in rows2 if jn == old_job), None)
                if hit and _open_detail(page, old_job, hit):
                    _report_docs(old_job, _collect_docs(page))
                    log.info("  ^ search worked — backfill can drive this box.")
                else:
                    log.info("  search did not surface %s on the board.", old_job)
            except Exception as e:  # noqa: BLE001
                log.info("  search attempt errored: %s", e)
        else:
            log.info("\n  No obvious search box found.")

        log.info("\nPaste everything above back so we can wire the exact old-order lookup "
                 "into backfill_orders.open_order_detail().")
        input("\nPress Enter to close the browser... ")
        browser.close()


def main() -> None:
    if not STORAGE_STATE_PATH.exists():
        raise SystemExit(f"No saved session at {STORAGE_STATE_PATH}. Run `python login.py` first.")
    args = sys.argv[1:]
    if args and args[0] == "--probe":
        if len(args) < 2:
            raise SystemExit("Usage: python discover_documents.py --probe <old-job#>")
        run_probe(args[1].strip())
    else:
        run_list(args[0].strip() if args else None)


if __name__ == "__main__":
    main()
