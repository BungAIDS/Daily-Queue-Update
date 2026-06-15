"""Discovery — see every document on a job, and how to reach an OLD order.

Two modes:

  python discover_documents.py [job#]
      Open the job's detail, list EVERY document with its pid type / revision,
      flag the Sales Order and the quote/drive run, and download both as
      samples. The job is opened from the board if it's there; otherwise it's
      looked up through the queue's search box (the same path the backfill
      uses), so an old job with a quote run can be inspected too. This confirms
      how the construction/quote run shows up and what its pid type is.

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
from pathlib import Path
from urllib.parse import urljoin

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

from config import CBC_URL, CBC_QUEUE_URL, STORAGE_STATE_PATH, OUTPUT_DIR
from scraper import CONTAINER_SELECTOR
# Reuse the SAME selection logic the daily run uses, so discovery reflects
# exactly what enrichment will pick.
from sales_orders import (
    _parse_doc, _norm_type, _latest_of_type, _run_docs, _is_run_name, _trigger_js,
    _find_autocad_folders, _run_files_in_folder, SO_TYPE,
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
    run_hrefs = {h for h, _ in _run_docs([(d["href"], d) for d in docs])}
    log.info("\n--- Documents for job %s (%d) ---", jn, len(docs))
    for d in docs:
        rev = f"rev {d['rev']}" if d["rev"] is not None else "rev ?"
        flag = ""
        if _norm_type(d["type"]) == _norm_type(SO_TYPE):
            flag = "   <== SALES ORDER"
        elif d["href"] in run_hrefs:
            how = "file name" if _is_run_name(d["fn"]) else "pid type"
            flag = f"   <== QUOTE RUN (by {how})"
        log.info("  %-26s %-7s  %s%s", d["type"], rev, (d["fn"] or "")[:48], flag)


def _download(context, page, doc: dict, label: str) -> Path | None:
    OUT.mkdir(parents=True, exist_ok=True)
    # Keep the document's own extension — quote runs come as .txt/.xlsx/.rtf too.
    ext = Path(doc.get("fn") or "").suffix or ".pdf"
    dest = OUT / f"{label}{ext}"
    try:
        resp = context.request.get(urljoin(page.url, doc["href"]))
        dest.write_bytes(resp.body())
        log.info("  downloaded %-12s -> %s (%d bytes)", label, dest, len(resp.body()))
        return dest
    except Exception as e:  # noqa: BLE001
        log.info("  %s download failed: %s", label, e)
        return None


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
        if chosen is not None:
            jn, args = chosen
            log.info("\nOpening detail for job %s (on the board) ...", jn)
            if not _open_detail(page, jn, args):
                log.info("Documents didn't appear within 120s.")
        elif want_job is None:
            raise SystemExit("No orders on the board to inspect.")
        else:
            # Not on the board — look it up through the queue's search box, the
            # same path the backfill uses (so an old job with a quote run can be
            # inspected without it being in the queue).
            from backfill_orders import open_order_detail
            jn = want_job
            log.info("\nJob %s isn't on the board — searching for it ...", jn)
            if not open_order_detail(page, jn):
                raise SystemExit(
                    f"Couldn't surface job {jn!r} via the search box. If the queue page has a "
                    "search/find-order field, set CBC_SEARCH_SELECTOR (and CBC_SEARCH_BUTTON if a "
                    "separate button submits) in .env, then retry. `--probe` prints the candidates.")

        docs = _collect_docs(page)
        _report_docs(jn, docs)

        so = _latest_of_type([(d["href"], d) for d in docs], SO_TYPE)
        runs = _run_docs([(d["href"], d) for d in docs])
        log.info("\n--- Summary ---")
        log.info("  Sales Order : %s", f"rev {so[1]['rev']}  {so[1]['fn']}" if so else "NONE FOUND")
        if runs:
            for _, d in runs:
                log.info("  Quote Run   : type %s  rev %s  %s", d["type"], d["rev"], d["fn"])
        else:
            log.info("  Quote Run   : none matched in the documents (if one is listed above, "
                     "add its pid type / file-name pattern to DRIVE_RUN_TYPES / "
                     "DRIVE_RUN_NAME_PATTERNS in .env)")
        if so:
            _download(context, page, so[1], f"{jn}_sales_order")
        for i, (_, d) in enumerate(runs):
            _download(context, page, d, f"{jn}_quote_run" + (f"_{i + 1}" if len(runs) > 1 else ""))
        pdf_i = next((i for i, (_, d) in enumerate(runs)
                      if (d.get("fn") or "").lower().endswith(".pdf")), None)
        if pdf_i is not None:
            name = f"{jn}_quote_run" + (f"_{pdf_i + 1}" if len(runs) > 1 else "") + ".pdf"
            log.info("\n  Now dump the pdf run to see its fields:")
            log.info("    python dump_pdf.py \"%s\"", OUT / name)

        # Some orders only keep the run in their AutoCAD folder — check there too.
        try:
            info = _find_autocad_folders([jn]).get(jn)
            if info:
                hits = _run_files_in_folder(info["path"])
                log.info("\n--- AutoCAD folder check (%s) ---", info["path"])
                for f in hits:
                    log.info("  run file: %s", f)
                if not hits:
                    log.info("  no run-named files found")
            else:
                log.info("\n(AutoCAD folder for %s not found — folder check skipped.)", jn)
        except Exception as e:  # noqa: BLE001 - discovery should still finish
            log.info("\n(AutoCAD folder check failed: %s)", e)

        input("\nPress Enter to close the browser... ")
        browser.close()


def run_probe(old_job: str) -> None:
    """Confirm the backfill's old-order lookup: list the page's text inputs, show
    what auto-detect picks, then run the REAL backfill search path for `old_job`
    and report whether its documents surface."""
    from backfill_orders import find_search_box, open_order_detail
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(storage_state=str(STORAGE_STATE_PATH), accept_downloads=True)
        page = context.new_page()
        page.goto(CBC_QUEUE_URL or CBC_URL, wait_until="domcontentloaded", timeout=30000)
        if "login" in page.url.lower() or page.locator('input[type="password"]').count() > 0:
            raise SystemExit("Landed on login — session expired. Re-run `python login.py`.")
        page.wait_for_selector(CONTAINER_SELECTOR, timeout=30000)

        # 1) Every text input — so you can read off the exact CBC_SEARCH_SELECTOR.
        log.info("=== text inputs on the queue page ===")
        for el in page.locator("input[type=text], input[type=search], input:not([type])").all():
            ident, name = el.get_attribute("id") or "", el.get_attribute("name") or ""
            ph, al = el.get_attribute("placeholder") or "", el.get_attribute("aria-label") or ""
            sel = f"#{ident}" if ident else (f"input[name='{name}']" if name else "input[..]")
            log.info("  %-28s placeholder=%r aria-label=%r", sel, ph, al)

        # 2) What the backfill's auto-detector picks.
        box = find_search_box(page)
        if box is None:
            log.info("\n  Auto-detect found NO search box. Choose the right input above and set its\n"
                     "  CSS selector as CBC_SEARCH_SELECTOR in .env (e.g. #MainContent_txtFindOrder).")
        else:
            log.info("\n  Auto-detect picked: id=%r name=%r placeholder=%r",
                     box.get_attribute("id") or "", box.get_attribute("name") or "",
                     box.get_attribute("placeholder") or "")

        # 3) Run the ACTUAL backfill lookup and report — this is the real path.
        log.info("\n=== running the real backfill lookup for %s ===", old_job)
        if open_order_detail(page, old_job):
            _report_docs(old_job, _collect_docs(page))
            log.info("  ^ SUCCESS — backfill can pull this order. You're ready: python backfill_orders.py")
        else:
            log.info("  Did NOT surface %s. If a box is listed above, set CBC_SEARCH_SELECTOR to it "
                     "(and CBC_SEARCH_BUTTON if a separate button submits), then re-probe.", old_job)

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
