"""Discovery helper: find out how cbcinsider exposes a job's SALES ORDER.

Run this ONCE, by hand, on the machine that already has a saved session
(cbc_session.json from login.py). It opens your dispatch queue, inspects the
FIRST order row, and reports exactly how the sales order is reachable — so we
can tell whether it's:

    OPTION 1  a direct URL per job   (e.g. OrderView.aspx?ord=12345)
    OPTION 2  a click that downloads a PDF
    OPTION 3  a click that opens an HTML order page (no PDF)

    python discover_sales_order.py

This does NOT change anything on cbcinsider — it only VIEWS one order, exactly
like clicking it yourself. A browser window opens so you can watch.

Everything it learns is printed to the console AND saved under
OUTPUT_DIR/so_discovery/ (HTML dumps + screenshots). When it finishes, paste
the console output back to me and attach:
    - so_discovery/first_container.html
    - so_discovery/result_page.html   (if it was created)
    - the screenshots
and I'll know exactly how to build the downloader.
"""
from __future__ import annotations

import logging

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

from config import CBC_URL, CBC_QUEUE_URL, STORAGE_STATE_PATH, OUTPUT_DIR
from scraper import CONTAINER_SELECTOR

logging.basicConfig(level=logging.INFO, format="%(message)s")
log = logging.getLogger("discover")

OUT = OUTPUT_DIR / "so_discovery"


def _save(name: str, text: str) -> None:
    path = OUT / name
    path.write_text(text or "", encoding="utf-8")
    log.info("  saved %s", path)


def _dump_clickables(container) -> None:
    """Print every <a> and every onclick element in the first order row — this
    alone often answers the question (a job-number <a href=...> = OPTION 1)."""
    log.info("\n--- Links & clickables in the FIRST order row ---")

    anchors = container.locator("a").all()
    log.info("Found %d <a> tag(s):", len(anchors))
    for i, a in enumerate(anchors):
        try:
            text = (a.inner_text() or "").strip().replace("\n", " ")
            log.info("  [a %d] text=%r id=%r target=%r",
                     i, text[:50], a.get_attribute("id"), a.get_attribute("target"))
            log.info("         href   = %r", a.get_attribute("href"))
            onclick = a.get_attribute("onclick")
            if onclick:
                log.info("         onclick= %r", onclick[:200])
        except Exception as e:  # noqa: BLE001 - discovery script, keep going
            log.info("  [a %d] (could not read: %s)", i, e)

    others = container.locator("[onclick]").all()
    log.info("\nFound %d element(s) with an onclick (non-anchor included):", len(others))
    for i, el in enumerate(others):
        try:
            tag = el.evaluate("e => e.tagName")
            log.info("  [click %d] <%s> id=%r onclick=%r",
                     i, tag, el.get_attribute("id"), (el.get_attribute("onclick") or "")[:200])
        except Exception as e:  # noqa: BLE001
            log.info("  [click %d] (could not read: %s)", i, e)


def _pick_job_anchor(container):
    """Best guess at the link that opens the order: the anchor in the Job cell
    (cell index 1 of the job row, matching scraper.py's column map). Falls back
    to the first anchor in the row."""
    job_row = container.locator("tr").first
    cells = job_row.locator("td").all()
    if len(cells) > 1:
        a = cells[1].locator("a")
        if a.count() > 0:
            return a.first
    anchors = container.locator("a")
    return anchors.first if anchors.count() > 0 else None


def main() -> None:
    if not STORAGE_STATE_PATH.exists():
        raise SystemExit(
            f"No saved session at {STORAGE_STATE_PATH}. Run `python login.py` first."
        )
    OUT.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # visible so you can watch
        context = browser.new_context(
            storage_state=str(STORAGE_STATE_PATH),
            accept_downloads=True,
        )
        page = context.new_page()

        target = CBC_QUEUE_URL or CBC_URL
        log.info("Loading queue with saved session: %s", target)
        page.goto(target, wait_until="domcontentloaded", timeout=30000)
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except PlaywrightTimeout:
            pass

        if "login" in page.url.lower() or page.locator('input[type="password"]').count() > 0:
            raise SystemExit("Landed on the login page — session expired. Re-run `python login.py`.")

        try:
            page.wait_for_selector(CONTAINER_SELECTOR, timeout=30000)
        except PlaywrightTimeout:
            raise SystemExit(
                "No order rows found. Set CBC_QUEUE_URL in .env to your exact "
                "dispatch.aspx URL, then re-run."
            )

        containers = page.locator(CONTAINER_SELECTOR).all()
        log.info("Found %d order rows on the board.", len(containers))
        page.screenshot(path=str(OUT / "queue_page.png"), full_page=True)

        first = containers[0]
        # The whole order row's HTML is the single most useful artifact — it
        # shows the markup of the job link / any 'view order' icon.
        _save("first_container.html", first.evaluate("e => e.outerHTML"))

        _dump_clickables(first)

        candidate = _pick_job_anchor(first)
        if candidate is None:
            log.info(
                "\nNo anchor to click in the first row. The order may open via a "
                "row-level onclick (see the onclick dump above) — paste that back "
                "to me and we'll wire it up."
            )
            input("\nDone inspecting. Press Enter to close the browser... ")
            browser.close()
            return

        cand_text = (candidate.inner_text() or "").strip()
        cand_href = candidate.get_attribute("href")
        log.info("\n--- Following the job link: text=%r href=%r ---", cand_text[:50], cand_href)

        # Wire up listeners for ALL three outcomes before clicking: a download
        # (option 2), a popup/new tab (often option 3), or in-place nav.
        downloads = []
        popups = []

        def _wire(pg):
            pg.on("download", downloads.append)

        _wire(page)

        def _on_page(pg):
            popups.append(pg)
            _wire(pg)

        context.on("page", _on_page)

        before_url = page.url
        try:
            candidate.click()
        except Exception as e:  # noqa: BLE001
            log.info("Click raised (%s) — it may have triggered a download/nav anyway.", e)
        page.wait_for_timeout(6000)  # let whatever it does happen

        log.info("\n========== RESULT ==========")
        if downloads:
            d = downloads[0]
            dest = OUT / (d.suggested_filename or "sales_order.pdf")
            try:
                d.save_as(str(dest))
                saved = str(dest)
            except Exception as e:  # noqa: BLE001
                saved = f"(could not save: {e})"
            log.info("Clicking triggered a DOWNLOAD  ->  this is OPTION 2 (easy).")
            log.info("  suggested_filename = %r", d.suggested_filename)
            log.info("  download url       = %r", d.url)
            log.info("  saved to           = %s", saved)
        elif popups:
            pg = popups[-1]
            try:
                pg.wait_for_load_state("domcontentloaded", timeout=10000)
            except PlaywrightTimeout:
                pass
            is_pdf = (pg.url or "").lower().endswith(".pdf")
            log.info("Clicking opened a NEW TAB/POPUP.")
            log.info("  popup url = %r", pg.url)
            log.info("  looks like a direct PDF? %s", is_pdf)
            try:
                pg.screenshot(path=str(OUT / "result_page.png"), full_page=True)
                _save("result_page.html", pg.content())
            except Exception as e:  # noqa: BLE001
                log.info("  (couldn't capture popup contents: %s — likely a PDF viewer)", e)
            log.info("  => If it's a .pdf URL we can fetch it directly with the session (OPTION 1).")
            log.info("     If it's an HTML order page, that's OPTION 3 (scrape or render to PDF).")
        else:
            after_url = page.url
            if after_url != before_url:
                log.info("The page NAVIGATED in place  ->  likely OPTION 1 or 3.")
                log.info("  before = %r", before_url)
                log.info("  after  = %r", after_url)
            else:
                log.info("No download, no popup, no navigation within 6s.")
                log.info("  It may load inline via AJAX, or need a different trigger.")
            try:
                page.screenshot(path=str(OUT / "result_page.png"), full_page=True)
                _save("result_page.html", page.content())
            except Exception as e:  # noqa: BLE001
                log.info("  (couldn't capture page: %s)", e)

        log.info("\n========== NEXT STEPS ==========")
        log.info("Everything was saved under: %s", OUT)
        log.info("Paste the console output above back to me, and attach:")
        log.info("  - first_container.html")
        log.info("  - result_page.html   (if present)")
        log.info("  - the .png screenshots")
        log.info("Then I'll build the real downloader to match.")

        input("\nPress Enter to close the browser... ")
        browser.close()


if __name__ == "__main__":
    main()
