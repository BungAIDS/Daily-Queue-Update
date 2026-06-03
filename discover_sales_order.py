"""Discovery helper #3: find and download a job's SALES ORDER pdf.

What we've learned so far:
  - Clicking a dispatch row runs loadDetail('<job>-<n>', ...), which opens a
    Bootstrap modal (#modalDetail) and loads the detail over an ASP.NET async
    postback (__doPostBackAsync("NotesPanel", jobid)). It's slow (~30s).
  - The loaded modal contains a "Documents" list, and one entry is the sales
    order, named like "421314 - Sales Order.pdf".

This script opens the FIRST order, waits for that Documents list, prints every
document link's href, and DOWNLOADS the "... - Sales Order.pdf" as proof — so
we learn whether the pdf is a direct URL (fast: fetch per job) or a
javascript/handler link (must be clicked).

    python discover_sales_order.py

Read-only: it just opens one order and downloads its sales order, the same as
clicking it yourself. Output is saved under OUTPUT_DIR/so_discovery/.
"""
from __future__ import annotations

import logging
import re
from urllib.parse import urljoin

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

from config import CBC_URL, CBC_QUEUE_URL, STORAGE_STATE_PATH, OUTPUT_DIR
from scraper import CONTAINER_SELECTOR

logging.basicConfig(level=logging.INFO, format="%(message)s")
log = logging.getLogger("discover")

OUT = OUTPUT_DIR / "so_discovery"

_STATIC = (".js", ".css", ".png", ".gif", ".jpg", ".jpeg", ".svg", ".woff", ".woff2", ".ico")
_SALES_ORDER_RE = re.compile(r"sales order", re.I)


def _save(name: str, text: str) -> None:
    (OUT / name).write_text(text or "", encoding="utf-8")
    log.info("  saved %s", OUT / name)


def main() -> None:
    if not STORAGE_STATE_PATH.exists():
        raise SystemExit(f"No saved session at {STORAGE_STATE_PATH}. Run `python login.py` first.")
    OUT.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(storage_state=str(STORAGE_STATE_PATH), accept_downloads=True)
        page = context.new_page()

        target = CBC_QUEUE_URL or CBC_URL
        log.info("Loading queue: %s", target)
        page.goto(target, wait_until="domcontentloaded", timeout=30000)
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except PlaywrightTimeout:
            pass

        if "login" in page.url.lower() or page.locator('input[type="password"]').count() > 0:
            raise SystemExit("Landed on login — session expired. Re-run `python login.py`.")

        page.wait_for_selector(CONTAINER_SELECTOR, timeout=30000)
        containers = page.locator(CONTAINER_SELECTOR).all()
        log.info("Found %d order rows.", len(containers))

        first = containers[0]
        onclick = first.get_attribute("onclick") or ""
        m = re.search(r"loadDetail\((.*?)\)", onclick)
        if not m:
            _save("first_container.html", first.evaluate("e => e.outerHTML"))
            raise SystemExit("Couldn't find loadDetail(...) in the first row — paste first_container.html.")
        args_js = m.group(1).strip()
        log.info("\nOpening detail via: loadDetail(%s)", args_js)

        # Open the modal (replicates the row's onclick).
        trigger = f"""() => {{
            if (typeof loadDetail !== 'function') return 'loadDetail missing';
            if (window.jQuery) {{
                jQuery('#modalDetail').off('show.bs.modal')
                    .on('show.bs.modal', function () {{ loadDetail({args_js}); }})
                    .modal('show');
            }} else {{ loadDetail({args_js}); }}
            return 'triggered';
        }}"""
        log.info("Trigger: %s", page.evaluate(trigger))

        # Wait (up to 90s) for the Documents list — specifically the Sales Order
        # link — to appear. The notes/documents load is the slow (~30s) part.
        log.info("Waiting for the Documents list / Sales Order link (up to 90s)...")
        so_link = page.locator("#modalDetail a").filter(has_text=_SALES_ORDER_RE)
        have_so = True
        try:
            so_link.first.wait_for(state="attached", timeout=90000)
            log.info("Sales Order link appeared.")
        except PlaywrightTimeout:
            have_so = False
            log.info("No 'Sales Order' link within 90s — dumping whatever links loaded.")
        page.wait_for_timeout(1500)

        # Save the rendered modal + a screenshot for the record.
        try:
            _save("modal_detail.html", page.locator("#modalDetail").inner_html(timeout=3000))
        except Exception as e:  # noqa: BLE001
            log.info("  (couldn't read #modalDetail: %s)", e)
        try:
            page.screenshot(path=str(OUT / "modal_detail.png"), full_page=True)
            log.info("  saved %s", OUT / "modal_detail.png")
        except Exception as e:  # noqa: BLE001
            log.info("  (screenshot failed: %s)", e)

        # List every document link with its href/onclick — this reveals the URL
        # pattern (direct pdf vs. handler) for ALL the docs, not just the SO.
        log.info("\n--- Links in the detail modal ---")
        anchors = page.locator("#modalDetail a").all()
        log.info("Found %d link(s):", len(anchors))
        for i, a in enumerate(anchors):
            try:
                text = (a.inner_text() or "").strip().replace("\n", " ")
                href = a.get_attribute("href")
                oc = a.get_attribute("onclick")
                log.info("  [%d] %r", i, text[:70])
                log.info("       href = %r", href)
                if oc:
                    log.info("       onclick = %r", oc[:200])
            except Exception as e:  # noqa: BLE001
                log.info("  [%d] (couldn't read: %s)", i, e)

        if not have_so:
            log.info("\n(No Sales Order link found — see the list above and modal_detail.html.)")
            input("\nPress Enter to close the browser... ")
            browser.close()
            return

        # Pull the Sales Order link's href and try the FAST path: fetch the pdf
        # directly with the logged-in session (works if href is a real URL).
        link = so_link.first
        so_text = (link.inner_text() or "").strip()
        so_href = link.get_attribute("href")
        so_onclick = link.get_attribute("onclick")
        log.info("\n=== SALES ORDER link ===")
        log.info("  text    = %r", so_text)
        log.info("  href    = %r", so_href)
        log.info("  onclick = %r", so_onclick)

        if so_href and not so_href.lower().startswith("javascript") and so_href.strip() != "#":
            full = urljoin(page.url, so_href)
            log.info("\nFetching directly with the session: %s", full)
            try:
                resp = context.request.get(full)
                body = resp.body()
                ct = resp.headers.get("content-type", "")
                dest = OUT / "sample_sales_order.pdf"
                dest.write_bytes(body)
                log.info("  -> %d bytes, content-type=%s", len(body), ct)
                log.info("  -> saved %s", dest)
                log.info("  => DIRECT URL works. The real tool can fetch each SO in one call (fast).")
            except Exception as e:  # noqa: BLE001
                log.info("  direct fetch failed (%s) — falling back to clicking.", e)
                so_href = None  # force the click path below

        if not so_href or so_href.lower().startswith("javascript") or so_href.strip() == "#":
            log.info("\nhref isn't a plain URL — clicking the link and capturing the result...")
            downloads, popups = [], []
            page.on("download", downloads.append)
            context.on("page", popups.append)
            try:
                link.click()
            except Exception as e:  # noqa: BLE001
                log.info("  click raised: %s", e)
            page.wait_for_timeout(8000)
            if downloads:
                d = downloads[0]
                dest = OUT / (d.suggested_filename or "sample_sales_order.pdf")
                try:
                    d.save_as(str(dest))
                    log.info("  -> DOWNLOAD: file=%r url=%r saved %s", d.suggested_filename, d.url, dest)
                except Exception as e:  # noqa: BLE001
                    log.info("  -> download save failed: %s", e)
                log.info("  => CLICK-TO-DOWNLOAD. The real tool will click each SO link.")
            elif popups:
                pg = popups[-1]
                try:
                    pg.wait_for_load_state("domcontentloaded", timeout=10000)
                except PlaywrightTimeout:
                    pass
                log.info("  -> opened a new tab: %s", pg.url)
                log.info("  => If that's a .pdf URL we can fetch it directly next round.")
            else:
                log.info("  -> no download or popup detected within 8s.")

        log.info("\n========== NEXT STEPS ==========")
        log.info("Saved under: %s", OUT)
        log.info("Paste the console output back to me (especially the SALES ORDER block")
        log.info("and the link list). If sample_sales_order.pdf saved and opens, we're done discovering.")

        input("\nPress Enter to close the browser... ")
        browser.close()


if __name__ == "__main__":
    main()
