"""Discovery helper #2: capture how a job's SALES ORDER detail is loaded.

We learned from round 1 that clicking a dispatch row has no link/anchor — the
row's own onclick calls a JS function:

    loadDetail('421314-0','19','DRIVE SET RPM',' ', 0)

...which opens a Bootstrap modal (#modalDetail) and fills it over AJAX. This
script triggers that for the FIRST order and captures:

  1. the SOURCE of loadDetail()      -> shows the exact backend URL it calls
  2. every network request it fires   -> the real endpoint + params
  3. the response body of the detail call
  4. the rendered modal HTML          -> what the "sales order" actually contains

    python discover_sales_order.py

Read-only: it just opens one order's detail pop-up, same as clicking it.
Everything is saved under OUTPUT_DIR/so_discovery/. When done, paste the
console output back to me and attach the files it lists.
"""
from __future__ import annotations

import logging
import re

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

from config import CBC_URL, CBC_QUEUE_URL, STORAGE_STATE_PATH, OUTPUT_DIR
from scraper import CONTAINER_SELECTOR

logging.basicConfig(level=logging.INFO, format="%(message)s")
log = logging.getLogger("discover")

OUT = OUTPUT_DIR / "so_discovery"

# Static assets we don't care about when logging what the modal fetched.
_STATIC = (".js", ".css", ".png", ".gif", ".jpg", ".jpeg", ".svg", ".woff", ".woff2", ".ico")


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

        # 1) The row's own onclick -> pull out the loadDetail(...) call.
        onclick = first.get_attribute("onclick") or ""
        m = re.search(r"loadDetail\((.*?)\)", onclick)
        if not m:
            _save("first_container.html", first.evaluate("e => e.outerHTML"))
            raise SystemExit(
                "Couldn't find loadDetail(...) in the row's onclick. Saved the row "
                "HTML to so_discovery/first_container.html — paste that back to me."
            )
        args_js = m.group(1).strip()
        log.info("\nRow opens its detail via:  loadDetail(%s)", args_js)

        # 2) Source of loadDetail() — usually reveals the endpoint URL directly.
        src = page.evaluate(
            "() => (typeof loadDetail === 'function') ? loadDetail.toString() : 'loadDetail NOT DEFINED'"
        )
        _save("loaddetail_source.js", src)
        log.info("\n--- loadDetail() source (also saved) ---\n%s", src[:1500])
        if len(src) > 1500:
            log.info("...(truncated; full version in loaddetail_source.js)")

        # 3) Capture network traffic while the modal loads.
        requests_seen: list[tuple[str, str]] = []
        responses_seen = []
        page.on("request", lambda r: requests_seen.append((r.method, r.url)))
        page.on("response", lambda r: responses_seen.append(r))

        start_idx = len(responses_seen)
        log.info("\nTriggering the detail modal for the first order...")
        trigger = f"""() => {{
            if (typeof loadDetail !== 'function') return 'loadDetail missing';
            if (window.jQuery) {{
                jQuery('#modalDetail').off('show.bs.modal')
                    .on('show.bs.modal', function () {{ loadDetail({args_js}); }})
                    .modal('show');
            }} else {{ loadDetail({args_js}); }}
            return 'triggered';
        }}"""
        log.info("  trigger result: %s", page.evaluate(trigger))
        page.wait_for_timeout(5000)  # let the AJAX land and the modal render

        # 4) Dump the rendered modal.
        try:
            modal_html = page.locator("#modalDetail").inner_html(timeout=3000)
            _save("modal_detail.html", modal_html)
        except Exception as e:  # noqa: BLE001
            log.info("  (couldn't read #modalDetail: %s)", e)
        try:
            page.screenshot(path=str(OUT / "modal_detail.png"), full_page=True)
            log.info("  saved %s", OUT / "modal_detail.png")
        except Exception as e:  # noqa: BLE001
            log.info("  (screenshot failed: %s)", e)

        # 5) Report the non-static requests fired during the load + save bodies.
        log.info("\n--- Network calls fired while opening the modal ---")
        new_responses = responses_seen[start_idx:]
        dumped = 0
        for resp in new_responses:
            url = resp.url
            low = url.lower().split("?", 1)[0]
            if low.endswith(_STATIC):
                continue
            ctype = (resp.headers or {}).get("content-type", "")
            log.info("  %s  %s  [%s]", resp.status, url, ctype)
            if any(t in ctype for t in ("html", "json", "text", "xml")):
                try:
                    body = resp.text()
                    dumped += 1
                    _save(f"response_{dumped}.txt", f"URL: {url}\nSTATUS: {resp.status}\nCONTENT-TYPE: {ctype}\n\n{body}")
                except Exception as e:  # noqa: BLE001
                    log.info("       (couldn't read body: %s)", e)
        if dumped == 0:
            log.info("  (no obvious HTML/JSON response captured — the source file above")
            log.info("   should still show the endpoint loadDetail calls.)")

        log.info("\n========== NEXT STEPS ==========")
        log.info("Saved under: %s", OUT)
        log.info("Paste the console output back to me, and attach:")
        log.info("  - loaddetail_source.js")
        log.info("  - modal_detail.html")
        log.info("  - response_*.txt   (if any)")
        log.info("  - modal_detail.png")
        log.info("That tells me the exact endpoint + what a 'sales order' contains.")

        input("\nPress Enter to close the browser... ")
        browser.close()


if __name__ == "__main__":
    main()
