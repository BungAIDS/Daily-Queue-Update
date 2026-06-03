"""Discovery probe #4: can we beat the ~30s detail load?

Two unknowns decide the fastest way to grab sales orders:
  1. Is the document list delivered by a SEPARATE, faster endpoint we could
     hit directly, or only by the slow NotesPanel postback to dispatch.aspx?
  2. Is the ~30s a one-time warmup, or does every order cost ~30s?

This opens the FIRST TWO orders, times each load until that order's documents
appear, prints the slowest network calls, and identifies which response
actually carries the 'downloaddoc.aspx' links. Read-only.

    python discover_sales_order.py
"""
from __future__ import annotations

import logging
import re
import time

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

from config import CBC_URL, CBC_QUEUE_URL, STORAGE_STATE_PATH, OUTPUT_DIR
from scraper import CONTAINER_SELECTOR

logging.basicConfig(level=logging.INFO, format="%(message)s")
log = logging.getLogger("discover")

OUT = OUTPUT_DIR / "so_discovery"
_STATIC = (".js", ".css", ".png", ".gif", ".jpg", ".jpeg", ".svg", ".woff", ".woff2", ".ico")


def _jobnum(args_js: str) -> str:
    """First loadDetail arg is like '421314-0' -> job number '421314'."""
    first = args_js.split(",", 1)[0].strip().strip("'\"")
    return first.split("-", 1)[0]


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
        log.info("Found %d order rows.\n", len(containers))

        # Per-request timing + the responses we'll scan for the doc-list carrier.
        reqs: dict = {}
        responses: list = []
        page.on("request", lambda r: reqs.__setitem__(r, {"url": r.url, "start": time.monotonic(), "end": None}))

        def _finish(r):
            if r in reqs:
                reqs[r]["end"] = time.monotonic()

        page.on("requestfinished", _finish)
        page.on("requestfailed", _finish)
        page.on("response", responses.append)

        for idx in (0, 1):
            if idx >= len(containers):
                break
            c = containers[idx]
            m = re.search(r"loadDetail\((.*?)\)", c.get_attribute("onclick") or "")
            if not m:
                log.info("[job %d] no loadDetail() — skipping", idx)
                continue
            args = m.group(1).strip()
            jn = _jobnum(args)
            log.info("=" * 60)
            log.info("[job %d] %s — opening detail...", idx, jn)

            reqs.clear()
            responses.clear()
            t0 = time.monotonic()

            page.evaluate(f"""() => {{
                if (window.jQuery) {{
                    jQuery('#modalDetail').off('show.bs.modal')
                        .on('show.bs.modal', function () {{ loadDetail({args}); }})
                        .modal('show');
                }} else {{ loadDetail({args}); }}
            }}""")

            # "Loaded" = this order's own documents (named with its job number)
            # have appeared in the modal.
            link = page.locator("#modalDetail a").filter(has_text=re.compile(re.escape(jn)))
            try:
                link.first.wait_for(state="attached", timeout=120000)
                elapsed = time.monotonic() - t0
                log.info("[job %d] documents appeared after %.1fs", idx, elapsed)
            except PlaywrightTimeout:
                log.info("[job %d] documents did NOT appear within 120s", idx)

            # Slowest network calls during this load.
            durs = [(info["end"] - info["start"], info["url"])
                    for info in reqs.values() if info["end"]]
            durs.sort(reverse=True)
            log.info("[job %d] slowest network calls:", idx)
            for dur, url in durs[:5]:
                log.info("    %6.1fs  %s", dur, url[:110])

            # Which response actually carried the document links?
            carrier = None
            for resp in responses:
                base = resp.url.lower().split("?", 1)[0]
                if base.endswith(_STATIC):
                    continue
                try:
                    body = resp.text()
                except Exception:  # noqa: BLE001
                    continue
                if "downloaddoc.aspx" in body:
                    carrier = resp
                    info = reqs.get(resp.request)
                    dur = (info["end"] - info["start"]) if info and info["end"] else None
                    log.info("[job %d] doc-list carrier: %s", idx, resp.url[:110])
                    log.info("           method=%s  took=%s  bytes=%d",
                             resp.request.method,
                             f"{dur:.1f}s" if dur else "?", len(body))
                    break
            if carrier is None:
                log.info("[job %d] couldn't pin the doc-list carrier response.", idx)

            # Close the modal before the next job.
            page.evaluate("() => { if (window.jQuery) jQuery('#modalDetail').modal('hide'); }")
            page.wait_for_timeout(1500)

        log.info("\n" + "=" * 60)
        log.info("READ ME: compare the two jobs' times (warmup vs. always-slow),")
        log.info("and look at the 'doc-list carrier' line:")
        log.info("  - if its URL is dispatch.aspx (a POST) and it's the slow call,")
        log.info("    the 30s is unavoidable per job -> we parallelize.")
        log.info("  - if it's a SEPARATE fast URL, we can hit that directly per job.")
        log.info("Paste this whole console output back to me.")

        input("\nPress Enter to close the browser... ")
        browser.close()


if __name__ == "__main__":
    main()
