"""Inspect ONE job's documents — to confirm how a CHANGE ORDER shows up.

Usage:
    python discover_sales_order.py            # first order on the board
    python discover_sales_order.py 421314     # a specific job number

It opens that order's detail modal, lists every document with its parsed
pid (doctype / internal id / REVISION), highlights the Sales Order revision
and any change-order ("...CO..." / "change") docs, and downloads the SO.

This is the signal we'll track to catch change orders: the Sales Order
revision number and the presence of change-order documents both come straight
from this list — no PDF parsing needed. Read-only.
"""
from __future__ import annotations

import logging
import re
import sys
from urllib.parse import urlparse, parse_qs, urljoin

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

from config import CBC_URL, CBC_QUEUE_URL, STORAGE_STATE_PATH, OUTPUT_DIR
from scraper import CONTAINER_SELECTOR

logging.basicConfig(level=logging.INFO, format="%(message)s")
log = logging.getLogger("discover")

OUT = OUTPUT_DIR / "so_discovery"

# pid looks like "CBC_SalesOrder-32045-3-LATEST" -> type / id / revision / tag.
PID_RE = re.compile(r"^(?P<type>.+?)-(?P<id>\d+)-(?P<rev>\d+)-(?P<tag>[A-Za-z0-9]+)$")


def _jobnum(args_js: str) -> str:
    first = args_js.split(",", 1)[0].strip().strip("'\"")
    return first.split("-", 1)[0]


def _parse_doc(href: str) -> dict:
    """Pull pid + filename out of a downloaddoc.aspx href and split the pid."""
    q = parse_qs(urlparse(href).query)
    pid = q.get("pid", [""])[0]
    fn = q.get("fn", [""])[0]
    m = PID_RE.match(pid)
    if m:
        return {"pid": pid, "fn": fn, "type": m["type"], "id": m["id"], "rev": int(m["rev"])}
    return {"pid": pid, "fn": fn, "type": pid, "id": "", "rev": None}


def _is_change_order(doc: dict) -> bool:
    t = (doc["type"] or "").lower()
    f = (doc["fn"] or "").lower()
    return ("co" in t and "sales" not in t) or "change" in f or "change" in t


def main() -> None:
    if not STORAGE_STATE_PATH.exists():
        raise SystemExit(f"No saved session at {STORAGE_STATE_PATH}. Run `python login.py` first.")
    OUT.mkdir(parents=True, exist_ok=True)
    want_job = sys.argv[1].strip() if len(sys.argv) > 1 else None

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

        # Pick the container: match the requested job number, else the first.
        chosen, args = None, None
        for c in containers:
            m = re.search(r"loadDetail\((.*?)\)", c.get_attribute("onclick") or "")
            if not m:
                continue
            a = m.group(1).strip()
            if want_job is None or _jobnum(a) == want_job:
                chosen, args = c, a
                break
        if chosen is None:
            raise SystemExit(
                f"Job {want_job!r} not found on the board. (It may have dropped off — "
                "try a job that's currently in the queue, or screenshot its SO instead.)"
            )

        jn = _jobnum(args)
        log.info("\nOpening detail for job %s ...", jn)
        page.evaluate(f"""() => {{
            if (window.jQuery) {{
                jQuery('#modalDetail').off('show.bs.modal')
                    .on('show.bs.modal', function () {{ loadDetail({args}); }})
                    .modal('show');
            }} else {{ loadDetail({args}); }}
        }}""")

        so_re = re.compile(re.escape(jn))
        try:
            page.locator("#modalDetail a").filter(has_text=so_re).first.wait_for(state="attached", timeout=120000)
        except PlaywrightTimeout:
            log.info("Documents didn't appear within 120s.")

        # Collect + parse every document link.
        docs = []
        for a in page.locator("#modalDetail a").all():
            href = a.get_attribute("href") or ""
            if "downloaddoc.aspx" not in href.lower():
                continue
            docs.append(_parse_doc(href))

        log.info("\n--- Documents for job %s (%d) ---", jn, len(docs))
        for d in docs:
            tag = "  <-- CHANGE ORDER" if _is_change_order(d) else ""
            rev = f"rev {d['rev']}" if d["rev"] is not None else "rev ?"
            log.info("  %-22s %-6s  %s%s", d["type"], rev, d["fn"][:55], tag)

        sales_orders = [d for d in docs if "salesorder" in (d["type"] or "").lower() and not _is_change_order(d)]
        change_orders = [d for d in docs if _is_change_order(d)]

        log.info("\n--- Summary ---")
        if sales_orders:
            top = max(sales_orders, key=lambda d: d["rev"] or 0)
            log.info("  Sales Order: %s  (revision %s)", top["fn"], top["rev"])
        else:
            log.info("  No Sales Order doc found.")
        log.info("  Change-order docs: %d %s", len(change_orders),
                 [d["fn"] for d in change_orders] or "")

        # Download the SO as a sample.
        so_link = page.locator("#modalDetail a").filter(has_text=re.compile(r"sales order", re.I))
        if so_link.count() > 0:
            href = so_link.first.get_attribute("href")
            full = urljoin(page.url, href)
            try:
                resp = context.request.get(full)
                dest = OUT / f"{jn}_sales_order.pdf"
                dest.write_bytes(resp.body())
                log.info("\n  Downloaded SO -> %s (%d bytes)", dest, len(resp.body()))
            except Exception as e:  # noqa: BLE001
                log.info("\n  SO download failed: %s", e)

        log.info("\nPaste the console output back to me — especially the Documents list")
        log.info("and Summary. That confirms how a change order shows up so I can track it.")

        input("\nPress Enter to close the browser... ")
        browser.close()


if __name__ == "__main__":
    main()
