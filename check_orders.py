"""Check out specific orders — hand it order numbers, it pulls each one's quote
run and tells you which template matched and what fields came out.

    python check_orders.py 421473 421492 420410 ...

For each order it opens the detail (from the board if it's there, otherwise via
the queue's search box — the same path the backfill uses), finds the quote /
construction run, downloads it (keeping its real extension), and runs it through
the template collection in `templates.py`. It prints, per order:

    - which template matched (and the score for each, so you can see why),
    - the fields it pulled + the compact summary,
    - the first reconstructed lines (so a not-yet-handled format is easy to add).

Runs headless and unattended (no Enter-to-close) so a whole list goes in one
shot; pass --show to watch the browser. Downloads land in OUTPUT_DIR/doc_discovery.

Paste a block back and the matching template's fields get pinned down — or, for
an unrecognized format, a new template gets added.
"""
from __future__ import annotations

import logging
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

from config import CBC_URL, CBC_QUEUE_URL, STORAGE_STATE_PATH
from scraper import CONTAINER_SELECTOR
from sales_orders import _run_docs, _find_autocad_folders, _run_files_in_folder
from templates import QuoteRunContext, TEMPLATES, parse_quote_run
# Reuse the discovery plumbing so this reflects exactly what the daily run sees.
from discover_documents import (
    OUT, _board_args, _open_detail, _collect_docs, _download,
)

logging.basicConfig(level=logging.INFO, format="%(message)s")
log = logging.getLogger("check")


def _print_run(jn: str, path: Path, design: str | None) -> None:
    """Run one downloaded quote run through the templates and print the result."""
    ctx = QuoteRunContext(path, design)
    log.info("    template scores: %s",
             ", ".join(f"{t.key}={t.score(ctx)}" for t in TEMPLATES if t.score(ctx)))
    r = parse_quote_run(path, design)
    log.info("    matched template : %s   (design=%s, ext=%s)",
             r["template"], r["design"], ctx.ext or "(none)")
    if r["fields"]:
        log.info("    fields (%d):", len(r["fields"]))
        for k, v in r["fields"].items():
            log.info("      %-26s %s", k, v)
        log.info("    summary: %s", r["summary"])
    else:
        log.info("    no fields pulled yet — this format needs a template (or its "
                 "fields pinned down). Raw lines below:")
    if r["raw_lines"]:
        log.info("    --- first lines ---")
        for ln in r["raw_lines"][:30]:
            log.info("      %s", ln)


def _check_one(context, page, board: dict, jn: str) -> None:
    log.info("\n%s\n=== Order %s ===", "=" * 72, jn)
    if jn in board:
        if not _open_detail(page, jn, board[jn]):
            log.info("  detail didn't load within 120s — skipped.")
            return
    else:
        from backfill_orders import open_order_detail
        log.info("  not on the board — searching for it ...")
        if not open_order_detail(page, jn):
            log.info("  couldn't surface %s via the search box — skipped. "
                     "(set CBC_SEARCH_SELECTOR in .env if the queue has a find-order field.)", jn)
            return

    docs = _collect_docs(page)
    runs = _run_docs([(d["href"], d) for d in docs])
    if runs:
        for i, (_, d) in enumerate(runs):
            label = f"{jn}_quote_run" + (f"_{i + 1}" if len(runs) > 1 else "")
            log.info("  quote run: %s", d.get("fn") or label)
            dest = _download(context, page, d, label)
            if dest:
                _print_run(jn, dest, None)
    else:
        log.info("  no quote run in the documents.")

    # Some orders keep the run only in their AutoCAD folder — check there too.
    try:
        info = _find_autocad_folders([jn]).get(jn)
        hits = _run_files_in_folder(info["path"]) if info else []
        for f in hits:
            log.info("  quote run (AutoCAD folder): %s", f)
            _print_run(jn, Path(f), None)
        if info and not hits and not runs:
            log.info("  (no run-named files in the AutoCAD folder either.)")
    except Exception as e:  # noqa: BLE001 - one order's folder miss shouldn't stop the batch
        log.info("  (AutoCAD folder check failed: %s)", e)


def check_orders(jobs: list[str], headless: bool = True) -> None:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(storage_state=str(STORAGE_STATE_PATH), accept_downloads=True)
        page = context.new_page()
        page.goto(CBC_QUEUE_URL or CBC_URL, wait_until="domcontentloaded", timeout=30000)
        if "login" in page.url.lower() or page.locator('input[type="password"]').count() > 0:
            raise SystemExit("Landed on login — session expired. Re-run `python login.py`.")
        page.wait_for_selector(CONTAINER_SELECTOR, timeout=30000)

        board = dict(_board_args(page))
        log.info("Board has %d orders; checking %d requested: %s",
                 len(board), len(jobs), ", ".join(jobs))
        for jn in jobs:
            try:
                _check_one(context, page, board, jn)
            except Exception as e:  # noqa: BLE001 - keep the batch going
                log.info("  error checking %s: %s", jn, e)

        log.info("\nDownloads saved under %s", OUT)
        log.info("Paste any block back to pin down that format's fields (or add a template).")
        if not headless:
            input("\nPress Enter to close the browser... ")
        browser.close()


def main() -> None:
    if not STORAGE_STATE_PATH.exists():
        raise SystemExit(f"No saved session at {STORAGE_STATE_PATH}. Run `python login.py` first.")
    args = [a for a in sys.argv[1:] if a != "--show"]
    headless = "--show" not in sys.argv[1:]
    jobs = [a.strip() for a in args if a.strip()]
    if not jobs:
        raise SystemExit("Usage: python check_orders.py <order#> [order# ...] [--show]")
    check_orders(jobs, headless=headless)


if __name__ == "__main__":
    main()
