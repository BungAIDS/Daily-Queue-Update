"""Playwright scraper for the cbcinsider work queue.

Auth model: this does NOT store your password. You log in once manually with
login.py, which saves your browser session (cookies) to cbc_session.json.
This script reuses that saved session. When the session eventually expires,
the run fails with a clear message and you re-run login.py.

The queue table is rendered as paired rows per job:
  Row 1 (job row):    Status | Job# | Oper | Item/Rep | Assigned | Checker | Start | End | Plan Hrs | FanNet | Total Price
  Row 2 (detail row): Customer name in col 1, "Ship With ####" note in the Plan Hrs col, others blank.

We pair them by walking rows two at a time.
"""
from __future__ import annotations

import logging
from typing import List, Dict, Any

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

from config import CBC_URL, CBC_QUEUE_URL, STORAGE_STATE_PATH

log = logging.getLogger(__name__)

FIELDS = [
    "status", "job", "oper", "item_rep", "assigned_to", "checker",
    "start_date", "end_date", "plan_hrs", "fannet_date", "total_price",
]


def _cell_text(cell) -> str:
    return (cell.inner_text() or "").strip()


def scrape_queue(headless: bool = True) -> List[Dict[str, Any]]:
    """Reuse the saved session, navigate to the queue, and return one dict per job."""
    if not STORAGE_STATE_PATH.exists():
        raise RuntimeError(
            f"No saved session found at {STORAGE_STATE_PATH}. "
            "Run `python login.py` first to log in and save your session."
        )

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(storage_state=str(STORAGE_STATE_PATH))
        page = context.new_page()

        target = CBC_QUEUE_URL or CBC_URL
        log.info("Loading queue with saved session: %s", target)
        page.goto(target, wait_until="domcontentloaded", timeout=30000)
        # Best-effort: many internal apps long-poll or hold a websocket open, so
        # "networkidle" may never settle. Don't let that fail the whole run — we
        # rely on the explicit wait_for_selector("table tbody tr") below anyway.
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except PlaywrightTimeout:
            log.info("networkidle didn't settle in 15s; proceeding to table wait")

        # If the saved session expired, we'll get bounced to a login page.
        if "login" in page.url.lower() or page.locator('input[type="password"]').count() > 0:
            browser.close()
            raise RuntimeError(
                "Saved session has expired (landed on the login page). "
                "Run `python login.py` again to refresh it."
            )

        # If we didn't go straight to the queue, try clicking a Queue nav link.
        if not CBC_QUEUE_URL:
            try:
                queue_link = page.locator("a:has-text('Queue'), a:has-text('Work Queue')").first
                if queue_link.count() > 0:
                    queue_link.click()
                    page.wait_for_load_state("networkidle", timeout=30000)
            except Exception:
                pass  # may already be on the queue page

        page.wait_for_selector("table tbody tr", timeout=30000)

        rows = page.locator("table tbody tr").all()
        log.info("Found %d raw rows in the queue table", len(rows))

        jobs: List[Dict[str, Any]] = []
        i = 0
        while i < len(rows):
            job_cells = rows[i].locator("td").all()
            if len(job_cells) < len(FIELDS):
                # Probably a header, spacer, or footer row — skip
                i += 1
                continue

            job_values = [_cell_text(c) for c in job_cells[: len(FIELDS)]]
            job = dict(zip(FIELDS, job_values))

            # Detail row sits immediately below the job row
            customer = ""
            ship_with = ""
            if i + 1 < len(rows):
                detail_cells = rows[i + 1].locator("td").all()
                if detail_cells:
                    customer = _cell_text(detail_cells[0])
                    if len(detail_cells) > 8:
                        ship_with = _cell_text(detail_cells[8])
                    i += 2
                else:
                    i += 1
            else:
                i += 1

            job["customer"] = customer
            job["ship_with"] = ship_with

            if job.get("job"):
                jobs.append(job)

        log.info("Parsed %d jobs from queue", len(jobs))
        browser.close()
        return jobs
