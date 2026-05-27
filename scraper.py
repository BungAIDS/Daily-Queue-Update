"""Playwright-based scraper for the cbcinsider work queue.

The queue table is rendered as paired rows per job:
  Row 1 (job row):    Status | Job# | Oper | Item/Rep | Assigned | Checker | Start | End | Plan Hrs | FanNet | Total Price
  Row 2 (detail row): Customer name in col 1, "Ship With ####" note in the Plan Hrs col, others blank.

We pair them by walking rows two at a time.
"""
from __future__ import annotations

import logging
from typing import List, Dict, Any

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

from config import CBC_URL, CBC_USERNAME, CBC_PASSWORD

log = logging.getLogger(__name__)

FIELDS = [
    "status", "job", "oper", "item_rep", "assigned_to", "checker",
    "start_date", "end_date", "plan_hrs", "fannet_date", "total_price",
]


def _cell_text(cell) -> str:
    return (cell.inner_text() or "").strip()


def scrape_queue(headless: bool = True) -> List[Dict[str, Any]]:
    """Log in, navigate to the queue, and return one dict per job (paired rows merged)."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context()
        page = context.new_page()

        log.info("Loading login page: %s", CBC_URL)
        page.goto(CBC_URL, wait_until="domcontentloaded", timeout=30000)

        # Fill login form — selectors are best-effort; adjust if the site changes.
        try:
            page.fill('input[name="username"], input[name="user"], input[type="email"]', CBC_USERNAME)
            page.fill('input[name="password"], input[type="password"]', CBC_PASSWORD)
            page.click('button[type="submit"], input[type="submit"]')
        except PWTimeout as e:
            browser.close()
            raise RuntimeError(f"Login form not found or not interactive: {e}")

        page.wait_for_load_state("networkidle", timeout=30000)

        if "login" in page.url.lower() or page.locator("text=/invalid|incorrect/i").count() > 0:
            browser.close()
            raise RuntimeError("Login failed — check CBC_USERNAME / CBC_PASSWORD.")

        # The queue is usually on the landing page or under a "Queue" nav item.
        try:
            queue_link = page.locator("a:has-text('Queue'), a:has-text('Work Queue')").first
            if queue_link.count() > 0:
                queue_link.click()
                page.wait_for_load_state("networkidle", timeout=30000)
        except Exception:
            pass  # already on the queue page

        # Wait for the main table to render
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
                    # Customer is in col 0 (Status/Customer column)
                    customer = _cell_text(detail_cells[0])
                    # "Ship With" note lives in the Plan Hrs column (index 8)
                    if len(detail_cells) > 8:
                        ship_with = _cell_text(detail_cells[8])
                    i += 2
                else:
                    i += 1
            else:
                i += 1

            job["customer"] = customer
            job["ship_with"] = ship_with

            # Only keep rows that look like real jobs (have a job number)
            if job.get("job"):
                jobs.append(job)

        log.info("Parsed %d jobs from queue", len(jobs))
        browser.close()
        return jobs
