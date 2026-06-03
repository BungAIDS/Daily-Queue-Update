"""Playwright scraper for the cbcinsider engineering dispatch queue.

Auth model: this does NOT store your password. You log in once manually with
login.py, which saves your browser session (cookies) to cbc_session.json.
This script reuses that saved session. When the session eventually expires,
the run fails with a clear message and you re-run login.py.

Page structure (ASP.NET WebForms, dispatch.aspx). The queue is NOT a single
table — each order is its own container div holding a 2-row table:

    <div id="MainContent_rptDispatch_Container_{i}">
        <table>
          <tr> ...job row...    </tr>
          <tr> ...detail row... </tr>
        </table>
    </div>

Job row cells (index : meaning):
    0 Status | 1 Job | 2 Oper | 3 Item | 4 Assigned To | 5 Checker |
    6 Start Date | 7 End Date | 8 Plan Hrs | 9 FanNet Date | 10 Total Price |
    11 status-note text + flag icons (notes / unapproved / credit-hold)

Detail row cells:
    0 Customer (colspan 3) | 1 Primary Rep | ... | 6 "Ship With ####" (if set)
"""
from __future__ import annotations

import logging
import re
from typing import List, Dict, Any

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout, Error as PlaywrightError

from config import CBC_URL, CBC_QUEUE_URL, CBC_WORK_CENTER, STORAGE_STATE_PATH

log = logging.getLogger(__name__)

# Each order row is one of these containers (the header div is "..._Header",
# which this prefix deliberately does not match).
CONTAINER_SELECTOR = 'div[id^="MainContent_rptDispatch_Container_"]'


def _cell_text(cell) -> str:
    return (cell.inner_text() or "").strip()


def _design_from_item(item: str) -> str:
    """Pull the design number out of the Item field.

    Items are typically "DD-x-xxxx" where DD is the design (e.g. "47-0-0000" =
    Design 47). Some are spelled out as "DESIGN 36". A few are codes like
    "EMSI" with no design number — keep those as-is for the column.
    """
    s = (item or "").strip()
    if not s:
        return ""
    if s.upper().startswith("DESIGN "):
        return s.split(None, 1)[1].strip()
    if "-" in s:
        head = s.split("-", 1)[0].strip()
        if head:
            return head
    return s


def _expected_count(page) -> int | None:
    """The page prints 'Results: N' — use it as a sanity check on our parse."""
    try:
        txt = page.locator("#MainContent_lblRecordCount").inner_text(timeout=2000)
    except Exception:
        return None
    m = re.search(r"(\d+)", txt or "")
    return int(m.group(1)) if m else None


def _parse_container(container) -> Dict[str, Any] | None:
    rows = container.locator("tr").all()
    if len(rows) < 2:
        return None

    jc = rows[0].locator("td").all()
    dc = rows[1].locator("td").all()
    if len(jc) < 11:
        return None  # not a real job row

    job: Dict[str, Any] = {
        "status": _cell_text(jc[0]),
        "job": _cell_text(jc[1]),
        "oper": _cell_text(jc[2]),
        "item": _cell_text(jc[3]),
        "design": _design_from_item(_cell_text(jc[3])),
        "assigned_to": _cell_text(jc[4]),
        "checker": _cell_text(jc[5]),
        "start_date": _cell_text(jc[6]),
        "end_date": _cell_text(jc[7]),
        "plan_hrs": _cell_text(jc[8]),
        "fannet_date": _cell_text(jc[9]),
        "total_price": _cell_text(jc[10]),
        "status_note": _cell_text(jc[11]) if len(jc) > 11 else "",
    }

    # Detail row: customer + primary rep + optional "Ship With ####".
    job["customer"] = _cell_text(dc[0]) if len(dc) > 0 else ""
    job["primary_rep"] = _cell_text(dc[1]) if len(dc) > 1 else ""
    job["ship_with"] = _cell_text(dc[6]) if len(dc) > 6 else ""

    # Flag icons live in the job row's last cell.
    job["unapproved"] = container.locator("img[id*='unapproved']").count() > 0
    job["credit_hold"] = container.locator("img[id*='credithold']").count() > 0
    job["has_notes"] = container.locator("img[id*='message']").count() > 0

    return job


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
        # Retry the initial nav a couple of times — at 5 AM the PC may have just
        # woken from sleep and the Wi-Fi/VPN may take a moment to come up, which
        # surfaces as ERR_CONNECTION_TIMED_OUT on the first attempt.
        import time as _time
        last_err = None
        for attempt in (1, 2, 3):
            try:
                page.goto(target, wait_until="domcontentloaded", timeout=30000)
                last_err = None
                break
            except (PlaywrightTimeout, PlaywrightError) as e:
                last_err = e
                if attempt == 3:
                    break
                wait_s = 5 * attempt  # 5s, then 10s
                log.warning("page.goto attempt %d failed (%s); retrying in %ds",
                            attempt, type(e).__name__, wait_s)
                _time.sleep(wait_s)
        if last_err is not None:
            raise last_err
        # Best-effort: this app uses AJAX/long-poll, so "networkidle" may never
        # settle. Don't fail the run on it — we gate on the order rows below.
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except PlaywrightTimeout:
            log.info("networkidle didn't settle in 15s; proceeding")

        # If the saved session expired, we get bounced to a login page.
        if "login" in page.url.lower() or page.locator('input[type="password"]').count() > 0:
            browser.close()
            raise RuntimeError(
                "Saved session has expired (landed on the login page). "
                "Run `python login.py` again to refresh it."
            )

        # Guard against silently scraping the wrong queue: the dispatch page is
        # filtered by Work Center, and the server remembers your last choice.
        if CBC_WORK_CENTER:
            sel = page.locator("#MainContent_wc")
            current = sel.input_value() if sel.count() > 0 else None
            if current and current != CBC_WORK_CENTER:
                browser.close()
                raise RuntimeError(
                    f"Work Center is '{current}' but CBC_WORK_CENTER expects "
                    f"'{CBC_WORK_CENTER}'. Re-run login.py with the right Work "
                    "Center selected, or update CBC_WORK_CENTER in .env."
                )

        try:
            page.wait_for_selector(CONTAINER_SELECTOR, timeout=30000)
        except PlaywrightTimeout:
            browser.close()
            hint = "" if CBC_QUEUE_URL else (
                " CBC_QUEUE_URL is not set — set it in .env to the exact "
                "dispatch page URL from your browser's address bar."
            )
            raise RuntimeError("Queue list never appeared (no order rows found)." + hint)

        containers = page.locator(CONTAINER_SELECTOR).all()
        log.info("Found %d order rows", len(containers))

        expected = _expected_count(page)
        if expected is not None and expected != len(containers):
            log.warning("Page reports %d results but parsed %d order rows", expected, len(containers))

        jobs: List[Dict[str, Any]] = []
        for c in containers:
            job = _parse_container(c)
            if job and job.get("job"):
                jobs.append(job)

        log.info("Parsed %d jobs from queue", len(jobs))
        browser.close()
        return jobs
