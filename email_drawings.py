"""CBC Insider "Email Drawings" helper — PROBE + PRE-FILL only. SEND is disabled.

Step 3 of *Completing Transmittals* is submitting the drawings through CBC
Insider (Engineering -> Email Drawings). It's a TWO-page flow:

  Page 1 (order lookup): type the order # and submit to advance.
  Page 2 (email form):   the recipients field, the file attach control, and the
                         "Send" button — the button we do NOT press yet.

This module reuses the saved browser session (login.py / scraper.py — no password
stored) to:

  * --probe : walk both pages read-only and dump their fields/selectors, so the
              order #, advance, emails, attach, and Send controls can be
              identified. (First run dumps page 1; once the order # + advance
              selectors are set, re-run with --order to also dump page 2.)
  * prepare : enter the order #, advance to page 2, paste the emails, and attach
              the files — then STOP on page 2 for a human to review. Send is off.

>>> THE ACTUAL SEND IS HARD-DISABLED <<<
The line that clicks the form's submit/send button is commented out and guarded
by `_SEND_HARD_DISABLED = True`, so there is NO code path that can mail a
transmittal to a customer. Re-enabling it is a deliberate edit (uncomment the
submit block AND flip the guard) — not something a stray call can trigger.

    python email_drawings.py --probe                 # discover the form fields
    python email_drawings.py --probe --order 421693  # also type the order # in
    python email_drawings.py 421693                  # gather + pre-fill (no send)
"""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path
from typing import List, Optional

from config import (
    CBC_URL, EMAIL_DRAWINGS_URL, STORAGE_STATE_PATH, TRANSMITTAL_MODE,
    EMAIL_DRAWINGS_ORDER_SELECTOR, EMAIL_DRAWINGS_ORDER_SUBMIT_SELECTOR,
    EMAIL_DRAWINGS_EMAILS_SELECTOR, EMAIL_DRAWINGS_ATTACH_SELECTOR,
    EMAIL_DRAWINGS_SUBMIT_SELECTOR,
)

log = logging.getLogger(__name__)

# Belt-and-suspenders kill switch. While True, no submit/send can run — and the
# submit code below is also commented out, so this is the second of two locks.
_SEND_HARD_DISABLED = True


# --------------------------------------------------------------------------- #
# Session / navigation (mirrors scraper.py's saved-session pattern)            #
# --------------------------------------------------------------------------- #
def _require_session() -> None:
    if not STORAGE_STATE_PATH.exists():
        raise RuntimeError(
            f"No saved session found at {STORAGE_STATE_PATH}. "
            "Run `python login.py` first to log in and save your session."
        )


def _goto_form(page, order: Optional[str] = None) -> None:
    """Navigate to the Email Drawings page (or the CBC landing page if its URL
    isn't configured yet, so the probe can still be pointed at it by hand)."""
    target = EMAIL_DRAWINGS_URL or CBC_URL
    log.info("Loading Email Drawings form with saved session: %s", target)
    page.goto(target, wait_until="domcontentloaded", timeout=30000)
    try:
        page.wait_for_load_state("networkidle", timeout=10000)
    except Exception:  # noqa: BLE001 - AJAX app may never go idle
        pass
    if "login" in page.url.lower() or page.locator('input[type="password"]').count() > 0:
        raise RuntimeError(
            "Saved session has expired (landed on the login page). "
            "Run `python login.py` again to refresh it."
        )


# --------------------------------------------------------------------------- #
# Shared helpers                                                               #
# --------------------------------------------------------------------------- #
def _dump_fields(page, label: str) -> None:
    """Print every input / textarea / select / button on the current page with
    its id, name, type, placeholder and visible text, so selectors can be read
    off and put in .env."""
    fields = page.evaluate(
        """() => {
            const out = [];
            for (const el of document.querySelectorAll('input, textarea, select, button, a[href]')) {
                out.push({
                    tag: el.tagName.toLowerCase(),
                    type: el.getAttribute('type') || '',
                    id: el.id || '',
                    name: el.getAttribute('name') || '',
                    placeholder: el.getAttribute('placeholder') || '',
                    value: (el.value || '').slice(0, 40),
                    text: (el.innerText || el.textContent || '').trim().slice(0, 40),
                });
            }
            return out;
        }"""
    )
    print(f"\n=== {label} ({len(fields)} controls) — {page.url} ===")
    for f in fields:
        sel = f"#{f['id']}" if f["id"] else (f"[name=\"{f['name']}\"]" if f["name"] else "(no id/name)")
        descr = f["placeholder"] or f["text"] or f["value"]
        print(f"  {f['tag']:<7} type={f['type']:<9} {sel:<36} {descr}")


def _advance_to_form(page, order: str) -> bool:
    """Page 1 -> page 2: type the order # and submit. Returns True if it tried
    (i.e. the order # selector is configured). Clicks the configured advance
    control, or presses Enter if none is set."""
    if not EMAIL_DRAWINGS_ORDER_SELECTOR:
        return False
    page.fill(EMAIL_DRAWINGS_ORDER_SELECTOR, str(order))
    if EMAIL_DRAWINGS_ORDER_SUBMIT_SELECTOR:
        page.click(EMAIL_DRAWINGS_ORDER_SUBMIT_SELECTOR)
    else:
        page.press(EMAIL_DRAWINGS_ORDER_SELECTOR, "Enter")
    try:
        page.wait_for_load_state("networkidle", timeout=20000)
    except Exception:  # noqa: BLE001 - AJAX app may not settle
        page.wait_for_timeout(1500)
    return True


# --------------------------------------------------------------------------- #
# Probe — read-only field discovery across both pages                         #
# --------------------------------------------------------------------------- #
def probe(order: Optional[str] = None, headless: bool = False) -> None:
    """Walk the two-page flow read-only and dump each page's controls. Advancing
    to page 2 (typing the order # and submitting) is just navigation — it does
    NOT send anything; the Send button on page 2 is never clicked."""
    _require_session()
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(storage_state=str(STORAGE_STATE_PATH))
        page = context.new_page()
        try:
            _goto_form(page, order)
            if not EMAIL_DRAWINGS_URL:
                print("\n! EMAIL_DRAWINGS_URL is not set in .env, so this opened the CBC landing")
                print("  page. Navigate to Engineering -> Email Drawings in the open window, then")
                print("  re-run --probe with that URL in EMAIL_DRAWINGS_URL for an accurate dump.\n")

            _dump_fields(page, "PAGE 1 (order lookup)")
            print("\nFrom page 1, set EMAIL_DRAWINGS_ORDER_SELECTOR (the order # input) and, if a")
            print("button — not Enter — advances, EMAIL_DRAWINGS_ORDER_SUBMIT_SELECTOR.")

            if order and _advance_to_form(page, order):
                _dump_fields(page, "PAGE 2 (email form)")
                print("\nFrom page 2, set EMAIL_DRAWINGS_EMAILS_SELECTOR (recipients),")
                print("EMAIL_DRAWINGS_ATTACH_SELECTOR (the file input), and")
                print("EMAIL_DRAWINGS_SUBMIT_SELECTOR (the Send button — recorded so we know what")
                print("NOT to click; sending stays disabled).")
            elif order:
                print("\n(Set EMAIL_DRAWINGS_ORDER_SELECTOR first, then re-run with --order to reach")
                print(" page 2 and dump the email form.)")

            if not headless:
                input("\nReview the open window, then press Enter to close it...")
        finally:
            browser.close()


# --------------------------------------------------------------------------- #
# Pre-fill — fills the form but NEVER submits                                  #
# --------------------------------------------------------------------------- #
def prepare(order: str, emails: List[str], files: List[str], headless: bool = False) -> None:
    """Navigate, fill the order #, paste the emails (';'-joined), and attach the
    files — then stop with the form on screen for a human. Does not submit."""
    _require_session()
    if not EMAIL_DRAWINGS_URL:
        raise RuntimeError(
            "EMAIL_DRAWINGS_URL is not set in .env. Run `python email_drawings.py --probe` "
            "first, find the Email Drawings page URL + field selectors, and set them."
        )
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(storage_state=str(STORAGE_STATE_PATH))
        page = context.new_page()
        try:
            _goto_form(page, order)

            # Page 1: type the order # and advance to the email form (page 2).
            if not _advance_to_form(page, order):
                log.warning("EMAIL_DRAWINGS_ORDER_SELECTOR not set — could not advance to the "
                            "email form. Run `--probe` and set the page-1 selectors first.")

            if EMAIL_DRAWINGS_EMAILS_SELECTOR:
                page.fill(EMAIL_DRAWINGS_EMAILS_SELECTOR, "; ".join(emails))
            else:
                log.warning("EMAIL_DRAWINGS_EMAILS_SELECTOR not set — skipping emails fill.")

            if EMAIL_DRAWINGS_ATTACH_SELECTOR:
                for f in files:
                    if Path(f).exists():
                        page.set_input_files(EMAIL_DRAWINGS_ATTACH_SELECTOR, f)
                    else:
                        log.warning("Attachment not found, skipping: %s", f)
            else:
                log.warning("EMAIL_DRAWINGS_ATTACH_SELECTOR not set — skipping attachments.")

            # ----------------------------------------------------------------- #
            # SEND IS HARD-DISABLED. The submit click below is intentionally     #
            # commented out so a transmittal can never be mailed by accident.    #
            # To enable sending you must (1) flip _SEND_HARD_DISABLED to False   #
            # AND (2) uncomment the submit block — a deliberate two-step edit.   #
            # ----------------------------------------------------------------- #
            if not _SEND_HARD_DISABLED and TRANSMITTAL_MODE == "send":
                raise RuntimeError("Sending is disabled in this build.")
            #     if EMAIL_DRAWINGS_SUBMIT_SELECTOR:
            #         page.click(EMAIL_DRAWINGS_SUBMIT_SELECTOR)   # <-- the actual SEND
            #         page.wait_for_load_state("networkidle", timeout=30000)
            #         log.info("Submitted Email Drawings form for order %s", order)

            print(f"\nPre-filled the Email Drawings form for order {order} "
                  f"({len(emails)} recipient(s), {len(files)} file(s)).")
            print("SEND is disabled — review the form and submit it yourself if it's correct.")
            if not headless:
                input("Press Enter to close the window...")
        finally:
            browser.close()


# --------------------------------------------------------------------------- #
# CLI                                                                          #
# --------------------------------------------------------------------------- #
def main(argv: Optional[List[str]] = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    ap = argparse.ArgumentParser(
        description="CBC Insider Email Drawings probe / pre-fill (never sends).")
    ap.add_argument("order", nargs="?", help="order number to pre-fill the form for")
    ap.add_argument("--probe", action="store_true", help="dump the form's fields/selectors (read-only)")
    ap.add_argument("--headless", action="store_true", help="run without a visible window")
    args = ap.parse_args(argv)

    if args.probe:
        probe(order=args.order, headless=args.headless)
        return 0

    if not args.order:
        ap.error("give an order number, or use --probe")

    # Gather the recipients + files from the order, then pre-fill (no send).
    import transmittal_data as td
    data = td.gather(args.order)
    if not data.emails:
        print("! No recipient emails found for this order — nothing to pre-fill. "
              "Check the Sales Order's Additional Features / Notes block.")
        for w in data.warnings:
            print(f"  ! {w}")
        return 1
    prepare(args.order, data.emails, data.attachments, headless=args.headless)
    return 0


if __name__ == "__main__":
    sys.exit(main())
