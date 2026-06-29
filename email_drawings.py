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
from urllib.parse import urljoin

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


def _click_email_drawings_link(page) -> bool:
    """Fallback for instances where EMAIL_DRAWINGS_URL isn't set: reach the form
    by following the 'Email Drawings' nav link. Reads the link's href straight
    from the DOM (more reliable than an accessibility-role lookup against a
    hover menu) and navigates; falls back to a JS click for __doPostBack links."""
    try:
        href = page.evaluate(
            """() => {
                const a = Array.from(document.querySelectorAll('a'))
                    .find(x => (x.innerText || x.textContent || '').trim() === 'Email Drawings');
                return a ? (a.getAttribute('href') || '') : null;
            }"""
        )
    except Exception:  # noqa: BLE001
        href = None
    if href is None:
        return False
    try:
        if href and not href.lower().startswith("javascript"):
            page.goto(urljoin(page.url, href), wait_until="domcontentloaded", timeout=30000)
        else:
            page.evaluate(
                """() => {
                    const a = Array.from(document.querySelectorAll('a'))
                        .find(x => (x.innerText || x.textContent || '').trim() === 'Email Drawings');
                    if (a) a.click();
                }"""
            )
            page.wait_for_load_state("domcontentloaded", timeout=30000)
        try:
            page.wait_for_load_state("networkidle", timeout=8000)
        except Exception:  # noqa: BLE001
            pass
        return True
    except Exception as e:  # noqa: BLE001
        log.warning("Could not follow the Email Drawings link: %s", e)
        return False


def _goto_form(page, order: Optional[str] = None) -> None:
    """Navigate to the Email Drawings page. If EMAIL_DRAWINGS_URL is set, go
    straight there; otherwise load the CBC landing page and follow the
    'Email Drawings' nav link so the probe still reaches the real form."""
    target = EMAIL_DRAWINGS_URL or CBC_URL
    log.info("Loading with saved session: %s", target)
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
    if not EMAIL_DRAWINGS_URL:
        if _click_email_drawings_link(page):
            print(f"\n>>> Email Drawings form is at: {page.url}")
            print(">>> Set EMAIL_DRAWINGS_URL to that in .env so future runs go straight there.\n")
        else:
            print("\n! Could not auto-find the 'Email Drawings' link. Navigate there by hand in")
            print("  the open window, copy the URL, and set EMAIL_DRAWINGS_URL in .env.\n")


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


import re as _re

# Heuristics for auto-detecting each control by its id/name/placeholder/label.
_ORDER_RX = _re.compile(r"order|job|search|\bso\b|\bwo\b|sales", _re.I)
_ADVANCE_RX = _re.compile(r"\bgo\b|find|search|lookup|open\s*order|continue|next|view|display|\bok\b", _re.I)
# The recipients box (e.g. #MainContent_txtTo) — match a trailing/standalone
# "to", "recipient", "email", "mail to". NOT the body textarea.
_RECIP_RX = _re.compile(r"recipient|e-?mail|mail\s*to|send\s*to|txt_?to|to\b|address", _re.I)
_SEND_RX = _re.compile(r"send|transmit", _re.I)


def _detect(page) -> dict:
    """Return the page's candidate controls with a computed CSS selector for
    each (prefer #id, else tag[name=...]). Visible elements only, so hidden
    ASP.NET state fields and off-screen menu links are ignored."""
    return page.evaluate(
        r"""() => {
            const visible = el => !!(el.offsetParent || el.getClientRects().length);
            const sel = el => {
                if (el.id) return '#' + CSS.escape(el.id);
                const n = el.getAttribute('name');
                if (n) return el.tagName.toLowerCase() + '[name="' + n + '"]';
                return null;
            };
            const desc = el => ((el.getAttribute('placeholder') || '') + ' ' +
                                (el.getAttribute('aria-label') || '') + ' ' +
                                (el.id || '') + ' ' + (el.getAttribute('name') || '') + ' ' +
                                (el.value || '') + ' ' + (el.innerText || el.textContent || '')).trim();
            const grab = q => Array.from(document.querySelectorAll(q))
                .filter(visible).map(el => ({sel: sel(el), desc: desc(el).slice(0, 80),
                                             type: (el.getAttribute('type') || '').toLowerCase()}))
                .filter(o => o.sel);
            return {
                texts: grab('input:not([type]), input[type=text], input[type=search], input[type=number]'),
                textareas: grab('textarea'),
                files: grab('input[type=file]'),
                buttons: grab('button, input[type=submit], input[type=button]'),
            };
        }"""
    )


def _pick(cands: list, rx, fallback: bool = True) -> Optional[str]:
    """First candidate whose description matches rx. Falls back to the first
    candidate only when `fallback` is True (off for fields where a wrong guess —
    e.g. grabbing the body textarea as the recipients box — is worse than none)."""
    for c in cands:
        if rx.search(c.get("desc", "")):
            return c["sel"]
    if fallback:
        return cands[0]["sel"] if cands else None
    return None


def _guess_page1(d: dict) -> dict:
    """Best-guess selectors for the order-lookup page."""
    return {
        "order": EMAIL_DRAWINGS_ORDER_SELECTOR or _pick(d["texts"], _ORDER_RX),
        "advance": EMAIL_DRAWINGS_ORDER_SUBMIT_SELECTOR or _pick(d["buttons"], _ADVANCE_RX),
    }


def _guess_page2(d: dict) -> dict:
    """Best-guess selectors for the email form page. The recipients box is a text
    input (e.g. #MainContent_txtTo), so match texts first and never fall back to
    the body textarea — a wrong recipients guess is worse than none."""
    emails = (EMAIL_DRAWINGS_EMAILS_SELECTOR
              or _pick(d["texts"], _RECIP_RX, fallback=False)
              or _pick(d["textareas"], _RECIP_RX, fallback=False))
    return {
        "emails": emails,
        "attach": EMAIL_DRAWINGS_ATTACH_SELECTOR or (d["files"][0]["sel"] if d["files"] else None),
        "send": EMAIL_DRAWINGS_SUBMIT_SELECTOR or _pick(d["buttons"], _SEND_RX, fallback=False),
    }


def _advance_to_form(page, order: str, order_sel: Optional[str],
                     advance_sel: Optional[str]) -> bool:
    """Page 1 -> page 2: type the order # into `order_sel` and submit (click
    `advance_sel`, or press Enter if none). Returns False if no order box is
    known."""
    if not order_sel:
        return False
    page.fill(order_sel, str(order))
    if advance_sel:
        page.click(advance_sel)
    else:
        page.press(order_sel, "Enter")
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
        env: dict = {}
        try:
            _goto_form(page, order)
            if not EMAIL_DRAWINGS_URL:
                env["EMAIL_DRAWINGS_URL"] = page.url

            # Page 1: auto-detect the order box + advance control.
            d1 = _detect(page)
            g1 = _guess_page1(d1)
            _dump_fields(page, "PAGE 1 (order lookup)")
            print(f"\n  -> detected order box : {g1['order'] or '(none found)'}")
            print(f"  -> detected advance   : {g1['advance'] or '(none — will press Enter)'}")
            env["EMAIL_DRAWINGS_ORDER_SELECTOR"] = g1["order"] or ""
            env["EMAIL_DRAWINGS_ORDER_SUBMIT_SELECTOR"] = g1["advance"] or ""

            # Page 2: advance with the test order, then auto-detect the email form.
            if order and _advance_to_form(page, order, g1["order"], g1["advance"]):
                d2 = _detect(page)
                g2 = _guess_page2(d2)
                _dump_fields(page, "PAGE 2 (email form)")
                print(f"\n  -> detected emails box: {g2['emails'] or '(none found)'}")
                print(f"  -> detected attach    : {g2['attach'] or '(none found)'}")
                print(f"  -> detected Send btn  : {g2['send'] or '(none found)'}  (never auto-clicked)")
                env["EMAIL_DRAWINGS_EMAILS_SELECTOR"] = g2["emails"] or ""
                env["EMAIL_DRAWINGS_ATTACH_SELECTOR"] = g2["attach"] or ""
                env["EMAIL_DRAWINGS_SUBMIT_SELECTOR"] = g2["send"] or ""
            elif order:
                print("\n(No order box detected on page 1, so couldn't advance to page 2.)")
            else:
                print("\n(Re-run with an order number — e.g. `--probe --order 421693` — to advance to")
                print(" page 2 and auto-detect the email form.)")

            # Hand back a ready-to-paste .env block of everything detected.
            if env:
                print("\n" + "=" * 64)
                print("Suggested .env (review the detected selectors above, then paste):")
                print("=" * 64)
                for k, v in env.items():
                    print(f"{k}={v}")
                print("=" * 64)

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
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(storage_state=str(STORAGE_STATE_PATH))
        page = context.new_page()
        try:
            _goto_form(page, order)

            # Page 1: type the order # and advance to the email form (page 2),
            # using configured selectors when set, else auto-detected ones.
            g1 = _guess_page1(_detect(page))
            if not _advance_to_form(page, order, g1["order"], g1["advance"]):
                log.warning("Could not find an order box to advance with. Run `--probe` to check "
                            "the page, or set EMAIL_DRAWINGS_ORDER_SELECTOR in .env.")

            g2 = _guess_page2(_detect(page))
            if g2["emails"]:
                page.fill(g2["emails"], "; ".join(emails))
            else:
                log.warning("No recipients field detected — skipping emails fill.")

            if g2["attach"]:
                for f in files:
                    if Path(f).exists():
                        page.set_input_files(g2["attach"], f)
                    else:
                        log.warning("Attachment not found, skipping: %s", f)
            else:
                log.warning("No file-attach control detected — skipping attachments.")

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
    ap.add_argument("--order", dest="order_opt", default=None,
                    help="order number (same as the positional; either works)")
    ap.add_argument("--probe", action="store_true", help="dump the form's fields/selectors (read-only)")
    ap.add_argument("--headless", action="store_true", help="run without a visible window")
    args = ap.parse_args(argv)
    order = args.order_opt or args.order

    if args.probe:
        probe(order=order, headless=args.headless)
        return 0

    if not order:
        ap.error("give an order number, or use --probe")

    # Gather the recipients + files from the order, then pre-fill (no send).
    import transmittal_data as td
    data = td.gather(order)
    if not data.emails:
        print("! No recipient emails found for this order — nothing to pre-fill. "
              "Check the Sales Order's Additional Features / Notes block.")
        for w in data.warnings:
            print(f"  ! {w}")
        return 1
    prepare(order, data.emails, data.attachments, headless=args.headless)
    return 0


if __name__ == "__main__":
    sys.exit(main())
