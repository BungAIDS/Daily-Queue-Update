"""One-time (and occasional) login helper.

Run this manually: `python login.py`

It opens a real browser window. You log into cbcinsider yourself — by hand,
with whatever 2FA/SSO your company uses — and navigate to your work queue.
Then you come back to this terminal and press Enter. Your session cookies are
saved to cbc_session.json, and the daily script reuses them.

No password is stored anywhere — only the logged-in session token, on your
own machine. When the session expires, just run this again.
"""
from __future__ import annotations

from playwright.sync_api import sync_playwright

from config import CBC_URL, CBC_QUEUE_URL, STORAGE_STATE_PATH


def main() -> None:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # visible window so you can log in
        context = browser.new_context()
        page = context.new_page()
        page.goto(CBC_QUEUE_URL or CBC_URL, wait_until="domcontentloaded")

        print("\n" + "=" * 70)
        print("A browser window has opened.")
        print("  1. Log into cbcinsider in that window.")
        print("  2. Navigate to your work queue so it's showing.")
        print("  3. Come back here and press Enter.")
        print("=" * 70)
        input("\nPress Enter once you're logged in and your queue is visible... ")

        context.storage_state(path=str(STORAGE_STATE_PATH))
        print(f"\nSession saved to {STORAGE_STATE_PATH}")
        print("You can now run `python main.py`. Re-run this login if it ever expires.")
        browser.close()


if __name__ == "__main__":
    main()
