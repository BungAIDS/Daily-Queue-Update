# Daily Queue Update

Logs into cbcinsider.com every morning at 5 AM, scrapes the work queue, compares it to yesterday, asks Claude for a natural-language briefing + anomaly flags + ranked action items, builds an Excel report, and emails you a plain-text summary.

## What it produces each run

- `queue_YYYY-MM-DD.xlsx` — two-tab Excel report in your `OUTPUT_DIR`:
  - **Changes** (first tab): AI briefing, anomalies, top action items, new orders, removed/completed orders, changed orders (with old → new values), persistent orders (3+ consecutive days in queue).
  - **Full Queue**: one row per job, AutoFilter enabled, red highlight for today/overdue End Dates, yellow for due within 3 days, summary row at the bottom with total job count and total dollar value.
- `snapshots/queue_YYYY-MM-DD.json` — full structured snapshot used for tomorrow's diff.
- A plain-text email with the briefing, counts, top action items, anomalies, and the Excel file path.
- An alert email if any step fails (login failure, site down, Claude API error, etc.).

## One-time setup

### 1. Install Python 3.11+ and dependencies

```bash
cd Daily-Queue-Update
python -m venv venv
# Windows:
venv\Scripts\activate
# Mac/Linux:
source venv/bin/activate

pip install -r requirements.txt
playwright install chromium
```

### 2. Get an Anthropic API key

Sign up at https://console.anthropic.com, add a payment method, and create an API key. This script uses `claude-opus-4-7` — expect ~$0.05–0.20 per daily run (a few pennies to a couple dimes depending on queue size).

### 3. Configure `.env`

Copy `.env.example` to `.env` and fill in every field:

```
CBC_USERNAME=your_cbc_login
CBC_PASSWORD=your_cbc_password
ANTHROPIC_API_KEY=sk-ant-...
OUTPUT_DIR=C:\Users\you\Documents\DailyQueue
SNAPSHOT_DIR=C:\Users\you\Documents\DailyQueue\snapshots
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=you@gmail.com
SMTP_PASSWORD=your_app_password
EMAIL_FROM=you@gmail.com
EMAIL_TO=you@gmail.com
EMAIL_ALERT_TO=you@gmail.com
```

**Gmail:** use an [App Password](https://myaccount.google.com/apppasswords) (not your real Gmail password). Requires 2FA on the account.
**Outlook:** set `SMTP_HOST=smtp.office365.com`, `SMTP_PORT=587`, and use your Outlook password (or an app password if MFA is on).

### 4. Test it manually

```bash
python main.py
```

Expected: console logs through scrape → diff → Claude → Excel → email, file appears in `OUTPUT_DIR`, email arrives.

If the scraper login fails, the CSS selectors in `scraper.py` may need tweaking for cbcinsider's current login page — open `scraper.py` and adjust the `page.fill(...)` / `page.click(...)` selectors.

## Scheduling at 5 AM daily

### Windows — Task Scheduler

1. Open **Task Scheduler** → **Create Basic Task**.
2. Name: `Daily Queue Update`. Trigger: **Daily**, start time **5:00 AM**.
3. Action: **Start a program**.
   - **Program/script:** the full path to `python.exe` inside your venv, e.g. `C:\path\to\Daily-Queue-Update\venv\Scripts\python.exe`
   - **Add arguments:** `main.py`
   - **Start in:** the full path to the project folder, e.g. `C:\path\to\Daily-Queue-Update`
4. Finish the wizard, then right-click the task → **Properties**:
   - Under **General**: check "Run whether user is logged on or not" and "Run with highest privileges".
   - Under **Settings**: check "Run task as soon as possible after a scheduled start is missed" (covers reboots).
5. Test it: right-click → **Run**. Check that the Excel file appears and the email arrives.

### Mac/Linux — cron

1. Make `main.py` runnable:
   ```bash
   chmod +x /full/path/to/Daily-Queue-Update/main.py
   ```
2. Edit your crontab:
   ```bash
   crontab -e
   ```
3. Add this line (runs daily at 5:00 AM):
   ```
   0 5 * * * cd /full/path/to/Daily-Queue-Update && /full/path/to/Daily-Queue-Update/venv/bin/python main.py >> /full/path/to/Daily-Queue-Update/cron.log 2>&1
   ```
4. Save and exit. Verify with `crontab -l`.
5. Mac users: macOS may need to grant cron Full Disk Access under **System Settings → Privacy & Security → Full Disk Access** so it can write to your output folder.

## File layout

```
Daily-Queue-Update/
├── main.py             # Entrypoint — orchestrates the whole run
├── scraper.py          # Playwright login + queue table parser (paired-row handling)
├── compare.py          # Diff today vs yesterday; persistence tracking
├── analyzer.py         # Claude API call — briefing + anomalies + action items
├── excel_writer.py     # Two-tab .xlsx report with AutoFilter and date highlights
├── emailer.py          # Plain-text email + failure alert email
├── config.py           # Loads .env
├── requirements.txt
├── .env.example
└── README.md
```

## Troubleshooting

- **Login fails:** Open `scraper.py` and look at the `page.fill(...)` selectors near the top of `scrape_queue()`. cbcinsider may use different input names — inspect the login form in a browser dev tools and update.
- **Scraper returns 0 jobs:** The table-row selector (`table tbody tr`) may not match. Run with `headless=False` (edit `main.py` to pass `headless=False` to `scrape_queue`) to watch what happens, and adjust selectors.
- **Claude returns invalid JSON:** Rare, but the analyzer raises and the script falls back to an empty briefing — you still get the Excel report and email. Check the alert email for the raw output.
- **No email arrives but no alert either:** Check the SMTP credentials. Gmail app passwords expire if you change your account password.
- **Date highlighting wrong:** `excel_writer._parse_date` tries `MM/DD/YYYY`, `MM/DD/YY`, `YYYY-MM-DD`. If cbcinsider uses a different format, add it to that list.
