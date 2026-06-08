# Daily Queue Update

Logs into cbcinsider.com every morning at 5 AM, scrapes the work queue, compares it to yesterday, asks Claude for a natural-language briefing + anomaly flags + ranked action items, builds an Excel report, and emails you a plain-text summary through your desktop Outlook.

> **Platform:** Windows with the Outlook desktop app installed and signed in. Email is sent through Outlook (no password stored), which is Windows-only. See the Mac/Linux note under Scheduling if you're not on Windows.

## What it produces each run

- `queue_YYYY-MM-DD.xlsx` — two-tab Excel report in your `OUTPUT_DIR`:
  - **Changes** (first tab): AI briefing, anomalies, top action items, new orders, removed/completed orders, changed orders (with old → new values), persistent orders (3+ consecutive days in queue).
  - **Full Queue**: one row per job, AutoFilter enabled, red highlight for today/overdue End Dates, yellow for due within 3 days, summary row at the bottom with total job count and total dollar value.
- `snapshots/queue_YYYY-MM-DD.json` — full structured snapshot used for tomorrow's diff.
- A plain-text email (via your desktop Outlook) with the briefing, counts, top action items, anomalies, and the Excel report attached.
- An alert email if any step fails (login failure, site down, Claude API error, etc.).

## One-time setup

### 1. Install Python 3.11+ and dependencies

You need the **Outlook desktop app installed and signed in** for email to work.

```bash
cd Daily-Queue-Update
python -m venv venv
# Windows:
venv\Scripts\activate

pip install -r requirements.txt
playwright install chromium
```

`pip install` pulls in `pywin32` (the Outlook bridge) automatically on Windows.

### 2. Get an Anthropic API key

Sign up at https://console.anthropic.com, add a payment method, and create an API key. This script uses `claude-opus-4-7` — expect ~$0.05–0.20 per daily run (a few pennies to a couple dimes depending on queue size).

### 3. Configure `.env`

Copy `.env.example` to `.env` and fill in every field:

```
ANTHROPIC_API_KEY=sk-ant-...
CBC_QUEUE_URL=                 # recommended: your dispatch.aspx page URL
CBC_WORK_CENTER=              # optional: e.g. ENGGL (guards against wrong queue)
STORAGE_STATE_PATH=./cbc_session.json
OUTPUT_DIR=C:\Users\you\Documents\DailyQueue
SNAPSHOT_DIR=C:\Users\you\Documents\DailyQueue\snapshots
EMAIL_TO=you@company.com
EMAIL_ALERT_TO=you@company.com
```

**No passwords are stored.** cbcinsider uses your saved session (step 4); email
goes through your signed-in Outlook desktop app, so no email password is needed
either. `EMAIL_TO` / `EMAIL_ALERT_TO` are just addresses — an address alone
isn't sensitive.

### 4. Log in once (saves your session — no password stored)

```bash
python login.py
```

A browser window opens. Log into cbcinsider yourself (with your normal
2FA/SSO), navigate to your work queue, then return to the terminal and press
Enter. Your session is saved to `cbc_session.json` and reused every day.

> **Important:** logged-in sessions eventually expire. When they do, the 5 AM
> run fails and emails you an alert — just run `python login.py` again to
> refresh it. (Tip: if cbcinsider has a "remember me" checkbox, tick it to
> make the session last longer.)

### 5. Test it manually

```bash
python main.py
```

Expected: console logs through scrape → diff → Claude → Excel → email, file appears in `OUTPUT_DIR`, email arrives.

If it lands on the login page, your session expired — re-run `python login.py`.
If it returns 0 jobs, set `CBC_QUEUE_URL` to your exact dispatch page URL (and
check `scraper.py`'s selectors against cbcinsider's current layout) — see
Troubleshooting below.

### Recovering a botched run (staged execution)

If the 5 AM run goes wrong, you don't have to re-run the whole thing. The
pipeline can run one stage at a time, reusing what the previous stage wrote to
disk so the slow scrape happens only once:

```bash
python main.py --no-ai      # 1. scrape + diff + Excel — no AI, no email
python main.py --ai-only    # 2. add the AI briefing to today's saved run — no email
python main.py --mail-only  # 3. email today's finished briefing + Excel
```

A plain `python main.py` still does all three in one shot. The history/tracking
state is advanced exactly once — by whichever stage does the scrape+diff
(`--no-ai` or a full run) — so re-running `--ai-only` or `--mail-only` is safe
and never double-counts. Run the stages in order; each tells you the next one.


## Scheduling at 5 AM daily

### Windows — Task Scheduler

1. Open **Task Scheduler** → **Create Basic Task**.
2. Name: `Daily Queue Update`. Trigger: **Daily**, start time **5:00 AM**.
3. Action: **Start a program**.
   - **Program/script:** the full path to `python.exe` inside your venv, e.g. `C:\path\to\Daily-Queue-Update\venv\Scripts\python.exe`
   - **Add arguments:** `main.py`
   - **Start in:** the full path to the project folder, e.g. `C:\path\to\Daily-Queue-Update`
4. Finish the wizard, then right-click the task → **Properties**:
   - Under **General**: select **"Run only when user is logged on"**. This is required — the Outlook desktop app can only be automated inside your interactive desktop session. ("Run whether user is logged on or not" will fail to send email.)
   - Under **Settings**: check "Run task as soon as possible after a scheduled start is missed" (covers reboots).
5. Test it: right-click → **Run**. Check that the Excel file appears and the email arrives.

> **Because email needs your logged-in session,** your computer must be **on and logged in** at 5 AM (it can be locked — locked is fine, logged-out is not). If your machine is usually off overnight, either leave it on, or schedule the task for a time you're logged in.

### Mac/Linux — cron

> **Note:** the Outlook email step is Windows-only. On Mac/Linux the scrape, diff, Claude analysis, and Excel report all still work, but you'd need a different notification method (ask me to swap the emailer for a file-drop or another option). The cron mechanics below are otherwise correct.

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
├── login.py            # Run once — log in by hand, save session (no password stored)
├── main.py             # Entrypoint — orchestrates the whole run
├── scraper.py          # Reuses saved session + dispatch parser (per-order container, job+detail rows)
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

- **Lands on the login page / "session expired":** Your saved session ran out. Run `python login.py` again to refresh it. Sessions expire periodically — this is expected.
- **Scraper returns 0 jobs:** Most often `CBC_QUEUE_URL` isn't set to your exact dispatch page (the URL ending in `dispatch.aspx`) — set it in `.env`. The parser keys off the per-order containers (`div[id^="MainContent_rptDispatch_Container_"]`); if cbcinsider changes that markup, run with `headless=False` (edit `main.py` to pass `headless=False` to `scrape_queue`) to watch what happens and adjust the selectors in `scraper.py`. If the log warns "Page reports N results but parsed M", the row markup drifted.
- **Scraping the wrong Work Center:** The dispatch page is filtered by Work Center and the site remembers your last pick. Set `CBC_WORK_CENTER` (e.g. `ENGGL`) in `.env` so the run aborts loudly instead of diffing the wrong queue.
- **Claude returns invalid JSON:** Rare, but the analyzer raises and the script falls back to an empty briefing — you still get the Excel report and email. Check the alert email for the raw output.
- **No email arrives but no alert either:** Make sure the Outlook desktop app is installed, signed in, and open. The script controls Outlook through your logged-in session, so the task must run while you're logged in (see scheduling note). If Outlook shows a security prompt the first time, allow it.
- **Date highlighting wrong:** `excel_writer._parse_date` tries `MM/DD/YYYY`, `MM/DD/YY`, `YYYY-MM-DD`. If cbcinsider uses a different format, add it to that list.
