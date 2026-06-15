# Daily Queue Update

Logs into cbcinsider.com every morning at 5 AM, scrapes the work queue, compares it to yesterday, asks Claude for a natural-language briefing + anomaly flags + ranked action items, builds an Excel report, and emails you a plain-text summary through your desktop Outlook.

> **Platform:** Windows with the Outlook desktop app installed and signed in. Email is sent through Outlook (no password stored), which is Windows-only. See the Mac/Linux note under Scheduling if you're not on Windows.

## What it produces each run

- `queue_YYYY-MM-DD.xlsx` — two-tab Excel report in your `OUTPUT_DIR`:
  - **Changes** (first tab): AI briefing, anomalies, top action items, new orders, removed/completed orders, changed orders (with old → new values), persistent orders (3+ consecutive days in queue).
  - **Full Queue**: one row per job, AutoFilter enabled, red highlight for today/overdue End Dates, yellow for due within 3 days, summary row at the bottom with total job count and total dollar value.
- `snapshots/queue_YYYY-MM-DD.json` — full structured snapshot used for tomorrow's diff.
- Housekeeping: reports/snapshots/diffs older than 60 days are **moved into `archive/` subfolders** (under `OUTPUT_DIR` and `SNAPSHOT_DIR`) — never deleted, so the complete record of every order stays on disk.
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

Sign up at https://console.anthropic.com, add a payment method, and create an API key. This script uses `claude-haiku-4-5` — expect ~$0.02–0.05 per daily run (a few pennies, depending on queue size). The model is set by `CLAUDE_MODEL` in `config.py`.

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

### The four scripts

`main.py` is the once-a-day job. The other three are its individual stages, so
if the 5 AM run goes wrong you can re-run just the part you need without redoing
the slow scrape:

```bash
python main.py     # everything: scrape -> AI overview -> Excel -> email  (the 5 AM job)

python scrape.py   # 1. scrape + diff + Excel        (no AI, no email)
python brief.py    # 2. add the AI overview           (no email; reuses today's scrape)
python send.py     # 3. email the most recent report
```

`main.py` runs stages 1→2→3 in one shot. Each stage reuses what the previous one
wrote to disk, so `brief.py` never re-scrapes — it just makes the one Claude
call. `send.py` picks the newest report that actually has an AI overview (pass a
path, or `--dry-run`, to override). The history/tracking state is advanced
exactly once — by whichever run does the scrape (`scrape.py` or `main.py`) — so
re-running `brief.py`/`send.py` is safe and never double-counts.


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
├── main.py             # The 5 AM job — runs scrape -> brief -> send in one shot
├── scrape.py           # Stage 1 — scrape + diff + Excel (no AI, no email)
├── brief.py            # Stage 2 — add the AI overview to today's run (no email)
├── send.py             # Stage 3 — email the most recent report
├── pipeline.py         # Shared stage logic the four scripts above call into
├── scraper.py          # Reuses saved session + dispatch parser (per-order container, job+detail rows)
├── sales_orders.py     # Per-job enrichment: Sales Order + construction/drive run + folder
├── templates.py        # Quote-run TEMPLATE collection — match a run by design#/format, pull its fields
├── drive_run.py        # PDF quote-run reader (the .pdf template in templates.py)
├── compare.py          # Diff today vs the most recent prior run; persistence tracking
├── analyzer.py         # Claude API call — briefing + anomalies + action items
├── excel_writer.py     # Two-tab .xlsx report with AutoFilter and date highlights
├── emailer.py          # Plain-text email + failure alert email
├── runstate.py         # Persists each day's diff/briefing/Excel so stages can hand off
├── config.py           # Loads .env
├── autocad_scan.py     # Backlog tool — sweep AutoCAD folders, record each fan's custom DWGs
├── backfill_orders.py  # Backlog tool — look up old orders 1-by-1 (resumable)
├── line_items.py       # SO line items — capture/normalize/tag logic + the lookup store
├── line_items_scan.py  # Build the line-items store from archived SO PDFs; tune the rules
├── find_orders.py      # Search orders by their line items (CLI + Excel inventory)
├── discover_documents.py  # Discovery — list a job's docs; probe how to reach old orders
├── requirements.txt
├── .env.example
└── README.md
```

## Custom fans & the backlog (construction runs + AutoCAD DWGs)

Three extra capabilities, on top of the daily run:

### Construction / quote run

The daily enrichment now grabs the **quote run** alongside the Sales Order.
Quote runs are recognized three ways, in order:

1. **Dedicated pid type** — only the HDX fans file the run under its own
   document type (`DRIVE_RUN_TYPES`, default `CBC_DriveRun,CBC_QuoteRun`, plus
   any other type ending in `Run` as a fallback).
2. **File name** — everything else files it under a generic type (usually
   `CBC_Inquiry`), so document names are matched against
   `DRIVE_RUN_NAME_PATTERNS` (default catches `... Qt Run.txt`,
   `... Quote Run ...`, and the D64 `... D64 Wheel Construction ....xlsx`).
3. **AutoCAD folder** — some orders never get the run attached to their
   documents at all; the job's folder is searched recursively for the same
   name patterns (e.g. `ENG REF\420410 qt  run.txt`) and the report links the
   file in place.

Every matching document is archived (keeping its real extension — runs come as
`.txt`, `.xlsx`, `.rtf`, or `.pdf`) under `DRIVE_RUN_DIR`. If zero runs match,
the log prints every pid type it saw on the board. Only highly-custom fans have
a run, so its presence is a signal in itself: the Full Queue tab gains a
**Quote Run** column — `YES`, hyperlinked straight to the run file; it shows a
plain `YES` only if the flag is set but no file was reachable.

**The quote-run "template" collection.** A quote run is not one format — which
one you get is mostly a function of the fan's **design number** (Design 64 →
a "D64 Wheel Construction" `.xlsx`, HDX → a plain-text "Qt Run", others → a
`.pdf`/`.rtf`). So reading them is a *collection of templates* in
`templates.py`: each template declares which runs it recognizes (by design #,
file extension, and/or file-name marker) and how to pull fields from that one
shape. The daily run calls `parse_quote_run(file, design=...)`, which picks the
best-matching template and returns its fields + a compact summary (shown next
to the `YES` flag). Adding a new fan format is just one more
`QuoteRunTemplate` in `templates.py` — see the "Adding a template" note there.

Until a real sample pins a format's exact headings down, each template does
resilient best-effort `Label: value` extraction, so unknown labels are still
captured. Inspect a job's docs and see which template matched:

```bash
python discover_documents.py <a-custom-job#>   # lists docs + pid types, downloads SO + run(s)
python templates.py "<path to a quote run>" [design#]  # shows the matched template + fields it pulled
python dump_pdf.py "<path to a pdf run>"       # dumps a pdf run's raw text/tables
```

Paste that back and the matching template's fields get pinned down. Until then
the report still shows the `YES` flag and a best-effort summary.

### AutoCAD DWG scan

Sweep the AutoCAD job folders and record each fan's custom drawings — fast,
filesystem-only, resumable:

```bash
python autocad_scan.py                 # every job folder under AUTOCAD_JOBS_DIR
python autocad_scan.py 421314 421388   # specific jobs
```

It writes `backlog/autocad_dwgs.xlsx`: one row per job with a column for **every
custom suffix** seen (`-51`, `-35`, …), each cell **green ✓** when the job has
that drawing and **red** when it doesn't, plus an `Extras` count. The standard
`-01`/`-02` (CW/CCW) drawings aren't shown — nearly every job has them, so they'd
just be noise — but a job missing **both** is flagged red as the rare exception.
Progress is saved after every batch, so an interrupted run resumes.

The **daily queue report carries the same green-✓/red matrix** on both the **Full
Queue** and **History** tabs, appended after the standard columns: every morning
the run scans each board job's AutoCAD folder live (reusing the folder lookup it
already does) and adds a column for each custom suffix found. History keeps it
per archived order too, so it builds into a complete per-order log over time
(orders archived once the scan went live carry their DWG data).

### Sales-order line items — capture, normalize, search

Every time a Sales Order is parsed (daily run, backfill, or the archive scan
below), **every line item on it** — the priced item/accessory rows and the
"Additional Features"-style lines — is captured into one lookup store
(`backlog/line_items.json`). The same option is rarely written the same way
twice ("SS SHAFT SLEEVE" / "Stainless Steel Shaft Sleeve" / "316SS sleeve"),
so each line is kept three ways:

- **raw** — exactly as printed on the SO (never altered, so rules can be
  re-tuned later without re-downloading anything),
- **normalized** — uppercased, the price columns and `L`/`C`/`N` type letter
  stripped, abbreviations expanded (`W/`→`WITH`, `SS`→`STAINLESS STEEL`,
  `IVD`→`INLET VANE DAMPER`, …) so variants converge,
- **details** — the unpriced continuation lines printed under an item
  (vendor, motor HP/enclosure, `Product: Damper`, …) — searchable, and they
  contribute to the item's tags,
- **tags** — canonical features (SHAFT SEAL, SPARK RESISTANT, COATING, …)
  matched by a rules table.

The daily report's **Full Queue / History tabs gain a "Features" column** with
each job's tags, and the AI briefing weaves notable features into its summary.

**Build the store from what's already archived** (no login, no browser — it
reads the PDFs under `SALES_ORDER_DIR`), then search:

```bash
python line_items_scan.py              # one local pass over the whole archive
python find_orders.py shaft seal       # orders whose SO matches BOTH terms
python find_orders.py --any teflon viton
python find_orders.py --tag "SHAFT SEAL"     # by canonical tag
python find_orders.py cermic felt --fuzzy    # typo-tolerant
python find_orders.py --job 421314     # what's stored for one job
python find_orders.py --list-tags      # the live tag vocabulary + counts
python find_orders.py --xlsx           # full inventory workbook (AutoFilter) —
                                       # filter line items straight in Excel
```

`--xlsx` writes two tabs: **Line Items** (one row per item) and a **Feature
Matrix** — one row per order, one column per feature tag, **green ✓** when the
order has that feature and **red** when it doesn't, exactly like the AutoCAD
DWG matrix. Searching first (`python find_orders.py shaft seal --xlsx`) limits
the matrix to the matching orders, but each row still shows that order's full
feature profile.

**Normalizing the long tail.** Order entry is free text, so plenty of lines
won't match any built-in rule at first. Three levers, in order:

1. **See what the extractor is doing** on a few real orders and paste the
   output back to get the capture/skip rules tuned:
   ```bash
   python line_items_scan.py --dump 421314 421473
   ```
   Each line is marked `ITEM $` (priced row), `ITEM +` (feature-section line),
   `skip [rule]`, or `.` (ignored), followed by the captured items with their
   normalized form and tags.
2. **Extend the rules** without touching code: point `LINE_ITEM_RULES` in
   `.env` at a small JSON file of site wording (extra abbreviations, skip
   patterns, tag patterns — see `.env.example`), then re-apply to everything
   already stored (raw text is kept verbatim, so this is instant and lossless):
   ```bash
   python line_items_scan.py --renorm
   ```
3. **Let Claude classify the rest.** Sends each still-untagged *unique* item
   to the API once (pennies on haiku), caches the answer forever in the store,
   and folds the tags in:
   ```bash
   python line_items_scan.py --ai
   ```

### Backfill old orders

Grind through a backlog of historical orders one at a time (run it all day):

```bash
python backfill_orders.py                  # all real AutoCAD job folders
python backfill_orders.py --list jobs.txt  # or a file of job numbers
python backfill_orders.py --range 420000 421000
```

A folder sweep only considers job numbers at or above `--min-job` (default
`400000`), so non-job folders (year/template/archive dirs with small or
non-numeric names) are skipped. Raise it once you know your exact lowest job,
e.g. `--min-job 403000` (and `--max-job` to cap the top). The same flags apply
to `autocad_scan.py`.

It downloads + parses each order's Sales Order and drive run, merges the DWG
scan, and writes `backlog/backlog.xlsx`. It's resumable (kill and re-run any
time).

Old orders are opened through the queue page's **"search order" / "find order"**
box — the backfill types each job number in and opens the surfaced order. The
box is auto-detected; the run preflights it and stops with a clear message (no
all-day grind) if it can't be found. Confirm it once, and grab the exact
selector if auto-detect misses:

```bash
python discover_documents.py --probe <a-real-job#>
```

That lists the page's text inputs, shows what auto-detect picked, and runs the
**real** lookup — `SUCCESS` means you're ready. If it misses, set
`CBC_SEARCH_SELECTOR` (and `CBC_SEARCH_BUTTON` if a button submits the search)
in `.env` to the selector it printed, then re-probe.

## Troubleshooting

- **Lands on the login page / "session expired":** Your saved session ran out. Run `python login.py` again to refresh it. Sessions expire periodically — this is expected.
- **Scraper returns 0 jobs:** Most often `CBC_QUEUE_URL` isn't set to your exact dispatch page (the URL ending in `dispatch.aspx`) — set it in `.env`. The parser keys off the per-order containers (`div[id^="MainContent_rptDispatch_Container_"]`); if cbcinsider changes that markup, run with `headless=False` (edit `main.py` to pass `headless=False` to `scrape_queue`) to watch what happens and adjust the selectors in `scraper.py`. If the log warns "Page reports N results but parsed M", the row markup drifted.
- **Scraping the wrong Work Center:** The dispatch page is filtered by Work Center and the site remembers your last pick. Set `CBC_WORK_CENTER` (e.g. `ENGGL`) in `.env` so the run aborts loudly instead of diffing the wrong queue.
- **Claude returns invalid JSON:** Rare, but the analyzer raises and the script falls back to an empty briefing — you still get the Excel report and email. Check the alert email for the raw output.
- **No email arrives but no alert either:** Make sure the Outlook desktop app is installed, signed in, and open. The script controls Outlook through your logged-in session, so the task must run while you're logged in (see scheduling note). If Outlook shows a security prompt the first time, allow it.
- **Date highlighting wrong:** `excel_writer._parse_date` tries `MM/DD/YYYY`, `MM/DD/YY`, `YYYY-MM-DD`. If cbcinsider uses a different format, add it to that list.
