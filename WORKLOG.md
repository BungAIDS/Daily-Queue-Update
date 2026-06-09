# Work Log — Construction Run + Backfill + AutoCAD DWG Scan

Running notes so progress survives across sessions. Newest status at the top of
each section. **If you're picking this up fresh, read this whole file first.**

## 2026-06-09 — full-codebase review pass: bug fixes + archiving

Whole-repo bug/optimization review, fixes applied (model stays **haiku**):

- **send.py**: the Changes-tab section is headed "Orders that have changed",
  but the parser looked for "Changed orders" — re-sent emails always said
  `changed=0`. Fixed (verified with a repro before/after).
- **Downloads validated** (`sales_orders`/`backfill_orders`/`fetch_sales_orders`):
  a non-200 or non-`%PDF-` body (e.g. the login page after session expiry) is
  no longer written to the archive — previously it was cached forever by the
  `dest.exists()` skip and silently stood in for the order's PDF.
- **backfill_orders**: a found-but-failed download now records `error` (retried
  on resume) instead of `ok`; missing search box exits 1; per-job modal wait
  45s (serial run, typical load ~30s).
- **autocad_scan `--limit`** now counts folders *scanned this run* — before,
  already-done folders consumed the limit and a resumed run could do nothing.
- **compare.py**: brief.py's read-only recompute now reads the day-start
  history baseline (no more relabeling a returning order as "new" when the
  diff file is missing); lookback off-by-one fixed (13 → 14 days); duplicate
  job numbers on the board now log a warning instead of being silently merged.
- **sales_orders**: AutoCAD folder lookup is one sweep of the Z: tree instead
  of one glob per job; jobs whose modal loaded with no Sales Order (HDX) are
  terminal and skip the 90s wait on the retry pass.
- **Archiving (new)**: each scrape moves per-run files older than 60 days
  (`queue_*.json`, `diff_*`, `briefing_*`, `excel_*`, `history_*_start`, old
  `queue_*.xlsx`) into `archive/` subfolders — **moved, never deleted**; the
  goal is a complete record of every order. `history.json` is never touched.
  See `runstate.archive_old_runs` (`ARCHIVE_AFTER_DAYS = 60`).
- Minor: config placeholder-tuple duplicate removed; off-Windows default
  output dir is `~/Documents/DailyQueue` (no more literal `%USERPROFILE%`
  folder when running tests off the work machine); README model/cost note
  matches the configured haiku.

## The goal (from the request)

1. **Construction run (`CBC_DriveRun`)** — alongside the Sales Order, open the
   construction/drive run for each job and pull fields from it. It's the same
   kind of document as the Sales Order, identified by its pid *type* prefix
   (`CBC_DriveRun`, just like `CBC_SalesOrder`). Few orders have one; the ones
   that do are **highly custom** fans, so its presence is itself a signal.
2. **Backfill old orders** — a resumable, run-it-all-day tool that looks up
   historical orders on cbcinsider one at a time and fills out a backlog store.
3. **AutoCAD DWG scan** — sweep the AutoCAD job folders and record, per job,
   which custom drawings exist:
   - `<job>-01` = CW, `<job>-02` = CCW (PDF, DWG, or both). Every order has these.
   - Any other `<job>-NN` suffix (`-51`, `-35`, …) → a yes/no column **per
     distinct suffix** across all jobs.
4. **Don't lose progress** — every long runner writes incremental progress to
   disk and resumes; this file tracks build progress.

## Hard constraint: discovery happens on the Windows machine

The live site (cbcinsider) needs the saved session (`cbc_session.json`, never
committed) and the AutoCAD folders live on the `Z:` drive. Neither is reachable
from the dev sandbox, so anything that touches them is written to mirror the
existing, working `sales_orders.py` patterns and is **verified by you** running
the discovery scripts. The pure logic (filename parsing, suffix matrix, Excel
shaping) is unit-tested in the sandbox.

## Discovery steps you run (these reveal "the exact commands")

1. **Drive run document** — confirm how `CBC_DriveRun` shows up and what's in it:
   ```
   python discover_documents.py <a-highly-custom-job#>
   ```
   It lists every document for the job with its pid type / revision, flags the
   `CBC_SalesOrder` and `CBC_DriveRun` docs, and downloads both. Then dump the
   drive-run text so we can pick fields to parse:
   ```
   python dump_pdf.py "<path to the downloaded drive run pdf>"
   ```
   Paste both back here.
2. **Old-order lookup** — find how to open an order that's **not** on the board
   (needed for backfill). `discover_documents.py --probe <old-job#>` tries the
   plausible entry points (search box, direct `loadDetail`, URL params) and
   reports which one surfaces the documents. Paste the result back.

## Merged main's restructure (2026-06-08) — still compatible

Main was restructured (PR #2): orchestration moved out of `main.py` into
`pipeline.py` (stages `scrape`/`brief`/`send`), and `compare.py` now diffs
against the most recent prior snapshot (`prev_date`). I merged `origin/main`
into this branch. Compatibility confirmed:

- `pipeline.scrape_and_diff()` still calls **`enrich_with_sales_orders(jobs)`**
  (unchanged signature) → the construction-run capture runs automatically.
- `pipeline.build_excel()` still calls **`build_workbook(...)`** → the Drive Run
  column flows through. Merged `compare.py` emits the same diff keys
  `excel_writer` consumes.
- Only README + config.py overlapped; config auto-merged, README resolved.
- The four backlog/discovery tools don't touch `main.py`/`pipeline.py`.

All modules compile; scanner tests pass on the merged tree.

## Status

- [x] Read the codebase, mapped the data flow, wrote this log.
- [x] **AutoCAD DWG scanner** (`autocad_scan.py`) — pure logic + CLI + Excel
      output. Unit-tested in the sandbox. **(Feature 3 — done, runnable now.)**
- [x] **Generic document finding by pid type** in `sales_orders.py` — captures
      `CBC_SalesOrder` and `CBC_DriveRun` in one pass; adds `has_drive_run` +
      drive-run download. SO behavior unchanged. **(Feature 1 plumbing.)**
- [x] **Drive-run parse** (`drive_run.py`) — best-effort generic extraction;
      field list finalized after discovery dump. **(Feature 1 — refine post-discovery.)**
- [x] **"Drive Run" column** in the Excel report (YES = highly custom).
- [x] **`discover_documents.py`** — generalized discovery: all doc types,
      downloads SO + drive run, `--probe` for old-order lookup.
- [x] **`backfill_orders.py`** — resumable 1-by-1 runner with a single
      `open_order_detail()` seam to fill in once discovery reveals the lookup.
- [x] **Old-order lookup wired in** — backfill opens off-board orders via the
      queue page's "search order" box (auto-detected; `CBC_SEARCH_SELECTOR` /
      `CBC_SEARCH_BUTTON` overrides). Preflighted so it stops loudly instead of
      grinding if the box isn't found. `discover_documents.py --probe` runs the
      real path to confirm. (Confirm/tune selector on the work machine.)
- [ ] **Finalize drive-run fields** in `drive_run.py` (blocked on discovery step 1).
- [x] README + `.env.example` updated for the new paths/flags.

## Key defaults I chose (change freely)

- **Backfill job source**: defaults to enumerating the AutoCAD job folders
  (every `<type>/<intermediate>/<job>` under `AUTOCAD_JOBS_DIR`) — self-contained
  and it's the same list the DWG scan walks. Also supports `--list FILE` and
  `--range FIRST LAST`.
- **Outputs go to their own store**, not the daily queue report: the backlog is
  huge, so the DWG matrix and backfill data live in their own workbook/JSON under
  `BACKLOG_DIR` (defaults under `OUTPUT_DIR`). The daily report only gains the
  compact "Drive Run" column.
- **Extra-suffix columns** show `yes`/`no` (per the request). The richer PDF/DWG
  format is kept in the JSON store and shown for the CW/CCW columns.

## Confirmed conventions (from you)

- **DWG/PDF filename form: `<job>-<suffix><revletter>`** (e.g. `421314-01A`,
  `421314-51B`). The scanner captures only the digit suffix and drops the
  revision letter, so revisions of one drawing share a single column. Verified
  by `test_revision_letters`.

## Deferred (save for later)

- Verifying the scan against real folders/filenames (spot-check with
  `python autocad_scan.py <job#>`), and loosening the folder enumeration if
  needed.
- **Capture the revision letter** and track the latest rev per drawing — pairs
  naturally with opening the DWG (or its `-01/-02` PDF title block) to read who
  drafted and who checked it.
