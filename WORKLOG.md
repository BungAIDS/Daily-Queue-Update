# Work Log — Construction Run + Backfill + AutoCAD DWG Scan

Running notes so progress survives across sessions. Newest status at the top of
each section. **If you're picking this up fresh, read this whole file first.**

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
- [ ] **Wire the confirmed old-order lookup** into `backfill_orders.py` (blocked
      on discovery step 2).
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
