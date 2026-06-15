# Work Log — Construction Run + Backfill + AutoCAD DWG Scan

Running notes so progress survives across sessions. Newest status at the top of
each section. **If you're picking this up fresh, read this whole file first.**

## 2026-06-15 — Quote-run TEMPLATE collection (`templates.py`)

Per DG: start a collection of templates that quote runs match so the program
knows how to pull info from each format. DG's steer: formats vary **mostly by
design #** (some real samples in hand, full variance unknown yet).

Built `templates.py` — a registry of `QuoteRunTemplate`s. Each declares what it
recognizes (design #, file extension, file-name markers) and how to extract
fields from that one shape; `parse_quote_run(path, design=None)` scores every
template and uses the best match. Return shape is a **superset** of the old
`drive_run.parse_drive_run_pdf` (`template`, `design`, `fields`, `raw_lines`,
`summary`), so it's a drop-in.

Seeded templates (the collection):
- **`d64_wheel_construction`** — Design 64, the `.xlsx` "D64 Wheel
  Construction" sheet (openpyxl, read-only; best-effort cell sweep).
- **`qt_run_text`** — HDX-style `.txt`/`.rtf` named "Qt Run"/"Quote Run"
  (requires the name marker so a plain `.txt` falls to the generic reader).
- **`pdf`** — any `.pdf` run; delegates to the existing `drive_run.py`.
- **`generic_text`** / **`unknown`** — fallbacks so something always matches
  and an unreadable format is named (not crashed) for adding later.

Scoring: handled extension = 1, design-# match +2, name-marker match +3, so
the most specific template wins (first wins a tie). `requires_signal` keeps
shared-extension templates (`.txt`) from grabbing files that lack their marker.

Wiring: `sales_orders.py` and `backfill_orders.py` now call `parse_quote_run`
for **every** run extension (was `.pdf`-only) and pass the job's design #;
jobs gain `drive_run_template`. `drive_run.parse_drive_run_pdf` stays as the
PDF reader the `pdf` template calls.

Tests: `test_templates.py` (14, in CI) — design parsing, matching by
design/ext/name, text + rtf + xlsx extraction, and the safe fallbacks. PDF
extraction isn't re-tested here (it just delegates to drive_run), so pdfplumber
isn't needed to run the suite.

**`check_orders.py` (same day, per DG: "I want to give you order numbers and
you check those out").** Hand it a list of order numbers; it opens each (board
or search-box), downloads the quote run, and runs it through the templates,
printing the matched template + per-template scores + fields + summary + first
raw lines. Headless/unattended (`--show` to watch); reuses the
`discover_documents` plumbing (refactored `_download` to return the saved
path). This is the loop for pinning formats down: run it on real orders, paste
a block back.

**Chicago Blower Qt Run parser (2026-06-15, DG ran `check_orders.py 421579`).**
First real sample in hand: job 421579's run is a Chicago Blower selection-program
text dump (`Z:\...\421579\ENG REF\QT RUN.txt`), and the generic `Label: value`
sweep mangled it (turning the outlet dimension table — `A`, `DA`, `DK`, … — into
junk fields). Added a dedicated **`cbc_qt_run_text`** template that matches by
content marker (`CHICAGO BLOWER` / `SN#`, so it beats `qt_run_text` even when the
file is just named `QT RUN.txt`) and pulls 29 real fields by targeted pattern:
serial, size/design/arr/%width/disch/rot, duty (CFM/SP/BHP/RPM/temp/density),
max HP/RPM/temp + ambient, tip speed, effective wheel dia, wheel construction
materials (blade/sideplate/backplate), shaft dia / brg centers / critical speed,
drive type, and the engineering-approval / non-std-materials / shrink-fit flags.
The dimension tables are deliberately left out. Verified against the real text in
`test_templates.py` (`REAL_CBC_QT_RUN`). NOTE: the run's `DESIGN 6195` is the CB
engineering design code, not the queue's design column.

**`check_orders.py` hardening (same run — it crawled / hung on later orders):**
the Z: AutoCAD-tree sweep now runs ONCE for all requested jobs (was per-order),
and the board page is reloaded between orders so a left-over detail modal or a
search that navigated away can't wedge the next lookup. Old/off-board orders
(419624, 420990) go through the search-box path; if that can't surface them the
order is skipped with a message instead of stalling the batch.

**Open (needs real samples on the work machine):** run
`python check_orders.py <order#> ...` (or `python templates.py "<a run file>"`)
on real orders, paste the output back, and pin the exact field headings into
the matching template's `extract`. Add a template per new design # as its
format turns up. Best-effort `Label: value` capture runs until then.

## 2026-06-11 — Sales-order LINE ITEMS: capture, normalize, search

New capability per DG's request: record **every line item** on each Sales
Order, store them per job, and make orders findable by what's on them. The
line items are free text and rarely written identically, so the end goal is
normalization for easy lookup.

What was built (all tested; 18 new tests in `test_line_items.py`, CI updated):

- **`line_items.py`** — pure logic + the store. Two capture signals (a line
  ending in a price/`N/C` column is an item anywhere; every line inside an
  "Additional Features"/"Accessories"-style section is an item even unpriced),
  conservative skip rules for totals/freight/footers/CO-history. Each item is
  stored as verbatim `raw` + normalized `norm` (qty/price stripped,
  abbreviations expanded: `W/`→WITH, `SS`→STAINLESS STEEL, `316SS`→`316
  STAINLESS STEEL`, …) + canonical `tags` (seeded fan vocabulary: SHAFT SEAL,
  SPARK RESISTANT, COATING, VIBRATION ISOLATION, …). Store:
  `BACKLOG_DIR/line_items.json` (`LINE_ITEMS_STORE` to move), atomic writes.
- **Wired into every parse path**: the daily run (`sales_orders.py` — also
  sets `j["line_items"]`/`j["line_item_tags"]`, so snapshots/history carry
  them), `backfill_orders.py`, and a new no-browser bootstrap
  **`line_items_scan.py`** that walks the already-archived PDFs under
  `SALES_ORDER_DIR` (latest CO# revision per job, resumable).
- **Search**: **`find_orders.py`** — AND/`--any` terms over raw+norm+tags,
  `--tag`, `--fuzzy` (typo-tolerant), `--job`, `--list-tags`, `--xlsx`
  inventory workbook (one row per item, AutoFilter). Term+tag AND at the JOB
  level (they may sit on different items).
- **Report**: Full Queue + History tabs gain a **Features** column (the job's
  tags, or `(N items)` when captured-but-untagged); the AI briefing receives
  `features` for new/returning orders.
- **Normalization levers** (raw is never lost, so all are lossless re-passes):
  `--dump <job>` shows per-line capture/skip decisions for tuning;
  `LINE_ITEM_RULES` (.env) points at a JSON that EXTENDS the built-in rules;
  `--renorm` re-applies current rules to the whole store; `--ai` classifies
  still-untagged unique items via the Claude API once (cached forever in
  `ai_tags`, pennies on haiku).

**Tuned against real documents (2026-06-11, dumps of jobs 421314 + 421473)** —
DG ran `--dump` on both and pasted the output back; the rules are now fitted
to the actual CBC SO anatomy and the two dumps live verbatim in
`test_line_items.py` (`REAL_LINES_*`) as the regression base:

- Item rows are `<description> <L|C|N> <Price Freight Markup Net Comm.>` —
  the type letter, then up to five money columns, **or `STD` / `INC` in the
  price column, or nothing at all** (`Weights on drawing ... L`). All
  captured now; the type letter is stripped from the norm (stored as
  `ptype`), and the LEFTMOST money column (Price) is kept, not Comm.
- **Unpriced continuation lines under an item are its `details`** (vendor,
  motor HP/enclosure/frame, `VFD Suitable`, `Product: Damper`): captured per
  item, searchable, and they contribute tags — the Ruskin `IVD C ...` row
  tags DAMPER + INLET VANES via the abbreviation (IVD → INLET VANE DAMPER)
  AND its `Product: Damper` detail. Page furniture interleaving a detail
  block at page breaks (`Chicago Blower ... (cont.)`, `v1.8.1.5 -1-`,
  ref-number rows) is excluded.
- New skips from the dumps: `List Total`, `Lead Time`, `Type Price Freight`
  header, commission/deduction lines, `Customs Invoice`, the
  drawings-distribution checklist (`Fan Drawings`, `O & M X`, ...), version
  footers, `^chicago blower`, `^order #`. The all-numeric CFM/RPM
  performance row and the spec-table row are rejected structurally. The
  email skip is now `\w@\w` so `Door Location: @9:00` stays a detail.
- New tags/abbrevs from the dumps: HEAVY DUTY, 3D STEP DRAWINGS,
  `grease fitting` → EXTENDED LUBE, `drive set` → V-BELT DRIVE, IVD.

**Feature Matrix (2026-06-11, per DG)** — `find_orders.py --xlsx` now writes a
second tab styled like the AutoCAD DWG matrix: one row per order, one column
per canonical tag (most-common first, rotated headers), green ✓ = the order
has that feature, red = it doesn't, Job # linked to the SO pdf, AutoFilter on.
With search terms the matrix covers only the matching orders but each row
shows the job's FULL profile from the store.

Remaining steps on the work machine:

1. `python line_items_scan.py` once over the archive (builds the store from
   every already-downloaded SO), then `python find_orders.py --list-tags` to
   see the vocabulary land.
2. `python line_items_scan.py --ai` to classify the long tail (MOUNTING
   CHARGE, WHEEL STEEL, CASH IN ADVANCE, ... — whatever the rules left
   untagged), cached forever.
3. If another order layout turns up (multi-fan orders?), `--dump` it and
   paste back.

Deferred:
- A wheel-type spec row that happens to end in a bare `L`/`C`/`N` could
  sneak past the structural guard (needs a ≥3-letter word, so most can't);
  revisit if a junk spec-row item ever shows up in the store.
- `qty` stays a guess (leading enumeration number) — the real dumps show
  item rows don't carry one (Qty lives in the spec table).

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

## Quote-run discovery results (2026-06-10) — how runs actually appear

Real listings (jobs 421473, 421492) settled discovery step 1. There is no
`CBC_DriveRun`/`CBC_QuoteRun` pid on normal fans:

- **Only HDX fans** give the run its own pid type. Everything else files it
  under **`CBC_Inquiry`** and it's recognizable only by FILE NAME, e.g.
  `421473_909-26-1604 Qt Run.txt` (note: a `.txt`).
- **Design 64** fans carry it as an Excel sheet instead:
  `421492_314-26-1647 D64 Wheel Construction (Inner...).xlsx` — needs its own
  parser later.
- **Some orders have no run in their documents at all** — it only lives in the
  job's AutoCAD folder, often in a subfolder:
  `Z:\AUTOCAD\CURRENT\JOBS\GENERAL LINE\420\420410\ENG REF\420410 qt  run.txt`.
- Other namings will exist; handled as they turn up via the env settings below.

Matching is now: pid type (`DRIVE_RUN_TYPES` + any `*Run` type) OR file name
(`DRIVE_RUN_NAME_PATTERNS` regexes: `qt\s*run`, `quote\s*run`,
`d64\s+wheel\s+construction`), with a recursive AutoCAD-folder fallback when
the documents carry none (daily run + backfill when the DWG scan knows the
folder). Downloads keep the original extension (the PDF-magic check now only
applies to `.pdf`); every matching file is archived; the report column is
labeled "Quote Run" and links the primary file (which may sit on Z: in the
folder-fallback case). `parse_drive_run_pdf` only runs on `.pdf` runs.

New deferred items:
- Parse the `Qt Run.txt` text format (fields TBD — dump one with any editor).
- Parse the D64 wheel-construction `.xlsx` (openpyxl; "pull the data
  differently" per DG).
- The `CS_SalesOrder` pid (seen at rev 2 on both jobs, an
  OrderVerificationReportViewer doc) is NOT counted as the Sales Order — the
  CO# still keys off `CBC_SalesOrder` revs only. Revisit if CO#s look low.
