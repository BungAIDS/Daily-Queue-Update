# Work Log — Construction Run + Backfill + AutoCAD DWG Scan

Running notes so progress survives across sessions. Newest status at the top of
each section. **If you're picking this up fresh, read this whole file first.**

## 2026-07-16 — Quote-Run review loop (the SO-review twin, per DG)

Update (same day): merged main's "Use-Note-as-direct-review-input" and ported
it to the twin — the yellow **Note** column IS the input now (no separate Add
Note column), matching the SO review. `sync`/`refresh` migrate an existing
workbook to the new layout, and legacy Add Note entries are still read once
during migration (`test_sync_upgrades_previous_add_note_layout`).

Fix (same day): DG's first Open QR Review hit Excel's "We found a problem /
Removed Records: Formula" repair dialog. Root cause: 3 real corpus lines start
with `=` (price sums like `=80240-6773-6056=67,411`) and openpyxl stores any
`=`-leading string as an Excel FORMULA — broken ones, which Excel deletes on
open. `_append_text_row` now re-types such cells as literal text in BOTH
review writers (quote_run_review + so_review, which had the same latent bug).
Verified against the published corpus: 27,220 rows, 0 formula cells, all 3
lines kept as text. After Git Update, run **Update QR Review** to rebuild the
sheet cleanly (the repaired copy lost those 3 cells; the store still has them).

DG: "how is quote run detection doing? could it use a boost? set up a similar
environment to sales-order review where I can review each order and give
feedback on each item." Detection state measured on the 2026-07-16 corpus:
468 runs / 382 orders, 389 OK / 62 CHECK VISION / 14 UNRECOGNIZED (.doc
dampers) / 3 NO FIELDS, 57.8 avg fields/run, 2 coverage tags (both legit
specials), ~3.5k uncaptured (MISSED) lines. The 62 CHECK VISION all have
attempts=None — the escalated re-read loop (built 07-06) hasn't been run yet;
`python pdf_vision.py` is still the pending P0 user action.

New **`quote_run_review.py`** — mirrors `so_review.py` exactly (same note
queue / handled-ledger / workbook contract), but unrolls the quote-run store
instead of the line-items store. Every order's runs (deduped via
`run_rank.dedupe_runs`, same as quote_runs.xlsx) become real filterable rows:
- **RUN** — the file itself (status — template — field count, hyperlinked to
  Z:), blue+bold; **FIELD** — one row per extracted field, name + value, in
  CORE_FIELDS order; **SUSPECT** — vision-QC complaints / NEEDS HUMAN reason,
  red; **MISSED** — each uncaptured `missed_data` line, amber. Real corpus:
  27,220 rows (421 RUN / 23,216 FIELD / 3,504 MISSED / 79 SUSPECT), ~6s build.
- Notes anchor by stable row key (FIELD keys on the field NAME, so a note
  survives the value being fixed; MISSED/SUSPECT key on their text and
  self-clear when a pattern lands), fall back to exact text, then the order's
  first row. Same open/handled lifecycle, Resolved history tab,
  `quote_run_review_handled.json` tracked return-ledger, notes published via
  data_push (`quote_run_review_notes.json` added to the file list).
- CLI: build / open / sync / refresh / reparse / list / handle. `reparse` =
  `quote_run_scan --reparse-stored` (patterns + vision QC, no Z:, no API) then
  rebuild, so one action shows the current parser's take for re-review.
- Launcher: new **Quote Run Review** category — Open / Update / Re-parse +
  Refresh, mirroring the SO trio. Tests: `test_quote_run_review.py` (12, in
  CI after test_so_review); full-loop smoke on the real corpus (build → type →
  sync → handle → refresh) verified.

## 2026-07-10 — Fix: workbook unusable during renders (phantom row rewrites)

DG: "the update takes so long the sheet is almost unusable for the whole 2-min
cycle." watch.log quantified it: 66 on board, new/returning/removed all 0,
field changes logged only 14x ALL DAY — yet Live Queue rewrote 5-15 rows every
poll and renders ran 30-40s+. Cause: the '#' board-position column. cbcinsider
reorders tied rows between scrapes, the jitter changed those rows' signatures,
and each one got a full slow rewrite (bulk values + style clear + per-cell
hyperlinks/comments — the expensive COM path).

- `row_sig` now masks VOLATILE cells' values (style still hashed) and the Live
  Queue's '#' cell is volatile — jitter can't force a row rewrite. NOTE: sig
  format changed → one full repaint on the first cycle after updating (the
  watcher resets lq_sigs on start anyway).
- `apply_upserts` gained `positions` ({job: pos} from the board scrape): the
  whole '#' column is refreshed in ONE bulk Range write, and only when the
  vector actually moved (`_POS_LAST`); sort/AutoFilter re-extend ride the same
  condition. Structural-CF gating unchanged.
- Considered and rejected: build the workbook offline and swap the file in —
  openpyxl-style replacement kicks every co-author out (the documented reason
  this module drives desktop Excel via COM). Staging-sheet swaps would reset
  filters/scroll and break internal links every cycle. Root-causing the render
  cost was the right lever: a normal cycle now touches ~0-4 rows.

## 2026-07-10 — Fix: Similar Orders tab went blank (repaint race + churn)

Field report: "enter 421507 → nothing pops up" while Changes' DWG Reuse showed
419623. watch.log (debug-logs) showed the mechanism at 10:13:26: both Similar
sheets' repaints died with OLE 0x800ac472 (Excel rejects COM writes while the
USER is mid-edit — they were typing in the picker). The repaint had already run
`Cells.Clear()`, wiping the FILTER formula — and `_RENDER_CACHE` still held the
last SUCCESSFUL fingerprint, so as long as the model didn't change, every later
cycle skipped the rewrite → blank tab until the model changed or a restart.

- Fix 1 (the bug): pop the sheet's `_RENDER_CACHE` entry BEFORE `render_sheet`
  and re-set it only after success — a failed paint now always retries next
  poll. Applied to both the repaint loop and the legacy `update_workbook`.
- Fix 2 (the amplifier): Similar Data was repainting nearly EVERY cycle
  because rows + the column-I dropdown followed live board-position order,
  which reshuffles per poll. Rows/queue list now sort by job number
  (`_sim_sort_key`), so the sheet repaints only on real content changes and
  the Live Queue 'Similar' anchors stop shifting every cycle. Every needless
  repaint was another window for the user-editing race.

## 2026-07-10 — Tab order enforced; Line Items tab retired

DG's layout: Changes | Live Queue | Order History | Similar Orders | Similar
Data, with the old Line Items tab gone (superseded by the Similar tabs +
`find_orders --xlsx`).

- `SHEET_ORDER` (live_excel) is now that order and actually ENFORCED:
  `_ensure_tab_order` snaps the managed tabs to the front of the tab bar each
  cycle (no-op when already right; restores the active sheet since Sheet.Move
  can steal focus). User-added tabs ride behind, untouched. A coworker
  dragging tabs around gets snapped back — by design.
- `_drop_obsolete_sheets` deletes the "Line Items" tab on sight
  (OBSOLETE_SHEETS): `line_items_sheet` had NO callers — the tab was already
  orphaned/stale from an older build — so the builder + LINE_ITEM_HEADERS +
  its two tests were removed with it. Data is regenerable from the stores.

## 2026-07-10 — Live Queue 'Similar' column: click -> that order's lookalikes

Asked: "click something in a Live Queue column that jumps to the new tab and
searches automatically — or does that need a macro?" Setting the picker cell
from a click DOES need VBA (rejected: .xlsm kills the no-macro/co-author
design). Macro-free equivalent shipped instead:

- Similar Data is now a VISIBLE grouped tab (grey band + bold order # on each
  group's first row, folder cells hyperlinked; Queue Order value still repeats
  every row because the picker tab's FILTER matches on it — do not blank it).
- New trailing Live Queue column **"Similar"** (between Last Out and #):
  lookalike count, internally hyperlinked (`#'Similar Data'!A<row>`) to that
  order's group — `live_sheets.similar_anchor`. watch stamps
  `_sim_count`/`_sim_anchor` on each on-board job BEFORE rows are planned.
  Anchors self-heal: row_sig includes the link, so when a group's row number
  shifts, the affected Live Queue rows re-plan on the same cycle.
- `live_excel._style_row`: links starting with `#` become internal hyperlinks
  (Address="", SubAddress=...). render_sheet now also RE-shows a sheet whose
  model isn't hidden (ws.Visible set both ways) so earlier hidden-build sheets
  resurface.
- LIVE_QUEUE_LAST_OUT_COL is now len-2 (new LIVE_QUEUE_SIMILAR_COL = len-1);
  removed_block untouched (its own header list, empty trailing).

## 2026-07-10 — Similar Orders tab: pick an order, see its lookalikes

Interactive tab in the live workbook (user asked for "select an order at the
top → it generates a list"):

- **Similar Orders** (visible): B1 = yellow picker cell with a data-validation
  dropdown of the on-board orders (typing any order # also allowed, ShowError
  off) + ONE `=IFERROR(FILTER('Similar Data'!...))` spill formula. Instant,
  no macros, works for every co-author. `&""` on both sides of the compare
  coerces numeric/text job cells so typed vs dropdown vs stored types all match.
- **Similar Data** (hidden, `Sheet.hidden`): flat (queue order × top-8 similar)
  table + the dropdown's source list in column I (every board order, even ones
  with no matches). Watcher computes rows in `watch._similar_orders_rows` —
  cached on (board ids, line-items store mtime, DWG scan mtime), ~1.3s to
  recompute for a 20-order board, skipped entirely when nothing changed.
- Renderer (`live_excel`): `Sheet.hidden` (ws.Visible=0) + `Sheet.picker`
  (`_apply_picker`: search-box styling, validation list, comment) + an explicit
  `.Formula` assignment pass for "="-prefixed cell values (locale-immune EN-US
  separators). The picker cell's typed value is read before `Cells.Clear` and
  restored after, and the visible tab's model is layout-only so its fingerprint
  almost never changes → repaints (which would blank the pick) are rare.
  `update_master_workbook` grew `extra_sheets` — Changes + extras now share the
  same fingerprint-cached repaint loop.
- Spill gotcha handled: rows below the formula aren't in the model, so nothing
  blocks the spill (COM writes of "" produce truly blank cells).
- Tests: layout + formula-range cases in test_live_sheets (34 now).

## 2026-07-10 — Similar-order suggester wired into the live queue ("DWG Reuse")

The `--like` ranking now runs automatically for every enriched order (new
arrival on the watch / change-order re-fetch / daily run):

- `find_orders`: `similar_jobs` split into `build_index` (one pass over the
  store: tag/line sets + rarity counts) + `similar_to_items` (score any items
  against it — works for orders NOT in the store, i.e. brand-new ones).
  `reuse_suggestions` = thresholded custom-DWG-only shortlist trimmed for
  storage on the job dict; `reuse_label`/`reuse_note` render it. Measured on
  the real corpus (6K orders / 75K lines): 0.09s index once per batch +
  ~16ms/order.
- `sales_orders.enrich_with_sales_orders`: after the line-items store is
  updated, builds one index and stamps `dwg_reuse` (list) +
  `dwg_reuse_label`/`dwg_reuse_note` (strings) on each job dict — flows into
  live_master via the normal upsert. Best-effort try/except; NOT in
  live_master._TRACKED so it can't spam the change log.
- New **"DWG Reuse"** column in `excel_writer.COLUMNS` — placed AFTER CO#
  because the Changes tab aligns Folder/Quote Run/CO# in fixed columns across
  its tables (test_changes_today_columns_align_across_sections). Cell = top
  candidate + suffixes (`421100 (-07,-51) +2`), hyperlinked to its CAD folder,
  full shortlist w/ shared SO lines as the hover comment (excel_writer +
  live_sheets both).
- `notify._order_facts`: new-order toast/Teams card gets a "DWG Reuse" fact.
- Config: `REUSE_MIN_SCORE` (default 0.5 — on this corpus that separates
  "same fan" from common-feature noise; 99 disables) and `REUSE_TOP` (3).
- Known noise source: address/routing boilerplate ("PO BOX ...", "ROUTE TO
  ...") is stored as line items and occasionally boosts same-customer matches.
  Mostly harmless (same customer IS a reuse signal); the proper fix is a
  line_items skip rule, someday.

## 2026-07-09 — DWG-aware search + `--like` similarity ranking (find_orders)

Goal: surface the backfilled SO data WITHOUT widening the live queue workbook
(way too many columns), and take the first programmatic step toward "new order
comes in → which backlog jobs already have a custom DWG for this?".

- `find_orders.py` now joins the AutoCAD scan store into every view: each hit
  prints its custom-DWG suffixes + CAD folder (`attach_dwg`/`_dwg_label`), and
  `--dwg` keeps only jobs the scan found custom drawings for.
- New `--like JOB` mode (`similar_jobs`): ranks every other order by
  rarity-weighted SO overlap — each shared canonical tag scores
  1/(#jobs with that tag), each IDENTICAL normalized line scores 2/(#jobs with
  that line) — so rare shared features dominate and MOTOR-on-everything counts
  for almost nothing. `--like 421314 --dwg` = the DWG-reuse shortlist. `--top`
  caps the list (default 15).
- `--xlsx`: "Custom DWGs" column on both sheets; on the Feature Matrix it
  hyperlinks to the job's CAD folder (link only set when there IS a label —
  openpyxl otherwise displays the bare target in the empty cell).
- Launcher `find_orders` action gained "Similar to job", "Only custom-DWG
  jobs", "Similar jobs to show" options. Tests: `test_find_orders.py` (pure
  dict-in/dict-out, in CI after test_line_items).
- Deliberately NOT added to the live workbook. Next step for the auto-recommend
  goal: watch.py calls `similar_jobs` for each NEW order against the store and
  puts the top DWG-reuse candidates in the notification/one compact column —
  `similar_jobs` is already pure and store-driven so it can be called as-is.
  An AI pass (send the new order's lines + the top-N shortlist, not the whole
  DB) can sit on top later if the ranking alone isn't judgey enough.

## 2026-07-09 — Hub/coupling/box fields + coverage tagging ("read right over it")

DG's asks: (1) fabricated hubs have no cast part number but do carry hub data
(HUB TUBE 3/4, HUB FLANGES 1/2, HUB BORE/OD) that we were reading right over;
(2) "there are similar cases where we miss info because it doesn't match — TAG
these"; (3) how much is still missing. All grounded on the real corpus
(`git show origin/order-data:quote_run_scan_progress.json`), tested, reparsed.

New CB fields (`templates.py` `_CB_PATTERNS`, all real corpus shapes):
- **Fabricated hub** (no part number): `Hub Tube Gauge/Material`,
  `Hub Flanges Gauge/Material`, `Hub Centers Material`, `Hub Bore`, `Hub OD`,
  `Hub Bushing` (`Q2 BUSHING ..`). Fixes the 421572-style hub. `Hub Bore` also
  fires on cast hubs.
- **Half-coupling shaft block**: `Coupling Max/Min/Nom Shaft Dia`,
  `Coupling Keyway` (`= 0.3750 X 0.1875`). ~36% of runs.
- **Inlet/damper box**: `Box B`, `Box C` (`BOX B X C: 73 IN. X 15 1/2 IN.`),
  `Inlet Box Angle` (spec-line `,BOX 270`; `BOX 0` = no box, skipped).
- **Inlet Cone** (`REINFORCED INLET CONE INCL. PER SK-19-72`) + **Safety Guard**
  presence flag (`SHAFT SAFETY GUARD`).
- Two bug fixes surfaced by the tag: `Sheave PD` now also reads `SPECIFIED PD`
  (customer item-59 override), and the `Hub` part-number pattern accepts the
  plural `HUBS 19-5-21`.

**Coverage tagging** (`templates.coverage_tags` / `missed_data_lines`, wired
through `quote_run_scan.apply_coverage`, stored on each run as `coverage_tags`
+ `missed_data`, surfaced in the new workbook **Review** column, amber):
- `coverage_tags` = high-precision, self-clearing probes — a data family whose
  keyword is in the doc but produced NO field ("read right over it"). After the
  new patterns only **2 of 450** runs are tagged, both legitimate specials: a
  fabricated FEA-markup billet hub (415970), and a dual-motor shared-sheave
  drive (417445). Add a pattern for a family → its tag disappears on reparse, so
  the tag count is a live coverage metric.
- `missed_data` = the actual uncaptured data lines (noise-filtered, capped 15)
  behind the tags, so a human/next pass sees exactly what's still read over.

Avg fields/run **54 → 58**. Full suite green (test_line_items still skipped —
sandbox pdfminer/cryptography panic, pre-existing).

**How much is still missing (uncaptured, non-noise data lines, uncapped):**
~8.0k lines across 294 text runs. Ranked remaining clusters (the next batch):
- **Alternate outline-dimension format** (`D = 56 1/16`, `A = 22 13/16`, the
  DA/DB/DC/CA-CE/FF/FR/FY... `code = inch  mm` list, ~113 runs). Same geometry
  family as the AXIAL/SIDE VIEW table we already parse, but a second compact
  format with ~40 extra detail codes not in the table. Biggest gap.
- **Spun sideplate** construction row (`SIDEPL,SPUN 0.075 (14) ..`, 179 lines) —
  a wheel-table variant our sideplate row regex misses.
- **Accessory rows**: `BURN TAPES`, `CHANNEL`/stiffener, and the trailing
  WR2/weight/price columns on every wheel-construction row (the P2 item).
- `GOOD FOR $..` (price variant), `OUT ANGLE`, `OUTLINE DIM. FORM SK-` refs.

**Damper docs** (DG: "only damper-relevant info should count"): investigated all
19 damper-flagged runs. The readable ones are *full fan runs* (the fan the
damper attaches to) — and for 419624/418421/405167 the damper-titled file is the
order's ONLY run, so stripping fan fields would blank the order. The 10 `.doc`
ones are unreadable binary (yield nothing already). So the correct,
non-destructive behavior is what we have: extract the fan run + carry the
`damper=True` flag (filterable). The new Box B/C/Inlet-Box/Cone fields now also
capture the damper-relevant geometry. Damper-*specific* fields (blade count,
damper size) would be additive — needs a sample of what damper data DG wants.

## 2026-07-01 — Shaft/bearing geometry + outline dims (BX/STB/N/F etc.)

DG pointed at the full job-421579 run (the earlier samples were trimmed above
BRG CENTERS). Added the shaft/bearing + outline fields DG asked for, all
grounded and tested against the real text — including a run against the FULL
noisy document to prove no false matches from the part-cost tables, the
factory-use number block, or the decoy `OUTLET N X A:` / flange-punching `N =`
lines.

New CB fields (`templates.py` `_CB_PATTERNS`):
- Shaft/rotor geometry line `LENGTH .. ,OH .. ,BX .. , STB .. , TG&P ..` + `STH`
  → **Shaft Length, OH, BX, STB, TG&P, STH**.
- Bearing spec block → **Bearing Size, Bearing Series, Bearing L10 Hr**
  (from the DRIVE-FLOAT row).
- Outline dims (AXIAL/SIDE VIEW), anchored on code+description so a bare letter
  can't false-match → **Housing Width (N)**, **Base to CL (F)** (=F/2).

Surfaced BX/STB/OH/STH/Bearing Size/Series/N/F in `_CB_SUMMARY_ORDER` (the Quote
Run Details column); all new fields added to `quote_run_scan.CORE_FIELDS`.
Test: `test_chicago_blower_shaft_bearing_and_outline_fields` on the real 421579
tail fixture.

Outline table: now pulled in FULL (all 16 codes) — see the block-parser entry
below.

## 2026-07-06 — Remove the storage truncation (raw_lines / transcript caps)

DG: the 250-line cap shouldn't be a thing. It only ever bit the OFFLINE corpus/
--reparse-stored loop (the live parse always used full text), but 42 runs were
being clipped, so tail sections (outline dims, totals on long/dual docs) were
lost to offline re-parsing. Fixes:
- `templates.RAW_LINES_CAP` 250 -> 10000 (a pure runaway backstop; the longest
  real run, a dual 4S/8S 8-pager, is well under 1000 lines). Applied to the CB,
  PDF, and generic-text templates (the generic one was also silently capped at
  40 lines).
- `drive_run.py` stored only the first 40 lines as raw_lines though it already
  kept the full `text` — dropped that slice.
- Vision transcript char cap 20000 -> 80000 and the vision `max_tokens`
  6000 -> 12000, so a long SCANNED doc's transcript isn't clipped either (output
  tokens are billed only when generated, so short docs — most — cost nothing
  extra). Store grows ~1-2 MB; negligible.
Takes full effect on the next `--rescan` (which re-stores the 42 clipped runs at
full length). 20 suites green.

## 2026-07-06 — Dual-arr selection, dedupe, and the OBSOLETE archiver

Per DG's answers to the cleanup questions:

1. **4S/8S dual files -> keep arr-4** (16 files). A doc that quotes the same fan
   in an arrangement-4 (motor-mounted) AND an arrangement-8/9 (on bearings) is
   two printouts in one file; DG: the arr-4 is the built unit.
   `templates.select_primary_run_text` splits the doc into pages, keeps only the
   arr-4 pages, drops the rest — called at the top of `_parse_chicago_blower`.
   This also fixes a real bug: the 8S run's bearing/rotor section was leaking
   onto the (bearing-less) 4S fan via first-match. Validated on the real 413224
   (4S+8S x6 -> just 4S, no bearing leak); 13 raw + a few vision duals fixed.

2. **Dedupe format-dupes/old revs** (DG: keep the .txt). `run_rank.dedupe_runs`
   collapses runs that share a fan spec (Size/Design/Arr) to the most-current
   copy; genuinely different fans survive. `run_rows` uses it, so the xlsx shows
   one row per distinct run. Real corpus: 467 run files -> 420 rows (47 dupes/
   old-revs collapsed).

   Plus **`archive_obsolete.py`** — moves non-active quote runs AND sales orders
   (format dupes, old CO/REV, older SO revisions) into `<job>\OBSOLETE\` so the
   live folder holds only the active file. SAFE: DRY RUN by default (--apply to
   move), MOVES never deletes, name clashes get a numbered suffix, writes an undo
   manifest to BACKLOG_DIR, `--undo <manifest>` reverses it. Launcher: Tools ->
   Archive Obsolete Runs/SOs (Apply is a confirmed checkbox). NOTE: acts on Z:,
   untestable from the sandbox — tested on synthetic folders; DG must DRY RUN and
   eyeball before --apply.

3. **Dampers: skipped** for now (per DG).

Tests: select_primary + no-leak (test_templates), dedupe_runs (test_run_rank),
archiver plan/dry-run/apply/undo/clash (test_archive_obsolete). 20 suites green.

## 2026-07-06 — Vision re-reads escalate, compare, and stop (no blind re-pay)

DG asked the right question: "are we confident scanning again solves anything?
will we compare the 2 results?" It didn't before — a CHECK VISION run was
re-read with identical settings (same model would repeat the same OCR error)
and the new result overwrote the old with no comparison, and a stubborn PDF
would re-read (and re-pay) every future batch. Reworked:

- **Escalated re-read**: `read_scanned_pdf(hints=, hi_res=)`. A re-read renders
  at 2576px/scale 3 (was 1568/2) AND `build_prompt(hints)` tells the model the
  exact prior complaints ("implausible CFM='4/100'", "odd Arrangement '781'").
- **Compare + give up**: `apply_vision_result` tracks `attempts` and stashes the
  prior reading's `prior_fields`. After a re-read still fails QC and
  `attempts >= MAX_VISION_ATTEMPTS (2)`, `escalate_to_human` compares the two
  reads (`compare_readings` on hard-number fields) and sets terminal **NEEDS
  HUMAN** with a reason — "two reads disagree — CFM: 28000 vs 26843" or "two
  reads agree but values look wrong — ...".
- **Terminal**: NEEDS HUMAN is never a re-read candidate and QC won't re-open it
  (but a later pattern fix that makes it clean clears it to OK). Excluded from
  `--reparse-attention` (a Z: re-parse can't fix a scan). Orange in the xlsx.
- Tests: hints-in-prompt, attempt/prior tracking, compare_readings,
  escalate disagree/agree, NEEDS HUMAN terminal + cl- ear-on-fix. 19 suites green.

## 2026-07-06 — Vision QC: repair what we can, flag the rest + HANDOFF_PLAN.md

Per DG (low on assistant usage; handing off): `pdf_vision.apply_vision_qc`
validates every vision run offline — numeric plausibility per field class,
arrangement whitelist (catches the systematic OCR "S read as 8/5": 7S1->781,
8S->88), and model-vs-clean-transcript disagreement on CFM/SP/BHP/RPM. Where
the transcript's targeted parse is clean and the model value garbled, the
field is REPAIRED in place; otherwise the run is flagged **CHECK VISION**
(amber in the xlsx) and the next `python pdf_vision.py` re-reads exactly those
(candidates include CHECK VISION without --redo). Runs inside
`--reparse-stored`, so one command applies patterns + QC.

Dry-run on the real store: 59 PDFs flagged (see list below), several fields
auto-repaired (incl. 9 arrangements from model slop like "ARR 8S, 100.0 PCT").

Flagged: 400934 401078 401195 401217 401266 401445 402395 402547 403049 404135 404216 404346 404357 404459 404641 404783 405693 406123 406678 406841 406906 407015 407189 407349 407497 408001 408015 408289 408290 408355 409440 409784 409960 410089 410887 411028 411091 411180 411306 411484 411820 411821 413044 413242 413265 413592 413680 413967 414648 414651 414686 415158 415956 416730 416809 417395 418342 419644 420848

**HANDOFF_PLAN.md** added at repo root: the full big-picture roadmap (system
map, working loop, P0-P7 priorities with why/how/verify, evidence bar,
confidence register, command cheat sheet) for the next session to continue.

## 2026-07-06 — Rank multiple quote runs by currency (71 orders have >1)

Per DG. 71/381 orders carry multiple run files; master.json and the daily
report used to take `runs[0]` — ALPHABETICAL — so 400567's 2021 base run beat
its `Qt Run CO#1.txt`. New import-light `run_rank.py`: most-current first by
(1) highest CO# in the name, (2) highest REV letter/number, (3) newest file
mtime (now captured by the sweep), (4) most fields; stable on full ties.

Applied everywhere one run represents the order:
- `master_sync.merge_quote_runs`: ranked head leads `drive_run`, and ALL runs
  (revisions included) are kept under a new **`drive_runs`** list — history is
  queryable instead of discarded.
- daily enrich (`sales_orders`): the archived-run fallback picks the ranked
  head for `drive_run_pdf` / Quote Run Details.
- `quote_runs.xlsx`: still one row per run, now ordered most-current first.

Dry-run on the real store: 31 multi-run orders change primary (400567→CO#1,
416869→CO#2, 408682→Rev1, 417821→REV 2, ...). Tests: `test_run_rank.py` +
a ranked-history case in `test_master_sync.py`.

## 2026-07-06 — Section templates for every arrangement (designed from the corpus)

Per DG: different arrangements carry different parts of the run (bearing
section, wheel info, base/motor variables) and we tracked almost none of the
arr-4 family's. Analyzed the FULL corpus from the order-data branch (447 docs:
324 stored raw texts + 123 vision transcripts) — built a per-arrangement
frequency matrix of every candidate line, then wrote SECTION-based patterns
(they fire when their section exists; no arrangement gating):

- **Arr-4 family block** (94+ runs, was 0% tracked): Wheel Weight/Thrust/WR2
  line, Housing to Wheel CG / Hub Inlet Face, motor base — now 92-99%.
- **Rotor/bearing block** (arr 1/3/7/8/9): Rotor WR2/Max RPM/Material, Stress
  Ratio at Hub/Bearing, bearing-loads rows (DRIVE-FIXED **and** -FLOAT,
  negative statics, wide layouts — the old L10 pattern only matched
  DRIVE-FLOAT; now anchored on the number before the P/C decimal) — 93-100%.
- **Universal**: Blades / Max RPM Wheel Only / RES CPM, Housing Construction
  (incl. arr-7 "SPLIT HOUSING AND BOX"), Stiffeners, Fan Outlet Area, Motor
  Frame/Position/Enclosure/Weight, Sheave PD + Min PD (belt), Inlet Box Size,
  Shaft Seal / Flanged Inlet flags, Total Weight + Total Price (GOOD FOR line),
  Drive="Motor mounted" for FR-MOTOR runs.
- Serial now defaults to the job number when the header omits SN# (93% vs
  27-60%). ~30 new CORE_FIELDS columns; summary adds Blades + motor block.

**`--reparse-stored`** (also in the launcher via Scan Quote Runs options):
re-runs the parser over the stored raw_lines/transcripts — new patterns apply
in seconds, no Z:, no API. Vision fields kept, pattern hits merged over them.
Dry-run on the real store: 416/450 runs updated, avg 54 fields/run.
Tests: real 401221 arr-4S fixture + bearing-row variants + reparse merge rules.

## 2026-07-02 — Read scanned (no-text-layer) PDFs with Claude vision

Per DG: the ~128 runs flagged "PDF (no text layer)" are unreadable by
pdfplumber (image-only scans/drawings). New `pdf_vision.py` sends exactly those
to Claude vision — classify + extract in one call. Cost reality-check for DG:
~1-2k input tokens/page on Haiku ≈ well under a cent per document; the whole
backlog is on the order of a dollar, one time (not $1/doc).

- Renders page 1-2 via pypdfium2 (already installed with pdfplumber; no new
  deps), downscaled to 1568px; asks for JSON {doc_type, fields, note} using the
  SAME field names as the text parser, so scanned runs land in the same
  workbook columns and master.json shape.
- Outcomes: quote_run -> fields, status OK, template `pdf_vision`; drawing ->
  new status **DRAWING** (grey in the xlsx, excluded from "needs attention" and
  `--reparse-attention` forever); error/refusal -> left flagged, retried next
  run for free. Progress saved after EVERY answer (they cost money).
- **Never re-pays**: runs with a vision result are skipped (`--redo` to force),
  and `quote_run_scan.carry_vision_forward` preserves vision results across a
  full `--rescan` (which starts from an empty store — it now loads the prior
  store just for this).
- **Full transcript stored** (per DG): every read also returns a complete
  transcription of the document, kept at `run["vision"]["transcript"]` in the
  progress store (capped 20k chars; survives rescans with the vision result).
  So when new fields are wanted later, we re-parse the stored text for free —
  no `--redo`, no second API charge. Roughly doubles output tokens per doc
  (`max_tokens` 6000); backlog total still ~a dollar or two.
- Config: `PDF_VISION_MODEL` (default = CLAUDE_MODEL, i.e. Haiku),
  `PDF_VISION_MAX_PAGES` (default 2). Needs the existing ANTHROPIC_API_KEY.
- Launcher: **Scans / Backfill -> Read Scanned PDFs (AI)** with Jobs/Limit/
  Model/Redo. TRIAL FIRST: run with Limit=5 (~3 cents), eyeball the fields in
  quote_runs.xlsx, then run with Limit blank for the rest.
- Tests: `test_pdf_vision.py` (parsing incl. fenced/garbage replies, run
  updating for all three outcomes, candidate selection, rescan carry-forward);
  plus an end-to-end dry run (real pypdfium2 render + mocked API) verified the
  CLI flow and the no-re-pay path.

## 2026-07-01 — Auto-publish order data on change (opt-in)

Per DG: keep the published snapshot current automatically so a remote reader
tracks the data as we gather more about orders. New `DATA_PUSH_ON_CHANGE` flag
(config, default off). When on:
- `master_sync.run()` republishes after it saves the master (so every scan/
  backfill auto-publishes — one chokepoint, all four scans funnel through it).
  Only fires when a source actually changed (`any(counts.values())`).
- `watch.py` publishes when a poll brings new orders or field changes, and once
  at session end (mirrors the existing `_publish_logs`). Idle polls don't push.
Both are best-effort (a failed/absent push never disturbs a scan or the watch)
and gated on `DATA_PUSH_ON_CHANGE and DATA_PUSH_BRANCH`, so tests (flag off)
never hit the network. Verified: master_sync tests still green, and gating
off/on/no-branch checked directly. The manual launcher task is unaffected.

## 2026-07-01 — Publish order data to a branch for remote access (data_push.py)

Per DG: make the order data readable remotely the way `log_push.py` already does
for the watch log, so it can be inspected without copying files off the Windows
box. New `data_push.py` mirrors that plumbing exactly — hash-object -> mktree ->
commit-tree -> force-push a single ORPHAN commit onto `DATA_PUSH_BRANCH` (default
`order-data`). Never touches the working tree/index; each push replaces the
branch (no history bloat).

- Publishes whichever exist: `live_master.json`, the quote-run / line-item /
  backfill / autocad JSON stores, and the `quote_runs`/`backlog`/`line_items`/
  `autocad_dwgs` xlsx sheets. `build_snapshot_commit()` is factored out so the
  tree/commit build is exercisable without a network push.
- Config `DATA_PUSH_BRANCH` (config.py, next to `LOG_PUSH_BRANCH`); empty
  disables. **Private repo only** — carries customers/prices.
- Launcher: **Tools → Publish Order Data** (`data_push.py`, optional branch arg).
- Read side: `git fetch origin order-data && git show order-data:live_master.json`
  (or check the files out) from any clone — binary xlsx round-trips intact.
- Validated end-to-end against the real remote on a throwaway `data-selftest`
  branch (orphan commit confirmed, binary bytes intact). NOTE: a sandbox token
  couldn't delete that test branch (GitHub 403 on delete); it's harmless dummy
  content and can be removed from the GitHub UI.

## 2026-07-01 — Pull the whole AXIAL/SIDE VIEW outline table (block parser)

Per DG: grab ALL the outline dimension codes (not just N/F). They don't all
belong in the Live Queue but should live in the master .json — which they do,
since `master_sync.merge_quote_runs` stores the full `drive_run` fields dict
unfiltered (only N/F stay in `CORE_FIELDS`/the summary; the rest ride the
inventory "Other" column + master).

- Replaced the two one-off N/F regexes with `_parse_outline_dims` (templates.py):
  finds the `AXIAL VIEW` marker, scans to the `PUNCHING DETAIL` / `PART NAME` /
  page-break boundary, and matches each `<code>  <DESC>  <inches>  <mm>` row with
  `_OUTLINE_ROW`. Captures all 16 codes (A/W/KK/E/RB/RM/F/H/TV/RH/LH/MA/D/K/N/LR)
  as "<Description> (<code>)"; a curated `_OUTLINE_DIMS` map gives the common
  codes clean names, unknown codes fall back to their own description text (so a
  different arrangement's extra codes are still captured).
- Verified on the full noisy doc: the flange-punching `A =`/`N =`/`DA =` table
  and the part-cost tables produce ZERO false matches (bounded + shape-anchored).
- Test: `test_chicago_blower_shaft_bearing_and_outline_fields` now asserts all 16
  dims and that exactly 16 are captured.
- DG says the earlier trimmed samples (421237/421572) are other good run
  examples; offered to supply more. Only 421579 currently has a full outline
  section, so a 2nd arrangement's full run would let us confirm the map/fallback
  across fan types (current parser already handles unknown codes generically).

## 2026-07-01 — Extract wheel-construction gauges + surface construction detail

Per DG, the report needs the wheel construction detail, not the aero fields
(CFM/SP/BHP/RPM already come from the Sales Order). The CB "Qt Run" wheel table
`WHEEL  THICK.(GA)  MATERIAL  WR2  WEIGHT` already had its material pulled but
the **gauge/thickness column was discarded**.

- New CB fields **Blade/Sideplate/Backplate/Liner Gauge** — capture the
  `THICK.(GA)` value ("1/4", "3/8", "0.048 (18)") using the same row anchors as
  the material patterns (`templates.py` `_CB_PATTERNS`). Verified on all three
  real samples (421579, 421237, 421572).
- `_CB_SUMMARY_ORDER` reordered to lead with construction: each material paired
  with its gauge, then Hub / Coupling, then Shaft Dia / Brg Centers / Critical
  Speed. This is what now shows in the **Quote Run Details** report column (Hub
  part number + bearing section were already parsed, just not surfaced).
- `quote_run_scan.CORE_FIELDS` gains the four gauge columns next to their
  materials so the inventory promotes them out of "Other".
- Tests: gauge assertions added to the three CB field tests (incl. the GUSSETS
  case that has no sideplate/backplate gauge).
- STILL NEEDED FROM DG: a real Qt Run showing the **BX / STB** bearing
  designations and the **"N F" section** — no current sample contains them, so
  those patterns can't be written without guessing.

## 2026-07-01 — Surface parsed quote-run fields in the report (Quote Run Details)

The daily run already parses every quote run into real engineering fields
(`sales_orders.enrich_with_sales_orders` sets `j["drive_run"]` +
`drive_run_summary` + `drive_run_template` via `templates.parse_quote_run`), but
every user-facing output showed only a `YES` / `YES (X)` presence flag — the
parsed fields were computed each morning and dropped (they reached only the
offline `quote_runs.xlsx` / `backlog.xlsx`). This was the intended-but-unbuilt
extension point called out in `drive_run.py:19-22`.

- New **`Quote Run Details`** column in `excel_writer.COLUMNS`, keyed on the
  already-computed `drive_run_summary` (Size/CFM/SP/BHP/RPM/materials/…). Because
  `QUEUE_HEADERS` derives from `COLUMNS`, it surfaces in the **Full Queue report**
  and the **Live Queue / Order History** (live_sheets) in one change; both writers
  render it through their generic `j.get(key, "")` fallback. Abbreviated to
  `Run Details` on the compact Live Queue via `_HEADER_ABBR`.
- Appended at the end of `COLUMNS` so the established front-block layout is
  undisturbed; `_COL_IDX`/`TOTAL_PRICE_COL`/`LIVE_QUEUE_*` are all derived, so
  nothing shifts. No new parsing.
- Tests: `test_quote_run_details_column_surfaces_summary` (report row renders the
  summary, YES flag untouched, no-run row stays blank). Full suite green except a
  pre-existing sandbox-only pdfminer/cryptography import panic in
  `test_line_items.py` (environment, not this change).
- NEXT: promote specific high-value fields (CFM/SP/BHP/RPM, materials, the
  Engineering Approval / FEA / Shrink-Fit flags) into their own sortable columns
  once DG picks which matter; the summary column is the interim surface.

## 2026-06-25 — Ctrl+C finishes the current poll (hardened + clearer)

The first Ctrl+C already only sets a stop flag (no raise), so the poll in progress
runs to completion before the watch saves and exits; a second Ctrl+C force-quits.
On Windows the OS console-control handler makes this robust (own thread, suppresses
the default KeyboardInterrupt) even while a poll is blocked in Playwright/Excel.
Hardening this run:
- Startup logs which Ctrl+C guarantee is active so it's visible.
- The 'console handler unavailable' fallback is now a WARNING (it's the only case
  a poll could be cut short — install pywin32 to avoid it).
- A poll-boundary `except KeyboardInterrupt` makes the rare cut-short case exit
  cleanly (save + stop) instead of unwinding with a traceback.
- Clearer interrupt message: "finishing the current poll, then saving and exiting".

## 2026-06-25 — Changes tab: 'Arr.' header + change-order table restructured

- Arrangement header now displays as 'Arr.' (display only — the internal label
  stays 'Arrangement', so change-log matching and column lookups are untouched).
  Applied via `_header_cells` and the Live Queue display-header constants.
- 'Change orders today' now reads like the other tables: Time, Job #, Folder,
  Quote Run, CO#, Oper, Design, Customer, What changed. The old free-text
  'Change' column is the CO# column (still shows CO#old -> CO#new). Folder /
  Quote Run reuse the standard cell builder.
- Tests: `test_change_orders_table_columns_and_abbrev_header`; updated
  `test_changes_today_log_sections`.

## 2026-06-25 — Changes tab: Arrangement/Size suffix -> hover comment

Mirror the Live Queue on the Changes tab to save width: Arrangement shows just
the 'A/X' code and Size its main value, with the descriptive suffix moved to a
hover comment. Applied to New orders / Removed (via `arrange_comment=True`), the
'orders that changed today' was-row + instance rows, and the change-order table
(new `_suffix_comment_cell` helper). Test:
`test_changes_arrangement_size_suffix_moves_to_comment`.

## 2026-06-25 — Changes tab: New/Removed columns aligned with the changed table

The 'Orders that changed today' table leads with a Time column, so its Job # sits
in column B and Folder in C. 'New orders today' and 'Removed / completed today'
had Job # in A and Folder in B, so everything from Folder on was off by one column
between sections. Fix: `_job_table` now inserts a blank spacer column right after
Job #, so Job # stays in A, B is blank, and Folder / Quote Run / CO# / … line up
in the same columns (C/D/E/…) across all three sections. Test:
`test_changes_today_columns_align_across_sections` in `test_live_sheets.py`.

## 2026-06-25 — Live tabs no longer drop rows after a busy-Excel write

Symptom: "all the orders vanished" — the Live Queue tab showed ~18 of 56 on-board
orders even though every poll logged `56 on board | removed=0` and the master log
was intact. No data was lost; the tab just failed to redraw.

Cause: the live tabs are drawn incrementally — each poll writes only the rows
whose signature changed (`master['lq_sigs']` / `['oh_sigs']`). `watch._plan`
committed those signatures as soon as the ops were *planned*, before the Excel
write. When a write failed (Excel busy / `OLE error 0x800ac472`, e.g. a dialog or
co-authoring sync), the store believed the rows were on the sheet, so the next
poll planned no op for them and they stayed missing until a restart — and if the
restart's first render also hit a busy Excel, it re-poisoned the store.

Fix (commit-after-success + idempotent appends):
- `watch._plan` now returns `(ops, commit)`; `_render_master` calls `commit()`
  only for tabs that `update_master_workbook` reports as rendered.
- `update_master_workbook` now returns the set of tab names that rendered without
  error (was a bare bool, which nobody used).
- `apply_upserts` / `apply_order_history` appends are now idempotent: if a key is
  already on the sheet (a re-planned row after a failed write), update it in place
  instead of adding a duplicate. A failed write therefore self-heals on the next
  poll — no restart, no duplicates.

Performance: unchanged in steady state — same ops/writes per poll; the only added
work is an O(1) keymap lookup (the keymap is already read each poll) and redrawing
rows after a failed write, which is the point.

Tests: `test_watch_render_commit.py` — a failed write leaves signatures
uncommitted (rows re-planned next poll); a successful write commits and then skips
unchanged rows.

## 2026-06-25 — Baseline poll no longer floods "Orders that changed today"

Symptom: the Changes tab showed a grey "changed today" row under (nearly) every
order, all stamped 5:00 AM, with the whole Sales-Order block (Size, Arrangement,
Description, Motor Pos, Class, …) flagged red — even when nothing had moved.

Cause: `watch.poll_once` appended `live_master.update`'s deltas to today's change
log on *every* poll, including the silent start-of-day **baseline** poll. The
baseline poll diffs the board against *yesterday's* saved master, so its deltas
are overnight moves — and, for orders the raw start-of-day seed re-enters
without their enrichment yet, `_keep_better_enrichment`'s guard only protects
fields when the stored order has a `so_pdf`; an order with SO fields but a blank
`so_pdf` has them recorded as `value -> ''`. Either way these are not changes
that happened *during* today's watch, so they shouldn't be "changed today".

Fix: `poll_once` now appends to the change log only when `not baseline`. The
master is still folded/updated on the baseline poll; we just don't record its
deltas. Later (non-baseline) polls log genuine intraday changes as before, so
each real move still gets its own grey row.

Tests: `test_watch_baseline.py` — baseline poll writes nothing to the change
log; a later normal poll still logs a real End-Date move.

## 2026-06-17 — Every helper feeds the one master store

DG: incorporate everything we know about each order into live_master, and have
every helper add what it collects.

- `live_master.merge_order(master, job, fields)` — the single merge primitive:
  writes only non-empty values, never regresses an existing value to empty, and
  creates an unseen order off-queue. Tested.
- `master_sync.py` — reads each helper's store off disk (no heavy imports) and
  merges via per-source adapters: autocad (`dwg_extras`/type/folder), quote_runs
  (drive-run fields), line_items (items + tags + customer/co/so_pdf), backfill
  (full SO spec). `run("<source>")` does load-merge-save; CLI
  `python master_sync.py [sources]` consolidates on demand.
- Hooked the end of each helper's main (`autocad_scan`, `quote_run_scan`,
  `line_items_scan`, `backfill_orders`) to `master_sync.run(...)` — best-effort,
  so any time a helper runs, its data lands in the master.
- Tests: `test_master_sync.py` (in CI) + a `merge_order` case in
  `test_live_master.py`.
- Note: this can grow live_master.json to the full backlog (~12K). Concurrency:
  run big syncs off the watcher's clock (both do load-merge-save).

## 2026-06-17 — Master JSON + change log; Changes-tab rebuild; UI fixes

Big batch from DG:

- **One master JSON** (`live_master.json`) is the source of truth: per order it
  holds all the info we have (board fields + SO design/size/arrangement/temps +
  CO# + DWGs + line-item features). `live_master.update` now compares each scan
  field-by-field (`tracked_values`/`_diffs`), updates the master, and RETURNS the
  modifications. Initial population (''->value, i.e. enrichment filling in) is
  skipped so the log stays meaningful.
- **`change_log.py`** — per-day `change_log_<date>.json` of field events
  {time, job, customer, field, old, new}. A field changing N times/day = N
  events (N lines). Archived with the other dated files (runstate).
- **Changes tab rebuilt** as a today log: New orders today, **Change orders
  today (CO# — restored)**, Orders that changed today (one time-stamped line per
  field modification, newest first), Removed/completed today. Replaces the old
  this-morning/vs-yesterday snapshot groups (which mis-flagged everything new).
- **"New today"** now reads main.py's today snapshot + diff (+ arrived-since-
  morning), so launching watch.py any time after the 5 AM run flags the right
  orders. Verified end-to-end.
- **Order History reorder**: data columns first (Job # pinned), then On Queue/
  Added/Left right before the DWG matrix; AutoFilter added (sortable).
- **Live Queue Added** written as Text -> shows the AM/PM label, not a 24h serial.
- Tests: `test_change_log.py` (new, in CI) + change-detection cases in
  `test_live_master.py`; Changes-tab test rewritten. 80+ pure tests green.

## 2026-06-16 — Order History = stable 12K log with DWG + Feature matrices

Refined the master log per DG:

- **Custom DWGs**: removed the text column from Live Queue; Order History now
  carries the full green-✓/red **AutoCAD DWG matrix** AND a new **line-item
  Feature matrix** (one column per tag, from `line_items` store tags), side by
  side with a **vertical divider** (a thin gray column).
- **12K backlog**: Order History merges the live master with the *whole*
  line-items store (`_oh_orders`), so it shows ~1 row per order ever (live +
  backlog), using the data we already gathered.
- **Stable log**: Order History shows only identity + SO-spec + the matrices +
  presence flags — NOT churny board fields — so a row's signature changes only on
  add / On-Queue flip. The 12K log isn't rewritten when a date/price ticks.
- **Efficient render**: `apply_order_history` bulk-writes all rows once, colors
  the matrices by **conditional formatting** (green ✓ / red blank, key-column
  guarded so the empty area isn't painted), draws the divider, and uses
  `=HYPERLINK()` formulas so 12K links go in the bulk write. Rebuilds the tab if
  the matrix column set grows (`reset_sheet` + sig reset).
- **Sigs**: moved to flat `lq_sigs` / `oh_sigs` maps in the master (so
  backlog-only orders are tracked, not just live ones).
- **"New today"**: now judged vs **this morning's** frozen baseline (not the
  previous day).

Smoke-tested the merge + matrix spec + sig planning across cycles (append all
once, no-op when unchanged, single update on an On-Queue flip). COM render needs
the on-PC smoke test.

## 2026-06-16 — Master log: incremental upsert instead of repaint

DG: stop repainting the whole tab every cycle (it reset filters). New model is a
chronological master log updated by **upsert keyed on order #**:

- **`live_master.py`** — the all-time log (`live_master.json`), one entry per
  order ever seen: `added` (set once), `left`, `on_queue`, latest `job`. Unlike
  history.json it's append-only (never pops a returning order). 3 tests.
- **`live_sheets.py`** — added stable-schema record builders
  (`live_queue_records`, `order_history_records`) keyed on order #, a compact
  **Custom DWGs** text column (the variable-width matrix fought the fixed
  schema; full matrix stays in the daily report), `row_sig` (md5 hex so it
  survives a JSON round-trip), and a pure `plan_upsert` (append/update/delete;
  unchanged rows -> no op). Order History gains On Queue/Added/Left columns.
- **`live_excel.py`** — `apply_upserts`: reads the key column for row positions
  (robust to a coworker sorting), writes/append/updates rows in place, deletes
  departed ones (Live Queue), re-extends AutoFilter only when the row count
  changed. No Cells.Clear() per cycle, so filters/sort/scroll persist. Each
  upsert tab is wiped+rebuilt once per process start (clean slate vs the old
  full-grid content; sigs reset to match via watch `_force_rebuild`).
- **`watch.py`** — maintains the master each cycle, plans upserts (storing sigs
  in the master so change-detection survives restarts), renders Live Queue
  (delete-on-leave) + Order History (append-only) incrementally, keeps Changes
  as a full-repaint snapshot. **Line Items tab dropped** (multiple rows/order
  was confusing).

Verified the op-planning end-to-end across cycles (unchanged->no-op,
changed->update, new->append, left->delete from Live Queue / stays in Order
History). COM apply still needs the on-PC smoke test.

## 2026-06-16 — Live workbook becomes the master (4 tabs) + email link

DG wants the live co-authored workbook to be the team's master sheet, not a
stripped board. Built the full multi-tab master, all written through Excel COM:

- **`live_sheets.py`** — a PURE model layer (`Cell`/`Sheet` with named style
  intents) that builds each tab, reusing `excel_writer`'s COLUMNS + label helpers
  and `compare.diff_queues` so the live master and the daily report can't drift.
  Tabs: **Live Queue** (Added col + every Full Queue column + DWG matrix + date/
  new fills + hyperlinks + totals + AutoFilter), **Changes** (two date-labeled
  groups — since this morning's frozen baseline, and vs yesterday), **History**,
  **Line Items** (one row per order×normalized item; AutoFilter the Normalized
  column to find orders — the in-workbook `find_orders`). 8 tests in
  `test_live_sheets.py` (added to CI).
- **`live_excel.py`** — rewritten as a GENERIC renderer: bulk-write values, then
  map named fills/fonts to Excel BGR colors, add hyperlinks, freeze, AutoFilter,
  autofit. A per-tab fingerprint cache repaints a tab only when its content
  changed, so a coworker's active filter/scroll isn't reset every cycle.
- **`watch.py`** — each cycle builds all four sheet models and calls
  `update_workbook`. Freezes an enriched start-of-day baseline
  (`live_baseline_<date>.json`) on the first poll for the intraday diff.
- **Email** — `emailer` now leads with an active `LIVE_WORKBOOK_LINK` and writes
  dates in full ("Tuesday, June 16, 2026" / "vs Monday, June 15, 2026"). The
  dated `queue_<date>.xlsx` is still saved for the archive; attach it only with
  `EMAIL_ATTACH_REPORT=1`. `compare.diff_queues` now returns `prev_date`.

Still Windows-only for the COM/notify paths (untestable in the sandbox); the
pure sheet model + diff are what's tested. NOTE: the "Features" column on the
board is kept as a quick tag summary; the real per-item detail is the new Line
Items tab.

DG: Line Items tab covers the WHOLE backlog (every order in the line-items
store), not just the board, so the item search spans all history. Empty until
the store is built (`line_items_scan.py`). Perf note: a large backlog makes that
tab's repaint heavier, but the fingerprint cache only repaints it when the store
changes (a few new orders/day) — revisit if it ever drags.

## 2026-06-16 — Live intraday watcher (queue stays fresh all day)

The queue went stale between 5 AM runs. New `watch.py` keeps a **co-authored
Excel workbook** live all day without re-paying the slow enrichment. The insight:
the cheap part of a run is the board scrape (`scrape_queue` — order numbers + row
data, no modals); the slow part is per-order enrichment (`enrich_with_sales_orders`).
So the watcher polls the board every couple of minutes and runs enrichment **only
for order numbers it hasn't seen yet today**, stamping each with its first-seen
("added") time. New modules:

- **`live_state.py`** — pure, unit-tested per-day memory (`live_state_<date>.json`):
  detects new/returning/removed orders, stamps a stable `first_seen`, refreshes
  volatile board fields each poll without clobbering enrichment, seeds the
  start-of-day baseline from the morning daily snapshot, sorts present orders
  newest-first. 8 tests in `test_live_state.py` (added to CI).
- **`live_excel.py`** — writes the live board into the co-authored workbook by
  driving the **desktop Excel app via COM** (mirrors `emailer.py`'s Outlook
  automation) so co-authoring/cursors sync to coworkers; openpyxl can't do this
  (it replaces the whole file). One bulk `Range.Value` write per cycle, light
  stable formatting, `SaveCopyAs` for the dated morning snapshot.
- **`notify.py`** — per-new-order **Windows toast** (winotify, PowerShell
  fallback) + **Microsoft Teams** Incoming-Webhook card (stdlib urllib, no dep)
  so coworkers + phones get pinged. Both best-effort.
- **`watch.py`** — the loop: 5am–5pm window (configurable), default 2-min
  interval, first poll = silent baseline + morning snapshot, restart-safe.
  `--once` / `--now` flags. Own lightweight logging (doesn't drag in anthropic).

Config in `.env`: `LIVE_WORKBOOK_PATH`, `POLL_INTERVAL_SECONDS`, `WATCH_START`/
`WATCH_END`, `TEAMS_WEBHOOK_URL`, `LIVE_TOAST`, `LIVE_MORNING_SNAPSHOT`. Can't be
exercised from the sandbox (Excel COM / Teams / toast are all on the Windows
box) — the pure state logic is what's tested; the COM/notify paths follow the
established lazy-import, best-effort pattern and need a smoke test on the PC.

## 2026-06-16 — Text PDFs parse as Qt Runs

After the rescan (unknown 55→18, CB text 250→265), the ~85 text-bearing PDF
runs were still going through the *generic* key/value sweep. But they're the
same CBC selection-program Qt Run, just saved as PDF (DG: "these are all
technically chicago blower runs" — the `CHICAGO BLOWER`/`SN#` header marks the
*selection-program layout*, not a CB-vs-other distinction). So:

- `drive_run.parse_drive_run_pdf` now also returns the FULL extracted `text`.
- `PdfQuoteRun.extract`: if that text carries the Qt Run header
  (`is_selection_program`, shared `SELECTION_PROGRAM_MARKERS`), parse it with
  the full `_parse_chicago_blower` field set; otherwise keep generic key/value
  (vendor quotes, markups, etc.). Routing tested by stubbing the PDF text
  extraction (pdfplumber can't run in the sandbox) — 2 new tests.

So the same Qt Run now yields the same fields whether it's `.txt`, `.docx`, or
`.pdf`. Re-run `quote_run_scan.py --rescan` to upgrade the stored PDF rows.

## 2026-06-15 — First full backlog sweep + refinements

DG ran `quote_run_scan.py` over the whole Z: tree: **12,873 jobs in ~22 min,
365 with a run, 554 runs**. Templates: cbc_qt_run_text 250, pdf 219 (146 of
them text-less drawings), unknown 55, d64 30. The log surfaced three fixes
(all done + tested):

- **CAD false positives** — HDX layout files named `QT RUN-...`
  (`.dwg/.sldasm/.slddrw/.dwl2/.bak`) were matching as "runs". Restricted the
  folder finder (`sales_orders._run_files_in_folder`) to document extensions
  (`RUN_DOC_EXTS`) so drawings are ignored. This also helps the daily run.
- **Office temp/lock files** (`~$...`) were parsed and errored ("not a zip");
  now skipped in the folder finder.
- **`.docx` is the biggest real gap** (`QT RUN.docx`, `quote run.docx`,
  `MARKUP PER FEA.docx`). Added stdlib `.docx` text extraction
  (`templates._docx_to_text`, zip→word/document.xml, no python-docx dep) and
  put `.docx` on the text templates, so a CB run saved as Word routes to the CB
  parser by content marker (markup versions fall to generic KV).

Deferred: `.doc` (old-binary Word — many are *damper* quote runs, needs a
heavier extractor) and `.msg` (Outlook) stay UNRECOGNIZED for now. D64 `.xlsx`
parses via the best-effort cell sweep (30 found) — refine when a real sheet is
pasted back.

**Re-run with `--rescan`** to re-evaluate the stored jobs against the new
extension filter + `.docx` support (the progress store holds the first pass).

## 2026-06-15 — Whole-backlog quote-run sweep (`quote_run_scan.py`)

DG: "check everything in history for quote runs." Chose (with DG) a **pure Z:
AutoCAD folder sweep** (no login, fast, resumable) over the slow online
backfill. New `quote_run_scan.py` mirrors `autocad_scan.py`: enumerates every
`<type>/<intermediate>/<job>` folder (reuses `iter_job_folders`), finds run
files recursively with `_run_files_in_folder` (catches `ENG REF\` + `history\`
copies), parses each through `parse_quote_run`, and writes
`backlog/quote_runs.xlsx` — one row per run with the matched template, the core
fields, an "Other" catch-all (so new template fields are never dropped), and a
**Status** column (OK / NO FIELDS / UNRECOGNIZED FORMAT / PDF-no-text) that
surfaces which formats still need a template. Resumable JSON store; same
`--min-job/--max-job/--limit` plus `--range` and `--list FILE`. Pure logic
(status, row-flatten core/Other split, recursive scan_one) tested in
`test_quote_run_scan.py` (5 tests, added to CI). Can't see runs that live only
in an order's online documents — that's still `backfill_orders.py`'s job.

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

**CB parser tuned on 4 real runs (2026-06-15, `check_orders.py 421579 421237
419624 420990`).** The batch ran clean after the hardening. Four GL/PFD runs
(designs 6195 / 1904 / 1904 / 1910B) refined the parser:
- Wheel-construction gauge column is a fraction **or** a decimal with the GA in
  parens (`0.048 (18)`, `0.179 ( 7)`), and blades carry a descriptor
  (`BLADES/2 RIB`). The fraction-only pattern was dropping Blade/Sideplate
  materials on the PFD fans — now handled (`_GA` sub-pattern + descriptor).
- Some runs have **no `SN#`** (order # on a bare line) — Serial now falls back
  to a bare-number header line. Added **Fan Type** (the `PACKAGED FORCED DRAFT
  FAN`-style descriptor after the spec line), **Hub** (`19-5-1056`),
  **Coupling** (`FALK T10`), and flags **FEA Analysis** / **Factory Run Test**.
  Direct drive is inferred from a `COUPLING` line when there's no `BELT DRIVEN`.
- Second real run (421237) locked in as `REAL_CBC_QT_RUN_421237` in the tests.
- Note: 420990 listed the same Qt-run doc twice → downloaded as `_1`/`_2`
  (the `drive_run_count > 1` "review" case); not a parser issue.

**More CB runs (2026-06-15, `check_orders.py 421473 421572`):**
- Confirmed the spec line also comes **space-delimited** (`SIZE 37 DESIGN 16A
  LS  ARR 9H  100.0 PCT ...`), not just comma-delimited — the field-by-field
  patterns already handle both.
- LS-class wheels have richer tables (LINER / GUSSETS / HUB TUBE). Added
  **Liner Material** (captures e.g. `PLAIN FIRMEX`, a notable wear liner), a
  **Wheel Material** fallback for runs with no construction table (just
  `WHEEL MATERIAL A569 HRS`), and fan **Class** (`CLASS 4`). Locked job 421572
  in as `REAL_CBC_QT_RUN_421572`.
- Jobs can have **several differing runs** (quote vs production vs history copy,
  different revisions/materials) across the documents + ENG REF + history
  folders. `check_orders` surfaces all of them — useful, and the daily run's
  `_run_docs` still picks the documents' run as primary.
- A PDF run can be a **drawing with no text layer** (421572's "Inlet Box Liners
  ONLY.pdf" yielded nothing); `check_orders` now says so instead of implying a
  missing template.

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
