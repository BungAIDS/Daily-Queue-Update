# Quote-Run Data Extraction — Handoff Plan

**For the next session/model picking this up: read this whole file, then
CLAUDE.md, then the 2026-07 WORKLOG entries.** The mission: extract *every*
useful variable from every quote run, keep it correct, and keep the remote
data loop working. This plan is ordered by value; each item says WHY, HOW,
and how to VERIFY.

## How this system works (the 60-second map)

```
Z:\AUTOCAD\...\JOBS ──quote_run_scan.py──> quote_run_scan_progress.json  (store)
   (text runs)             │                  runs[] carry: fields, raw_lines
scanned PDFs ──pdf_vision.py (Claude API)──>  (full doc text), vision.transcript,
                           │                  vision.suspect, mtime, status
                           ▼
        quote_runs.xlsx  +  live_master.json (master_sync; drive_runs history)
                           │
        data_push.py ──> git branch `order-data`  <── remote assistant reads this
```

- **The store is a corpus.** Every run keeps its full text (`raw_lines`, cap
  250) or vision `transcript`. New extraction patterns are applied WITHOUT
  touching Z: or the API: `python quote_run_scan.py --reparse-stored`
  (seconds). This is the core loop: design pattern → test on corpus →
  reparse-stored → publish.
- **Parser**: `templates.py::_parse_chicago_blower` (targeted regexes,
  first-match-wins per label, `re.I|re.M`) + `_parse_outline_dims` (block
  parser for the AXIAL/SIDE VIEW table). Section-based — patterns fire when
  their section exists; do NOT gate on arrangement.
- **Vision runs**: model fields win overlaps (it saw the image; the transcript
  is derivative OCR). Patterns fill gaps. `pdf_vision.apply_vision_qc` repairs
  garbled values from clean transcript parses and flags the rest
  `CHECK VISION`. A CHECK VISION run is re-read ONCE by the next
  `python pdf_vision.py`, but the re-read is **escalated** (higher-res render +
  the model is told exactly what looked wrong) and the two readings are
  **compared**; if they still don't converge after `MAX_VISION_ATTEMPTS` (2) the
  run goes to `NEEDS HUMAN` — terminal, never auto-re-read/re-paid again, orange
  in the workbook with a reason. A later parser/repair that makes it clean
  clears it back to OK. So: re-reads escalate, compare, and stop — no infinite
  re-pay, no silent same-model repeat.
- **Multiple runs per order** (71/381): `run_rank.py` ranks by CO# > REV >
  file mtime > field count. Master keeps ALL runs under `drive_runs`.
- **Money rules**: a vision result is never re-paid (`--redo` overrides;
  `carry_vision_forward` preserves across `--rescan`). Never break these.

## Analysis loop with the remote assistant

1. User runs a scan / reparse / vision batch (launcher tasks exist for all).
2. `data_push.py` publishes the 9 data files to the `order-data` branch
   (auto if `DATA_PUSH_ON_CHANGE=1`).
3. Remote session: `git fetch origin order-data && git show FETCH_HEAD:quote_run_scan_progress.json`
   → full corpus. Measure BEFORE building (see "evidence bar" below).

## Current state (2026-07-06)

- 467 runs / 381 orders with runs. Status: 450 OK, 3 NO FIELDS (DG: ignore),
  14 UNRECOGNIZED (.doc/.xls). 128 scanned PDFs read by vision (all quote
  runs, none drawings); 123 have transcripts (5 from the pre-transcript trial
  don't — QC flags them).
- ~54 fields/run average after the arrangement-section templates.
- **59 vision runs flagged `CHECK VISION`** by QC (list in WORKLOG / filter
  the xlsx Status column): dominant cause is OCR misreading **S as 8/5** in
  arrangements ("7S1"→"781", "8S"→"88"), plus model/transcript disagreements
  on CFM/SP/RPM. Fix = re-read: `python pdf_vision.py` picks them up
  automatically.
- ~25% of content lines in the corpus are still untouched by any pattern
  (measured; see P2 for the ranked clusters).

## Priority roadmap

### P0 — Close the correctness loop (do first, mostly user actions)
- User: `git pull` → `python quote_run_scan.py --reparse-stored` → publish.
- User: `python pdf_vision.py` re-reads the 59 flagged PDFs (~50¢ total).
  Then `--reparse-stored` again → publish. Remote: re-run QC counts; confirm
  flags shrink; investigate any repeat offenders (model may be consistently
  wrong on those — needs eyeballs).
- User spot-checks the 8-order list (419103, 421311, 412477, 420402, 404346,
  404795, 400567, 410087) against the PDFs. Each ❌ becomes a pattern fix +
  regression test.

### P1 — Coupling-shaft block + accessory lines (next pattern batch)
Measured top uncaptured clusters (counts = lines across corpus):
- 164x `MAX SHAFT DIAMETER AT FAN SHAFT HALF COUPLING = N.N` (+ MIN 164x,
  NOM 110x) → fields `Coupling Max/Min/Nom Shaft Dia`.
- 94x `KEYWAY DIMENSIONS FOR HALF COUPLING = N.N X N.N` → `Coupling Keyway`.
- 164x `THE BEARING TEMPERATURE AT N F AMBIENT IS ACCEPTABLE.` → flag field.
- 108x `REINFORCED INLET CONE INCL. PER SK-N-N` → `Inlet Cone SK`.
- 106x `SHAFT SAFETY GUARD  n  n` → accessory presence (+weight/price cols).
- 81x `FRAME BASED ON ODP, OTHERWISE SPECIFY ITEM N` → Motor Enclosure
  default when no explicit TEFC/ODP (don't overwrite an explicit one).
- 88x `BASE NOT SIZED FOR GEAR BOX...` note; 122x `ORDER N` (cross-check vs
  job#); 76x `OVERALL DIMENSIONS:` block (L×W×H — parse the lines after it).
METHOD: same as before — pin shapes from the corpus, add to `_CB_PATTERNS`,
add real-fixture tests, run the coverage matrix script (in WORKLOG/git
history of this chat's scripts), `--reparse-stored`, verify counts.

### P2 — Wheel-construction table: per-component WR2 + weight
The table rows (`BLADES 1/4 ASTM... 81 54`) already yield material+gauge; the
two trailing columns (WR2, weight) are dropped. Extend the row regexes with
two more capture patterns per component (`Blade WR2`, `Blade Weight Lb`, ...).
Watch the HUB row (5 numbers: wr2, weight, total weight, price).

### P3 — Multi-printout files (correctness risk)
Some files concatenate several run printouts (400567 has 4 SIZE values in one
file). First-match-wins reads the FIRST printout. Split text on the
`---...--- / CHICAGO BLOWER CORP.` header boundary, parse each segment, pick
the most-current segment (reuse `run_rank.revision_key` on in-text CO/REV
markers + completeness). Verify against 400567, 400076.

### P4 — The 14 UNRECOGNIZED (.doc/.xls) — mostly damper quote runs
Old binary formats; no Python reader in-repo. Options: (a) DG bulk-converts
to .docx via Word once (folder script), then existing docx path reads them;
(b) add a `docx2txt`-style .doc extractor dep. Then design a DAMPER template
(different product line — fields unknown; get a sample first). The `damper`
column already flags them.

### P5 — D64 wheel-construction xlsx (27 runs)
Still generic key/value. Needs ONE real sheet from DG (or read cells from the
stored raw_lines cell-dump) to pin a real mapping. Ask DG which fields matter
(inner/outer wheel rows?).

### P6 — mtime backfill (one full rescan)  [cap: DONE]
The raw_lines cap is now a 10000-line runaway backstop (was 250) — no real doc
truncates. A full `--rescan` still backfills `mtime` on every run (strengthens
run_rank) and re-stores the 42 previously-clipped runs at full length; DG runs
it once (~10 min; vision results carry forward), then publishes.

### P7 — Surfacing decisions (DG's call, then trivial)
Which new fields deserve their own Live Queue / report columns vs staying in
the details summary + xlsx? Candidates: Motor Frame, Blades, Total Price.
One-line changes in `excel_writer.COLUMNS` / `_CB_SUMMARY_ORDER`.

## Evidence bar (how we work — keep it)
- **Measure on the corpus before writing a pattern** (frequency + real line
  shapes); never guess a format that isn't in the corpus — ask DG for a sample.
- Every pattern lands with a REAL-text fixture test (`test_templates.py`
  style) including the variants that almost broke it (negative numbers,
  hyphens, single-space vision transcripts).
- After changes: full suite (`for t in test_*.py: python t`), then
  `--reparse-stored` dry-run numbers in the commit message.
- All work on branch `claude/quirky-hypatia-7pwmqp`, PR to main when DG says.
- Repo must stay PRIVATE (order data on `order-data` branch).

## Confidence register (known weak spots)
1. OCR-garbled numbers that PASS plausibility (an 8 read as 3) — undetectable
   without the re-read loop / human spot-check.
2. Field semantics named by inference, unverified by an engineer ("Total
   Price" = GOOD FOR line; "Wheel Thrust"; STH/TG&P meanings).
3. Multi-printout files until P3 lands.
4. Truncated tails until P6 lands.
5. `Ambient Temp F` is extracted but not in CORE_FIELDS (lands in Other) —
   minor, fix with P7.

## Command cheat sheet
```
python quote_run_scan.py --reparse-stored   # apply new patterns, free, seconds
python quote_run_scan.py --rescan           # full Z: sweep (~10 min), keeps vision
python pdf_vision.py                        # read flagged/new scanned PDFs (API)
python pdf_vision.py --limit 5              # trial batch
python data_push.py                         # publish data to order-data branch
git show origin/order-data:quote_run_scan_progress.json   # remote: the corpus
```
