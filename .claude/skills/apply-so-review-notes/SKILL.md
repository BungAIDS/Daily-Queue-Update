---
name: apply-so-review-notes
description: >-
  Work through the user's Sales-Order review notes and apply each comment to the
  parser that normalizes sales-order line items (line_items.py). Use this
  whenever the user asks to read/apply/work through their "review notes", "SO
  review notes", "the review queue", "my most recent notes", "my comments on the
  sales orders", or says something like "go through my notes and fix the parser"
  — even if they don't name the files. The notes live on the `order-data` branch
  (so_review_notes.json); resolutions ride back in the tracked so_review_handled.json.
  This is the recurring loop for turning a human's per-line comments ("this should
  be its own component", "this shouldn't be captured", "wrong material") into
  changes to the tag rules and component builders, with tests. Prefer this skill
  over ad-hoc edits any time the request is grounded in the SO review workbook/queue.
---

# Apply Sales-Order review notes to the parser

## What this is

The user reviews how each Sales-Order line item was captured and leaves a **note**
on any row that's wrong — a flared discharge outlet that should be its own
`[EVASE]` component, a "Ship To" block that shouldn't be a component at all, a
material on the wrong part, two lines that should merge. Your job is to turn
those notes into changes to the **parser vocabulary and component logic** in
`line_items.py`, verify them, resolve each note, and hand the ambiguous ones back
to the user. This is the same loop you've run before (the `EVASE` change is the
canonical example).

The parser is **rules-as-data plus small component builders**, so most notes map
to a tidy, local change — not a rewrite. Read `line_items.py`'s header and
`DEFAULT_RULES` before you start if you haven't this session.

## The single most important rule: don't guess

A note is a terse human comment. You cannot judge it from the note text alone —
you must look at the **real line item it's attached to** and often the rest of
the order. And some notes are genuinely unclear ("all these"), placeholders
("test"), or need a decision only the user can make. **It is correct and expected
to implement only the notes you're confident about this run, leave the rest OPEN,
and ask the user about them.** A wrong parser change quietly corrupts the whole
corpus on the next re-normalize; a deferred note costs nothing. When in doubt,
defer.

## Workflow

### 0. First, pick up any answered clarifications

Before triaging fresh notes, check whether the user has answered a previous
round of your questions:

```
python .claude/skills/apply-so-review-notes/scripts/read_clarifications.py
```

This reads the answered clarification file from `order-data` and prints
`{id: answer}` for every note the user filled in. Those answers unblock notes you
deferred last time — fold them into your triage below (usually straight into
*implement now*). After you apply and resolve them, **clear the request**:
regenerate `so_review_clarifications.md` with only the notes that are *still*
deferred (an empty list writes a header with no `NOTE` blocks, which the launcher
reads as "nothing pending") and commit it, so answered questions don't come back.
If nothing is published yet, this just prints nothing — carry on.

### 1. Read the open notes

```
python .claude/skills/apply-so-review-notes/scripts/list_open_notes.py
```

This fetches `origin/order-data`, lists every OPEN note (newest first), and skips
ids already in this repo's `so_review_handled.json` ledger. "My most recent
collection" usually means the newest cluster by `created_at` (often one review
session / day) — if there are many and the scope is unclear, show the user the
list and confirm which batch they mean.

Each note gives you: `id`, `order`, an anchor (`item_no` or a `row_key` like
`component|0|...`, `attr|10|...`, `review|10|0|...`), the displayed row
(`item_text`), and the free-text `note`.

### 2. Ground every note in the real data

```
python .claude/skills/apply-so-review-notes/scripts/show_order.py <order>
```

This prints how that order currently parses (verbatim `raw`, `tags`, `component`,
`attributes`, `details`) from the published corpus. Find the row the note anchors
to and understand **what the parser does today** versus **what the note wants**.
Also check whether the change is **already implemented** in the current
`line_items.py` (rules evolve between when a note was written and now) — if so,
skip straight to resolving it (step 6) with a note that it's already handled.

### 3. Triage by confidence

Sort each open note into one of three buckets:

- **Implement now** — the intent is unambiguous and you can see the exact rule/
  builder to change. Example: *"VIBRATION BASE IS ITS OWN COMPONENT WITH NO
  ATTRIBUTES, CALLED [VIBRATION BASE]"* on a `Vibration Base` line currently
  tagged `VIBRATION ISOLATION` → split it into its own `[VIBRATION BASE]`
  component.
- **Defer and ask** — you can't confidently determine the meaning, the scope is
  unclear ("all these" on a Ship To block — all of *what*? drop them from
  capture? which rows?), or it needs a product decision (naming, whether a
  distinction matters, how two options should merge). Do **not** implement a
  best-guess. Leave the note OPEN (do not `handle` it) and collect a concrete
  question — these become the clarification file in step 7. Deferring is a
  first-class outcome, not a failure: a wrong parser change corrupts the whole
  corpus on the next re-normalize; a deferred note costs a round-trip.
- **Noise / placeholder** — "test", empty, obviously not a parser instruction.
  Leave OPEN and mention it; don't invent work. If unsure whether it's noise,
  treat it as *defer and ask*.

Only the *implement now* bucket gets code changes this run.

### 4. Apply the change to `line_items.py`

Match the existing idioms — the file is large and internally consistent, so a new
rule should look like its neighbors. Common note shapes and where they land:

- **"X should be its own component `[NAME]`"** → add a tag pattern to
  `DEFAULT_RULES["tags"]` (`re.I`, lowercase source), add a small
  `_<name>_attributes(primary, norm_blob, tags)` builder returning
  `{"component": "NAME", ...}`, and wire it into `component_attributes`. If X is
  currently a *sub-type* of another component (e.g. a `vibration_isolation_type`
  attribute), split it out and make sure the old builder no longer claims it. Mind
  precedence: whichever builder runs last wins the `component` slot — re-assert
  late if a generic builder would otherwise grab it (see how `EVASE`/`DRIVE
  COMPONENTS` re-assert near the end of `component_attributes`).
- **"this shouldn't be captured"** (page furniture, Ship To, address, a heading) →
  add a `DEFAULT_RULES["skip_patterns"]` regex.
- **spelling/abbreviation should converge** ("Évasé"→"EVASE", "IVD"→inlet vane
  damper) → `DEFAULT_RULES["abbreviations"]` or `normalize_text`.
- **wrong attribute/material/scope on a component** → the relevant
  `_<component>_attributes` builder.

When you change tagging or normalization, bump `NORMALIZER_VERSION` so a
re-normalize is signalled. Keep changes tight and precise — a pattern like
`\bevas[eé]e?\b` beats a broad one that risks false positives on other lines.

### 5. Prove it

Add or extend cases in `test_line_items.py` (plain `test_*` functions + the
`main()` runner; run with `python test_line_items.py`). Assert the noted wording
now gets the intended tag/component/attribute, and add a negative case so the new
rule doesn't over-match. Some tests in that file import `sales_orders`
(pdfplumber) and can fail in a bare container for environment reasons unrelated
to your change — if so, run the affected `test_*` functions directly (import the
test module and call them) and note that CI runs the full suite.

Mention to the user that the change re-applies to the whole existing corpus when
they run `python line_items_scan.py --renorm` (or the launcher's "Re-parse +
Refresh SO Review") — **no re-parse from PDF is needed**, because every line's
verbatim `raw` is stored.

### 6. Resolve the handled notes

For each note you actually implemented (or found already implemented):

```
python so_review.py handle <id> "<what the parser now does>"
```

This appends to the tracked `so_review_handled.json`, which rides back to the
user's machine on their next Git Update and clears the review row. Write the
resolution the way the existing ledger does — a concrete description of the **new
behavior**, not "done". Good: *"`Vibration Base` now groups as its own
`[VIBRATION BASE]` component instead of a `vibration_isolation_type` under
VIBRATION ISOLATION."* Do **not** run `handle` for deferred or noise notes — they
must stay open.

### 7. Ask about the deferred ones (the clarification round-trip)

Turn every deferred note into a question in the standard clarification document,
so the user answers on their own time and it comes back deterministically. Build
a JSON array of the deferred items — each `{"id", "order", "row", "note",
"current", "question"}`, where `current` is how the row parses today and
`question` states your best guess and exactly what you need decided — then:

```
python .claude/skills/apply-so-review-notes/scripts/write_clarifications.py \
    <deferred.json> so_review_clarifications.md
```

Write `so_review_clarifications.md` at the **repo root** (it's tracked, like
`so_review_handled.json`) so it rides down to the user on their next Git Update.
Then, on their machine:

1. **"Answer Clarifications"** (launcher) copies the request into an editable
   working file next to the workbook and opens it. The user types inside each
   `>>>>>` box, saves, closes.
2. **"Send Clarifications"** (launcher) publishes the answered file up the
   `order-data` branch with the rest of the order data.

The deferred notes stay OPEN — you do **not** `handle` them until their answers
come back. If the user is actively in the chat, `AskUserQuestion` is a fine fast
path for an enumerable choice; the file is the durable, asynchronous channel and
the one the launcher buttons are built around.

### 8. Quantify the impact

Don't just say "done" — measure whether the change moved the parser closer to
fully capturing the sales orders. Run the full corpus report:

```
python .claude/skills/apply-so-review-notes/scripts/impact_report.py
```

It re-derives the **whole** published corpus from stored `raw` text with your new
rules (what the user's `--renorm` will do) and diffs it against how the corpus is
stored today. This is the real measure — a change can ripple beyond the orders
you touched — so let it run to completion. It takes a few minutes and **prints
progress to stderr** (shards loaded, orders re-derived, percent done) so you and
the user can see it's still alive; wait for it. It reports: how many **line items
changed** (newly given a component, reclassified, or dropped by a skip rule), how
many **fewer MARKED-FOR-REVIEW rows** remain (the workbook's red rows, via the
same `parser_review_metrics` the sheet uses), and the shift in **component-capture
coverage** — the share of line items the parser resolves instead of leaving for a
human, plus which components were gained. Read the summary back to the user; that
IS the efficacy statement they're after.

On the user's machine, `--store <line_items.json>` runs the same full report
against the local store (no shard download). `--jobs <order...>` is available for
a ~1s spot-check of specific orders, but it only sees changes inside those
orders — prefer the full run for the real number.

### 9. Commit and report

Commit `line_items.py`, `test_line_items.py`, `so_review_handled.json`, and
(when there are deferrals) `so_review_clarifications.md` to the working branch and
push (follow the session's branch rules). Then tell the user: what you
implemented and resolved (id → change), **the impact numbers from step 8**, that
a clarification file is waiting if you deferred any (point them at "Answer
Clarifications"), and anything you treated as noise.

## Worked example (the shape to aim for)

Note #229, order 422029: *"VIBRATION BASE IS ITS OWN COMPONENT WITH NO
ATTRIBUTES, CALLED [VIBRATION BASE]"*. `show_order.py 422029` shows `Vibration
Base L 1,227.00` → `tags=['VIBRATION ISOLATION']`, `component='VIBRATION
ISOLATION'`, `vibration_isolation_type='VIBRATION BASE'`. Confident → implement:
add a `VIBRATION BASE` tag/component, ensure the vibration-isolation builder no
longer labels a bare vibration base as a sub-type, add tests, then
`python so_review.py handle 229 "Vibration Base now groups as its own [VIBRATION
BASE] component, not a vibration_isolation_type under VIBRATION ISOLATION."`.

Contrast note #124: row `[SHIP TO]`, note *"all these"*. Ambiguous — which rows,
and should they be dropped from capture entirely? Defer, leave open, and ask.
