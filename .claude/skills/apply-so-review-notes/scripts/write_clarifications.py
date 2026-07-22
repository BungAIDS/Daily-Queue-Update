#!/usr/bin/env python3
"""Write the standard "clarifications needed" document for deferred review notes.

Input: a JSON file (arg 1) that is a list of deferred items, each:
  {"id": 124, "order": "421967", "row": "[SHIP TO]", "note": "all these",
   "current": "captured as a [SHIP TO] component",
   "question": "which rows, and should they be dropped from capture?"}

Output: a simple text document (arg 2, default so_review_clarifications.md) with
one block per note and an `ANSWER>` line for the user to type under. The format
is deliberately plain so it opens in Notepad and is unambiguous to type in and
to parse back (see read_clarifications.py). Keep questions concrete — the user
answers fastest when the real row, how it parses now, and the exact ambiguity
are already on the page.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path

RULE = "=" * 62
ANSWER_OPEN = ">>>>>>>>>>  TYPE YOUR ANSWER ON THE BLANK LINES BELOW  >>>>>>>>>>"


def _answer_close(note_id: int) -> str:
    return f"<<<<<<<<<<  (end of answer for note #{note_id})  <<<<<<<<<<"


HEADER = """\
# SALES-ORDER REVIEW — CLARIFICATIONS NEEDED
#
# I couldn't confidently apply the notes below, so I left them OPEN and I'm
# asking instead of guessing.
#
# HOW TO ANSWER: for each note, type on the blank lines in the box between the
#   >>>>>  TYPE YOUR ANSWER ...  >>>>>
#   ...your answer here...
#   <<<<<  (end of answer ...)  <<<<<
# markers. Write as much as you like; blank lines inside your answer are fine.
# To skip a note for now, just leave its box empty.
#
# When done: save, close, and click "Send Clarifications" in the launcher.
# Don't edit the "NOTE #..." or ">>>>>"/"<<<<<" marker lines — they tell me
# which answer goes with which note. Lines starting with "#" are ignored.
#
# Generated: {when}
"""


def main() -> int:
    if not 2 <= len(sys.argv) <= 3:
        print("usage: write_clarifications.py <deferred.json> [out.md]", file=sys.stderr)
        return 2
    items = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    out = Path(sys.argv[2]) if len(sys.argv) == 3 else Path("so_review_clarifications.md")

    blocks = [HEADER.format(when=datetime.now().isoformat(timespec="seconds"))]
    for it in items:
        lines = [
            "",
            RULE,
            f" NOTE #{it['id']}   ·   order {it.get('order','')}   ·   row: {it.get('row','')}",
            RULE,
            f" Your note : {it.get('note','')}",
        ]
        if it.get("current"):
            lines.append(f" Parses now: {it['current']}")
        lines.append(" My question:")
        for qline in str(it.get("question", "")).splitlines() or [""]:
            lines.append(f"   {qline}")
        lines += [
            "",
            ANSWER_OPEN,
            "",
            "",
            "",
            _answer_close(it["id"]),
        ]
        blocks.append("\n".join(lines))

    out.write_text("\n".join(blocks) + "\n", encoding="utf-8")
    print(f"Wrote {out} with {len(items)} clarification(s).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
