#!/usr/bin/env python3
"""List the OPEN Sales-Order review notes waiting to be applied to the parser.

Reads the note queue the user's machine publishes to the `order-data` branch
(so_review_notes.json), drops any note already resolved in this repo's tracked
ledger (so_review_handled.json), and prints what's left newest-first. Run from
the repo root:  python .claude/skills/apply-so-review-notes/scripts/list_open_notes.py
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[4]
BRANCH = "order-data"


def _git(args: list[str]) -> subprocess.CompletedProcess:
    return subprocess.run(["git", "-C", str(REPO), *args],
                          capture_output=True, text=True)


def main() -> int:
    fetch = _git(["fetch", "origin", BRANCH])
    if fetch.returncode != 0:
        print(f"Could not fetch origin/{BRANCH}:\n{fetch.stderr.strip()}", file=sys.stderr)
        print("The note queue only exists once the user has published order data "
              "(launcher 'Publish Order Data'). Ask them to publish, or check "
              "DATA_PUSH_BRANCH.", file=sys.stderr)
        return 1

    show = _git(["show", f"origin/{BRANCH}:so_review_notes.json"])
    if show.returncode != 0:
        print(f"No so_review_notes.json on origin/{BRANCH} — nothing published yet.",
              file=sys.stderr)
        return 1
    notes = (json.loads(show.stdout) or {}).get("notes") or []

    handled_path = REPO / "so_review_handled.json"
    already = set()
    if handled_path.exists():
        led = json.loads(handled_path.read_text(encoding="utf-8"))
        already = {int(m.get("id", -1)) for m in led.get("handled", [])}

    open_notes = [n for n in notes
                  if n.get("status") != "handled" and int(n.get("id", -1)) not in already]
    open_notes.sort(key=lambda n: str(n.get("created_at") or ""), reverse=True)

    if not open_notes:
        print("No open notes. (All published notes are handled or already in the "
              "local ledger.)")
        return 0

    print(f"{len(open_notes)} open note(s), newest first "
          f"(already-resolved ids skipped):\n")
    for n in open_notes:
        anchor = n.get("item_no") or n.get("row_key") or ""
        print(f"#{n['id']}  order {n.get('order')}  [{anchor}]  ({n.get('created_at','')})")
        print(f"    row : {n.get('item_text','')}")
        print(f"    note: {n.get('note','')}")
        print()
    print("Look up how each order currently parses with:  "
          "python .claude/skills/apply-so-review-notes/scripts/show_order.py <order>")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
