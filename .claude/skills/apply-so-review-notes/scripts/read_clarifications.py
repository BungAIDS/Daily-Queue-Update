#!/usr/bin/env python3
"""Parse the answered clarifications document back into {id: answer}.

Reads the document the user filled in — by default from the published copy on
the order-data branch (git show origin/order-data:so_review_clarifications.md),
or from a local path passed as arg 1 — and prints JSON mapping each note id to
the text the user typed after its `ANSWER>`. Notes left blank are omitted, so
what you get back is exactly the set that's now ready to implement.

    python .claude/skills/apply-so-review-notes/scripts/read_clarifications.py
    python .claude/skills/apply-so-review-notes/scripts/read_clarifications.py path/to/file.md
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[4]
NOTE_RE = re.compile(r"^NOTE\s+#(\d+)\b")


def _load(argv: list[str]) -> str | None:
    if len(argv) == 2:
        return Path(argv[1]).read_text(encoding="utf-8")
    subprocess.run(["git", "-C", str(REPO), "fetch", "origin", "order-data"],
                   capture_output=True, text=True)
    r = subprocess.run(
        ["git", "-C", str(REPO), "show", "origin/order-data:so_review_clarifications.md"],
        capture_output=True, text=True)
    if r.returncode != 0:
        print("No so_review_clarifications.md on origin/order-data yet — the user "
              "hasn't published answers.", file=sys.stderr)
        return None
    return r.stdout


def parse(text: str) -> dict[str, str]:
    """Capture only what the user typed inside each note's answer box.

    Each note is `NOTE #<id> ...` followed by a box delimited by a `>>>>>` open
    marker and a `<<<<<` close marker. The metadata lines (Your note / Parses
    now / My question) sit before the box and are ignored; capture runs strictly
    between the markers. Blank boxes are dropped, so the result is exactly the
    notes the user actually answered.
    """
    answers: dict[str, str] = {}
    current: str | None = None
    capturing = False
    buf: list[str] = []

    def flush() -> None:
        nonlocal capturing
        if current is not None and capturing:
            ans = "\n".join(buf).strip()
            if ans:
                answers[current] = ans
        capturing = False

    for raw in text.splitlines():
        line = raw.rstrip("\n")
        stripped = line.strip()
        m = NOTE_RE.match(stripped)
        if m:
            flush()
            current, buf = m.group(1), []
            continue
        if stripped.startswith(">>>"):        # answer box opens
            capturing, buf = True, []
            continue
        if stripped.startswith("<<<"):        # answer box closes
            flush()
            continue
        if capturing:
            buf.append(line)
    flush()
    return answers


def main() -> int:
    text = _load(sys.argv)
    if text is None:
        return 1
    answers = parse(text)
    print(json.dumps(answers, indent=2, ensure_ascii=False))
    if not answers:
        print("(no answers filled in yet)", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
