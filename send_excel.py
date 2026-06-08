"""Email an existing Excel report as-is — no scrape, no diff, no AI, no stages.

    python send_excel.py                          # newest queue_*.xlsx in OUTPUT_DIR
    python send_excel.py "C:\\path\\queue_2026-06-08.xlsx"

Use this when you already have a report you're happy with and just want it in
your inbox. It regenerates nothing.

The email is built to be identical to a normal daily run's: the briefing text,
anomalies, action items, and the counts line are read back out of the report's
Changes tab and sent through the same emailer the daily run uses, with the
workbook attached. (If anything about the sheet can't be parsed, it still sends
the file with a short body rather than failing.)
"""
from __future__ import annotations

import re
import sys
from datetime import date
from pathlib import Path
from typing import Any, Dict, Tuple

from openpyxl import load_workbook

from config import EMAIL_TO, OUTPUT_DIR
from emailer import _send, send_daily_briefing

# Changes-tab section headers -> the diff key whose count the email reports.
# Matched by prefix; the trailing "(N)" carries the count.
_SECTIONS = [
    ("New orders", "new"),
    ("Returning orders", "returning"),
    ("Completed / Removed", "removed"),
    ("Changed orders", "changed"),
    ("Persistent orders", "persistent"),
]


def _latest() -> Path:
    files = sorted(OUTPUT_DIR.glob("queue_*.xlsx"), key=lambda p: p.stat().st_mtime)
    if not files:
        raise SystemExit(f"No queue_*.xlsx found in {OUTPUT_DIR}")
    return files[-1]


def _report_date(path: Path) -> str:
    """The date in the filename (queue_YYYY-MM-DD.xlsx), else today."""
    m = re.search(r"\d{4}-\d{2}-\d{2}", path.stem)
    return m.group(0) if m else date.today().isoformat()


def _parse_changes_tab(path: Path) -> Tuple[Dict[str, Any], Dict[str, int]]:
    """Reconstruct the briefing dict and the section counts from the Changes
    tab, so the email can be rebuilt to match a normal run exactly."""
    ws = load_workbook(path, read_only=True)["Changes"]
    briefing_text = ""
    anomalies: list[str] = []
    action_items: list[dict] = []
    counts: Dict[str, int] = {}
    mode: str | None = None

    for row in ws.iter_rows(values_only=True):
        a = row[0]

        if isinstance(a, str):
            section = next((key for prefix, key in _SECTIONS if a.startswith(prefix)), None)
            if section is not None:
                m = re.search(r"\((\d+)\)", a)
                counts[section] = int(m.group(1)) if m else 0
                mode = None
                continue
            if a.startswith("AI Briefing"):
                mode = "briefing"
                continue
            if a.startswith("Anomalies"):
                mode = "anomalies"
                continue
            if a.startswith("Top Action Items"):
                mode = "action_header"  # the row after this is the column header
                continue

        if mode == "briefing" and isinstance(a, str) and a.strip():
            briefing_text = a.strip()
            mode = None
        elif mode == "anomalies" and isinstance(a, str) and a.lstrip().startswith("-"):
            anomalies.append(a.lstrip("- ").strip())
        elif mode == "action_header":
            mode = "action_data"  # skip the Rank/Job #/Reason header row
        elif mode == "action_data":
            rank, job, reason = (list(row) + [None, None, None])[:3]
            if not job and not rank:
                mode = None
            else:
                action_items.append({"rank": rank, "job": str(job or ""), "reason": reason or ""})

    briefing = {
        "briefing": briefing_text or "(briefing in the attached report)",
        "anomalies": anomalies,
        "action_items": action_items,
    }
    return briefing, counts


def main() -> int:
    if not EMAIL_TO:
        print("EMAIL_TO is not set in .env — nothing to send to.")
        return 1

    path = Path(sys.argv[1]) if len(sys.argv) > 1 else _latest()
    if not path.exists():
        print(f"File not found: {path}")
        return 1

    report_date = _report_date(path)
    try:
        briefing, counts = _parse_changes_tab(path)
        # send_daily_briefing only needs each section's length, so size throwaway
        # lists to the parsed counts — the email's "Counts:" line then matches.
        diff = {k: [None] * counts.get(k, 0)
                for k in ("new", "returning", "removed", "changed", "persistent")}
        send_daily_briefing(briefing, diff, path, report_date)
    except Exception as e:  # never fail to send just because parsing drifted
        print(f"(Could not rebuild the full briefing body: {e}; sending file with a short note.)")
        _send(EMAIL_TO, f"Daily Queue Briefing — {report_date}",
              f"Daily queue report attached: {path.name}", attachment=path)

    print(f"Sent {path.name} to {EMAIL_TO}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
