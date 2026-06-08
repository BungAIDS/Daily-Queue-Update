"""Email an existing Excel report as-is — no scrape, no diff, no AI, no stages.

    python send_excel.py                          # newest queue_*.xlsx in OUTPUT_DIR
    python send_excel.py "C:\\path\\queue_2026-06-08.xlsx"

Use this when you already have a report you're happy with and just want it in
your inbox. It regenerates nothing: it attaches the file and sends it through
Outlook to EMAIL_TO. The email body reuses the AI overview already written at
the top of the Changes tab, so the message still reads like the daily briefing.
"""
from __future__ import annotations

import re
import sys
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from config import EMAIL_TO, OUTPUT_DIR
from emailer import _send


def _latest() -> Path:
    files = sorted(OUTPUT_DIR.glob("queue_*.xlsx"), key=lambda p: p.stat().st_mtime)
    if not files:
        raise SystemExit(f"No queue_*.xlsx found in {OUTPUT_DIR}")
    return files[-1]


def _ai_overview(path: Path) -> str:
    """Pull the AI overview off the top of the Changes tab (up to the 'New
    orders' section) for the email body. Best-effort: empty string on any issue."""
    try:
        ws = load_workbook(path, read_only=True)["Changes"]
        lines: list[str] = []
        for row in ws.iter_rows(min_row=1, max_row=60, max_col=5, values_only=True):
            first = row[0]
            if isinstance(first, str) and first.startswith("New orders"):
                break
            cells = [str(c) for c in row if c not in (None, "")]
            if cells:
                lines.append(" ".join(cells))
        return "\n".join(lines).strip()
    except Exception:
        return ""


def _report_date(path: Path) -> str:
    """The date in the filename (queue_YYYY-MM-DD.xlsx), else today."""
    m = re.search(r"\d{4}-\d{2}-\d{2}", path.stem)
    return m.group(0) if m else date.today().isoformat()


def main() -> int:
    if not EMAIL_TO:
        print("EMAIL_TO is not set in .env — nothing to send to.")
        return 1

    path = Path(sys.argv[1]) if len(sys.argv) > 1 else _latest()
    if not path.exists():
        print(f"File not found: {path}")
        return 1

    body = _ai_overview(path) or "Daily queue report attached."
    body += f"\n\nReport attached: {path.name}"
    _send(EMAIL_TO, f"Daily Queue Briefing — {_report_date(path)}", body, attachment=path)
    print(f"Sent {path.name} to {EMAIL_TO}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
