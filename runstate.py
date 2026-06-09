"""Persist a single day's run artifacts so the pipeline can run in stages.

The daily run can be split into three resumable stages (see pipeline.py):
  1. scrape.py  scrape + diff + Excel  (writes snapshot, diff, Excel)
  2. brief.py   add the AI briefing     (reads snapshot + diff, writes briefing)
  3. send.py    email it                (reads the most recent report off disk)

The scraped jobs are already persisted as the day's snapshot (compare.py). This
module persists the other hand-off artifacts — the diff, the AI briefing, and
the path of the Excel that was written — keyed by date under SNAPSHOT_DIR, so a
botched 5 AM run can be recovered one stage at a time without re-scraping.
"""
from __future__ import annotations

import json
import logging
import re
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Dict

from config import OUTPUT_DIR, SNAPSHOT_DIR

log = logging.getLogger(__name__)


def _path(kind: str, d: date) -> Path:
    return SNAPSHOT_DIR / f"{kind}_{d.isoformat()}.json"


def save_diff(diff: Dict[str, Any], d: date) -> None:
    # default=str guards any stray non-JSON value; tuples (the per-field change
    # triples) already serialize as arrays and reload fine for the Excel writer.
    _path("diff", d).write_text(json.dumps(diff, indent=2, default=str), encoding="utf-8")


def load_diff(d: date) -> Dict[str, Any] | None:
    p = _path("diff", d)
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else None


def save_briefing(briefing: Dict[str, Any], d: date) -> None:
    _path("briefing", d).write_text(json.dumps(briefing, indent=2), encoding="utf-8")


def load_briefing(d: date) -> Dict[str, Any] | None:
    p = _path("briefing", d)
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else None


def save_excel_path(path: Path, d: date) -> None:
    _path("excel", d).write_text(json.dumps({"path": str(path)}), encoding="utf-8")


def load_excel_path(d: date) -> Path | None:
    p = _path("excel", d)
    if not p.exists():
        return None
    return Path(json.loads(p.read_text(encoding="utf-8"))["path"])


# --- archiving ---------------------------------------------------------------

# Dated per-run files older than this many days are swept into an archive/
# subfolder. They are MOVED, never deleted — the point of this program is a
# complete record of every order ever run, so nothing is thrown away; the
# sweep just keeps the working folders small enough to browse.
ARCHIVE_AFTER_DAYS = 60

_DATED_NAME = re.compile(r"\d{4}-\d{2}-\d{2}")

# (folder, filename patterns) swept by archive_old_runs. history.json — the
# live long-term store — deliberately matches none of these.
_ARCHIVE_SWEEPS = (
    (SNAPSHOT_DIR, ("queue_*.json", "diff_*.json", "briefing_*.json",
                    "excel_*.json", "history_*_start.json")),
    (OUTPUT_DIR, ("queue_*.xlsx",)),
)


def archive_old_runs(today: date, keep_days: int = ARCHIVE_AFTER_DAYS) -> None:
    """Move per-run files older than `keep_days` into <folder>/archive/.

    Runs once per scrape (pipeline.scrape_and_diff). The lookback that the
    daily diff and persistence tracking need is 14 days, so 60 is comfortably
    safe. Never raises: a locked or unreadable file is logged and left for the
    next day's sweep.
    """
    cutoff = today - timedelta(days=keep_days)
    moved = 0
    for folder, patterns in _ARCHIVE_SWEEPS:
        dest = folder / "archive"
        for pattern in patterns:
            try:
                matches = list(folder.glob(pattern))
            except OSError as e:
                log.warning("Archive sweep of %s/%s failed: %s", folder, pattern, e)
                continue
            for p in matches:
                m = _DATED_NAME.search(p.name)
                if not m:
                    continue
                try:
                    d = date.fromisoformat(m.group(0))
                except ValueError:
                    continue
                if d >= cutoff:
                    continue
                try:
                    dest.mkdir(parents=True, exist_ok=True)
                    p.replace(dest / p.name)
                    moved += 1
                except OSError as e:
                    log.warning("Could not archive %s (%s); leaving it in place", p.name, e)
    if moved:
        log.info("Archived %d run file(s) older than %d days into archive/ (moved, not deleted)",
                 moved, keep_days)
