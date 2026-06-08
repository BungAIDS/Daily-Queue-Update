"""Persist a single day's run artifacts so the pipeline can run in stages.

The daily run can be split into three resumable stages (see main.py):
  1. --no-ai     scrape + diff + Excel  (writes snapshot, diff, Excel)
  2. --ai-only   add the AI briefing      (reads snapshot + diff, writes briefing)
  3. --mail-only email it                 (reads briefing + diff + Excel path)

The scraped jobs are already persisted as the day's snapshot (compare.py). This
module persists the other hand-off artifacts — the diff, the AI briefing, and
the path of the Excel that was written — keyed by date under SNAPSHOT_DIR, so a
botched 5 AM run can be recovered one stage at a time without re-scraping.
"""
from __future__ import annotations

import json
from datetime import date
from pathlib import Path
from typing import Any, Dict

from config import SNAPSHOT_DIR


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
