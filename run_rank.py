"""Rank an order's multiple quote runs by how CURRENT they are.

~1 in 5 orders carries more than one run file (a .txt and a .pdf of the same
run, or genuine revisions: `Qt Run CO#1.txt`, `QT RUN REV A.pdf`, an old dated
copy). Wherever one run must represent the order — master.json, the daily
report's Quote Run Details — picking `runs[0]` (alphabetical) can silently
serve stale engineering data (a 2021 base run sorting ahead of its CO#1).

Ranking, most-current first:
  1. highest CO# in the file name          (a change-order rerun supersedes)
  2. highest REV letter/number in the name (REV B > REV A > no rev)
  3. newest file modified time             (recency when names don't say)
  4. most fields extracted                 (richer parse breaks remaining ties)
Stable: runs that tie keep their original order. Pure stdlib, import-light.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List, Sequence, Tuple

# "CO#1" / "CO # 2" in a file name. The '#' is required so part numbers and
# words like CONSTRUCTION can't false-match.
_CO_RE = re.compile(r"\bCO\s*#\s*(\d+)", re.I)
# "REV A" / "REV-B" / "REV 2" in a file name.
_REV_RE = re.compile(r"\bREV\s*[-. ]?\s*([A-Z]\b|\d+)", re.I)


def revision_key(name: str) -> Tuple[int, int]:
    """(co, rev) revision signals in a run file name; (0, 0) = base run."""
    co = max((int(m) for m in _CO_RE.findall(name or "")), default=0)
    rev = 0
    for m in _REV_RE.findall(name or ""):
        val = int(m) if m.isdigit() else ord(m.upper()) - ord("A") + 1
        rev = max(rev, val)
    return co, rev


def rank_runs(runs: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """The store's run dicts, most-current first (stable for full ties)."""
    def key(run: Dict[str, Any]) -> Tuple[int, int, float, int]:
        co, rev = revision_key(run.get("file", ""))
        return (co, rev, float(run.get("mtime") or 0), len(run.get("fields") or {}))
    return sorted(runs, key=key, reverse=True)


def rank_paths(paths: Sequence[Path]) -> List[Path]:
    """Run FILES on disk, most-current first (name signals, then mtime)."""
    def key(p: Path) -> Tuple[int, int, float]:
        co, rev = revision_key(p.name)
        try:
            mtime = p.stat().st_mtime
        except OSError:
            mtime = 0.0
        return (co, rev, mtime)
    return sorted(paths, key=key, reverse=True)
