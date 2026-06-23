"""Engineer roster + name matching.

Tags each order with the engineer(s) named in its **Assigned To**, **Checker**,
or **Note** fields, so every order an engineer ever touched can be found by the
new "Engineer" column (on the Live Queue, the daily report, and Order History)
and by the `engineers` list stored on each order in live_master.json.

Matching is case-insensitive and *token-bounded*: an alias only matches when it
is not butted up against another letter/digit, so bare or dotted initials
("JD", "J.D.") and last names match cleanly without firing inside a longer word
(e.g. "JD" will not match inside "AJDx"). Per the project decision, initials are
matched in every scanned field, including the free-text Note.

Association is **cumulative**: once an engineer is detected on an order they stay
attached forever (see `merge`), even if the order is later reassigned — so a
look-up of "every order associated with X" stays complete.

To add or change people, edit `ROSTER` below: map each canonical display name to
the spellings/initials it may appear as on the board.
"""
from __future__ import annotations

import re
from typing import Any, Dict, Iterable, List, Pattern, Tuple

# Canonical engineer name -> every spelling/initial it may appear as in the
# Assigned To / Checker / Note fields. The canonical name is always matched too,
# so it need not be repeated in its own alias list. EDIT THIS to add coworkers.
ROSTER: Dict[str, List[str]] = {
    # "John Doe":   ["John", "JD", "J.D.", "Doe"],
    # "Maria Ruiz": ["Maria", "MR", "Ruiz"],
}

# The order fields scanned for names: the two board assignment columns and the
# Sales-Order status note.
_SCAN_FIELDS: Tuple[str, ...] = ("assigned_to", "checker", "status_note")


def _alias_pattern(alias: str) -> Pattern:
    """A token-bounded, case-insensitive matcher for one alias. We bound with
    'not adjacent to a letter/digit' rather than \\b so dotted initials like
    'J.D.' and bare initials like 'JD' match without firing mid-word."""
    return re.compile(r"(?<![A-Za-z0-9])" + re.escape(alias.strip()) + r"(?![A-Za-z0-9])",
                      re.IGNORECASE)


def _compile(roster: Dict[str, List[str]]) -> List[Tuple[str, List[Pattern]]]:
    out: List[Tuple[str, List[Pattern]]] = []
    for name, aliases in roster.items():
        spellings = [name] + list(aliases or [])
        pats = [_alias_pattern(a) for a in spellings if a and a.strip()]
        if pats:
            out.append((name, pats))
    return out


_COMPILED: List[Tuple[str, List[Pattern]]] = _compile(ROSTER)


def reload_roster() -> None:
    """Recompile after ROSTER is edited at runtime (used by the tests)."""
    global _COMPILED
    _COMPILED = _compile(ROSTER)


def detect(job: Dict[str, Any]) -> List[str]:
    """The canonical engineer names mentioned in an order's Assigned To / Checker
    / Note fields — sorted and de-duplicated. An empty roster yields []."""
    text = " \n ".join(str(job.get(f) or "") for f in _SCAN_FIELDS)
    if not text.strip() or not _COMPILED:
        return []
    return sorted(name for name, pats in _COMPILED if any(p.search(text) for p in pats))


def merge(existing: Iterable[str], detected: Iterable[str]) -> List[str]:
    """Union of the engineers already attached to an order with those detected
    this poll (sorted) — so an association, once made, is never dropped."""
    return sorted({*(existing or ()), *(detected or ())})


def cell_text(job: Dict[str, Any]) -> str:
    """The Engineer cell text for an order, comma-joined. Uses the stored
    (accumulated) `engineers` list when present — the live document / Order
    History path — and otherwise detects on the fly, so the daily report (whose
    snapshot jobs carry no stored list) still shows the engineers."""
    return ", ".join(job.get("engineers") or detect(job))


def backfill(master: Dict[str, Any]) -> int:
    """Tag every order already in the master log from its stored Assigned To /
    Checker / Note — so historical and already-departed orders are searchable by
    engineer too, not just ones re-seen after this feature shipped. Cumulative
    (union with anything already attached) and idempotent. Returns the number of
    orders whose engineer list changed (e.g. after a ROSTER edit)."""
    changed = 0
    for entry in (master.get("orders") or {}).values():
        job = entry.get("job")
        if not isinstance(job, dict):
            continue
        merged = merge(job.get("engineers"), detect(job))
        if merged != (job.get("engineers") or []):
            job["engineers"] = merged
            changed += 1
    return changed
