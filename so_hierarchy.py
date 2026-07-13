"""Component hierarchy over one order's captured Sales-Order line items.

The store keeps what the SO *prints*: flat priced lines, each with its unpriced
detail lines and derived attributes/tags (line_items.extract_items). But CBC
sells one physical option across several printed lines — 421966's inlet volume
control is three separate charges ("Inlet Volume Control, Low Leak, Automatic",
"...Handle Location, Non-standard", "Inlet, Flanged, Punched (with IVC)") — so
a flat view hides that they all describe ONE component.

This module rolls those lines up into the component view a human expects:

    [IVC] — 3 lines                                    FAMILY  (derived)
      Inlet Volume Control, Low Leak, Automatic        LINE    the component
        leakage_class=LOW LEAKAGE; operation=Automatic FACTS   (derived attrs)
        · Actuator Manufacturer: By Others             DETAIL  (stored sub-line)
      + Inlet Volume Control Handle Location, Non-std  LINE    satellite charge
      + Inlet, Flanged, Punched (with IVC)             LINE    satellite charge

The grouping key is the `used_on` attribute line_items' component extractors
already stamp on every line that modifies a component (all three IVC lines
carry used_on=IVC). Rules:

  - Two or more lines sharing a used_on -> a FAMILY node. Its main LINE is the
    priciest member (the component itself); the rest are satellites, kept
    visible but demoted with a leading '+' — a satellite that looks redundant
    ("with IVC") is *shown as* subordinate rather than standing top-level.
  - A single line with a used_on, or a line with none, stays top-level — a
    family of one would just repeat the line.
  - Under every LINE: its FACTS (the structured attributes, minus *_review
    keys), each stored DETAIL verbatim, and a REVIEW row for parser flags.
  - Every row carries the item's 1-based position in the stored item list, the
    same '#' the flat capture table shows — so tree and flat cross-reference.

Nothing here re-parses text: the tree is a pure re-arrangement of what the
store already holds, so a wrong tree means a wrong capture/attribute — exactly
what it exists to surface. Import-light (no pdf/COM); used by the Sales Order
tab's Hierarchy block (live_sheets) and the so_tree.py CLI.
"""
from __future__ import annotations

import re
from typing import Any, Dict, List

from line_items import split_lead, split_price_tail, split_type_tail

# Row kinds — WHERE in the store (or from what derivation) a tree row comes.
KIND_FAMILY = "FAMILY"    # derived group node (shared used_on)
KIND_LINE = "LINE"        # a stored item (a printed SO line)
KIND_FACTS = "FACTS"      # the item's derived attributes
KIND_DETAIL = "DETAIL"    # a stored unpriced sub-line of the item
KIND_REVIEW = "REVIEW"    # parser review flags on the item


def parse_price(p: Any) -> float:
    """A stored price string as a float: '3,531.00' -> 3531.0; the no-charge
    marks (STD / INC / NC / N/C / blank) -> 0."""
    s = str(p or "").replace("$", "").replace(",", "").strip()
    return float(s) if re.fullmatch(r"\d+(\.\d+)?", s) else 0.0


def line_text(item: Dict[str, Any]) -> str:
    """The item as a readable description: the raw line with the leading
    item-number/qty and the trailing price columns / L-C-N letter stripped
    (they have their own columns), falling back to the normalized form."""
    body, _qty = split_lead(item.get("raw") or "")
    body, _price = split_price_tail(body)
    body, _ptype, _mark = split_type_tail(body)
    return body or item.get("norm", "")


def family_key(item: Dict[str, Any]) -> str:
    """The component this line belongs to: its `used_on` attribute ('' when the
    extractors didn't tie it to one)."""
    attrs = item.get("attributes") or {}
    return str(attrs.get("used_on") or "").strip().upper() if isinstance(attrs, dict) else ""


def _facts(item: Dict[str, Any]) -> str:
    """The item's structured attributes as 'k=v; ...'. *_review keys are left
    out — they surface on the REVIEW row instead."""
    attrs = item.get("attributes") or {}
    if not isinstance(attrs, dict):
        return ""
    return "; ".join(f"{k}={attrs[k]}"
                     for k in sorted(attrs) if attrs[k] and not k.endswith("_review"))


def _line_rows(item: Dict[str, Any], no: int, depth: int, satellite: bool) -> List[Dict[str, Any]]:
    prefix = "+ " if satellite else ""
    rows = [{"depth": depth, "kind": KIND_LINE, "text": prefix + line_text(item),
             "price": item.get("price") or "", "item_no": no}]
    facts = _facts(item)
    if facts:
        rows.append({"depth": depth + 1, "kind": KIND_FACTS, "text": facts,
                     "price": "", "item_no": no})
    for d in item.get("details") or []:
        rows.append({"depth": depth + 1, "kind": KIND_DETAIL, "text": "· " + str(d),
                     "price": "", "item_no": no})
    flags = item.get("review_flags") or []
    if flags:
        rows.append({"depth": depth + 1, "kind": KIND_REVIEW,
                     "text": "; ".join(str(f) for f in flags), "price": "", "item_no": no})
    return rows


def tree_rows(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """The order's hierarchy as flat render rows, each
    {depth, kind, text, price, item_no} — families where they first appear in
    SO print order, everything else in place."""
    groups: Dict[str, List[int]] = {}
    for i, it in enumerate(items or []):
        # \x00<i> keeps ungrouped lines unique AND un-groupable.
        groups.setdefault(family_key(it) or f"\x00{i}", []).append(i)

    rows: List[Dict[str, Any]] = []
    for key, idxs in groups.items():           # insertion order = SO print order
        if key.startswith("\x00") or len(idxs) == 1:
            for i in idxs:                     # top-level: no family of one
                rows += _line_rows(items[i], i + 1, depth=0, satellite=False)
            continue
        total = sum(parse_price(items[i].get("price")) for i in idxs)
        anchor = max(idxs, key=lambda i: parse_price(items[i].get("price")))
        rows.append({"depth": 0, "kind": KIND_FAMILY,
                     "text": f"[{key}] — {len(idxs)} lines",
                     "price": f"{total:,.2f}" if total else "", "item_no": ""})
        for i in [anchor] + [i for i in idxs if i != anchor]:
            rows += _line_rows(items[i], i + 1, depth=1, satellite=(i != anchor))
    return rows


def indent_text(row: Dict[str, Any], unit: str = "    ") -> str:
    """The row's text with its depth rendered as leading indentation — what the
    Excel Hierarchy column and the CLI print."""
    return unit * int(row.get("depth") or 0) + str(row.get("text") or "")
