"""What we KNOW about a job, built from its captured Sales-Order line items.

The store keeps what the SO *prints*: flat priced lines, each with its unpriced
detail lines and derived attributes/tags (line_items.extract_items). But a
printed line is EVIDENCE, not necessarily a thing: CBC sells one physical
option across several charges — 421966's inlet volume control is three lines
("Inlet Volume Control, Low Leak, Automatic", "...Handle Location,
Non-standard", "Inlet, Flanged, Punched (with IVC)") — and three lines must
not become three IVCs. One is sufficient to know an IVC exists; every other
line may (or may not) add information about it.

`components()` builds that knowledge: ONE record per real thing, accumulating
every contributing line's facts, details and review flags, with the lines kept
underneath as its sources:

    [IVC]  (3 lines, 6,104.00)                COMPONENT  one thing, not three
      leakage class: LOW LEAKAGE              FACT       merged from any line
      operation: Automatic                    FACT
      · Actuator Manufacturer: By Others      DETAIL     stored sub-line
      Inlet Volume Control, Low Leak, Auto…   SOURCE     the evidence lines
      + Inlet Volume Control Handle Location… SOURCE
      + Inlet, Flanged, Punched (with IVC)    SOURCE

Grouping is deliberately conservative — lines merge ONLY where the extractors
already tied them to the same thing:

  - `used_on` attribute first (all three IVC lines carry used_on=IVC), then
    the `component` attribute (a line that IS the thing, e.g. MOTOR). Never
    on loose tag overlap: two lines tagged FLANGE are usually two flanges.
  - A line with neither stands as its own component, named by its text.
  - When merged lines disagree on a fact (same key, different values) BOTH
    values are kept ('A | B') and the conflict lands in the component's
    review list — surfaced, never silently resolved.
  - The priciest line is the component's own line (shown first, un-prefixed);
    other sources keep capture order with a leading '+'.

Everything is derived, never stored: the tree is a pure re-arrangement of the
capture, so a wrong tree means a wrong capture/attribute — exactly what it
exists to surface — and grouping rules can be tuned and re-run any time.
Import-light (no pdf/COM); used by the Sales Order tab's Hierarchy block
(live_sheets) and the so_tree.py CLI.
"""
from __future__ import annotations

import re
from typing import Any, Dict, List

from line_items import split_lead, split_price_tail, split_type_tail

# Row kinds — where each tree row comes from.
KIND_COMPONENT = "COMPONENT"  # one real thing we know the job has
KIND_FACT = "FACT"            # one merged attribute of it (derived)
KIND_DETAIL = "DETAIL"        # a stored unpriced sub-line of a source
KIND_REVIEW = "REVIEW"        # parser review flags + fact conflicts
KIND_SOURCE = "SOURCE"        # a printed SO line that contributed

# Attribute keys that never render as FACT rows: used_on names the component
# itself, and *_review keys surface through review_flags/REVIEW instead.
_NON_FACT_KEYS = {"used_on"}

# Keys the extractors themselves build as ", "-joined collections (features,
# scopes, subcategories, ...). Across merged lines these UNION quietly — three
# IVC lines carrying three different ivc_subcategory values is expected, not a
# conflict. Every other key is expected to be single-valued; a disagreement
# there is kept ('A | B') and flagged for review.
_ACCUMULATIVE_SUFFIXES = ("_subcategory", "_feature", "_scope", "_instruction",
                          "_state", "_method")


def _accumulative(key: str) -> bool:
    return str(key).endswith(_ACCUMULATIVE_SUFFIXES)


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


def _attrs(item: Dict[str, Any]) -> Dict[str, Any]:
    a = item.get("attributes") or {}
    return a if isinstance(a, dict) else {}


def group_key(item: Dict[str, Any]) -> str:
    """The thing this line is evidence of: its `used_on` attribute (an explicit
    'this line belongs to that component' link), else its `component` attribute
    (this line IS that component), else '' — the line stands alone. Loose tag
    overlap is deliberately NOT a merge signal."""
    a = _attrs(item)
    return (str(a.get("used_on") or "").strip().upper()
            or str(a.get("component") or "").strip().upper())


def components(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """One record per real thing, in SO print order (a merged component sits
    where its first evidence line appeared):

        {name, keyed, facts {k: v}, details [..], review [..],
         sources [{item_no, text, price, primary}], price (summed float)}

    `keyed` says the name came from a used_on/component link (rendered
    [BRACKETED]); an un-keyed record is a lone line named by its own text."""
    groups: Dict[str, List[int]] = {}
    for i, it in enumerate(items or []):
        groups.setdefault(group_key(it) or f"\x00{i}", []).append(i)

    out: List[Dict[str, Any]] = []
    for key, idxs in groups.items():           # insertion order = SO print order
        keyed = not key.startswith("\x00")
        primary = max(idxs, key=lambda i: parse_price(items[i].get("price")))
        facts: Dict[str, Any] = {}
        conflict_keys: List[str] = []
        details: List[str] = []
        review: List[str] = []
        sources: List[Dict[str, Any]] = []
        for i in [primary] + [i for i in idxs if i != primary]:
            it = items[i]
            sources.append({"item_no": i + 1, "text": line_text(it),
                            "price": it.get("price") or "", "primary": i == primary})
            for k in sorted(_attrs(it)):
                v = _attrs(it)[k]
                if not v or k in _NON_FACT_KEYS or k.endswith("_review"):
                    continue
                have = facts.get(k)
                if have is None:
                    facts[k] = v
                    continue
                seen = [x.strip() for x in re.split(r"[|,]", str(have))]
                for part in (str(v).split(", ") if _accumulative(k) else [str(v)]):
                    if part.strip() in seen:
                        continue
                    if _accumulative(k):        # collection key -> union quietly
                        facts[k] = f"{facts[k]}, {part}"
                    else:                       # single-valued key -> keep both, flag
                        facts[k] = f"{facts[k]} | {part}"
                        if k not in conflict_keys:
                            conflict_keys.append(k)
                    seen.append(part.strip())
            for d in it.get("details") or []:
                if str(d) not in details:
                    details.append(str(d))
            for f in it.get("review_flags") or []:
                if str(f) not in review:
                    review.append(str(f))
        conflicts = [f"CONFLICTING {k}: {facts[k]}" for k in conflict_keys]
        total = sum(parse_price(items[i].get("price")) for i in idxs)
        out.append({"name": key if keyed else (line_text(items[idxs[0]]) or "?"),
                    "keyed": keyed,
                    "facts": facts, "details": details,
                    "review": review + conflicts, "sources": sources,
                    "price": total})
    return out


def _fact_text(key: str, value: Any) -> str:
    return f"{str(key).replace('_', ' ')}: {value}"


def tree_rows(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """The job's knowledge as flat render rows, each
    {depth, kind, text, price, item_no}. Per component: its COMPONENT row, one
    FACT row per merged attribute, the stored DETAIL sub-lines, a REVIEW row
    when flagged, then the SOURCE lines (skipped for a lone line — the
    component row IS the line, and carries its item # directly)."""
    rows: List[Dict[str, Any]] = []
    for c in components(items):
        multi = len(c["sources"]) > 1
        name = f"[{c['name']}]" if c["keyed"] else c["name"]
        if multi:
            name += f" — {len(c['sources'])} lines"
        # A lone line keeps its printed price mark (STD / NC stay visible, like
        # the flat table); a merged component shows the sum of its charges.
        price = (f"{c['price']:,.2f}" if c["price"] else "") if multi \
            else c["sources"][0]["price"]
        rows.append({"depth": 0, "kind": KIND_COMPONENT, "text": name,
                     "price": price,
                     "item_no": "" if multi else c["sources"][0]["item_no"]})
        for k, v in c["facts"].items():
            rows.append({"depth": 1, "kind": KIND_FACT, "text": _fact_text(k, v),
                         "price": "", "item_no": ""})
        for d in c["details"]:
            rows.append({"depth": 1, "kind": KIND_DETAIL, "text": "· " + d,
                         "price": "", "item_no": ""})
        if c["review"]:
            rows.append({"depth": 1, "kind": KIND_REVIEW,
                         "text": "; ".join(c["review"]), "price": "", "item_no": ""})
        if multi:
            for s in c["sources"]:
                rows.append({"depth": 1, "kind": KIND_SOURCE,
                             "text": ("" if s["primary"] else "+ ") + s["text"],
                             "price": s["price"], "item_no": s["item_no"]})
    return rows


def indent_text(row: Dict[str, Any], unit: str = "    ") -> str:
    """The row's text with its depth rendered as leading indentation — what the
    Excel Hierarchy column and the CLI print."""
    return unit * int(row.get("depth") or 0) + str(row.get("text") or "")
