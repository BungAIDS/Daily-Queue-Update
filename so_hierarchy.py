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
every contributing line's attributes and review flags, with the printed lines
and their verbatim continuation details kept underneath as sources:

    [IVC]  (3 lines, 6,104.00)                COMPONENT  one thing, not three
      leakage class: LOW LEAKAGE              ATTRIBUTE  merged from any line
      operation: Automatic                    ATTRIBUTE
      Inlet Volume Control, Low Leak, Auto…   SOURCE     the evidence lines
      + Inlet Volume Control Handle Location… SOURCE
      + Inlet, Flanged, Punched (with IVC)    SOURCE

Grouping is deliberately conservative — lines merge ONLY where the extractors
already tied them to the same thing:

  - `used_on` attribute first (all three IVC lines carry used_on=IVC), then
    the `component` attribute (a line that IS the thing, e.g. MOTOR). Never
    on loose tag overlap: two lines tagged FLANGE are usually two flanges.
  - A line with neither stands as its own component, named by its text.
  - When merged lines disagree on an attribute (same key, different values) BOTH
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
KIND_ATTRIBUTE = "ATTRIBUTE"  # one merged, parser-derived attribute of it
KIND_REVIEW = "REVIEW"        # parser review flags + attribute conflicts
KIND_SOURCE = "SOURCE"        # a printed SO line that contributed

# Structural keys that never render as ATTRIBUTE rows: used_on/component build
# the grouping itself, vfd_context only records that the VFD wording belongs to
# the motor, and *_review keys surface through review_flags/REVIEW instead.
_NON_ATTRIBUTE_KEYS = {"used_on", "component", "vfd_context"}

# Keys the extractors themselves build as ", "-joined collections (features,
# scopes, subcategories, ...). Across merged lines these UNION quietly — three
# IVC lines carrying three different ivc_subcategory values is expected, not a
# conflict. Every other key is expected to be single-valued; a disagreement
# there is kept ('A | B') and flagged for review.
_ACCUMULATIVE_SUFFIXES = ("_subcategory", "_feature", "_scope", "_instruction",
                          "_state", "_method")
_NESTED_COMPONENTS = {"ACTUATOR"}
_ACTUATOR_PREFIX = "actuator_"


def _accumulative(key: str) -> bool:
    return str(key).endswith(_ACCUMULATIVE_SUFFIXES)


def parse_price(p: Any) -> float:
    """A stored price string as a float: '3,531.00' -> 3531.0; the no-charge
    marks (STD / INC / NC / N/C / blank) -> 0."""
    s = str(p or "").replace("$", "").replace(",", "").strip()
    return float(s) if re.fullmatch(r"\d+(\.\d+)?", s) else 0.0


def _paid_inquiry_override(key: str, primary: Dict[str, Any], other: Dict[str, Any]) -> bool:
    """Whether a later quoted modification supersedes a no-charge base fact.

    CBC prints the original STD outlet and a paid Inquiry line for its changed
    construction on the same CO.  For single-valued flange type, that paid
    inquiry is the final construction; keeping both values would knowingly
    report the obsolete base value as current.
    """
    if key not in {"flange_type", "flanged", "punched"}:
        return False
    return (parse_price(primary.get("price")) > 0
            and bool(_attrs(primary).get("inquiry_num"))
            and parse_price(other.get("price")) == 0
            and not _attrs(other).get("inquiry_num"))


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


# Fan locations that ARE a component in their own right when a standalone line
# describes one (e.g. "Outlet, Flanged, Punched"). Used only as a fallback, when
# the extractor didn't already tie the line to a component/used_on — so two
# lines describing the SAME location become ONE component (never two), and any
# single-valued attribute they disagree on — e.g. flange_type PUNCHED vs
# UNPUNCHED — is surfaced as a conflict rather than sitting silently in two
# look-alike components.
_LOCATION_COMPONENTS = ("OUTLET", "INLET")


def _derived_component(a: Dict[str, Any]) -> str:
    """A canonical component name inferred from a standalone line's own scope,
    for parts the extractor named only in prose. Kept deliberately narrow (a
    clear single-location flange scope) so it can only ever MERGE genuine
    duplicates, never fuse unrelated lines."""
    scope = str(a.get("flange_scope") or "").strip().upper()
    return scope if scope in _LOCATION_COMPONENTS else ""


_HANDLE_LOCATION_VALUE = re.compile(
    r"^(?:(?P<clock>\d{1,2}:\d{2})(?:\s*\((?P<clock_status>STD|NON-STD)\))?"
    r"|(?P<status>STD|NON-STD))$",
    re.I,
)


def _merged_handle_location(first: Any, second: Any) -> str:
    """Combine a clock row and its separate STD/NON-STD row without conflict."""
    clocks: List[str] = []
    statuses: List[str] = []
    for value in (first, second):
        match = _HANDLE_LOCATION_VALUE.fullmatch(str(value).strip())
        if not match:
            return ""
        clock = match.group("clock")
        status = match.group("clock_status") or match.group("status")
        if clock and clock not in clocks:
            clocks.append(clock)
        if status and status.upper() not in statuses:
            statuses.append(status.upper())
    if len(clocks) > 1 or len(statuses) > 1:
        return ""
    if clocks:
        return f"{clocks[0]} ({statuses[0]})" if statuses else clocks[0]
    return statuses[0] if statuses else ""


def group_key(item: Dict[str, Any]) -> str:
    """The thing this line is evidence of: its `used_on` attribute (an explicit
    'this line belongs to that component' link), else its `component` attribute
    (this line IS that component), else a component derived from its own scope
    (so two lines for the same location merge), else '' — the line stands alone.
    Loose tag overlap is deliberately NOT a merge signal."""
    a = _attrs(item)
    return (str(a.get("used_on") or "").strip().upper()
            or str(a.get("component") or "").strip().upper()
            or _derived_component(a))


def _used_on_parents(value: Any) -> List[str]:
    parents: List[str] = []
    for part in re.split(r"\s*(?:,|\|)\s*", str(value or "")):
        parent = part.strip().upper()
        if parent and parent not in parents:
            parents.append(parent)
    return parents


def _hierarchy_views(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Expand actuator ownership without changing the captured item records.

    A damper line can contain actuator fields (stored with ``actuator_`` keys),
    and a separately priced actuator line can explicitly name one or several
    dampers. Both become the same nested [ACTUATOR] child. A shared buyout line
    is visible/priced once, while its facts are linked to every named parent.
    """
    views: List[Dict[str, Any]] = []

    def add(item: Dict[str, Any], item_no: int, attrs: Dict[str, Any],
            parent: str = "", price: Any = None, hide_source: bool = False) -> None:
        view = dict(item)
        view["attributes"] = attrs
        view["_item_no"] = item_no
        view["_hierarchy_parent"] = parent
        view["_hide_source"] = hide_source
        if price is not None:
            view["price"] = price
        views.append(view)

    for item_no, item in enumerate(items or [], start=1):
        attrs = dict(_attrs(item))
        component = str(attrs.get("component") or "").strip().upper()
        parents = _used_on_parents(attrs.get("used_on"))
        if component in _NESTED_COMPONENTS and parents:
            for parent_index, parent in enumerate(parents):
                child_attrs = dict(attrs)
                child_attrs["used_on"] = parent
                add(item, item_no, child_attrs, parent=parent,
                    price=item.get("price") if parent_index == 0 else "",
                    hide_source=parent_index > 0)
            continue

        child_attrs = {
            key[len(_ACTUATOR_PREFIX):]: value
            for key, value in attrs.items()
            if key.startswith(_ACTUATOR_PREFIX) and value
        }
        parent_attrs = {
            key: value for key, value in attrs.items()
            if not key.startswith(_ACTUATOR_PREFIX)
        }
        add(item, item_no, parent_attrs)
        if child_attrs and component:
            child_attrs["component"] = "ACTUATOR"
            child_attrs["used_on"] = component
            child_item = dict(item)
            # The source row's review belongs to the parent fact. Do not repeat
            # it on the synthetic child created from actuator_* attributes.
            child_item["review_flags"] = []
            add(child_item, item_no, child_attrs, parent=component, price="",
                hide_source=True)
    return views


def components(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """One record per real thing, in SO print order (a merged component sits
    where its first evidence line appeared):

        {name, keyed, attributes {k: v}, review [{text, item_no}],
         sources [{item_no, text, details, price, primary}], price (summed float)}

    `keyed` says the name came from a used_on/component link (rendered
    [BRACKETED]); an un-keyed record is a lone line named by its own text."""
    views = _hierarchy_views(items or [])
    groups: Dict[str, List[int]] = {}
    for i, it in enumerate(views):
        parent = str(it.get("_hierarchy_parent") or "").strip().upper()
        if parent:
            component = str(_attrs(it).get("component") or "?").strip().upper()
            key = f"\x01{parent}\x1f{component}"
        else:
            key = group_key(it) or f"\x00{i}"
        groups.setdefault(key, []).append(i)

    records: List[Dict[str, Any]] = []
    for order, (key, idxs) in enumerate(groups.items()):
        nested = key.startswith("\x01")
        parent = ""
        if nested:
            parent, name = key[1:].split("\x1f", 1)
            keyed = True
        else:
            keyed = not key.startswith("\x00")
            name = key if keyed else (line_text(views[idxs[0]]) or "?")
        primary = max(idxs, key=lambda i: parse_price(views[i].get("price")))
        attributes: Dict[str, Any] = {}
        conflict_keys: List[str] = []
        review: List[Dict[str, Any]] = []
        sources: List[Dict[str, Any]] = []
        for i in [primary] + [i for i in idxs if i != primary]:
            it = views[i]
            item_no = int(it.get("_item_no") or 0)
            if not it.get("_hide_source"):
                sources.append({
                    "item_no": item_no,
                    "text": line_text(it),
                    "details": [str(d) for d in it.get("details") or []],
                    "price": it.get("price") or "",
                    "primary": i == primary,
                })
            for k in sorted(_attrs(it)):
                v = _attrs(it)[k]
                if not v or k in _NON_ATTRIBUTE_KEYS or k.endswith("_review"):
                    continue
                if k == "note":
                    notes = attributes.setdefault(k, [])
                    if str(v) not in notes:
                        notes.append(str(v))
                    continue
                have = attributes.get(k)
                if have is None:
                    attributes[k] = v
                    continue
                if k == "handle_location":
                    merged = _merged_handle_location(have, v)
                    if merged:
                        attributes[k] = merged
                        continue
                if nested and name == "ACTUATOR" and k == "model":
                    have_model = re.sub(r"[^A-Z0-9]", "", str(have).upper())
                    new_model = re.sub(r"[^A-Z0-9]", "", str(v).upper())
                    # A damper line often says only RPED; its separately priced
                    # actuator line supplies the exact RPED150/RPED200 model.
                    if (have_model == "RPED" and new_model.startswith("RPED")
                            and new_model != "RPED"):
                        attributes[k] = v
                        continue
                    if (new_model == "RPED" and have_model.startswith("RPED")
                            and have_model != "RPED"):
                        continue
                if k == "testing_status":
                    equivalents = {str(have).upper(), str(v).upper()}
                    if equivalents <= {"N/A", "NOT AVAILABLE"}:
                        attributes[k] = "NOT AVAILABLE"
                        continue
                seen = [x.strip() for x in re.split(r"[|,]", str(have))]
                for part in (str(v).split(", ") if _accumulative(k) else [str(v)]):
                    if part.strip() in seen:
                        continue
                    if i != primary and _paid_inquiry_override(k, views[primary], it):
                        continue
                    if _accumulative(k):
                        attributes[k] = f"{attributes[k]}, {part}"
                    else:
                        attributes[k] = f"{attributes[k]} | {part}"
                        if k not in conflict_keys:
                            conflict_keys.append(k)
                    seen.append(part.strip())
            for flag in it.get("review_flags") or []:
                entry = {"text": str(flag), "item_no": item_no}
                if entry not in review:
                    review.append(entry)
        conflicts = [
            {"text": f"CONFLICTING {k}: {attributes[k]}", "item_no": ""}
            for k in conflict_keys
        ]
        records.append({
            "name": name,
            "keyed": keyed,
            "attributes": attributes,
            "review": review + conflicts,
            "sources": sources,
            "price": sum(parse_price(views[i].get("price")) for i in idxs),
            "children": [],
            "_parent": parent,
            "_order": order,
        })

    top: Dict[str, Dict[str, Any]] = {}
    output: List[Dict[str, Any]] = []
    for record in records:
        if record["_parent"]:
            continue
        top[str(record["name"]).upper()] = record
        output.append(record)
    for child in (record for record in records if record["_parent"]):
        parent_name = str(child["_parent"]).upper()
        parent = top.get(parent_name)
        if parent is None:
            parent = {
                "name": parent_name,
                "keyed": True,
                "attributes": {},
                "review": [],
                "sources": [],
                "price": 0.0,
                "children": [],
                "_parent": "",
                "_order": child["_order"],
            }
            top[parent_name] = parent
            output.append(parent)
        parent["_order"] = min(parent["_order"], child["_order"])
        parent["children"].append(child)

    output.sort(key=lambda record: record["_order"])
    for record in output:
        record["children"].sort(key=lambda child: child["_order"])
        record.pop("_parent", None)
        record.pop("_order", None)
        for child in record["children"]:
            child.pop("_parent", None)
            child.pop("_order", None)
    return output


def _attribute_text(key: str, value: Any) -> str:
    return f"{str(key).replace('_', ' ')}: {value}"


def tree_rows(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """The job's knowledge as flat render rows, each
    {depth, kind, text, price, item_no}. Per component: its COMPONENT row, one
    ATTRIBUTE row per merged attribute, individual REVIEW rows when flagged,
    then the SOURCE lines (skipped for a lone line — the
    component row IS the line, and carries its item # directly)."""
    rows: List[Dict[str, Any]] = []

    def append_component(component: Dict[str, Any], depth: int) -> None:
        sources = component.get("sources") or []
        multi = len(sources) > 1
        name = f"[{component['name']}]" if component["keyed"] else component["name"]
        if multi:
            price = f"{component['price']:,.2f}" if component["price"] else ""
            item_no: Any = ""
        elif sources:
            price = sources[0]["price"]
            item_no = sources[0]["item_no"]
        else:
            price, item_no = "", ""
        rows.append({"depth": depth, "kind": KIND_COMPONENT, "text": name,
                     "price": price, "item_no": item_no})
        for key, value in component["attributes"].items():
            values = value if isinstance(value, list) else [value]
            for part in values:
                rows.append({"depth": depth + 1, "kind": KIND_ATTRIBUTE,
                             "text": _attribute_text(key, part),
                             "price": "", "item_no": ""})
        for review in component["review"]:
            rows.append({"depth": depth + 1, "kind": KIND_REVIEW,
                         "text": review["text"], "price": "",
                         "item_no": review["item_no"]})
        for child in component.get("children") or []:
            append_component(child, depth + 1)
        if multi:
            for source in sources:
                rows.append({
                    "depth": depth + 1,
                    "kind": KIND_SOURCE,
                    "text": ("" if source["primary"] else "+ ") + source["text"],
                    "price": source["price"],
                    "item_no": source["item_no"],
                })

    for component in components(items):
        append_component(component, 0)
    return rows


def indent_text(row: Dict[str, Any], unit: str = "    ") -> str:
    """The row's text with its depth rendered as leading indentation — what the
    Excel Hierarchy column and the CLI print."""
    return unit * int(row.get("depth") or 0) + str(row.get("text") or "")
