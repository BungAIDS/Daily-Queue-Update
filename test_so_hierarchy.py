"""Tests for the component-hierarchy rollup over captured Sales-Order line
items (so_hierarchy.py).

No pytest — run directly:

    python test_so_hierarchy.py

Fitted against the real 421966 shape: one physical option (the IVC) sold as
three separate printed charges that share a used_on attribute.
"""
from __future__ import annotations

import sys

import so_hierarchy as soh


def _item(raw, price="", used_on=None, details=None, tags=None, flags=None,
          attrs=None, norm=""):
    it = {"raw": raw, "norm": norm or raw.upper(), "qty": "", "price": price,
          "ptype": "L", "section": "", "details": details or [], "tags": tags or []}
    a = dict(attrs or {})
    if used_on:
        a["used_on"] = used_on
    if a:
        it["attributes"] = a
    if flags:
        it["review_flags"] = flags
    return it


def _ivc_items():
    main = _item("Inlet Volume Control, Low Leak, Automatic L 3,531.00",
                 "3,531.00", used_on="IVC",
                 attrs={"operation": "Automatic", "leakage_class": "LOW LEAKAGE"},
                 details=["Actuator Manufacturer: By Others"])
    handle = _item("Inlet Volume Control Handle Location, Non-standard L 1,014.00",
                   "1,014.00", used_on="IVC",
                   details=["IVC handle location for Discharge"])
    flange = _item("Inlet, Flanged, Punched (with IVC) L 1,559.00",
                   "1,559.00", used_on="IVC")
    return main, handle, flange


def test_family_rollup_anchor_and_satellites():
    main, handle, flange = _ivc_items()
    rows = soh.tree_rows([handle, main, flange])   # capture order: handle first
    fam = rows[0]
    assert fam["kind"] == soh.KIND_FAMILY
    assert fam["text"].startswith("[IVC]") and "3 lines" in fam["text"]
    assert fam["price"] == "6,104.00" and fam["item_no"] == ""
    lines = [r for r in rows if r["kind"] == soh.KIND_LINE]
    # The priciest member is the component's own line: first, un-prefixed,
    # nested under the family, and cross-referenced to its flat-table row.
    assert lines[0]["text"].startswith("Inlet Volume Control, Low Leak")
    assert lines[0]["depth"] == 1 and lines[0]["item_no"] == 2
    # The other charges stay visible but demoted with '+', in capture order.
    assert lines[1]["text"].startswith("+ Inlet Volume Control Handle Location")
    assert lines[1]["item_no"] == 1
    assert lines[2]["text"].startswith("+ Inlet, Flanged, Punched")
    assert lines[2]["item_no"] == 3


def test_price_and_type_tails_are_stripped_from_line_text():
    main, _h, _f = _ivc_items()
    assert soh.line_text(main) == "Inlet Volume Control, Low Leak, Automatic"


def test_single_member_family_stays_top_level():
    # "unless it's the only line item referring to it" — no family of one.
    _m, _h, flange = _ivc_items()
    rows = soh.tree_rows([flange])
    assert rows[0]["kind"] == soh.KIND_LINE and rows[0]["depth"] == 0
    assert not rows[0]["text"].startswith("+")
    assert all(r["kind"] != soh.KIND_FAMILY for r in rows)


def test_facts_details_and_review_rows():
    it = _item("Motor C 3,194.37", "3,194.37",
               attrs={"vendor": "Baldor", "used_on_review": "INCONCLUSIVE"},
               details=["75 HP, 1800 RPM"], flags=["UNTAGGED"])
    rows = soh.tree_rows([it])
    kinds = [r["kind"] for r in rows]
    assert kinds == [soh.KIND_LINE, soh.KIND_FACTS, soh.KIND_DETAIL, soh.KIND_REVIEW]
    # FACTS carries the attributes but never *_review keys (REVIEW's job)...
    assert "vendor=Baldor" in rows[1]["text"] and "used_on_review" not in rows[1]["text"]
    assert rows[2]["text"] == "· 75 HP, 1800 RPM"
    assert rows[3]["text"] == "UNTAGGED"
    # ...and every sub-row points back at the same flat-table item #.
    assert all(r["item_no"] == 1 for r in rows)
    assert all(r["depth"] == 1 for r in rows[1:])


def test_ungrouped_items_keep_capture_order():
    a = _item("Base Fan L 16,649.00", "16,649.00")
    b = _item("Lifting Lugs L STD", "STD")
    rows = soh.tree_rows([a, b])
    lines = [r for r in rows if r["kind"] == soh.KIND_LINE]
    assert [r["item_no"] for r in lines] == [1, 2]
    assert all(r["depth"] == 0 for r in lines)


def test_family_appears_where_first_member_did():
    main, handle, _f = _ivc_items()
    solo = _item("Base Fan L 16,649.00", "16,649.00")
    rows = soh.tree_rows([handle, solo, main])
    # The [IVC] family sits where its first member (handle) appeared: before
    # Base Fan.
    assert rows[0]["kind"] == soh.KIND_FAMILY and "[IVC]" in rows[0]["text"]
    base = next(r for r in rows if "Base Fan" in r["text"])
    assert base["depth"] == 0 and base["item_no"] == 2


def test_parse_price_and_no_charge_marks():
    assert soh.parse_price("3,531.00") == 3531.0
    assert soh.parse_price("$1,014.00") == 1014.0
    for mark in ("STD", "INC", "NC", "N/C", "", None):
        assert soh.parse_price(mark) == 0.0


def test_indent_text_follows_depth():
    assert soh.indent_text({"depth": 0, "text": "x"}) == "x"
    assert soh.indent_text({"depth": 2, "text": "x"}) == "        x"


def main() -> int:
    passed = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
            print(f"  ok  {name}")
            passed += 1
    print(f"\n{passed} tests passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
