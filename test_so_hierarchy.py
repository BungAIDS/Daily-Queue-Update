"""Tests for the job-knowledge rollup over captured Sales-Order line items
(so_hierarchy.py): every printed line ADDS information; only lines the
extractors tied together become one component.

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
                 attrs={"operation": "Automatic", "leakage_class": "LOW LEAKAGE",
                        "ivc_subcategory": "IVC ACTUATOR"},
                 details=["Actuator Manufacturer: By Others"])
    handle = _item("Inlet Volume Control Handle Location, Non-standard L 1,014.00",
                   "1,014.00", used_on="IVC",
                   attrs={"ivc_subcategory": "IVC"},
                   details=["IVC handle location for Discharge"])
    flange = _item("Inlet, Flanged, Punched (with IVC) L 1,559.00",
                   "1,559.00", used_on="IVC",
                   attrs={"ivc_subcategory": "IVC INLET FLANGE"})
    return main, handle, flange


def test_three_lines_one_component():
    main, handle, flange = _ivc_items()
    comps = soh.components([handle, main, flange])
    assert len(comps) == 1                       # one IVC, not three
    c = comps[0]
    assert c["name"] == "IVC" and c["keyed"]
    assert c["price"] == 6104.0
    # Attributes accumulate from every contributing line...
    assert c["attributes"]["operation"] == "Automatic"
    assert c["attributes"]["leakage_class"] == "LOW LEAKAGE"
    # ...collection keys union quietly (three different subcategories is
    # expected, not a conflict).
    assert c["attributes"]["ivc_subcategory"] == "IVC ACTUATOR, IVC, IVC INLET FLANGE"
    assert c["review"] == []
    # Sources: the priciest line is the component's own (primary), the rest in
    # capture order, retaining their verbatim details and flat-table item #s.
    assert [s["item_no"] for s in c["sources"]] == [2, 1, 3]
    assert c["sources"][0]["primary"] and not c["sources"][1]["primary"]
    assert c["sources"][1]["details"] == ["IVC handle location for Discharge"]


def test_tree_rows_component_attributes_sources():
    main, handle, flange = _ivc_items()
    rows = soh.tree_rows([handle, main, flange])
    assert rows[0]["kind"] == soh.KIND_COMPONENT
    assert rows[0]["text"] == "[IVC] — 3 lines"
    assert rows[0]["price"] == "6,104.00" and rows[0]["item_no"] == ""
    attributes = [r["text"] for r in rows if r["kind"] == soh.KIND_ATTRIBUTE]
    assert "operation: Automatic" in attributes  # keys prettified for reading
    srcs = [r for r in rows if r["kind"] == soh.KIND_SOURCE]
    assert srcs[0]["text"].startswith("Inlet Volume Control, Low Leak")
    assert srcs[0]["item_no"] == 2
    assert srcs[1]["text"].startswith("+ ") and srcs[2]["text"].startswith("+ ")
    assert all(r["depth"] == 1 for r in rows[1:])


def test_conflicting_single_valued_attribute_is_kept_and_flagged():
    a = _item("IVC, Manual L 100.00", "100.00", used_on="IVC",
              attrs={"operation": "Manual"})
    b = _item("IVC Actuator L 50.00", "50.00", used_on="IVC",
              attrs={"operation": "Automatic"})
    c = soh.components([a, b])[0]
    assert c["attributes"]["operation"] == "Manual | Automatic"  # both kept...
    assert any("CONFLICTING operation" in r["text"] for r in c["review"])
    rows = soh.tree_rows([a, b])
    assert any(r["kind"] == soh.KIND_REVIEW and "CONFLICTING" in r["text"]
               for r in rows)


def test_lone_line_stays_its_own_component():
    # A single line referring to a component: no group header, no SOURCE rows —
    # the component row IS the line, keeping its item # and printed price mark.
    _m, _h, flange = _ivc_items()
    rows = soh.tree_rows([flange])
    assert rows[0]["kind"] == soh.KIND_COMPONENT
    assert rows[0]["text"] == "[IVC]" and rows[0]["item_no"] == 1
    assert all(r["kind"] != soh.KIND_SOURCE for r in rows)
    std = _item("Lifting Lugs L STD", "STD")
    r = soh.tree_rows([std])[0]
    assert r["text"] == "Lifting Lugs" and r["price"] == "STD"


def test_component_attr_groups_without_used_on():
    # `component` links a line that IS the thing; two of them merge.
    motor = _item("Motor C 3,194.37", "3,194.37", attrs={"component": "MOTOR",
                                                         "vendor": "Baldor"})
    base = _item("Motor Slide Base L 500.00", "500.00",
                 attrs={"component": "MOTOR", "motor_base": "SLIDE BASE"})
    comps = soh.components([motor, base])
    assert len(comps) == 1 and comps[0]["name"] == "MOTOR"
    assert comps[0]["attributes"]["vendor"] == "Baldor"
    assert comps[0]["attributes"]["motor_base"] == "SLIDE BASE"
    assert "component" not in comps[0]["attributes"]


def test_same_location_lines_merge_to_one_component_and_flag_conflict():
    # 421967: two "Outlet, Flanged, ..." lines describe the ONE outlet, but the
    # extractor named it only in prose (flange_scope=OUTLET, no component). They
    # must fold into a single [OUTLET] component, and their disagreement on a
    # single-valued attribute (flange_type) is surfaced — not left as two
    # look-alike components, one of them wrong.
    punched = _item("Outlet, Flanged, Punched L STD", "STD",
                    attrs={"flange_scope": "OUTLET", "flange_type": "PUNCHED"})
    unpunched = _item("Outlet, Flanged, Unpunched L 250.00", "250.00",
                      attrs={"flange_scope": "OUTLET", "flange_type": "UNPUNCHED"})
    comps = soh.components([punched, unpunched])
    assert len(comps) == 1 and comps[0]["name"] == "OUTLET" and comps[0]["keyed"]
    assert set(comps[0]["attributes"]["flange_type"].split(" | ")) == {"PUNCHED", "UNPUNCHED"}
    assert any("CONFLICTING flange_type" in r["text"] for r in comps[0]["review"])
    # A line already tied to another component keeps that tie — an inlet flange
    # that is part of the IVC stays under IVC, it does not become its own INLET.
    ivc_flange = _item("Inlet, Flanged, Punched (with IVC) L 1,559.00", "1,559.00",
                       used_on="IVC", attrs={"flange_scope": "INLET"})
    assert soh.group_key(ivc_flange) == "IVC"


def test_no_merge_on_loose_tag_overlap():
    # Two flange lines share the FLANGE tag but nothing ties them to one
    # thing -> two components (an inlet flange and an outlet flange).
    inlet = _item("Inlet, Flanged, Punched L 100.00", "100.00",
                  tags=["FLANGE", "INLET"], attrs={"flange_scope": "INLET"})
    outlet = _item("Outlet, Flanged, Punched L 100.00", "100.00",
                   tags=["FLANGE", "OUTLET"], attrs={"flange_scope": "OUTLET"})
    assert len(soh.components([inlet, outlet])) == 2


def test_review_flags_and_review_keys_stay_off_attributes():
    it = _item("Motor C 3,194.37", "3,194.37",
               attrs={"vendor": "Baldor", "used_on_review": "INCONCLUSIVE"},
               flags=["UNTAGGED"])
    c = soh.components([it])[0]
    assert "used_on_review" not in c["attributes"]
    assert c["review"] == [{"text": "UNTAGGED", "item_no": 1}]


def test_unclassified_detail_is_review_tied_to_source_item():
    it = _item("Motor C 3,194.37", "3,194.37",
               details=["Odd winding instruction"],
               attrs={"component": "MOTOR"},
               flags=["UNCLASSIFIED DETAIL: Odd winding instruction"])
    rows = soh.tree_rows([it])
    reviews = [row for row in rows if row["kind"] == soh.KIND_REVIEW]
    assert reviews == [{
        "depth": 1,
        "kind": soh.KIND_REVIEW,
        "text": "UNCLASSIFIED DETAIL: Odd winding instruction",
        "price": "",
        "item_no": 1,
    }]
    assert all(row["kind"] != "DETAIL" for row in rows)


def test_component_sits_where_first_evidence_appeared():
    main, handle, _f = _ivc_items()
    solo = _item("Base Fan L 16,649.00", "16,649.00")
    rows = soh.tree_rows([handle, solo, main])
    assert rows[0]["text"] == "[IVC] — 2 lines"   # first member was line 1
    base = next(r for r in rows if "Base Fan" in r["text"])
    assert base["kind"] == soh.KIND_COMPONENT and base["item_no"] == 2


def test_price_and_type_tails_are_stripped_from_line_text():
    main, _h, _f = _ivc_items()
    assert soh.line_text(main) == "Inlet Volume Control, Low Leak, Automatic"


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
