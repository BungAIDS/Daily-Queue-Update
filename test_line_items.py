"""Tests for the sales-order line-item capture / normalization / search logic.

No pytest needed — run it directly:

    python test_line_items.py

Real SO PDFs only exist on the work machine, so extraction is tested against
synthetic reconstructed-text lines shaped like the dumps (priced item rows,
an Additional Features section, totals/footer noise, CO-history lines). The
normalization, tagging, store, and search logic is pure Python and fully
checked here.
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import line_items as li


SO_LINES = [
    "Sales Order 421314",
    "Sold To: ACME CORP   Ship To: ACME CORP PLANT 2",
    "Customer PO: 4500123  Terms: Net 30",
    "Design 53 Industrial Exhauster",
    "Qty Design Size Arrangement MotorPos Class",
    "1 53 365 DB W 2",
    "1 BASE FAN SIZE 365 ARR. 1 CW THD 1,847.00 1,847.00",
    "1 SPARK RESISTANT CONST. TYPE B 312.00",
    "SS SHAFT SLEEVE 184.00",
    "SHAFT SEAL - CERAMIC FELT N/C",
    "EPOXY COATED INTERIOR & EXTERIOR 96.50",
    "Fan suitable for 3600 rpm Motor",          # unpriced, outside a section
    "ADDITIONAL FEATURES",
    "EXTENDED LUBE LINES W/ ZERK FITTINGS",
    "VIBRATION ISOLATORS - RUBBER IN SHEAR",
    "2 ACCESS DOOR, QUICK OPEN",
    "NOTES:",
    "CO#1 5/21/26 DG - ADDED SHAFT SEAL PER CUSTOMER",
    "TOTAL BILLING 2,439.50",
    "Freight Allowed",
    "Page 1 of 2",
]


# Verbatim line sequences from real CBC Sales-Order dumps (jobs 421473 and
# 421314, 2026-06-11 discovery) — the regression base the rules were fitted to.
REAL_LINES_473 = [
    "Chicago Blower Corporation Sales Order",
    "Date:",
    "Design 16A SW",
    "Order # Rep Ref. # Customer P.O. # Fan Serial Number:",
    "421473 7074-49840-00-AI26 7074-49840-00-AI26",
    "Sold To: Ship To:",
    "INNO-VENT INDUSTRIAL INC. INNO-VENT INDUSTRIAL INC.",
    "Total Commission: - BUY SELL w/ comm to dest rep (Rev 0%)",
    "Qty Design Size Arrangement Motor Pos Class Rotation Discharge % Width Wheel Type",
    "1 D16A 37 A/9H Z 3 CW TH 100 LS",
    "CFM OV SP RPM BHP Design Temp Max Temp Elev. Density BHP @ 41° F",
    "34000 4552 21 963 184.4 70 95 1352 0.0714 194.99",
    "Type Price Freight Markup Net Comm.",
    "Base Fan (Base fan Suitable for 447T motor frame, L 69,648.00",
    "Inquiry Num: 909-26-1604)",
    "Motor (Model ECP84407TR-5 complete with : C 11,435.83 457.00 1,903.00 13,796.00 571.00",
    "Vendor: Baldor/Reliance",
    "200 HP, 1800 RPM, Enclosure: TEFC Severe",
    "Duty",
    "447T, Cast Iron, Foot Mounted, 3/60/575, F1,",
    "1.15 SF",
    "Quote Num: 1101260457",
    "Mounting Charge L 1,336.00",
    "Drive N 4,366.00 437.00",
    "Constant Speed, SF: 1.3",
    "CBC Mount",
    "Mounting Charge L INC",
    "Access Door, Quick Clamp L 617.00",
    "Door Location: @9:00",
    "Belt Guard, Painted Safety Yellow L 2,338.00",
    "Extended Grease Fittings L 280.00",
    "Housing Drain with Plug L 129.00",
    "Housing, Heavy Duty L 4,754.00",
    "Include 3D STEP Drawings L 991.00",
    "Inlet, Flanged, Punched L 312.00",
    "v1.8.1.5 -1-",
    "Chicago Blower Corporation Sales Order (cont.)",
    "Order # Rep Ref. # Customer P.O. # Page 2 of 2",
    "421473 7074-49840-00-AI26 7074-49840-00-AI26",
    "Mechanical Run Test, Not Available L STD",
    "Outlet, Flanged, Punched L STD",
    "Shaft and Bearing Guard, Painted Safety Yellow L 805.00",
    "Weights on drawing, Inquiry Num: 909-26-1812 L",
    "Inlet flanged C 776.00 213.00 989.00 64.00",
    "Vendor: FlexibleCompensators",
    "Quote Num: 34592",
    "Ship Loose",
    "Product: Expansion Joint",
    "Freight to CBC N 125.00",
    "Ship Loose Charge N 49.00 5.00",
    "List Total Each 81,710.00 0.38 31,050.00 2,484.00",
    "Lead Time: 76 working days Product 51,866.00",
    "See Additional Features / Notes below Freight",
    "Total 51,866.00",
    "Sales Tax (NT)",
    "Buy/Sell Deduction -3,631.00",
    "Total Billing 48,235.00",
    "Additional Features / Notes:",
    "E-Mail Prints to: hmallette@inno-vent.ca",
    "Run Test - N/A (Send to Sales for Run Test Availability Check if CBC is mounting the motor)",
    "Fan Drawings:",
    "Emailed Mailed",
    "Fan Drawings Both",
    "O & M X",
    "Motor Prints X",
    "Motor Data Sheets X",
    "Buyout Prints (e.g. silencer, filter, etc.) X",
    "Other",
    "v1.8.1.5 -2-",
]

REAL_LINES_314 = [
    "Chicago Blower Corporation CO #2 Sales Order",
    "Design 34 Vaneaxial Belt Drive",
    "Qty Design Size Arrangement Motor Pos Class Rotation Discharge % Width Wheel Type",
    "1 D34 15 N/A N/A 2 100 M",
    "CFM OV SP RPM BHP Design Temp Max Temp Elev. Density BHP @ 70° F",
    "3200 2607 2 3344 2.52 70 70 0 0.075 2.52",
    "Type Price Freight Markup Net Comm.",
    "Base Fan (Base fan, Suitable for 3600rpm Motor, L 5,761.00",
    "Inquiry Num: 317-26-1510)",
    "Motor C 254.83 5.00 62.00 322.00 19.00",
    "Vendor: Toshiba or equivalent",
    "3 HP, 3600 RPM, Enclosure: TEFC Premium",
    "182T, Cast Iron, Foot Mounted, 3/60/230/460,",
    "F1, 1.15 SF",
    "Model #0032SDSR41A-P, PN 06-6-0020-01",
    "VFD Suitable",
    "CBC Mount",
    "Mounting Charge L 479.00",
    "Drive (Drive Set, Constant Speed (CBC mounted), N 360.00 36.00",
    "Belt Guard, Painted Safety Yellow L 945.00",
    "Mechanical Run Test, Standard L STD",
    "Wheel, Steel L STD",
    "IVD C 2,750.79 440.00 3,191.00 132.00",
    "Vendor: Ruskin",
    "Quote Num: 042426GCM2",
    "Ship Loose",
    "Product: Damper",
    "Freight to CBC N 223.00",
    "List Total Each 8,181.00 0.5000 4,090.00 818.00",
    "Lead Time: 60 working days Product 8,346.00",
    "Total Billing 9,564.00",
    "Additional Features / Notes:",
    "C/O #2 5/15/26 ECR: CORRECTED TOTAL BILLING.",
    "CO#1 050826 AMF - CORRECTED CLASS",
    "CASH IN ADVANCE",
    "NO TAXES",
    "Total Billing $ $9,023.00",
    "Run Test - Required",
]


def test_real_so_std_inc_and_bare_type_letter():
    items = {it["norm"]: it for it in li.extract_items(REAL_LINES_473)}
    # STD / INC in the price column (after the L/C/N type letter) are items...
    assert items["MECHANICAL RUN TEST NOT AVAILABLE"]["price"] == "STD"
    assert "OUTLET FLANGED PUNCHED" in items
    inc = {it["norm"]: it for it in li.extract_items(["Temporary Fan Feature L INC"])}
    assert inc["TEMPORARY FAN FEATURE"]["price"] == "INC"
    assert "MOUNTING CHARGE" not in items
    # ...and so is a row whose price column is simply empty (trailing bare L).
    assert any(n.startswith("WEIGHTS ON DRAWING") for n in items), items.keys()
    # But "INC." at the end of a company name never makes an address an item.
    assert not any("INNO" in n for n in items), items.keys()


def test_real_so_type_letter_stripped_and_price_column():
    items = {it["norm"]: it for it in li.extract_items(REAL_LINES_473)}
    bg = items["BELT GUARD PAINTED SAFETY YELLOW"]   # no trailing " L" in norm
    assert bg["ptype"] == "L" and bg["price"] == "2,338.00", bg
    # Multi-column money tail keeps the LEFTMOST (the Price column), not Comm.
    motor = items["MOTOR MODEL ECP84407TR 5 COMPLETE WITH"]
    assert motor["price"] == "11,435.83", motor


def test_real_so_noise_excluded():
    for lines in (REAL_LINES_473, REAL_LINES_314):
        joined = " | ".join(it["norm"] for it in li.extract_items(lines))
        for bad in ("LIST TOTAL", "LEAD TIME", "CUSTOMS", "FAN DRAWINGS",
                    "EMAILED", "BUYOUT PRINTS", "CHICAGO BLOWER", "DEDUCTION",
                    "COMMISSION", "CORRECTED CLASS"):
            assert bad not in joined, (bad, joined)
        # The CFM/RPM performance values row (numbers only) is not an item,
        # and neither is the spec-table row under the Qty/Design/Size header.
        assert "34000" not in joined and "D16A" not in joined, joined
        assert "3200 2607" not in joined and "D34" not in joined, joined


def test_real_so_details_attached():
    items = {it["norm"]: it for it in li.extract_items(REAL_LINES_314)}
    motor = items["MOTOR"]
    det = " | ".join(motor["details"])
    assert "Toshiba" in det and "3 HP" in det and "VFD Suitable" in det, det
    assert "VFD" in motor["tags"], motor["tags"]  # detail lines drive tags too
    # Page furniture between items must never attach as a detail.
    items473 = {it["norm"]: it for it in li.extract_items(REAL_LINES_473)}
    det = " | ".join(items473["INLET FLANGED PUNCHED"]["details"])
    assert "Chicago" not in det and "7074" not in det and "v1.8" not in det, det


def test_real_so_ivd_and_buyout_tagging():
    items = {it["norm"]: it for it in li.extract_items(REAL_LINES_314)}
    ivd = items["INLET VANE DAMPER"]            # IVD abbreviation expanded
    assert "DAMPER" in ivd["tags"] and "INLET VANES" in ivd["tags"], ivd
    assert any("Ruskin" in d for d in ivd["details"]), ivd["details"]
    # The flanged expansion-joint buyouts tag FLEX CONNECTOR via "Product:".
    items473 = {it["norm"]: it for it in li.extract_items(REAL_LINES_473)}
    assert "FLEX CONNECTOR" in items473["INLET FLANGED"]["tags"]
    assert "EXTENDED LUBE" in items473["EXTENDED GREASE FITTINGS"]["tags"]
    assert "HEAVY DUTY" in items473["HOUSING HEAVY DUTY"]["tags"]
    assert "3D STEP DRAWINGS" in items473["INCLUDE 3D STEP DRAWINGS"]["tags"]


def test_real_so_search_reaches_details():
    store = li.load_store(Path("/nonexistent/store.json"))
    li.record_job(store, "421314", li.extract_items(REAL_LINES_314))
    li.record_job(store, "421473", li.extract_items(REAL_LINES_473))
    assert [h["job"] for h in li.search(store, ["toshiba"])] == ["421314"]
    assert [h["job"] for h in li.search(store, ["baldor"])] == ["421473"]
    assert [h["job"] for h in li.search(store, ["200 HP"])] == ["421473"]
    assert {h["job"] for h in li.search(store, [], tag="DAMPER")} == {"421314"}


def test_priced_lines_captured():
    items = li.extract_items(SO_LINES)
    raws = [it["raw"] for it in items]
    assert any("BASE FAN" in r for r in raws), raws
    assert any("SPARK RESISTANT" in r for r in raws), raws
    assert any("SHAFT SLEEVE" in r for r in raws), raws
    assert any("CERAMIC FELT" in r for r in raws), raws  # N/C counts as priced


def test_section_lines_captured_unpriced():
    items = {it["norm"]: it for it in li.extract_items(SO_LINES)}
    assert any("EXTENDED LUBE" in n for n in items), items.keys()
    assert any("VIBRATION ISOLATORS" in n for n in items), items.keys()
    it = next(v for n, v in items.items() if "ACCESS DOOR" in n)
    assert it["section"] == "ADDITIONAL FEATURES"
    assert it["qty"] == "2"


def test_noise_skipped():
    norms = " | ".join(it["norm"] for it in li.extract_items(SO_LINES))
    for bad in ("TOTAL BILLING", "FREIGHT", "PAGE 1", "SOLD TO", "CUSTOMER PO",
                "CO#1", "ADDED SHAFT SEAL PER CUSTOMER", "NOTES"):
        assert bad not in norms, (bad, norms)
    # The spec-table row ("1 53 365 DB W 2") has no price and no section.
    assert "365 DB" not in norms, norms


def test_section_closes_at_notes():
    # The CO-history line sits after NOTES:, so even though section capture was
    # on, nothing after the section_end marker is captured.
    items = li.extract_items(["ACCESSORIES", "BELT GUARD", "NOTES:", "SOME NOTE TEXT"])
    assert [it["norm"] for it in items] == ["BELT GUARD"], items


def test_price_and_qty_parsed():
    items = {it["norm"]: it for it in li.extract_items(SO_LINES)}
    spark = next(v for n, v in items.items() if "SPARK" in n)
    assert spark["price"] == "312.00" and spark["qty"] == "1", spark
    base = next(v for n, v in items.items() if "BASE FAN" in n)
    assert base["price"] in ("1,847.00", "1847.00"), base  # double money tail stripped
    assert "1,847" not in base["norm"], base
    seal = next(v for n, v in items.items() if "CERAMIC" in n)
    assert seal["price"].upper().replace("/", "") == "NC", seal


def test_bare_integers_are_not_prices():
    # "3600" in "suitable for 3600 rpm Motor" must not read as a price column.
    desc, price = li.split_price_tail("FAN SUITABLE FOR 3600 RPM MOTOR")
    assert price == "" and desc.endswith("MOTOR")
    desc, price = li.split_price_tail("DAMPER 1,250.00")
    assert price == "1,250.00" and desc == "DAMPER"
    desc, price = li.split_price_tail("GUARD $96.50 EACH")
    assert price == "$96.50" and desc == "GUARD"


def test_normalization_converges_variants():
    # Different order-entry spellings of the same option meet at one norm.
    a = li.normalize_text("SS SHAFT SLEEVE 184.00")
    b = li.normalize_text("1  Stainless Steel Shaft Sleeve")
    c = li.normalize_text("S/S SHAFT SLEEVE - N/C")
    assert a == b == c == "STAINLESS STEEL SHAFT SLEEVE", (a, b, c)
    assert li.normalize_text("316SS sleeve") == "316 STAINLESS STEEL SLEEVE"
    assert li.normalize_text("SST exterior") == "STAINLESS STEEL EXTERIOR"
    assert li.normalize_text("Wheel Aluminium AMCA B") == "WHEEL ALUMINUM AMCA B"
    # "EXT" stays unexpanded (ambiguous with EXTERIOR) — the tag converges both.
    assert "EXTENDED LUBE" in li.tag_item(li.normalize_text("EXT LUBE LINES W/ ZERKS"))
    assert "EXTENDED LUBE" in li.tag_item(li.normalize_text("Extended lube lines w/ zerks"))
    assert li.normalize_text("W/DRAIN") == "WITH DRAIN"


def test_tagging():
    assert "SHAFT SEAL" in li.tag_item("SHAFT SEAL CERAMIC FELT")
    assert "SHAFT SEAL" in li.tag_item(li.normalize_text("Teflon shaft seal"))
    assert "SPARK RESISTANT" in li.tag_item("SPARK RESISTANT CONSTRUCTION TYPE B")
    assert "STAINLESS STEEL" in li.tag_item(li.normalize_text("316SS wheel"))
    assert "COATING" in li.tag_item("EPOXY COATED INTERIOR AND EXTERIOR")
    assert "VIBRATION ISOLATION" in li.tag_item("VIBRATION ISOLATORS RUBBER IN SHEAR")
    assert "V-BELT DRIVE" in li.tag_item(li.normalize_text("Drive"))
    assert "VIBRATION ISOLATION" in li.tag_item(li.normalize_text("Vibration Base"))
    assert "INLET VANES" in li.tag_item(li.normalize_text("Inlet Volume Control Automatic"))
    assert "COATING" in li.tag_item(li.normalize_text("Passivation of Welds"))
    assert "LINING" in li.tag_item(li.normalize_text("Firmex liners on blades and housing scroll"))
    assert "EXTENDED LUBE" in li.tag_item(li.normalize_text("Extended Grease Leads"))
    assert "FLEXIBLE COUPLING" in li.tag_item(li.normalize_text("Flexible Coupling Falk Type T Steelflex"))
    assert "ALUMINUM" in li.tag_item(li.normalize_text("Wheel Aluminium AMCA B"))
    assert "WHEEL" in li.tag_item(li.normalize_text("Percent Width 78.7%"))
    assert "LIFTING LUGS" in li.tag_item(li.normalize_text("Lifting Lugs"))
    assert "NAMEPLATE" in li.tag_item(li.normalize_text("Fan Nameplate without Chicago Blower Name"))
    assert "PACKAGING" in li.tag_item(li.normalize_text("ISPM Wood Inspection Stamp"))
    assert "SHIPPING" in li.tag_item(li.normalize_text("Ship Loose Freight Included"))
    assert "ACTUATOR" in li.tag_item(li.normalize_text("Actuator for IVC Bettis #RPED100"))
    assert "DRIVE COMPONENTS" in li.tag_item(li.normalize_text("Motor Sheave/Bushing 3B5V74/B"))
    assert "V-BELT DRIVE" in li.tag_item(li.normalize_text("Drive (Max/Min RPM: 1531/1531, 3 belts: B112"))
    assert "SPECIAL CONSTRUCTION" in li.tag_item(li.normalize_text("Tie Rod Support"))
    assert "SPECIAL CONSTRUCTION" in li.tag_item(li.normalize_text("Loc Tite on the set screw threads"))
    assert "SPECIAL CONSTRUCTION" in li.tag_item(li.normalize_text("Continuous Weld Airstream"))
    assert "SPECIAL CONSTRUCTION" in li.tag_item(li.normalize_text("Earthing Boss"))
    assert "INSPECTION" in li.tag_item(li.normalize_text("Customer Final Inspection"))
    assert "INSPECTION" in li.tag_item(li.normalize_text("General Mill Certifications"))
    assert "LABEL" in li.tag_item(li.normalize_text("FEI Label Inquiry Num"))
    assert li.tag_item("SOMETHING NOBODY EVER ORDERED") == []


def test_tags_label():
    items = li.extract_items(SO_LINES)
    label = li.tags_label(items)
    assert "SHAFT SEAL" in label and "SPARK RESISTANT" in label, label
    assert li.tags_label([]) == ""
    assert li.tags_label([{"raw": "X", "norm": "MYSTERY OPTION", "tags": []}]) == "(1 items)"


def test_rules_extension_file(tmp: Path):
    ext = tmp / "rules_ext.json"
    ext.write_text('{"tags": {"SHAFT GROUNDING": ["grounding\\\\s*ring"]},'
                   ' "abbreviations": {"GRND": "GROUNDING"}}', encoding="utf-8")
    rules = li.load_rules(ext)
    norm = li.normalize_text("SHAFT GRND RING", rules)
    assert norm == "SHAFT GROUNDING RING"
    assert "SHAFT GROUNDING" in li.tag_item(norm, rules)
    # Defaults are still present (extended, not replaced).
    assert "SHAFT SEAL" in li.tag_item("SHAFT SEAL", rules)


def test_store_roundtrip_and_merge(tmp: Path):
    path = tmp / "store.json"
    store = li.load_store(path)
    items = li.extract_items(SO_LINES)
    li.record_job(store, "421314", items, customer="ACME CORP", co_number=1,
                  so_pdf="Z:/SO/421314.pdf")
    li.save_store(store, path)
    store2 = li.load_store(path)
    assert store2["jobs"]["421314"]["customer"] == "ACME CORP"
    assert len(store2["jobs"]["421314"]["items"]) == len(items)
    # A sparse re-record (archive scan: no customer known) keeps the metadata.
    li.record_job(store2, "421314", items)
    assert store2["jobs"]["421314"]["customer"] == "ACME CORP"
    assert store2["jobs"]["421314"]["co_number"] == 1
    assert store2["jobs"]["421314"]["so_pdf"] == "Z:/SO/421314.pdf"


def _seeded_store() -> dict:
    store = li.load_store(Path("/nonexistent/line_items.json"))
    li.record_job(store, "421314", li.extract_items(SO_LINES), customer="ACME CORP")
    li.record_job(store, "421999", li.extract_items(
        ["1 BASE FAN 900.00", "TEFLON SHAFT SEAL 55.00", "BELT GUARD 80.00"]),
        customer="ZEECO")
    return store


def test_search_and_terms():
    store = _seeded_store()
    # AND (default): both jobs have a shaft seal...
    hits = li.search(store, ["shaft seal"])
    assert [h["job"] for h in hits] == ["421999", "421314"]  # newest job first
    # ...but only one also has spark-resistant construction.
    hits = li.search(store, ["shaft seal", "spark"])
    assert [h["job"] for h in hits] == ["421314"]
    # OR mode.
    hits = li.search(store, ["spark", "belt guard"], any_mode=True)
    assert {h["job"] for h in hits} == {"421314", "421999"}
    # Tag filter (canonical lookup): only the tagged items come back.
    hits = li.search(store, [], tag="SHAFT SEAL")
    assert {h["job"] for h in hits} == {"421314", "421999"}
    assert all("SHAFT SEAL" in it["tags"] for h in hits for it in h["matches"])
    # Tag + term AND at the JOB level — they may sit on different items.
    hits = li.search(store, ["spark"], tag="SHAFT SEAL")
    assert [h["job"] for h in hits] == ["421314"]
    raws = " | ".join(it["raw"] for it in hits[0]["matches"])
    assert "SPARK" in raws and "CERAMIC FELT" in raws, raws
    # Search hits normalized text, so "stainless" finds the raw "SS" item;
    # terms are ANDed, each may land on a different word of the same item.
    hits = li.search(store, ["stainless", "sleeve"])
    assert [h["job"] for h in hits] == ["421314"]


def test_search_fuzzy():
    store = _seeded_store()
    assert li.search(store, ["cermic"]) == []          # typo: no substring hit
    hits = li.search(store, ["cermic"], fuzzy=0.8)     # fuzzy catches it
    assert [h["job"] for h in hits] == ["421314"]


def test_ai_cache_applied():
    store = li.load_store(Path("/nonexistent/line_items.json"))
    items = li.extract_items(["Mystery Option L 12.00"])
    li.record_job(store, "421314", items)
    rec = store["jobs"]["421314"]
    it = rec["items"][0]
    assert it["norm"] in li.unknown_norms(store)
    store["ai_tags"][it["norm"]] = ["BASE FAN"]
    li.apply_ai_cache(rec["items"], store)
    assert "BASE FAN" in it["tags"]
    # Once cached it's never sent to the API again.
    assert it["norm"] not in li.unknown_norms(store)


def test_ai_cache_does_not_override_current_rules():
    store = li.load_store(Path("/nonexistent/line_items.json"))
    items = li.extract_items(["Base Fan L 900.00"])
    assert items[0]["tags"] == ["BASE FAN"]
    store["ai_tags"][items[0]["norm"]] = ["UNITARY BASE"]
    li.apply_ai_cache(items, store)
    assert items[0]["tags"] == ["BASE FAN"]


def test_renormalize_uses_current_rules():
    store = _seeded_store()
    # Sabotage the stored derived fields; renormalize must rebuild from raw.
    for it in store["jobs"]["421314"]["items"]:
        it["norm"], it["tags"] = "WRONG", ["BOGUS"]
    n = li.renormalize_store(store)
    assert n >= len(store["jobs"]["421314"]["items"])
    norms = [it["norm"] for it in store["jobs"]["421314"]["items"]]
    assert "STAINLESS STEEL SHAFT SLEEVE" in norms, norms


def test_audit_untagged_uses_current_rules():
    store = li.load_store(Path("/nonexistent/line_items.json"))
    items = li.extract_items(["Drive L 100.00", "Mystery Option L 12.00"])
    items.append({"raw": "Product:", "norm": "PRODUCT", "qty": "", "price": "",
                  "ptype": "", "section": "", "details": [], "tags": []})
    li.record_job(store, "421000", items)
    rows = li.audit_untagged(store)
    assert [r["norm"] for r in rows] == ["MYSTERY OPTION"], rows


def test_actuator_attributes():
    items = {it["norm"]: it for it in li.extract_items([
        "Actuator for IVC, Bettis #RPED100 Double Acting C 5,054.00 809.00",
        "Pneumatic Actuator. Complete with a VAC",
        "Vendor: Novaspect",
        "Product: Actuator",
        "Operation: Automatic",
        "Actuator Manufacturer: Bettis RPD",
        "Actuator Supplied By: CBC (Mounted)",
        "Fail Position Upon Loss of Supply/Air/Power: In",
        "Place",
        "Fail Position Upon Loss of Signal: Closed",
        "Actuator Mounting: Bracket and actuator",
    ])}
    actuator = items["ACTUATOR FOR IVC BETTIS #RPED100 DOUBLE ACTING"]
    attrs = actuator["attributes"]
    assert {"ACTUATOR", "INLET VANES"} <= set(actuator["tags"])
    assert attrs["component"] == "ACTUATOR"
    assert attrs["used_on"] == "IVC"
    assert attrs["vendor"] == "Novaspect"
    assert attrs["product"] == "Actuator"
    assert attrs["manufacturer"] == "Bettis RPD"
    assert attrs["model"].upper().replace(" ", "") == "BETTIS#RPED100"
    assert attrs["size"] == "RPED100"
    assert attrs["operation"] == "Automatic"
    assert attrs["supplied_by"] == "CBC (Mounted)"
    assert attrs["mounting"] == "Bracket and actuator"
    assert attrs["fail_power"] == "In Place"
    assert attrs["fail_signal"] == "Closed"


def test_drive_attributes():
    items = {it["norm"]: it for it in li.extract_items([
        "Drive (Max/Min RPM: 1531/1531, 3 belts: B112, C 409.00 125.00",
        "Motor Sheave/Bushing: 3B5V74/B (1 7/8\"), Fan",
        "Sheave/Bushing: 3B5V86/B (2 3/16\"), Actual SF:",
        "1.31, Actual CD: 44.34)",
        "Constant Speed, SF: 1.3",
    ])}
    drive = items["DRIVE MAX MIN RPM 1531 1531 3 BELTS B112"]
    attrs = drive["attributes"]
    assert {"V-BELT DRIVE", "DRIVE COMPONENTS"} <= set(drive["tags"])
    assert attrs["component"] == "V-BELT DRIVE"
    assert attrs["belt_qty"] == "3"
    assert attrs["belt"] == "B112"
    assert attrs["max_rpm"] == "1531"
    assert attrs["min_rpm"] == "1531"
    assert attrs["drive_sheave_bushing"] == '3B5V74/B (1 7/8")'
    assert attrs["driven_sheave_bushing"] == '3B5V86/B (2 3/16")'
    assert attrs["actual_sf"] == "1.31"
    assert attrs["actual_cd"] == "44.34"
    assert attrs["service_factor"] == "1.3"


def test_selected_drive_table_attributes():
    items = li.extract_items([
        "1515/1515 B116 3 3TB80 Q1 3B5V94 B 1.49 45.24 436.00",
        "Notes:",
        "* Selected Drive",
        "Specified minimum belt service factor: 1.3",
        "Center Distance with allowance for install and take-up: 43.5 - 46.5",
    ])
    drive = items[0]
    attrs = drive["attributes"]
    assert "DRIVE COMPONENTS" in drive["tags"]
    assert attrs["belt"] == "B116"
    assert attrs["belt_qty"] == "3"
    assert attrs["drive_sheave_bushing"] == "3TB80 Q1"
    assert attrs["driven_sheave_bushing"] == "3B5V94 B"
    assert attrs["actual_sf"] == "1.49"
    assert attrs["actual_cd"] == "45.24"
    assert attrs["service_factor"] == "1.3"
    assert attrs["center_distance_range"] == "43.5 - 46.5"
    bare = li.extract_items([
        "1515/1515 B116 3 3B5V80 B 3B5V94 B 1.49 45.24 428.00",
    ])[0]
    assert "DRIVE COMPONENTS" in bare["tags"]
    assert bare["attributes"]["belt"] == "B116"
    assert bare["attributes"]["drive_sheave_bushing"] == "3B5V80 B"
    assert bare["attributes"]["driven_sheave_bushing"] == "3B5V94 B"


def test_inquiry_number_attributes():
    items = {it["norm"]: it for it in li.extract_items([
        "3D Drawings, InquiryNum:333-25-1622 L",
        "Access Door, Quick Clamp @ 10:30 position, Inquiry L 645.00",
        "Num: 352-23-2696",
        "Base Fan (Base fan, Suitable for 3600rpm Motor, L 5,761.00",
        "Inquiry Num: 317-26-1510)",
    ])}
    assert items["3D DRAWINGS INQUIRYNUM 333 25 1622"]["attributes"]["inquiry_num"] == "333-25-1622"
    assert items["ACCESS DOOR QUICK CLAMP 10 30 POSITION INQUIRY"]["attributes"]["inquiry_num"] == "352-23-2696"
    assert items["BASE FAN BASE FAN SUITABLE FOR 3600RPM MOTOR"]["attributes"]["inquiry_num"] == "317-26-1510"


def test_search_reaches_inquiry_attributes():
    store = li.load_store(Path("/nonexistent/line_items.json"))
    li.record_job(store, "421463", li.extract_items([
        "3D Drawings, InquiryNum:333-25-1622 L",
    ]))
    assert [h["job"] for h in li.search(store, ["333-25-1622"])] == ["421463"]


def test_inquiry_counts():
    store = li.load_store(Path("/nonexistent/line_items.json"))
    li.record_job(store, "421463", li.extract_items([
        "3D Drawings, InquiryNum:333-25-1622 L",
        "Add COG and Weights to the fan drawing, Inquiry L",
        "Num: 333-25-1622",
    ]))
    li.record_job(store, "421464", li.extract_items([
        "Plan View on Customer Drawing, Inquiry Num: 333-25-1622 L",
    ]))
    counts = {num: (jobs, items, job_list) for num, jobs, items, job_list in li.inquiry_counts(store)}
    assert counts["333-25-1622"] == (2, 3, ["421464", "421463"])


def test_data_branch_noise_skipped():
    lines = [
        "ADDITIONAL FEATURES",
        "Product:",
        "Product 7,623.00",
        "Prints",
        "Warranty Exclusive 3 Year",
        "Paymode-X invoice processing system N 52.00",
        "Includes Paymode-X invoice processing system",
        "be necessary. For more information on this process, contact Chicago Blower.",
        "Additional Shipping Notes must use BOL provided by CH Robinson",
        "Do Not Stack sticker on all 4 sides of skid",
        "Drive L 100.00",
    ]
    items = {it["norm"]: it for it in li.extract_items(lines)}
    assert "DRIVE" in items
    for bad in ("PRODUCT", "PRINTS", "WARRANTY", "SHIPPING NOTES", "DO NOT STACK"):
        assert not any(bad in n for n in items), items


def test_used_on_requires_damper_context():
    flange = li.extract_items(["Outlet, Flanged, Punched L STD"])[0]
    assert "used_on" not in flange["attributes"]
    damper = li.extract_items(["Outlet Damper, Opposed L 100.00"])[0]
    assert damper["attributes"]["used_on"] == "OUTLET DAMPER"


def test_tag_counts():
    store = _seeded_store()
    counts = {t: (j, i) for t, j, i in li.tag_counts(store)}
    assert counts["SHAFT SEAL"][0] == 2   # both jobs
    assert counts["SPARK RESISTANT"][0] == 1


def test_inventory_xlsx(tmp: Path):
    import find_orders
    from openpyxl import load_workbook
    store = _seeded_store()
    hits = li.search(store, [])  # no filter -> the full inventory
    out = find_orders.write_xlsx(hits, tmp / "inv.xlsx")
    ws = load_workbook(out).active
    assert ws.cell(1, 1).value == "Job #"
    assert ws.max_row == 1 + sum(len(h["matches"]) for h in hits)


def test_feature_matrix_sheet(tmp: Path):
    import find_orders
    from openpyxl import load_workbook
    store = _seeded_store()
    out = find_orders.write_xlsx(li.search(store, []), tmp / "inv.xlsx", store)
    ws = load_workbook(out)["Feature Matrix"]
    headers = [c.value for c in ws[1]]
    rows = {ws.cell(r, 1).value: r for r in range(2, ws.max_row + 1)}
    assert {"421314", "421999"} <= set(rows)
    # Green ✓ where the order has the feature, red (blank) where it doesn't.
    seal = headers.index("SHAFT SEAL") + 1
    spark = headers.index("SPARK RESISTANT") + 1
    assert ws.cell(rows["421314"], seal).value == "✓"
    assert ws.cell(rows["421999"], seal).value == "✓"
    assert ws.cell(rows["421314"], spark).value == "✓"
    assert ws.cell(rows["421999"], spark).value is None
    assert str(ws.cell(rows["421999"], spark).fill.start_color.rgb).endswith("FFC7CE")
    assert str(ws.cell(rows["421999"], seal).fill.start_color.rgb).endswith("C6EFCE")


def test_feature_matrix_full_profile_when_filtered(tmp: Path):
    # A search narrows the hits, but each matched job's matrix row still shows
    # its WHOLE feature profile (from the store), not just the matched items.
    import find_orders
    from openpyxl import load_workbook
    store = _seeded_store()
    hits = li.search(store, ["spark"])           # only 421314, one matched item
    out = find_orders.write_xlsx(hits, tmp / "f.xlsx", store)
    ws = load_workbook(out)["Feature Matrix"]
    headers = [c.value for c in ws[1]]
    assert ws.max_row == 2 and ws.cell(2, 1).value == "421314"
    assert ws.cell(2, headers.index("SHAFT SEAL") + 1).value == "✓"  # not the spark item


def _mini_pdf(lines: list, path: Path) -> None:
    """A minimal one-page text PDF (Helvetica, one string per line) — just
    enough for pdfplumber to extract real words back out, so the whole
    parse_sales_order_pdf -> extract_items path runs against an actual PDF."""
    content = ["BT", "/F1 10 Tf"]
    y = 760
    for ln in lines:
        esc = str(ln).replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
        content.append(f"1 0 0 1 40 {y} Tm ({esc}) Tj")
        y -= 14
    content.append("ET")
    stream = "\n".join(content).encode("latin-1", "replace")
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R "
        b"/Resources << /Font << /F1 5 0 R >> >> >>",
        b"<< /Length " + str(len(stream)).encode() + b" >>\nstream\n" + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    out = bytearray(b"%PDF-1.4\n")
    offsets = []
    for i, body in enumerate(objs, start=1):
        offsets.append(len(out))
        out += f"{i} 0 obj\n".encode() + body + b"\nendobj\n"
    xref_pos = len(out)
    out += f"xref\n0 {len(objs) + 1}\n".encode() + b"0000000000 65535 f \n"
    for off in offsets:
        out += f"{off:010d} 00000 n \n".encode()
    out += (f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\n"
            f"startxref\n{xref_pos}\n%%EOF").encode()
    path.write_bytes(bytes(out))


def test_pdf_roundtrip(tmp: Path):
    # End to end: a real (tiny) PDF through parse_sales_order_pdf.
    try:
        import pdfplumber  # noqa: F401
    except ImportError:
        print("      (pdfplumber not installed - skipping the pdf round-trip)")
        return
    from sales_orders import parse_sales_order_pdf
    pdf = tmp / "421314 - Sales Order CO#1.pdf"
    _mini_pdf(SO_LINES, pdf)
    res = parse_sales_order_pdf(pdf)
    norms = [it["norm"] for it in res["line_items"]]
    assert any("SPARK RESISTANT" in n for n in norms), norms
    assert any("EXTENDED LUBE" in n for n in norms), norms     # section capture
    assert not any("TOTAL" in n for n in norms), norms
    assert res["co_history"] and "CO#1" in res["co_history"][0].replace(" ", "")


def main() -> int:
    passed = 0
    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d)
        for name, fn in sorted(globals().items()):
            if not name.startswith("test_") or not callable(fn):
                continue
            (fn(tmp) if "tmp" in fn.__code__.co_varnames else fn())
            print(f"  ok  {name}")
            passed += 1
    print(f"\n{passed} tests passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
