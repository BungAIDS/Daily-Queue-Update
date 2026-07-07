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
    assert {"BEARINGS", "MOTOR"} <= set(items473["EXTENDED GREASE FITTINGS"]["tags"])
    assert "EXTENDED LUBE" not in items473["EXTENDED GREASE FITTINGS"]["tags"]
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
    assert any("ZERK FITTINGS" in n for n in items), items.keys()
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
    assert li.normalize_text("304SST Airstream") == "304 STAINLESS STEEL AIRSTREAM"
    assert li.normalize_text("316LSS housing") == "316L STAINLESS STEEL HOUSING"
    assert li.normalize_text("SST exterior") == "STAINLESS STEEL EXTERIOR"
    assert li.normalize_text("Wheel Aluminium AMCA B") == "WHEEL ALUMINUM AMCA B"
    # "EXT" stays unexpanded (ambiguous with EXTERIOR).
    assert li.normalize_text("EXT LUBE LINES W/ ZERKS") == "EXT LUBE LINES WITH ZERKS"
    assert li.normalize_text("W/DRAIN") == "WITH DRAIN"


def test_tagging():
    assert "SHAFT SEAL" in li.tag_item("SHAFT SEAL CERAMIC FELT")
    assert "SHAFT SEAL" in li.tag_item(li.normalize_text("Teflon shaft seal"))
    assert "SPARK RESISTANT" in li.tag_item("SPARK RESISTANT CONSTRUCTION TYPE B")
    assert "STAINLESS STEEL" in li.tag_item(li.normalize_text("316SS wheel"))
    assert "MATERIALS" in li.tag_item(li.normalize_text("304SST Airstream"))
    assert "COATING" in li.tag_item("EPOXY COATED INTERIOR AND EXTERIOR")
    assert "VIBRATION ISOLATION" in li.tag_item("VIBRATION ISOLATORS RUBBER IN SHEAR")
    assert "V-BELT DRIVE" in li.tag_item(li.normalize_text("Drive"))
    assert "VIBRATION ISOLATION" in li.tag_item(li.normalize_text("Vibration Base"))
    assert "INLET VANES" in li.tag_item(li.normalize_text("Inlet Volume Control Automatic"))
    assert "COATING" not in li.tag_item(li.normalize_text("Passivation of Welds"))
    assert "STAINLESS STEEL" in li.tag_item(li.normalize_text("Passivation of Welds"))
    assert "LINING" in li.tag_item(li.normalize_text("Firmex liners on blades and housing scroll"))
    assert "COUPLING" in li.tag_item(li.normalize_text("Flexible Coupling Falk Type T Steelflex"))
    assert "EXTREME TEMP" in li.tag_item(li.normalize_text("High Temp Fan"))
    assert "LOW LEAKAGE" in li.tag_item(li.normalize_text("Low Leak IVC"))
    assert "ALUMINUM" in li.tag_item(li.normalize_text("Wheel Aluminium AMCA B"))
    assert "MATERIALS" in li.tag_item(li.normalize_text("Wheel Aluminium AMCA B"))
    assert "WHEEL" in li.tag_item(li.normalize_text("Percent Width 78.7%"))
    assert "LIFTING LUGS" in li.tag_item(li.normalize_text("Lifting Lugs"))
    assert "NAMEPLATE" in li.tag_item(li.normalize_text("Fan Nameplate without Chicago Blower Name"))
    assert "PACKAGING" in li.tag_item(li.normalize_text("ISPM Wood Inspection Stamp"))
    assert "SHIPPING" in li.tag_item(li.normalize_text("Ship Loose Freight Included"))
    assert "ACTUATOR" in li.tag_item(li.normalize_text("Actuator for IVC Bettis #RPED100"))
    assert "3D STEP DRAWINGS" in li.tag_item(li.normalize_text("3D STEP File Drawings"))
    assert "DRIVE COMPONENTS" in li.tag_item(li.normalize_text("Motor Sheave/Bushing 3B5V74/B"))
    assert "SPLIT HOUSING" in li.tag_item(li.normalize_text("Drawing and split housing released into production"))
    assert "SPLIT HOUSING" in li.tag_item(li.normalize_text("Shipping Split"))
    assert "V-BELT DRIVE" in li.tag_item(li.normalize_text("Drive (Max/Min RPM: 1531/1531, 3 belts: B112"))
    assert "SPECIAL CONSTRUCTION" in li.tag_item(li.normalize_text("Tie Rod Support"))
    assert "SPECIAL CONSTRUCTION" in li.tag_item(li.normalize_text("Loc Tite on the set screw threads"))
    assert "SPECIAL CONSTRUCTION" in li.tag_item(li.normalize_text("Continuous Weld Airstream"))
    assert "SPECIAL CONSTRUCTION" in li.tag_item(li.normalize_text("Earthing Boss"))
    assert "INSPECTION" in li.tag_item(li.normalize_text("Customer Final Inspection"))
    assert "INSPECTION" in li.tag_item(li.normalize_text("General Mill Certifications"))
    assert "LABEL" in li.tag_item(li.normalize_text("FEI Label Inquiry Num"))
    assert "BALANCE" in li.tag_item(li.normalize_text("G2.5 Balance"))
    assert "BALANCE" in li.tag_item(li.normalize_text("Welded Balance Weights"))
    assert "BALANCE" not in li.tag_item(li.normalize_text("Balance Report"))
    assert "BALANCE" not in li.tag_item(li.normalize_text(
        "All Chicago Blower wheels are precision balanced"))
    assert "BEARINGS" in li.tag_item(li.normalize_text("Bearings, Standard"))
    assert "BEARINGS" in li.tag_item(li.normalize_text("Bearings, Split Pillow Block"))
    assert "BEARINGS" in li.tag_item(li.normalize_text("Repair Bearings (Pair), 2-3/16 Bore"))
    guard_tags = li.tag_item(li.normalize_text("Shaft, Bearing, and Coupling Guard, Painted Safety Yellow"))
    assert "SHAFT/BEARING/COUPLING GUARD" in guard_tags
    assert "BEARINGS" not in guard_tags
    assert "BEARINGS" not in li.tag_item(li.normalize_text("Motor with insulated bearings"))
    assert "BEARINGS" not in li.tag_item(li.normalize_text(
        "Paint Interior Wheel Exterior Motor Base Channel Base and Bearing Base"))
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
    store["jobs"]["421314"]["items"].append(
        {"raw": "Prints:", "norm": "PRINTS", "qty": "", "price": "",
         "ptype": "", "section": "", "details": [], "tags": ["DRAWINGS"], "attributes": {}}
    )
    # Sabotage the stored derived fields; renormalize must rebuild from raw.
    for it in store["jobs"]["421314"]["items"]:
        it["norm"], it["tags"] = "WRONG", ["BOGUS"]
    n = li.renormalize_store(store)
    assert n >= len(store["jobs"]["421314"]["items"])
    norms = [it["norm"] for it in store["jobs"]["421314"]["items"]]
    assert "STAINLESS STEEL SHAFT SLEEVE" in norms, norms
    assert "PRINTS" not in norms


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
    assert attrs["drive_sheave"] == "3B5V74"
    assert attrs["drive_bushing"] == 'B (1 7/8")'
    assert attrs["driven_sheave"] == "3B5V86"
    assert attrs["driven_bushing"] == 'B (2 3/16")'
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
    assert attrs["drive_sheave"] == "3TB80"
    assert attrs["drive_bushing"] == "Q1"
    assert attrs["driven_sheave"] == "3B5V94"
    assert attrs["driven_bushing"] == "B"
    assert attrs["actual_sf"] == "1.49"
    assert attrs["actual_cd"] == "45.24"
    assert attrs["selected_drive"] == "YES"
    assert attrs["service_factor"] == "1.3"
    assert attrs["center_distance_range"] == "43.5 - 46.5"
    bare = li.extract_items([
        "1515/1515 B116 3 3B5V80 B 3B5V94 B 1.49 45.24 428.00",
    ])[0]
    assert "DRIVE COMPONENTS" in bare["tags"]
    assert bare["attributes"]["belt"] == "B116"
    assert bare["attributes"]["drive_sheave_bushing"] == "3B5V80 B"
    assert bare["attributes"]["driven_sheave_bushing"] == "3B5V94 B"

    compact = li.extract_items([
        "*2567/2567 A52 1 AK46 AK30H H 1.25 21.14 99.00",
        "Notes:",
        "* Selected Drive",
        "Motor conduit box in F1 position.",
        "Specified minimum belt service factor: 1.2",
        "Center Distance range assumes adjustable base per drawing 7-2-94.",
    ])[0]
    compact_attrs = compact["attributes"]
    assert "DRIVE COMPONENTS" in compact["tags"]
    assert "DRAWINGS" not in compact["tags"]
    assert "MOTOR" not in compact["tags"]
    assert "SPECIAL CONSTRUCTION" not in compact["tags"]
    assert compact_attrs["selected_drive"] == "YES"
    assert compact_attrs["drive_sheave"] == "AK46"
    assert "drive_bushing" not in compact_attrs
    assert compact_attrs["driven_sheave"] == "AK30H"
    assert compact_attrs["driven_bushing"] == "H"
    assert compact_attrs["actual_sf"] == "1.25"
    assert compact_attrs["actual_cd"] == "21.14"


def test_inquiry_number_attributes():
    items = {it["norm"]: it for it in li.extract_items([
        "3D Drawings, InquiryNum:333-25-1622 L",
        "Access Door, Quick Clamp @ 10:30 position, Inquiry L 645.00",
        "Num: 352-23-2696",
        "Base Fan (Base fan, Suitable for 3600rpm Motor, L 5,761.00",
        "Inquiry Num: 317-26-1510)",
    ])}
    drawing_attrs = items["3D DRAWINGS INQUIRYNUM 333 25 1622"]["attributes"]
    assert "3D STEP DRAWINGS" in items["3D DRAWINGS INQUIRYNUM 333 25 1622"]["tags"]
    assert drawing_attrs["inquiry_num"] == "333-25-1622"
    assert drawing_attrs["drawing_type"] == "3D STEP DRAWINGS"
    assert drawing_attrs["drawing_scope"] == "3D FILE"
    assert items["ACCESS DOOR QUICK CLAMP 10 30 POSITION INQUIRY"]["attributes"]["inquiry_num"] == "352-23-2696"
    assert items["BASE FAN BASE FAN SUITABLE FOR 3600RPM MOTOR"]["attributes"]["inquiry_num"] == "317-26-1510"


def test_drawing_note_attributes():
    items = li.extract_items([
        "Add COG and Weights to the fan drawing, Inquiry L",
        "Num: 333-25-1622",
        "Add BOM to fan drawing, Inquiry Num: 333-25-1622 L",
        "Plan View on Customer Drawing, Inquiry Num: 333-25-1622 L",
        "ADDITIONAL FEATURES",
        "Static and Dymanic Loads on drawing, Inquiry Num: L 500.00",
        "ADD CG TO FAN DRAWING",
        "SHOW GROUNDING LUGS ON FAN DRAWING",
        "Please update the fan / drawing to show the inlet box rotated 270 degrees from the motor side to the 9 o'clock position.",
        "Add tag to be included on the expansion joint drawing: 1A-EXJ-AXS812",
        "Should add DO NOT STACK and ISPM Stamping to drawing.",
        "Change Subject line on Drawing Transmittal to include tag information as it appears on the order face, Inquiry Num: 364-20-200",
        "Only change on drawings are for CAT purposes added a 1E2999 Symbols",
        "DRAWING AND SPLIT HOUSING AND RELEASED INTO PRODUCTION",
    ])
    by_raw = {it["raw"]: it for it in items}
    weights = by_raw["Add COG and Weights to the fan drawing, Inquiry L"]["attributes"]
    assert weights["inquiry_num"] == "333-25-1622"
    assert weights["drawing_type"] == "WEIGHTS/COG/LOADS"
    assert weights["drawing_scope"] == "FAN"
    assert by_raw["Add BOM to fan drawing, Inquiry Num: 333-25-1622 L"]["attributes"]["drawing_type"] == "BOM"
    assert by_raw["Plan View on Customer Drawing, Inquiry Num: 333-25-1622 L"]["attributes"]["drawing_type"] == "PLAN VIEW/CUSTOMER DRAWING"
    assert by_raw["Static and Dymanic Loads on drawing, Inquiry Num: L 500.00"]["attributes"]["drawing_type"] == "WEIGHTS/COG/LOADS"
    assert by_raw["ADD CG TO FAN DRAWING"]["attributes"]["drawing_type"] == "WEIGHTS/COG/LOADS"
    assert by_raw["SHOW GROUNDING LUGS ON FAN DRAWING"]["attributes"]["drawing_type"] == "GROUNDING/LUGS"
    assert by_raw["Please update the fan / drawing to show the inlet box rotated 270 degrees from the motor side to the 9 o'clock position."]["attributes"]["drawing_type"] == "ORIENTATION/ARRANGEMENT"
    assert by_raw["Add tag to be included on the expansion joint drawing: 1A-EXJ-AXS812"]["attributes"]["drawing_type"] == "TAG/MARKING"
    assert by_raw["Should add DO NOT STACK and ISPM Stamping to drawing."]["attributes"]["drawing_type"] == "PACKAGING/MARKING"
    assert by_raw["Change Subject line on Drawing Transmittal to include tag information as it appears on the order face, Inquiry Num: 364-20-200"]["attributes"]["drawing_type"] == "TAG/MARKING, DRAWING TRANSMITTAL"
    assert by_raw["Only change on drawings are for CAT purposes added a 1E2999 Symbols"]["attributes"]["drawing_type"] == "CUSTOMER SYMBOLS/NOTES"
    split = by_raw["DRAWING AND SPLIT HOUSING AND RELEASED INTO PRODUCTION"]
    assert "SPLIT HOUSING" in split["tags"]
    assert "HOUSING" not in split["tags"]
    assert split["attributes"]["drawing_type"] == "SPLIT HOUSING"


def test_split_housing_attributes():
    items = {it["raw"]: it for it in li.extract_items([
        "Horizontal Split Housing L 3,567.00",
        "Pie Wedge Split Housing L 2,303.00",
        "Shipping Split L 1,000.00",
        "Housing split for shipment L 1,000.00",
    ])}
    horizontal = items["Horizontal Split Housing L 3,567.00"]
    assert horizontal["tags"] == ["SPLIT HOUSING"]
    assert horizontal["attributes"]["component"] == "SPLIT HOUSING"
    assert horizontal["attributes"]["split_type"] == "HORIZONTAL"
    assert items["Pie Wedge Split Housing L 2,303.00"]["attributes"]["split_type"] == "PIE WEDGE"
    assert items["Shipping Split L 1,000.00"]["attributes"]["split_type"] == "SHIPPING"
    assert items["Housing split for shipment L 1,000.00"]["attributes"]["split_type"] == "SHIPPING"


def test_explosion_proof_split_detail_line():
    motor = li.extract_items([
        "Motor C 633.61",
        "Vendor: WEG",
        "Enclosure: Premium Explosion",
        "Proof",
    ])[0]
    assert "MOTOR" in motor["tags"]
    assert "EXPLOSION PROOF" not in motor["tags"]
    assert motor["attributes"]["component"] == "MOTOR"
    assert motor["attributes"]["motor_enclosure"] == "EXPLOSION PROOF"

    warranty = li.extract_items([
        "Exclusive 3 Year Warranty L INC",
        "Enclosure: Premium Explosion",
        "Proof",
    ])[0]
    assert warranty["tags"] == ["WARRANTY"]
    assert "motor_enclosure" not in warranty["attributes"]


def test_flange_scope_motor_and_non_wheel_end():
    motor = li.extract_items(["C-Flange Motor Replate L 100.00"])[0]
    assert "MOTOR" in motor["tags"]
    assert "FLANGE" not in motor["tags"]
    assert motor["attributes"]["component"] == "MOTOR"
    assert motor["attributes"]["motor_mounting"] == "C-FLANGE"
    assert motor["attributes"]["flange_scope"] == "MOTOR"

    housing = li.extract_items(["Non-Wheel End Flange Reinforcing Gussets L 200.00"])[0]
    assert "FLANGE" in housing["tags"]
    assert "HOUSING" in housing["tags"]
    assert "WHEEL" not in housing["tags"]
    assert housing["attributes"]["flange_scope"] == "HOUSING"
    assert housing["attributes"]["flange_location"] == "NON-WHEEL END"


def test_housing_modified_flange_and_length_attributes():
    thickness = li.extract_items(['Housing Flange Thickness, 2-7/8", Inquiry Num: L 250.00'])[0]
    assert {"FLANGE", "HOUSING"} <= set(thickness["tags"])
    assert thickness["attributes"]["flange_scope"] == "HOUSING"
    assert thickness["attributes"]["housing_subcategory"] == "MODIFIED FLANGE"
    assert thickness["attributes"]["housing_feature"] == "FLANGE THICKNESS"

    bolt_pattern = li.extract_items([
        'Drum Housing Bolt Pattern, QTY (12) 9/16" L',
        'Diameter Holes, 34" B.C. Diameter, Straddle',
    ])[0]
    assert "HOUSING" in bolt_pattern["tags"]
    assert bolt_pattern["attributes"]["housing_subcategory"] == "MODIFIED FLANGE"
    assert bolt_pattern["attributes"]["housing_feature"] == "BOLT PATTERN"

    length = li.extract_items(['Housing Length 30-3/4" Flange to Flange, Inquiry L 650.00'])[0]
    assert {"FLANGE", "HOUSING"} <= set(length["tags"])
    assert length["attributes"]["housing_subcategory"] == "DIMENSION NOTE"
    assert length["attributes"]["housing_dimension"] == "FLANGE-TO-FLANGE LENGTH"

    mixing_box = li.extract_items(['Mixing Box Length 30" Flange to Flange L 650.00'])[0]
    assert "HOUSING" not in mixing_box["tags"]
    assert "housing_subcategory" not in mixing_box["attributes"]


def test_housing_construction_and_nameplate_mount_attributes():
    square = li.extract_items(["Housing, Square L 284.00"])[0]
    assert square["attributes"]["housing_subcategory"] == "SQUARE"

    universal = li.extract_items(["Housing, Universal Housing L 75.00"])[0]
    assert universal["attributes"]["housing_subcategory"] == "UNIVERSAL"

    stiffener = li.extract_items(["Housing Stiffner Plates, Inquiry Num: 311-19-2159 L 114.00"])[0]
    assert stiffener["attributes"]["housing_subcategory"] == "STIFFENER PLATES"

    casing = li.extract_items(["Housing, Casing Extension for Outlet L 3,213.00"])[0]
    assert casing["attributes"]["housing_subcategory"] == "CASING EXTENSION"
    assert casing["attributes"]["used_on"] == "OUTLET"

    nameplate = li.extract_items([
        "316 SS Nameplate, Riveted to Housing with 316 SS L 1,179.00",
        "Hardware, Inquiry Num: 253-26-1710",
    ])[0]
    assert {"HOUSING", "MATERIALS", "NAMEPLATE", "STAINLESS STEEL"} <= set(nameplate["tags"])
    assert nameplate["attributes"]["housing_subcategory"] == "NAMEPLATE LOCATION"
    assert nameplate["attributes"]["mount_location"] == "HOUSING"
    assert nameplate["attributes"]["material_scope"] == "NAMEPLATE, HARDWARE"


def test_conduit_box_location_is_motor_not_housing():
    conduit = li.extract_items([
        "mount conduit box as close to the housing as L",
        "possible., Inquiry Num: 311-24-1312",
    ])[0]
    assert "MOTOR" in conduit["tags"]
    assert "HOUSING" not in conduit["tags"]
    assert conduit["attributes"]["component"] == "MOTOR"
    assert conduit["attributes"]["motor_feature"] == "CONDUIT BOX LOCATION"
    assert conduit["attributes"]["motor_conduit_box_location"] == "CLOSE TO HOUSING"

    mixed = li.extract_items([
        "HOUSING FLANGE THICKNESS, DRUM HOUSING BOLT PATTERN AND MOTOR CONDUIT BOX HUGGING HOUSING L 100.00",
    ])[0]
    assert {"HOUSING", "MOTOR", "SPECIAL CONSTRUCTION"} <= set(mixed["tags"])
    assert mixed["attributes"]["component"] == "MOTOR"
    assert mixed["attributes"]["motor_conduit_box_location"] == "HUGGING HOUSING"
    assert mixed["attributes"]["housing_subcategory"] == "MODIFIED FLANGE"
    assert "motor_mounting" not in mixed["attributes"]


def test_packaging_housing_reference_does_not_tag_housing():
    packaging = li.extract_items([
        "For all CAT D/38 Fans: bolted fan to the skid with N 122.00",
        "scrap wood for blocking under the fan housing,",
    ])[0]
    assert "PACKAGING" in packaging["tags"]
    assert "HOUSING" not in packaging["tags"]


def test_flex_connector_flange_scope_and_attrs():
    flex = li.extract_items([
        "Inlet Flanged Expansion Joint L 100.00",
        "Vendor: FlexCom",
        "Product: Expansion Joint",
    ])[0]
    assert "FLEX CONNECTOR" in flex["tags"]
    assert "FLANGE" not in flex["tags"]
    assert flex["attributes"]["component"] == "FLEX CONNECTOR"
    assert flex["attributes"]["flex_connector_type"] == "EXPANSION JOINT"
    assert flex["attributes"]["used_on"] == "INLET"
    assert flex["attributes"]["flange_scope"] == "FLEX CONNECTOR, INLET"


def test_flexible_coupling_is_coupling_subcategory():
    coupling = li.extract_items([
        "Flexible Coupling Falk Type T Steelflex, Size 1070T, Clearance, Horizontal Split Cover T10, Set Screws, CBC Mount L 1,000.00",
    ])[0]
    assert "COUPLING" in coupling["tags"]
    assert "FLEXIBLE COUPLING" not in coupling["tags"]
    attrs = coupling["attributes"]
    assert attrs["component"] == "COUPLING"
    assert attrs["coupling_subcategory"] == "FLEXIBLE COUPLING"
    assert attrs["coupling_type"] == "FALK TYPE T STEELFLEX"
    assert attrs["size"] == "1070T"
    assert attrs["fit"] == "CLEARANCE"
    assert attrs["cover_type"] == "HORIZONTAL SPLIT COVER T10"
    assert attrs["set_screws"] == "YES"
    assert attrs["mounting"] == "CBC MOUNT"


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


def test_material_attributes():
    wheel = li.extract_items([
        "Wheel and Hub, 316 SS Construction, Inquiry Num: L 15,589.00",
    ])[0]
    attrs = wheel["attributes"]
    assert {"MATERIALS", "STAINLESS STEEL", "WHEEL"} <= set(wheel["tags"])
    assert attrs["material"] == "STAINLESS STEEL"
    assert attrs["material_grade"] == "316 SS"
    assert attrs["material_scope"] == "WHEEL AND HUB"

    base = li.extract_items([
        "Base Fan (Base Fan, Design 53 Size F1, 304SST L 9,212.00",
        "Airstream, Arrangement 4 (Less Motor), Inquiry",
        "Num: 333-26-1680)",
    ])[0]
    attrs = base["attributes"]
    assert {"BASE FAN", "MATERIALS", "STAINLESS STEEL"} <= set(base["tags"])
    assert "MOTOR" not in base["tags"]
    assert attrs["material"] == "STAINLESS STEEL"
    assert attrs["material_grade"] == "304 SS"
    assert attrs["material_scope"] == "BASE FAN, AIRSTREAM"

    base_with_motor_detail = li.extract_items([
        "Base Fan L 3,684.00",
        "Motor (Customer Provided) N NA",
        "Vendor: CBC Option",
        "5 HP, 1800 RPM, Enclosure: TEFC",
        "184T, Cast Iron, 3/60/230/460, F1, 1.15 SF",
        "Mounted by Others",
    ])[0]
    assert base_with_motor_detail["tags"] == ["BASE FAN"]
    assert "vendor" not in base_with_motor_detail["attributes"]

    alum = li.extract_items(["Wheel, Aluminum (AMCA B) L INC"])[0]
    attrs = alum["attributes"]
    assert {"ALUMINUM", "MATERIALS", "WHEEL"} <= set(alum["tags"])
    assert attrs["material"] == "ALUMINUM"
    assert attrs["material_scope"] == "WHEEL"

    grades = li.extract_items([
        "Housing and Base, 304L SS / 316L SS Construction L 1,200.00",
    ])[0]
    attrs = grades["attributes"]
    assert attrs["material_grade"] == "304L SS, 316L SS"
    assert attrs["material_scope"] == "HOUSING AND BASE"


def test_material_scope_requires_material():
    report = li.extract_items(["CBC Runtest Results and Wheel Balance Report, L"])[0]
    assert "BALANCE" not in report["tags"]
    assert "material_scope" not in report["attributes"]


def test_balance_attributes():
    g25 = li.extract_items(["G2.5 Balance L 341.00"])[0]
    assert "BALANCE" in g25["tags"]
    assert g25["attributes"]["balance_type"] == "GRADED BALANCE"
    assert g25["attributes"]["balance_grade"] == "G2.5"

    g10 = li.extract_items([
        "Special Product Design (EMD Wheel PN 40214482 N 4,393.00",
        "(CBC PN 08-5-4281) Includes G1.0 Balance on",
        "clearance fit arbor, Inquiry Num: 410-13-237)",
    ])[0]
    assert {"BALANCE", "WHEEL"} <= set(g10["tags"])
    assert g10["attributes"]["balance_type"] == "GRADED BALANCE"
    assert g10["attributes"]["balance_grade"] == "G1.0"

    welded = li.extract_items([
        "Welded Balance Weights, Inquiry Num: 339-26-2098 L 126.00",
    ])[0]
    assert "BALANCE" in welded["tags"]
    assert welded["attributes"]["balance_type"] == "WELDED BALANCE WEIGHTS"
    assert "balance_grade" not in welded["attributes"]


def test_low_leakage_and_extreme_temperature_attributes():
    low_leak = li.extract_items([
        "Inlet Volume Control, Low Leak, Automatic L 3,639.00",
        "Size 4014 Low Leak IVC with rotating ring arm only",
    ])[0]
    assert {"INLET VANES", "LOW LEAKAGE"} <= set(low_leak["tags"])
    assert low_leak["attributes"]["leakage_class"] == "LOW LEAKAGE"
    assert low_leak["attributes"]["used_on"] == "IVC"

    cold_fan = li.extract_items([
        "Base Fan Suitable for -45 C Temperature L 21,101.00",
    ])[0]
    assert {"BASE FAN", "EXTREME TEMP"} <= set(cold_fan["tags"])
    assert cold_fan["attributes"]["temperature_service"] == "LOW TEMPERATURE"
    assert cold_fan["attributes"]["temperature_rating"] == "-45C"

    hot_ivc = li.extract_items([
        "Fan and Low Leakage IVC suitable for 175F, Inquiry L",
        "Num: 9-14-164",
    ])[0]
    assert {"EXTREME TEMP", "LOW LEAKAGE", "INLET VANES"} <= set(hot_ivc["tags"])
    assert hot_ivc["attributes"]["temperature_service"] == "HIGH TEMPERATURE"
    assert hot_ivc["attributes"]["temperature_rating"] == "175F"
    assert hot_ivc["attributes"]["leakage_class"] == "LOW LEAKAGE"

    motor_grease = li.extract_items([
        "Motor with High Temperature Grease C 1,000.00",
    ])[0]
    assert "MOTOR" in motor_grease["tags"]
    assert "EXTREME TEMP" not in motor_grease["tags"]
    assert motor_grease["attributes"]["component"] == "MOTOR"
    assert motor_grease["attributes"]["grease_type"] == "HIGH TEMPERATURE GREASE"

    shaft_seal = li.extract_items([
        "John Crane Double Carbon Shaft Seal, High Temp N 6,040.00",
    ])[0]
    assert "SHAFT SEAL" in shaft_seal["tags"]
    assert "EXTREME TEMP" not in shaft_seal["tags"]
    assert shaft_seal["attributes"]["component"] == "SHAFT SEAL"
    assert shaft_seal["attributes"]["temperature_service"] == "HIGH TEMPERATURE"


def test_heavy_duty_component_attributes():
    wheel = li.extract_items(["Wheel, Heavy Duty L 302.00"])[0]
    assert {"HEAVY DUTY", "WHEEL"} <= set(wheel["tags"])
    assert wheel["attributes"]["component"] == "WHEEL"
    assert wheel["attributes"]["duty_rating"] == "HEAVY DUTY"

    housing = li.extract_items(["Housing, Heavy Duty L 4,754.00"])[0]
    assert {"HEAVY DUTY", "HOUSING"} <= set(housing["tags"])
    assert housing["attributes"]["component"] == "HOUSING"
    assert housing["attributes"]["duty_rating"] == "HEAVY DUTY"

    severe_motor = li.extract_items(["Motor TEFC Severe Duty C 1,200.00"])[0]
    assert "HEAVY DUTY" not in severe_motor["tags"]
    assert "duty_rating" not in severe_motor["attributes"]


def test_bearing_attributes():
    repair = li.extract_items([
        'Repair Bearings (Pair), 2-3/16" Bore (Qty: 1), N 552.00 110.00',
        "Inquiry Num: 340-25-943RP",
    ])[0]
    assert "BEARINGS" in repair["tags"]
    assert repair["attributes"]["bearing_type"] == "REPAIR BEARINGS"
    assert repair["attributes"]["bearing_bore"] == '2-3/16"'
    assert repair["attributes"]["inquiry_num"] == "340-25-943RP"

    split = li.extract_items(["Bearings, Split Pillow Block L 2,387.00"])[0]
    assert "BEARINGS" in split["tags"]
    assert split["attributes"]["bearing_type"] == "SPLIT PILLOW BLOCK"

    spare = li.extract_items(["Spare Bearings, Inquiry Num: 340-26-1112 L 3,583.00"])[0]
    assert "BEARINGS" in spare["tags"]
    assert spare["attributes"]["bearing_type"] == "SPARE BEARINGS"
    assert spare["attributes"]["inquiry_num"] == "340-26-1112"

    adder = li.extract_items(["Bearing ADDER for 200,00 hours, Inquiry Num: L"])[0]
    assert "BEARINGS" in adder["tags"]
    assert adder["attributes"]["bearing_type"] == "BEARING ADDER"


def test_shaft_bearing_guard_tag_uses_primary_line():
    guard = li.extract_items(["Shaft and Bearing Guard, Painted Safety Yellow L 949.00"])[0]
    assert "SHAFT/BEARING/COUPLING GUARD" in guard["tags"]

    grease = li.extract_items([
        "Extended Grease Fittings - Shipped Loose, Inquiry L 349.00",
        "Num: 6-5-1551",
        "Size 16-1/2 Shaft and Bearing Guard, Inquiry Num: L -940.00",
    ])[0]
    assert "BEARINGS" in grease["tags"]
    assert "MOTOR" not in grease["tags"]
    assert "EXTENDED LUBE" not in grease["tags"]
    assert "SHAFT/BEARING/COUPLING GUARD" not in grease["tags"]


def test_lube_accessories_map_to_bearing_or_motor():
    assert "EXTENDED LUBE" not in li.load_rules(refresh=True)["tags"]

    unknown = li.extract_items(["Extended Lube Lines with Zerk Fittings L 100.00"])[0]
    assert {"BEARINGS", "MOTOR"} <= set(unknown["tags"])

    motor = li.extract_items(["Motor Grease Lines L 100.00"])[0]
    assert "MOTOR" in motor["tags"]
    assert "BEARINGS" not in motor["tags"]

    bearing = li.extract_items(["Extended Grease Leads to Fan Bearings L 100.00"])[0]
    assert "BEARINGS" in bearing["tags"]
    assert "MOTOR" not in bearing["tags"]

    both = li.extract_items(["Motor Bearing Grease Fittings L 100.00"])[0]
    assert {"BEARINGS", "MOTOR"} <= set(both["tags"])

    items = {it["norm"]: it for it in li.extract_items([
        "Motor C 1,000.00",
        "Extended Grease Leads",
        "Extended Grease Leads L 100.00",
    ])}
    assert "MOTOR" in items["MOTOR"]["tags"]
    assert "BEARINGS" not in items["MOTOR"]["tags"]
    assert {"BEARINGS", "MOTOR"} <= set(items["EXTENDED GREASE LEADS"]["tags"])


def test_assembly_note_is_misc_note_not_component():
    assembly = li.extract_items([
        "Assembly, Assemble Panel, Wheel, Shaft, and L INC",
        "Bearings",
    ])[0]
    assert "MISC NOTE" in assembly["tags"]
    assert "WHEEL" not in assembly["tags"]
    assert "BEARINGS" not in assembly["tags"]
    assert assembly["attributes"]["note_type"] == "ASSEMBLY"


def test_paint_line_does_not_become_component_tags():
    paint = li.extract_items([
        "Paint: Interior, Wheel, Exterior, Motor Base, L 983.00",
        "Channel Base and Bearing Base",
    ])[0]
    assert paint["tags"] == ["COATING"]
    assert paint["attributes"]["coating_context"] == "FAN"
    assert paint["attributes"]["coating_category"] == "PAINT"
    assert "WHEEL" not in paint["tags"]

    paint_with_warranty_detail = li.extract_items([
        "Paint: Exterior, Motor Base L 714.00",
        "Extended Warranty: Chicago Blower standard warranty extended to 24 months.",
    ])[0]
    assert paint_with_warranty_detail["tags"] == ["COATING"]

    pre_coating = li.extract_items(["Pre-Coating Assembly/Disassembly L 1,185.00"])[0]
    assert pre_coating["tags"] == ["COATING"]
    assert pre_coating["attributes"]["coating_category"] == "PRE-COATING PROCESS"
    assert pre_coating["attributes"]["coating_process"] == "PRE-COATING ASSEMBLY/DISASSEMBLY"

    special = li.extract_items([
        "Special Paint, exterior and airstream: SSPC-SP10, N 5,310.00 220.00",
        "Zinc Rich Epoxy primer, 3 mils, mid coat epoxy 7 mil, top coat poly, 3-mil",
        "(RAL 6019 if not available use RAL 7035), Inquiry Num: 253-24-1651",
    ])[0]
    assert special["attributes"]["coating_category"] == "SPECIAL COATING"
    assert special["attributes"]["coating_type"] == "EPOXY"
    assert special["attributes"]["coating_color"] == "RAL 6019"
    assert special["attributes"]["alternate_coating_color"] == "RAL 7035"

    veg_oil = li.extract_items([
        "Airstream to be unpainted and Coated with L 714.00",
        "vegetable oil, Inquiry Num: 317-26-1059",
    ])[0]
    assert veg_oil["attributes"]["coating_category"] == "SPECIAL COATING"
    assert veg_oil["attributes"]["coating_type"] == "VEGETABLE OIL"
    assert veg_oil["attributes"]["coating_state"] == "UNPAINTED"

    unpainted = li.extract_items(["Wheel, Cast Aluminum CCW (Unpainted) (Bore: L 739.00"])[0]
    assert unpainted["attributes"]["coating_category"] == "UNPAINTED"

    note = li.extract_items([
        "Specification and updated coating Note # B to show wheel is un-coated L 100.00",
    ])[0]
    assert note["attributes"]["coating_category"] == "COATING NOTE"


def test_accessory_coating_is_attribute_not_fan_coating():
    belt = li.extract_items([
        "Belt Guard, Painted Safety Yellow L 1,581.00",
        "CBC Mount",
        "Tach Hole in Guard: with Plug",
        "Fan End",
        "Motor End",
    ])[0]
    assert belt["tags"] == ["BELT GUARD"]
    attrs = belt["attributes"]
    assert attrs["component"] == "BELT GUARD"
    assert attrs["coating_context"] == "ACCESSORY"
    assert attrs["coating_scope"] == "BELT GUARD"
    assert attrs["coating_color"] == "SAFETY YELLOW"
    assert attrs["mounting"] == "CBC MOUNT"
    assert attrs["tach_hole"] == "WITH PLUG"
    assert attrs["tach_hole_location"] == "FAN END, MOTOR END"

    shaft_guard = li.extract_items(["Shaft and Bearing Guard, Painted Safety Yellow L 949.00"])[0]
    assert "SHAFT/BEARING/COUPLING GUARD" in shaft_guard["tags"]
    assert "COATING" not in shaft_guard["tags"]
    assert shaft_guard["attributes"]["coating_context"] == "ACCESSORY"


def test_passivation_is_stainless_treatment_not_coating():
    passivation = li.extract_items([
        "Passivation of Welds (Passivation of Welds L 1,412.00",
        "(Airstream and Exterior), Inquiry Num: 333-26-1234",
    ])[0]
    assert "COATING" not in passivation["tags"]
    assert {"MATERIALS", "STAINLESS STEEL"} <= set(passivation["tags"])
    attrs = passivation["attributes"]
    assert attrs["material"] == "STAINLESS STEEL"
    assert attrs["material_treatment"] == "PASSIVATION OF WELDS"
    assert attrs["material_scope"] == "AIRSTREAM, EXTERIOR, WELDS"


def test_accessory_coating_does_not_make_silencer_fan_coating():
    silencer = li.extract_items([
        "Inlet Silencer for 85 dBA @ 3' Outdoors, Model C 2,267.00 363.00",
        "includes Leg, 2 coats of paint",
        "Vendor: VAW Systems",
        "Product: Inlet Silencer",
    ])[0]
    assert "SILENCER" in silencer["tags"]
    assert "COATING" not in silencer["tags"]
    assert silencer["attributes"]["coating_context"] == "ACCESSORY"
    assert silencer["attributes"]["coating_scope"] == "SILENCER"


def test_admin_notes_are_misc_notes():
    note = li.extract_items([
        "Slow Pay Addition N 339.00",
        "3-coat paint system to match fan.",
    ])[0]
    assert note["tags"] == ["MISC NOTE"]
    assert note["attributes"] == {"note_type": "ADMIN"}


def test_component_materials_do_not_count_as_fan_materials():
    actuator = li.extract_items([
        "Bettis #RPED150 Double Acting Pneumatic Actuator C 6,537.14 1,046.00",
        "Fisher #67CFR Filter/ Regulator, SS Tubing.",
        "Vendor: Novaspect",
        "Product: Actuator",
    ])[0]
    attrs = actuator["attributes"]
    assert "ACTUATOR" in actuator["tags"]
    assert not {"MATERIALS", "STAINLESS STEEL"} & set(actuator["tags"])
    assert attrs["component"] == "ACTUATOR"
    assert attrs["component_material"] == "STAINLESS STEEL"
    assert attrs["component_material_scope"] == "TUBING"
    assert "material" not in attrs

    motor = li.extract_items([
        "Motor (Multimounting IE3 5.5 HP 2P 112M 3Ph C 620.49",
        "112M, Cast Aluminum, Foot Mounted",
        "Vendor: Weg",
    ])[0]
    attrs = motor["attributes"]
    assert "MOTOR" in motor["tags"]
    assert not {"MATERIALS", "ALUMINUM"} & set(motor["tags"])
    assert attrs["component"] == "MOTOR"
    assert attrs["component_material"] == "ALUMINUM"
    assert attrs["component_material_scope"] == "MOTOR"
    assert "material" not in attrs


def test_data_branch_noise_skipped():
    lines = [
        "ADDITIONAL FEATURES",
        "Product:",
        "Product 7,623.00",
        "Prints",
        "Prints:",
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
    paymode = next(v for n, v in items.items() if "PAYMODE" in n)
    assert paymode["tags"] == ["MISC NOTE"]
    assert paymode["attributes"]["note_type"] == "ADMIN"
    for bad in ("PRODUCT", "PRINTS", "WARRANTY", "SHIPPING NOTES", "DO NOT STACK"):
        assert not any(bad in n for n in items), items


def test_performance_header_noise_skipped():
    items = li.extract_items([
        "MAX TEMP, ELEVATION, DENSITY, COLD START BHP, MOTOR DESCRIPTION, FRAME, HP, INLET BOX, MIXING BOX, FRESH AIR",
    ])
    assert items == []


def test_used_on_requires_damper_context():
    flange = li.extract_items(["Outlet, Flanged, Punched L STD"])[0]
    assert "used_on" not in flange["attributes"]
    damper = li.extract_items(["Outlet Damper, Opposed L 100.00"])[0]
    assert damper["attributes"]["used_on"] == "OUTLET DAMPER"
    discharge = li.extract_items(["Actuator for Discharge Damper, Bettis #RPED200 C 5,192.00"])[0]
    assert discharge["attributes"]["used_on"] == "OUTLET DAMPER"
    fresh_air = li.extract_items(["Change Actuator Location for the FA damper to be on motor side of fan L 100.00"])[0]
    assert fresh_air["attributes"]["used_on"] == "FRESH AIR DAMPER"
    assert "used_on_review" not in fresh_air["attributes"]
    prespin = li.extract_items(["Actuator for Pre-spin Damper, Bettis #RPED200 C 5,192.00"])[0]
    assert prespin["attributes"]["used_on"] == "PRESPIN DAMPER"

    mounted_damper = li.extract_items([
        "Fresh Air Damper, Opposed Blade, Without Stuffing L 13,946.00",
        "Boxes (Shipped Loose), Mounted on Oversized Inlet Box, Automatic",
        "Product: Actuator",
        "Vendor: Bettis RPD",
    ])[0]
    assert "DAMPER" in mounted_damper["tags"]
    assert "ACTUATOR" in mounted_damper["tags"]
    assert "INLET" not in mounted_damper["tags"]
    assert mounted_damper["attributes"]["used_on"] == "FRESH AIR DAMPER"
    assert mounted_damper["attributes"]["mount_location"] == "INLET BOX"


def test_without_ivc_does_not_tag_inlet_vanes_or_used_on():
    item = li.extract_items(["Inlet, Flanged, Punched (without IVC) L 468.00"])[0]
    assert {"INLET", "FLANGE"} <= set(item["tags"])
    assert "INLET VANES" not in item["tags"]
    assert "used_on" not in item["attributes"]


def test_inlet_cone_width_does_not_tag_wheel_without_wheel_wording():
    item = li.extract_items(["Aluminum Inlet Cone, Inlet Cone 100% Width Cartoned (Unpainted) L 752.00"])[0]
    assert {"INLET", "MATERIALS", "ALUMINUM"} <= set(item["tags"])
    assert "WHEEL" not in item["tags"]
    assert item["attributes"]["material_scope"] == "INLET CONE"
    assert item["attributes"]["coating_scope"] == "INLET"


def test_inlet_subcategory_attributes():
    open_inlet = li.extract_items(["Inlet, Open L STD"])[0]
    assert open_inlet["attributes"]["inlet_subcategory"] == "OPEN"

    slip = li.extract_items(["Inlet, Slip N STD"])[0]
    assert slip["attributes"]["inlet_subcategory"] == "SLIP"

    bell = li.extract_items(["Inlet, Bell L 656.00"])[0]
    assert bell["attributes"]["inlet_subcategory"] == "BELL"

    tube = li.extract_items(["Inlet, Tube L STD"])[0]
    assert tube["attributes"]["inlet_subcategory"] == "TUBE"

    cone = li.extract_items([
        "INLET CONE 72%, D62 SW, SIZE 200 (Customer N 102.00 10.00",
        "Part #: 90515, CBC Part #: 62-0-0062)",
    ])[0]
    assert cone["attributes"]["inlet_subcategory"] == "INLET CONE"
    assert cone["attributes"]["inlet_cone_width_percent"] == "72"
    assert cone["attributes"]["inlet_size"] == "200"


def test_inlet_flange_direction_and_box_attributes():
    punched = li.extract_items(["Inlet, Flanged, Punched (without IVC) L 468.00"])[0]
    assert punched["attributes"]["inlet_subcategory"] == "FLANGED/PUNCHED"
    assert punched["attributes"]["inlet_feature"] == "PUNCHED"
    assert punched["attributes"]["ivc_relation"] == "WITHOUT IVC"

    with_ivc = li.extract_items(["Inlet, Flanged, Punched (with IVC) L 1,453.00"])[0]
    assert {"INLET", "INLET VANES"} <= set(with_ivc["tags"])
    assert with_ivc["attributes"]["ivc_relation"] == "WITH IVC"
    assert with_ivc["attributes"]["used_on"] == "IVC"

    bolted = li.extract_items(['Inlet, Flanged, Standard Bolted (Inlet Dia 10") N STD'])[0]
    assert bolted["attributes"]["inlet_subcategory"] == "FLANGED"
    assert bolted["attributes"]["inlet_feature"] == "STANDARD BOLTED"

    direction = li.extract_items(["Inlet Direction: Vertical Inlet Down L STD"])[0]
    assert direction["attributes"]["inlet_subcategory"] == "DIRECTION"
    assert direction["attributes"]["inlet_direction"] == "VERTICAL INLET DOWN"

    box = li.extract_items([
        "Inlet Box, Bolt-on (Shipped Loose), Oversized Inlet L 7,161.00",
        "Box Size 300 (Inlet Box, Bolt-on (Shipped Loose),",
        "Bolt-On Inlet Box Position: @ 0",
    ])[0]
    assert box["attributes"]["inlet_subcategory"] == "INLET BOX"
    assert box["attributes"]["inlet_box_type"] == "BOLT-ON"
    assert box["attributes"]["inlet_box_size"] == "300"
    assert box["attributes"]["inlet_box_position"] == "0"
    assert box["attributes"]["shipping_state"] == "SHIPPED LOOSE"


def test_mixing_box_is_not_inlet_category():
    mixing = li.extract_items([
        "Mixing Box with Flanged FGR Port (Shipped Loose) L 8,297.00",
        "(Mixing Box With Flanged FGR Port (Shipped Loose), Suitable for Size 300 Inlet Box)",
    ])[0]
    assert "MIXING BOX" in mixing["tags"]
    assert "INLET" not in mixing["tags"]
    assert mixing["attributes"]["component"] == "MIXING BOX"
    assert mixing["attributes"]["flange_scope"] == "MIXING BOX"
    assert mixing["attributes"]["mixing_box_feature"] == "FGR PORT, FLANGED"
    assert mixing["attributes"]["used_on"] == "INLET BOX"
    assert mixing["attributes"]["used_on_size"] == "300"


def test_drain_attributes():
    housing = li.extract_items(["Housing Drain with Plug, 304 SS Construction L 328.00"])[0]
    attrs = housing["attributes"]
    assert "HOUSING" not in housing["tags"]
    assert attrs["drain_type"] == "HOUSING DRAIN"
    assert attrs["drain_closure"] == "PLUG"
    assert attrs["material"] == "STAINLESS STEEL"
    assert attrs["material_grade"] == "304 SS"
    assert attrs["material_scope"] == "HOUSING DRAIN"

    housing_plural = li.extract_items(["Housing Drains with Plugs, Inquiry Num: 333-26-1234 L 252.00"])[0]
    assert housing_plural["tags"] == ["DRAIN"]
    assert housing_plural["attributes"]["drain_type"] == "HOUSING DRAIN"

    inlet = li.extract_items(['Inlet Box, Drain Plug 3/4" Diameter L 51.00'])[0]
    assert "INLET" not in inlet["tags"]
    assert inlet["attributes"]["drain_type"] == "INLET BOX DRAIN"
    assert inlet["attributes"]["drain_closure"] == "PLUG"

    motor = li.extract_items([
        "Motor (CEM3711T-10hp motor C 618.96 25.00",
        "M7A Add Condensation Drain Holes - Vertical Shaft",
        "Down)",
    ])[0]
    assert motor["attributes"]["drain_type"] == "MOTOR CONDUIT BOX DRAIN"
    assert motor["attributes"]["drain_detail"] == "CONDENSATION DRAIN HOLES"

    motor_ss = li.extract_items(["Motor (Conduit Box Drain, 304 SS Construction C 100.00"])[0]
    attrs = motor_ss["attributes"]
    assert attrs["drain_type"] == "MOTOR CONDUIT BOX DRAIN"
    assert attrs["component_material"] == "STAINLESS STEEL"
    assert attrs["component_material_grade"] == "304 SS"
    assert attrs["component_material_scope"] == "MOTOR CONDUIT BOX DRAIN"


def test_standalone_actuator_used_on_review():
    actuator = li.extract_items([
        "Bettis #RPED150 Double Acting Pneumatic Actuator C 6,537.14 1,046.00",
        "Vendor: Novaspect",
        "Product: Actuator",
    ])[0]
    attrs = actuator["attributes"]
    assert attrs["component"] == "ACTUATOR"
    assert "used_on" not in attrs
    assert attrs["used_on_review"] == "INCONCLUSIVE - INLET/OUTLET/PRESPIN/IVC"


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
    assert any("ZERK FITTINGS" in n for n in norms), norms     # section capture
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
