"""Sales-order line items: capture, normalize, store, and search.

Every parsed Sales Order yields its LINE ITEMS — the priced item/accessory
lines and the "Additional Features"-style feature lines that describe what was
actually sold (shaft seals, spark-resistant construction, coatings, isolators,
special motors, ...). The same option is rarely written the same way twice
("SS SHAFT SLEEVE" / "Stainless Steel Shaft Sleeve" / "316SS sleeve"), so each
captured line is kept several ways:

    raw      the line exactly as printed on the Sales Order (never altered)
    norm     a normalized form: uppercased, item numbers / qty / the price
             columns and trailing L/C/N type letter stripped, punctuation
             collapsed, known abbreviations expanded (W/ -> WITH, SS ->
             STAINLESS STEEL, IVD -> INLET VANE DAMPER, ...) so spelling
             variants of the same option converge
    details  the unpriced continuation lines printed under the item (vendor,
             motor HP/enclosure, "Product: Damper", ...) — searchable, and
             they contribute to the item's tags
    tags     canonical feature tags matched by the rules table (SHAFT SEAL,
             SPARK RESISTANT, COATING, ...) — the lookup vocabulary

CBC Sales-Order anatomy (fitted against real dumps, jobs 421314/421473): the
item table lists "<description> <L|C|N> <Price Freight Markup Net Comm.>" —
the type letter then up to five money columns, or STD / INC / N/C in place of
a price, or nothing at all. Unpriced lines between two items are that item's
detail block. Page furniture (the "Chicago Blower Corporation ... (cont.)"
headers, v-version footers, ref-number rows), the totals/freight/tax block,
the Sold To / commission / terms front matter, the drawings-distribution
checklist, and the CO-history lines are all excluded by skip rules.

Everything lands in one resumable JSON store (LINE_ITEMS_STORE, default
BACKLOG_DIR/line_items.json), keyed by job:

    {"jobs":    {"421314": {"customer": ..., "co_number": 1, "so_pdf": ...,
                            "scanned_at": ..., "items": [{raw, norm, qty,
                            price, ptype, section, details, tags}, ...]}},
     "ai_tags": {"<norm>": ["TAG", ...]}}   # cached Claude classifications

The store is fed three ways: the daily run (sales_orders.py) records every
board job it parses, backfill_orders.py records each historical order, and
line_items_scan.py walks the already-archived PDFs under SALES_ORDER_DIR.
Search it with find_orders.py. Because `raw`/`details` are stored verbatim,
the normalization/tag rules can be tuned any time and re-applied with
`python line_items_scan.py --renorm` — nothing is ever lost.

Rules are data, not code: DEFAULT_RULES below seeds the vocabulary, and an
optional JSON file (env LINE_ITEM_RULES) EXTENDS it — new abbreviations, skip
patterns, and tag patterns merge over the defaults, so site-specific wording
can be added without touching code (and survives updates).

This module is pure logic (no pdfplumber/playwright) so it stays unit-testable
off the work machine; callers hand it the reconstructed text lines of a PDF.
"""
from __future__ import annotations

import json
import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, List, Tuple

from config import BACKLOG_DIR, LINE_ITEM_RULES, LINE_ITEMS_STORE

log = logging.getLogger(__name__)

# --------------------------------------------------------------------------- #
# Rules (defaults; an optional LINE_ITEM_RULES json file extends these)        #
# --------------------------------------------------------------------------- #
DEFAULT_RULES: Dict[str, Any] = {
    # Token -> expansion, applied during normalization (token-wise, after
    # trailing punctuation is stripped; keys are uppercase).
    "abbreviations": {
        "W/": "WITH", "W/O": "WITHOUT", "C/W": "COMPLETE WITH",
        "SS": "STAINLESS STEEL", "S/S": "STAINLESS STEEL",
        "SST": "STAINLESS STEEL", "STL": "STEEL",
        "GALV": "GALVANIZED", "ALUM": "ALUMINUM",
        "ALUMINIUM": "ALUMINUM",
        "CONST": "CONSTRUCTION", "CONSTR": "CONSTRUCTION",
        "ARR": "ARRANGEMENT", "ARRG": "ARRANGEMENT", "ARRGT": "ARRANGEMENT",
        "ASSY": "ASSEMBLY", "MTR": "MOTOR", "TEMP": "TEMPERATURE",
        "FA": "FRESH AIR",
        "BRG": "BEARING", "BRGS": "BEARINGS", "SLV": "SLEEVE",
        "HSG": "HOUSING", "WHL": "WHEEL", "CONN": "CONNECTOR",
        "HVY": "HEAVY", "DTY": "DUTY", "HD": "HEAVY DUTY",
        "XP": "EXPLOSION PROOF", "EXPL": "EXPLOSION",
        "IVD": "INLET VANE DAMPER",
    },
    # Lines matching any of these are never items (case-insensitive regexes):
    # page furniture, totals/charges, address & terms front matter, the
    # drawings-distribution checklist, and the change-order history (captured
    # separately as co_history). Fitted against real CBC SO dumps.
    "skip_patterns": [
        r"^\s*c\s*/?\s*o\s*#?\s*\d",                  # CO#1 ... change-order notes
        r"^chicago\s+blower\b",                       # page header / (cont.) header
        r"^v\d+\.\d+",                                # "v1.8.1.5 -1-" version footer
        r"^(sub\s*)?total\b", r"total\s+(billing|price|order|commission)",
        r"^amount\s+due", r"^list\s+total\b", r"deduction",
        r"^freight\b", r"^fob\b", r"^(sales\s+)?tax\b", r"surcharge",
        r"^lead\s+time\b", r"^customs\s+invoice\b",
        r"^type\s+price\b",                           # "Type Price Freight Markup Net Comm."
        r"^max\s+temp,\s+elevation,\s+density\b",     # fan performance table header
        r"^page\s+\d+", r"^\d+\s+of\s+\d+$",
        r"^sales\s+order\b", r"^order\s*#", r"^order\s+(no|number|date)\b",
        r"^job\s+(no|number)\b",
        r"^sold\s+to\b", r"^ship\s+to\b", r"^bill\s+to\b",
        r"^customer\s+(po|p\.o\.|order|#|contact)", r"^p\.?\s*o\.?\s*(no|number|#)",
        r"^terms\b", r"^net\s+\d+", r"^quote\b", r"^quotation\b",
        r"^phone\b", r"^fax\b", r"^www\.", r"\w@\w",  # \w@\w = emails, not "@9:00"
        r"^date\b", r"^entered\s+by\b", r"^salesman\b", r"^rep\b",
        r"^sales\s+office\b", r"^splits?\b",
        r"list\s+price", r"multiplier", r"^unit\s+price", r"^price\s+each",
        r"^discount\b", r"^see\s+(additional|special)\b",
        r"^warranty\b", r"^commission\s+override\b",
        r"\b(charge|fee)\b",
        r"^prints?\s*:?\s*$", r"^product\s*:?\s*$", r"^product\s+\$?\d",
        # Shipping/admin notes can appear inside Additional Features / Notes;
        # keep them out of the item inventory while preserving real ship-loose
        # priced rows.
        r"^additional\s+shipping\s+notes?\b", r"^traffic\s+note\b",
        r"shipping\s+barcode", r"^(?!.*\bdrawings?\b).*do\s+not\s+stack", r"no\s+metal\s+banding",
        r"^order\s+is\s+shipping\s+overseas\b", r"^customer\s+broker\b",
        r"^standard\s+address\b", r"^https?\b", r"^ispm\s+wood\b",
        r"^(last\s+choice\s+)?(fedex|ups)\b", r"^please\s+send\b",
        r"^if\s+the\s+shipment\b", r"^lbs\.?\s+contact\b",
        r"^above\s+email\s+after\b", r"^as\s+well\b",
        r"^appointment\s+required\b", r"^for\s+(orders|packages)\b",
        r"^with\s+a\s+ship\s+date\b", r"^contact\s*$",
        r"^must\s+(appear|be)\b",
        r"^reference\s+sn\b", r"\bppap\b", r"^kindly\s+note\b",
        r"^please\s+confirm\b",
        r"^be\s+necessary\b.*\bfor\s+more\s+information\b",
        # Drawings-distribution checklist (trails the Notes section).
        r"^fan\s+drawings?\b", r"^motor\s+prints?\b", r"^motor\s+data\s+sheets?\b",
        r"^buyout\s+prints?\b", r"^emailed\b", r"^mailed\b", r"^o\s*&\s*m\b",
        r"^other\s*$", r"^special\s+invoicing\b",
    ],
    # A short line matching one of these opens a feature section: every
    # following line is captured (priced or not) until a section_end marker.
    "section_start": [
        r"^additional\s+features?\b", r"^special\s+features?\b",
        r"^accessor(y|ies)\b", r"^features?\s*:?\s*$", r"^options?\s*:?\s*$",
        r"^includes?\s*:?\s*$", r"^scope\s+of\s+supply\b",
    ],
    # ...and one of these closes it.
    "section_end": [
        r"^notes?\b", r"^tag\b", r"^nameplate\b", r"^total\b",
        r"^sold\s+to\b", r"^ship\s+to\b", r"^terms\b",
        r"^approval\b", r"^signature\b", r"^change\s+order",
        r"^fan\s+drawings?\b", r"^special\s+invoicing\b",
    ],
    # Canonical feature tags: tag -> regexes matched against the NORMALIZED
    # text + details of an item. An item can carry several.
    "tags": {
        "BASE FAN": [r"^base\s+fan\b"],
        "ACTUATOR": [r"\bactuator\b"],
        "SPARK RESISTANT": [r"spark", r"\bAMCA\s+(?:TYPE\s+)?[ABC]\b"],
        "SHAFT SEAL": [r"shaft\s*seal", r"stuffing\s*box", r"lip\s*seal",
                       r"ceramic\s*felt"],
        "SHAFT SLEEVE": [r"shaft\s*sleeve"],
        "SHAFT COOLER": [r"shaft\s*cooler", r"heat\s*slinger"],
        "MATERIALS": [r"stainless", r"\b304L?\b", r"\b316L?\b", r"passivat",
                      r"aluminum", r"aluminium"],
        "STAINLESS STEEL": [r"stainless", r"\b304L?\b", r"\b316L?\b", r"passivat"],
        "ALUMINUM": [r"aluminum", r"aluminium"],
        "EXTREME TEMP": [
            r"high\s*temp", r"heat\s*fan", r"low\s*temp",
            r"\bsuitable\s+for\s+-?\d{2,3}\s*(?:F|C)\b",
            r"\brated\s+for\s+\d{3}\s*F\b",
            r"\bto\s*-?\d{2,3}\s*(?:deg(?:ree)?\s*)?C\b",
            r"\b-\s*\d{2,3}\s*C\s+temperature\b",
        ],
        "HEAVY DUTY": [r"heavy\s*duty"],
        "LOW LEAKAGE": [r"low\s*[- ]?leak(?:age)?"],
        "COATING": [r"epoxy", r"\bcoat(?:ed|ing|s)?\b", r"paint", r"primer",
                    r"plasite", r"heresite", r"\bzinc\b"],
        "LINING": [r"rubber\s*lin", r"\bliners?\b", r"\blined\b", r"\blining\b", r"abrasion",
                   r"firmex"],
        "INSULATION": [r"insulat",
                       r"\b(?:housing|fan)\b.*\b(lagging|jacket(?:ed)?|mineral\s+wool|fiberglass|fibre\s*glass)\b",
                       r"\b(lagging|jacket(?:ed)?|mineral\s+wool|fiberglass|fibre\s*glass)\b.*\b(?:housing|fan)\b"],
        "VIBRATION ISOLATION": [r"isolat", r"rubber[\s-]*in[\s-]*shear",
                                r"\bRIS\b", r"spring\s*mount", r"seismic",
                                r"vibration\s*base"],
        "VIBRATION SWITCH": [r"vibration\s*(switch|detector|monitor|sensor)"],
        "DAMPER": [r"damper", r"backdraft", r"volume\s*control"],
        "INLET VANES": [r"inlet\s*vane", r"\bVIV\b", r"\bIVC\b",
                        r"variable\s*inlet", r"inlet\s*volume\s*control"],
        "INLET": [r"\binlet\s+(open|slip|bell|cone|box|tube|flanged|punched)",
                  r"\binlet\s+direction"],
        "MIXING BOX": [r"\bmixing\s+box\b"],
        "OUTLET": [r"\boutlet\s+(open|slip|flanged|punched|pressure\s*tap|volume\s*control)",
                   r"\bdischarge\s+elbow"],
        "WHEEL": [r"\bwheel\b", r"\bpercent\s+width\b", r"%\s*width\b"],
        "HOUSING": [r"\bhousing\b"],
        "SPLIT HOUSING": [r"\bsplit\s+housings?\b",
                          r"\bshipping\s+splits?\b",
                          r"\bhousings?\b.*\bsplit\b.*\b(ship|shipping|shipment)\b"],
        "SILENCER": [r"silencer", r"muffler", r"sound\s*atten"],
        "ACCESS DOOR": [r"access\s*door", r"inspection\s*door", r"clean\s*out",
                        r"quick\s*open"],
        "DRAIN": [r"\bdrain"],
        "BELT GUARD": [r"belt\s*guard"],
        "SHAFT/BEARING/COUPLING GUARD": [
            r"\bshaft\b.*\bbearing\b.*\bguard\b",
            r"\bguard\b.*\bshaft\b.*\bbearing\b",
        ],
        "WEATHER COVER": [r"weather\s*(cover|hood|proof)"],
        "SCREEN": [r"\bscreen"],
        "FLANGE": [r"flange"],
        "FLEX CONNECTOR": [r"flex(ible)?\s*conn", r"expansion\s*joint"],
        "COUPLING": [r"flexible\s*coupling", r"steelflex",
                     r"thomas\s+series", r"\bhalf\s+coupling\b"],
        "UNITARY BASE": [r"unitary\s*base", r"structural\s*(steel\s*)?base",
                         r"channel\s*base"],
        "BEARINGS": [r"^bearings?\s+(standard|split\s+pillow\s+block)\b",
                     r"^repair\s+bearings?\b", r"^spare\s+bearings?\b",
                     r"^bearing\s+adder\b"],
        "LIFTING LUGS": [r"lifting\s*lugs?"],
        "LABEL": [r"\blabel\b", r"\bmark\s+all\s+items\b", r"\bmarked\s+with\s+this\s+information\b"],
        "NAMEPLATE": [r"nameplate"],
        "PACKAGING": [r"\bcrate\b", r"\bcrating\b", r"shrink\s*wrap",
                      r"\bskid\b", r"ispm\s*(?:wood\s*)?(?:inspection\s*)?(?:stamp|stamping)",
                      r"ispm\s*wood", r"do\s*not\s*stack",
                      r"metal\s*banding"],
        "SHIPPING": [r"\bship(?:ped|ping|ments?)?\b", r"freight", r"\bBOL\b",
                     r"\bUPS\s+ground\b"],
        "WARRANTY": [r"warranty"],
        "MOUNTING": [r"\bmounting\b", r"\bmounted\b", r"\bcbc\s*mount\b"],
        "SPECIAL CONSTRUCTION": [r"set\s*screws?", r"loc\s*tite", r"\bcaulk(?:ing)?\b",
                                 r"\bweld(?:ing)?\b", r"tie\s*rod\s*support",
                                 r"buffer\s*tube", r"cast\s*hub",
                                 r"plug\s*panel", r"pressure\s*tap",
                                 r"\bconduit\b", r"\boverhang\b",
                                 r"effective\s*diameter", r"threaded\s*plug",
                                 r"hole\s*diameters?", r"earthing\s*boss"],
        "INSPECTION": [r"\binspection\b",
                       r"ispm\s*(?:wood\s*)?(?:inspection\s*)?(?:stamp|stamping)",
                       r"mill\s*certifications?"],
        "DRAWINGS": [r"\bcertified\s*drawings?\b", r"\bprints?\b",
                     r"\bdrawings?\b"],
        "MOTOR": [r"\bmotor\b", r"\bc\s*[- ]?\s*flange\b"],
        "VFD": [r"\bVFD\b", r"variable\s*freq", r"inverter"],
        "EXPLOSION PROOF": [r"explosion(?:\s|;|-)*proof", r"class\s*i+\b.*div"],
        "DRIVE COMPONENTS": [r"sheave\s*/?\s*bushing", r"\bbushing\b",
                             r"\bactual\s+sf\b", r"\bactual\s+cd\b",
                             r"selected\s+drive", r"center\s+distance",
                             r"^\d{3,4}\s+\d{3,4}\s+[A-Z]{1,2}\d+\s+\d+\s+"],
        "V-BELT DRIVE": [r"^drive\b", r"v[\s-]*belt", r"sheave",
                         r"bushing", r"drive\s*set"],
        "BALANCE": [r"\bG\s*\d+(?:\.\d+)?\s+balance\b",
                    r"welded\s+balance\s+weights?"],
        "TESTING": [r"witness", r"\btest"],
        "SPARE PARTS": [r"\bspare\b", r"^(?:npo\s+)?(?:repair|replacement)\b"],
        "3D STEP DRAWINGS": [r"3d\s+(step\s+)?(file\s+)?drawings?"],
    },
}

_rules_cache: Dict[str, Any] | None = None
_MATERIAL_TAGS = {"MATERIALS", "STAINLESS STEEL", "ALUMINUM"}
_GUARD_TAG = "SHAFT/BEARING/COUPLING GUARD"
_PAINT_SURFACE_TAGS = {"MOTOR", "UNITARY BASE", "WHEEL"}
_MISC_NOTE_TAG = "MISC NOTE"
_MISC_NOTE_COMPONENT_TAGS = {"WHEEL"}
_BASE_FAN_DETAIL_TAGS = {"MOTOR", "MOUNTING"}
_BELT_GUARD_DETAIL_TAGS = {"COATING", "MOUNTING", "MOTOR"}
_DRIVE_TABLE_DETAIL_TAGS = {"DRAWINGS", "MOTOR", "SPECIAL CONSTRUCTION"}
_LABEL_DETAIL_TAGS = {"MOTOR", "NAMEPLATE"}
_EXTREME_TEMPERATURE_TAG = "EXTREME TEMP"
_PACKAGING_INSPECTION_DETAIL_TAGS = {
    "COATING", "FLANGE", "INLET", "MATERIALS", "OUTLET", "SCREEN",
    "SILENCER", "STAINLESS STEEL", "ALUMINUM",
}
_SHIP_VIA_COMPONENT_TAGS = {
    "ACTUATOR", "DAMPER", "FLEX CONNECTOR", "INLET VANES", "SILENCER", "OUTLET",
}
_ACCESSORY_COATING_TAGS = {
    "BELT GUARD", _GUARD_TAG, "MOTOR", "SILENCER", "FLEX CONNECTOR",
    "WEATHER COVER", "SCREEN",
}
_LUBE_ACCESSORY = re.compile(
    r"\b(EXTENDED\s+LUBE|LUBE\s+LINES?|GREASE\s+(LINES?|FITTINGS?|LEADS?)|ZERK\s+FITTINGS?)\b",
    re.I,
)
_COATING_WORD = re.compile(
    r"\b(PAINT(?:ED|ING)?|COAT(?:ED|ING|S)?|EPOXY|PRIMER|PLASITE|HERESITE|"
    r"ZINC|ENAMEL|UNPAINTED|GALVANIZ(?:ED|ING)?|VEGETABLE\s+OIL)\b",
    re.I,
)
_RAL_STOP_WORDS = {"COAT", "COATS", "IF", "NOT", "AVAILABLE", "USE", "INQUIRY", "NUM"}


def load_rules(path: str | Path | None = None, refresh: bool = False) -> Dict[str, Any]:
    """The compiled rules: DEFAULT_RULES merged with the optional extension
    file (arg, or env LINE_ITEM_RULES). Extension entries ADD to the defaults —
    same-named tags gain extra patterns — so a site file stays a short list of
    local wording, not a fork of the whole table. Compiled once and cached."""
    global _rules_cache
    if _rules_cache is not None and not refresh and path is None:
        return _rules_cache
    raw = {k: (dict(v) if isinstance(v, dict) else list(v)) for k, v in DEFAULT_RULES.items()}
    src = Path(path) if path else (Path(LINE_ITEM_RULES) if LINE_ITEM_RULES else None)
    if src is not None:
        try:
            ext = json.loads(src.read_text(encoding="utf-8"))
            raw["abbreviations"].update(
                {k.upper(): str(v).upper() for k, v in (ext.get("abbreviations") or {}).items()})
            for key in ("skip_patterns", "section_start", "section_end"):
                raw[key] = raw[key] + list(ext.get(key) or [])
            for tag, pats in (ext.get("tags") or {}).items():
                raw["tags"].setdefault(tag.upper(), [])
                raw["tags"][tag.upper()] = list(raw["tags"][tag.upper()]) + list(pats)
        except (OSError, json.JSONDecodeError, AttributeError, TypeError) as e:
            log.warning("Could not read line-item rules %s (%s); using defaults", src, e)
    compiled = {
        "abbreviations": raw["abbreviations"],
        "skip": [re.compile(p, re.I) for p in raw["skip_patterns"]],
        "start": [re.compile(p, re.I) for p in raw["section_start"]],
        "end": [re.compile(p, re.I) for p in raw["section_end"]],
        "tags": {t: [re.compile(p, re.I) for p in pats] for t, pats in raw["tags"].items()},
    }
    if path is None:
        _rules_cache = compiled
    return compiled


# --------------------------------------------------------------------------- #
# Normalization                                                               #
# --------------------------------------------------------------------------- #
# Trailing money: "$1,234.56", "1,234", "1234.56". A bare integer (no $ , or
# decimals) is NOT a price — it could be "3600 RPM". Item rows end in up to
# five money columns (Price Freight Markup Net Comm.), so the tail is stripped
# repeatedly and the LAST one stripped (the leftmost = the Price column) wins.
_PRICE_TAIL = re.compile(
    r"""(?:^|\s)(?P<price>
          \$\s*\d[\d,]*(?:\.\d{2})?
        | \d{1,3}(?:,\d{3})+(?:\.\d{2})?
        | \d+\.\d{2}
        | N/?C\b | NO\s+CHARGE | INCLUDED
        )\s*$""",
    re.I | re.X,
)
_EACH_TAIL = re.compile(r"(?:^|\s)(?:EA(?:CH)?|/EA|PER\s+UNIT)\s*$", re.I)
# The price-type letter ending a CBC item row: "<desc> L 945.00", "<desc> L
# STD", "<desc> L INC", or a bare trailing "<desc> L" (empty price column).
# STD/INC only count as a price right after the type letter — never let the
# "INC." of a company name turn an address line into an item.
_TYPE_PRICE_TAIL = re.compile(r"\s+(?P<ptype>[LCN])\s+(?P<mark>STD|INC|INCL|INCLUDED|N/?C)\.?\s*$", re.I)
_TYPE_TAIL = re.compile(r"\s+(?P<ptype>[LCN])\s*$")
# Leading enumeration: "1.", "(1)", "1)", "ITEM 2:", "2 -". Qty guess only.
_LEAD = re.compile(r"^\s*(?:ITEM\s*)?\(?(?P<num>\d{1,3})\)?\s*(?:[.):]|-(?=\s))?\s+", re.I)
_NONDECIMAL_DOT = re.compile(r"(?<!\d)\.|\.(?!\d)")
_DIGITS_SS = re.compile(r"^T?(\d{3}L?)(SST?|SS)$", re.I)


def split_price_tail(text: str) -> Tuple[str, str]:
    """Strip the trailing price columns / 'EACH' markers off a line. Returns
    (description, leftmost price stripped — the Price column on multi-column
    rows like 'Motor C 254.83 5.00 62.00 322.00 19.00')."""
    price = ""
    for _ in range(8):
        m = _EACH_TAIL.search(text)
        if m:
            text = text[: m.start()].rstrip()
            continue
        m = _PRICE_TAIL.search(text)
        if not m:
            break
        price = m.group("price").strip()
        text = text[: m.start()].rstrip()
    return text.strip(), price


def split_type_tail(text: str) -> Tuple[str, str, str]:
    """Strip the trailing L/C/N price-type letter (and an STD/INC/N-C mark in
    the price column). Returns (description, ptype, price-mark or '')."""
    m = _TYPE_PRICE_TAIL.search(text)
    if m:
        return text[: m.start()].rstrip(" ,;:-"), m.group("ptype").upper(), m.group("mark").upper()
    m = _TYPE_TAIL.search(text)
    if m:
        return text[: m.start()].rstrip(" ,;:-"), m.group("ptype").upper(), ""
    return text.strip(), "", ""


def split_lead(text: str) -> Tuple[str, str]:
    """Strip a leading item-number/qty marker. Returns (rest, qty guess)."""
    m = _LEAD.match(text)
    if not m:
        return text.strip(), ""
    return text[m.end():].strip(), m.group("num")


def normalize_text(raw: str, rules: Dict[str, Any] | None = None) -> str:
    """The canonical lookup form of a line: uppercase, enumeration/qty, price
    columns and the L/C/N type letter stripped, known abbreviations expanded,
    punctuation collapsed to spaces (decimals inside numbers survive).
    Spelling variants of the same option converge here — this is what search
    and tagging run against."""
    rules = rules or load_rules()
    s, _ = split_lead(raw.strip().upper())
    s, _ = split_price_tail(s)
    s, _, _ = split_type_tail(s)
    abbrev = rules["abbreviations"]
    out: List[str] = []
    for tok in s.split():
        t = tok.strip(".,;:()[]")
        if t in abbrev:
            out.append(abbrev[t])
            continue
        m = _DIGITS_SS.match(t)  # "316SS"/"304SST" -> "316 STAINLESS STEEL"
        if m:
            out.append(f"{m.group(1)} STAINLESS STEEL")
            continue
        if t.startswith("W/") and len(t) > 2:  # "W/DRAIN" -> "WITH DRAIN"
            out.append("WITH " + t[2:])
            continue
        out.append(tok)
    s = " ".join(out)
    s = _NONDECIMAL_DOT.sub(" ", s)
    s = re.sub(r"[^A-Z0-9.%#]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def tag_item(norm: str, rules: Dict[str, Any] | None = None) -> List[str]:
    """Canonical feature tags whose patterns match the normalized text (pass
    norm + normalized details to let an item's detail block contribute)."""
    rules = rules or load_rules()
    return sorted(t for t, pats in rules["tags"].items() if any(p.search(norm) for p in pats))


def _primary_norm(item: Dict[str, Any], rules: Dict[str, Any] | None = None) -> str:
    return str(item.get("norm") or normalize_text(str(item.get("raw", "")), rules)).upper()


def _item_blob(item: Dict[str, Any]) -> str:
    parts = [str(item.get("raw", ""))]
    parts += [str(d) for d in item.get("details") or []]
    return " ".join(re.sub(r"\s+", " ", p).strip() for p in parts if str(p).strip())


_DETAIL_LABELS = (
    "Vendor", "Product", "Operation", "Actuator Manufacturer",
    "Actuator Supplied By", "Actuator Mounting",
    "Actuator Size",
    "Fail Position Upon Loss of Supply/Air/Power",
    "Fail Position Upon Loss of Signal",
)


def _label_value(item: Dict[str, Any], label: str) -> str:
    parts = [str(item.get("raw", ""))]
    parts += [str(d) for d in item.get("details") or []]
    for i, part in enumerate(parts):
        m = re.match(rf"\s*{re.escape(label)}\s*:\s*(.+?)\s*$", part, re.I)
        if m:
            val = m.group(1).strip(" ,;")
            if (label.startswith("Fail Position") and val.upper() == "IN"
                    and i + 1 < len(parts) and parts[i + 1].strip().upper() == "PLACE"):
                val = "In Place"
            return val
    blob = _item_blob(item)
    labels = "|".join(re.escape(x) for x in _DETAIL_LABELS)
    m = re.search(rf"\b{re.escape(label)}\s*:\s*(.*?)(?=\s+(?:{labels})\s*:|$)",
                  blob, re.I)
    return m.group(1).strip(" ,;") if m else ""


_INQUIRY_NUM = r"[A-Z0-9]{1,6}-\d{1,2}-[A-Z0-9]{2,10}"


def inquiry_numbers(item: Dict[str, Any]) -> List[str]:
    """Inquiry numbers printed on or under this line item.

    CBC orders split these several ways: "Inquiry Num: 333-25-1622",
    "InquiryNum:333-25-1622", or "Inquiry L 645.00" followed by
    "Num: 352-23-2696" on the detail line.
    """
    parts = [re.sub(r"\s+", " ", str(item.get("raw", ""))).strip()]
    parts += [re.sub(r"\s+", " ", str(d)).strip() for d in item.get("details") or []]
    patterns = [
        re.compile(rf"\bInquiry\s*(?:Num(?:ber)?|#)\s*:?\s*({_INQUIRY_NUM})", re.I),
        re.compile(rf"\bInquiry\b.{{0,80}}?\bNum\s*:?\s*({_INQUIRY_NUM})", re.I),
    ]
    seen, nums = set(), []
    for i, part in enumerate(parts):
        candidates = [part]
        if "inquiry" in part.lower() and i + 1 < len(parts):
            candidates.append(f"{part} {parts[i + 1]}")
        for text in candidates:
            for pat in patterns:
                for m in pat.finditer(text):
                    num = m.group(1).strip(" .;,()").upper()
                    if num not in seen:
                        seen.add(num)
                        nums.append(num)
    return nums


def _used_on(norm_blob: str) -> str:
    if _is_without_ivc(norm_blob):
        return ""
    component_context = ("DAMPER" in norm_blob or "ACTUATOR" in norm_blob
                         or "VOLUME CONTROL" in norm_blob)
    if not component_context and "IVC" not in norm_blob:
        return ""
    if "FRESH AIR" in norm_blob or re.search(r"\bFA\s+DAMPER\b", norm_blob):
        return "FRESH AIR DAMPER"
    if "PRESPIN" in norm_blob or "PRE SPIN" in norm_blob:
        return "PRESPIN DAMPER"
    if "OUTLET" in norm_blob and ("DAMPER" in norm_blob or "VOLUME CONTROL" in norm_blob):
        return "OUTLET DAMPER"
    if "DISCHARGE" in norm_blob and ("DAMPER" in norm_blob or "ACTUATOR" in norm_blob):
        return "OUTLET DAMPER"
    if "INLET VANE DAMPER" in norm_blob or re.search(r"\bIVD\b", norm_blob):
        return "INLET VANE DAMPER"
    if "IVC" in norm_blob or "INLET VOLUME CONTROL" in norm_blob:
        return "IVC"
    if "INLET" in norm_blob and "DAMPER" in norm_blob:
        return "INLET DAMPER"
    return ""


def _needs_used_on_review(norm_blob: str, attrs: Dict[str, str]) -> bool:
    if attrs.get("used_on"):
        return False
    if attrs.get("product", "").upper() == "ACTUATOR":
        return True
    if attrs.get("model") or attrs.get("manufacturer"):
        return True
    if re.search(r"\b(BETTIS|UNIC|EMERSON|ACTUATOR\s*:?)\b", norm_blob, re.I):
        return True
    return False


def _is_ivc_actuator_context(norm_blob: str, tags: set[str]) -> bool:
    return bool(
        "INLET VANES" in tags
        and ("ACTUATOR" in tags or re.search(r"\bACTUATOR\b", norm_blob))
        and re.search(r"\b(IVC|INLET\s+VOLUME\s+CONTROL)\b", norm_blob)
    )


def _is_ship_via_note(primary: str, norm_blob: str) -> bool:
    return bool(
        re.search(r"\b(UPDATED|CHANGE|CHANGED)\s+SHIP\s+VIA\b", norm_blob)
        or re.search(r"\bSHIP\s+VIA\b", primary)
    )


def _is_inspection_line(primary: str, norm_blob: str) -> bool:
    inspection_patterns = [
        r"\bISPM\b.*\b(STAMP|STAMPING|INSPECTION)\b",
        r"\bORDER\s+IS\s+SHIPPING\s+OVERSEAS\b.*\bINSPECTION\s*/?\s*CRATE\s+REPORT\b",
        r"\bCUSTOMER\s+FINAL\s+INSPECTION\b",
        r"\bUNWITNESSED\s+DIMENSIONAL\s+INSPECTION\b",
        r"\bFINAL\s+INSPECTION\s+REPORT\b",
        r"\bGENERAL\s+MILL\s+CERTIFICATIONS?\b",
    ]
    return any(re.search(pattern, primary, re.I) or re.search(pattern, norm_blob, re.I)
               for pattern in inspection_patterns)


def _motor_insulation_reference(norm_blob: str) -> bool:
    return bool(re.search(
        r"\b((?:CLASS\s+)?[ABFH]\s+INSULATION|INSULATED\s+BEARINGS?|"
        r"DUAL\s+INSULATED\s+BEARINGS?|(?:DE|NDE)\s+INSULATED\s+BEARING|"
        r"INSULATED\s+(?:DE|NDE)\s+BEARING)\b",
        norm_blob,
        re.I,
    ))


def _fan_insulation_reference(norm_blob: str) -> bool:
    return bool(re.search(
        r"\b(PLUG\s+PANEL|HOUSING|INLET\s+BOX|SILENCER|DUCT|FAN)\b.{0,80}\b(INSULAT|LAGGING|JACKET(?:ED)?|MINERAL\s+WOOL|FIBERGLASS|FIBRE\s*GLASS)\b"
        r"|\b(INSULAT\w*|LAGGING|JACKET(?:ED)?|MINERAL\s+WOOL|FIBERGLASS|FIBRE\s*GLASS)\b.{0,80}\b(PLUG\s+PANEL|HOUSING|INLET\s+BOX|SILENCER|DUCT|FAN)\b",
        norm_blob,
        re.I,
    ))


def _is_label_instruction_line(primary: str, norm_blob: str, tags: set[str]) -> bool:
    if "LABEL" not in tags:
        return False
    if re.match(r"^MOTOR\b", primary, re.I):
        return False
    return bool(re.search(
        r"\b(ONLY\s+APPLY\s+CBC.*WARNING\s+LABEL|SHIPPING\s+BARCODE\s+LABEL|"
        r"FEI\s+LABEL|LABEL\s+REQUIRED\s+ON\s+EACH\s+ITEM|ON\s+FEDEX\s+SHIPPING\s+LABEL|"
        r"MARK\s+ALL\s+ITEMS|MARKED\s+WITH\s+THIS\s+INFORMATION)\b",
        norm_blob,
        re.I,
    ))


def _is_primary_lifting_lugs(primary: str) -> bool:
    return bool(re.search(r"^LIFTING\s+LUGS?\b", primary, re.I))


def _is_flex_connector_flow_liner(norm_blob: str, tags: set[str]) -> bool:
    return "FLEX CONNECTOR" in tags and bool(re.search(r"\bFLOW\s+LINERS?\b", norm_blob))


def _is_motor_insulation_only(primary: str, norm_blob: str, tags: set[str]) -> bool:
    if "INSULATION" not in tags or "MOTOR" not in tags:
        return False
    if _fan_insulation_reference(primary) or _fan_insulation_reference(norm_blob):
        return False
    return _motor_insulation_reference(norm_blob)


def _is_non_fan_shaft_seal_context(primary: str, norm_blob: str, tags: set[str]) -> bool:
    if "SHAFT SEAL" not in tags:
        return False
    if re.search(r"\bSHAFT\s+SEALS?\b|\bCERAMIC\s+FELT\b", norm_blob):
        return False
    if "DAMPER" in tags and re.search(r"\bSTUFFING\s+BOX(?:ES)?\b", norm_blob):
        return True
    if re.search(r"\bDOUBLE\s+LIP\s+SEALS?\s+BOTH\s+ENDS\b", norm_blob):
        return True
    return False


def _is_incidental_shipping_reference(primary: str, norm_blob: str, tags: set[str]) -> bool:
    if "SHIPPING" not in tags:
        return False
    if "LABEL" in tags:
        real_shipping_label = re.search(r"\b(SHIPPING\s+BAR\s*CODE\s+LABEL|SHIPPING\s+BARCODE\s+LABEL|FEDEX\s+SHIPPING\s+LABEL|SHIPPING\s+LABEL)\b", norm_blob)
        if not real_shipping_label and re.search(r"\b(SHIPPED|SHIP)\s+WITH\s+(?:THE\s+)?FAN\b", norm_blob):
            return True
    warranty_shipment = re.search(
        r"\b(WARRANTY|MONTHS?)\b.{0,120}\b(SHIPMENT|SHIP\s+DATE|DATE\s+OF\s+SHIP)\b|"
        r"\b(SHIPMENT|SHIP\s+DATE|DATE\s+OF\s+SHIP)\b.{0,120}\b(WARRANTY|MONTHS?)\b",
        norm_blob,
    )
    real_shipping_instruction = re.search(
        r"\b(SHIP\s+DIRECT|DIRECT\s+SHIP|SHIP\s+LOOSE|SHIPPED\s+LOOSE|SHIP\s+VIA|"
        r"SHIPPING\s+INSTRUCTIONS?|SHIPMENT\s+ONLINE|BOL|FREIGHT|SHIP\s+COMPLETE|"
        r"SHIP\s+WITH\s+FANS?)\b",
        norm_blob,
    )
    return bool(warranty_shipment and not real_shipping_instruction)


def _is_spare_parts_primary(primary: str) -> bool:
    return bool(re.search(r"^(?:NPO\s+)?(?:SPARE|REPAIR|REPLACEMENT)\b", primary))


def _is_packaging_inspection_primary(primary: str, tags: set[str]) -> bool:
    return bool(
        "INSPECTION" in tags
        and "PACKAGING" in tags
        and re.search(r"\b(ISPM|WOOD\s+INSPECTION\s+STAMP|LUMBER|SKID)\b", primary)
    )


_MATERIAL_GRADE = re.compile(r"\bT?(304L?|316L?)\s+STAINLESS\s+STEEL\b", re.I)
_BALANCE_GRADE = re.compile(r"\bG\s*(\d+(?:\.\d+)?)\s+BALANCE\b", re.I)
_TEMP_VALUE = re.compile(r"(?<![A-Z0-9])(-?\d{2,3})\s*(?:(?:DEG(?:REE)?|°)\s*)?([FC])\b", re.I)


def _add_unique(values: List[str], value: str) -> None:
    if value and value not in values:
        values.append(value)


def _drawing_attributes(primary: str, norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if not ({"DRAWINGS", "3D STEP DRAWINGS"} & tags):
        return {}

    attrs: Dict[str, str] = {}
    types: List[str] = []
    scopes: List[str] = []

    def add_type(value: str) -> None:
        _add_unique(types, value)

    def add_scope(value: str) -> None:
        _add_unique(scopes, value)

    if "3D STEP DRAWINGS" in tags or re.search(r"\b3D\s+(STEP\s+)?(FILE\s+)?DRAWINGS?\b", norm_blob):
        add_type("3D STEP DRAWINGS")
    if re.search(r"\bCERTIFIED\s+DRAWINGS?\b", norm_blob):
        add_type("CERTIFIED DRAWINGS")
    if re.search(r"\b(UNCERTIFIED|PRELIMINARY)\s+DRAWINGS?\b", norm_blob):
        add_type("PRELIMINARY DRAWINGS")
    if re.search(r"\b(BILL\s+OF\s+MATERIALS|BOM)\b", norm_blob):
        add_type("BOM")
    if re.search(r"\bBUYOUT\b.*\b(PART\s+NUMBERS?|ITEMS?|DRAWINGS?)\b", norm_blob):
        add_type("BUYOUT ITEMS")
    if re.search(r"\b(CG|COG|CENTER\s+OF\s+GRAVITY|WEIGHTS?|"
                 r"STATIC\s+AND\s+(?:DYNAMIC|DYMANIC)\s+LOADS?)\b", norm_blob):
        add_type("WEIGHTS/COG/LOADS")
    if re.search(r"\bGROUNDING\s+LUGS?\b", norm_blob):
        add_type("GROUNDING/LUGS")
    if re.search(r"\b(WHEEL\s+REMOVAL\s+CLEARANCES?|LIFTING\s+LOCATIONS?|SUGGESTED\s+SCROLL\s+DIMENSIONS)\b",
                 norm_blob):
        add_type("CLEARANCES/LOCATIONS")
    if re.search(r"\b(ROTATED?|ROTATION|ORIENTATION|ARRANGEMENT)\b", norm_blob):
        add_type("ORIENTATION/ARRANGEMENT")
    if (re.search(r"\b(TAG|MARK)\b", norm_blob)
            and re.search(r"\b(DRAWING|INCLUDED|INCLUDE|TRANSMITTAL)\b", norm_blob)):
        add_type("TAG/MARKING")
    if re.search(r"\b(DO\s+NOT\s+STACK|ISPM|STAMP(?:ING)?|STICKER|LABEL)\b", norm_blob):
        add_type("PACKAGING/MARKING")
    if re.search(r"\b(PLAN\s*VIEW|CUSTOMER\s+DRAWINGS?|MARKED\s+UP\s+CUSTOMER\s+DRAWINGS?)\b", norm_blob):
        add_type("PLAN VIEW/CUSTOMER DRAWING")
    if re.search(r"\b(REFERENCE\s+)?(?:CBC\s+)?OEM\s+DRAWINGS?\b", norm_blob):
        add_type("OEM REFERENCE")
    if re.search(r"\b(RED\s+MARK|LAST\s+SUBMITTED|REFERENCE\s+(?:CBC\s+)?DRAWINGS?)\b", norm_blob):
        add_type("REFERENCE/REDMARK")
    if re.search(r"\b(PROVIDE|INCLUDE)\b.*\b(DRAWING|WIRING\s+DIAGRAM|PERFORMANCE\s+DATA\s+SHEET|OMI)\b",
                 norm_blob):
        add_type("VENDOR DOCUMENTATION")
    if re.search(r"\bPROVIDE\b.*\bDRAWINGS?\s+FOR\s+APPROVAL\b", norm_blob):
        add_type("VENDOR APPROVAL DRAWING")
    if re.search(r"\bDRAWING\s+TRANSMITTAL\b", norm_blob):
        add_type("DRAWING TRANSMITTAL")
    if re.search(r"\bSYMBOLS?\b", norm_blob):
        add_type("CUSTOMER SYMBOLS/NOTES")
    if "SPLIT HOUSING" in tags or re.search(r"\bSPLIT\s+HOUSINGS?\b", norm_blob):
        add_type("SPLIT HOUSING")
    if re.search(r"\b(DRAWING\s+NOTES?|ADD\s+NOTES?\s+ON\s+THE\s+DRAWING|ON\s+DRAWINGS?\s+ADD)\b",
                 norm_blob):
        add_type("DRAWING NOTES")
    if not types:
        add_type("DRAWING NOTE")

    if "3D STEP DRAWINGS" in types:
        add_scope("3D FILE")
    if re.search(r"\bFAN\b", norm_blob):
        add_scope("FAN")
    if "MOTOR" in tags or re.search(r"\bMOTOR\b", norm_blob):
        add_scope("MOTOR")
    if "FLEX CONNECTOR" in tags or re.search(r"\b(FLEX\s+CONNECTOR|EXPANSION\s+JOINT|EJ)\b", norm_blob):
        add_scope("FLEX CONNECTOR")
    if "SILENCER" in tags or re.search(r"\bSILENCER\b", norm_blob):
        add_scope("SILENCER")
    if "WHEEL" in tags or re.search(r"\bWHEEL\b", norm_blob):
        add_scope("WHEEL")
    if "INLET" in tags or re.search(r"\bINLET\b", norm_blob):
        add_scope("INLET")
    if "OUTLET" in tags or re.search(r"\bOUTLET\b", norm_blob):
        add_scope("OUTLET")
    if "SPLIT HOUSING" in tags or re.search(r"\bSPLIT\s+HOUSINGS?\b", norm_blob):
        add_scope("SPLIT HOUSING")
    if re.search(r"\bCUSTOMER\s+DRAWINGS?\b", norm_blob):
        add_scope("CUSTOMER DRAWING")

    attrs["drawing_type"] = ", ".join(types)
    if scopes:
        attrs["drawing_scope"] = ", ".join(scopes)
    return attrs


def _split_housing_attributes(norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "SPLIT HOUSING" not in tags:
        return {}
    attrs = {"component": "SPLIT HOUSING"}
    split_types: List[str] = []
    if re.search(r"\bHORIZONTAL\s+SPLIT\s+HOUSINGS?\b", norm_blob):
        _add_unique(split_types, "HORIZONTAL")
    if re.search(r"\bPIE\s+WEDGE\s+SPLIT\s+HOUSINGS?\b", norm_blob):
        _add_unique(split_types, "PIE WEDGE")
    if re.search(
        r"\b(SHIPPING|SHIPMENT)\s+SPLITS?\b|"
        r"\bSPLIT\s+(?:HOUSINGS?\s+)?(?:FOR\s+)?(?:SHIPPING|SHIPMENT)\b|"
        r"\bSPLIT\s+(?:HOUSINGS?\s+)?TO\s+SHIP\b|"
        r"\bHOUSINGS?\s+SPLIT\s+(?:FOR\s+)?(?:SHIPPING|SHIPMENT)\b",
        norm_blob,
    ):
        _add_unique(split_types, "SHIPPING")
    if split_types:
        attrs["split_type"] = ", ".join(split_types)
    return attrs


def _is_motor_flange_line(primary: str, norm_blob: str) -> bool:
    if (re.search(r"\b(HOUSING|INLET|OUTLET)\s+FLANGE\b", norm_blob)
            and re.search(r"\bMOTOR\s+CONDUIT\s+BOX\b", norm_blob)):
        return False
    return bool(
        re.search(r"\bC\s*[- ]?\s*FLANGE\b", norm_blob)
        or re.search(r"\bMOTOR\b.{0,50}\bFLANGE\b", norm_blob)
        or re.search(r"\bFLANGE\b.{0,50}\bMOTOR\b", norm_blob)
    )


def _is_motor_nameplate_context(primary: str, norm_blob: str) -> bool:
    if not re.search(r"\b(NAMEPLATE|REPLAT(?:E|ED|ING))\b", norm_blob):
        return False
    if re.match(r"^MOTOR\b", primary, re.I):
        return True
    if re.search(r"\bMOTOR\b.{0,80}\bNAMEPLATE\b|\bNAMEPLATE\b.{0,80}\bMOTOR\b", norm_blob):
        return True
    if re.search(r"\bMOTOR\b", norm_blob) and re.search(r"\bREPLAT(?:E|ED|ING)\b", norm_blob):
        return True
    return bool(
        re.search(r"\b\d+(?:\.\d+)?\s*HP\b", norm_blob)
        and re.search(r"\b(RPM|FRAME|TEFC|TENV|XP[A-Z]*|PH|HZ|SF)\b", norm_blob)
    )


def _is_non_wheel_end_location(norm_blob: str) -> bool:
    return bool(re.search(r"\bNON\s*[- ]?\s*WHEEL\s+END\b", norm_blob))


def _is_flex_connector_line(norm_blob: str) -> bool:
    return bool(re.search(r"\b(FLEX(IBLE)?\s+CONNECTOR|EXPANSION\s+JOINT|EJ)\b", norm_blob))


def _is_mixing_box_line(primary: str, norm_blob: str) -> bool:
    return bool(re.match(r"^MIXING\s+BOX\b", primary, re.I) or "MIXING BOX" in norm_blob)


def _is_non_inlet_component_mounted_to_inlet_box(primary: str, norm_blob: str, tags: set[str]) -> bool:
    return bool(
        "INLET" in tags
        and not re.search(r"^INLET\b", primary, re.I)
        and ({"DAMPER", "ACTUATOR", "FLEX CONNECTOR"} & tags)
        and re.search(r"\bMOUNTED\s+ON\s+(?:OVERSIZED\s+)?INLET\s+BOX\b", norm_blob)
    )


def _is_without_ivc(norm_blob: str) -> bool:
    return bool(
        re.search(r"\b(WITHOUT|LESS|NO)\s+IVC\b", norm_blob)
        or re.search(r"\bWITHOUT\s+INLET\s+VOLUME\s+CONTROL\b", norm_blob)
    )


def _is_inlet_cone_width_without_wheel(norm_blob: str) -> bool:
    return bool(
        re.search(r"\bINLET\s+CONE\b", norm_blob)
        and re.search(r"\b\d+(?:\.\d+)?\s*%\s*WIDTH\b|\bPERCENT\s+WIDTH\b", norm_blob)
        and not re.search(r"\bWHEEL\b", norm_blob)
    )


def _is_motor_conduit_box_location(norm_blob: str) -> bool:
    return bool(
        re.search(r"\bCONDUIT\s+BOX\b", norm_blob)
        and re.search(r"\bHOUSING\b", norm_blob)
        and re.search(
            r"\b(HUGGING|CLOSE\s+TO|AS\s+CLOSE\s+TO|TACK\s+AND\s+WELD|"
            r"WELD\s+CONDUIT\s+BOX|CONDUIT\s+BOX\s+TO\s+HOUSING)\b",
            norm_blob,
        )
    )


def _is_motor_conduit_box_context(primary: str, norm_blob: str) -> bool:
    if not re.search(r"\bCONDUIT\s+BOX\b", norm_blob):
        return False
    return bool(
        re.search(r"^MOTOR\s+CONDUIT\s+BOX\s+LOCATION\b", primary)
        or re.search(r"\bVIEWED\s+FROM\s+OUTLET\b", norm_blob)
        or re.search(r"\bF[123]\s+CONDUIT\s+BOX\b", norm_blob)
        or re.search(r"\b(ROTATE\s+MOTOR\s+CONDUIT\s+BOX|KNOCKOUT\s+FACES|"
                     r"CONDUIT\s+BOX\s+(?:HUGGING|LOCATION)|"
                     r"MOUNT\s+CONDUIT\s+BOX\s+AS\s+CLOSE)\b", norm_blob)
        or _is_motor_conduit_box_location(norm_blob)
    )


def _is_pure_motor_conduit_box_location(primary: str, norm_blob: str) -> bool:
    if not _is_motor_conduit_box_context(primary, norm_blob):
        return False
    return not bool(
        re.search(r"\b(RUN\s+SECOND\s+CONDUIT|THREADED\s+PLUG|VERTICAL\s+MOUNTING\s+PLATE|"
                  r"FLEXIBLE\s+CONDUIT|AUXILIARY\s+BOX)\b", norm_blob)
    )


def _is_testing_context(primary: str, norm_blob: str, tags: set[str]) -> bool:
    if re.search(r"\b(TEST|WITNESS|INSPECTION|AMP\s+DRAW|ROUTINE\s+TEST|IEEE\s*112)\b", primary):
        return True
    if "MOTOR" in tags and re.search(r"\b(IEEE\s*112|ROUTINE\s+TEST|UNWITNESSED)\b", norm_blob):
        return True
    return False


def _has_housing_engineering_feature(norm_blob: str) -> bool:
    return bool(
        re.search(
            r"\b(HOUSING\s+FLANGE|DRUM\s+HOUSING\s+BOLT\s+PATTERN|"
            r"HOUSING\s+LENGTH|HOUSING\s+THICK(?:NESS)?|"
            r"HOUSING\s+STIFF(?:E|NE)R|STIFF(?:E|NE)R\s+PLATES?|"
            r"HOUSING\s+MOUNTING|MOUNT\s+HOUSING|"
            r"HOUSING\s+TO\s+DRIVE\s+COVER|TAP\s+DRIVE\s+SIDE\s+HOUSING|"
            r"CASING\s+EXTENSION|SQUARE\s+HOUSING|UNIVERSAL\s+HOUSING|"
            r"RIVETED\s+TO\s+HOUSING|NAMEPLATE\s+TO\s+(?:FAN\s+)?HOUSING|"
            r"LINERS?\b.*\bHOUSING|HOUSING\s+SCROLL)\b",
            norm_blob,
        )
        or _is_non_wheel_end_location(norm_blob)
    )


def _is_housing_packaging_reference(norm_blob: str, tags: set[str]) -> bool:
    if "HOUSING" not in tags or not ({"PACKAGING", "SHIPPING"} & tags):
        return False
    if _has_housing_engineering_feature(norm_blob):
        return False
    return bool(re.search(r"\b(SKID|SCRAP\s+WOOD|BLOCKING|BANDING|CRAT(?:E|ING)|CARTON|PACKAG)\b", norm_blob))


def _is_explosion_proof_motor_context(primary: str) -> bool:
    return bool(
        re.search(r"\bMOTOR\b", primary)
        or re.search(r"\bEXPLOSION\s+PROOF\b", primary)
        or re.search(r"\bCLASS\s*[0-9IVX]+\b.*\bDIV(?:ISION)?\b", primary)
    )


def _explosion_proof_attributes(primary: str, norm_blob: str, raw_tags: set[str]) -> Dict[str, str]:
    if "EXPLOSION PROOF" not in raw_tags and not re.search(r"\bEXPLOSION\s+PROOF\b", norm_blob):
        return {}
    if not _is_explosion_proof_motor_context(primary):
        return {}
    attrs: Dict[str, str] = {
        "component": "MOTOR",
        "motor_enclosure": "EXPLOSION PROOF",
    }
    classes: List[str] = []
    groups: List[str] = []
    divisions: List[str] = []
    for m in re.finditer(r"\bCL(?:ASS|S)?\.?\s*([0-9IVX]+)\b", norm_blob):
        _add_unique(classes, m.group(1).upper())
    for m in re.finditer(
        r"\bGR(?:OU)?PS?\.?\s+([A-Z](?:[\s,]+[A-Z])*)"
        r"(?=\s+CL(?:ASS|S)?\b|\s+DIV(?:ISION)?\b|$)",
        norm_blob,
    ):
        for letter in re.findall(r"\b[A-Z]\b", m.group(1).upper()):
            _add_unique(groups, letter)
    for m in re.finditer(r"\bDIV(?:ISION)?\.?\s*([0-9]+)\b", norm_blob):
        _add_unique(divisions, m.group(1))
    if classes:
        attrs["motor_explosion_class"] = ", ".join(classes)
    if groups:
        attrs["motor_explosion_groups"] = ", ".join(groups)
    if divisions:
        attrs["motor_explosion_division"] = ", ".join(divisions)
    return attrs


def _flange_attributes(primary: str, norm_blob: str, tags: set[str], raw_tags: set[str]) -> Dict[str, str]:
    if "FLANGE" not in raw_tags and not re.search(r"\bFLANG(?:E|ED)\b", norm_blob):
        return {}
    attrs: Dict[str, str] = {}
    scopes: List[str] = []
    if _is_motor_flange_line(primary, norm_blob):
        attrs["component"] = "MOTOR"
        attrs["motor_mounting"] = "C-FLANGE" if re.search(r"\bC\s*[- ]?\s*FLANGE\b", norm_blob) else "FLANGE"
        _add_unique(scopes, "MOTOR")
    if "FLEX CONNECTOR" in tags or _is_flex_connector_line(norm_blob):
        _add_unique(scopes, "FLEX CONNECTOR")
    if _is_non_wheel_end_location(norm_blob):
        attrs["flange_location"] = "NON-WHEEL END"
        _add_unique(scopes, "HOUSING")
    if "INLET" in tags or ("MIXING BOX" not in tags and re.search(r"\bINLET\b", norm_blob)):
        _add_unique(scopes, "INLET")
    if "OUTLET" in tags or re.search(r"\bOUTLET\b", norm_blob):
        _add_unique(scopes, "OUTLET")
    if "HOUSING" in tags and "HOUSING" not in scopes:
        _add_unique(scopes, "HOUSING")
    if "MIXING BOX" in tags:
        _add_unique(scopes, "MIXING BOX")
    if scopes:
        attrs["flange_scope"] = ", ".join(scopes)
    if re.search(r"\bC\s*[- ]?\s*FLANGE\b", norm_blob):
        attrs["flange_type"] = "C-FLANGE"
    elif "PUNCHED" in norm_blob:
        attrs["flange_type"] = "PUNCHED"
    elif re.search(r"\bFLANGED\b", norm_blob):
        attrs["flange_type"] = "FLANGED"
    return attrs


def _inlet_attributes(primary: str, norm_blob: str, tags: set[str], raw_blob: str) -> Dict[str, str]:
    if "INLET" not in tags:
        return {}
    attrs: Dict[str, str] = {}
    subcategories: List[str] = []
    features: List[str] = []

    def add_subcategory(value: str) -> None:
        _add_unique(subcategories, value)

    def add_feature(value: str) -> None:
        _add_unique(features, value)

    if re.search(r"\bINLET,\s*OPEN\b|\bINLET\s+OPEN\b", primary):
        add_subcategory("OPEN")
    if re.search(r"\bINLET,\s*SLIP\b|\bINLET\s+SLIP\b", primary):
        add_subcategory("SLIP")
    if re.search(r"\bINLET,\s*BELL\b|\bINLET\s+BELL\b", primary):
        add_subcategory("BELL")
    if re.search(r"\bINLET\s+CONE\b|\bINTEGRAL\s+INLET\s+CONE\b", primary):
        add_subcategory("INLET CONE")
    if re.search(r"\bINLET\s+BOX\b", primary):
        add_subcategory("INLET BOX")
    if re.search(r"\bINLET,\s*TUBE\b|\bINLET\s+TUBE\b", primary):
        add_subcategory("TUBE")

    if "PUNCHED" in primary:
        add_subcategory("FLANGED/PUNCHED" if "FLANGED" in primary else "PUNCHED")
        add_feature("PUNCHED")
    elif "FLANGED" in primary:
        add_subcategory("FLANGED")
    if "STANDARD BOLTED" in primary:
        add_feature("STANDARD BOLTED")
    if re.search(r"\bWELDED\b", primary):
        add_feature("WELDED")
    if re.search(r"\bWITHOUT\s+IVC\b", primary):
        attrs["ivc_relation"] = "WITHOUT IVC"
    elif re.search(r"\bWITH\s+IVC\b", primary):
        attrs["ivc_relation"] = "WITH IVC"

    if "INLET DIRECTION" in primary:
        add_subcategory("DIRECTION")
        m = re.search(r"\bINLET\s+DIRECTION\s+((?:VERTICAL|HORIZONTAL)\s+INLET\s+(?:UP|DOWN|LEFT|RIGHT))\b", primary)
        if m:
            attrs["inlet_direction"] = m.group(1)
    if "BOLT ON" in primary or "BOLT-ON" in raw_blob.upper():
        add_feature("BOLT-ON")
        attrs["inlet_box_type"] = "BOLT-ON"
    if "ASSEMBLY" in primary and "INLET BOX" in primary:
        add_feature("ASSEMBLY")
        attrs["inlet_box_type"] = "ASSEMBLY"
    if "GASKET" in primary:
        add_feature("GASKET")
    if re.search(r"\bSUPPORT\s+LEGS\b", primary):
        add_feature("SUPPORT LEGS")
    if "OVERSIZED" in norm_blob and "INLET BOX" in primary:
        add_feature("OVERSIZED")
    if "SHIPPED LOOSE" in norm_blob or "SHIP LOOSE" in norm_blob:
        attrs["shipping_state"] = "SHIPPED LOOSE"
    m = re.search(r"\bINLET\s+BOX\s+SIZE\s+(\d+)\b|\bBOX\s+SIZE\s+(\d+)\b", norm_blob)
    if m:
        attrs["inlet_box_size"] = next(v for v in m.groups() if v)
    m = re.search(r"\b(?:INLET\s+CONE\s+)?(\d+)\s*%", primary)
    if m and "INLET CONE" in primary:
        attrs["inlet_cone_width_percent"] = m.group(1)
    m = re.search(r"\bSIZE\s+(\d+)\b", primary)
    if m and "INLET CONE" in primary:
        attrs["inlet_size"] = m.group(1)
    m = re.search(r"INLET\s+BOX\s+POSITION\s*:\s*@\s*([0-9]+)", raw_blob, re.I)
    if m:
        attrs["inlet_box_position"] = m.group(1)

    if "ACCESS DOOR" in tags:
        add_feature("ACCESS DOOR")
    if "SCREEN" in tags:
        add_feature("SCREEN")
    if "LINING" in tags:
        add_feature("LINING")
    if re.search(r"\bMOUNTING\s+HARDWARE\b", primary):
        add_feature("MOUNTING HARDWARE")
    if "WITHOUT CAULK" in primary:
        add_feature("WITHOUT CAULK")
    if "CARTONED" in norm_blob and "INLET CONE" in primary:
        add_feature("CARTONED")

    if subcategories:
        attrs["inlet_subcategory"] = ", ".join(subcategories)
    if features:
        attrs["inlet_feature"] = ", ".join(features)
    return attrs


def _mixing_box_attributes(norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "MIXING BOX" not in tags:
        return {}
    attrs: Dict[str, str] = {"component": "MIXING BOX"}
    features: List[str] = []
    if "FGR PORT" in norm_blob:
        _add_unique(features, "FGR PORT")
    if "FLANGED" in norm_blob:
        _add_unique(features, "FLANGED")
    if "SHIPPED LOOSE" in norm_blob or "SHIP LOOSE" in norm_blob:
        attrs["shipping_state"] = "SHIPPED LOOSE"
    if "INLET BOX" in norm_blob:
        attrs["used_on"] = "INLET BOX"
    m = re.search(r"\bSIZE\s+(\d+)\s+INLET\s+BOX\b|\bINLET\s+BOX\s+SIZE\s+(\d+)\b", norm_blob)
    if m:
        attrs["used_on_size"] = next(v for v in m.groups() if v)
    if features:
        attrs["mixing_box_feature"] = ", ".join(features)
    return attrs


def _inlet_mount_attributes(norm_blob: str) -> Dict[str, str]:
    if re.search(r"\bMOUNTED\s+ON\s+(?:OVERSIZED\s+)?INLET\s+BOX\b", norm_blob):
        return {"mount_location": "INLET BOX"}
    return {}


def _inlet_vane_attributes(primary: str, norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "INLET VANES" not in tags:
        return {}
    attrs: Dict[str, str] = {}
    subcategories: List[str] = []
    features: List[str] = []

    def add_subcategory(value: str) -> None:
        _add_unique(subcategories, value)

    def add_feature(value: str) -> None:
        _add_unique(features, value)

    if _is_ivc_actuator_context(norm_blob, tags):
        add_subcategory("IVC ACTUATOR")
        attrs["ivc_component"] = "ACTUATOR"
    if "INLET VANE DAMPER" in norm_blob or re.search(r"\bIVD\b", norm_blob):
        add_subcategory("INLET VANE DAMPER")
        attrs["damper_subcategory"] = "INLET VANE DAMPER"
    if re.search(r"\bINLET\b.*\bFLANGED\b.*\bPUNCHED\b.*\bWITH\s+IVC\b", norm_blob):
        add_subcategory("IVC INLET FLANGE")
    if not subcategories:
        add_subcategory("IVC")

    if re.search(r"\bAUTOMATIC\b", norm_blob):
        attrs["operation"] = "Automatic"
    elif re.search(r"\bMANUAL\b", norm_blob):
        attrs["operation"] = "Manual"

    if re.search(r"\bLOCKING(?:\s+QUADRANT)?\b", norm_blob):
        add_feature("LOCKING QUADRANT")
    if re.search(r"\bROTATING\s+RING\s+ARM\b", norm_blob):
        add_feature("ROTATING RING ARM")
    m = re.search(r"\bSIZE\s+(\d{3,4})\b|\b(\d{3,4})\s+LOW\s+LEAK\s+IVC\b", norm_blob)
    if m:
        attrs["ivc_size"] = next(v for v in m.groups() if v)
    m = re.search(r"(?:@\s*)?([0-9]{1,2})\s+O\s*CLOCK", norm_blob)
    if m:
        attrs["ivc_arm_position"] = f"{m.group(1)} O'CLOCK"

    if subcategories:
        attrs["ivc_subcategory"] = ", ".join(subcategories)
    if features:
        attrs["ivc_feature"] = ", ".join(features)
    attrs.setdefault("used_on", "IVC")
    return attrs


def _inspection_attributes(primary: str, norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "INSPECTION" not in tags:
        return {}
    attrs: Dict[str, str] = {}
    if re.search(r"\bISPM\b.*\b(STAMP|STAMPING|INSPECTION)\b", norm_blob):
        attrs["inspection_subcategory"] = "ISPM WOOD STAMP"
        if "PACKAGING" in tags:
            attrs["inspection_scope"] = "PACKAGING"
    elif re.search(r"\bORDER\s+IS\s+SHIPPING\s+OVERSEAS\b", norm_blob):
        attrs["inspection_subcategory"] = "OVERSEAS CRATE REPORT"
        attrs["inspection_scope"] = "PACKAGING"
    elif re.search(r"\bCUSTOMER\s+FINAL\s+INSPECTION\b", primary):
        attrs["inspection_subcategory"] = "CUSTOMER FINAL INSPECTION"
    elif re.search(r"\bUNWITNESSED\s+DIMENSIONAL\s+INSPECTION\b", primary):
        attrs["inspection_subcategory"] = "DIMENSIONAL INSPECTION"
        attrs["witnessed"] = "NO"
    elif re.search(r"\bFINAL\s+INSPECTION\s+REPORT\b", norm_blob):
        attrs["inspection_subcategory"] = "FINAL INSPECTION REPORT"
    return attrs


def _motor_insulation_attributes(norm_blob: str, tags: set[str], raw_tags: set[str]) -> Dict[str, str]:
    if "MOTOR" not in tags and "MOTOR" not in raw_tags and not re.search(r"\bMOTOR\b", norm_blob):
        return {}
    attrs: Dict[str, str] = {}
    classes = []
    for m in re.finditer(r"\b(?:CLASS\s+)?([ABFH])\s+INSULATION\b", norm_blob):
        _add_unique(classes, m.group(1).upper())
    if classes:
        attrs["motor_insulation_class"] = ", ".join(classes)

    if re.search(r"\b(DUAL\s+INSULATED\s+BEARINGS?|INSULATED\s+BEARINGS?\b.*\bDE\b.*\bNDE\b)", norm_blob):
        attrs["motor_insulated_bearing"] = "DE AND NDE"
    elif re.search(r"\b(NDE\s+INSULATED\s+BEARING|INSULATED\s+NDE\s+BEARING|INSULATED\s+BEARINGS?\s*\(?NDE)\b", norm_blob):
        attrs["motor_insulated_bearing"] = "NDE"
    elif re.search(r"\b(DE\s+INSULATED\s+BEARING|INSULATED\s+DE\s+BEARING|INSULATED\s+BEARINGS?\s*\(?DE)\b", norm_blob):
        attrs["motor_insulated_bearing"] = "DE"
    elif re.search(r"\bINSULATED\s+BEARINGS?\b", norm_blob):
        attrs["motor_insulated_bearing"] = "YES"

    if re.search(r"\bAEGIS\s+RING\b", norm_blob):
        attrs["motor_shaft_grounding"] = "AEGIS RING"
    elif re.search(r"\bSHAFT\s+GROUNDING\s+BRUSH\b", norm_blob):
        attrs["motor_shaft_grounding"] = "SHAFT GROUNDING BRUSH"
    elif re.search(r"\bSHAFT\s+GROUNDING\s+RING\b", norm_blob):
        attrs["motor_shaft_grounding"] = "SHAFT GROUNDING RING"
    elif re.search(r"\bSHAFT\s+GROUNDING\b", norm_blob):
        attrs["motor_shaft_grounding"] = "SHAFT GROUNDING"

    if re.search(r"\b(VPI|VACUUM\s+PRESSURE\s+IMPREGNATION)\b", norm_blob):
        attrs["motor_construction"] = "VPI"
    return attrs


def _insulation_attributes(item: Dict[str, Any], primary: str, norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "INSULATION" not in tags:
        return {}
    attrs: Dict[str, str] = {}
    if "PLUG PANEL" in norm_blob:
        attrs["insulation_scope"] = "PLUG PANEL"
    elif "HOUSING" in norm_blob:
        attrs["insulation_scope"] = "HOUSING"
    elif "INLET BOX" in norm_blob:
        attrs["insulation_scope"] = "INLET BOX"
    elif "SILENCER" in norm_blob:
        attrs["insulation_scope"] = "SILENCER"
    elif "DUCT" in norm_blob:
        attrs["insulation_scope"] = "DUCT"
    elif "FAN" in norm_blob:
        attrs["insulation_scope"] = "FAN"
    if re.search(r"\bACOUSTIC\b|\bSOUND\s+PILLOW\b", norm_blob):
        attrs["insulation_type"] = "ACOUSTIC"
    elif re.search(r"\bTHERMAL\b", norm_blob):
        attrs["insulation_type"] = "THERMAL"
    elif re.search(r"\bLAGGING\b", norm_blob):
        attrs["insulation_type"] = "LAGGING"
    elif re.search(r"\bJACKET(?:ED)?\b", norm_blob):
        attrs["insulation_type"] = "JACKET"
    elif re.search(r"\bMINERAL\s+WOOL\b", norm_blob):
        attrs["insulation_material"] = "MINERAL WOOL"
    elif re.search(r"\bFIBERGLASS\b|\bFIBRE\s*GLASS\b", norm_blob):
        attrs["insulation_material"] = "FIBERGLASS"

    raw_blob = _item_blob(item)
    m = re.search(
        r"\b(?:INSULATION|LAGGING)\s+([0-9]+(?:\.[0-9]+)?)\s*(\")?"
        r"|\b([0-9]+(?:\.[0-9]+)?)\s*(\")?\s+(?:INSULATION|LAGGING)\b",
        raw_blob,
        re.I,
    )
    if m:
        value = m.group(1) or m.group(3)
        suffix = '"' if m.group(2) or m.group(4) or attrs.get("insulation_scope") else ""
        attrs["insulation_thickness"] = f"{value}{suffix}"
    insulated_by = _label_value(item, "Insulated By")
    if insulated_by:
        attrs["insulated_by"] = insulated_by
    return attrs


def _label_attributes(norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "LABEL" not in tags:
        return {}
    attrs: Dict[str, str] = {}
    if re.search(r"\bSHIPPING\s+BAR\s*CODE\s+LABEL\b|\bSHIPPING\s+BARCODE\s+LABEL\b", norm_blob):
        attrs["label_type"] = "SHIPPING BARCODE LABEL"
        attrs["label_scope"] = "SHIPPING"
    elif re.search(r"\bFEI\s+LABEL\b", norm_blob):
        attrs["label_type"] = "FEI LABEL"
    elif re.search(r"\bWARNING\s+LABEL\b", norm_blob):
        attrs["label_type"] = "WARNING LABEL"
        attrs["label_scope"] = "FAN" if "FAN" in norm_blob else "WARNING"
    elif re.search(r"\bRETIE\s+LABEL\b", norm_blob):
        attrs["label_type"] = "RETIE LABEL"
        attrs["label_scope"] = "MOTOR"
    elif re.search(r"\bLABEL\s+REQUIRED\s+ON\s+EACH\s+ITEM\b", norm_blob):
        attrs["label_type"] = "ITEM LABEL"
        attrs["label_scope"] = "EACH ITEM"
    elif re.search(r"\bMARK\s+ALL\s+ITEMS\b|\bMARKED\s+WITH\s+THIS\s+INFORMATION\b", norm_blob):
        attrs["label_type"] = "ITEM MARKING"
        attrs["label_scope"] = "EACH ITEM"
    elif re.search(r"\bFEDEX\s+SHIPPING\s+LABEL\b|\bSHIPPING\s+LABEL\b", norm_blob):
        attrs["label_type"] = "SHIPPING LABEL"
        attrs["label_scope"] = "SHIPPING"

    if re.search(r"\bSTICKERS?\s*/?\s*LABELS?.*\bNAMEPLATE\b.*\bBAG\b", norm_blob):
        attrs["label_handling"] = "OTHER LABELS/NAMEPLATE BAGGED"
    if re.search(r"\bVENDOR\s+S\s+MOTOR\s+NAMEPLATE\b.*\bAPPLIED\b", norm_blob):
        attrs["related_nameplate_handling"] = "VENDOR MOTOR NAMEPLATE APPLIED"
    return attrs


def _mounting_attributes(primary: str, norm_blob: str, tags: set[str]) -> Dict[str, str]:
    attrs: Dict[str, str] = {}
    if "MOTOR" in tags or _is_motor_nameplate_context(primary, norm_blob):
        values: List[str] = []
        if re.search(r"\bC\s*[- ]?\s*FLANGE\b|\bC\s*[- ]?\s*FACE\b", norm_blob):
            _add_unique(values, "C-FLANGE")
        if re.search(r"\bFOOT\s*MOUNT(?:ED|ING)?\b|\bFOOTMOUNT(?:ED|ING)?\b", norm_blob):
            _add_unique(values, "FOOT MOUNTED")
        if re.search(r"\bWITH\s+FEET\b", norm_blob):
            _add_unique(values, "WITH FEET")
        if re.search(r"\b(NO|WITHOUT|LESS)\s+FEET\b", norm_blob):
            _add_unique(values, "NO FEET")
        if re.search(r"\bMULTI\s*MOUNT(?:ING)?\b|\bMULTIMOUNT(?:ING)?\b", norm_blob):
            _add_unique(values, "MULTIMOUNTING")
        if re.search(r"\bMOTOR\s+MOUNTING\s+FRAME\b", norm_blob):
            _add_unique(values, "MOTOR MOUNTING FRAME")
        if re.search(r"\bADJUSTABLE\s+MOTOR\s+BASE\b", norm_blob):
            attrs["motor_base"] = "ADJUSTABLE MOTOR BASE"
        if values:
            attrs["component"] = "MOTOR"
            attrs["motor_mounting"] = ", ".join(values)
    if re.search(r"\bCBC\s+MOUNT(?:ED)?\b", norm_blob):
        attrs.setdefault("mounting", "CBC MOUNT")
    return attrs


def _nameplate_attributes(primary: str, norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "LABEL" in tags and "NAMEPLATE" not in tags:
        return {}
    if "NAMEPLATE" not in tags and not _is_motor_nameplate_context(primary, norm_blob):
        return {}
    attrs: Dict[str, str] = {}
    if "MOTOR" in tags or _is_motor_nameplate_context(primary, norm_blob):
        attrs["component"] = "MOTOR"
        attrs["motor_nameplate"] = "YES"
        if re.search(r"\bREPLACE\s+NAMEPLATE\b", norm_blob):
            attrs["motor_nameplate_action"] = "REPLACE NAMEPLATE"
        elif re.search(r"\bREPLAT(?:E|ED|ING)\b", norm_blob):
            attrs["motor_nameplate_action"] = "REPLATE"
        elif re.search(r"\bNAMEPLATE\s+RR\b", norm_blob):
            attrs["motor_nameplate_action"] = "REPLACE NAMEPLATE"
    elif "HOUSING" in tags or _is_nameplate_housing_mount(norm_blob):
        attrs["nameplate_mount_location"] = "HOUSING"
        if re.search(r"\bRIVETED\s+TO\s+HOUSING\b", norm_blob):
            attrs["nameplate_mounting"] = "RIVETED"
    else:
        attrs["component"] = "NAMEPLATE"
        if "STAINLESS STEEL" in tags or "STAINLESS STEEL" in norm_blob:
            attrs["nameplate_type"] = "STAINLESS STEEL"
    return attrs


def _lifting_lug_attributes(tags: set[str]) -> Dict[str, str]:
    if "LIFTING LUGS" not in tags:
        return {}
    return {
        "component": "HOUSING",
        "housing_feature": "LIFTING LUGS",
        "lifting_lugs": "YES",
    }


def _screen_attributes(primary: str, norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "SCREEN" not in tags:
        return {}
    attrs: Dict[str, str] = {}
    subcategories: List[str] = []
    features: List[str] = []
    used_on: List[str] = []

    def add_subcategory(value: str) -> None:
        _add_unique(subcategories, value)

    def add_feature(value: str) -> None:
        _add_unique(features, value)

    def add_used_on(value: str) -> None:
        _add_unique(used_on, value)

    direct_screen = bool(
        re.match(r"^(?:OUTLET\s+AND\s+INLET\s+SCREEN|(?:INLET|OUTLET)\s+(?:WIRE\s+)?SCREEN|SCREEN)\b", primary)
        or re.search(r"\bINLET\s+SCREEN\s+MOUNTING\s+RING\b", primary)
        or re.search(r"\bSELF\s+TAPPERS?\b.*\bINLET\s+SCREEN\b", primary)
        or re.match(r"^(?:LASER\s*[- ]?\s*CUT|OVERSIZED|STANDARD)\b.*\bSCREEN\b", primary)
    )

    if re.search(r"\bTRASH\s+SCREEN\b", norm_blob):
        add_subcategory("TRASH SCREEN")
    if re.search(r"\bWIRE\s+SCREEN\s+GUARD\b|\bWIRE\s+SCREEN\b", norm_blob):
        add_subcategory("WIRE SCREEN GUARD")
    if re.search(r"\bSCREEN\s+MOUNTING\s+RING\b", norm_blob):
        add_subcategory("MOUNTING RING")
    if re.search(
        r"\bSELF\s+TAPPERS?\b.*\bSCREEN\b|"
        r"\bSCREEN\b.{0,40}\bHARDWARE\b|"
        r"\bHARDWARE\b.{0,40}\bSCREEN\b",
        norm_blob,
    ):
        add_subcategory("HARDWARE")
    if re.search(r"\b(RAIN\s*HOOD|RAINHOOD|WEATHER\s+HOOD|WEATHER\s+COVER)\b", norm_blob):
        add_subcategory("RAINHOOD SCREEN")
        add_used_on("WEATHER COVER")
    if not subcategories:
        add_subcategory("SCREEN")

    if re.search(r"\bOVERSIZED\b", norm_blob):
        add_feature("OVERSIZED")
    if re.search(r"\bSTANDARD\b", norm_blob):
        add_feature("STANDARD")
    if re.search(r"\bLASER\s*[- ]?\s*CUT\b", norm_blob):
        add_feature("LASER-CUT")
    if re.search(r"\bBOLT\s*[- ]?\s*ON\b", norm_blob):
        add_feature("BOLT-ON")
    if re.search(r"\bCBC\s+MOUNT(?:ED)?\b", norm_blob):
        add_feature("CBC MOUNT")
    if re.search(r"\bSHIPPED\s+LOOSE\b|\bSHIP\s+LOOSE\b", norm_blob):
        attrs["shipping_state"] = "SHIPPED LOOSE"

    silencer_context = "SILENCER" in tags or "SILENCER" in norm_blob
    if silencer_context and not direct_screen:
        if re.search(r"\bINLET\s+SCREEN\b|\bSCREEN\b.{0,20}\bINLET\b", norm_blob):
            add_used_on("INLET")
        if re.search(r"\b(OUTLET|DISCHARGE)\s+SCREEN\b|\bSCREEN\b.{0,20}\b(OUTLET|DISCHARGE)\b", norm_blob):
            add_used_on("OUTLET")
    else:
        if re.search(r"\bINLET\b.{0,40}\bSCREEN\b|\bSCREEN\b.{0,40}\bINLET\b", norm_blob):
            add_used_on("INLET")
        if re.search(r"\b(OUTLET|DISCHARGE)\b.{0,40}\bSCREEN\b|\bSCREEN\b.{0,40}\b(OUTLET|DISCHARGE)\b", norm_blob):
            add_used_on("OUTLET")
    if silencer_context:
        add_used_on("SILENCER")

    if re.search(r"\bMOUNTED\s+TO\s+BELL\b", norm_blob):
        attrs["mount_location"] = "BELL"
    elif re.search(r"\bMOUNTED\s+TO\s+FAN\b", norm_blob):
        attrs["mount_location"] = "FAN"
    elif re.search(r"\bMOUNTED\s+TO\s+INLET\s+CONE\b", norm_blob):
        attrs["mount_location"] = "INLET CONE"

    m = re.search(r"\bINLET\s+DIA\s+([0-9]+(?:\.[0-9]+)?)\b", norm_blob)
    if m:
        attrs["screen_diameter"] = m.group(1)
    m = re.search(r"\bSIZE\s+([0-9]{2,4})\b", primary)
    if m:
        attrs["screen_size"] = m.group(1)

    if direct_screen:
        attrs["component"] = "SCREEN"
    if subcategories:
        attrs["screen_subcategory"] = ", ".join(subcategories)
    if features:
        attrs["screen_feature"] = ", ".join(features)
    if used_on:
        attrs["used_on"] = ", ".join(used_on)
    return attrs


def _shaft_cooler_attributes(primary: str, norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "SHAFT COOLER" not in tags:
        return {}
    attrs: Dict[str, str] = {"shaft_cooler": "YES"}
    if re.match(r"^(SHAFT\s+COOLER|HEAT\s+SLINGER)\b", primary):
        attrs["component"] = "SHAFT COOLER"
    if re.search(r"\bHEAT\s+SLINGER\b", norm_blob):
        attrs["shaft_cooler_type"] = "HEAT SLINGER"
    else:
        attrs["shaft_cooler_type"] = "SHAFT COOLER"
    if "CAST" in norm_blob:
        attrs["shaft_cooler_construction"] = "CAST"
    return attrs


def _shaft_seal_attributes(norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "SHAFT SEAL" not in tags:
        return {}
    attrs: Dict[str, str] = {"component": "SHAFT SEAL"}
    types: List[str] = []
    if re.search(r"\bNOT\s+GAS\s+TIGHT\b", norm_blob):
        _add_unique(types, "NOT GAS TIGHT")
    if re.search(r"\bLEAK\s+RESISTANT\b", norm_blob):
        _add_unique(types, "LEAK RESISTANT")
    if re.search(r"\bPTFE\b", norm_blob):
        _add_unique(types, "PTFE SEAL RING")
    if re.search(r"\bDOUBLE\s+CARBON\b", norm_blob):
        _add_unique(types, "DOUBLE CARBON")
    if re.search(r"\bCERAMIC\s+FELT\b", norm_blob):
        _add_unique(types, "CERAMIC FELT")
    if re.search(r"\bLIP\s+SEALS?\b", norm_blob):
        _add_unique(types, "LIP SEAL")
    if re.search(r"\bSTUFFING\s+BOX(?:ES)?\b", norm_blob):
        _add_unique(types, "STUFFING BOX")
    if types:
        attrs["shaft_seal_type"] = ", ".join(types)
    if re.search(r"\bJOHN\s+CRANE\b", norm_blob):
        attrs["manufacturer"] = "JOHN CRANE"
    return attrs


def _shaft_sleeve_attributes(norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "SHAFT SLEEVE" not in tags:
        return {}
    attrs: Dict[str, str] = {
        "component": "SHAFT SLEEVE",
        "shaft_sleeve": "YES",
        "used_on": "SHAFT",
    }
    if re.search(r"\bSPLIT\s+SLEEVE\b", norm_blob):
        attrs["shaft_sleeve_type"] = "SPLIT SLEEVE"
    return attrs


def _shipping_attributes(norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "SHIPPING" not in tags:
        return {}
    attrs: Dict[str, str] = {}
    states: List[str] = []
    methods: List[str] = []
    instructions: List[str] = []
    scopes: List[str] = []

    def add_state(value: str) -> None:
        _add_unique(states, value)

    def add_method(value: str) -> None:
        _add_unique(methods, value)

    def add_instruction(value: str) -> None:
        _add_unique(instructions, value)

    def add_scope(value: str) -> None:
        _add_unique(scopes, value)

    if re.search(r"\bSHIP(?:PED)?\s+LOOSE\b|\bSHIP\s+LOOSE\b", norm_blob):
        add_state("SHIPPED LOOSE")
    if re.search(r"\bDIRECT\s+SHIP\b|\bSHIP\b(?:\s+\S+){0,8}\s+DIRECT\b", norm_blob):
        add_method("SHIP DIRECT")
    if re.search(r"\bFREIGHT\s+INCLUDED\b", norm_blob):
        add_method("FREIGHT INCLUDED")
    if re.search(r"\bSHIP\s+VIA\b", norm_blob):
        add_instruction("SHIP VIA")
    if re.search(r"\bUPDATED\s+SHIP\s+VIA\b", norm_blob):
        add_instruction("UPDATED SHIP VIA")
    if re.search(r"\bUPS\s+GROUND\b", norm_blob):
        add_method("UPS GROUND")
    if re.search(r"\bDOMESTIC\s+FREIGHT\b", norm_blob):
        add_method("DOMESTIC FREIGHT")
    if re.search(r"\bCATERPILLAR\b", norm_blob):
        add_method("CATERPILLAR SHIPPING PORTAL")

    if re.search(r"\bSHIP\s+WITH\s+FANS?\b", norm_blob):
        add_instruction("SHIP WITH REFERENCED FANS")
        add_scope("FAN")
    if re.search(r"\bALL\s+AUXILIARY\s+ITEMS\s+EXCEPT\s+MOTOR\b", norm_blob):
        add_instruction("SHIP AUXILIARY ITEMS LOOSE")
        add_scope("AUXILIARY ITEMS EXCEPT MOTOR")
    if re.search(r"\bPICTURES?\b.*\bPRIOR\s+TO\s+SHIPPING\b", norm_blob):
        add_instruction("PICTURES PRIOR TO SHIPPING")
        add_scope("FAN")
    if re.search(r"\b(SHIPPING\s+(?:PAPERS|DOCUMENTS|BOXES)|PACKING\s+LIST|BOL)\b", norm_blob):
        add_instruction("SHIPPING DOCUMENTS/MARKING")
    if re.search(r"\b(SHIPPING\s+INSTRUCTIONS?|SHIPPING\s+INFORMATION|INVOICE\s+INSTRUCTIONS)\b", norm_blob):
        add_instruction("SHIPPING/INVOICE INSTRUCTIONS")
    if re.search(r"\bSHIPPING\s+PAPERS\b", norm_blob):
        add_instruction("MARK SHIPPING PAPERS")
    if re.search(r"\bENTER\s+SHIPMENT\s+ONLINE\b|\bSHIPMENT\s+ONLINE\b", norm_blob):
        add_instruction("ENTER SHIPMENT ONLINE")
    if re.search(r"\bDIMENSIONS?\s+TO\s+SET\s+UP\s+SHIPMENT\b|\bSET\s+UP\s+SHIPMENT\b", norm_blob):
        add_instruction("SHIPMENT SETUP DIMENSIONS")
    if re.search(r"\bSEND\b.{0,80}\bCUSTOMS?\s+DOCUMENTS?\b.{0,80}\bSHIPMENT\b", norm_blob):
        add_instruction("SEND CUSTOMS DOCUMENTS")
    if re.search(r"\bDOES?\s+NOT\s+SHIP\s+EARLY\b|\bNOT\s+SHIP\s+EARLY\b", norm_blob):
        add_instruction("DO NOT SHIP EARLY")
    if re.search(r"\bSHIP\s+COMPLETE\b|\bNO\s+PARTIAL\b", norm_blob):
        add_instruction("SHIP COMPLETE")
    if re.search(r"\bWEIGHT\b.*\bDIMS?\b|\bDIMS?\b.*\bWEIGHT\b", norm_blob):
        add_instruction("PROVIDE WEIGHT/DIMS")
    if re.search(r"\bMOTOR\s+BASE\b.*\bSHIPPING\s+PURPOSES\b", norm_blob):
        add_instruction("MOTOR BASE FOR SHIPPING ONLY")
        add_scope("MOTOR BASE")
    if re.search(r"\bSHIPPING\s+COVERS?\b", norm_blob):
        attrs["component"] = "SHIPPING COVER"
        if "INLET" in norm_blob:
            add_scope("INLET")
        if "OUTLET" in norm_blob:
            add_scope("OUTLET")

    for tag, scope in (
        ("MOTOR", "MOTOR"),
        ("FLEX CONNECTOR", "FLEX CONNECTOR"),
        ("DAMPER", "DAMPER"),
        ("INLET VANES", "IVC"),
        ("SILENCER", "SILENCER"),
        ("VIBRATION ISOLATION", "VIBRATION ISOLATION"),
        ("PACKAGING", "PACKAGING"),
        ("LABEL", "LABEL"),
    ):
        if tag in tags:
            add_scope(scope)

    if states:
        attrs["shipping_state"] = ", ".join(states)
    if methods:
        attrs["shipping_method"] = ", ".join(methods)
    if instructions:
        attrs["shipping_instruction"] = ", ".join(instructions)
    if scopes:
        attrs["shipping_scope"] = ", ".join(scopes)
    return attrs


def _silencer_attributes(primary: str, norm_blob: str, tags: set[str], product: str = "",
                         blob: str = "") -> Dict[str, str]:
    if "SILENCER" not in tags:
        return {}
    attrs: Dict[str, str] = {}
    product_norm = normalize_text(product) if product else ""
    primary_context = f"{primary} {product_norm}".strip()
    direction_source = primary_context or norm_blob
    used_on: List[str] = []
    if re.search(r"\bINLET\b", direction_source):
        _add_unique(used_on, "INLET")
    if re.search(r"\b(OUTLET|DISCHARGE)\b", direction_source):
        _add_unique(used_on, "OUTLET")
    if not used_on and re.search(r"\bINLET\b", norm_blob):
        _add_unique(used_on, "INLET")
    if not used_on and re.search(r"\b(OUTLET|DISCHARGE)\b", norm_blob):
        _add_unique(used_on, "OUTLET")
    if used_on:
        attrs["silencer_used_on"] = ", ".join(used_on)

    if re.search(r"\bCIRCULAR\s+DISCHARGE\s+SILENCER\b|\bCIB\b", norm_blob):
        attrs["silencer_subcategory"] = "CIRCULAR DISCHARGE SILENCER"
    elif "OUTLET" in used_on:
        attrs["silencer_subcategory"] = "OUTLET SILENCER"
    elif "INLET" in used_on:
        attrs["silencer_subcategory"] = "INLET SILENCER"
    else:
        attrs["silencer_subcategory"] = "SILENCER"

    model_patterns = (
        r"\bMODEL\s*:?\s*([A-Z0-9]+(?:[-\s][A-Z0-9]+){0,3})\b",
        r"\b(?:VAW\s+)?SILENCER\s+([0-9A-Z]+VRSB[-\s][A-Z0-9]+)\b",
        r"\b([0-9]{2}VCIB[-\s]V99[-\s]SN\d+)\b",
        r"\b((?:CI|IB|SI)[-\s]\d+[-\s][A-Z0-9.]+|[0-9]+[-\s]TA[-\s][A-Z0-9]+)\b",
    )
    for pattern in model_patterns:
        m = re.search(pattern, norm_blob)
        if m:
            model = re.sub(r"\s+", "-", m.group(1).strip())
            model = re.sub(
                r"-(AND|WITH|INCLUDES?|INLET|OUTLET|DISCHARGE|SILENCER|UNIT|VENDOR|PRODUCT).*$",
                "",
                model,
            )
            attrs["silencer_model"] = model
            break

    m = re.search(r"\b(\d{2,3})\s*DBA\b", norm_blob)
    if m:
        attrs["silencer_noise_target"] = f"{m.group(1)} DBA"
    pressure_blob = re.sub(r"\s+", " ", blob or norm_blob)
    m = re.search(
        r"\b(?:Pressure\s+Drop|total\s+pressure\s+drop(?:\s+of)?)\s*:?\s*(\d+(?:\.\d+)?|\.\d+)",
        pressure_blob,
        re.I,
    )
    if m and m.group(1):
        attrs["pressure_drop"] = m.group(1)

    features: List[str] = []
    for label, pattern in (
        ("PIEZOMETER TUBE", r"\bPIEZOMETER\s+TUBE\b"),
        ("VELOCITY TUBE", r"\bVELOCITY\s+TUBE\b"),
        ("THERMOWELL PORT", r"\bTHERMOWELL\s+PORT\b"),
        ("TRASH SCREEN", r"\bTRASH\s+SCREEN\b"),
        ("RAIN HOOD", r"\bRAIN\s+HOOD\b|\bRAINHOOD\b"),
        ("FILTER", r"\bFILTER\b"),
        ("SUPPORT LEGS", r"\b(SUPPORT|EXTENDED)\s+LEGS?\b"),
        ("HARDWARE/GASKET", r"\b(H\s*&\s*G|H\s+G|HARDWARE\s+AND\s+GASKET|GASKET)\b"),
        ("LIFTING LUGS", r"\bLIFTING\s+LUGS?\b"),
        ("MOUNTING HARDWARE", r"\bMOUNTING\s+HARDWARE\b"),
    ):
        if re.search(pattern, norm_blob):
            _add_unique(features, label)
    if features:
        attrs["silencer_feature"] = ", ".join(features)
    return attrs


def _spark_resistant_attributes(norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "SPARK RESISTANT" not in tags:
        return {}
    attrs = {"spark_resistant": "YES"}
    m = re.search(r"\bAMCA\s+(?:TYPE\s+)?([ABC])\b", norm_blob)
    if not m:
        m = re.search(r"\bSPARK\s+RESISTANT\s+CONSTRUCTION\s+([ABC])\b", norm_blob)
    if m:
        attrs["spark_resistant_type"] = f"AMCA {m.group(1).upper()}"
    return attrs


def _spare_parts_attributes(primary: str, norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "SPARE PARTS" not in tags:
        return {}
    attrs: Dict[str, str] = {}
    if re.match(r"^SPARE\b", primary):
        attrs["spare_part_type"] = "SPARE"
    elif re.match(r"^(?:NPO\s+)?REPAIR\b", primary):
        attrs["spare_part_type"] = "REPAIR"
    elif re.match(r"^(?:NPO\s+)?REPLACEMENT\b", primary):
        attrs["spare_part_type"] = "REPLACEMENT"

    component = ""
    if "BEARINGS" in tags:
        component = "BEARINGS"
    elif "V-BELT DRIVE" in tags or "DRIVE COMPONENTS" in tags:
        component = "V-BELT DRIVE"
    elif "WHEEL" in tags:
        component = "WHEEL"
    elif "INLET VANES" in tags:
        component = "IVC"
    elif "INLET" in tags or "INLET CONE" in norm_blob:
        component = "INLET CONE"
    elif re.search(r"\bROTOR\s+ASSEMBLY\b", norm_blob):
        component = "ROTOR ASSEMBLY"
    elif re.search(r"\bSHAFT\b", norm_blob):
        component = "SHAFT"
    elif re.search(r"\bFAN\s+KIT\b", norm_blob):
        component = "FAN KIT"
    elif re.search(r"\bFAN\b", norm_blob):
        component = "FAN"
    if component:
        attrs["spare_part_component"] = component
    return attrs


def _special_construction_attributes(primary: str, norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "SPECIAL CONSTRUCTION" not in tags:
        return {}
    attrs: Dict[str, str] = {}
    types: List[str] = []
    scopes: List[str] = []
    details: List[str] = []

    def add_type(value: str) -> None:
        _add_unique(types, value)

    def add_scope(value: str) -> None:
        _add_unique(scopes, value)

    def add_detail(value: str) -> None:
        _add_unique(details, value)

    if re.search(r"\bEFFECTIVE\s+DIAMETER\b", norm_blob):
        add_type("EFFECTIVE DIAMETER")
        m = re.search(r"\b(\d+(?:\.\d+)?)\s*%\s+EFFECTIVE\s+DIAMETER\b", norm_blob)
        if m:
            attrs["effective_diameter_percent"] = m.group(1)
    if re.search(r"\bCONTINUOUS\s+WELD\b", norm_blob):
        add_type("CONTINUOUS WELD")
        if "AIRSTREAM" in norm_blob:
            add_scope("AIRSTREAM")
        if "EXTERIOR" in norm_blob:
            add_scope("EXTERIOR")
    if re.search(r"\bAWS\s+D\d+(?:\.\d+)?\s+CODE\s+WELDING\b", norm_blob):
        add_type("CODE WELDING")
        codes = re.findall(r"\bAWS\s+D\d+(?:\.\d+)?\b", norm_blob)
        if codes:
            attrs["welding_code"] = ", ".join(dict.fromkeys(codes))
        if "STATIC COMPONENTS" in norm_blob:
            add_scope("STATIC COMPONENTS")
        if "ROTATING COMPONENTS" in norm_blob:
            add_scope("ROTATING COMPONENTS")
    if re.search(r"\bEARTHING\s+BOSS\b", norm_blob):
        add_type("EARTHING BOSS")
    if re.search(r"\bPRESSURE\s+TAP\b", norm_blob):
        add_type("PRESSURE TAP")
        if "OUTLET" in norm_blob:
            add_scope("OUTLET")
        elif "INLET" in norm_blob:
            add_scope("INLET")
    if re.search(r"\bTIE\s+ROD\s+SUPPORT\b", norm_blob):
        add_type("TIE ROD SUPPORT")
    if re.search(r"\bPLUG\s+PANEL\b", norm_blob):
        add_type("PLUG PANEL")
    if re.search(r"\bTHREADED\s+PLUG\b", norm_blob):
        add_type("THREADED PLUG")
        if "CONDUIT BOX" in norm_blob:
            add_scope("CONDUIT BOX")
        if "GUARD" in norm_blob:
            add_scope("GUARD")
    if re.search(r"\bRUN\s+SECOND\s+CONDUIT\b", norm_blob):
        add_type("AUXILIARY CONDUIT")
        add_detail("SECOND CONDUIT")
    if re.search(r"\bSET\s+SCREWS?\b", norm_blob):
        add_type("SET SCREWS")
        if re.search(r"\bDOUBLE\s+BLADE\s+PITCH\b", norm_blob):
            add_detail("DOUBLE BLADE PITCH")
    if re.search(r"\bLOC\s*TITE\b|\bLOCTITE\b", norm_blob):
        add_type("LOC TITE")
    if re.search(r"\bCAULK(?:ING)?\b", norm_blob):
        add_type("CAULKING")
        if "SILICONE FREE" in norm_blob:
            add_detail("SILICONE-FREE")
        if "WITHOUT CAULK" in norm_blob:
            add_detail("WITHOUT CAULK")
    if re.search(r"\bBUFFER\s+TUBE\b", norm_blob):
        add_type("BUFFER TUBE")
    if re.search(r"\bCAST\s+HUB\b", norm_blob):
        add_type("CAST HUB")
        if "STRAIGHT BORE" in norm_blob:
            add_detail("STRAIGHT BORE")
    if re.search(r"\bHOLE\s+DIAMETERS?\b", norm_blob):
        add_type("HOLE DIAMETERS")
    if re.search(r"\bOVERHANG\b", norm_blob):
        add_type("OVERHANG")
    if re.search(r"\bWELD\s+NUTS?\b", norm_blob):
        add_type("WELD NUTS")
        if "INLET" in norm_blob:
            add_scope("INLET")
    if re.search(r"\bVERTICAL\s+MOUNTING\s+PLATE\b", norm_blob):
        add_type("VERTICAL MOUNTING PLATE")
        add_scope("MOTOR CONDUIT BOX")

    if types:
        attrs["special_construction_type"] = ", ".join(types)
    if scopes:
        attrs["special_construction_scope"] = ", ".join(scopes)
    if details:
        attrs["special_construction_detail"] = ", ".join(details)
    return attrs


def _testing_attributes(primary: str, norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "TESTING" not in tags:
        return {}
    attrs: Dict[str, str] = {}
    types: List[str] = []
    statuses: List[str] = []
    measurements: List[str] = []

    def add_type(value: str) -> None:
        _add_unique(types, value)

    def add_status(value: str) -> None:
        _add_unique(statuses, value)

    def add_measurement(value: str) -> None:
        _add_unique(measurements, value)

    if re.search(r"\bMECHANICAL\s+RUN\s+TEST\b", norm_blob):
        add_type("MECHANICAL RUN TEST")
    elif re.search(r"\bRUN\s+TEST\b", norm_blob):
        add_type("RUN TEST")
    if re.search(r"\bSOAP\s+BUBBLE\s+PRESSURE\s+TEST\b", norm_blob):
        add_type("SOAP BUBBLE PRESSURE TEST")
    if re.search(r"\bOVERSPEED\s+TEST\b", norm_blob):
        add_type("OVERSPEED TEST")
    if re.search(r"\bDIMENSIONAL\s+INSPECTION\b", norm_blob):
        add_type("DIMENSIONAL INSPECTION")
    if re.search(r"\b(IEEE\s*112|ROUTINE\s+TEST)\b", norm_blob):
        add_type("MOTOR ROUTINE TEST")
    if re.search(r"\bAMP\s+DRAW\b", norm_blob):
        add_type("AMP DRAW MEASUREMENT")
        add_measurement("AMP DRAW")
    if re.search(r"\bRUN\s+TEST\s+REPORTS?\b|\bTEST\s+REPORTS?\b", norm_blob):
        add_type("TEST REPORT")

    if re.search(r"\bSTANDARD\b", primary):
        add_status("STANDARD")
    if re.search(r"\b(REQUIRED|CUSTOMER\s+WITNESS)\b", norm_blob):
        add_status("REQUIRED")
    if re.search(r"\bNOT\s+AVAILABLE\b", norm_blob):
        add_status("NOT AVAILABLE")
    if re.search(r"\bN\s*/?\s*A\b", norm_blob):
        add_status("N/A")
    if re.search(r"\bCUSTOMER\s+WITNESS\b", norm_blob):
        attrs["witnessed"] = "CUSTOMER"
    elif re.search(r"\bUNWITNESSED\b", norm_blob):
        attrs["witnessed"] = "NO"
    m = re.search(r"\b(\d+)\s+HOUR\b", norm_blob)
    if m:
        attrs["testing_duration"] = f"{m.group(1)} HOUR"
    if re.search(r"\bVIBRATION\s+READINGS?\b", norm_blob):
        add_measurement("VIBRATION READINGS")
    voltages = re.findall(r"\b\d{3,4}\s*V\b", norm_blob)
    if voltages:
        attrs["testing_voltage"] = ", ".join(dict.fromkeys(v.replace(" ", "") for v in voltages))
    if types:
        attrs["testing_type"] = ", ".join(types)
    if statuses:
        attrs["testing_status"] = ", ".join(statuses)
    if measurements:
        attrs["testing_measurements"] = ", ".join(measurements)
    return attrs


def _lining_attributes(norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "LINING" not in tags:
        return {}
    attrs: Dict[str, str] = {}
    scopes: List[str] = []
    if re.search(r"\bINLET\s+BOX\b", norm_blob):
        _add_unique(scopes, "INLET BOX")
    if re.search(r"\bHOUSING\s+SCROLL\b|\bSCROLL\b", norm_blob):
        _add_unique(scopes, "HOUSING SCROLL")
    if re.search(r"\bSIDE\s+SHEET\b", norm_blob):
        _add_unique(scopes, "SIDE SHEET")
    if re.search(r"\bWHEEL\s+BLADES?\b|\bBLADES?\b", norm_blob):
        _add_unique(scopes, "WHEEL BLADES")
    if "FIRMEX" in norm_blob:
        attrs["lining_type"] = "FIRMEX"
    elif re.search(r"\bRUBBER\s+LIN", norm_blob):
        attrs["lining_type"] = "RUBBER"
    if "ABRASION" in norm_blob:
        attrs["lining_service"] = "ABRASION"
    if scopes:
        attrs["lining_scope"] = ", ".join(scopes)
    return attrs


def _housing_attributes(norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "HOUSING" not in tags:
        return {}
    attrs: Dict[str, str] = {}
    subcategories: List[str] = []
    features: List[str] = []

    def add_subcategory(value: str) -> None:
        _add_unique(subcategories, value)

    def add_feature(value: str) -> None:
        _add_unique(features, value)

    if re.search(r"\bSQUARE\s+HOUSING\b|\bHOUSING\s+SQUARE\b", norm_blob):
        add_subcategory("SQUARE")
    if re.search(r"\bUNIVERSAL\s+HOUSING\b|\bHOUSING\s+UNIVERSAL\b", norm_blob):
        add_subcategory("UNIVERSAL")
    if "HEAVY DUTY" in tags:
        add_subcategory("HEAVY DUTY")
    if "LINING" in tags or re.search(r"\b(FIRMEX|LINERS?|LINING)\b.*\bHOUSING\b", norm_blob):
        add_subcategory("LINING")

    if re.search(r"\b(FLANGE\s+THICKNESS|BOLT\s+PATTERN|REINFORCING\s+GUSSETS?)\b", norm_blob):
        add_subcategory("MODIFIED FLANGE")
        if "FLANGE THICKNESS" in norm_blob:
            add_feature("FLANGE THICKNESS")
        if "BOLT PATTERN" in norm_blob:
            add_feature("BOLT PATTERN")
        if re.search(r"\bREINFORCING\s+GUSSETS?\b", norm_blob):
            add_feature("REINFORCING GUSSETS")

    if re.search(r"\bHOUSING\s+LENGTH\b", norm_blob) and re.search(r"\bFLANGE\s+TO\s+FLANGE\b", norm_blob):
        add_subcategory("DIMENSION NOTE")
        attrs["housing_dimension"] = "FLANGE-TO-FLANGE LENGTH"

    if re.search(r"\b(?:\d+\s*GA|[0-9/.\-]+\s*\")\s+HOUSING\s+THICKNESS\b|\bHOUSING\s+THICKNESS\b", norm_blob):
        add_subcategory("THICKNESS")
        m = re.search(r"\b(\d+\s*GA)\s+HOUSING\s+THICKNESS\b", norm_blob)
        if m:
            attrs["housing_thickness"] = re.sub(r"\s+", " ", m.group(1)).strip()

    if re.search(r"\bHOUSING\s+STIFF(?:E|NE)R\s+PLATES?\b", norm_blob):
        add_subcategory("STIFFENER PLATES")
    if "CASING EXTENSION" in norm_blob:
        add_subcategory("CASING EXTENSION")
        if "OUTLET" in norm_blob:
            attrs["used_on"] = "OUTLET"
    if re.search(r"\b(TAP\s+DRIVE\s+SIDE\s+HOUSING|TAP\s+.*HOUSING\s+HALF)\b", norm_blob):
        add_subcategory("TAPPED HOUSING")
    if re.search(r"\b(MOUNT\s+HOUSING|HOUSING\s+MOUNTING|HOUSING\s+TO\s+DRIVE\s+COVER)\b", norm_blob):
        add_subcategory("MOUNTING/SUPPORT")
    if re.search(r"\b(NAMEPLATE\b.*\bHOUSING|RIVETED\s+TO\s+HOUSING)\b", norm_blob):
        add_subcategory("NAMEPLATE LOCATION")
        attrs["mount_location"] = "HOUSING"

    if subcategories:
        attrs["housing_subcategory"] = ", ".join(subcategories)
    if features:
        attrs["housing_feature"] = ", ".join(features)
    return attrs


def _motor_conduit_box_attributes(primary: str, norm_blob: str) -> Dict[str, str]:
    if not _is_motor_conduit_box_context(primary, norm_blob):
        return {}
    attrs = {
        "component": "MOTOR",
        "motor_feature": "CONDUIT BOX LOCATION",
    }
    if "HUGGING" in norm_blob:
        attrs["motor_conduit_box_location"] = "HUGGING HOUSING"
    elif re.search(r"\b(CLOSE\s+TO|AS\s+CLOSE\s+TO)\b", norm_blob):
        attrs["motor_conduit_box_location"] = "CLOSE TO HOUSING"
    elif re.search(r"\b(TACK\s+AND\s+WELD|WELD\s+CONDUIT\s+BOX|CONDUIT\s+BOX\s+TO\s+HOUSING)\b", norm_blob):
        attrs["motor_conduit_box_location"] = "MOUNTED TO HOUSING"
    else:
        m = re.search(r"\b(F[123]|RF)\s+CONDUIT\s+BOX\b", norm_blob)
        if m:
            attrs["motor_conduit_box_location"] = m.group(1).upper()
    m = re.search(r"(?:@\s*)?([0-9]{1,2})\s*:?\s*00\b", norm_blob)
    if m and "VIEWED FROM OUTLET" in norm_blob:
        attrs["motor_conduit_box_position"] = f"{m.group(1)}:00 VIEWED FROM OUTLET"
    if re.search(r"\bKNOCKOUT\s+FACES?.{0,20}\bDOWNWARD\b", norm_blob):
        attrs["motor_conduit_box_orientation"] = "KNOCKOUT FACES DOWNWARD"
    return attrs


def _flex_connector_attributes(norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "FLEX CONNECTOR" not in tags:
        return {}
    attrs: Dict[str, str] = {"component": "FLEX CONNECTOR"}
    if re.search(r"\b(EXPANSION\s+JOINT|EJ)\b", norm_blob):
        attrs["flex_connector_type"] = "EXPANSION JOINT"
    elif re.search(r"\bFLEXIBLE\s+CONNECTOR\b", norm_blob):
        attrs["flex_connector_type"] = "FLEXIBLE CONNECTOR"
    elif re.search(r"\bFLEX\s+CONNECTOR\b", norm_blob):
        attrs["flex_connector_type"] = "FLEX CONNECTOR"
    if re.search(r"\bFLOW\s+LINERS?\b", norm_blob):
        attrs["flex_connector_feature"] = "FLOW LINER"
    if re.search(r"\bFIBERGLASS\s+SOUND\s+PILLOW\b", norm_blob):
        attrs["flex_connector_insulation"] = "FIBERGLASS SOUND PILLOW"
    used_on: List[str] = []
    if "INLET" in tags or re.search(r"\bINLET\b", norm_blob):
        _add_unique(used_on, "INLET")
    if "OUTLET" in tags or re.search(r"\bOUTLET\b", norm_blob):
        _add_unique(used_on, "OUTLET")
    if used_on:
        attrs["used_on"] = ", ".join(used_on)
    return attrs


def _coupling_attributes(norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "COUPLING" not in tags:
        return {}
    attrs: Dict[str, str] = {
        "component": "COUPLING",
        "coupling_subcategory": "FLEXIBLE COUPLING",
    }
    if "FALK" in norm_blob or "STEELFLEX" in norm_blob:
        attrs["manufacturer"] = "FALK"
        attrs["coupling_type"] = "FALK TYPE T STEELFLEX"
        attrs["model"] = "TYPE T STEELFLEX"
    if "REXNORD" in norm_blob:
        attrs["manufacturer"] = "REXNORD"
        m = re.search(r"\bREXNORD\s+(THOMAS\s+SERIES\s+\d+)\b", norm_blob)
        attrs["coupling_type"] = f"REXNORD {m.group(1)}" if m else "REXNORD THOMAS SERIES"
        attrs["model"] = attrs["coupling_type"].replace("REXNORD ", "")
    if re.search(r"\bHALF\s+COUPLING\b", norm_blob):
        attrs["coupling_type"] = "HALF COUPLING"
    m = re.search(r"\bSIZE\s+([A-Z0-9-]+)\b", norm_blob)
    if m:
        attrs["size"] = m.group(1).upper()
    if "CLEARANCE" in norm_blob:
        attrs["fit"] = "CLEARANCE"
    elif "INTERFERENCE" in norm_blob:
        attrs["fit"] = "INTERFERENCE"
    m = re.search(r"\b(HORIZONTAL\s+SPLIT\s+COVER\s+T10)\b", norm_blob)
    if m:
        attrs["cover_type"] = m.group(1)
    if re.search(r"\bSET\s+SCREWS?\b", norm_blob):
        attrs["set_screws"] = "YES"
    if "CBC MOUNT" in norm_blob:
        attrs["mounting"] = "CBC MOUNT"
    return attrs


def _low_leakage_attributes(norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "LOW LEAKAGE" not in tags and not re.search(r"\bLOW\s*[- ]?LEAK(?:AGE)?\b", norm_blob):
        return {}
    attrs = {"leakage_class": "LOW LEAKAGE"}
    if "IVC" in norm_blob or "INLET VOLUME CONTROL" in norm_blob or "INLET VANES" in tags:
        attrs["used_on"] = "IVC"
    return attrs


def _temperature_values(norm_blob: str) -> List[Tuple[int, str]]:
    values: List[Tuple[int, str]] = []
    for m in re.finditer(r"\bTO\s*(-?\d{2,3})\s*(?:(?:DEG(?:REE)?|°)\s*)?([FC])\b", norm_blob, re.I):
        value = int(m.group(1))
        unit = m.group(2).upper()
        if (value, unit) not in values:
            values.append((value, unit))
    for m in _TEMP_VALUE.finditer(norm_blob):
        context = norm_blob[max(0, m.start() - 35):m.end() + 35]
        if not (
            "°" in m.group(0)
            or
            re.search(r"\b(DEG|TEMPERATURE|TEMP|AMBIENT|OPERATION|SERVICE|SUITABLE|RATED|MAX)\b", context)
            or re.search(r"\b(FOR|AT|TO)\s*$", norm_blob[max(0, m.start() - 10):m.start()])
            or re.search(r"\bTO\b", norm_blob[m.end():m.end() + 10])
        ):
            continue
        value = int(m.group(1))
        unit = m.group(2).upper()
        if (value, unit) not in values:
            values.append((value, unit))
    return values


def _extreme_temperature_values(norm_blob: str) -> List[Tuple[int, str]]:
    values: List[Tuple[int, str]] = []
    for value, unit in _temperature_values(norm_blob):
        if value < 0 or (unit == "F" and value >= 130) or (unit == "C" and value >= 50):
            values.append((value, unit))
    return values


def _temperature_label(value: int, unit: str) -> str:
    return f"{value}{unit}"


def _temperature_direction(norm_blob: str, values: List[Tuple[int, str]]) -> str:
    directions: List[str] = []
    if re.search(r"\bLOW\s+TEMP(?:ERATURE)?\b", norm_blob) or any(value < 0 for value, _ in values):
        _add_unique(directions, "LOW TEMPERATURE")
    if (re.search(r"\b(HIGH\s+TEMP(?:ERATURE)?|HEAT\s+FAN)\b", norm_blob)
            or any((unit == "F" and value >= 130) or (unit == "C" and value >= 50)
                   for value, unit in values)):
        _add_unique(directions, "HIGH TEMPERATURE")
    return ", ".join(directions)


def _temperature_component(primary: str, norm_blob: str, tags: set[str]) -> str:
    if "SHAFT SEAL" in tags:
        return "SHAFT SEAL"
    if _is_top_level_extreme_temperature(primary, norm_blob, tags):
        return ""
    if "MOTOR" in tags and not re.match(r"^BASE\s+FAN\b", primary, re.I):
        return "MOTOR"
    return ""


def _is_top_level_extreme_temperature(primary: str, norm_blob: str, tags: set[str]) -> bool:
    if re.match(r"^MOTOR\b", primary, re.I) or "SHAFT SEAL" in tags:
        return False
    if "BASE FAN" in tags or re.search(r"\bFANS?\b", norm_blob):
        return True
    if ("INLET VANES" in tags or "DAMPER" in tags or "SPARK RESISTANT" in tags
            or "IVC" in norm_blob or "INLET VOLUME CONTROL" in norm_blob):
        return True
    return False


def _temperature_attributes(primary: str, norm_blob: str, tags: set[str], raw_blob: str) -> Dict[str, str]:
    temp_blob = f"{norm_blob} {raw_blob.upper()}"
    values = _extreme_temperature_values(temp_blob)
    direction = _temperature_direction(temp_blob, values)
    if not direction and _EXTREME_TEMPERATURE_TAG not in tags:
        return {}

    attrs: Dict[str, str] = {}
    component = _temperature_component(primary, norm_blob, tags)
    if component:
        attrs["component"] = component
    if direction:
        attrs["temperature_service"] = direction
    elif _is_top_level_extreme_temperature(primary, norm_blob, tags):
        attrs["temperature_service"] = "EXTREME TEMP"
    if direction and "," not in direction:
        attrs["temperature_direction"] = direction
    if values:
        attrs["temperature_rating"] = ", ".join(_temperature_label(value, unit) for value, unit in values)
    if re.search(r"\bHIGH\s+TEMP(?:ERATURE)?\s+GREASE\b", temp_blob):
        attrs["grease_type"] = "HIGH TEMPERATURE GREASE"
    elif re.search(r"\bLOW\s+TEMP(?:ERATURE)?\s+GREASE\b", temp_blob):
        attrs["grease_type"] = "LOW TEMPERATURE GREASE"
    return attrs


def _heavy_duty_attributes(norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "HEAVY DUTY" not in tags and not re.search(r"\bHEAVY\s+DUTY\b", norm_blob):
        return {}
    attrs = {"duty_rating": "HEAVY DUTY"}
    if "WHEEL" in tags or re.search(r"\bWHEEL\b", norm_blob):
        attrs["component"] = "WHEEL"
    elif "HOUSING" in tags or re.search(r"\bHOUSING\b", norm_blob):
        attrs["component"] = "HOUSING"
    return attrs


def _is_number_token(value: str) -> bool:
    return bool(re.fullmatch(r"\d+(?:\.\d+)?", value))


def _set_sheave_attrs(attrs: Dict[str, str], prefix: str, sheave: str, bushing: str = "") -> None:
    sheave = sheave.strip(" ,;")
    bushing = bushing.strip(" ,;")
    if not sheave:
        return
    attrs[f"{prefix}_sheave"] = sheave
    combo = sheave
    if bushing:
        attrs[f"{prefix}_bushing"] = bushing
        combo = f"{sheave} {bushing}"
    combo_key = "drive_sheave_bushing" if prefix == "drive" else "driven_sheave_bushing"
    attrs.setdefault(combo_key, combo)


def _set_sheave_attrs_from_text(attrs: Dict[str, str], prefix: str, value: str) -> None:
    value = re.sub(r"\s+", " ", value).strip(" ,;")
    if not value:
        return
    combo_key = "drive_sheave_bushing" if prefix == "drive" else "driven_sheave_bushing"
    attrs[combo_key] = value

    m = re.match(r"^(?P<sheave>[A-Z0-9]+)\s*/\s*(?P<bushing>.+)$", value, re.I)
    if m:
        _set_sheave_attrs(attrs, prefix, m.group("sheave"), m.group("bushing"))
        return
    parts = value.split()
    if len(parts) >= 2:
        _set_sheave_attrs(attrs, prefix, parts[0], " ".join(parts[1:]))
    else:
        _set_sheave_attrs(attrs, prefix, value)


def _is_drive_table_line(primary: str) -> bool:
    return bool(re.match(r"^\*?\s*\d{3,4}\s*/?\s+\d{3,4}\s+[A-Z]{1,3}\d+\s+\d+\b", primary, re.I))


def _drive_table_attributes(item: Dict[str, Any]) -> Dict[str, str]:
    raw = re.sub(r"\s+", " ", str(item.get("raw", ""))).strip()
    tokens = raw.split()
    attrs: Dict[str, str] = {}
    if len(tokens) < 8:
        return attrs

    selected = False
    rpm_token = tokens[0]
    if rpm_token.startswith("*"):
        selected = True
        rpm_token = rpm_token.lstrip("*")
    rpm_match = re.fullmatch(r"(\d{3,4})/(\d{3,4})", rpm_token)
    if not rpm_match:
        return attrs
    if not re.fullmatch(r"[A-Z]{1,3}\d+", tokens[1], re.I) or not tokens[2].isdigit():
        return attrs

    if all(_is_number_token(tok) for tok in tokens[-3:]):
        sf_index = len(tokens) - 3
    elif all(_is_number_token(tok) for tok in tokens[-2:]):
        sf_index = len(tokens) - 2
    else:
        return attrs

    sheave_tokens = tokens[3:sf_index]
    if len(sheave_tokens) < 2:
        return attrs

    attrs["drive_rpm"] = rpm_match.group(1)
    attrs["driven_rpm"] = rpm_match.group(2)
    attrs["belt"] = tokens[1].upper()
    attrs["belt_qty"] = tokens[2]

    if len(sheave_tokens) == 2:
        _set_sheave_attrs(attrs, "drive", sheave_tokens[0])
        _set_sheave_attrs(attrs, "driven", sheave_tokens[1])
    elif len(sheave_tokens) == 3:
        _set_sheave_attrs(attrs, "drive", sheave_tokens[0])
        _set_sheave_attrs(attrs, "driven", sheave_tokens[1], sheave_tokens[2])
    else:
        _set_sheave_attrs(attrs, "drive", sheave_tokens[0], sheave_tokens[1])
        _set_sheave_attrs(attrs, "driven", sheave_tokens[2], " ".join(sheave_tokens[3:]))

    attrs["actual_sf"] = tokens[sf_index]
    attrs["actual_cd"] = tokens[sf_index + 1]
    blob = _item_blob(item)
    if selected or re.search(r"\*\s*selected\s+drive\b", blob, re.I):
        attrs["selected_drive"] = "YES"
    return attrs


def _balance_attributes(norm_blob: str) -> Dict[str, str]:
    attrs: Dict[str, str] = {}
    types: List[str] = []
    grades: List[str] = []
    for m in _BALANCE_GRADE.finditer(norm_blob):
        _add_unique(grades, f"G{m.group(1).upper()}")
    if grades:
        _add_unique(types, "GRADED BALANCE")
        attrs["balance_grade"] = ", ".join(grades)
    if re.search(r"\bWELDED\s+BALANCE\s+WEIGHTS?\b", norm_blob):
        _add_unique(types, "WELDED BALANCE WEIGHTS")
    if types:
        attrs["balance_type"] = ", ".join(types)
    return attrs


def _is_housing_drain(norm_blob: str) -> bool:
    return bool(re.search(r"\bHOUSING\s+DRAINS?\b", norm_blob))


def _is_inlet_box_drain(norm_blob: str) -> bool:
    return bool(re.search(r"\bINLET\s+BOX\b", norm_blob) and "DRAIN" in norm_blob)


def _is_motor_drain(primary: str, norm_blob: str) -> bool:
    if "DRAIN" not in norm_blob:
        return False
    return bool(
        re.match(r"^MOTOR\b", primary, re.I)
        or re.search(r"\b(CONDENSATION\s+DRAIN|CONDUIT\s+BOX\s+DRAIN|DRAIN\s+HOLES?)\b", norm_blob)
    )


def _drain_type(primary: str, norm_blob: str) -> str:
    if "DRAIN" not in norm_blob:
        return ""
    if _is_housing_drain(norm_blob):
        return "HOUSING DRAIN"
    if _is_inlet_box_drain(norm_blob):
        return "INLET BOX DRAIN"
    if _is_motor_drain(primary, norm_blob):
        return "MOTOR CONDUIT BOX DRAIN"
    return "DRAIN"


def _drain_attributes(primary: str, norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if "DRAIN" not in tags and "DRAIN" not in norm_blob:
        return {}
    drain_type = _drain_type(primary, norm_blob)
    if not drain_type:
        return {}
    attrs = {"drain_type": drain_type}
    if "PLUG" in norm_blob:
        attrs["drain_closure"] = "PLUG"
    if "CONDENSATION DRAIN" in norm_blob or re.search(r"\bDRAIN\s+HOLES?\b", norm_blob):
        attrs["drain_detail"] = "CONDENSATION DRAIN HOLES"
    return attrs


def _bearing_attributes(blob: str, norm_blob: str) -> Dict[str, str]:
    attrs: Dict[str, str] = {}
    for label, pattern in (
        ("SPLIT PILLOW BLOCK", r"\bBEARINGS?\s+SPLIT\s+PILLOW\s+BLOCK\b"),
        ("STANDARD", r"\bBEARINGS?\s+STANDARD\b"),
        ("REPAIR BEARINGS", r"\bREPAIR\s+BEARINGS?\b"),
        ("SPARE BEARINGS", r"\bSPARE\s+BEARINGS?\b"),
        ("BEARING ADDER", r"\bBEARING\s+ADDER\b"),
    ):
        if re.search(pattern, norm_blob):
            attrs["bearing_type"] = label
            break

    m = re.search(
        r"\b(\d+(?:[-\s]\d+/\d+)?|\d+/\d+)\s*(?:\"|in(?:ch(?:es)?)?\.?)?\s*bore\b",
        blob,
        re.I,
    )
    if m:
        attrs["bearing_bore"] = re.sub(r"\s+", "-", m.group(1).strip()) + '"'
    return attrs


def _material_attributes(norm_blob: str) -> Dict[str, str]:
    """Material roll-up with grade and where the material applies."""
    attrs: Dict[str, str] = {}
    materials: List[str] = []
    passivation = bool(re.search(r"\bPASSIVAT", norm_blob))
    if re.search(r"\bALUMINUM\b", norm_blob):
        _add_unique(materials, "ALUMINUM")
    if re.search(r"\bSTAINLESS\s+STEEL\b", norm_blob) or passivation:
        _add_unique(materials, "STAINLESS STEEL")
    if materials:
        attrs["material"] = ", ".join(materials)
    if passivation:
        attrs["material_treatment"] = "PASSIVATION OF WELDS" if "WELDS" in norm_blob else "PASSIVATION"

    grades: List[str] = []
    for m in _MATERIAL_GRADE.finditer(norm_blob):
        grade = f"{m.group(1).upper()} SS"
        if grade not in grades:
            grades.append(grade)
    if grades:
        attrs["material_grade"] = ", ".join(grades)
    if not attrs:
        return {}

    scopes: List[str] = []

    def add_scope(scope: str) -> None:
        if scope not in scopes:
            scopes.append(scope)

    if (re.search(r"\bHOUSING\s+(AND|&)\s+BASE\b", norm_blob)
            or re.search(r"\bBASE\s+(AND|&)\s+HOUSING\b", norm_blob)
            or ("HOUSING" in norm_blob and "BASE" in norm_blob and "BASE FAN" not in norm_blob)):
        add_scope("HOUSING AND BASE")
    if "BASE FAN" in norm_blob:
        add_scope("BASE FAN")
    if "AIRSTREAM" in norm_blob:
        add_scope("AIRSTREAM")
    if "EXTERIOR" in norm_blob:
        add_scope("EXTERIOR")
    if re.search(r"\bWHEEL\s+AND\s+HUB\b", norm_blob):
        add_scope("WHEEL AND HUB")
    elif "WHEEL" in norm_blob:
        add_scope("WHEEL")
    if re.search(r"\bINLET\s+CONE\b", norm_blob):
        add_scope("INLET CONE")
    elif "INLET" in norm_blob and not _is_inlet_box_drain(norm_blob):
        add_scope("INLET")
    if "OUTLET" in norm_blob or "DISCHARGE" in norm_blob:
        add_scope("OUTLET")
    if passivation and "WELDS" in norm_blob:
        add_scope("WELDS")
    if "ACCESS DOOR" in norm_blob:
        add_scope("ACCESS DOOR")
    primary = norm_blob.split(";", 1)[0].strip()
    if _is_housing_drain(norm_blob):
        add_scope("HOUSING DRAIN")
    if _is_inlet_box_drain(norm_blob):
        add_scope("INLET BOX DRAIN")
    if _is_motor_drain(primary, norm_blob):
        add_scope("MOTOR CONDUIT BOX DRAIN")
    if "HOUSING" in norm_blob and not any(s.startswith("HOUSING") for s in scopes):
        add_scope("HOUSING")
    for scope, pattern in (
        ("DRAIN", r"\bDRAIN\b"),
        ("SHAFT COOLER", r"\bSHAFT\s+COOLER\b"),
        ("SHAFT SEAL", r"\bSHAFT\s+SEAL\b"),
        ("SHAFT SLEEVE", r"\bSHAFT\s+SLEEVE\b"),
        ("SCREEN", r"\bSCREEN\b"),
        ("FLEX CONNECTOR", r"\b(FLEX(IBLE)?\s+CONNECTOR|EXPANSION\s+JOINT|EJ)\b"),
        ("SILENCER", r"\bSILENCER\b"),
        ("NAMEPLATE", r"\bNAMEPLATE\b"),
        ("HARDWARE", r"\bHARDWARE\b"),
        ("TUBING", r"\bTUBING\b"),
        ("BACKING RINGS", r"\bBACKING\s+RINGS?\b"),
        ("BACKING BARS", r"\bBACKING\s+BARS?\b"),
        ("MOTOR", r"^MOTOR\b"),
    ):
        if scope == "DRAIN" and any(s.endswith("DRAIN") for s in scopes):
            continue
        if re.search(pattern, norm_blob):
            add_scope(scope)
    if "SHAFT" in norm_blob and not {"SHAFT COOLER", "SHAFT SEAL", "SHAFT SLEEVE"} & set(scopes):
        add_scope("SHAFT")
    if _is_nameplate_housing_mount(norm_blob):
        scopes = [s for s in scopes if s != "HOUSING"]
    if scopes:
        attrs["material_scope"] = ", ".join(scopes)
    return attrs


def _component_material_owner(item: Dict[str, Any], material_attrs: Dict[str, str],
                              rules: Dict[str, Any] | None = None) -> str:
    if not material_attrs:
        return ""
    primary = _primary_norm(item, rules)
    product = _label_value(item, "Product").upper()
    if "ACTUATOR" in primary or product == "ACTUATOR":
        return "ACTUATOR"
    if re.match(r"^MOTOR\b", primary) and not re.match(r"^MOTOR\s+MOUNTING\b", primary):
        return "MOTOR"
    return ""


def _component_material_attributes(owner: str, material_attrs: Dict[str, str]) -> Dict[str, str]:
    attrs = dict(material_attrs)
    scopes = [s.strip() for s in attrs.get("material_scope", "").split(",") if s.strip()]
    if owner == "ACTUATOR":
        component_scopes = [s for s in scopes if s in {"TUBING", "HARDWARE"}]
        attrs["material_scope"] = ", ".join(component_scopes or ["ACTUATOR"])
    elif owner == "MOTOR":
        component_scopes = [s for s in scopes if s == "MOTOR CONDUIT BOX DRAIN"]
        attrs["material_scope"] = ", ".join(component_scopes or ["MOTOR"])
    return attrs


def _is_nameplate_housing_mount(norm_blob: str) -> bool:
    return bool(
        "NAMEPLATE" in norm_blob
        and "HOUSING" in norm_blob
        and re.search(r"\b(RIVETED\s+TO\s+HOUSING|NAMEPLATE\b.*\bHOUSING)\b", norm_blob)
    )


def _is_shaft_bearing_guard_line(primary: str) -> bool:
    return bool(
        re.search(r"\bshaft\b.*\bbearing\b.*\bguard\b", primary, re.I)
        or re.search(r"\bguard\b.*\bshaft\b.*\bbearing\b", primary, re.I)
    )


def _is_belt_guard_line(primary: str) -> bool:
    return bool(re.search(r"\bBELT\s+GUARD\b", primary, re.I))


def _is_base_fan_line(primary: str) -> bool:
    return bool(re.match(r"^BASE\s+FAN\b", primary, re.I))


def _is_paint_line(primary: str) -> bool:
    return bool(re.match(r"^PAINT\b", primary, re.I))


def _is_assembly_note(primary: str) -> bool:
    return bool(re.match(r"^ASSEMBLY\b", primary, re.I))


def _is_admin_note(primary: str) -> bool:
    return bool(re.search(r"\b(SLOW\s+PAY\s+ADDITION|PAYMODE[-\s]*X|FEE|CHARGE)\b", primary, re.I))


def _is_fan_coating_line(primary: str, tags: set[str]) -> bool:
    if _is_paint_line(primary) or re.match(r"^(SPECIAL\s+PAINT|PRE\s+COATING)\b", primary, re.I):
        return True
    if re.search(r"\b(AIRSTREAM|EXTERIOR)\b", primary, re.I) and _COATING_WORD.search(primary):
        return True
    if {"BASE FAN", "WHEEL", "INLET", "OUTLET", "HOUSING"} & tags and _COATING_WORD.search(primary):
        return True
    return False


def _is_accessory_coating(primary: str, tags: set[str], norm_blob: str) -> bool:
    if "COATING" not in tags:
        return False
    if "SILENCER" in tags and _COATING_WORD.search(norm_blob):
        return True
    if _is_fan_coating_line(primary, tags):
        return False
    if _is_belt_guard_line(primary) or _is_shaft_bearing_guard_line(primary):
        return True
    if primary.startswith("MOTOR"):
        return True
    return bool((_ACCESSORY_COATING_TAGS & tags) and _COATING_WORD.search(norm_blob))


def _lube_component_tags(primary: str, norm_blob: str) -> List[str]:
    if not (_LUBE_ACCESSORY.search(primary) or _LUBE_ACCESSORY.search(norm_blob)):
        return []
    tags: List[str] = []
    if re.search(r"\bMOTOR\b", norm_blob, re.I):
        _add_unique(tags, "MOTOR")
    if re.search(r"\bBEARINGS?\b", norm_blob, re.I):
        _add_unique(tags, "BEARINGS")
    if not tags:
        tags = ["BEARINGS", "MOTOR"]
    return tags


def _lube_review_attributes(primary: str, norm_blob: str) -> Dict[str, str]:
    if not (_LUBE_ACCESSORY.search(primary) or _LUBE_ACCESSORY.search(norm_blob)):
        return {}
    has_motor = bool(re.search(r"\bMOTOR\b", norm_blob, re.I))
    has_bearing = bool(re.search(r"\bBEARINGS?\b", norm_blob, re.I))
    if has_motor or has_bearing:
        return {}
    return {"component_review": "UNCLEAR GREASE TARGET - VERIFY MOTOR/BEARINGS/ARRANGEMENT"}


def _guard_attributes(primary: str, norm_blob: str) -> Dict[str, str]:
    attrs: Dict[str, str] = {}
    if _is_belt_guard_line(primary):
        attrs["component"] = "BELT GUARD"
        if "PROVIDED BY OTHERS" in norm_blob:
            attrs["supplied_by"] = "OTHERS"
        if "CBC MOUNT" in norm_blob:
            attrs["mounting"] = "CBC MOUNT"
        if "HINGED" in norm_blob:
            attrs["guard_type"] = "HINGED"
        if "STANDARD STEEL" in norm_blob:
            attrs["guard_material"] = "STANDARD STEEL"
        if "TACH HOLE IN GUARD WITH PLUG" in norm_blob:
            attrs["tach_hole"] = "WITH PLUG"
        elif "TACH HOLE IN GUARD NONE" in norm_blob:
            attrs["tach_hole"] = "NONE"
        locs: List[str] = []
        if "FAN END" in norm_blob:
            locs.append("FAN END")
        if "MOTOR END" in norm_blob:
            locs.append("MOTOR END")
        if locs:
            attrs["tach_hole_location"] = ", ".join(locs)
    elif _is_shaft_bearing_guard_line(primary):
        attrs["component"] = "SHAFT/BEARING/COUPLING GUARD"
        if "PROVIDED BY OTHERS" in norm_blob:
            attrs["supplied_by"] = "OTHERS"
        if "STANDARD STEEL" in norm_blob:
            attrs["guard_material"] = "STANDARD STEEL"
    return attrs


def _coating_scope(primary: str, tags: set[str], norm_blob: str) -> str:
    scopes: List[str] = []
    for tag in ("BELT GUARD", _GUARD_TAG, "MOTOR", "SILENCER", "FLEX CONNECTOR",
                "WEATHER COVER", "SCREEN", "WHEEL", "INLET", "OUTLET", "HOUSING"):
        if tag == "MOTOR" and "MOTOR BASE" in norm_blob and not primary.startswith("MOTOR"):
            continue
        if tag in tags:
            _add_unique(scopes, tag)
    for scope, pattern in (
        ("AIRSTREAM", r"\bAIRSTREAM\b"),
        ("EXTERIOR", r"\bEXTERIOR\b"),
        ("INTERIOR", r"\bINTERIOR\b"),
        ("MOTOR BASE", r"\bMOTOR\s+BASE\b"),
        ("CHANNEL BASE", r"\bCHANNEL\s+BASE\b"),
        ("BEARING BASE", r"\bBEARING\s+BASE\b"),
    ):
        if re.search(pattern, norm_blob):
            _add_unique(scopes, scope)
    return ", ".join(scopes)


def _ral_color(match: re.Match[str]) -> str:
    words = [
        w for w in (match.group(2), match.group(3))
        if w and w not in _RAL_STOP_WORDS
    ]
    return " ".join(["RAL", match.group(1), *words])


def _coating_category(primary: str, norm_blob: str, coating_type: str | None) -> str:
    if re.search(r"\b(UPDATED\s+)?COATING\s+NOTE\b", norm_blob):
        return "COATING NOTE"
    if re.match(r"^PRE\s+COATING\b", primary, re.I):
        return "PRE-COATING PROCESS"
    if "VEGETABLE OIL" in norm_blob:
        return "SPECIAL COATING"
    if re.match(r"^SPECIAL\s+PAINT\b", primary, re.I):
        return "SPECIAL COATING"
    if any(term in norm_blob for term in ("PLASITE", "HERESITE")):
        return "SPECIAL COATING"
    if "UNPAINTED" in norm_blob:
        return "UNPAINTED"
    if coating_type in {"EPOXY", "PRIMER"}:
        return "EPOXY/PRIMER"
    if coating_type in {"PAINT", "ENAMEL"}:
        return "PAINT"
    return coating_type or "COATING"


def _coating_attributes(primary: str, norm_blob: str, tags: set[str], raw_tags: set[str] | None = None) -> Dict[str, str]:
    if not _COATING_WORD.search(norm_blob):
        return {}
    attrs: Dict[str, str] = {}
    context_tags = raw_tags or tags
    accessory = (
        _is_accessory_coating(primary, context_tags, norm_blob)
        or ("SILENCER" in tags and _COATING_WORD.search(norm_blob))
        or _is_belt_guard_line(primary)
    )
    attrs["coating_context"] = "ACCESSORY" if accessory else "FAN"
    if _is_belt_guard_line(primary):
        scope = "BELT GUARD"
    elif _is_shaft_bearing_guard_line(primary):
        scope = "SHAFT/BEARING/COUPLING GUARD"
    else:
        scope = _coating_scope(primary, tags, norm_blob)
    if scope:
        attrs["coating_scope"] = scope

    if "UNPAINTED" in norm_blob:
        attrs["coating_state"] = "UNPAINTED"
    coating_type = None
    if "GALVANIZED" in norm_blob or "GALVANIZING" in norm_blob:
        coating_type = "GALVANIZED"
    elif "VEGETABLE OIL" in norm_blob:
        coating_type = "VEGETABLE OIL"
    elif "PLASITE" in norm_blob:
        coating_type = "PLASITE"
    elif "HERESITE" in norm_blob:
        coating_type = "HERESITE"
    elif "EPOXY" in norm_blob:
        coating_type = "EPOXY"
    elif "PRIMER" in norm_blob or "ZINC" in norm_blob:
        coating_type = "PRIMER"
    elif "ENAMEL" in norm_blob:
        coating_type = "ENAMEL"
    elif re.search(r"\bPAINT", norm_blob):
        coating_type = "PAINT"
    elif re.search(r"\bCOAT", norm_blob):
        coating_type = "COATING"
    if coating_type:
        attrs["coating_type"] = coating_type
    attrs["coating_category"] = _coating_category(primary, norm_blob, coating_type)
    if attrs["coating_category"] == "PRE-COATING PROCESS":
        attrs["coating_process"] = "PRE-COATING ASSEMBLY/DISASSEMBLY"

    if "SAFETY YELLOW" in norm_blob:
        attrs["coating_color"] = "SAFETY YELLOW"
    elif "STANDARD CBC BLACK" in norm_blob:
        attrs["coating_color"] = "STANDARD CBC BLACK"
    else:
        ral_colors = [
            _ral_color(match)
            for match in re.finditer(r"\bRAL\s+(\d{4})(?:\s+([A-Z]+)(?:\s+([A-Z]+))?)?", norm_blob)
        ]
        if ral_colors:
            attrs["coating_color"] = ral_colors[0]
            if len(ral_colors) > 1:
                attrs["alternate_coating_color"] = ", ".join(ral_colors[1:])
        elif "IMPERIAL GRAY" in norm_blob:
            attrs["coating_color"] = "IMPERIAL GRAY"
        elif "BLACK" in norm_blob:
            attrs["coating_color"] = "BLACK"

    m = re.search(r"\bCOATS?\s*:?\s*(\d+)\b", norm_blob)
    if m:
        attrs["coats"] = m.group(1)
    return attrs


def _final_tags(item: Dict[str, Any], rules: Dict[str, Any] | None = None) -> List[str]:
    rules = rules or load_rules()
    tags = tag_item(_taggable_text(item, rules), rules)
    tag_set = set(tags)
    primary = _primary_norm(item, rules)
    norm_blob = normalize_text(_item_blob(item), rules)
    if _is_admin_note(primary):
        return [_MISC_NOTE_TAG]
    if "WARRANTY" in tags and not re.search(r"\bWARRANTY\b", primary, re.I):
        tags = [t for t in tags if t != "WARRANTY"]
    if _GUARD_TAG in tags and not _is_shaft_bearing_guard_line(primary):
        tags = [t for t in tags if t != _GUARD_TAG]
    if _is_base_fan_line(primary):
        tags = [t for t in tags if t not in _BASE_FAN_DETAIL_TAGS]
    if "SPLIT HOUSING" in tags:
        tags = [t for t in tags if t != "HOUSING"]
    if "EXPLOSION PROOF" in tags:
        tags = [t for t in tags if t != "EXPLOSION PROOF"]
        if _is_explosion_proof_motor_context(primary):
            _add_unique(tags, "MOTOR")
    if "MIXING BOX" in tags and _is_mixing_box_line(primary, norm_blob):
        tags = [t for t in tags if t != "INLET"]
    if _is_non_inlet_component_mounted_to_inlet_box(primary, norm_blob, set(tags)):
        tags = [t for t in tags if t != "INLET"]
    if _is_label_instruction_line(primary, norm_blob, set(tags)):
        tags = [t for t in tags if t not in _LABEL_DETAIL_TAGS]
    if "NAMEPLATE" in tags and _is_motor_nameplate_context(primary, norm_blob):
        _add_unique(tags, "MOTOR")
    if _is_ship_via_note(primary, norm_blob):
        tags = [t for t in tags if t not in _SHIP_VIA_COMPONENT_TAGS]
        _add_unique(tags, "SHIPPING")
    if "LIFTING LUGS" in tags and not _is_primary_lifting_lugs(primary):
        tags = [t for t in tags if t != "LIFTING LUGS"]
    if _is_flex_connector_flow_liner(norm_blob, set(tags)):
        tags = [t for t in tags if t != "LINING"]
    if _is_motor_flange_line(primary, norm_blob):
        tags = [t for t in tags if t != "FLANGE"]
        _add_unique(tags, "MOTOR")
    if _is_non_wheel_end_location(norm_blob):
        tags = [t for t in tags if t != "WHEEL"]
        _add_unique(tags, "HOUSING")
    if _is_motor_conduit_box_context(primary, norm_blob):
        _add_unique(tags, "MOTOR")
        if _is_pure_motor_conduit_box_location(primary, norm_blob) and not _has_housing_engineering_feature(norm_blob):
            tags = [t for t in tags if t != "SPECIAL CONSTRUCTION"]
        if not _has_housing_engineering_feature(norm_blob):
            tags = [t for t in tags if t != "HOUSING"]
    if _is_without_ivc(norm_blob):
        tags = [t for t in tags if t != "INLET VANES"]
    if _is_ivc_actuator_context(norm_blob, set(tags)):
        _add_unique(tags, "DAMPER")
    if _is_inlet_cone_width_without_wheel(norm_blob):
        tags = [t for t in tags if t != "WHEEL"]
    if "FLEX CONNECTOR" in tags and "FLANGE" in tags and _is_flex_connector_line(norm_blob):
        tags = [t for t in tags if t != "FLANGE"]
    temp_blob = f"{norm_blob} {_item_blob(item).upper()}"
    has_extreme_temp = bool(_temperature_direction(temp_blob, _extreme_temperature_values(temp_blob)))
    if (_EXTREME_TEMPERATURE_TAG in tags or has_extreme_temp) and _is_top_level_extreme_temperature(
        primary, norm_blob, set(tags)
    ):
        _add_unique(tags, _EXTREME_TEMPERATURE_TAG)
    elif _EXTREME_TEMPERATURE_TAG in tags:
        tags = [t for t in tags if t != _EXTREME_TEMPERATURE_TAG]
    drain_type = _drain_type(primary, norm_blob)
    if drain_type == "HOUSING DRAIN":
        tags = [t for t in tags if t != "HOUSING"]
    elif drain_type == "INLET BOX DRAIN":
        tags = [t for t in tags if t != "INLET"]
    if "OUTLET" in tags and "DAMPER" in tags and _used_on(norm_blob) == "OUTLET DAMPER":
        tags = [t for t in tags if t != "OUTLET"]
    if _is_non_fan_shaft_seal_context(primary, norm_blob, set(tags)):
        tags = [t for t in tags if t != "SHAFT SEAL"]
    if "SHIPPING" in tags and re.search(r"\bALL\s+AUXILIARY\s+ITEMS\s+EXCEPT\s+MOTOR\b", norm_blob):
        tags = [t for t in tags if t != "MOTOR"]
    if _is_incidental_shipping_reference(primary, norm_blob, set(tags)):
        tags = [t for t in tags if t != "SHIPPING"]
    if "SPARE PARTS" in tags and not _is_spare_parts_primary(primary):
        tags = [t for t in tags if t != "SPARE PARTS"]
    if _is_belt_guard_line(primary):
        tags = [t for t in tags if t not in _BELT_GUARD_DETAIL_TAGS]
    elif _is_accessory_coating(primary, tag_set, norm_blob):
        tags = [t for t in tags if t != "COATING"]
    if _is_drive_table_line(primary):
        tags = [t for t in tags if t not in _DRIVE_TABLE_DETAIL_TAGS]
    if "COATING" in tags and _is_paint_line(primary):
        tags = [t for t in tags if t not in _PAINT_SURFACE_TAGS]
    if "INSPECTION" in tags and not _is_inspection_line(primary, norm_blob):
        tags = [t for t in tags if t != "INSPECTION"]
    if "TESTING" in tags and not _is_testing_context(primary, norm_blob, set(tags)):
        tags = [t for t in tags if t != "TESTING"]
    if _is_packaging_inspection_primary(primary, set(tags)):
        tags = [t for t in tags if t not in _PACKAGING_INSPECTION_DETAIL_TAGS]
    if _is_motor_insulation_only(primary, norm_blob, set(tags)):
        tags = [t for t in tags if t != "INSULATION"]
    if "LINING" in tags:
        if re.search(r"\b(SCROLL|SIDE\s+SHEET)\b", norm_blob):
            _add_unique(tags, "HOUSING")
        if re.search(r"\b(WHEEL\s+BLADES?|BLADES?)\b", norm_blob):
            _add_unique(tags, "WHEEL")
    if _is_housing_packaging_reference(norm_blob, set(tags)):
        tags = [t for t in tags if t != "HOUSING"]
    if _is_assembly_note(primary):
        tags = [t for t in tags if t not in _MISC_NOTE_COMPONENT_TAGS]
        _add_unique(tags, _MISC_NOTE_TAG)
    for tag in _lube_component_tags(primary, norm_blob):
        _add_unique(tags, tag)
    if _component_material_owner(item, _material_attributes(norm_blob), rules):
        tags = [t for t in tags if t not in _MATERIAL_TAGS]
    if "MOUNTING" in tags:
        tags = [t for t in tags if t != "MOUNTING"]
    return sorted(tags)


def component_attributes(item: Dict[str, Any], rules: Dict[str, Any] | None = None) -> Dict[str, str]:
    """Structured fan-component details pulled from raw text + detail lines."""
    rules = rules or load_rules()
    blob = _item_blob(item)
    norm_blob = normalize_text(blob, rules)
    raw_tags = set(tag_item(_taggable_text(item, rules), rules))
    tags = set(_final_tags(item, rules))
    primary = _primary_norm(item, rules)
    attrs: Dict[str, str] = {}

    inquiries = inquiry_numbers(item)
    if inquiries:
        attrs["inquiry_num"] = ", ".join(inquiries)
    if _is_packaging_inspection_primary(primary, tags):
        attrs.update(_inspection_attributes(primary, norm_blob, tags))
        attrs.update(_shipping_attributes(norm_blob, tags))
        return attrs
    attrs.update(_drawing_attributes(primary, norm_blob, tags))
    attrs.update(_split_housing_attributes(norm_blob, tags))
    attrs.update(_explosion_proof_attributes(primary, norm_blob, raw_tags))
    attrs.update(_special_construction_attributes(primary, norm_blob, tags))
    attrs.update(_flange_attributes(primary, norm_blob, tags, raw_tags))
    attrs.update(_flex_connector_attributes(norm_blob, tags))
    attrs.update(_coupling_attributes(norm_blob, tags))
    attrs.update(_low_leakage_attributes(norm_blob, tags))
    attrs.update(_temperature_attributes(primary, norm_blob, tags, blob))
    attrs.update(_heavy_duty_attributes(norm_blob, tags))
    attrs.update(_inlet_attributes(primary, norm_blob, tags, blob))
    attrs.update(_mixing_box_attributes(norm_blob, tags))
    attrs.update(_inlet_mount_attributes(norm_blob))
    attrs.update(_inlet_vane_attributes(primary, norm_blob, tags))
    attrs.update(_inspection_attributes(primary, norm_blob, tags))
    attrs.update(_motor_insulation_attributes(norm_blob, tags, raw_tags))
    attrs.update(_insulation_attributes(item, primary, norm_blob, tags))
    attrs.update(_label_attributes(norm_blob, tags))
    attrs.update(_mounting_attributes(primary, norm_blob, tags))
    attrs.update(_nameplate_attributes(primary, norm_blob, tags))
    attrs.update(_lifting_lug_attributes(tags))
    attrs.update(_screen_attributes(primary, norm_blob, tags))
    attrs.update(_shaft_cooler_attributes(primary, norm_blob, tags))
    attrs.update(_shaft_seal_attributes(norm_blob, tags))
    attrs.update(_shaft_sleeve_attributes(norm_blob, tags))
    attrs.update(_lining_attributes(norm_blob, tags))
    attrs.update(_housing_attributes(norm_blob, tags))
    attrs.update(_motor_conduit_box_attributes(primary, norm_blob))

    vendor = _label_value(item, "Vendor")
    product = _label_value(item, "Product")
    base_fan = _is_base_fan_line(primary)
    if vendor and not base_fan:
        attrs["vendor"] = vendor
    if product and not base_fan:
        attrs["product"] = product
    attrs.update(_silencer_attributes(primary, norm_blob, tags, product, blob))
    attrs.update(_spark_resistant_attributes(norm_blob, tags))
    attrs.update(_spare_parts_attributes(primary, norm_blob, tags))
    admin_note = _is_admin_note(primary)
    if admin_note:
        attrs["note_type"] = "ADMIN"
        return attrs
    elif _is_assembly_note(primary):
        attrs["note_type"] = "ASSEMBLY"
    attrs.update(_shipping_attributes(norm_blob, tags))
    attrs.update(_testing_attributes(primary, norm_blob, tags))
    attrs.update(_guard_attributes(primary, norm_blob))
    attrs.update(_coating_attributes(primary, norm_blob, tags, raw_tags))
    attrs.update(_drain_attributes(primary, norm_blob, tags))
    attrs.update(_lube_review_attributes(primary, norm_blob))

    material_attrs = _material_attributes(norm_blob)
    material_owner = _component_material_owner(item, material_attrs, rules)
    if material_owner:
        attrs.setdefault("component", material_owner)
        for key, val in _component_material_attributes(material_owner, material_attrs).items():
            attrs[f"component_{key}"] = val
    else:
        attrs.update(material_attrs)

    attrs.update(_balance_attributes(norm_blob))
    attrs.update(_bearing_attributes(blob, norm_blob))

    used_on = "" if _is_ship_via_note(primary, norm_blob) else _used_on(norm_blob)
    if used_on:
        attrs["used_on"] = used_on

    if "ACTUATOR" in tags or "ACTUATOR" in norm_blob:
        attrs["component"] = "ACTUATOR"
        for label, key in (
            ("Actuator Manufacturer", "manufacturer"),
            ("Actuator Supplied By", "supplied_by"),
            ("Actuator Mounting", "mounting"),
            ("Fail Position Upon Loss of Supply/Air/Power", "fail_power"),
            ("Fail Position Upon Loss of Signal", "fail_signal"),
            ("Operation", "operation"),
        ):
            val = _label_value(item, label)
            if val:
                attrs[key] = val
        m = re.search(
            r"\b(UNIC\s*(?P<unic_size>\d+(?:/\d+)?)|"
            r"BETTIS\s*#?\s*(?P<bettis_size>[A-Z0-9-]+)|"
            r"EMERSON\s+FIELD\s+Q)\b",
            blob, re.I,
        )
        if m:
            attrs["model"] = re.sub(r"\s+", " ", m.group(1)).strip()
            size = (_label_value(item, "Actuator Size")
                    or m.groupdict().get("unic_size")
                    or m.groupdict().get("bettis_size")
                    or "")
            if size:
                attrs["size"] = size.upper()
        if _needs_used_on_review(norm_blob, attrs):
            attrs["used_on_review"] = "INCONCLUSIVE - INLET/OUTLET/PRESPIN/IVC"

    is_drive = bool({"V-BELT DRIVE", "DRIVE COMPONENTS"} & tags)
    if is_drive:
        attrs.setdefault("component", "V-BELT DRIVE")
        m = re.search(r"Max/Min RPM:\s*(\d+)\s*/\s*(\d+)", blob, re.I)
        if m:
            attrs["max_rpm"], attrs["min_rpm"] = m.group(1), m.group(2)
        m = re.search(r"\b(\d+)\s+belts?\s*:\s*([A-Z0-9-]+)", blob, re.I)
        if m:
            attrs["belt_qty"], attrs["belt"] = m.group(1), m.group(2).upper()
        m = re.search(r"Motor\s+Sheave/Bushing:\s*(.*?)\s*,?\s*Fan\s+Sheave/Bushing:\s*(.*?)\s*,?\s*Actual\s+SF",
                      blob, re.I)
        if m:
            _set_sheave_attrs_from_text(attrs, "drive", m.group(1))
            _set_sheave_attrs_from_text(attrs, "driven", m.group(2))
        for label, key in (("Actual SF", "actual_sf"), ("Actual CD", "actual_cd")):
            m = re.search(rf"\b{re.escape(label)}\s*:\s*([0-9.]+)", blob, re.I)
            if m:
                attrs[key] = m.group(1)
        m = re.search(r"Constant Speed,\s*SF:\s*([0-9.]+)", blob, re.I)
        if m:
            attrs["service_factor"] = m.group(1)
        m = re.search(r"Specified minimum belt service factor:\s*([0-9.]+)", blob, re.I)
        if m:
            attrs["service_factor"] = m.group(1)
        m = re.search(r"Center Distance with allowance for install and take-up:\s*([0-9.]+\s*-\s*[0-9.]+)",
                      blob, re.I)
        if m:
            attrs["center_distance_range"] = m.group(1)

        for key, value in _drive_table_attributes(item).items():
            attrs.setdefault(key, value)

    return attrs


def _review_flags(tags: List[str], attrs: Dict[str, Any]) -> List[str]:
    flags: List[str] = []
    if not tags:
        flags.append("UNTAGGED")
    if isinstance(attrs, dict):
        for key, value in sorted(attrs.items()):
            if not value or not str(key).endswith("_review"):
                continue
            label = str(key).replace("_", " ").upper()
            flags.append(f"{label}: {value}")
    return flags


def _apply_review_flags(item: Dict[str, Any]) -> None:
    flags = _review_flags(list(item.get("tags") or []), item.get("attributes") or {})
    if flags:
        item["review_flags"] = flags
    else:
        item.pop("review_flags", None)


def derive_item_fields(item: Dict[str, Any], rules: Dict[str, Any] | None = None) -> Dict[str, Any]:
    """Re-derive normalized fields for a stored item without mutating it."""
    rules = rules or load_rules()
    body, qty = split_lead(item.get("raw", ""))
    body, price = split_price_tail(body)
    body, ptype, mark = split_type_tail(body)
    norm = normalize_text(body, rules)
    probe = dict(item)
    probe["norm"] = norm
    tags = _final_tags(probe, rules)
    attrs = component_attributes(probe, rules)
    return {"norm": norm, "qty": qty, "price": price or mark, "ptype": ptype,
            "tags": tags, "attributes": attrs, "review_flags": _review_flags(tags, attrs)}


# --------------------------------------------------------------------------- #
# Extraction from a Sales Order's text lines                                  #
# --------------------------------------------------------------------------- #
def classify_line(line: str, section: str, rules: Dict[str, Any]) -> Tuple[str, str]:
    """How one reconstructed text line is treated (stateless part — the
    item-vs-detail decision needs the running state in iter_classified).
    Returns (kind, detail):

        blank          empty line
        section-start  opens a feature section (detail = its title)
        section-end    closes it (detail = the matched pattern)
        skip           never an item (detail = the skip pattern that hit)
        item-priced    an item line — it ends in price columns, an L/C/N
                       price-type letter, or both (detail = what anchored it)
        item-section   an unpriced line captured because a section is open
        text           everything else
    """
    s = line.strip()
    if not s:
        return "blank", ""
    if len(s) <= 48:
        for p in rules["start"]:
            if p.search(s):
                return "section-start", s
    if section:  # end markers only mean anything while a section is open
        for p in rules["end"]:
            if p.search(s):
                return "section-end", p.pattern
    for p in rules["skip"]:
        if p.search(s):
            return "skip", p.pattern
    body, price = split_price_tail(s)
    body, ptype, mark = split_type_tail(body)
    if price or ptype:
        if not re.search(r"[A-Za-z]", body):
            return "text", ""  # all-numeric row (the CFM/RPM performance values)
        if ptype and not (price or mark) and not re.search(r"[A-Za-z]{3}", body):
            return "text", ""  # spec-table row that happens to end in L/C/N
        anchor = " ".join(x for x in (ptype, mark or price) if x)
        return "item-priced", anchor
    if section:
        return "item-section", section
    return "text", ""


def _detail_worthy(s: str) -> bool:
    """Is an unpriced line a plausible item detail? Needs real words, and rows
    of bare reference numbers ('421473 7074-49840-00-AI26') don't qualify —
    every token containing a digit means it's an order/PO ref row, not specs
    ('200 HP, 1800 RPM' has the unit words)."""
    if not re.search(r"[A-Za-z]{2}", s):
        return False
    if all(any(ch.isdigit() for ch in t) for t in s.split()):
        return False
    return True


MAX_DETAILS = 10  # per item — continuation blocks run ~7 lines (motors)


def iter_classified(lines: Iterable[str], rules: Dict[str, Any] | None = None,
                    ) -> Iterator[Tuple[str, str, str]]:
    """The full extraction state machine, yielding (kind, detail, line) for
    every line — classify_line's kinds plus "detail" (an unpriced line
    attached to the item above: vendor, motor specs, "Product: Damper", ...).
    Skips do NOT break a detail block (page furniture interleaves one), but a
    new item, a section event, or MAX_DETAILS does. Single source of truth for
    extract_items AND the --dump tuning view, so the dump shows exactly what
    the extractor does."""
    rules = rules or load_rules()
    section = ""
    have_item = False
    n_details = 0
    for line in lines:
        s = re.sub(r"\s+", " ", str(line)).strip()
        kind, detail = classify_line(s, section, rules)
        if kind == "section-start":
            section, have_item = detail, False
        elif kind == "section-end":
            section, have_item = "", False
        elif kind in ("item-priced", "item-section"):
            have_item, n_details = True, 0
        elif kind == "text" and have_item and n_details < MAX_DETAILS and _detail_worthy(s):
            kind = "detail"
            n_details += 1
        yield kind, detail, s


def extract_items(lines: Iterable[str], rules: Dict[str, Any] | None = None) -> List[Dict[str, Any]]:
    """Pull the line items out of a Sales Order's reconstructed text lines.

    Two capture signals, covering both ways SOs lay items out: any line ending
    in the price columns / L-C-N type letter is an item wherever it sits, and
    every line inside a recognized feature section ("Additional Features",
    "Accessories", ...) is an item even unpriced. Unpriced lines under an item
    become its `details`. Duplicate lines (same normalized form) collapse to
    one. Items come back as {raw, norm, qty, price, ptype, section, details,
    tags} — tags consider the details too."""
    rules = rules or load_rules()
    by_norm: Dict[str, Dict[str, Any]] = {}
    last: Dict[str, Any] | None = None
    section = ""
    for kind, detail, s in iter_classified(lines, rules):
        if kind == "section-start":
            section, last = detail, None
            continue
        if kind == "section-end":
            section, last = "", None
            continue
        if kind == "detail":
            if last is not None:
                last["details"].append(s)
            continue
        if kind not in ("item-priced", "item-section"):
            continue
        body, qty = split_lead(s)
        body, price = split_price_tail(body)
        body, ptype, mark = split_type_tail(body)
        norm = normalize_text(body, rules)
        if len(norm) < 3:  # stray cell fragments ("X", "1") aren't items
            last = None
            continue
        prev = by_norm.get(norm)
        if prev is not None:
            if (price or mark) and not prev["price"]:
                prev["price"] = price or mark
            last = prev
            continue
        last = by_norm[norm] = {
            "raw": s,
            "norm": norm,
            "qty": qty,
            "price": price or mark,
            "ptype": ptype,
            "section": section,
            "details": [],
            "tags": [],
            "attributes": {},
        }
    # Tags consider the details too — "Product: Damper" under an "IVD" row is
    # what identifies it.
    items = list(by_norm.values())
    for it in items:
        it["tags"] = _final_tags(it, rules)
        it["attributes"] = component_attributes(it, rules)
        _apply_review_flags(it)
    return items


def _taggable_text(item: Dict[str, Any], rules: Dict[str, Any] | None = None) -> str:
    """norm + normalized details — everything a tag pattern may match."""
    parts = [item.get("norm", "")]
    parts += [normalize_text(d, rules) for d in item.get("details") or []]
    return " ; ".join(p for p in parts if p)


def tags_label(items: List[Dict[str, Any]] | None) -> str:
    """Compact cell text for the report's Features column: the union of the
    job's tags, or '(N items)' when items were captured but none tagged yet."""
    items = items or []
    tags = sorted({t for it in items for t in it.get("tags") or []})
    if tags:
        return ", ".join(tags)
    return f"({len(items)} items)" if items else ""


# --------------------------------------------------------------------------- #
# Store                                                                       #
# --------------------------------------------------------------------------- #
def store_path() -> Path:
    return LINE_ITEMS_STORE if LINE_ITEMS_STORE else BACKLOG_DIR / "line_items.json"


def load_store(path: Path | None = None) -> Dict[str, Any]:
    p = path or store_path()
    if p.exists():
        try:
            store = json.loads(p.read_text(encoding="utf-8"))
            store.setdefault("jobs", {})
            store.setdefault("ai_tags", {})
            return store
        except (json.JSONDecodeError, OSError) as e:
            log.warning("Could not read line-items store %s (%s); starting fresh", p, e)
    return {"jobs": {}, "ai_tags": {}}


def save_store(store: Dict[str, Any], path: Path | None = None) -> None:
    p = path or store_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(store, indent=2), encoding="utf-8")
    tmp.replace(p)  # atomic — a crash mid-write never corrupts the store


def apply_ai_cache(items: List[Dict[str, Any]], store: Dict[str, Any]) -> None:
    """Fill still-untagged items from cached Claude classifications in place."""
    ai = store.get("ai_tags") or {}
    for it in items:
        extra = ai.get(it.get("norm", ""))
        if extra and not it.get("tags"):
            it["tags"] = sorted(set(extra))
            _apply_review_flags(it)


def _is_skipped_stored_item(item: Dict[str, Any], rules: Dict[str, Any]) -> bool:
    raw = str(item.get("raw", "")).strip()
    return bool(raw and any(p.search(raw) for p in rules["skip"]))


def audit_untagged(store: Dict[str, Any], limit: int = 50) -> List[Dict[str, Any]]:
    """Most common normalized items that current rules still keep but do not tag."""
    rules = load_rules(refresh=True)
    ai = store.get("ai_tags") or {}
    rows: Dict[str, Dict[str, Any]] = {}
    for job, rec in (store.get("jobs") or {}).items():
        for item in rec.get("items") or []:
            if _is_skipped_stored_item(item, rules):
                continue
            derived = derive_item_fields(item, rules)
            norm = derived["norm"]
            if not norm or derived["tags"]:
                continue
            row = rows.setdefault(norm, {"norm": norm, "count": 0, "jobs": [], "ai_tags": []})
            row["count"] += 1
            if len(row["jobs"]) < 5:
                row["jobs"].append(str(job))
            for tag in ai.get(norm) or []:
                if tag not in row["ai_tags"]:
                    row["ai_tags"].append(tag)
    out = sorted(rows.values(), key=lambda r: (-r["count"], r["norm"]))
    return out[:limit] if limit else out


def audit_review(store: Dict[str, Any], limit: int = 50, tag: str = "") -> List[Dict[str, Any]]:
    """Most common line-item templates that need a human/rule decision."""
    rules = load_rules(refresh=True)
    want_tag = tag.upper().strip()
    rows: Dict[Tuple[str, Tuple[str, ...], Tuple[str, ...]], Dict[str, Any]] = {}
    for job, rec in (store.get("jobs") or {}).items():
        for item in rec.get("items") or []:
            if _is_skipped_stored_item(item, rules):
                continue
            derived = derive_item_fields(item, rules)
            flags = derived.get("review_flags") or []
            tags = derived.get("tags") or []
            if not flags:
                continue
            if want_tag and want_tag not in [t.upper() for t in tags]:
                continue
            norm = derived["norm"]
            key = (norm, tuple(tags), tuple(flags))
            row = rows.setdefault(key, {
                "norm": norm,
                "count": 0,
                "jobs": [],
                "tags": tags,
                "review_flags": flags,
                "sample": str(item.get("raw", "")),
            })
            row["count"] += 1
            if len(row["jobs"]) < 8:
                row["jobs"].append(str(job))
    out = sorted(rows.values(), key=lambda r: (-r["count"], r["norm"]))
    return out[:limit] if limit else out


def record_job(store: Dict[str, Any], job: str, items: List[Dict[str, Any]],
               customer: str = "", co_number: int | None = None, so_pdf: str = "") -> None:
    """Record (or refresh) one job's line items. The latest parse wins, but
    metadata never regresses: a blank customer/co/pdf from a sparse source
    (e.g. the archive scan, which has no board context) keeps the old value."""
    prev = store["jobs"].get(job) or {}
    store["jobs"][job] = {
        "customer": customer or prev.get("customer", ""),
        "co_number": co_number if co_number is not None else prev.get("co_number"),
        "so_pdf": so_pdf or prev.get("so_pdf", ""),
        "scanned_at": datetime.now().isoformat(timespec="seconds"),
        "items": items,
    }


# --------------------------------------------------------------------------- #
# Search                                                                      #
# --------------------------------------------------------------------------- #
def _job_sort_key(job: str):
    return (0, -int(job)) if job.isdigit() else (1, job)


def _term_matches(term: str, item: Dict[str, Any], fuzzy: float = 0.0) -> bool:
    """One search term vs one item: substring of the normalized/raw text, the
    detail lines, or any tag; with fuzzy > 0, also a close difflib match
    against any same-length word window of the normalized text."""
    t = term.upper().strip()
    hay = [item.get("norm", ""), item.get("raw", "").upper()]
    hay += [x.upper() for x in item.get("tags") or []]
    attrs = item.get("attributes") or {}
    if isinstance(attrs, dict):
        for k, v in attrs.items():
            val = " ".join(str(x) for x in v) if isinstance(v, list) else str(v)
            hay.extend([str(k).upper(), val.upper(), normalize_text(val)])
    for d in item.get("details") or []:
        hay.append(d.upper())
        hay.append(normalize_text(d))
    if any(t in h for h in hay):
        return True
    if fuzzy <= 0:
        return False
    from difflib import SequenceMatcher
    words = item.get("norm", "").split()
    n = max(1, len(t.split()))
    for i in range(0, max(1, len(words) - n + 1)):
        window = " ".join(words[i:i + n])
        if window and SequenceMatcher(None, t, window).ratio() >= fuzzy:
            return True
    return False


def search(store: Dict[str, Any], terms: List[str], any_mode: bool = False,
           tag: str = "", fuzzy: float = 0.0) -> List[Dict[str, Any]]:
    """Find jobs by their line items. Criteria AND at the JOB level: every term
    must match some item of the job (any single term with any_mode), and with
    `tag` the job must also carry an item with that canonical tag — the term
    and the tag may sit on different items. Returns newest-job-first records
    {job, customer, co_number, so_pdf, scanned_at, matches: [items]} where
    matches are the items that satisfied any criterion."""
    out = []
    for job in sorted(store.get("jobs") or {}, key=_job_sort_key):
        rec = store["jobs"][job]
        items = rec.get("items") or []
        hit_lists: List[List[Dict[str, Any]]] = []
        if tag:
            tagged = [it for it in items
                      if tag.upper() in [x.upper() for x in it.get("tags") or []]]
            if not tagged:
                continue
            hit_lists.append(tagged)
        if terms:
            per_term = [[it for it in items if _term_matches(t, it, fuzzy)] for t in terms]
            if not (any(per_term) if any_mode else all(per_term)):
                continue
            hit_lists.extend(per_term)
        if not hit_lists:
            if not items:
                continue
            hit_lists.append(items)  # no criteria: the full inventory
        matched, seen = [], set()
        for lst in hit_lists:
            for it in lst:
                if id(it) not in seen:
                    seen.add(id(it))
                    matched.append(it)
        out.append({"job": job, "customer": rec.get("customer", ""),
                    "co_number": rec.get("co_number"), "so_pdf": rec.get("so_pdf", ""),
                    "scanned_at": rec.get("scanned_at", ""), "matches": matched})
    return out


def tag_counts(store: Dict[str, Any]) -> List[Tuple[str, int, int]]:
    """The live vocabulary: (tag, #jobs, #items), most-used first."""
    jobs: Dict[str, set] = {}
    items: Dict[str, int] = {}
    for job, rec in (store.get("jobs") or {}).items():
        for it in rec.get("items") or []:
            for t in it.get("tags") or []:
                jobs.setdefault(t, set()).add(job)
                items[t] = items.get(t, 0) + 1
    return sorted(((t, len(jobs[t]), items[t]) for t in jobs),
                  key=lambda r: (-r[1], -r[2], r[0]))


def inquiry_counts(store: Dict[str, Any]) -> List[Tuple[str, int, int, List[str]]]:
    """Inquiry number rollup: (inquiry_num, #jobs, #items, job list)."""
    rules = load_rules()
    jobs: Dict[str, set] = {}
    items: Dict[str, int] = {}
    for job, rec in (store.get("jobs") or {}).items():
        for it in rec.get("items") or []:
            attrs = it.get("attributes") if isinstance(it.get("attributes"), dict) else {}
            inquiry = (attrs or component_attributes(it, rules)).get("inquiry_num", "")
            for num in [n.strip().upper() for n in str(inquiry).split(",") if n.strip()]:
                jobs.setdefault(num, set()).add(str(job))
                items[num] = items.get(num, 0) + 1
    return sorted(((n, len(jobs[n]), items[n], sorted(jobs[n], key=_job_sort_key))
                   for n in jobs),
                  key=lambda r: (-r[1], -r[2], r[0]))


# --------------------------------------------------------------------------- #
# AI normalization of the long tail (optional, cached forever)                #
# --------------------------------------------------------------------------- #
_AI_SYSTEM = """You normalize line items from industrial fan sales orders into canonical feature tags.

You receive a JSON list of normalized line-item strings. For EACH string, return the canonical feature tags that apply (several may apply; an empty list means the line is not a fan feature/accessory — e.g. a heading, address fragment, or pricing row).

Prefer tags from the existing vocabulary below. Only invent a new tag when nothing in the vocabulary fits, and make it a SHORT UPPERCASE noun phrase (e.g. "SHAFT GROUNDING RING").

EXISTING VOCABULARY:
__VOCAB__

Output STRICT JSON only — an object mapping every input string (exactly as given) to its list of tags:
{"<item>": ["TAG", ...], ...}"""


def unknown_norms(store: Dict[str, Any]) -> List[str]:
    """Unique normalized strings that no rule tagged and the AI hasn't seen."""
    ai = store.get("ai_tags") or {}
    seen = set()
    for rec in (store.get("jobs") or {}).values():
        for it in rec.get("items") or []:
            n = it.get("norm", "")
            if n and not it.get("tags") and n not in ai:
                seen.add(n)
    return sorted(seen)


def ai_classify_unknowns(store: Dict[str, Any], batch_size: int = 60) -> int:
    """Send the un-tagged unique items to Claude (CLAUDE_MODEL) and cache the
    classifications in store["ai_tags"] — each unique string is ever classified
    once, so re-runs are free. Returns how many strings were classified. Apply
    the cache to stored items afterwards (see line_items_scan.py --ai)."""
    from config import ANTHROPIC_API_KEY, CLAUDE_MODEL
    if not ANTHROPIC_API_KEY:
        raise RuntimeError("ANTHROPIC_API_KEY is not set (check your .env) — "
                           "the --ai pass needs it.")
    todo = unknown_norms(store)
    if not todo:
        return 0
    import anthropic
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    rules = load_rules()
    done = 0
    for i in range(0, len(todo), batch_size):
        batch = todo[i:i + batch_size]
        vocab = sorted(set(rules["tags"]) |
                       {t for tags in (store.get("ai_tags") or {}).values() for t in tags})
        response = client.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=min(8000, 300 + 45 * len(batch)),
            thinking={"type": "disabled"},
            system=_AI_SYSTEM.replace("__VOCAB__", ", ".join(vocab)),
            messages=[{"role": "user", "content": json.dumps(batch, indent=0)}],
        )
        text = next((b.text for b in response.content if b.type == "text"), "").strip()
        if text.startswith("```"):  # strip ```json fences if Claude added them
            text = text.strip("`")
            if text.startswith("json"):
                text = text[4:]
            text = text.strip()
            if text.endswith("```"):
                text = text[:-3].strip()
        try:
            result = json.loads(text)
        except json.JSONDecodeError as e:
            log.error("Claude returned invalid JSON for a line-item batch (%s); raw:\n%s", e, text)
            continue  # this batch stays unknown; a later run retries it
        for norm in batch:
            tags = result.get(norm)
            if isinstance(tags, list):
                store.setdefault("ai_tags", {})[norm] = sorted(
                    {str(t).upper().strip() for t in tags if str(t).strip()})
                done += 1
        log.info("  AI-classified %d/%d unique items...", min(i + batch_size, len(todo)), len(todo))
    return done


def renormalize_store(store: Dict[str, Any]) -> int:
    """Re-derive every stored item's norm + tags from its verbatim raw text
    (and details) using the CURRENT rules + the AI cache. Run after editing
    the rules file — nothing is re-downloaded or re-parsed. Returns the item
    count."""
    rules = load_rules(refresh=True)
    n = 0
    for rec in (store.get("jobs") or {}).values():
        kept: List[Dict[str, Any]] = []
        for it in rec.get("items") or []:
            if _is_skipped_stored_item(it, rules):
                continue
            derived = derive_item_fields(it, rules)
            it["norm"] = derived["norm"]
            it["tags"] = derived["tags"]
            it["attributes"] = derived["attributes"]
            if derived.get("review_flags"):
                it["review_flags"] = derived["review_flags"]
            else:
                it.pop("review_flags", None)
            qty = derived["qty"]
            price = derived["price"]
            ptype = derived["ptype"]
            if qty and not it.get("qty"):
                it["qty"] = qty
            if price and not it.get("price"):
                it["price"] = price
            if ptype and not it.get("ptype"):
                it["ptype"] = ptype
            kept.append(it)
            n += 1
        rec["items"] = kept
        apply_ai_cache(kept, store)
    return n
