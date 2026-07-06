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
        r"^prints?\s*$", r"^product\s*:?\s*$", r"^product\s+\$?\d",
        # Shipping/admin notes can appear inside Additional Features / Notes;
        # keep them out of the item inventory while preserving real ship-loose
        # priced rows.
        r"^additional\s+shipping\s+notes?\b", r"^traffic\s+note\b",
        r"shipping\s+barcode", r"do\s+not\s+stack", r"no\s+metal\s+banding",
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
        "SPARK RESISTANT": [r"spark"],
        "SHAFT SEAL": [r"shaft\s*seal", r"stuffing\s*box", r"lip\s*seal",
                       r"ceramic\s*felt"],
        "SHAFT SLEEVE": [r"shaft\s*sleeve"],
        "SHAFT COOLER": [r"shaft\s*cooler", r"heat\s*slinger"],
        "MATERIALS": [r"stainless", r"\b304L?\b", r"\b316L?\b", r"passivat",
                      r"aluminum", r"aluminium"],
        "STAINLESS STEEL": [r"stainless", r"\b304L?\b", r"\b316L?\b", r"passivat"],
        "ALUMINUM": [r"aluminum", r"aluminium"],
        "HIGH TEMPERATURE": [r"high\s*temp", r"heat\s*fan"],
        "HEAVY DUTY": [r"heavy\s*duty"],
        "COATING": [r"epoxy", r"\bcoat(?:ed|ing|s)?\b", r"paint", r"primer",
                    r"plasite", r"heresite", r"\bzinc\b"],
        "LINING": [r"rubber\s*lin", r"\blined\b", r"\blining\b", r"abrasion",
                   r"firmex"],
        "INSULATION": [r"insulat"],
        "VIBRATION ISOLATION": [r"isolat", r"rubber[\s-]*in[\s-]*shear",
                                r"\bRIS\b", r"spring\s*mount", r"seismic",
                                r"vibration\s*base"],
        "VIBRATION SWITCH": [r"vibration\s*(switch|detector|monitor|sensor)"],
        "DAMPER": [r"damper", r"backdraft", r"volume\s*control"],
        "INLET VANES": [r"inlet\s*vane", r"\bVIV\b", r"\bIVC\b",
                        r"variable\s*inlet", r"inlet\s*volume\s*control"],
        "INLET": [r"\binlet\s+(open|slip|bell|cone|box|tube|flanged|punched)",
                  r"\binlet\s+direction"],
        "OUTLET": [r"\boutlet\s+(open|slip|flanged|punched|pressure\s*tap|volume\s*control)",
                   r"\bdischarge\s+elbow"],
        "WHEEL": [r"\bwheel\b", r"\bpercent\s+width\b", r"%\s*width\b"],
        "HOUSING": [r"\bhousing\b"],
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
        "FLEXIBLE COUPLING": [r"flexible\s*coupling", r"steelflex",
                              r"thomas\s+series"],
        "UNITARY BASE": [r"unitary\s*base", r"structural\s*(steel\s*)?base",
                         r"channel\s*base"],
        "BEARINGS": [r"^bearings?\s+(standard|split\s+pillow\s+block)\b",
                     r"^repair\s+bearings?\b", r"^spare\s+bearings?\b",
                     r"^bearing\s+adder\b"],
        "LIFTING LUGS": [r"lifting\s*lugs?"],
        "LABEL": [r"\blabel\b"],
        "NAMEPLATE": [r"nameplate"],
        "PACKAGING": [r"\bcrate\b", r"\bcrating\b", r"shrink\s*wrap",
                      r"\bskid\b", r"ispm\s*wood", r"do\s*not\s*stack",
                      r"metal\s*banding"],
        "SHIPPING": [r"ship\s*loose", r"freight\s*included", r"\bshipping\b"],
        "WARRANTY": [r"warranty"],
        "MOUNTING": [r"\bmounting\b", r"\bmounted\b", r"\bcbc\s*mount\b"],
        "SPECIAL CONSTRUCTION": [r"set\s*screws?", r"loc\s*tite", r"caulking",
                                 r"\bweld(?:ing)?\b", r"tie\s*rod\s*support",
                                 r"buffer\s*tube", r"cast\s*hub",
                                 r"plug\s*panel", r"pressure\s*tap",
                                 r"\bconduit\b", r"\boverhang\b",
                                 r"effective\s*diameter", r"threaded\s*plug",
                                 r"hole\s*diameters?", r"earthing\s*boss"],
        "INSPECTION": [r"\binspection\b", r"mill\s*certifications?"],
        "DRAWINGS": [r"\bcertified\s*drawings?\b", r"\bprints?\b",
                     r"\bdrawings?\b"],
        "MOTOR": [r"\bmotor\b"],
        "VFD": [r"\bVFD\b", r"variable\s*freq", r"inverter"],
        "EXPLOSION PROOF": [r"explosion\s*proof", r"class\s*i+\b.*div"],
        "DRIVE COMPONENTS": [r"sheave\s*/?\s*bushing", r"\bbushing\b",
                             r"\bactual\s+sf\b", r"\bactual\s+cd\b",
                             r"selected\s+drive", r"center\s+distance",
                             r"^\d{3,4}\s+\d{3,4}\s+[A-Z]{1,2}\d+\s+\d+\s+"],
        "V-BELT DRIVE": [r"^drive\b", r"v[\s-]*belt", r"sheave",
                         r"bushing", r"drive\s*set"],
        "BALANCE": [r"\bG\s*\d+(?:\.\d+)?\s+balance\b",
                    r"welded\s+balance\s+weights?"],
        "TESTING": [r"witness", r"\btest"],
        "SPARE PARTS": [r"spare"],
        "3D STEP DRAWINGS": [r"3d\s+(step\s+)?drawings?"],
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
    component_context = ("DAMPER" in norm_blob or "ACTUATOR" in norm_blob
                         or "VOLUME CONTROL" in norm_blob)
    if not component_context and "IVC" not in norm_blob:
        return ""
    if "FRESH AIR" in norm_blob:
        return "FRESH AIR DAMPER"
    if "PRESPIN" in norm_blob or "PRE SPIN" in norm_blob:
        return "PRESPIN DAMPER"
    if "OUTLET" in norm_blob and ("DAMPER" in norm_blob or "VOLUME CONTROL" in norm_blob):
        return "OUTLET DAMPER"
    if "DISCHARGE" in norm_blob and ("DAMPER" in norm_blob or "ACTUATOR" in norm_blob):
        return "OUTLET DAMPER"
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


_MATERIAL_GRADE = re.compile(r"\bT?(304L?|316L?)\s+STAINLESS\s+STEEL\b", re.I)
_BALANCE_GRADE = re.compile(r"\bG\s*(\d+(?:\.\d+)?)\s+BALANCE\b", re.I)


def _add_unique(values: List[str], value: str) -> None:
    if value and value not in values:
        values.append(value)


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
    elif "INLET" in norm_blob:
        add_scope("INLET")
    if "OUTLET" in norm_blob or "DISCHARGE" in norm_blob:
        add_scope("OUTLET")
    if passivation and "WELDS" in norm_blob:
        add_scope("WELDS")
    if "ACCESS DOOR" in norm_blob:
        add_scope("ACCESS DOOR")
    if re.search(r"\bHOUSING\s+DRAIN\b", norm_blob):
        add_scope("HOUSING DRAIN")
    if "HOUSING" in norm_blob and not any(s.startswith("HOUSING") for s in scopes):
        add_scope("HOUSING")
    for scope, pattern in (
        ("DRAIN", r"\bDRAIN\b"),
        ("SHAFT COOLER", r"\bSHAFT\s+COOLER\b"),
        ("SHAFT SEAL", r"\bSHAFT\s+SEAL\b"),
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
        if scope == "DRAIN" and "HOUSING DRAIN" in scopes:
            continue
        if re.search(pattern, norm_blob):
            add_scope(scope)
    if "SHAFT" in norm_blob and not {"SHAFT COOLER", "SHAFT SEAL"} & set(scopes):
        add_scope("SHAFT")
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
        attrs["material_scope"] = "MOTOR"
    return attrs


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


def _coating_attributes(primary: str, norm_blob: str, tags: set[str]) -> Dict[str, str]:
    if not _COATING_WORD.search(norm_blob):
        return {}
    attrs: Dict[str, str] = {}
    accessory = _is_accessory_coating(primary, tags, norm_blob) or _is_belt_guard_line(primary)
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
    if _is_belt_guard_line(primary):
        tags = [t for t in tags if t not in _BELT_GUARD_DETAIL_TAGS]
    elif _is_accessory_coating(primary, tag_set, norm_blob):
        tags = [t for t in tags if t != "COATING"]
    if "COATING" in tags and _is_paint_line(primary):
        tags = [t for t in tags if t not in _PAINT_SURFACE_TAGS]
    if _is_assembly_note(primary):
        tags = [t for t in tags if t not in _MISC_NOTE_COMPONENT_TAGS]
        _add_unique(tags, _MISC_NOTE_TAG)
    for tag in _lube_component_tags(primary, norm_blob):
        _add_unique(tags, tag)
    if _component_material_owner(item, _material_attributes(norm_blob), rules):
        tags = [t for t in tags if t not in _MATERIAL_TAGS]
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

    vendor = _label_value(item, "Vendor")
    product = _label_value(item, "Product")
    base_fan = _is_base_fan_line(primary)
    if vendor and not base_fan:
        attrs["vendor"] = vendor
    if product and not base_fan:
        attrs["product"] = product
    admin_note = _is_admin_note(primary)
    if admin_note:
        attrs["note_type"] = "ADMIN"
        return attrs
    elif _is_assembly_note(primary):
        attrs["note_type"] = "ASSEMBLY"
    attrs.update(_guard_attributes(primary, norm_blob))
    attrs.update(_coating_attributes(primary, norm_blob, raw_tags))

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

    used_on = _used_on(norm_blob)
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
            attrs["drive_sheave_bushing"] = m.group(1).strip(" ,;")
            attrs["driven_sheave_bushing"] = m.group(2).strip(" ,;")
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

        tokens = blob.split()
        if (len(tokens) >= 9 and re.fullmatch(r"\d{3,4}/\d{3,4}", tokens[0])
                and re.fullmatch(r"[A-Z]{1,2}\d+", tokens[1], re.I)
                and tokens[2].isdigit()):
            rpms = tokens[0].split("/")
            attrs.setdefault("drive_rpm", rpms[0])
            attrs.setdefault("driven_rpm", rpms[1])
            attrs.setdefault("belt", tokens[1].upper())
            attrs.setdefault("belt_qty", tokens[2])
            attrs.setdefault("drive_sheave_bushing", f"{tokens[3]} {tokens[4]}")
            attrs.setdefault("driven_sheave_bushing", f"{tokens[5]} {tokens[6]}")
            if re.fullmatch(r"\d+(?:\.\d+)?", tokens[7]):
                attrs.setdefault("actual_sf", tokens[7])
            if re.fullmatch(r"\d+(?:\.\d+)?", tokens[8]):
                attrs.setdefault("actual_cd", tokens[8])

    return attrs


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
    return {"norm": norm, "qty": qty, "price": price or mark, "ptype": ptype,
            "tags": tags, "attributes": component_attributes(probe, rules)}


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


def audit_untagged(store: Dict[str, Any], limit: int = 50) -> List[Dict[str, Any]]:
    """Most common normalized items that current rules still keep but do not tag."""
    rules = load_rules(refresh=True)
    ai = store.get("ai_tags") or {}
    rows: Dict[str, Dict[str, Any]] = {}
    for job, rec in (store.get("jobs") or {}).items():
        for item in rec.get("items") or []:
            raw = str(item.get("raw", "")).strip()
            if any(p.search(raw) for p in rules["skip"]):
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
        for it in rec.get("items") or []:
            derived = derive_item_fields(it, rules)
            it["norm"] = derived["norm"]
            it["tags"] = derived["tags"]
            it["attributes"] = derived["attributes"]
            qty = derived["qty"]
            price = derived["price"]
            ptype = derived["ptype"]
            if qty and not it.get("qty"):
                it["qty"] = qty
            if price and not it.get("price"):
                it["price"] = price
            if ptype and not it.get("ptype"):
                it["ptype"] = ptype
            n += 1
        apply_ai_cache(rec.get("items") or [], store)
    return n
