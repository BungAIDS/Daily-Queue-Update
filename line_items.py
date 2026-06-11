"""Sales-order line items: capture, normalize, store, and search.

Every parsed Sales Order yields its LINE ITEMS — the priced item/accessory
lines and the "Additional Features"-style feature lines that describe what was
actually sold (shaft seals, spark-resistant construction, coatings, isolators,
special motors, ...). The same option is rarely written the same way twice
("SS SHAFT SLEEVE" / "Stainless Steel Shaft Sleeve" / "316SS sleeve"), so each
captured line is kept three ways:

    raw   the line exactly as printed on the Sales Order (never altered)
    norm  a normalized form: uppercased, item numbers / qty / trailing prices
          stripped, punctuation collapsed, known abbreviations expanded
          (W/ -> WITH, SS -> STAINLESS STEEL, ...) so spelling variants of the
          same option converge
    tags  canonical feature tags matched by the rules table (SHAFT SEAL,
          SPARK RESISTANT, COATING, ...) — the lookup vocabulary

Everything lands in one resumable JSON store (LINE_ITEMS_STORE, default
BACKLOG_DIR/line_items.json), keyed by job:

    {"jobs":    {"421314": {"customer": ..., "co_number": 1, "so_pdf": ...,
                            "scanned_at": ..., "items": [{raw, norm, qty,
                            price, section, tags}, ...]}},
     "ai_tags": {"<norm>": ["TAG", ...]}}   # cached Claude classifications

The store is fed three ways: the daily run (sales_orders.py) records every
board job it parses, backfill_orders.py records each historical order, and
line_items_scan.py walks the already-archived PDFs under SALES_ORDER_DIR.
Search it with find_orders.py. Because `raw` is stored verbatim, the
normalization/tag rules can be tuned any time and re-applied with
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
from typing import Any, Dict, Iterable, List, Tuple

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
        "STL": "STEEL", "GALV": "GALVANIZED", "ALUM": "ALUMINUM",
        "CONST": "CONSTRUCTION", "CONSTR": "CONSTRUCTION",
        "ARR": "ARRANGEMENT", "ARRG": "ARRANGEMENT", "ARRGT": "ARRANGEMENT",
        "ASSY": "ASSEMBLY", "MTR": "MOTOR", "TEMP": "TEMPERATURE",
        "BRG": "BEARING", "BRGS": "BEARINGS", "SLV": "SLEEVE",
        "HSG": "HOUSING", "WHL": "WHEEL", "CONN": "CONNECTOR",
        "HVY": "HEAVY", "DTY": "DUTY", "HD": "HEAVY DUTY",
        "XP": "EXPLOSION PROOF", "EXPL": "EXPLOSION",
    },
    # Lines matching any of these are never items (case-insensitive regexes):
    # page furniture, totals/charges, address blocks, order metadata, and the
    # change-order history (captured separately as co_history).
    "skip_patterns": [
        r"^\s*c\s*/?\s*o\s*#?\s*\d",                  # CO#1 ... change-order notes
        r"^(sub\s*)?total\b", r"total\s+(billing|price|order)", r"^amount\s+due",
        r"^freight\b", r"^fob\b", r"^(sales\s+)?tax\b", r"surcharge",
        r"^page\s+\d+", r"^\d+\s+of\s+\d+$",
        r"^sales\s+order\b", r"^order\s+(no|number|date)\b", r"^job\s+(no|number)\b",
        r"^sold\s+to\b", r"^ship\s+to\b", r"^bill\s+to\b",
        r"^customer\s+(po|p\.o\.|order)\b", r"^p\.?\s*o\.?\s*(no|number|#)",
        r"^terms\b", r"^net\s+\d+", r"^quote\b", r"^quotation\b",
        r"^phone\b", r"^fax\b", r"^www\.", r"@",
        r"^date\b", r"^entered\s+by\b", r"^salesman\b", r"^rep\b",
        r"list\s+price", r"multiplier", r"^unit\s+price", r"^price\s+each",
        r"^discount\b",
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
    ],
    # Canonical feature tags: tag -> regexes matched against the NORMALIZED
    # text (uppercase, abbreviations expanded). An item can carry several.
    "tags": {
        "SPARK RESISTANT": [r"spark"],
        "SHAFT SEAL": [r"shaft\s*seal", r"stuffing\s*box", r"lip\s*seal",
                       r"ceramic\s*felt"],
        "SHAFT SLEEVE": [r"shaft\s*sleeve"],
        "SHAFT COOLER": [r"shaft\s*cooler", r"heat\s*slinger"],
        "STAINLESS STEEL": [r"stainless", r"\b304L?\b", r"\b316L?\b"],
        "HIGH TEMPERATURE": [r"high\s*temp", r"heat\s*fan"],
        "COATING": [r"epoxy", r"\bcoat", r"galvaniz", r"paint", r"primer",
                    r"plasite", r"heresite", r"\bzinc\b"],
        "LINING": [r"rubber\s*lin", r"\blined\b", r"\blining\b", r"abrasion"],
        "INSULATION": [r"insulat"],
        "VIBRATION ISOLATION": [r"isolat", r"rubber[\s-]*in[\s-]*shear",
                                r"\bRIS\b", r"spring\s*mount", r"seismic"],
        "VIBRATION SWITCH": [r"vibration\s*(switch|detector|monitor|sensor)"],
        "DAMPER": [r"damper", r"backdraft"],
        "INLET VANES": [r"inlet\s*vane", r"\bVIV\b", r"\bIVC\b",
                        r"variable\s*inlet"],
        "SILENCER": [r"silencer", r"muffler", r"sound\s*atten"],
        "ACCESS DOOR": [r"access\s*door", r"inspection\s*door", r"clean\s*out",
                        r"quick\s*open"],
        "DRAIN": [r"\bdrain"],
        "BELT GUARD": [r"belt\s*guard"],
        "WEATHER COVER": [r"weather\s*(cover|hood|proof)"],
        "SCREEN": [r"\bscreen"],
        "FLANGE": [r"flange"],
        "FLEX CONNECTOR": [r"flex(ible)?\s*conn", r"expansion\s*joint"],
        "UNITARY BASE": [r"unitary\s*base", r"structural\s*(steel\s*)?base",
                         r"channel\s*base"],
        "BEARINGS": [r"bearing"],
        "EXTENDED LUBE": [r"ext(ended)?\s*lube", r"lube\s*line", r"grease\s*line"],
        "MOTOR": [r"\bmotor\b"],
        "VFD": [r"\bVFD\b", r"variable\s*freq", r"inverter"],
        "EXPLOSION PROOF": [r"explosion\s*proof", r"class\s*i+\b.*div"],
        "V-BELT DRIVE": [r"v[\s-]*belt", r"sheave", r"bushing"],
        "BALANCE": [r"balanc"],
        "TESTING": [r"witness", r"\btest"],
        "SPARE PARTS": [r"spare"],
    },
}

_rules_cache: Dict[str, Any] | None = None


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
# Trailing money / no-charge marker: "$1,234.56", "1,234", "1234.56", "N/C",
# "NO CHARGE", "INCLUDED". A bare integer (no $ , or decimals) is NOT a price —
# it could be "3600 RPM". Items often end "qty  unit  extended", so the tail is
# stripped repeatedly.
_PRICE_TAIL = re.compile(
    r"""(?:^|\s)(?P<price>
          \$\s*\d[\d,]*(?:\.\d{2})?
        | \d{1,3}(?:,\d{3})+(?:\.\d{2})?
        | \d+\.\d{2}
        | N/?C\b | NO\s+CHARGE | INCL(?:UDED)?
        )\s*$""",
    re.I | re.X,
)
_EACH_TAIL = re.compile(r"(?:^|\s)(?:EA(?:CH)?|/EA|PER\s+UNIT)\s*$", re.I)
# Leading enumeration: "1.", "(1)", "1)", "ITEM 2:", "2 -". Qty guess only.
_LEAD = re.compile(r"^\s*(?:ITEM\s*)?\(?(?P<num>\d{1,3})\)?\s*(?:[.):]|-(?=\s))?\s+", re.I)
_NONDECIMAL_DOT = re.compile(r"(?<!\d)\.|\.(?!\d)")
_DIGITS_SS = re.compile(r"^(\d+)(SS)$", re.I)


def split_price_tail(text: str) -> Tuple[str, str]:
    """Strip trailing price columns / 'EACH' markers off a line. Returns
    (description, rightmost price found). Repeats so 'desc  1  847.00  847.00'
    loses the whole money tail."""
    price = ""
    for _ in range(4):
        m = _EACH_TAIL.search(text)
        if m:
            text = text[: m.start()].rstrip()
            continue
        m = _PRICE_TAIL.search(text)
        if not m:
            break
        if not price:
            price = m.group("price").strip()
        text = text[: m.start()].rstrip()
    return text.strip(), price


def split_lead(text: str) -> Tuple[str, str]:
    """Strip a leading item-number/qty marker. Returns (rest, qty guess)."""
    m = _LEAD.match(text)
    if not m:
        return text.strip(), ""
    return text[m.end():].strip(), m.group("num")


def normalize_text(raw: str, rules: Dict[str, Any] | None = None) -> str:
    """The canonical lookup form of a line: uppercase, enumeration/qty and
    price columns stripped, known abbreviations expanded, punctuation collapsed
    to spaces (decimals inside numbers survive). Spelling variants of the same
    option converge here — this is what search and tagging run against."""
    rules = rules or load_rules()
    s, _ = split_lead(raw.strip().upper())
    s, _ = split_price_tail(s)
    abbrev = rules["abbreviations"]
    out: List[str] = []
    for tok in s.split():
        t = tok.strip(".,;:()[]")
        if t in abbrev:
            out.append(abbrev[t])
            continue
        m = _DIGITS_SS.match(t)  # "316SS" -> "316 STAINLESS STEEL"
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
    """Canonical feature tags whose patterns match the normalized text."""
    rules = rules or load_rules()
    return sorted(t for t, pats in rules["tags"].items() if any(p.search(norm) for p in pats))


# --------------------------------------------------------------------------- #
# Extraction from a Sales Order's text lines                                  #
# --------------------------------------------------------------------------- #
def classify_line(line: str, section: str, rules: Dict[str, Any]) -> Tuple[str, str]:
    """How one reconstructed text line is treated. Returns (kind, detail):

        blank          empty line
        section-start  opens a feature section (detail = its title)
        section-end    closes it (detail = the matched pattern)
        skip           never an item (detail = the skip pattern that hit)
        item-priced    an item line — it ends in a price / N/C column
        item-section   an unpriced line captured because a section is open
        text           everything else (ignored)

    Single source of truth for extract_items AND the --dump tuning view, so
    what the dump shows is exactly what the extractor does."""
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
    _, price = split_price_tail(s)
    if price:
        return "item-priced", price
    if section:
        return "item-section", section
    return "text", ""


def extract_items(lines: Iterable[str], rules: Dict[str, Any] | None = None) -> List[Dict[str, Any]]:
    """Pull the line items out of a Sales Order's reconstructed text lines.

    Two capture signals, covering both ways SOs lay items out: any line ending
    in a price/N-C column is an item wherever it sits, and every line inside a
    recognized feature section ("Additional Features", "Accessories", ...) is
    an item even unpriced. Duplicate lines (same normalized form) collapse to
    one. Items come back as {raw, norm, qty, price, section, tags}."""
    rules = rules or load_rules()
    section = ""
    by_norm: Dict[str, Dict[str, Any]] = {}
    for line in lines:
        s = re.sub(r"\s+", " ", str(line)).strip()
        kind, detail = classify_line(s, section, rules)
        if kind == "section-start":
            section = detail
            continue
        if kind == "section-end":
            section = ""
            continue
        if kind not in ("item-priced", "item-section"):
            continue
        body, qty = split_lead(s)
        body, price = split_price_tail(body)
        norm = normalize_text(body, rules)
        if len(norm) < 3:  # stray cell fragments ("X", "1") aren't items
            continue
        prev = by_norm.get(norm)
        if prev is not None:
            if price and not prev["price"]:
                prev["price"] = price
            continue
        by_norm[norm] = {
            "raw": s,
            "norm": norm,
            "qty": qty,
            "price": price,
            "section": section,
            "tags": tag_item(norm, rules),
        }
    return list(by_norm.values())


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
    """Merge cached Claude classifications into each item's tags (in place)."""
    ai = store.get("ai_tags") or {}
    for it in items:
        extra = ai.get(it.get("norm", ""))
        if extra:
            it["tags"] = sorted(set(it.get("tags") or []) | set(extra))


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
    """One search term vs one item: substring of the normalized/raw text or any
    tag; with fuzzy > 0, also a close difflib match against any same-length
    word window of the normalized text (catches typos/odd spellings)."""
    t = term.upper().strip()
    hay = [item.get("norm", ""), item.get("raw", "").upper()] + [
        x.upper() for x in item.get("tags") or []]
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
    using the CURRENT rules (and the AI cache). Run after editing the rules
    file — nothing is re-downloaded or re-parsed. Returns the item count."""
    rules = load_rules(refresh=True)
    n = 0
    for rec in (store.get("jobs") or {}).values():
        for it in rec.get("items") or []:
            body, qty = split_lead(it.get("raw", ""))
            body, price = split_price_tail(body)
            it["norm"] = normalize_text(body, rules)
            it["tags"] = tag_item(it["norm"], rules)
            if qty and not it.get("qty"):
                it["qty"] = qty
            if price and not it.get("price"):
                it["price"] = price
            n += 1
        apply_ai_cache(rec.get("items") or [], store)
    return n
