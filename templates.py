"""Quote-run TEMPLATES — a collection of the run formats we know how to read.

A quote (construction) run is the document that says how a highly-custom fan is
built. It is NOT one format: discovery (see WORKLOG.md, 2026-06-10) showed the
same logical document arrives as a `.txt`, an `.xlsx`, a `.pdf`, or an `.rtf`,
and **which format you get is mostly a function of the fan's design number** —
e.g. Design 64 fans carry a "D64 Wheel Construction" Excel sheet, HDX fans a
plain-text "Qt Run". So rather than one do-everything parser, this module is a
*registry of templates*: each template declares which runs it recognizes (by
design #, file extension, and/or file-name markers) and knows how to pull
fields out of that one shape. New fan formats are added by dropping one more
`QuoteRunTemplate` into `TEMPLATES` — that's the whole "collection".

Entry point:

    parse_quote_run(path, design=None) -> {
        "template":  "<template key that matched>",
        "design":    <int design # if known, else None>,
        "fields":    {label: value, ...},   # structured fields pulled from the run
        "raw_lines": [...],                  # first lines, for spotting new fields
        "summary":   "k1=v1; k2=v2; ...",    # compact one-liner for the report
    }

This is a superset of the old `drive_run.parse_drive_run_pdf` shape (which is
now just the PDF template), so it is a drop-in for the daily run / backfill.

Adding a template
-----------------
Subclass `QuoteRunTemplate`, set `key`/`label`, optionally `designs`,
`extensions`, and `name_patterns`, then implement `extract(ctx)`. Register it by
listing it in `TEMPLATES` ABOVE the generic fallbacks. Until a real sample pins
the exact field names down, an extractor can lean on the shared best-effort
key/value helpers — `fields` is free-form, so unknown labels are still captured
and surfaced. When you have a real dump, tighten that one template's `extract`.
"""
from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

log = logging.getLogger(__name__)


def _design_num(design: Any) -> Optional[int]:
    """'64' -> 64, '36P' -> 36, 'EMSI'/''/None -> None. Leading digits only."""
    m = re.match(r"\s*(\d+)", str(design or ""))
    return int(m.group(1)) if m else None


# ---------------------------------------------------------------------------
# Context: everything a template needs about one run file, loaded lazily so a
# template that only cares about the file name never pays to open the bytes.   #
# ---------------------------------------------------------------------------

class QuoteRunContext:
    """One run file presented to the registry. Heavy reads (text, pdf pages,
    workbook) are cached on first access and only happen if a template asks."""

    def __init__(self, path: str | Path, design: Any = None):
        self.path = Path(path)
        self.filename = self.path.name
        self.ext = self.path.suffix.lower()
        self.design = _design_num(design)
        self._text: Optional[str] = None

    @property
    def stem_lower(self) -> str:
        return self.filename.lower()

    def text(self) -> str:
        """Decoded text of a text-like run (.txt/.rtf), '' for binary formats.
        RTF is lightly de-marked-up — runs only need the words, not styling."""
        if self._text is None:
            self._text = self._read_text()
        return self._text

    def _read_text(self) -> str:
        if self.ext == ".docx":
            return _docx_to_text(self.path)
        if self.ext not in (".txt", ".rtf", ".csv", ""):
            return ""
        try:
            data = self.path.read_text(encoding="utf-8", errors="replace")
        except OSError as e:
            log.warning("Could not read quote-run text %s: %s", self.path, e)
            return ""
        return _rtf_to_text(data) if self.ext == ".rtf" else data


def _docx_to_text(path: Path) -> str:
    """Extract the visible text from a .docx (a zip of XML) with the stdlib only
    — no python-docx dependency. Paragraph and tab markers become newlines/tabs
    so labeled lines survive; all other tags are stripped."""
    import html
    import zipfile
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("word/document.xml").decode("utf-8", "replace")
    except (OSError, KeyError, zipfile.BadZipFile) as e:
        log.warning("Could not read .docx %s: %s", path, e)
        return ""
    xml = re.sub(r"<w:tab\b[^>]*/>", "\t", xml)
    xml = re.sub(r"</w:p>", "\n", xml)            # paragraph end -> newline
    xml = re.sub(r"<w:br\b[^>]*/>", "\n", xml)
    xml = re.sub(r"<[^>]+>", "", xml)             # drop the remaining tags
    return html.unescape(xml)


def _rtf_to_text(data: str) -> str:
    """Best-effort RTF -> plain text: drop control words and group braces. Good
    enough to recover field labels; not a full RTF reader."""
    data = re.sub(r"\\'[0-9a-fA-F]{2}", "", data)        # hex-escaped chars
    data = re.sub(r"\\par[d]?\b", "\n", data)            # paragraph breaks -> newlines
    data = re.sub(r"\\[a-zA-Z]+-?\d*\s?", "", data)      # other control words
    data = data.replace("{", "").replace("}", "")
    data = re.sub(r"[ \t]+\n", "\n", data)
    return data


# ---------------------------------------------------------------------------
# Shared best-effort field extraction (a "Label: value" pair on a line, or a   #
# two-column table row). Templates reuse these until a real sample lets us pin #
# the exact layout for that one format.                                        #
# ---------------------------------------------------------------------------

# Labels a construction run most likely carries — used only to ORDER the
# compact summary, never to gate capture (unknown labels are still kept).
FIELDS_OF_INTEREST = [
    "material", "construction", "gauge", "weld", "welder", "coating", "paint",
    "finish", "bearing", "shaft", "seal", "flange", "spark", "duty", "service",
    "wheel", "hub", "rim", "blade", "backplate",
]

# "Label: value" / "Label = value" on one line. Label is short-ish so a prose
# sentence with a stray colon doesn't masquerade as a field.
_KV_RE = re.compile(r"^\s*([A-Za-z][A-Za-z0-9 /%#.\-]{1,34}?)\s*[:=]\s*(.+?)\s*$")


def kv_from_lines(lines: List[str]) -> Dict[str, str]:
    fields: Dict[str, str] = {}
    for ln in lines:
        m = _KV_RE.match(ln)
        if m:
            label, val = m.group(1).strip(), m.group(2).strip()
            if label and val:
                fields.setdefault(label, val)  # first occurrence wins
    return fields


def kv_from_rows(rows: List[List[Any]]) -> Dict[str, str]:
    """Two-column rows (Label | value) — the spreadsheet/table layout."""
    fields: Dict[str, str] = {}
    for row in rows or []:
        cells = [str(c).replace("\n", " ").strip() for c in row if c is not None]
        cells = [c for c in cells if c != ""]
        if len(cells) == 2 and re.match(r"^[A-Za-z]", cells[0]) and len(cells[0]) <= 34:
            fields.setdefault(cells[0].rstrip(":"), cells[1])
    return fields


def summarize(fields: Dict[str, str]) -> str:
    """Compact one-liner — fields of interest first, then the rest, capped so it
    stays readable in a spreadsheet cell."""
    def rank(label: str) -> int:
        low = label.lower()
        for i, key in enumerate(FIELDS_OF_INTEREST):
            if key in low:
                return i
        return len(FIELDS_OF_INTEREST)
    ordered = sorted(fields.items(), key=lambda kv: (rank(kv[0]), kv[0]))
    return "; ".join(f"{k}={v}" for k, v in ordered[:12])


# ---------------------------------------------------------------------------
# Template base + the collection.                                              #
# ---------------------------------------------------------------------------

class QuoteRunTemplate:
    """One known run format. Declare what it matches; implement how to read it.

    Matching is a confidence score so the most specific template wins: a design
    number match and a file-name marker each add confidence on top of a handled
    extension. The fallbacks score 1, so something always matches.
    """
    key: str = "base"
    label: str = "Quote run"
    # Design numbers this format is used for (e.g. {64}); empty = any design.
    designs: frozenset = frozenset()
    # File extensions this template can read (lowercase, with the dot).
    extensions: frozenset = frozenset()
    # Case-insensitive regexes matched against the file NAME.
    name_patterns: tuple = ()
    # Case-insensitive substrings that must appear in the file's TEXT for this
    # template to apply (e.g. "CHICAGO BLOWER" for a CB engineering run). When
    # set, the template ONLY matches if a marker is present — it reads content.
    content_markers: tuple = ()
    # True when a shared extension means this template may ONLY claim a file if
    # its design # or name marker also matches (e.g. .txt is shared with the
    # generic text fallback, so a plain .txt must not be grabbed as a Qt Run).
    requires_signal: bool = False

    def __init__(self):
        self._name_res = tuple(re.compile(p, re.I) for p in self.name_patterns)

    def score(self, ctx: QuoteRunContext) -> int:
        """0 = does not apply. Higher = more confident this is the right reader."""
        if self.extensions and ctx.ext not in self.extensions:
            return 0
        design_hit = bool(self.designs and ctx.design in self.designs)
        name_hit = any(r.search(ctx.filename) for r in self._name_res)
        content_hit = False
        if self.content_markers:
            up = ctx.text().upper()
            content_hit = any(m.upper() in up for m in self.content_markers)
            if not content_hit:
                return 0  # a content-keyed template requires its marker
        if self.requires_signal and not (design_hit or name_hit or content_hit):
            return 0
        s = 1  # handled extension
        if design_hit:
            s += 2
        if name_hit:
            s += 3
        if content_hit:
            s += 4
        return s

    def extract(self, ctx: QuoteRunContext) -> Dict[str, Any]:
        """Return {"fields": {...}, "raw_lines": [...]}. Never raises — a bad
        file must not sink the daily run."""
        raise NotImplementedError


class _TextLineMixin:
    """Shared text -> lines + key/value extraction for text-shaped runs."""

    def _from_text(self, ctx: QuoteRunContext) -> Dict[str, Any]:
        lines = [ln.rstrip() for ln in ctx.text().splitlines() if ln.strip()]
        return {"fields": kv_from_lines(lines), "raw_lines": lines[:40]}


# Chicago Blower engineering "Qt Run" text — the selection-program dump. The
# useful data sits on a handful of consistent labeled lines (a comma-delimited
# spec line, a performance line, a wheel-construction table); the rest is a
# dimension table that a generic key/value sweep turns into noise. So pull the
# known fields by targeted pattern and leave the dimension tables out.
# A wheel-construction row's thickness column is a fraction ("1/4"), a decimal
# ("0.179"), or a decimal with the gauge in parens ("0.048 (18)", "0.179 ( 7)").
_GA = r"[\d./]+(?:\s*\(\s*\d+\s*\))?"
# (label, regex) — first match in the whole text wins; group(1) is the value.
_CB_PATTERNS = [
    ("Serial", r"SN#\s*(\d+)"),
    ("Size", r"\bSIZE\s+([0-9A-Za-z./\-]+)"),
    ("Design", r"\bDESIGN\s+([0-9A-Za-z./\-]+)"),
    ("Arrangement", r"\bARR\s+([0-9A-Za-z./\-]+)"),
    ("% Width", r"([\d.]+)\s*PCT\b"),
    ("Discharge", r"\bDISCH\s+([0-9A-Za-z./\-]+)"),
    ("Rotation", r"\bROT\s+([0-9A-Za-z./\-]+)"),
    ("Effective Wheel Dia", r"EFFECTIVE WHEEL DIA\.?\s+([\d /]+?)\s*$"),
    ("CFM", r"([\d.,]+)\s*CFM\b"),
    ("SP", r"([\d.]+)\s*SP\b"),
    ("BHP", r"([\d.]+)\s*BHP\b"),
    ("RPM", r"([\d.]+)\s*RPM\b"),
    ("Air Temp F", r"([\d.]+)\s*DEG F\b"),
    ("Density", r"DENSITY\s+([\d.]+)"),
    ("Max HP", r"MAX HP\s+([\d.]+)"),
    ("Max RPM", r"MAX RPM\s+([\d.]+)"),
    ("Max Temp F", r"MAX TEMP\s+([\d.]+)"),
    ("Ambient Temp F", r"AMBIENT TEMP\s+([\d.]+)"),
    ("Tip Speed FPM", r"TIP SPEED\s+([\d.]+)\s*FPM"),
    ("Shaft Dia", r"SHAFT DIA\s+([\d /]+?)\s*,"),
    ("Brg Centers", r"BRG CENTERS\s+([\d.]+)"),
    ("Critical Speed RPM", r"CRITICAL SPEED\s+([\d.]+)"),
    # Wheel-construction rows: <component> <gauge> <MATERIAL> <WR2> <weight>.
    # Component can carry a descriptor ("BLADES/2 RIB", "SIDEPL,SPUN").
    ("Blade Material", r"\bBLADES(?:/\d+\s*RIB)?\s+" + _GA + r"\s+([A-Z][A-Z0-9 .\-]+?)\s+\d+\s+\d"),
    ("Sideplate Material", r"\bSIDEPL\S*\s+" + _GA + r"\s+([A-Z][A-Z0-9 .\-]+?)\s+\d+\s+\d"),
    ("Backplate Material", r"\bBACKPLATE\b\s+" + _GA + r"\s+([A-Z][A-Z0-9 .\-]+?)\s+\d+\s+\d"),
    ("Liner Material", r"\bLINER\S*\s+" + _GA + r"\s+([A-Z][A-Z0-9 .\-]+?)\s+\d+\s+\d"),
    # Some runs have no construction table — just a single wheel-material line.
    ("Wheel Material", r"^\s*WHEEL MATERIAL\s+([A-Z0-9][A-Z0-9 .\-]+?)\s*$"),
    ("Class", r"\bCLASS\s+(\d+)\b"),
    ("Hub", r"\bHUB\s+([0-9][0-9\-]+)"),
    ("Coupling", r"\bCOUPLING\s+([A-Z][A-Z0-9 ]+?)\s*,\s*SIZE"),
]
# Compact summary, in engineering-useful order (only the present fields show).
_CB_SUMMARY_ORDER = [
    "Size", "Design", "Fan Type", "Arrangement", "% Width", "Discharge", "Rotation",
    "CFM", "SP", "BHP", "RPM", "Max Temp F", "Effective Wheel Dia", "Blade Material",
]


def _parse_chicago_blower(text: str) -> Dict[str, str]:
    fields: Dict[str, str] = {}
    for label, pat in _CB_PATTERNS:
        m = re.search(pat, text, re.I | re.M)
        if m and m.group(1).strip():
            fields[label] = re.sub(r"\s{2,}", " ", m.group(1).strip())

    lines = [ln.rstrip() for ln in text.splitlines()]
    # Serial: prefer "SN#NNNN"; some runs print the order number on a bare line.
    if "Serial" not in fields:
        for ln in lines[:8]:
            m = re.match(r"\s*(\d{5,7})\s*$", ln)
            if m:
                fields["Serial"] = m.group(1)
                break
    # Fan type: a descriptor line (e.g. "PACKAGED FORCED DRAFT FAN") right after
    # the SIZE/DESIGN spec line — letters only, so data/dimension lines are out.
    for i, ln in enumerate(lines):
        if re.search(r"\bSIZE\b.*\bDESIGN\b", ln, re.I):
            for nxt in lines[i + 1:i + 3]:
                s = nxt.strip()
                if re.fullmatch(r"[A-Z][A-Z ]{4,40}", s) and "WHEEL" not in s and s != "CONFIDENTIAL":
                    fields["Fan Type"] = s
                    break
            break

    up = text.upper()
    if "BELT DRIVEN" in up:
        fields["Drive"] = "Belt"
    elif "COUPLING" in up or ("DIRECT" in up and "DRIV" in up):
        fields["Drive"] = "Direct"
    if "ENGINEERING APPROVAL" in up:
        fields["Engineering Approval"] = "Required"
    if "FEA ANALYSIS REQUIRED" in up:
        fields["FEA Analysis"] = "Required"
    if "NON STD WHEEL MATERIAL" in up:
        fields["Non-Std Wheel Materials"] = "Yes"
    if "SHRINK FIT" in up:
        fields["Shrink Fit"] = "Yes"
    if "RUN TEST AT FACTORY" in up:
        fields["Factory Run Test"] = "Yes"
    return fields


def _cb_summary(fields: Dict[str, str]) -> str:
    parts = [f"{k}={fields[k]}" for k in _CB_SUMMARY_ORDER if k in fields]
    return "; ".join(parts)


class ChicagoBlowerQtRun(_TextLineMixin, QuoteRunTemplate):
    """The Chicago Blower selection-program "Qt Run" text (header: CHICAGO
    BLOWER CORP. / SN#...). Pulls the fan spec, duty/performance, wheel
    construction materials, drive, and shaft/bearing data by targeted pattern
    — the generic key/value sweep mis-reads its dimension tables."""
    key = "cbc_qt_run_text"
    label = "Chicago Blower Qt Run (text)"
    extensions = frozenset({".txt", ".rtf", ".docx"})
    name_patterns = (r"qt\s*run", r"quote\s*run")
    content_markers = ("CHICAGO BLOWER", "SN#")

    def extract(self, ctx: QuoteRunContext) -> Dict[str, Any]:
        text = ctx.text()
        fields = _parse_chicago_blower(text)
        lines = [ln.rstrip() for ln in text.splitlines() if ln.strip()]
        return {"fields": fields, "raw_lines": lines[:40], "summary": _cb_summary(fields)}


class QtRunText(_TextLineMixin, QuoteRunTemplate):
    """Any other `.txt`/`.rtf` run named like a quote run (non-Chicago-Blower).
    Best-effort key/value until a real dump of that format pins its headings."""
    key = "qt_run_text"
    label = "Qt Run (text, generic)"
    extensions = frozenset({".txt", ".rtf", ".docx"})
    name_patterns = (r"qt\s*run", r"quote\s*run")
    requires_signal = True  # plain .txt belongs to GenericTextRun

    def extract(self, ctx: QuoteRunContext) -> Dict[str, Any]:
        return self._from_text(ctx)


class D64WheelConstruction(QuoteRunTemplate):
    """Design 64 fans carry the run as a "D64 Wheel Construction" Excel sheet
    (inner/outer wheel). Best-effort cell sweep — adjacent Label|value pairs and
    in-cell "Label: value" — until a real sheet pins the exact cells."""
    key = "d64_wheel_construction"
    label = "D64 Wheel Construction (xlsx)"
    designs = frozenset({64})
    extensions = frozenset({".xlsx", ".xlsm"})
    name_patterns = (r"d\s*64\s+wheel\s+construction", r"wheel\s+construction")

    def extract(self, ctx: QuoteRunContext) -> Dict[str, Any]:
        try:
            import openpyxl
        except ImportError:
            log.warning("openpyxl not installed; cannot parse D64 wheel-construction xlsx")
            return {"fields": {}, "raw_lines": []}
        fields: Dict[str, str] = {}
        raw_lines: List[str] = []
        try:
            wb = openpyxl.load_workbook(str(ctx.path), read_only=True, data_only=True)
        except Exception as e:  # noqa: BLE001 - a bad workbook must not fail the run
            log.warning("Could not open D64 xlsx %s: %s", ctx.path, e)
            return {"fields": {}, "raw_lines": []}
        try:
            for ws in wb.worksheets:
                rows = [[c for c in row] for row in ws.iter_rows(values_only=True)]
                for k, v in kv_from_rows(rows).items():
                    fields.setdefault(k, v)
                for row in rows:
                    cells = [str(c).strip() for c in row if c is not None and str(c).strip()]
                    if cells:
                        if len(raw_lines) < 40:
                            raw_lines.append(" | ".join(cells))
                        for k, v in kv_from_lines(cells).items():
                            fields.setdefault(k, v)
        except Exception as e:  # noqa: BLE001
            log.warning("Could not scan D64 xlsx %s: %s", ctx.path, e)
        finally:
            wb.close()
        return {"fields": fields, "raw_lines": raw_lines}


class PdfQuoteRun(QuoteRunTemplate):
    """Any `.pdf` run (HDX and others file the construction run as a PDF). Reuses
    the position-aware PDF extraction in `drive_run.py`."""
    key = "pdf"
    label = "Quote run (pdf)"
    extensions = frozenset({".pdf"})

    def extract(self, ctx: QuoteRunContext) -> Dict[str, Any]:
        from drive_run import parse_drive_run_pdf
        r = parse_drive_run_pdf(ctx.path)
        return {"fields": r.get("fields", {}), "raw_lines": r.get("raw_lines", [])}


class GenericTextRun(_TextLineMixin, QuoteRunTemplate):
    """Fallback for any text-shaped run that didn't match a named template."""
    key = "generic_text"
    label = "Quote run (text, generic)"
    extensions = frozenset({".txt", ".rtf", ".csv", ".docx", ""})

    def extract(self, ctx: QuoteRunContext) -> Dict[str, Any]:
        return self._from_text(ctx)


class UnknownRun(QuoteRunTemplate):
    """Last resort: an extension we have no reader for (yet). Records nothing but
    keeps the pipeline flowing and names the format so it can be added."""
    key = "unknown"
    label = "Quote run (unrecognized format)"

    def score(self, ctx: QuoteRunContext) -> int:
        return 1  # always applies, but only as the absolute last choice

    def extract(self, ctx: QuoteRunContext) -> Dict[str, Any]:
        log.info("Quote run %s: no template reads %s files yet — add one to "
                 "templates.TEMPLATES.", ctx.filename, ctx.ext or "(no ext)")
        return {"fields": {}, "raw_lines": []}


# The collection. ORDER MATTERS for ties only (max score wins; first template
# wins a tie), so list the specific formats before the generic fallbacks.
TEMPLATES: List[QuoteRunTemplate] = [
    D64WheelConstruction(),
    ChicagoBlowerQtRun(),
    QtRunText(),
    PdfQuoteRun(),
    GenericTextRun(),
    UnknownRun(),
]


def match_template(ctx: QuoteRunContext) -> QuoteRunTemplate:
    """Pick the highest-confidence template for this run (first wins a tie)."""
    best, best_score = TEMPLATES[-1], 0
    for t in TEMPLATES:
        s = t.score(ctx)
        if s > best_score:
            best, best_score = t, s
    return best


def parse_quote_run(path: str | Path, design: Any = None) -> Dict[str, Any]:
    """Read a quote run with whichever template matches. Never raises."""
    res: Dict[str, Any] = {"template": "", "design": None, "fields": {},
                           "raw_lines": [], "summary": ""}
    try:
        ctx = QuoteRunContext(path, design)
        res["design"] = ctx.design
        template = match_template(ctx)
        res["template"] = template.key
        out = template.extract(ctx) or {}
        res["fields"] = out.get("fields", {}) or {}
        res["raw_lines"] = out.get("raw_lines", []) or []
        # A template may supply its own summary (its fields have a natural order);
        # otherwise rank the fields generically.
        res["summary"] = out.get("summary") or summarize(res["fields"])
    except Exception as e:  # noqa: BLE001 - belt-and-suspenders; extract already guards
        log.warning("Could not parse quote run %s: %s", path, e)
    return res


def main() -> None:
    """`python templates.py <path-to-run> [design#]` — show which template
    matched and what it pulled, so fields can be confirmed before wiring in."""
    import sys
    if len(sys.argv) < 2:
        raise SystemExit("Usage: python templates.py <path to a quote run> [design#]")
    path = sys.argv[1]
    design = sys.argv[2] if len(sys.argv) > 2 else None
    ctx = QuoteRunContext(path, design)
    print(f"\n{'=' * 72}\n{ctx.filename}\n{'=' * 72}")
    print(f"ext={ctx.ext or '(none)'}  design={ctx.design}")
    print("Template scores:")
    for t in TEMPLATES:
        print(f"  {t.score(ctx):>2}  {t.key:<24} {t.label}")
    r = parse_quote_run(path, design)
    print(f"\nMatched template: {r['template']}")
    print(f"Fields found ({len(r['fields'])}):")
    for k, v in r["fields"].items():
        print(f"  {k:<28} {v}")
    print(f"\nSummary: {r['summary']}")
    print("\n--- first reconstructed lines (for spotting fields to add) ---")
    for ln in r["raw_lines"]:
        print(f"  {ln}")
    print("\nPaste this back so the matching template's fields can be pinned down.")


if __name__ == "__main__":
    main()
