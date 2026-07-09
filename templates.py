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
        if self.ext not in (".txt", ".rtf", ".csv", ".md", ""):
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
        # Store the whole document (RAW_LINES_CAP is a runaway backstop, not a
        # real limit) so non-CB text runs are a full corpus for future patterns.
        return {"fields": kv_from_lines(lines), "raw_lines": lines[:RAW_LINES_CAP]}


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
    # Shaft/rotor geometry line, comma-delimited under the bearing line:
    #   LENGTH 35 3/8 ,OH 8.64,BX 15 3/8 , STB 12 , TG&P 68   +   STH 2.12
    ("Shaft Length", r"\bLENGTH\s+([\d./ ]+?)\s*,"),
    ("OH", r"\bOH\s+([\d.]+)"),
    ("BX", r"\bBX\s+([\d./ ]+?)\s*,"),
    ("STB", r"\bSTB\s+([\d./ ]+?)\s*,"),
    ("TG&P", r"TG&P\s+(\d+(?:\s+\d+/\d+)?)"),
    ("STH", r"^\s*STH\s+([\d.]+)"),
    # Bearing spec block: "SIZE 2 15/16 BEARINGS, LINK BELT SERIES 6800" and the
    # DRIVE-FLOAT row's L10 hours (…STATIC DYN THRUST <L10> P/C).
    ("Bearing Size", r"\bSIZE\s+([\d /]+?)\s+BEARINGS"),
    ("Bearing Series", r"BEARINGS,\s*(.+?SERIES\s+\d+)"),
    # Bearing-loads rows: "DRIVE-FIXED   301   3   0   47882 0.0494" (or -FLOAT;
    # static can be negative; some layouts add extra columns). L10 is the big
    # integer right before the P/C decimal, so anchor on that, not on position.
    ("Bearing L10 Hr", r"DRIVE-(?:FIXED|FLOAT)[^\n]*?(\d{4,})\s+0?\.\d+"),
    ("Drive Brg Mount", r"DRIVE-(FIXED|FLOAT)"),
    ("Drive Brg Static Lb", r"DRIVE-(?:FIXED|FLOAT)\s+(-?\d+)"),
    ("Other Brg Mount", r"OTHER-(FIXED|FLOAT)"),
    ("Other Brg Static Lb", r"OTHER-(?:FIXED|FLOAT)\s+(-?\d+)"),
    ("Brg Thrust Lb", r"OTHER-(?:FIXED|FLOAT)\s+-?\d+\s+-?\d+\s+(-?\d+)"),
    # Rotor line (arr 1/3/7/8/9): "ROTOR WR2  267 LB-FT2, ROTOR MAX RPM 2860, MTL. 1045 STEEL"
    ("Rotor WR2", r"ROTOR WR2\s+([\d.]+)"),
    ("Rotor Max RPM", r"ROTOR MAX RPM\s+([\d.]+)"),
    ("Rotor Material", r"ROTOR MAX RPM[^\n]*MTL\.\s*(.+?)\s*$"),
    ("Stress Ratio at Hub", r"STRESS RATIO AT HUB\s+([\d.]+)"),
    ("Stress Ratio at Bearing", r"STRESS RATIO AT HUB[^\n]*AT BEARING\s+([\d.]+)"),
    # Wheel dynamics line (all arrangements): "MAX RPM, WHEEL ONLY 2346.  10 BLADES.  RES 3110 CPM"
    ("Max RPM Wheel Only", r"MAX RPM,?\s*WHEEL ONLY\s+(\d+(?:\.\d+)?)"),
    ("Blades", r"(\d+)\s+BLADES\."),
    ("Wheel Resonance CPM", r"\bRES\s+([\d.]+)\s*CPM"),
    # Arr-4-family wheel/housing block (wheel on the motor shaft — these runs
    # have no shaft/bearing section; this block is what they carry instead):
    #   WHEEL WEIGHT  267 LB, THRUST  190 LB, WR2  387 LB-FT2
    #   HOUSING TO WHEEL CG   5.33, HOUSING TO HUB INLET FACE  6.38 IN.
    ("Wheel Weight Lb", r"WHEEL WEIGHT\s+([\d.]+)\s*LB"),
    ("Wheel Thrust Lb", r"WHEEL WEIGHT[^\n]*THRUST\s+([\d.]+)\s*LB"),
    ("Wheel WR2", r"WHEEL WEIGHT[^\n]*\bWR2\s+([\d.]+)"),
    ("Housing to Wheel CG", r"HOUSING TO WHEEL CG\s+([\d.]+)"),
    ("Housing to Hub Inlet Face", r"HOUSING TO HUB INLET FACE\s+([\d.]+)"),
    # Belt-drive sheave data (arr 9/1): "FAN SHEAVE  ASSUMED PD  6.2 IN, 4000 FPM"
    # and "AT 2585 RPM, MIN PD  3.0 IN, 2030 FPM". The PD is ASSUMED (auto-sized)
    # or SPECIFIED (customer item-59 override) — both are the fan sheave PD.
    ("Sheave PD", r"(?:ASSUMED|SPECIFIED)\s+PD\s+([\d.]+)\s*IN"),
    ("Min Sheave PD", r"MIN PD\s+([\d.]+)"),
    # Motor / base (motor-mounted arrangements): "BASE , 405T FR MOTOR",
    # "ADJUSTABLE BASE, 365TS FR MOTOR, MOTOR POSITION Z".
    ("Motor Frame", r"\b(\d{2,3}T?S?)\s+FR\.?\s+MOTOR"),
    ("Motor Position", r"MOTOR POSITION\s+([A-Z])\b"),
    ("Motor Weight Lb", r"MOTOR WEIGHT(?: AND MOUNTING)?[^0-9\n]*(\d+)"),
    # Housing construction: "HOUSING  7 GA C.Q. HRS ..." / "HOUSING 3/16 A240
    # 304L SS ..." / arr-7 "SPLIT HOUSING AND BOX  3/8 C.Q. HRS ...". The
    # capture must end on a letter so trailing weight/price columns fall off.
    ("Housing Construction",
     r"^\s*(?:SPLIT\s+)?HOUSING(?:\s+AND\s+BOX)?\s+((?:\d+\s*GA|[\d/]+)\s+[A-Z][A-Za-z0-9 .\-]*?[A-Za-z])(?:\s+\d+)*\s*$"),
    ("Stiffeners", r"STIFFENERS\s+(SK-[0-9-]+\s+\w+\s*PRESSURE(?:,\s*\d+\s*IN\.?\s*CENTERS)?)"),
    ("Fan Outlet Area FT2", r"FAN OUTLET AREA\s+([\d.]+)\s*FT2"),
    ("Inlet Box Size", r"INLET BOX SIZE\s+([\d.]+)"),
    ("Shaft Seal Height", r"SPECIAL SHAFT SEAL(?: OR SPACER)? HEIGHT\s+([\d.]+)"),
    # Quote totals: "GOOD FOR 60 DAYS   3773   32851" (weight, price).
    ("Total Weight Lb", r"GOOD FOR\s+\d+\s*DAYS\s+(\d+)\s+\d+"),
    ("Total Price", r"GOOD FOR\s+\d+\s*DAYS\s+\d+\s+(\d+)"),
    # The AXIAL/SIDE VIEW outline-dimension table is pulled as a block by
    # _parse_outline_dims (below) rather than one pattern per code.
    # Wheel-construction rows: <component> <gauge> <MATERIAL> <WR2> <weight>.
    # Component can carry a descriptor ("BLADES/2 RIB", "SIDEPL,SPUN").
    ("Blade Material", r"\bBLADES(?:/\d+\s*RIB)?\s+" + _GA + r"\s+([A-Z][A-Z0-9 .\-]+?)\s+\d+\s+\d"),
    ("Sideplate Material", r"\bSIDEPL\S*\s+" + _GA + r"\s+([A-Z][A-Z0-9 .\-]+?)\s+\d+\s+\d"),
    ("Backplate Material", r"\bBACKPLATE\b\s+" + _GA + r"\s+([A-Z][A-Z0-9 .\-]+?)\s+\d+\s+\d"),
    ("Liner Material", r"\bLINER\S*\s+" + _GA + r"\s+([A-Z][A-Z0-9 .\-]+?)\s+\d+\s+\d"),
    # The THICK.(GA) column of each wheel-construction row — the gauge/thickness
    # ("1/4", "3/8", "0.048 (18)") that sits between the component and its
    # material. Same anchors as the material patterns, capturing _GA instead.
    ("Blade Gauge", r"\bBLADES(?:/\d+\s*RIB)?\s+(" + _GA + r")\s+[A-Z]"),
    ("Sideplate Gauge", r"\bSIDEPL\S*\s+(" + _GA + r")\s+[A-Z]"),
    ("Backplate Gauge", r"\bBACKPLATE\b\s+(" + _GA + r")\s+[A-Z]"),
    ("Liner Gauge", r"\bLINER\S*\s+(" + _GA + r")\s+[A-Z]"),
    # Some runs have no construction table — just a single wheel-material line.
    ("Wheel Material", r"^\s*WHEEL MATERIAL\s+([A-Z0-9][A-Z0-9 .\-]+?)\s*$"),
    ("Class", r"\bCLASS\s+(\d+)\b"),
    ("Hub", r"\bHUBS?\s+([0-9][0-9\-]+)"),
    ("Coupling", r"\bCOUPLING\s+([A-Z][A-Z0-9 ]+?)\s*,\s*SIZE"),
    # Fabricated hub (no cast part number): a HUB TUBE / HUB FLANGES / HUB
    # CENTERS block in the wheel-construction table, each a gauge + material row
    # like the blade rows. e.g. "HUB TUBE  3/4  AISI 1026 HFSM  1  23".
    ("Hub Tube Gauge", r"\bHUB TUBE\s+(" + _GA + r")\s+[A-Z]"),
    ("Hub Tube Material", r"\bHUB TUBE\s+" + _GA + r"\s+([A-Z][A-Z0-9 .\-]+?)\s+\d+\s+\d"),
    ("Hub Flanges Gauge", r"\bHUB FLANGES\s+(" + _GA + r")\s+[A-Z]"),
    ("Hub Flanges Material", r"\bHUB FLANGES\s+" + _GA + r"\s+([A-Z][A-Z0-9 .\-]+?)\s+\d+\s+\d"),
    ("Hub Centers Material", r"\bHUB CENTERS\s+\S+\s+([A-Z][A-Z0-9 .\-]+?)\s+\d+\s+\d"),
    # Hub bore/OD — present on both cast ("HUB 19-5-16 BORE 2 15/16 CAST IRON")
    # and fabricated ("HUB BORE 2 11/16, HUB OD 7.00") hubs.
    ("Hub Bore", r"\bBORE\s+([\d ./]+?)\s*(?:,|CAST|C\.\s*IRON|HUB OD|$)"),
    ("Hub OD", r"\bHUB OD\s+([\d.]+)"),
    ("Hub Bushing", r"\b(Q\d\s+BUSHING\s+[\d ./]+)"),
    # Shaft/coupling geometry (belt & coupled runs): the half-coupling shaft
    # diameters and keyway, a large uncaptured cluster.
    ("Coupling Max Shaft Dia", r"MAX SHAFT DIA(?:METER)?\s+AT FAN SHAFT HALF COUPLING\s*=\s*([\d.]+)"),
    ("Coupling Min Shaft Dia", r"MIN SHAFT DIA(?:METER)?\s+AT FAN SHAFT HALF COUPLING\s*=\s*([\d.]+)"),
    ("Coupling Nom Shaft Dia", r"NOM SHAFT DIA(?:METER)?\s+AT FAN SHAFT HALF COUPLING\s*=\s*([\d.]+)"),
    ("Coupling Keyway", r"KEYWAY DIMENSIONS FOR HALF COUPLING\s*=\s*([\d. /]+X[\d. /]+)"),
    # Inlet-box / damper box dimensions (fans with an inlet box or damper):
    # "BOX B X C:  80  IN. X  17  IN." and the box code on the spec line.
    ("Box B", r"BOX B X C:\s+([\d ./]+?)\s+IN"),
    ("Box C", r"BOX B X C:[^\n]*?IN\.?\s*X\s+([\d ./]+?)\s+IN"),
    # Inlet-box orientation angle on the spec line (",BOX 270"). BOX 0 means no
    # inlet box, so only a non-zero angle is captured.
    ("Inlet Box Angle", r"\bBOX\s+([1-9]\d*)\b"),
    ("Inlet Cone", r"REINFORCED INLET CONE INCL\. PER (SK[E]?-[0-9-]+)"),
]
# Compact summary, in engineering-useful order (only the present fields show).
_CB_SUMMARY_ORDER = [
    "Size", "Design", "Fan Type", "Arrangement", "% Width", "Discharge", "Rotation",
    "Class", "CFM", "SP", "BHP", "RPM", "Max Temp F", "Effective Wheel Dia",
    # Wheel construction — material paired with its gauge, the detail the report
    # cares about most (aero fields above already come from the Sales Order).
    "Blade Material", "Blade Gauge", "Sideplate Material", "Sideplate Gauge",
    "Backplate Material", "Backplate Gauge", "Liner Material", "Liner Gauge",
    "Wheel Material", "Hub", "Coupling", "Blades",
    # Shaft / bearing section (shaft geometry + bearing spec + key outline dims).
    "Shaft Dia", "Brg Centers", "Critical Speed RPM",
    "BX", "STB", "OH", "STH", "Bearing Size", "Bearing Series",
    "Housing Width (N)", "Base to CL (F)",
    # Motor/base block (the arr-4 family's back half) + drive.
    "Motor Frame", "Motor Position", "Motor Enclosure", "Drive",
]


# Outline (AXIAL/SIDE VIEW) dimension codes -> friendly names. The section lists
# one dimension per line as "<code>  <DESCRIPTION>  <inches>  <mm>". Codes not in
# this map are still captured, named from their own description text.
_OUTLINE_DIMS = {
    "A": "Discharge Height (A)", "W": "Bottom of Disch to CL (W)",
    "KK": "Outlet/Inlet Flange (KK)", "E": "Discharge Flange to CL (E)",
    "RB": "Unitary Base to CL (RB)", "RM": "Motor CL to Fan CL (RM)",
    "F/2": "Base to CL (F)", "H": "Shaft Height (H)",
    "TV": "Total Vert Height (TV)", "RH": "Max Right of CL to Hsg (RH)",
    "LH": "Max Left of CL to Disch (LH)", "MA": "Mounting Channel (MA)",
    "D": "Base Flange to CL (D)", "K": "Drive End of Shaft to CL (K)",
    "N": "Housing Width (N)", "LR": "Mtg Flange to CL (LR)",
}
# One outline row: a short left code, a text description, the inches value (whole
# + optional fraction), then the mm column at line end. Anchored top-and-tail so
# the flange-punching "A = .." table and the part-cost tables can't leak in.
_OUTLINE_ROW = re.compile(
    r"^\s*([A-Z]{1,3}(?:/\d)?)\s{2,}([A-Z][A-Z0-9 ./&\-]+?)\s{2,}"
    r"(\d+(?:\s+\d+/\d+)?)\s{2,}\d+\s*$")


def _parse_outline_dims(text: str) -> Dict[str, str]:
    """Pull the whole AXIAL/SIDE VIEW outline-dimension table (one dim per line),
    bounded to that section. Returns {friendly name: inches value}."""
    lines = text.splitlines()
    try:
        start = next(i for i, ln in enumerate(lines) if "AXIAL VIEW" in ln.upper())
    except StopIteration:
        return {}
    out: Dict[str, str] = {}
    for ln in lines[start + 1:]:
        up = ln.upper()
        if "PUNCHING DETAIL" in up or "PART NAME" in up or ln.strip().startswith("---"):
            break
        m = _OUTLINE_ROW.match(ln)
        if not m:
            continue
        code, desc, inches = m.group(1), m.group(2).strip(), m.group(3)
        name = _OUTLINE_DIMS.get(code) or f"{desc.title()} ({code})"
        out.setdefault(name, re.sub(r"\s{2,}", " ", inches.strip()))
    return out


# A run file sometimes quotes the SAME fan in two mounting arrangements — an
# arrangement-4 (wheel on the motor shaft) and an arrangement-8/9 (fan on its
# own bearings) — as two printouts in one document. Per DG the arrangement-4 is
# the built unit, so it wins; the others are dropped BEFORE parsing, otherwise
# the 8/9 run's bearing/rotor section leaks onto the (bearing-less) 4 fan.
_ARR_LINE = re.compile(r"\bARR\s+(\d[0-9A-Z]*)", re.I)
_PAGE_BOUNDARY = re.compile(r"^\s*-{15,}\s*$")


def select_primary_run_text(text: str) -> str:
    """When a document concatenates runs in more than one arrangement family and
    one of them is arrangement 4, return only that run's pages. Single-run docs
    (one family) are returned unchanged."""
    lines = text.splitlines()
    pages: List[List[str]] = []
    cur: List[str] = []
    for ln in lines:
        if _PAGE_BOUNDARY.match(ln):
            if cur:
                pages.append(cur)
            cur = []
        else:
            cur.append(ln)
    if cur:
        pages.append(cur)
    if len(pages) <= 1:
        return text
    arr_of: List[Optional[str]] = []
    last: Optional[str] = None
    for pg in pages:
        m = _ARR_LINE.search("\n".join(pg))
        a = m.group(1).upper().rstrip(",") if m else last
        arr_of.append(a)
        last = a or last
    families = {a[0] for a in arr_of if a}
    if "4" not in families or len(families) <= 1:
        return text                       # no arr-4 to prefer, or a single run
    kept = ["\n".join(pg) for pg, a in zip(pages, arr_of) if a and a[0] == "4"]
    return "\n".join(kept) if kept else text


def _parse_chicago_blower(text: str) -> Dict[str, str]:
    """Parse a CB run. For a dual-arrangement doc (a fan quoted as arr-4 AND
    arr-8/9), the arr-4 is authoritative (DG's rule), but anything in the arr-8
    section that the arr-4 didn't provide is kept — the arr-4 never gets
    overwritten, the 8/9 only fills gaps (extra bearing/outline/accessory data)."""
    primary = select_primary_run_text(text)
    fields = _cb_extract(primary)
    if primary != text:                    # a dual doc was trimmed to arr-4
        for k, v in _cb_extract(text).items():
            fields.setdefault(k, v)         # keep non-conflicting arr-8 data
    return fields


def _cb_extract(text: str) -> Dict[str, str]:
    fields: Dict[str, str] = {}
    for label, pat in _CB_PATTERNS:
        m = re.search(pat, text, re.I | re.M)
        if m and m.group(1).strip():
            fields[label] = re.sub(r"\s{2,}", " ", m.group(1).strip())
    for name, val in _parse_outline_dims(text).items():
        fields.setdefault(name, val)

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
    elif re.search(r"\bFR\.?\s+MOTOR", up):
        fields["Drive"] = "Motor mounted"   # arr 4 family: wheel on the motor shaft
    # Motor enclosure: printed in the machine-readable tail as 113 'TEFC', and
    # sometimes in prose. Restricted to known enclosure codes to avoid noise.
    m = re.search(r"'(TEFC|ODP|WPII|XPFC)'|\b(TEFC|ODP|WPII|XPFC)\b", up)
    if m:
        fields["Motor Enclosure"] = m.group(1) or m.group(2)
    if "FLANGED INLET" in up:
        fields["Flanged Inlet"] = ("Included" if "FLANGED INLET  INCLUDED" in up
                                   or "FLANGED INLET INCLUDED" in up else "Yes")
    if "SHAFT SEAL NOT INCLUDED" in up:
        fields["Shaft Seal"] = "Not included"
    elif "SHAFT SEAL" in up:
        fields.setdefault("Shaft Seal", "Included")
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
    if "SHAFT SAFETY GUARD" in up:
        fields["Safety Guard"] = "Yes"
    return fields


def _cb_summary(fields: Dict[str, str]) -> str:
    parts = [f"{k}={fields[k]}" for k in _CB_SUMMARY_ORDER if k in fields]
    return "; ".join(parts)


# --- Coverage tagging: "we read right over it because it doesn't match" -------
# DG's ask: when a run carries recognizable engineering data (hub construction,
# a sheave PD, a half-coupling) but our patterns produced NO field for that
# family, we must TAG the run so the miss is visible instead of silently dropped.
# Each probe is (tag, keyword-present regex, the field keys that mean the family
# WAS structured). Keyword present + none of those keys captured = a real miss.
# The tag is self-clearing: add a pattern that captures the family and the tag
# disappears on the next reparse — the tag count is a live coverage metric.
_COVERAGE_PROBES = [
    # A hub-construction line: a hub keyword, "HUB SPECIAL", or a hub part
    # number ("HUB 19-5-218", "HUBS 19-5-21"). NOT "STRESS RATIO AT HUB 0.34"
    # (a ratio, not a hub) — hence the AT-HUB exclusion and the dash requirement
    # on the part-number form.
    ("hub", re.compile(
        r"(?<!AT )\bHUBS?\s+(?:TUBE|FLANGES|CENTERS|BORE|OD|HARDWARE|SPECIAL|\d+\s*-)", re.I),
     ("Hub", "Hub Tube Gauge", "Hub Tube Material", "Hub Flanges Gauge",
      "Hub Flanges Material", "Hub Centers Material", "Hub Bore", "Hub OD",
      "Hub Bushing")),
    ("coupling", re.compile(r"\bHALF COUPLING\b|\bCOUPLING\b", re.I),
     ("Coupling", "Coupling Max Shaft Dia", "Coupling Min Shaft Dia",
      "Coupling Nom Shaft Dia", "Coupling Keyway")),
    ("inlet_box", re.compile(r"\bBOX B X C\b|\bINLET BOX\b", re.I),
     ("Box B", "Box C", "Inlet Box Size")),
    ("sheave", re.compile(r"\bFAN SHEAVE\b", re.I),
     ("Sheave PD", "Min Sheave PD")),
    ("inlet_cone", re.compile(r"REINFORCED INLET CONE", re.I),
     ("Inlet Cone",)),
]

# Lines that carry no field-worthy data — timestamps, the numeric parameter
# dump, item-code dumps, inquiry/order numbers, note banners. Excluded from the
# "missed data lines" review list so what's left is genuine uncaptured content.
_MISS_NOISE = re.compile(
    r"^[\s\-=*_/.]*$"
    r"|^\s*\d+\s+[\d.]+\s*,"                       # param dump: '  10  2800.0000,'
    r"|'\s*[A-Z0-9 ]{1,4}\s*'"                     # item-code dump: "2 'LS  '"
    r"|CONFIDENTIAL|CHICAGO BLOWER|INQUIRY|SN#|ORDER #"
    r"|\*\*\*"                                     # note banners
    r"|^\s*(?:MON|TUE|WED|THU|FRI|SAT|SUN)\b"      # timestamp lines
    r"|^\s*\d{3}-\d{2}-\d+\s*$"                    # bare inquiry number
    r"|^\s*(?:\d+\s+){5,}\d*\s*$"                  # numeric matrix rows
    r"|NOTE|ACCEPTABLE|NOT VALID|APPROVED BY|SPECIFY ITEM"
    r"|BASED ON|OTHERWISE|ASSUMED|CAN SPECIFY|IGNORED|CAN TRY",
    re.I)
_MISS_NUM = re.compile(r"\d+(?:\s+\d+/\d+|/\d+|\.\d+)?")


def _num_tokens(s: str) -> set:
    return {m.group(0).strip() for m in _MISS_NUM.finditer(s)}


def coverage_tags(text: str, fields: Dict[str, str]) -> List[str]:
    """High-precision tags: a data family whose keyword is in the text but which
    produced no field — the "we read right over it" case DG flagged. Empty for a
    fully-captured run. Returned sorted for stable output/tests."""
    up = (text or "").upper()
    tags = []
    for tag, kw, keys in _COVERAGE_PROBES:
        if kw.search(up) and not any(fields.get(k) for k in keys):
            tags.append(tag)
    return sorted(tags)


def missed_data_lines(text: str, fields: Dict[str, str], cap: int = 15) -> List[str]:
    """The review aid behind the coverage tags: the actual document lines that
    carry a value not reflected in any captured field — so a human (or the next
    pattern-design pass) can see exactly what's still being read over. Noise
    (timestamps, param/code dumps, banners) is filtered out; capped so a run
    with a long uncaptured tail doesn't bloat the store."""
    captured_nums: set = set()
    captured_words: set = set()
    for k, v in fields.items():
        captured_nums |= _num_tokens(str(v))
        captured_words |= {t for t in re.split(r"[^A-Za-z]+", k.upper()) if len(t) > 2}
    out: List[str] = []
    seen = set()
    for ln in (text or "").splitlines():
        s = ln.strip()
        if len(s) < 8 or _MISS_NOISE.search(s) or not any(c.isdigit() for c in s):
            continue
        nums = _num_tokens(s)
        if not nums or nums <= captured_nums:
            continue                                  # every value already captured
        label = re.split(r"\d", s, 1)[0]
        lwords = {t for t in re.split(r"[^A-Za-z]+", label.upper()) if len(t) > 2}
        if lwords & captured_words:
            continue                                  # label maps to a captured field
        if s not in seen:
            seen.add(s)
            out.append(s)
        if len(out) >= cap:
            break
    return out


# The header that marks the CBC selection-program "Qt Run" layout — the
# structured dump the parser above understands. It shows up the same whether the
# run was saved as .txt, .docx, or .pdf, so any of those can route to that
# parser. (It does NOT mean "is this Chicago Blower" — every run here is; it
# means "is this the selection-program output" vs. some other doc that merely
# matched the file name, like a vendor quote or a markup.)
SELECTION_PROGRAM_MARKERS = ("CHICAGO BLOWER", "SN#")

# How many document lines a template returns as raw_lines. The sweep persists
# these to its store, making the store a re-parsable corpus (design new patterns,
# re-extract without re-reading Z:), so a truncated tail = fields we can't
# re-parse offline. This is NOT a functional limit — it's a pure runaway
# backstop set far above any real run (the longest, a dual 4S/8S 8-pager, is
# well under 1000 lines). No real quote run is ever truncated.
RAW_LINES_CAP = 10000


def is_selection_program(text: str) -> bool:
    up = (text or "").upper()
    return any(m in up for m in SELECTION_PROGRAM_MARKERS)


class ChicagoBlowerQtRun(_TextLineMixin, QuoteRunTemplate):
    """The Chicago Blower selection-program "Qt Run" text (header: CHICAGO
    BLOWER CORP. / SN#...). Pulls the fan spec, duty/performance, wheel
    construction materials, drive, and shaft/bearing data by targeted pattern
    — the generic key/value sweep mis-reads its dimension tables."""
    key = "cbc_qt_run_text"
    label = "Chicago Blower Qt Run (text)"
    extensions = frozenset({".txt", ".rtf", ".docx", ".md"})
    name_patterns = (r"qt\s*run", r"quote\s*run")
    content_markers = SELECTION_PROGRAM_MARKERS

    def extract(self, ctx: QuoteRunContext) -> Dict[str, Any]:
        text = ctx.text()
        fields = _parse_chicago_blower(text)
        lines = [ln.rstrip() for ln in text.splitlines() if ln.strip()]
        # Keep the whole document (not just a peek): the sweep persists these
        # lines so new per-arrangement patterns can be designed/re-parsed from
        # the store without re-reading Z:.
        return {"fields": fields, "raw_lines": lines[:RAW_LINES_CAP],
                "summary": _cb_summary(fields)}


class QtRunText(_TextLineMixin, QuoteRunTemplate):
    """Any other `.txt`/`.rtf` run named like a quote run (non-Chicago-Blower).
    Best-effort key/value until a real dump of that format pins its headings."""
    key = "qt_run_text"
    label = "Qt Run (text, generic)"
    extensions = frozenset({".txt", ".rtf", ".docx", ".md"})
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
    """Any `.pdf` run. The same selection-program Qt Run is sometimes saved as a
    PDF, so if the extracted text carries the Qt Run header we parse it with the
    full Qt Run field set; otherwise we fall back to generic key/value (a vendor
    quote, a markup, etc.). Uses the position-aware extraction in `drive_run.py`."""
    key = "pdf"
    label = "Quote run (pdf)"
    extensions = frozenset({".pdf"})

    def extract(self, ctx: QuoteRunContext) -> Dict[str, Any]:
        from drive_run import parse_drive_run_pdf
        r = parse_drive_run_pdf(ctx.path)
        text = r.get("text", "")
        lines = [ln.rstrip() for ln in text.splitlines() if ln.strip()]
        if is_selection_program(text):
            fields = _parse_chicago_blower(text)
            if fields:
                return {"fields": fields, "raw_lines": lines[:RAW_LINES_CAP],
                        "summary": _cb_summary(fields)}
        return {"fields": r.get("fields", {}),
                "raw_lines": lines[:RAW_LINES_CAP] or r.get("raw_lines", [])}


class GenericTextRun(_TextLineMixin, QuoteRunTemplate):
    """Fallback for any text-shaped run that didn't match a named template."""
    key = "generic_text"
    label = "Quote run (text, generic)"
    extensions = frozenset({".txt", ".rtf", ".csv", ".docx", ".md", ""})

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
        # Debug, not info: a full-backlog sweep hits many of these (.doc/.msg),
        # and the inventory's Status column + end-of-run summary already report
        # them. Run with -v / DEBUG logging to see each one.
        log.debug("Quote run %s: no template reads %s files yet — add one to "
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
