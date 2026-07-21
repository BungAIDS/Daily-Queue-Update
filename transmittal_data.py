"""Gather everything a Drawing Transmittal needs from an order's Sales Order.

This is step 2 of *Completing Transmittals* — "fill out the Transmittal Word
document" — reduced to its data: who it goes to, which box to check, and which
drawing rows the table gets. It reads the **order face** (the archived Sales
Order PDF) plus the job's **AutoCAD folder**, and returns a `TransmittalData`
the doc-filler (`transmittal_doc.py`) and the Email-Drawings submitter
(`fill_transmittal_insider.py`) consume.

Where each field comes from (all confirmed against a real filled transmittal,
job 421693, and the SO regression dumps in test_line_items.py):

  emails      "Additional Features / Notes:" block, the "E-Mail Prints to:" line(s)
  P.O. #      the "Order # Rep Ref. # Customer P.O. #" header's data row
  customer    the board record (preferred) -> SO "Sold To:" fallback
  box         "record only / released for fabrication" when the SO says
              STATUS: APPROVED - RELEASED FOR PRODUCTION (or the board `flags`
              say so); otherwise "approval only"
  design #    the "Design <n> <desc>" header (e.g. Design 1904 PFD)
  table rows  a fixed recipe driven by the SO "Fan Drawings:" checklist +
              whether "Include 3D STEP Drawings" is a line item
  IMI no.     the O&M manual drawing number — a file in the AutoCAD folder

The parsing functions are pure (operate on the SO's reconstructed text lines) so
they're unit-tested directly; the I/O wrappers (open the PDF, find the folder)
sit at the bottom. Nothing here sends anything — see TRANSMITTAL_MODE.

    python transmittal_data.py 421693            # show what we'd put on the doc
    python transmittal_data.py --backtest 421693 # same, for validating against
                                                 # the real …TRANSMITTAL-01.doc
"""
from __future__ import annotations

import argparse
import json
import logging
import re
import sys
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Dict, List, Optional

from config import AUTOCAD_JOBS_DIR, SALES_ORDER_DIR, SNAPSHOT_DIR

log = logging.getLogger(__name__)

# --------------------------------------------------------------------------- #
# Patterns                                                                     #
# --------------------------------------------------------------------------- #
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
# "Design 16A SW" / "Design 1904 PFD" -> ("16A"/"1904", "SW"/"PFD").
DESIGN_HDR = re.compile(r"^\s*Design\s+(\S+)\s*(.*)$", re.IGNORECASE)
# The order/PO header row, then its data row carries the Customer P.O. #.
PO_HEADER_RE = re.compile(r"Order\s*#.*Customer\s*P\.?\s*O\.?\s*#", re.IGNORECASE)
ADDL_FEATURES_RE = re.compile(r"additional\s+features", re.IGNORECASE)
FAN_DRAWINGS_RE = re.compile(r"^\s*fan\s+drawings?\s*:", re.IGNORECASE)
EMAIL_PRINTS_RE = re.compile(r"e-?mail\s+prints?\s+to", re.IGNORECASE)
# "3D STEP" with hyphen/spacing variants ("3-D STEP", "3D-STEP", "3DSTEP"); \s
# also matches a newline so the phrase survives being split across two
# reconstructed lines when it's searched in joined text (see has_step).
STEP_RE = re.compile(r"\b3\s*-?\s*d[\s\-_]*step\b", re.IGNORECASE)
# "STATUS: APPROVED - RELEASED FOR PRODUCTION" (spacing/punct vary on real SOs).
# \b keeps "UNRELEASED" from matching; "fabrication" is the transmittal box's
# own wording and appears on some SOs.
RELEASED_RE = re.compile(r"\breleased\s+for\s+(?:production|fabrication)", re.IGNORECASE)
APPROVED_STATUS_RE = re.compile(r"status\s*:?\s*approved", re.IGNORECASE)
# Words that negate or condition a release/approved mention earlier on the SAME
# line: "will NOT be released for production", "held UNTIL released for
# production", "PRIOR TO release...". Checked against the text before the match.
NEGATION_RE = re.compile(r"\b(?:not|never|until|unless|prior\s+to|before|pending)\b", re.IGNORECASE)
# A page footer like "v1.8.1.5 -2-" that closes the Notes block.
PAGE_FOOTER_RE = re.compile(r"^\s*v\d+(\.\d+)+\s", re.IGNORECASE)

# The Fan-Drawings checklist rows we recognize, in transmittal order. Each label
# is matched at the start of a line; the trailing text is its Emailed/Mailed mark.
FAN_DRAWING_ROWS = [
    ("fan_drawings", re.compile(r"^\s*fan\s+drawings?\b", re.IGNORECASE)),
    ("om", re.compile(r"^\s*o\s*&\s*m\b", re.IGNORECASE)),
    ("motor_prints", re.compile(r"^\s*motor\s+prints?\b", re.IGNORECASE)),
    ("motor_data", re.compile(r"^\s*motor\s+data\s+sheets?\b", re.IGNORECASE)),
    ("buyout_prints", re.compile(r"^\s*buyout\s+prints?\b", re.IGNORECASE)),
    ("other", re.compile(r"^\s*other\b", re.IGNORECASE)),
]

CW_SUFFIX = "01"   # the standard CW main-assembly drawing suffix
CCW_SUFFIX = "02"  # CCW


# --------------------------------------------------------------------------- #
# Data model                                                                   #
# --------------------------------------------------------------------------- #
@dataclass
class DrawingRow:
    """One row of the transmittal's drawing table."""
    drawing_no: str
    description: str
    email: bool = False
    print: bool = False
    manual: bool = False
    rev: str = ""


@dataclass
class TransmittalData:
    order: str
    customer: str = ""
    po: str = ""
    emails: List[str] = field(default_factory=list)
    box: str = "approval"               # "sales" | "approval" | "record"
    released: bool = False              # True => "record only, released for fab"
    box_evidence: str = ""              # the SO line / flag the box choice came from
    design_no: str = ""
    design_desc: str = ""
    fan_checklist: Dict[str, str] = field(default_factory=dict)
    include_step: bool = False
    drawing_rows: List[DrawingRow] = field(default_factory=list)
    imi_number: str = ""
    so_pdf: Optional[str] = None
    folder: Optional[str] = None
    attachments: List[str] = field(default_factory=list)
    so_verified_at: Optional[str] = None   # when the watcher last read this order's SO
    so_read_today: bool = False            # was that read today?
    warnings: List[str] = field(default_factory=list)

    def warn(self, msg: str) -> None:
        self.warnings.append(msg)


# --------------------------------------------------------------------------- #
# Pure parsing (operate on the SO's reconstructed text lines)                  #
# --------------------------------------------------------------------------- #
def _notes_block(lines: List[str]) -> List[str]:
    """The lines of the 'Additional Features / Notes:' section to the end of the
    document (or the next page footer), where the email + checklist live."""
    start = next((i for i, ln in enumerate(lines) if ADDL_FEATURES_RE.search(ln)), None)
    if start is None:
        return []
    out: List[str] = []
    for ln in lines[start + 1:]:
        out.append(ln)
    return out


def parse_emails(lines: List[str]) -> List[str]:
    """Recipient emails from the Notes block — every address after the
    'E-Mail Prints to:' line(s), de-duplicated, order preserved.

    Restricting to the Notes block keeps internal addresses elsewhere on the SO
    out; if the explicit 'E-Mail Prints to:' line is present we anchor on it,
    otherwise we sweep the whole block as a fallback."""
    block = _notes_block(lines)
    if not block:
        return []
    anchored: List[str] = []
    capturing = False
    for ln in block:
        if EMAIL_PRINTS_RE.search(ln):
            capturing = True
        if capturing:
            anchored.extend(EMAIL_RE.findall(ln))
        # An email line can be followed by more bare-address lines; stop only at
        # the Fan Drawings checklist or a page footer.
        if capturing and (FAN_DRAWINGS_RE.search(ln) or PAGE_FOOTER_RE.search(ln)):
            break
    found = anchored or [m for ln in block for m in EMAIL_RE.findall(ln)]
    # De-duplicate case-insensitively, preserving first-seen casing/order.
    out, done = [], set()
    for e in found:
        k = e.lower()
        if k not in done:
            done.add(k)
            out.append(e)
    return out


def parse_po(lines: List[str]) -> str:
    """Customer P.O. # from the data row under the order/PO header.

    Header: 'Order # Rep Ref. # Customer P.O. # ...'; the next line carrying the
    order number holds the values. The P.O. is the last token of that row (rep
    ref sits between the order number and the P.O.)."""
    for i, ln in enumerate(lines):
        if PO_HEADER_RE.search(ln):
            for data in lines[i + 1:i + 4]:
                toks = data.split()
                if len(toks) >= 2 and toks[0].isdigit():
                    return toks[-1]
            break
    return ""


def _sold_to_only(t: str) -> str:
    """The Sold To half of a 'Sold To  Ship To' line that table extraction glued
    together. The two names sit side by side and usually start the same way, so
    we cut at the second occurrence of the first word ('JOHN ZINK COMPANY LLC
    JOHN ZINK COMPANY' -> 'JOHN ZINK COMPANY LLC'). Exact-halves and no-repeat
    lines fall through unchanged."""
    words = t.split()
    if len(words) >= 2:
        for p in range(1, len(words)):
            if words[p] == words[0]:
                return " ".join(words[:p])
    return t


def parse_customer(lines: List[str]) -> str:
    """Customer from the 'Sold To:' line that follows the 'Sold To: Ship To:'
    header. Sold To and Ship To sit side by side, so we keep only the Sold To
    name (see _sold_to_only)."""
    for i, ln in enumerate(lines):
        if re.match(r"\s*sold\s+to\s*:", ln, re.IGNORECASE):
            for nxt in lines[i + 1:i + 3]:
                t = nxt.strip()
                if not t:
                    continue
                return _sold_to_only(t)
            break
    return ""


def parse_design(lines: List[str]) -> tuple[str, str]:
    """('1904', 'PFD') from the 'Design <n> <desc>' header near the top."""
    for ln in lines[:12]:
        m = DESIGN_HDR.match(ln)
        if m:
            return m.group(1).strip(), m.group(2).strip()
    return "", ""


def _release_signal(ln: str) -> Optional[bool]:
    """One line's released-for-production signal: True for a clean positive
    ('STATUS: APPROVED', '... RELEASED FOR PRODUCTION'), False when the mention
    is negated/conditioned earlier on the line ('will NOT be released for
    production until ...'), None when the line says nothing about release."""
    m = RELEASED_RE.search(ln) or APPROVED_STATUS_RE.search(ln)
    if not m:
        return None
    return not NEGATION_RE.search(ln[: m.start()])


def parse_approval_evidence(lines: List[str], flags: str = "") -> tuple[str, bool, str]:
    """Which transmittal box to check, whether the order is released, and the SO
    line (or board flag) the decision came from.

    A clean 'STATUS: APPROVED / RELEASED FOR PRODUCTION' line (or the board
    `flags` saying so) means ('record', True) — negated mentions like
    'will not be released for production until approval' never count, so an
    unapproved order's boilerplate can't tick the released box. Everything else
    is ('approval', False); the 'sales purposes only' box is never auto-selected.
    Signals are evaluated per line, so a stray 'STATUS' at one line's end can't
    pair with an 'APPROVED' opening the next."""
    candidates = list(lines) + ([flags] if flags else [])
    negated = ""
    for ln in candidates:
        sig = _release_signal(ln)
        if sig:
            return "record", True, ln.strip()
        if sig is False and not negated:
            negated = ln.strip()
    if negated:
        return "approval", False, f"release mention is negated: {negated}"
    return "approval", False, "no APPROVED / RELEASED FOR PRODUCTION status found"


def parse_approval(lines: List[str], flags: str = "") -> tuple[str, bool]:
    """(box, released) — see parse_approval_evidence for the rules."""
    box, released, _ = parse_approval_evidence(lines, flags)
    return box, released


def parse_fan_drawings(lines: List[str]) -> Dict[str, str]:
    """The 'Fan Drawings:' send checklist -> {row_key: mark}, where mark is the
    trailing 'Both' / 'Emailed' / 'Mailed' / 'X' text (empty if unmarked)."""
    start = next((i for i, ln in enumerate(lines) if FAN_DRAWINGS_RE.search(ln)), None)
    if start is None:
        return {}
    checklist: Dict[str, str] = {}
    for ln in lines[start:]:
        if PAGE_FOOTER_RE.search(ln):
            break
        for key, rx in FAN_DRAWING_ROWS:
            m = rx.match(ln)
            if m:
                mark = ln[m.end():].strip()
                # Drop a trailing parenthetical (e.g. Buyout Prints "(e.g. ...)").
                mark = re.sub(r"^\(.*?\)\s*", "", mark).strip()
                checklist[key] = mark
                break
    return checklist


def _emailed(mark: str) -> bool:
    m = mark.lower()
    return ("email" in m) or ("both" in m) or (m.strip() == "x")


def _mailed(mark: str) -> bool:
    # Only a mailed-ONLY mark prints. On real transmittals a "Both" (or a bare
    # "X") still renders as EMAIL-only — see job 421693 — so "both" is not print.
    m = mark.lower()
    return "mail" in m and "email" not in m and "both" not in m


def has_step(lines: List[str]) -> bool:
    """True if a '... 3D STEP Drawings' line item is on the order. Searched in
    the joined text so the phrase still matches when the PDF reconstruction
    splits '3D' and 'STEP' onto separate lines, and tolerant of hyphen/spacing
    variants ('3-D STEP', '3D-STEP')."""
    return bool(STEP_RE.search("\n".join(lines)))


def apply_board_approval(d: "TransmittalData", board_unapproved_flag: Optional[bool]) -> None:
    """Let today's board state veto a stale/misread SO: when the queue board
    shows the order UNAPPROVED, the 'record only / released for fabrication'
    box must never be ticked, whatever the SO text seemed to say."""
    if board_unapproved_flag and d.released:
        d.warn("The queue board shows this order UNAPPROVED today, but the SO text "
               f"read as released ({d.box_evidence}) — checking 'approval only' "
               "instead. Verify the order's status before sending.")
        d.box, d.released = "approval", False
        d.box_evidence = f"board flag: UNAPPROVED today (overrode SO text: {d.box_evidence})"


def build_drawing_rows(
    order: str,
    design_no: str,
    design_desc: str,
    checklist: Dict[str, str],
    include_step: bool,
    imi_number: str = "",
    suffix: str = CW_SUFFIX,
) -> List[DrawingRow]:
    """Assemble the transmittal table from the SO, per the confirmed recipe:

      <order>-NN  DESIGN <n> <desc> FAN ASSEMBLY (AUTOCAD/PDF)   (always)
      IMI-…       FAN OPERATING AND MAINTENANCE MANUAL           (if O & M sent)
      <order>-NN  3D STEP DRAWING                                (if STEP on order)
      (blank)     MOTOR DOCUMENTS TO FOLLOW                      (if motor docs pending)
    """
    rows: List[DrawingRow] = []
    fan_mark = checklist.get("fan_drawings", "X")  # the assembly always goes
    desc_bits = " ".join(b for b in (design_no, design_desc) if b)
    rows.append(DrawingRow(
        drawing_no=f"{order}-{suffix}",
        description=f"DESIGN {desc_bits} FAN ASSEMBLY (AUTOCAD/PDF)".replace("  ", " ").strip(),
        email=_emailed(fan_mark) or True,   # assembly is emailed by default
        print=_mailed(fan_mark),
    ))
    om_mark = checklist.get("om", "")
    if om_mark:
        rows.append(DrawingRow(
            drawing_no=imi_number or "[IMI-?]",
            description="FAN OPERATING AND MAINTENANCE MANUAL",
            email=_emailed(om_mark) or True,
            print=_mailed(om_mark),
        ))
    if include_step:
        rows.append(DrawingRow(
            drawing_no=f"{order}-{suffix}",
            description="3D STEP DRAWING",
            email=True,
        ))
    if checklist.get("motor_prints") or checklist.get("motor_data"):
        rows.append(DrawingRow(drawing_no="", description="MOTOR DOCUMENTS TO FOLLOW"))
    return rows


def build_transmittal_data(
    lines: List[str],
    order: str,
    customer: str = "",
    flags: str = "",
    imi_number: str = "",
    suffix: str = CW_SUFFIX,
    so_is_current: bool = True,
) -> TransmittalData:
    """The pure core: SO text lines (+ a few board-known fields) -> TransmittalData.
    No file or network I/O, so this is what the tests exercise. The current SO
    text is the sole authority for whether a 3D STEP drawing was requested."""
    d = TransmittalData(order=str(order))
    d.customer = customer or parse_customer(lines)
    d.po = parse_po(lines)
    d.emails = parse_emails(lines)
    d.design_no, d.design_desc = parse_design(lines)
    d.box, d.released, d.box_evidence = parse_approval_evidence(lines, flags)
    d.fan_checklist = parse_fan_drawings(lines)
    d.include_step = has_step(lines) and so_is_current
    d.imi_number = imi_number
    d.drawing_rows = build_drawing_rows(
        order=d.order, design_no=d.design_no, design_desc=d.design_desc,
        checklist=d.fan_checklist, include_step=d.include_step,
        imi_number=imi_number, suffix=suffix,
    )

    if not d.emails:
        d.warn("No recipient emails found in the Additional Features / Notes block.")
    if not d.po:
        d.warn("No Customer P.O. # found on the order face.")
    if not d.design_no:
        d.warn("No Design number found — assembly row description will be incomplete.")
    if not d.fan_checklist:
        d.warn("No 'Fan Drawings:' checklist found — table built from defaults only.")
    if d.fan_checklist.get("om") and not imi_number:
        d.warn("O&M manual is to be sent but no IMI-… number was found in the AutoCAD folder.")
    return d


# --------------------------------------------------------------------------- #
# I/O wrappers                                                                 #
# --------------------------------------------------------------------------- #
def _recon_lines(page, x_tol: float = 1.5) -> List[str]:
    """Rebuild a page's text lines from word positions so spaces survive (plain
    extraction glues the Notes text together). Mirrors sales_orders._recon_lines,
    inlined so reading an archived PDF here doesn't pull in the browser stack."""
    words = page.extract_words(x_tolerance=x_tol, keep_blank_chars=False, use_text_flow=False)
    rows: Dict[int, list] = {}
    for w in words:
        rows.setdefault(round(w["top"]), []).append(w)
    out = []
    for top in sorted(rows):
        ws = sorted(rows[top], key=lambda w: w["x0"])
        out.append(" ".join(w["text"] for w in ws))
    return out


def load_so_lines(pdf_path: str | Path) -> List[str]:
    """Reconstructed text lines for an SO PDF, using the same word-position
    reconstruction as the daily run so glued-together Notes text survives."""
    lines: List[str] = []
    try:
        import pdfplumber
    except ImportError:
        log.warning("pdfplumber not installed; cannot read SO PDFs (pip install pdfplumber)")
        return lines
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            for page in pdf.pages:
                lines.extend(_recon_lines(page))
    except Exception as e:  # noqa: BLE001 - a bad PDF must not crash the caller
        log.warning("Could not read SO PDF %s: %s", pdf_path, e)
    return lines


def find_job_folder(job: str) -> Optional[Path]:
    """The job's AutoCAD folder under AUTOCAD_JOBS_DIR (<type>/<intermediate>/<job>)."""
    root = AUTOCAD_JOBS_DIR
    if not root.exists():
        return None
    matches = list(root.glob(f"*/*/{job}")) or list(root.glob(f"*/*/{job}*"))
    return matches[0] if matches else None


def find_imi_number(folder: Optional[Path]) -> str:
    """The O&M manual number — a file named like 'IMI-HD_A4' / 'IMI-GL-2021' in
    the job's AutoCAD folder. Returns the file stem, or '' if none is found."""
    if not folder or not folder.exists():
        return ""
    for f in sorted(folder.rglob("*")):
        if f.is_file() and re.match(r"IMI[-_]", f.name, re.IGNORECASE):
            return f.stem
    return ""


def so_last_verified(order: str, ref: Optional[date] = None) -> Optional[str]:
    """When the watcher last read (re-verified) this order's Sales Order, from the
    per-day live state file it writes (`verified_at`). Returns the ISO timestamp,
    or None if the order hasn't been verified in that day's state. Reads the JSON
    directly so this stays decoupled from the watcher's (browser) imports."""
    ref = ref or date.today()
    path = SNAPSHOT_DIR / f"live_state_{ref.isoformat()}.json"
    if not path.exists():
        return None
    try:
        state = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None
    entry = state.get(str(order)) if isinstance(state, dict) else None
    va = (entry or {}).get("verified_at") if isinstance(entry, dict) else None
    return va or None


def board_unapproved(order: str, ref: Optional[date] = None) -> Optional[bool]:
    """Today's board 'unapproved' flag for this order, from the watcher's
    per-day live state file. None when the order isn't in today's state (or the
    stored job dict doesn't carry the flag) — only a definite True/False from
    the board itself is returned. Reads the JSON directly, like so_last_verified,
    so this stays decoupled from the watcher's (browser) imports."""
    ref = ref or date.today()
    path = SNAPSHOT_DIR / f"live_state_{ref.isoformat()}.json"
    if not path.exists():
        return None
    try:
        state = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None
    entry = state.get(str(order)) if isinstance(state, dict) else None
    job = (entry or {}).get("job") if isinstance(entry, dict) else None
    if not isinstance(job, dict) or "unapproved" not in job:
        return None
    return bool(job.get("unapproved"))


def so_read_today(order: str, ref: Optional[date] = None) -> bool:
    """True if the watcher read this order's SO today (per the recorded
    verified_at). The transmittal should be built from a SO confirmed current
    today, so a False here is a warning to refresh the order first."""
    ref = ref or date.today()
    va = so_last_verified(order, ref)
    return bool(va) and va[:10] == ref.isoformat()


def find_so_pdf(job: str) -> Optional[Path]:
    """The latest archived Sales Order PDF for a job (SALES_ORDER_DIR/<job>/…).

    Picks the highest change-order revision when the filename carries one
    ('… CO#2.pdf'), else the most recently modified PDF."""
    folder = SALES_ORDER_DIR / str(job)
    if not folder.exists():
        return None
    pdfs = [p for p in folder.glob("*.pdf") if p.is_file()]
    if not pdfs:
        return None

    def _key(p: Path):
        m = re.search(r"CO\s*#?\s*(\d+)", p.stem, re.IGNORECASE)
        co = int(m.group(1)) if m else 0
        try:
            mtime = p.stat().st_mtime
        except OSError:
            mtime = 0.0
        return (co, mtime)

    return max(pdfs, key=_key)


def find_attachments(
    folder: Optional[Path],
    order: str,
    imi_number: str = "",
    *,
    drawing_suffix: str = CW_SUFFIX,
    include_step: bool = False,
) -> List[Path]:
    """Files approved for the transmittal from the AutoCAD folder.

    Only the selected standard assembly drawing (-01 CW or -02 CCW) is included,
    in DWG/PDF form, plus the requested O&M manual. STEP exports are considered
    only when the current Sales Order explicitly requested 3D STEP drawings; in
    that case loose STEP names are accepted in the folder or one level below it.
    """
    if not folder or not folder.exists():
        return []
    if drawing_suffix not in (CW_SUFFIX, CCW_SUFFIX):
        raise ValueError(f"unsupported transmittal drawing suffix: {drawing_suffix}")

    import autocad_scan

    out: List[Path] = []
    step_exts = {".step", ".stp"}
    for f in sorted(folder.glob("*")):
        if not f.is_file():
            continue
        drawing = autocad_scan.parse_drawing(f.name, str(order))
        if drawing and drawing[0] == drawing_suffix:
            out.append(f)
        elif imi_number and f.stem.lower() == imi_number.lower() and f.suffix.lower() == ".pdf":
            out.append(f)
    if include_step:
        order_in_name = re.compile(rf"(?<!\d){re.escape(str(order))}(?!\d)")
        for f in sorted(list(folder.glob("*")) + list(folder.glob("*/*"))):
            if (f.is_file() and f.suffix.lower() in step_exts
                    and order_in_name.search(f.stem) and f not in out):
                out.append(f)
    return out


def _refresh_so(order: str) -> bool:
    """Fetch this order's SO fresh online (reusing the watcher/daily-run fetch via
    sales_orders.refresh_order_so) so the transmittal reads the latest revision.
    Lazy-imported so transmittal_data still loads without the browser stack;
    returns True on success, False (with a logged reason) if it couldn't run —
    no saved session, the order isn't on the board, the network, etc."""
    try:
        import sales_orders  # lazy: pulls in Playwright only when we actually refresh
        sales_orders.refresh_order_so(order)
        return True
    except Exception as e:  # noqa: BLE001 - a failed refresh must not break preview
        log.warning("Could not refresh the SO for %s online: %s", order, e)
        return False


def gather(order: str, customer: str = "", flags: str = "",
           refresh_stale: bool = True,
           require_current_so_for_step: bool = True) -> TransmittalData:
    """Full pipeline for one order: locate the SO PDF + AutoCAD folder, read the
    IMI number and candidate attachments, and build the TransmittalData. Picks
    the assembly suffix (-01 CW / -02 CCW) by which drawing the folder actually
    has. Never sends anything.

    If the order's SO hasn't been re-read today (per the watcher's verified_at)
    and `refresh_stale` is set, it fetches a fresh SO first so the transmittal is
    built from the latest revision. Live preparation also withholds STEP unless
    that current read succeeded; historical backtests can explicitly opt out."""
    order = str(order)

    # Confirm today's SO; if stale, pull a fresh one before reading anything.
    refreshed = False
    if refresh_stale and not so_read_today(order):
        log.info("SO for %s not re-read today — fetching a fresh copy...", order)
        refreshed = _refresh_so(order)

    so_pdf = find_so_pdf(order)
    folder = find_job_folder(order)
    imi = find_imi_number(folder)

    suffix = CW_SUFFIX
    if folder and folder.exists():
        names = [f.name for f in folder.glob("*") if f.is_file()]
        import autocad_scan
        found = autocad_scan.scan_files(names, order)
        if CW_SUFFIX not in found and CCW_SUFFIX in found:
            suffix = CCW_SUFFIX

    lines = load_so_lines(so_pdf) if so_pdf else []
    current_so = so_read_today(order)
    step_source_allowed = current_so or not require_current_so_for_step
    d = build_transmittal_data(lines, order, customer=customer, flags=flags,
                               imi_number=imi, suffix=suffix,
                               so_is_current=step_source_allowed)
    if has_step(lines) and require_current_so_for_step and not current_so:
        d.warn("The archived SO mentions 3D STEP, but this order's SO was not "
               "successfully re-read today — the STEP row and file were withheld.")

    # Today's board state beats a stale/misread SO for the approval box.
    apply_board_approval(d, board_unapproved(order))

    d.so_pdf = str(so_pdf) if so_pdf else None
    d.folder = str(folder) if folder else None
    d.attachments = [str(p) for p in find_attachments(
        folder, order, imi, drawing_suffix=suffix, include_step=d.include_step)]
    if d.include_step and not any(a.lower().endswith((".stp", ".step")) for a in d.attachments):
        d.warn("3D STEP is on the order but no .stp/.step file was found in the "
               "AutoCAD folder — export/attach the STEP file manually.")

    # Re-check freshness: after a successful refresh this is now today's SO.
    d.so_verified_at = so_last_verified(order)
    d.so_read_today = so_read_today(order)
    if not d.so_read_today:
        when = d.so_verified_at or "not since the watcher last ran"
        detail = ("an auto-refresh was attempted but failed — check the saved session / "
                  "that the order is on the board") if refresh_stale else "auto-refresh was off"
        d.warn(f"This order's Sales Order has NOT been re-read today (last: {when}); "
               f"{detail}. Verify the order is current before sending.")

    if not so_pdf:
        d.warn(f"No archived Sales Order PDF found under {SALES_ORDER_DIR / order}.")
    if not folder:
        d.warn(f"No AutoCAD folder found for {order} under {AUTOCAD_JOBS_DIR}.")
    return d


# --------------------------------------------------------------------------- #
# CLI / back-test                                                              #
# --------------------------------------------------------------------------- #
def print_summary(d: TransmittalData) -> None:
    """Dump every decision the transmittal will be built from — also called by
    fill_transmittal_insider so the launcher's per-run log captures it."""
    print(f"\n=== Transmittal data for order {d.order} ===")
    print(f"  Customer : {d.customer}")
    print(f"  P.O. #   : {d.po}")
    print(f"  Emails   : {'; '.join(d.emails) or '(none)'}")
    print(f"  Box      : {d.box}  (released={d.released})")
    print(f"  Because  : {d.box_evidence or '(no status evidence recorded)'}")
    print(f"  3D STEP  : {'yes' if d.include_step else 'no'}")
    print(f"  Design   : {d.design_no} {d.design_desc}")
    print(f"  IMI #    : {d.imi_number or '(not found)'}")
    print(f"  Checklist: {d.fan_checklist or '(none)'}")
    print(f"  Drawing table:")
    for r in d.drawing_rows:
        flags = "".join(c if v else "·" for c, v in (("E", r.email), ("P", r.print), ("M", r.manual)))
        print(f"     [{flags}] {r.drawing_no or '(blank)':<14} {r.description}")
    if d.attachments:
        print(f"  Attachments ({len(d.attachments)}):")
        for a in d.attachments:
            print(f"     - {a}")
    print(f"  SO today : {'yes' if d.so_read_today else 'NO'}  (last verified: {d.so_verified_at or 'never'})")
    if d.so_pdf:
        print(f"  SO PDF   : {d.so_pdf}")
    if d.folder:
        print(f"  Folder   : {d.folder}")
    if d.warnings:
        print(f"  WARNINGS:")
        for w in d.warnings:
            print(f"     ! {w}")


def main(argv: Optional[List[str]] = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    ap = argparse.ArgumentParser(description="Show the data a transmittal would be built from.")
    ap.add_argument("orders", nargs="+", help="order number(s)")
    ap.add_argument("--backtest", action="store_true",
                    help="offline read of historical orders (no online SO refresh) for diffing "
                         "against the real …TRANSMITTAL-01.doc")
    ap.add_argument("--no-refresh", action="store_true",
                    help="don't fetch a fresh SO even if it wasn't re-read today")
    ap.add_argument("--customer", default="", help="override customer (else read from SO)")
    args = ap.parse_args(argv)

    # Back-test diffs history offline; otherwise refresh a stale SO before reading.
    refresh = not (args.backtest or args.no_refresh)
    for order in args.orders:
        d = gather(order, customer=args.customer, refresh_stale=refresh,
                   require_current_so_for_step=not args.backtest)
        print_summary(d)
    return 0


if __name__ == "__main__":
    sys.exit(main())
