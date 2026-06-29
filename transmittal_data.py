"""Gather everything a Drawing Transmittal needs from an order's Sales Order.

This is step 2 of *Completing Transmittals* — "fill out the Transmittal Word
document" — reduced to its data: who it goes to, which box to check, and which
drawing rows the table gets. It reads the **order face** (the archived Sales
Order PDF) plus the job's **AutoCAD folder**, and returns a `TransmittalData`
the doc-filler (`transmittal_doc.py`) and the Email-Drawings submitter
(`email_drawings.py`) consume.

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
import logging
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional

from config import AUTOCAD_JOBS_DIR, SALES_ORDER_DIR

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
STEP_RE = re.compile(r"3d\s+step", re.IGNORECASE)
# "STATUS: APPROVED - RELEASED FOR PRODUCTION" (spacing/punct vary on real SOs).
RELEASED_RE = re.compile(r"released\s+for\s+production", re.IGNORECASE)
APPROVED_STATUS_RE = re.compile(r"status\s*:?\s*approved", re.IGNORECASE)
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
    design_no: str = ""
    design_desc: str = ""
    fan_checklist: Dict[str, str] = field(default_factory=dict)
    include_step: bool = False
    drawing_rows: List[DrawingRow] = field(default_factory=list)
    imi_number: str = ""
    so_pdf: Optional[str] = None
    folder: Optional[str] = None
    attachments: List[str] = field(default_factory=list)
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


def parse_customer(lines: List[str]) -> str:
    """Customer from the 'Sold To:' line that follows the 'Sold To: Ship To:'
    header. Sold To and Ship To sit side by side and are usually identical, so we
    take the first half when the line is two equal halves, else the whole line."""
    for i, ln in enumerate(lines):
        if re.match(r"\s*sold\s+to\s*:", ln, re.IGNORECASE):
            for nxt in lines[i + 1:i + 3]:
                t = nxt.strip()
                if not t:
                    continue
                half = len(t) // 2
                left, right = t[:half].strip(), t[half:].strip()
                if left and left == right:
                    return left
                return t
            break
    return ""


def parse_design(lines: List[str]) -> tuple[str, str]:
    """('1904', 'PFD') from the 'Design <n> <desc>' header near the top."""
    for ln in lines[:12]:
        m = DESIGN_HDR.match(ln)
        if m:
            return m.group(1).strip(), m.group(2).strip()
    return "", ""


def parse_approval(lines: List[str], flags: str = "") -> tuple[str, bool]:
    """Which transmittal box to check, and whether the order is released.

    Returns ('record', True) when the SO is APPROVED / RELEASED FOR PRODUCTION
    (or the board `flags` say so), else ('approval', False). The 'sales purposes
    only' box is never auto-selected."""
    text = "\n".join(lines)
    flags = flags or ""
    released = bool(
        RELEASED_RE.search(text)
        or (APPROVED_STATUS_RE.search(text) and "approv" in text.lower())
        or RELEASED_RE.search(flags)
        or APPROVED_STATUS_RE.search(flags)
    )
    return ("record", True) if released else ("approval", False)


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
    m = mark.lower()
    return ("mail" in m and "email" not in m) or ("both" in m)


def has_step(lines: List[str]) -> bool:
    """True if a '... 3D STEP Drawings' line item is on the order."""
    return any(STEP_RE.search(ln) for ln in lines)


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
) -> TransmittalData:
    """The pure core: SO text lines (+ a few board-known fields) -> TransmittalData.
    No file or network I/O, so this is what the tests exercise."""
    d = TransmittalData(order=str(order))
    d.customer = customer or parse_customer(lines)
    d.po = parse_po(lines)
    d.emails = parse_emails(lines)
    d.design_no, d.design_desc = parse_design(lines)
    d.box, d.released = parse_approval(lines, flags)
    d.fan_checklist = parse_fan_drawings(lines)
    d.include_step = has_step(lines)
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


def find_attachments(folder: Optional[Path], order: str, imi_number: str = "") -> List[Path]:
    """Candidate files to attach from the AutoCAD folder: the order's drawings
    (DWG/PDF/STEP for any suffix) and the O&M manual PDF. The engineer confirms
    this list in preview mode before anything is attached."""
    if not folder or not folder.exists():
        return []
    out: List[Path] = []
    exts = {".dwg", ".pdf", ".step", ".stp"}
    for f in sorted(folder.glob("*")):
        if not f.is_file():
            continue
        if re.match(rf"{re.escape(str(order))}-\d+", f.stem, re.IGNORECASE) and f.suffix.lower() in exts:
            out.append(f)
        elif imi_number and f.stem.lower() == imi_number.lower() and f.suffix.lower() == ".pdf":
            out.append(f)
    return out


def gather(order: str, customer: str = "", flags: str = "") -> TransmittalData:
    """Full pipeline for one order: locate the SO PDF + AutoCAD folder, read the
    IMI number and candidate attachments, and build the TransmittalData. Picks
    the assembly suffix (-01 CW / -02 CCW) by which drawing the folder actually
    has. Never sends anything."""
    order = str(order)
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
    d = build_transmittal_data(lines, order, customer=customer, flags=flags,
                               imi_number=imi, suffix=suffix)
    d.so_pdf = str(so_pdf) if so_pdf else None
    d.folder = str(folder) if folder else None
    d.attachments = [str(p) for p in find_attachments(folder, order, imi)]

    if not so_pdf:
        d.warn(f"No archived Sales Order PDF found under {SALES_ORDER_DIR / order}.")
    if not folder:
        d.warn(f"No AutoCAD folder found for {order} under {AUTOCAD_JOBS_DIR}.")
    return d


# --------------------------------------------------------------------------- #
# CLI / back-test                                                              #
# --------------------------------------------------------------------------- #
def _print(d: TransmittalData) -> None:
    print(f"\n=== Transmittal data for order {d.order} ===")
    print(f"  Customer : {d.customer}")
    print(f"  P.O. #   : {d.po}")
    print(f"  Emails   : {'; '.join(d.emails) or '(none)'}")
    print(f"  Box      : {d.box}  (released={d.released})")
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
                    help="(same output; intended for diffing against the real …TRANSMITTAL-01.doc)")
    ap.add_argument("--customer", default="", help="override customer (else read from SO)")
    args = ap.parse_args(argv)

    for order in args.orders:
        d = gather(order, customer=args.customer)
        _print(d)
    return 0


if __name__ == "__main__":
    sys.exit(main())
