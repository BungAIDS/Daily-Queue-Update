"""Sales-order enrichment for the daily run.

For every job on the board this:
  1. opens its detail modal (in parallel across SO_CONCURRENCY tabs) and reads
     the CBC_SalesOrder revision  ->  CO# = rev - 1  (CO#1 = rev 2),
  2. downloads the latest Sales Order pdf into SALES_ORDER_DIR/<job>/ if that
     revision isn't already on disk (keeping older revisions),
  3. parses Design / Size / Arrangement + the change-order history out of the
     pdf, and
  4. looks up the job's AutoCAD folder, which also yields its type.

`enrich_with_sales_orders(jobs)` mutates each job dict in place, adding:
    co_number      int   (0 = no change orders)
    co_history     list[str]  (the "CO#N date initials - description" lines)
    so_design_desc str   (e.g. "Vaneaxial Belt Drive")
    so_size        str
    so_arrangement str
    so_pdf         str   (path to the latest SO pdf, or "")
    has_drive_run  bool  (True = a quote/construction run exists -> highly custom fan)
    drive_run_pdf  str   (path to the run file: archived download, or the file
                          in the job's AutoCAD folder; .pdf/.txt/.xlsx; or "")
    drive_run_count int  (how many files matched; >1 -> report shows "YES (X)"
                          so someone reviews which is the real run)
    drive_run      dict  (parsed quote-run fields, any format; see templates.py)
    drive_run_summary str (compact one-liner of the quote-run fields)
    drive_run_template str (which templates.py template read the run)
    job_type       str   (e.g. "AXIAL" / "GENERAL LINE", or "")
    job_folder     str   (AutoCAD folder if found, else the SO archive folder)

It is resilient: any job that errors, has no sales order (e.g. HDX), or whose
folder isn't found simply gets blank/zero fields rather than failing the run.
"""
from __future__ import annotations

import asyncio
import contextlib
import logging
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List
from urllib.parse import urlparse, parse_qs, urljoin

from playwright.async_api import async_playwright, TimeoutError as PWTimeout, Error as PWError

from config import (
    CBC_URL, CBC_QUEUE_URL, STORAGE_STATE_PATH,
    SALES_ORDER_DIR, DRIVE_RUN_DIR, DRIVE_RUN_TYPES, DRIVE_RUN_NAME_PATTERNS,
    SO_CONCURRENCY, AUTOCAD_JOBS_DIR,
)
from templates import parse_quote_run
from scraper import CONTAINER_SELECTOR
import autocad_scan
import line_items
from process_lock import cbc_fetch_lock
from sales_order_validation import (
    DOCUMENT_KIND_ORDER_VERIFICATION,
    DOCUMENT_KIND_SALES_ORDER,
    accept_existing,
    classify_sales_order_document,
    clear_sales_order_data,
    failed_acceptance,
    finalize_candidate,
    staging_path,
    sales_order_sha256,
)

log = logging.getLogger(__name__)

PID_RE = re.compile(r"^(?P<type>.+?)-(?P<id>\d+)-(?P<rev>\d+)-(?P<tag>[A-Za-z0-9]+)$")

# Documents are identified by the *type* prefix of their pid
# (CBC_SalesOrder-<id>-<rev>-<tag>). That prefix is the reliable key for the
# Sales Order. The quote/construction run only has its own pid type on HDX
# fans (config.DRIVE_RUN_TYPES); on everything else it's filed under a generic
# type like CBC_Inquiry and is recognizable only by its file name
# (config.DRIVE_RUN_NAME_PATTERNS).
SO_TYPE = "CBC_SalesOrder"
RUN_NAME_RES = [re.compile(p, re.I) for p in DRIVE_RUN_NAME_PATTERNS]
# A quote run is a *document*. Plenty of CAD files (HDX layouts) are named
# "QT RUN-..." (.dwg/.sldasm/.slddrw/.dwl2/.bak); those are drawings, not runs,
# so the folder scan only accepts document-like extensions. (Document matching
# by pid type is unaffected — that path doesn't go through here.)
RUN_DOC_EXTS = {".txt", ".pdf", ".rtf", ".xlsx", ".xlsm", ".xls", ".doc", ".docx", ".csv", ".md"}
CO_START = re.compile(r"^\s*C\s*/?\s*O\s*#?\s*\d", re.I)
CO_HISTORY_END = re.compile(
    r"^\s*(?:_{3,}|Design\s+Info\b|Order\s+Verification\s+Report\b|"
    r"Total\s+Billing\b|Freight\b|Page\s+\d+\b|CSIV\w*\b)",
    re.I,
)
DESIGN_HDR = re.compile(r"^\s*Design\s+(\S+)\s*(.*)$")
# Spec-row cells look like "Label value" (e.g. "Size M2", "WheelType BI").
SPEC_LABELS = {
    "design": "Design", "size": "Size", "arrangement": "Arrangement",
    "motorpos": "MotorPos", "class": "Class", "rotation": "Rotation",
    "discharge": "Discharge", "%width": "%Width", "wheeltype": "WheelType",
    "designtemp": "DesignTemp", "maxtemp": "MaxTemp",
}
SPEC_CELL = re.compile(
    r"^(DesignTemp|MaxTemp|Design|Size|Arrangement|MotorPos|Class|Rotation|Discharge|%Width|WheelType)\b\s*(.*)$",
    re.I,
)
# Some older true Sales Orders have no usable spec table. Their fan summary
# instead lives in two plain-text sections:
#
#   Design Info
#   D95 Backward Curved SW, SIZE 270, A/4, CW, TH, 44.5%, WHEEL TYPE Backward Curved
#   Performance
#   ..., DESIGN TEMP 95, MAX TEMP 95, ...
_LEGACY_DESIGN_INFO = re.compile(r"^\s*Design\s+Info\s*$", re.I)
_LEGACY_SECTION_STOP = re.compile(
    r"^(?:_+|Performance\b|Base\s+Fan\b|Inquiry\b|Motor\b|CSIV\w*\b|\S+@\S+)",
    re.I,
)
_LEGACY_ARRANGEMENT = re.compile(
    r"^((?:A/[A-Z0-9-]+(?:\s+[A-Z])?)|(?:Arrangement\s+[A-Z0-9/-]+))(?:\s|$)",
    re.I,
)
_LEGACY_WHEEL_CODES = {
    "airfoil": "AF",
    "backward curved": "BC",
    "backward inclined": "BI",
    "forward curved": "FC",
    "radial blade": "RB",
}
# Special temperature rating, written in the Base Fan line as "Suitable for
# <temp>" (e.g. "Suitable for -45C", "Suitable for -40°"). Distinct from the
# DesignTemp/MaxTemp airstream values and the BHP@ reference temp. Requires a
# degree symbol or C/F unit so it doesn't catch "Suitable for 3600 rpm Motor".
TEMP_RE = re.compile(r"suitable\s*for\s*(-?\d+\s*(?:°\s*[CF]?|[CF]))", re.I)


def _special_temp(design_temp: str, max_temp: str, suitable: str) -> str:
    """Headline temp: the high airstream temp if Design/Max > 150, else the
    low 'Suitable for' rating if present, else '0' (a standard-temp fan)."""
    def _num(s):
        m = re.search(r"-?\d+", s or "")
        return int(m.group()) if m else None
    highs = [t for t in (_num(design_temp), _num(max_temp)) if t is not None]
    if highs and max(highs) > 150:
        return str(max(highs))
    return suitable or "0"


# --------------------------------------------------------------------------- #
# AutoCAD folder / job-type lookup                                            #
# --------------------------------------------------------------------------- #
# A subfolder named "history" or "hist" (any case, e.g. "Qt History", "RUN HIST")
# holds superseded copies of the run — never the live one, so the sweep skips it.
_HISTORY_DIR = re.compile(r"(?i)\bhist(ory)?\b")


def _in_history_dir(f: Path, root: Path) -> bool:
    """True if any subfolder between *root* and *f* is a history/hist folder."""
    try:
        parts = f.relative_to(root).parts[:-1]  # directories only, drop the filename
    except ValueError:
        parts = f.parts[:-1]
    return any(_HISTORY_DIR.search(p) for p in parts)


def _run_files_in_folder(folder: Path) -> List[Path]:
    """Quote-run files in a job's AutoCAD folder, searched recursively — they
    are often tucked in a subfolder (e.g. ENG REF\\420410 qt  run.txt). Some
    orders never get the run attached to their cbcinsider documents at all,
    so the folder is the only place it lives. Superseded copies under a
    `history\\` / `hist\\` subfolder are skipped — they are not the live run."""
    try:
        return sorted(
            f for f in folder.rglob("*")
            if f.is_file() and not f.name.startswith("~$")   # Office lock/temp files
            and f.suffix.lower() in RUN_DOC_EXTS              # a document, not a CAD drawing
            and not _in_history_dir(f, folder)               # not a superseded history copy
            and _is_run_name(f.name)
        )
    except OSError as e:
        log.warning("  could not scan %s for quote-run files (%s)", folder, e)
        return []


def _archived_runs(job: str) -> List[Path]:
    """The quote-run files already in a job's Quote-Runs archive
    (DRIVE_RUN_DIR/<job>/) — exactly what the report's 'YES (X)' link opens."""
    d = DRIVE_RUN_DIR / job
    try:
        return sorted(f for f in d.iterdir()
                      if f.is_file() and not f.name.startswith("~$"))
    except OSError:
        return []


def _archive_folder_runs(job: str, autocad_folder: "Path | None") -> List[Path]:
    """Copy any quote-run files that live only in the job's AutoCAD folder into
    its Quote-Runs archive (DRIVE_RUN_DIR/<job>/), so every run the 'YES (X)'
    count includes sits in the one folder the link opens. A file already there
    (by name) is left as-is, and copy failures are logged but never fatal.
    Returns all run files in the archive afterward."""
    dest_dir = DRIVE_RUN_DIR / job
    if autocad_folder:
        for src in _run_files_in_folder(autocad_folder):
            dest = dest_dir / src.name
            if dest.exists():
                continue
            try:
                dest_dir.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dest)
            except OSError as e:  # noqa: BLE001 - a copy must not fail enrichment
                log.warning("  could not copy run %s into the archive (%s)", src.name, e)
    return _archived_runs(job)


def _find_autocad_folders(job_numbers: List[str], deep: bool = True) -> Dict[str, Dict[str, Any]]:
    """Locate each job's AutoCAD folder, which is AUTOCAD_JOBS_DIR/<type>/<first 3
    digits of job#>/<job#> (e.g. JOBS/AXIAL/421/421303, JOBS/GENERAL LINE/421/
    421034). Returns {job: {type, path, dwg_extras, dwg_missing_std}}; {} if the
    drive isn't reachable. The <type> a job sits under is its job type.

    We build each job's expected path directly and check it per type, rather than
    globbing the whole ~12K-folder tree — that's far faster and reliable on the
    network share (the old full sweep could time out before reaching some types,
    which is why axial fans came back as 'Open'). When `deep`, a glob fallback also
    covers any job that doesn't follow the <type>/<first3>/<job> convention; the
    per-poll re-check passes deep=False to skip that expensive sweep, since the
    direct lookup already finds a standard folder the moment it's created. While we
    have a folder we also scan it for the job's custom drawings."""
    out: Dict[str, Dict[str, Any]] = {}
    root = AUTOCAD_JOBS_DIR
    wanted = [str(j).strip() for j in job_numbers if str(j).strip()]
    try:
        if not root.exists():
            log.warning("AutoCAD jobs root not reachable: %s (folder links disabled)", root)
            return out
        type_dirs = [d for d in root.iterdir() if d.is_dir()]
    except OSError as e:
        log.warning("Could not list AutoCAD job types under %s (%s); folder links disabled", root, e)
        return out

    def _record(job: str, folder: Path, jtype: str) -> None:
        info: Dict[str, Any] = {"type": jtype, "path": folder,
                                "dwg_extras": {}, "dwg_missing_std": False, "imi": ""}
        try:  # live scan of this job's custom DWGs (names only — never opens a file)
            names = [f.name for f in folder.glob("*") if f.is_file()]
            rec = autocad_scan.build_record(job, jtype, str(folder),
                                            autocad_scan.scan_files(names, job))
            info["dwg_extras"], info["dwg_missing_std"] = rec["extras"], rec["missing_std"]
            # The O&M manual number for the transmittal (a file like IMI-HD_A4).
            info["imi"] = next((Path(n).stem for n in names
                                if re.match(r"IMI[-_]", n, re.I)), "")
        except OSError as e:
            log.warning("  could not scan DWGs for %s (%s)", job, e)
        out[job] = info

    # Direct lookup: <type>/<first 3 digits>/<job>, checked against each type.
    for job in wanted:
        if job in out:
            continue
        inter_name = job[:3]
        for td in type_dirs:
            inter = td / inter_name
            try:
                cand = inter / job
                if cand.is_dir():                       # exact "421303"
                    _record(job, cand, td.name)
                    break
                if inter.is_dir():                      # "421303 ACME CORP"
                    hits = [m for m in inter.glob(f"{job}*") if m.is_dir()]
                    if hits:
                        _record(job, hits[0], td.name)
                        break
            except OSError:
                continue

    # Fallback for any job whose folder doesn't follow <type>/<first3>/<job>:
    # one depth-3 sweep, matching the leaf either exactly (keeps a trailing
    # letter like 352366A) or by its leading number ("421034 ACME" -> 421034).
    # Skipped when deep=False (the per-poll re-check) — that sweep walks the whole
    # ~12K-folder tree and would run every cycle for jobs whose folder simply
    # doesn't exist yet.
    unfound = {j for j in wanted if j not in out}
    if deep and unfound:
        try:
            for m in root.glob("*/*/*"):
                if not m.is_dir():
                    continue
                name = m.name.strip()
                hit = name if name in unfound else (autocad_scan.job_key(name)
                                                    if autocad_scan.job_key(name) in unfound else None)
                if hit:
                    _record(hit, m, m.relative_to(root).parts[0])
                    unfound.discard(hit)
                    if not unfound:
                        break
        except OSError as e:
            log.warning("AutoCAD fallback sweep failed (%s)", e)

    log.info("Located %d/%d AutoCAD job folders under %s", len(out), len(wanted), root)
    if unfound:
        log.info("  no folder for: %s", ", ".join(sorted(unfound))[:300])
    return out


def refresh_autocad_folders(jobs: List[Dict[str, Any]]) -> int:
    """Re-run the (now cheap) AutoCAD folder lookup for `jobs`, updating each
    dict's job_type / job_folder / dwg_extras / dwg_missing_std in place. Used to
    fill in orders whose folder wasn't found by an earlier lookup (they showed
    'Open') — the watcher enriches an order only once, so without this a folder
    found by an improved lookup would never reach an already-known order.
    Returns how many got a folder this pass."""
    by_job = {j["job"]: j for j in jobs if j.get("job")}
    if not by_job:
        return 0
    # Direct lookup only (no full-tree sweep): cheap enough to run every poll, and
    # it still finds a standard <type>/<first3>/<job> folder the moment it exists.
    index = _find_autocad_folders(list(by_job.keys()), deep=False)
    n = 0
    for jn, j in by_job.items():
        info = index.get(jn)
        if not info:
            continue
        j["job_type"] = info["type"]
        j["job_folder"] = str(info["path"])
        j["dwg_extras"] = info.get("dwg_extras", {})
        j["dwg_missing_std"] = info.get("dwg_missing_std", False)
        j["so_imi"] = info.get("imi", "")
        n += 1
    return n


# Matches both the current '… CO2.pdf' naming and the legacy '… CO#2.pdf' one
# (the '#' had to go — Excel hyperlinks can't hold a '#' in a file path).
_CO_IN_SO_NAME = re.compile(r"CO#?(\d+)", re.I)


def _latest_so_in_folder(folder: Path, expected_job: str):
    """Latest verified true Sales Order in a job folder."""
    best = None
    try:
        for p in folder.glob("*.pdf"):
            if "sales order" not in p.name.lower():
                continue
            accepted = accept_existing(p, expected_job, DOCUMENT_KIND_SALES_ORDER)
            if not accepted or not accepted.path:
                if accepted:
                    log.warning(
                        "Quarantined unverified Sales Order for %s: internal=%s status=%s -> %s",
                        expected_job,
                        accepted.validation.internal_order or "?",
                        accepted.validation.status,
                        accepted.quarantine_path,
                    )
                continue
            m = _CO_IN_SO_NAME.search(p.name)
            co = int(m.group(1)) if m else 0
            key = (co, p.stat().st_mtime)
            if best is None or key > best[0]:
                best = (key, p, co)
    except OSError:
        return None
    return (best[1], best[2]) if best else None


def refresh_sales_orders(jobs: List[Dict[str, Any]]) -> int:
    """Repoint an order's so_pdf at the latest Sales Order PDF actually on disk
    whenever the stored one has gone — a change order renames the file
    ('… (original).pdf' -> '… CO1.pdf'), which dead-links a path captured before
    the CO. Syncs co_number to the file found, so the change order also shows up
    (red text, change log). Cheap: a single stat per order, and a folder listing
    only for the ones whose link is broken. Mutates the job dicts in place;
    returns how many were repointed."""
    n = 0
    for j in jobs:
        jn = str(j.get("job") or "").strip()
        if not jn:
            continue
        cur = (j.get("so_pdf") or "").strip()
        if cur and Path(cur).exists():
            accepted = accept_existing(cur, jn, DOCUMENT_KIND_SALES_ORDER)
            if accepted and accepted.path:
                continue                  # link still resolves and its printed Order# matches
            if accepted:
                log.warning(
                    "Quarantined stored Sales Order link for %s: internal=%s status=%s -> %s",
                    jn,
                    accepted.validation.internal_order or "?",
                    accepted.validation.status,
                    accepted.quarantine_path,
                )
                if (
                    accepted.validation.document_kind
                    == DOCUMENT_KIND_ORDER_VERIFICATION
                ):
                    clear_sales_order_data(
                        j, datetime.now().isoformat(timespec="seconds")
                    )
        folder = Path(cur).parent if cur else (SALES_ORDER_DIR / jn)
        hit = _latest_so_in_folder(folder, jn)
        if not hit:
            continue
        path, co = hit
        # Never regress: if the only Sales Order on disk is older than the change
        # order we already know about (e.g. the new CO's PDF hasn't downloaded
        # yet), don't point the link back at the original — keep the known CO#.
        if co < int(j.get("co_number") or 0):
            continue
        j["so_pdf"] = str(path)
        j["co_number"] = co
        n += 1
    return n


# --------------------------------------------------------------------------- #
# PDF parsing                                                                 #
# --------------------------------------------------------------------------- #
def _recon_line_records(page, page_number: int,
                        x_tol: float = 1.5) -> List[Dict[str, Any]]:
    """Rebuild text lines from word positions so spaces survive (plain
    extraction glues the Notes text together), retaining their PDF source."""
    words = page.extract_words(x_tolerance=x_tol, keep_blank_chars=False, use_text_flow=False)
    rows: Dict[int, list] = {}
    for w in words:
        rows.setdefault(round(w["top"]), []).append(w)
    out: List[Dict[str, Any]] = []
    for row_number, top in enumerate(sorted(rows), start=1):
        ws = sorted(rows[top], key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in ws)
        out.append({
            "text": text,
            "source": {
                "page": page_number,
                "row": row_number,
                "top": top,
                "source_type": "pdf-row",
                "source_text": text,
            },
        })
    return out


def _recon_lines(page, x_tol: float = 1.5) -> List[str]:
    return [record["text"] for record in _recon_line_records(page, 1, x_tol)]


def _co_history_from_lines(lines: List[str]) -> List[str]:
    """Collect CO notes, joining descriptions that wrap onto later PDF lines."""
    history: List[str] = []
    current: List[str] = []
    for raw in lines:
        line = str(raw or "").strip()
        if CO_START.match(line):
            if current:
                history.append(" ".join(current))
            current = [line]
            continue
        if not current or not line:
            continue
        if CO_HISTORY_END.match(line):
            history.append(" ".join(current))
            current = []
            continue
        current.append(line)
    if current:
        history.append(" ".join(current))
    return history


def _respace_value(value: str, recon_text: str) -> str:
    """Re-insert spaces that table extraction glued out of a value (e.g.
    'Flangemount' -> 'Flange mount'), using the page's word-position
    reconstruction which recovers the small inter-word gaps. Each token is
    looked up in the despaced recon and replaced with its properly-spaced span;
    anything not found is left exactly as-is, so this never loses content."""
    if not value:
        return value
    chars, idxmap = [], []
    for i, ch in enumerate(recon_text):
        if not ch.isspace():
            chars.append(ch)
            idxmap.append(i)
    blob = "".join(chars)
    out = []
    for token in value.split():
        pos = blob.find(token)
        if pos < 0:
            out.append(token)
            continue
        s, e = idxmap[pos], idxmap[pos + len(token) - 1] + 1
        out.append(re.sub(r"\s+", " ", recon_text[s:e]).strip())
    return " ".join(out)


def _spec_from_tables(tables) -> Dict[str, str]:
    fields: Dict[str, str] = {}
    for table in tables or []:
        for row in table:
            for cell in row:
                if not cell:
                    continue
                m = SPEC_CELL.match(cell.replace("\n", " ").strip())
                if m:
                    label, val = SPEC_LABELS[m.group(1).lower()], m.group(2).strip()
                    if label not in fields and val:  # first wins (vaneaxial repeats "Design")
                        fields[label] = val
    return fields


def _legacy_spec_from_text(lines: List[str]) -> Dict[str, str]:
    """Parse the fan summary from an Infor ``Order Verification Report``.

    These PDFs have no extractable spec table, but the same values are present
    in stable comma-delimited Design Info and Performance lines. Return keys in
    the same shape as :func:`parse_sales_order_pdf`; an unrecognized block is a
    harmless empty dict.
    """
    block: List[str] = []
    for index, line in enumerate(lines):
        if not _LEGACY_DESIGN_INFO.match(line or ""):
            continue
        for raw in lines[index + 1:]:
            value = re.sub(r"\s+", " ", raw or "").strip()
            if not value:
                continue
            if _LEGACY_SECTION_STOP.match(value):
                break
            block.append(value)
        break
    if not block:
        return {}

    text = " ".join(block)
    design_size = re.split(r",\s*SIZE\s+", text, maxsplit=1, flags=re.I)
    if len(design_size) != 2:
        return {}
    raw_design, rest = (part.strip() for part in design_size)
    wheel_parts = re.split(r",\s*WHEEL\s+TYPE\s*", rest, maxsplit=1, flags=re.I)
    before_wheel = wheel_parts[0]
    raw_wheel = wheel_parts[1] if len(wheel_parts) == 2 else ""

    parts = [part.strip() for part in before_wheel.split(",") if part.strip()]
    if not parts:
        return {}
    size = parts[0]
    rotation_index = next(
        (i for i, part in enumerate(parts[1:], start=1)
         if part.upper() in {"CW", "CCW"}),
        None,
    )
    if rotation_index is None:
        return {}

    pre_rotation = parts[1:rotation_index]
    arrangement = ""
    fan_class = ""
    for part in pre_rotation:
        arrangement_match = _LEGACY_ARRANGEMENT.match(part)
        if arrangement_match and not arrangement:
            arrangement = arrangement_match.group(1)
        if re.fullmatch(r"C/[A-Z0-9-]+", part, re.I) and not fan_class:
            fan_class = part

    design_match = re.match(
        r"^(?:Design\s+[0-9A-Z-]+|D[0-9A-Z-]+)\s+(.+)$",
        raw_design,
        re.I,
    )
    design_desc = design_match.group(1).strip() if design_match else raw_design
    wheel = re.sub(r"\s+", " ", raw_wheel).strip(" ,")
    wheel = _LEGACY_WHEEL_CODES.get(wheel.casefold(), wheel)
    discharge = parts[rotation_index + 1] if len(parts) > rotation_index + 1 else ""
    pct_width = parts[rotation_index + 2] if len(parts) > rotation_index + 2 else ""
    pct_width = pct_width.rstrip("%").strip()

    all_text = "\n".join(lines)
    # Horizontal whitespace only: ``\s`` could cross a newline and steal the C
    # from the following ``CSIV10C`` page footer as a temperature unit.
    temp_value = r"(-?\d+(?:\.\d+)?(?:[ \t]*(?:°[ \t]*)?[CF]\b)?)"
    design_temp = re.search(r"\bDESIGN\s+TEMP\s+" + temp_value, all_text, re.I)
    max_temp = re.search(r"\bMAX\s+TEMP\s+" + temp_value, all_text, re.I)
    return {
        "design_desc": design_desc,
        "size": size,
        "arrangement": arrangement or "N/A",
        "motor_pos": "N/A",
        "fan_class": fan_class or "N/A",
        "rotation": parts[rotation_index].upper(),
        "discharge": discharge,
        "pct_width": pct_width,
        "wheel_type": wheel,
        "design_temp": design_temp.group(1).strip() if design_temp else "",
        "max_temp": max_temp.group(1).strip() if max_temp else "",
    }


def parse_sales_order_pdf(path: str | Path) -> Dict[str, Any]:
    """Pull Design/Size/Arrangement + change-order history out of an SO pdf."""
    res = {"design_desc": "", "size": "", "arrangement": "", "motor_pos": "", "fan_class": "",
           "rotation": "", "discharge": "", "pct_width": "", "wheel_type": "", "temp": "",
           "design_temp": "", "max_temp": "", "special_temp": "0",
           "header_co": None, "co_history": [], "line_items": [], "parts_only": False,
           "job_number": "",
           # Transmittal fields (see transmittal_data.py): who the drawings go to,
           # the customer P.O. #, and whether the order is released for production.
           "emails": [], "customer_po": "", "released": False}
    try:
        import pdfplumber
    except ImportError:
        log.warning("pdfplumber not installed; cannot parse SO pdfs (pip install pdfplumber)")
        return res
    try:
        with pdfplumber.open(str(path)) as pdf:
            page_texts = [(page.extract_text() or "") for page in pdf.pages]
            if (
                classify_sales_order_document(page_texts[0])
                == DOCUMENT_KIND_ORDER_VERIFICATION
            ):
                log.warning(
                    "Refusing to parse Order Verification Report as a Sales Order: %s",
                    path,
                )
                return res
            page_tables = [page.extract_tables() for page in pdf.pages]
            page_recon_records = [
                _recon_line_records(page, page_number)
                for page_number, page in enumerate(pdf.pages, start=1)
            ]
            page_recon = [[record["text"] for record in records]
                          for records in page_recon_records]
            for ln in page_texts[0].splitlines()[:8]:
                if res["header_co"] is None:
                    m = re.search(r"CO\s*#\s*(\d+)", ln)
                    if m:
                        res["header_co"] = int(m.group(1))
                d = DESIGN_HDR.match(ln)
                if d and not res["design_desc"]:
                    res["design_desc"] = d.group(2).strip()
            # The Qty/Design/Size/Arrangement spec row is normally on page 1, but
            # a long Tag (nameplate) section can push it onto a later page — so
            # scan every page until the spec row turns up.
            for tables, recon_lines in zip(page_tables, page_recon):
                spec = _spec_from_tables(tables)
                if spec.get("Size") or spec.get("Arrangement"):
                    recon = "\n".join(recon_lines)
                    res["size"] = _respace_value(spec.get("Size", ""), recon)
                    res["arrangement"] = _respace_value(spec.get("Arrangement", "") or "N/A", recon)
                    # These six are short codes (DB, CCW, BI, 100, …) — take them
                    # verbatim; re-spacing would wrongly split e.g. "DB" -> "D B".
                    res["motor_pos"] = spec.get("MotorPos", "")
                    res["fan_class"] = spec.get("Class", "")
                    res["rotation"] = spec.get("Rotation", "")
                    res["discharge"] = spec.get("Discharge", "")
                    res["pct_width"] = spec.get("%Width", "")
                    res["wheel_type"] = spec.get("WheelType", "")
                    res["design_temp"] = spec.get("DesignTemp", "")
                    res["max_temp"] = spec.get("MaxTemp", "")
                    break
            # Some older true Sales Orders have no extractable spec table. Fill
            # missing summary values from their Design Info / Performance text.
            plain_lines = [line for text in page_texts for line in text.splitlines()]
            for key, value in _legacy_spec_from_text(plain_lines).items():
                if not res.get(key) and value:
                    res[key] = value
            recon_all_records: List[Dict[str, Any]] = []
            document_facts: Dict[str, Dict[str, Any]] = {}
            for page_number, (tables, recon_lines, recon_records) in enumerate(
                    zip(page_tables, page_recon, page_recon_records), start=1):
                for fact in line_items.document_fact_items_from_tables(
                        tables, recon_lines, page_number=page_number):
                    document_facts.setdefault(str(fact.get("document_fact") or ""), fact)
                recon_all_records.extend(
                    line_items.strip_continuation_metadata(recon_records, tables)
                )
            recon_all = [record["text"] for record in recon_all_records]
            res["co_history"] = _co_history_from_lines(recon_all)
            item_context = line_items.order_context_from_lines(
                recon_all, arrangement=res["arrangement"]
            )
            res["parts_only"] = item_context["parts_only"]
            res["job_number"] = item_context["job_number"]
            # Every line item on the order — the priced item/accessory rows and
            # the "Additional Features"-style lines — normalized + tagged so
            # orders can be looked up by what's on them (see line_items.py).
            res["line_items"] = line_items.extract_items(
                recon_all_records,
                order_context=item_context,
            ) + list(document_facts.values())
            # Transmittal fields off the same reconstructed text: recipient emails
            # (Additional Features / Notes -> "E-Mail Prints to:"), the customer
            # P.O. #, and the released-for-production status.
            import transmittal_data as _td
            res["emails"] = _td.parse_emails(recon_all)
            res["customer_po"] = _td.parse_po(recon_all)
            res["released"] = _td.parse_approval(recon_all)[1]
            # Special temperature rating from the "Suitable for <temp>" phrase.
            raw_all = "\n".join(page_texts)
            mt = TEMP_RE.search(raw_all)
            if mt:
                res["temp"] = re.sub(r"\s+", "", mt.group(1))
            res["special_temp"] = _special_temp(res["design_temp"], res["max_temp"], res["temp"])
    except Exception as e:  # noqa: BLE001 - never let a bad pdf fail the run
        log.warning("Could not parse SO pdf %s: %s", path, e)
    return res


_SO_SUMMARY_FIELDS = {
    "so_design_desc": "design_desc",
    "so_size": "size",
    "so_arrangement": "arrangement",
    "so_motor_pos": "motor_pos",
    "so_class": "fan_class",
    "so_rotation": "rotation",
    "so_discharge": "discharge",
    "so_pct_width": "pct_width",
    "so_wheel_type": "wheel_type",
    "so_design_temp": "design_temp",
    "so_max_temp": "max_temp",
    "so_special_temp": "special_temp",
}
_SO_SUMMARY_CORE = ("so_design_desc", "so_size", "so_arrangement")


def repair_missing_sales_order_summaries(jobs: List[Dict[str, Any]]) -> int:
    """Fill missing summary fields from already-archived Sales Order PDFs.

    Source summary fields are intentionally fill-only.  The derived Special
    Temp is recomputed while repairing a legacy summary, and an obvious high
    temperature can also be corrected from already-stored Design/Max Temp
    fields.  Returns the number of job dicts repaired.
    """
    repaired = 0
    for job in jobs:
        core_missing = not all(job.get(field) for field in _SO_SUMMARY_CORE)
        high_temp = _special_temp(
            str(job.get("so_design_temp") or ""),
            str(job.get("so_max_temp") or ""),
            "",
        )
        high_temp_mismatch = (
            high_temp != "0"
            and str(job.get("so_special_temp") or "") != high_temp
        )
        if not core_missing and not high_temp_mismatch:
            continue

        changed = False
        if high_temp_mismatch:
            job["so_special_temp"] = high_temp
            changed = True

        pdf = str(job.get("so_pdf") or "").strip()
        if core_missing and pdf:
            accepted = accept_existing(
                pdf,
                str(job.get("job") or ""),
                DOCUMENT_KIND_SALES_ORDER,
            )
            if not accepted or not accepted.path:
                if (
                    accepted
                    and accepted.validation.document_kind
                    == DOCUMENT_KIND_ORDER_VERIFICATION
                ):
                    clear_sales_order_data(
                        job, datetime.now().isoformat(timespec="seconds")
                    )
                continue
            missing = [field for field in _SO_SUMMARY_FIELDS if not job.get(field)]
            parsed = parse_sales_order_pdf(accepted.path)
            for field in missing:
                value = parsed.get(_SO_SUMMARY_FIELDS[field])
                if value not in (None, "", [], {}):
                    job[field] = value
                    changed = True

            parsed_special_temp = parsed.get("special_temp")
            stored_high_temp = _special_temp(
                str(job.get("so_design_temp") or ""),
                str(job.get("so_max_temp") or ""),
                "",
            )
            if stored_high_temp != "0":
                parsed_special_temp = stored_high_temp
            current_special_temp = str(job.get("so_special_temp") or "")
            if (
                parsed_special_temp not in (None, "")
                and not (
                    str(parsed_special_temp) == "0"
                    and current_special_temp not in ("", "0")
                )
                and current_special_temp != str(parsed_special_temp)
            ):
                job["so_special_temp"] = str(parsed_special_temp)
                changed = True
        repaired += int(changed)
    return repaired


# --------------------------------------------------------------------------- #
# Parallel fetch of each job's sales order                                    #
# --------------------------------------------------------------------------- #
_STATIC = (".js", ".css", ".png", ".gif", ".jpg", ".jpeg", ".svg", ".woff", ".woff2", ".ico")


def _jobnum(args_js: str) -> str:
    return args_js.split(",", 1)[0].strip().strip("'\"").split("-", 1)[0]


def _parse_doc(href: str) -> Dict[str, Any]:
    q = parse_qs(urlparse(href).query)
    pid, fn = q.get("pid", [""])[0], q.get("fn", [""])[0]
    m = PID_RE.match(pid)
    return {"fn": fn, "type": m["type"] if m else pid, "rev": int(m["rev"]) if m else None}


def _norm_type(t: str | None) -> str:
    """Normalize a pid type for comparison: lowercase, drop site prefixes."""
    t = (t or "").lower()
    for prefix in ("cbc_", "cs_"):
        if t.startswith(prefix):
            return t[len(prefix):]
    return t


def _latest_of_type(docs: List, type_name: str):
    """Highest revision for a pid type.

    Sales Orders are deliberately exact-only: ``CS_SalesOrder`` is an Order
    Verification Report and must never stand in for ``CBC_SalesOrder``.
    Other document families retain the site's prefix-normalized matching.
    """
    want = _norm_type(type_name)
    matches = [hd for hd in docs if _norm_type(hd[1].get("type")) == want]
    exact = [
        hd for hd in matches
        if str(hd[1].get("type") or "").casefold() == str(type_name or "").casefold()
    ]
    candidates = exact if str(type_name or "").casefold() == SO_TYPE.casefold() else (exact or matches)
    return max(candidates, key=lambda hd: hd[1].get("rev") or 0) if candidates else None


def _required_so_document_kind(doc: Dict[str, Any]) -> str | None:
    """Every document entering the Sales Order bank must be a true SO."""
    return DOCUMENT_KIND_SALES_ORDER


def _co_number_for_so_doc(doc: Dict[str, Any], parsed: Dict[str, Any] | None = None) -> int:
    """Return a trustworthy CO number for the selected Sales Order document."""
    if str(doc.get("type") or "").casefold() == SO_TYPE.casefold():
        rev = doc.get("rev")
        return (int(rev) - 1) if rev and int(rev) > 1 else 0
    return 0


def _is_run_name(fn: str) -> bool:
    """True if a document/file name looks like a quote run (DRIVE_RUN_NAME_PATTERNS)."""
    return any(rx.search(fn or "") for rx in RUN_NAME_RES)


def _run_docs(docs: List) -> List:
    """Every quote/construction-run document in `docs`, best match first.

    A doc qualifies by pid type — DRIVE_RUN_TYPES, or any other non-SO type
    ending in "run" (the HDX fans have a dedicated run pid) — or by file name
    (DRIVE_RUN_NAME_PATTERNS; most fans file the run under a generic type like
    CBC_Inquiry as "<job> ... Qt Run.txt", "... D64 Wheel Construction ...").
    Type matches sort ahead of name matches, higher revisions first within."""
    known = {_norm_type(t) for t in DRIVE_RUN_TYPES}
    matches = []
    for hd in docs:
        t = _norm_type(hd[1].get("type"))
        by_type = t != _norm_type(SO_TYPE) and (t in known or t.endswith("run"))
        if by_type and t not in known:
            log.warning("Run document matched by pid-type fallback (%r) — add it to "
                        "DRIVE_RUN_TYPES in .env to make this explicit.", hd[1].get("type"))
        if by_type or _is_run_name(hd[1].get("fn")):
            matches.append((by_type, hd))
    matches.sort(key=lambda m: (0 if m[0] else 1, -(m[1][1].get("rev") or 0)))
    return [hd for _, hd in matches]


def _so_filename(job: str, rev: int | None) -> str:
    # No '#' in the change-order suffix: Excel treats everything after a '#' in a
    # hyperlink address as an in-document anchor, so a Job # linked to a path like
    # "... CO#1.pdf" fails with "Cannot open the specified file". Legacy '#'
    # archives are still recognized (_CO_IN_SO_NAME), but new downloads use CO<n>.
    if rev and rev > 1:
        return f"{job} - Sales Order CO{rev - 1}.pdf"
    return f"{job} - Sales Order (original).pdf"


def _doc_filename(job: str, label: str, rev: int | None) -> str:
    """Archive filename for a non-SO document (e.g. the drive run)."""
    return f"{job} - {label} rev {rev}.pdf" if rev and rev > 1 else f"{job} - {label}.pdf"


def _run_filename(job: str, doc: Dict[str, Any]) -> str:
    """Archive name for a quote-run document. Keeps the site's own file name —
    it carries the identifying naming and the real extension (.txt qt runs,
    .xlsx D64 wheel constructions, .pdf HDX runs) — prefixed with the job
    number when it isn't already in it. '#' is sanitized alongside the
    Windows-invalid characters: Excel hyperlinks can't hold a '#' in a path."""
    fn = re.sub(r'[<>:"/\\|?*#]', "_", (doc.get("fn") or "").strip())
    if not fn:
        return _doc_filename(job, "Quote Run", doc.get("rev"))
    return fn if job in fn else f"{job} - {fn}"


def _download_error(status: int, body: bytes, expect_pdf: bool = True) -> str | None:
    """Why a downloaded document isn't usable, or None if it looks fine.

    The doc server can return an error page (HTTP 5xx) or — once the session
    expires — the login page itself, with HTTP 200. Writing either to disk
    would poison the archive: the dest.exists() check skips re-downloading
    forever, so the bad file would permanently stand in for the order's PDF.
    Quote runs aren't always PDFs (.txt, .xlsx, .rtf), so for those we only
    reject what is recognizably an HTML page.
    """
    if status != 200:
        return f"HTTP {status}"
    head = body[:1024].lstrip()
    if expect_pdf and not head.startswith(b"%PDF-"):
        return "response is not a PDF (expired-session login page or error page?)"
    if not expect_pdf and head[:15].lower().startswith((b"<!doctype", b"<html")):
        return "response is an HTML page (expired-session login page or error page?)"
    return None


async def _download(context, page_url: str, href: str, dest: Path) -> str | None:
    """Download a document to `dest` (skipping if present), retrying transient
    doc-server timeouts. Returns the path on success, else None."""
    if dest.exists():
        return str(dest)
    url = urljoin(page_url, href)
    for attempt in (1, 2, 3):
        try:
            resp = await context.request.get(url, timeout=60000)
            body = await resp.body()
            err = _download_error(resp.status, body, dest.suffix.lower() == ".pdf")
            if err:
                raise RuntimeError(err)
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_bytes(body)
            return str(dest)
        except Exception as e:  # noqa: BLE001
            if attempt == 3:
                log.warning("download failed for %s after %d tries: %s", dest.name, attempt, e)
            else:
                await asyncio.sleep(2 * attempt)
    return None


async def _download_sales_order(
    context,
    page_url: str,
    href: str,
    destination: Path,
    expected_job: str,
    required_document_kind: str | None = None,
):
    existing = accept_existing(destination, expected_job, required_document_kind)
    if existing and existing.path:
        return existing
    if existing:
        log.warning(
            "Rejected existing Sales Order for %s: internal=%s status=%s -> %s",
            expected_job,
            existing.validation.internal_order or "?",
            existing.validation.status,
            existing.quarantine_path,
        )

    staged = staging_path(destination, expected_job)
    downloaded = await _download(context, page_url, href, staged)
    if not downloaded:
        return failed_acceptance(expected_job, f"download failed for {destination.name}")
    accepted = finalize_candidate(
        downloaded, destination, expected_job, required_document_kind
    )
    if not accepted.path:
        log.warning(
            "Rejected downloaded Sales Order for %s: internal=%s status=%s -> %s",
            expected_job,
            accepted.validation.internal_order or "?",
            accepted.validation.status,
            accepted.quarantine_path,
        )
    return accepted


def _trigger_js(args_js: str) -> str:
    return f"""() => {{
        if (window.jQuery) {{
            jQuery('#modalDetail').off('show.bs.modal')
                .on('show.bs.modal', function () {{ loadDetail({args_js}); }})
                .modal('show');
        }} else {{ loadDetail({args_js}); }}
    }}"""


async def _open_board(context, url):
    """Load the dispatch board, retrying transient nav timeouts (the server can
    be slow/congested, especially during the retry pass)."""
    page = await context.new_page()
    last = None
    for attempt in (1, 2, 3):
        try:
            await page.goto(url, wait_until="domcontentloaded", timeout=60000)
            await page.wait_for_selector(CONTAINER_SELECTOR, timeout=45000)
            return page
        except (PWTimeout, PWError) as e:
            last = e
            if attempt < 3:
                await page.wait_for_timeout(3000 * attempt)
    await page.close()
    raise last


async def _args_map(page) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for c in await page.locator(CONTAINER_SELECTOR).all():
        m = re.search(r"loadDetail\((.*?)\)", await c.get_attribute("onclick") or "")
        if m:
            out[_jobnum(m.group(1).strip())] = m.group(1).strip()
    return out


async def _process_job(page, context, job: str, args_js: str) -> Dict[str, Any]:
    res = {"rev": None, "pdf_path": None, "dr_rev": None, "dr_pdf_path": None, "no_so": False}
    await page.evaluate(_trigger_js(args_js))
    link = page.locator("#modalDetail a").filter(has_text=re.compile(re.escape(job)))
    try:
        await link.first.wait_for(state="attached", timeout=90000)
    except PWTimeout:
        return res  # modal never showed its docs — worth retrying on the next pass

    # Collect every document link once, then pick the latest of each type we
    # want by its pid type prefix — the Sales Order and the drive run.
    docs = []
    for a in await page.locator("#modalDetail a").all():
        href = await a.get_attribute("href") or ""
        if "downloaddoc.aspx" in href.lower():
            docs.append((href, _parse_doc(href)))
    # Surface the raw pid types so the run log can name what's actually there
    # (the key diagnostic when the quote/drive run isn't being recognized).
    res["doc_types"] = sorted({d.get("type") or "?" for _, d in docs})

    so = _latest_of_type(docs, SO_TYPE)
    if so:
        href, doc = so
        res["rev"] = doc["rev"]
        res["so_source_type"] = doc.get("type", "")
        accepted = await _download_sales_order(
            context,
            page.url,
            href,
            SALES_ORDER_DIR / job / _so_filename(job, doc["rev"]),
            job,
            _required_so_document_kind(doc),
        )
        res["pdf_path"] = accepted.path
        res["so_validation"] = accepted.validation.status
        res["so_internal_order"] = accepted.validation.internal_order
        res["so_validation_method"] = accepted.validation.method
        res["so_document_kind"] = accepted.validation.document_kind
        res["so_quarantine"] = accepted.quarantine_path
    else:
        # The docs DID load and there's just no Sales Order among them (e.g.
        # HDX). Terminal — don't burn another 90s wait on it in the retry pass.
        res["no_so"] = True

    # Construction / quote run — only the highly-custom orders have one. More
    # than one file can match (a qt-run txt and a D64 wheel-construction xlsx,
    # say); archive them all, and link the best as the primary.
    runs = _run_docs(docs) if not so or res.get("pdf_path") else []
    if runs:
        res["dr_rev"] = runs[0][1]["rev"]
        res["dr_count"] = len(runs)
        for href, doc in runs:
            got = await _download(context, page.url, href, DRIVE_RUN_DIR / job / _run_filename(job, doc))
            if got and not res["dr_pdf_path"]:
                res["dr_pdf_path"] = got

    return res


async def _worker(context, url, queue, results, total):
    # A worker that can't even load the board sits the round out rather than
    # crashing the whole fetch — the shared queue is drained by the others.
    try:
        page = await _open_board(context, url)
        amap = await _args_map(page)
    except Exception as e:  # noqa: BLE001
        log.warning("SO worker could not open the board (%s); sitting out this pass", e)
        return
    while True:
        try:
            job = queue.get_nowait()
        except asyncio.QueueEmpty:
            break
        try:
            args_js = amap.get(job)
            results[job] = await _process_job(page, context, job, args_js) if args_js else {"rev": None, "pdf_path": None}
        except Exception as e:  # noqa: BLE001
            log.warning("SO fetch error for %s: %s", job, e)
            results.setdefault(job, {"rev": None, "pdf_path": None})
        finally:
            r = results.get(job) or {}
            if r.get("pdf_path"):
                mark = "ok"
            elif r.get("no_so"):
                mark = "no SO"
            elif r.get("rev") is not None:
                mark = "no pdf"
            else:
                mark = "no docs (timeout)"
            if r.get("dr_pdf_path"):
                mark += " +DriveRun"
            log.info("  sales orders %d/%d  (%s: %s)", len(results), total, job, mark)
            with contextlib.suppress(Exception):
                await page.evaluate("() => window.jQuery && jQuery('#modalDetail').modal('hide')")
                await page.wait_for_timeout(300)
            queue.task_done()
    with contextlib.suppress(Exception):
        await page.close()


async def _afetch_all_unlocked(job_numbers: List[str]) -> Dict[str, Dict[str, Any]]:
    if not STORAGE_STATE_PATH.exists():
        raise RuntimeError(f"No saved session at {STORAGE_STATE_PATH}. Run `python login.py`.")
    url = CBC_QUEUE_URL or CBC_URL
    total = len(job_numbers)
    results: Dict[str, Dict[str, Any]] = {}
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(storage_state=str(STORAGE_STATE_PATH), accept_downloads=True)
        queue: asyncio.Queue = asyncio.Queue()
        for j in job_numbers:
            queue.put_nowait(j)
        n = min(SO_CONCURRENCY, total) or 1
        # return_exceptions so one worker dying never cancels the others.
        await asyncio.gather(
            *[asyncio.create_task(_worker(context, url, queue, results, total)) for _ in range(n)],
            return_exceptions=True,
        )
        with contextlib.suppress(Exception):
            await browser.close()
    return results


async def _afetch_all(job_numbers: List[str]) -> Dict[str, Dict[str, Any]]:
    """Fetch a watcher batch without overlapping the historical search flow."""
    with cbc_fetch_lock():
        return await _afetch_all_unlocked(job_numbers)


# --------------------------------------------------------------------------- #
# Public entry point                                                          #
# --------------------------------------------------------------------------- #
def _terminal(r: Dict[str, Any]) -> bool:
    """A fetch result that shouldn't be retried: we have the pdf, or the modal's
    documents loaded and there's genuinely no Sales Order to fetch (e.g. HDX)."""
    return bool(r.get("pdf_path") or r.get("no_so"))


def enrich_with_sales_orders(jobs: List[Dict[str, Any]], max_passes: int = 2,
                             deep_folders: bool = True) -> None:
    """Mutate `jobs` in place, attaching sales-order + folder fields (see module
    docstring). Opens every job's detail modal in parallel — the slow step.

    Under heavy parallel load the doc server occasionally lets a modal time out,
    leaving a job empty. So we make up to `max_passes` passes, re-running only
    the jobs that came back incomplete; that leftover set is small and far less
    contended, so the stragglers come through — without changing concurrency.

    `deep_folders=False` (the intraday watcher) skips the full-tree folder sweep —
    a standard folder is still found by direct lookup, and non-standard ones are
    picked up by the next daily run — so a folderless job never costs a ~12K-folder
    walk on every enrichment.
    """
    by_job = {j["job"]: j for j in jobs if j.get("job")}
    if not by_job:
        return

    index = _find_autocad_folders(list(by_job.keys()), deep=deep_folders)

    so_results: Dict[str, Dict[str, Any]] = {}
    seen_types: set = set()
    todo = list(by_job.keys())
    for p in range(1, max_passes + 1):
        log.info("Sales-order fetch pass %d: %d job(s), %d parallel...", p, len(todo), SO_CONCURRENCY)
        try:
            res = asyncio.run(_afetch_all(todo))
        except Exception as e:  # noqa: BLE001 - keep earlier passes' results
            log.warning("Sales-order fetch pass %d failed (%s); keeping results so far", p, e)
            break
        for k, v in res.items():
            seen_types.update(v.get("doc_types") or [])
            old = so_results.get(k)
            # Keep the best result seen: a downloaded pdf beats a confirmed
            # no-SO beats a bare rev beats nothing.
            if old is None or (v.get("pdf_path") and not old.get("pdf_path")) \
                    or (v.get("rev") is not None and old.get("rev") is None) \
                    or (v.get("no_so") and not _terminal(old)):
                # Don't lose an earlier pass's drive run if this result missed it.
                if old and old.get("dr_pdf_path") and not v.get("dr_pdf_path"):
                    v = {**v, "dr_pdf_path": old["dr_pdf_path"], "dr_rev": old.get("dr_rev"),
                         "dr_count": old.get("dr_count")}
                so_results[k] = v
        todo = [k for k in by_job if not _terminal(so_results.get(k) or {})]
        if not todo:
            break
        if p < max_passes:
            log.info("  %d job(s) still incomplete; retrying those.", len(todo))

    n_co = n_dl = n_dr = n_dr_folder = n_items = 0
    line_item_updates: List[Dict[str, Any]] = []
    for jn, j in by_job.items():
        r = so_results.get(jn, {})
        rev = r.get("rev")
        pdf = r.get("pdf_path")
        parsed = parse_sales_order_pdf(pdf) if pdf else {}
        selected_doc = {"type": r.get("so_source_type", ""), "rev": rev}
        j["co_number"] = _co_number_for_so_doc(selected_doc, parsed) if pdf else 0
        if j["co_number"]:
            n_co += 1

        j["so_validation"] = r.get("so_validation", "")
        j["so_internal_order"] = r.get("so_internal_order", "")
        j["so_validation_method"] = r.get("so_validation_method", "")
        j["so_document_kind"] = r.get("so_document_kind", "")
        j["so_source_type"] = r.get("so_source_type", "")
        j["so_quarantine"] = r.get("so_quarantine", "")
        j["co_history"] = parsed.get("co_history", [])
        j["so_design_desc"] = parsed.get("design_desc", "")
        j["so_size"] = parsed.get("size", "")
        j["so_arrangement"] = parsed.get("arrangement", "")
        j["so_motor_pos"] = parsed.get("motor_pos", "")
        j["so_class"] = parsed.get("fan_class", "")
        j["so_rotation"] = parsed.get("rotation", "")
        j["so_discharge"] = parsed.get("discharge", "")
        j["so_pct_width"] = parsed.get("pct_width", "")
        j["so_wheel_type"] = parsed.get("wheel_type", "")
        j["so_design_temp"] = parsed.get("design_temp", "")
        j["so_max_temp"] = parsed.get("max_temp", "")
        j["so_special_temp"] = parsed.get("special_temp", "") if pdf else ""
        j["so_pdf"] = pdf or ""
        # Transmittal data carried on every order so it lives in the master.
        j["so_emails"] = parsed.get("emails", [])
        j["so_po"] = parsed.get("customer_po", "")
        j["so_released"] = bool(parsed.get("released", False))
        if pdf:
            n_dl += 1
            # Freshness stamp (flows into the state + master): lets merge_backfill
            # skip re-imposing an older backfill scan over this verification.
            j["so_verified_at"] = datetime.now().isoformat(timespec="seconds")

        # Line items: tag (rules + AI cache), surface on the job for the report
        # and snapshot, and record in the lookup store.
        items = parsed.get("line_items") or []
        if pdf:
            line_item_updates.append({
                "job": jn,
                "items": items,
                "customer": j.get("customer", ""),
                "co_number": j["co_number"],
                "so_pdf": pdf,
                "arrangement": parsed.get("arrangement", ""),
                "parts_only": bool(parsed.get("parts_only", False)),
                "job_number": parsed.get("job_number", ""),
                "so_design_desc": parsed.get("design_desc", ""),
                "so_size": parsed.get("size", ""),
                "so_arrangement": parsed.get("arrangement", ""),
                "so_motor_pos": parsed.get("motor_pos", ""),
                "so_class": parsed.get("fan_class", ""),
                "so_rotation": parsed.get("rotation", ""),
                "so_discharge": parsed.get("discharge", ""),
                "so_pct_width": parsed.get("pct_width", ""),
                "so_wheel_type": parsed.get("wheel_type", ""),
                "so_design_temp": parsed.get("design_temp", ""),
                "so_max_temp": parsed.get("max_temp", ""),
                "so_special_temp": parsed.get("special_temp", ""),
                "source_pdf_sha256": sales_order_sha256(pdf),
            })
            n_items += len(items)
        j["line_items"] = items
        j["line_item_tags"] = line_items.tags_label(items)

        info = index.get(jn)
        if info:
            j["job_type"] = info["type"]
            j["job_folder"] = str(info["path"])
            j["dwg_extras"] = info.get("dwg_extras", {})
            j["dwg_missing_std"] = info.get("dwg_missing_std", False)
            j["so_imi"] = info.get("imi", "")
        else:
            j["job_type"] = ""
            # Fall back to the SO archive folder when there's no AutoCAD folder yet.
            j["job_folder"] = str(SALES_ORDER_DIR / jn) if pdf else ""
            j["dwg_extras"] = {}
            j["dwg_missing_std"] = False
            j["so_imi"] = ""

        # Construction / quote run: presence alone flags a highly-custom fan.
        # Gather every run into the job's Quote-Runs archive — the modal
        # downloads plus any that only live in the AutoCAD folder — so the count
        # and the link both describe that one folder (clicking 'YES (X)' lands you
        # on exactly X files). More than one means someone should review which is
        # the real run.
        dr_pdf = r.get("dr_pdf_path")
        had_doc_run = bool(dr_pdf or r.get("dr_rev") is not None)
        archived = _archive_folder_runs(jn, info["path"] if info else None)
        if archived and not had_doc_run:
            n_dr_folder += 1
        j["has_drive_run"] = had_doc_run or bool(archived)
        if not dr_pdf and archived:
            from run_rank import rank_paths
            dr_pdf = str(rank_paths(archived)[0])   # most-current run represents the order
        j["drive_run_pdf"] = dr_pdf or ""
        # X = the distinct run files actually in the archive folder the link opens.
        j["drive_run_count"] = len(archived) if j["has_drive_run"] else 0
        j["drive_run_rev"] = r.get("dr_rev")
        # Read the run with whichever template matches its format — keyed mostly
        # by design # (Design 64 -> wheel-construction xlsx, HDX -> Qt Run text,
        # others -> pdf). Handles every run extension, not just .pdf.
        dparsed = parse_quote_run(dr_pdf, design=j.get("design")) if dr_pdf else {}
        j["drive_run"] = dparsed.get("fields", {})
        j["drive_run_summary"] = dparsed.get("summary", "")
        j["drive_run_template"] = dparsed.get("template", "")
        if j["has_drive_run"]:
            n_dr += 1

    try:
        line_items.record_jobs_atomic(line_item_updates)
        # record_jobs_atomic applies any AI cache entries to these same item
        # lists, so refresh the labels after the transaction.
        for update in line_item_updates:
            by_job[update["job"]]["line_item_tags"] = line_items.tags_label(update["items"])
        log.info("Line items: %d captured across %d parsed order(s) -> %s",
                 n_items, n_dl, line_items.store_path())
    except OSError as e:  # never let the lookup store sink the daily run
        log.warning("Could not save the line-items store (%s)", e)

    # Similar-order suggester: for each enriched order, shortlist the backlog
    # orders that share its rare SO features AND already have custom drawings
    # on file. One index for the whole batch; lands on the job dict -> master ->
    # the live workbook's "DWG Reuse" column and the new-order notification.
    try:
        import find_orders
        ridx = find_orders.build_index(line_items.load_store(),
                                       dwg=autocad_scan.load_progress())
        n_sugg = 0
        for jn, j in by_job.items():
            sugg = find_orders.reuse_suggestions(ridx, j.get("line_items") or [],
                                                 exclude_job=jn)
            j["dwg_reuse"] = sugg
            j["dwg_reuse_label"] = find_orders.reuse_label(sugg)
            j["dwg_reuse_note"] = find_orders.reuse_note(sugg)
            if sugg:
                n_sugg += 1
        if n_sugg:
            log.info("DWG reuse: %d order(s) matched backlog jobs with custom "
                     "drawings for the same rare features.", n_sugg)
    except Exception as e:  # noqa: BLE001 - a suggestion is a nicety, never fatal
        log.warning("Similar-order suggestions skipped (%s)", e)

    log.info("Sales orders: %d jobs have a SO, %d at a change order, %d still missing a SO.",
             n_dl, n_co, len(by_job) - n_dl)
    log.info("Quote/drive runs: %d job(s) have one (highly custom; %d found in the "
             "AutoCAD folder rather than the documents).", n_dr, n_dr_folder)
    if n_dr == 0 and seen_types:
        log.info("No quote run matched DRIVE_RUN_TYPES=%s or DRIVE_RUN_NAME_PATTERNS. "
                 "pid types seen on the board: %s — a run filed under another "
                 "type/name needs adding to those settings in .env.",
                 DRIVE_RUN_TYPES, sorted(seen_types))


def _stamp_verified_today(order: str) -> None:
    """Record that this order's Sales Order was just (re-)read, in the watcher's
    per-day live state — the same `verified_at` the watcher writes — so a
    'read today' check sees it. Best-effort; a stamp failure never breaks a
    refresh. Lazy-imports live_state to avoid an import cycle at module load."""
    try:
        import live_state
        from datetime import date, datetime
        d = date.today()
        state = live_state.load_state(d)
        entry = state.get(str(order)) or {}
        entry["verified_at"] = datetime.now().isoformat(timespec="seconds")
        state[str(order)] = entry
        live_state.save_state(state, d)
    except Exception as e:  # noqa: BLE001
        log.warning("Could not stamp verified_at for %s: %s", order, e)


def refresh_order_so(order: str, job: Dict[str, Any] | None = None,
                     deep_folders: bool = False) -> Dict[str, Any]:
    """Fetch ONE order's Sales Order online and re-parse it, returning the
    enriched job dict (so_pdf, so_emails, so_po, so_released, so_imi, …). This is
    the single-order entry point for anything that needs a fresh SO on demand —
    e.g. the transmittal flow when an order's SO hasn't been re-read today. It
    reuses enrich_with_sales_orders (the same fetch the watcher/daily run use) and
    stamps the order's verified_at so the read is recorded.

    The order must be reachable on the board (the live queue) for the fetch to
    open its detail modal. Raises if there's no saved session."""
    j = dict(job or {})
    j.setdefault("job", str(order))
    enrich_with_sales_orders([j], deep_folders=deep_folders)
    _stamp_verified_today(str(order))
    return j
