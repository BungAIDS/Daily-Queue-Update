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
    drive_run      dict  (parsed drive-run fields, pdf runs only; see drive_run.py)
    drive_run_summary str (compact one-liner of the drive-run fields)
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
from pathlib import Path
from typing import Any, Dict, List
from urllib.parse import urlparse, parse_qs, urljoin

from playwright.async_api import async_playwright, TimeoutError as PWTimeout, Error as PWError

from config import (
    CBC_URL, CBC_QUEUE_URL, STORAGE_STATE_PATH,
    SALES_ORDER_DIR, DRIVE_RUN_DIR, DRIVE_RUN_TYPES, DRIVE_RUN_NAME_PATTERNS,
    SO_CONCURRENCY, AUTOCAD_JOBS_DIR,
)
from drive_run import parse_drive_run_pdf
from scraper import CONTAINER_SELECTOR
import autocad_scan

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
CO_START = re.compile(r"^\s*C\s*/?\s*O\s*#?\s*\d", re.I)
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
def _run_files_in_folder(folder: Path) -> List[Path]:
    """Quote-run files in a job's AutoCAD folder, searched recursively — they
    are often tucked in a subfolder (e.g. ENG REF\\420410 qt  run.txt). Some
    orders never get the run attached to their cbcinsider documents at all,
    so the folder is the only place it lives."""
    try:
        return sorted(f for f in folder.rglob("*") if f.is_file() and _is_run_name(f.name))
    except OSError as e:
        log.warning("  could not scan %s for quote-run files (%s)", folder, e)
        return []


def _find_autocad_folders(job_numbers: List[str]) -> Dict[str, Dict[str, Any]]:
    """Locate each job under AUTOCAD_JOBS_DIR/<type>/<intermediate>/<job>.

    One sweep of the two directory levels builds the whole index — a glob per
    job re-scans every <type>/<intermediate> dir on the network share once per
    job, which is N full sweeps for an N-job board. Returns {job: {type, path,
    dwg_extras, dwg_missing_std}}; {} if the drive isn't reachable. The <type>
    a job sits under is its job type. While we have the folder we also scan it
    for the job's custom drawings (the extra -NN suffixes), reusing autocad_scan.
    """
    out: Dict[str, Dict[str, Any]] = {}
    root = AUTOCAD_JOBS_DIR
    wanted = set(job_numbers)
    try:
        if not root.exists():
            log.warning("AutoCAD jobs root not reachable: %s (folder links disabled)", root)
            return out
        for m in root.glob("*/*/*"):
            job = autocad_scan.job_key(m.name)  # "421314 ACME CORP" -> "421314"
            if job not in wanted or job in out or not m.is_dir():
                continue
            info: Dict[str, Any] = {"type": m.relative_to(root).parts[0], "path": m,
                                    "dwg_extras": {}, "dwg_missing_std": False}
            try:  # live scan of this job's custom DWGs (names only — never opens a file)
                names = [f.name for f in m.glob("*") if f.is_file()]
                rec = autocad_scan.build_record(job, info["type"], str(m),
                                                autocad_scan.scan_files(names, job))
                info["dwg_extras"], info["dwg_missing_std"] = rec["extras"], rec["missing_std"]
            except OSError as e:
                log.warning("  could not scan DWGs for %s (%s)", job, e)
            out[job] = info
            if len(out) == len(wanted):
                break  # found every job on the board — stop walking the share
        log.info("Located %d/%d AutoCAD job folders under %s", len(out), len(job_numbers), root)
    except OSError as e:
        log.warning("Could not look up AutoCAD folders (%s); folder links disabled", e)
    return out


# --------------------------------------------------------------------------- #
# PDF parsing                                                                 #
# --------------------------------------------------------------------------- #
def _recon_lines(page, x_tol: float = 1.5) -> List[str]:
    """Rebuild text lines from word positions so spaces survive (plain
    extraction glues the Notes text together)."""
    words = page.extract_words(x_tolerance=x_tol, keep_blank_chars=False, use_text_flow=False)
    rows: Dict[int, list] = {}
    for w in words:
        rows.setdefault(round(w["top"]), []).append(w)
    out = []
    for top in sorted(rows):
        ws = sorted(rows[top], key=lambda w: w["x0"])
        out.append(" ".join(w["text"] for w in ws))
    return out


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


def parse_sales_order_pdf(path: str | Path) -> Dict[str, Any]:
    """Pull Design/Size/Arrangement + change-order history out of an SO pdf."""
    res = {"design_desc": "", "size": "", "arrangement": "", "motor_pos": "", "fan_class": "",
           "rotation": "", "discharge": "", "pct_width": "", "wheel_type": "", "temp": "",
           "design_temp": "", "max_temp": "", "special_temp": "0",
           "header_co": None, "co_history": []}
    try:
        import pdfplumber
    except ImportError:
        log.warning("pdfplumber not installed; cannot parse SO pdfs (pip install pdfplumber)")
        return res
    try:
        with pdfplumber.open(str(path)) as pdf:
            p1 = pdf.pages[0]
            for ln in (p1.extract_text() or "").splitlines()[:8]:
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
            for page in pdf.pages:
                spec = _spec_from_tables(page.extract_tables())
                if spec.get("Size") or spec.get("Arrangement"):
                    recon = "\n".join(_recon_lines(page))
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
            for page in pdf.pages:
                for ln in _recon_lines(page):
                    if CO_START.match(ln):
                        res["co_history"].append(ln.strip())
            # Special temperature rating from the "Suitable for <temp>" phrase.
            raw_all = "\n".join((page.extract_text() or "") for page in pdf.pages)
            mt = TEMP_RE.search(raw_all)
            if mt:
                res["temp"] = re.sub(r"\s+", "", mt.group(1))
            res["special_temp"] = _special_temp(res["design_temp"], res["max_temp"], res["temp"])
    except Exception as e:  # noqa: BLE001 - never let a bad pdf fail the run
        log.warning("Could not parse SO pdf %s: %s", path, e)
    return res


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
    """Normalize a pid type for comparison: lowercase, drop a leading 'cbc_'."""
    t = (t or "").lower()
    return t[4:] if t.startswith("cbc_") else t


def _latest_of_type(docs: List, type_name: str):
    """Highest-revision (href, doc) whose pid type matches type_name, or None."""
    want = _norm_type(type_name)
    matches = [hd for hd in docs if _norm_type(hd[1].get("type")) == want]
    return max(matches, key=lambda hd: hd[1].get("rev") or 0) if matches else None


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
    if rev and rev > 1:
        return f"{job} - Sales Order CO#{rev - 1}.pdf"
    return f"{job} - Sales Order (original).pdf"


def _doc_filename(job: str, label: str, rev: int | None) -> str:
    """Archive filename for a non-SO document (e.g. the drive run)."""
    return f"{job} - {label} rev {rev}.pdf" if rev and rev > 1 else f"{job} - {label}.pdf"


def _run_filename(job: str, doc: Dict[str, Any]) -> str:
    """Archive name for a quote-run document. Keeps the site's own file name —
    it carries the identifying naming and the real extension (.txt qt runs,
    .xlsx D64 wheel constructions, .pdf HDX runs) — prefixed with the job
    number when it isn't already in it."""
    fn = re.sub(r'[<>:"/\\|?*]', "_", (doc.get("fn") or "").strip())
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
        res["pdf_path"] = await _download(
            context, page.url, href, SALES_ORDER_DIR / job / _so_filename(job, doc["rev"]))
    else:
        # The docs DID load and there's just no Sales Order among them (e.g.
        # HDX). Terminal — don't burn another 90s wait on it in the retry pass.
        res["no_so"] = True

    # Construction / quote run — only the highly-custom orders have one. More
    # than one file can match (a qt-run txt and a D64 wheel-construction xlsx,
    # say); archive them all, and link the best as the primary.
    runs = _run_docs(docs)
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


async def _afetch_all(job_numbers: List[str]) -> Dict[str, Dict[str, Any]]:
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


# --------------------------------------------------------------------------- #
# Public entry point                                                          #
# --------------------------------------------------------------------------- #
def _terminal(r: Dict[str, Any]) -> bool:
    """A fetch result that shouldn't be retried: we have the pdf, or the modal's
    documents loaded and there's genuinely no Sales Order to fetch (e.g. HDX)."""
    return bool(r.get("pdf_path") or r.get("no_so"))


def enrich_with_sales_orders(jobs: List[Dict[str, Any]], max_passes: int = 2) -> None:
    """Mutate `jobs` in place, attaching sales-order + folder fields (see module
    docstring). Opens every job's detail modal in parallel — the slow step.

    Under heavy parallel load the doc server occasionally lets a modal time out,
    leaving a job empty. So we make up to `max_passes` passes, re-running only
    the jobs that came back incomplete; that leftover set is small and far less
    contended, so the stragglers come through — without changing concurrency.
    """
    by_job = {j["job"]: j for j in jobs if j.get("job")}
    if not by_job:
        return

    index = _find_autocad_folders(list(by_job.keys()))

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

    n_co = n_dl = n_dr = n_dr_folder = 0
    for jn, j in by_job.items():
        r = so_results.get(jn, {})
        rev = r.get("rev")
        j["co_number"] = (rev - 1) if rev and rev > 1 else 0
        if j["co_number"]:
            n_co += 1

        pdf = r.get("pdf_path")
        parsed = parse_sales_order_pdf(pdf) if pdf else {}
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
        if pdf:
            n_dl += 1

        info = index.get(jn)
        if info:
            j["job_type"] = info["type"]
            j["job_folder"] = str(info["path"])
            j["dwg_extras"] = info.get("dwg_extras", {})
            j["dwg_missing_std"] = info.get("dwg_missing_std", False)
        else:
            j["job_type"] = ""
            # Fall back to the SO archive folder when there's no AutoCAD folder yet.
            j["job_folder"] = str(SALES_ORDER_DIR / jn) if pdf else ""
            j["dwg_extras"] = {}
            j["dwg_missing_std"] = False

        # Construction / quote run: presence alone flags a highly-custom fan.
        # More than one match (drive_run_count > 1) means someone should review
        # which file is the real run — the report flags it.
        dr_pdf = r.get("dr_pdf_path")
        dr_count = r.get("dr_count") or 0
        j["has_drive_run"] = bool(dr_pdf or r.get("dr_rev") is not None)
        if not j["has_drive_run"] and info:
            # Not attached to the order's documents — some runs only live in
            # the job's AutoCAD folder. Link the file in place (no download).
            hits = _run_files_in_folder(info["path"])
            if hits:
                dr_pdf = str(hits[0])
                dr_count = len(hits)
                j["has_drive_run"] = True
                n_dr_folder += 1
        j["drive_run_pdf"] = dr_pdf or ""
        j["drive_run_count"] = dr_count if j["has_drive_run"] else 0
        j["drive_run_rev"] = r.get("dr_rev")
        dparsed = parse_drive_run_pdf(dr_pdf) if dr_pdf and str(dr_pdf).lower().endswith(".pdf") else {}
        j["drive_run"] = dparsed.get("fields", {})
        j["drive_run_summary"] = dparsed.get("summary", "")
        if j["has_drive_run"]:
            n_dr += 1

    log.info("Sales orders: %d jobs have a SO, %d at a change order, %d still missing a SO.",
             n_dl, n_co, len(by_job) - n_dl)
    log.info("Quote/drive runs: %d job(s) have one (highly custom; %d found in the "
             "AutoCAD folder rather than the documents).", n_dr, n_dr_folder)
    if n_dr == 0 and seen_types:
        log.info("No quote run matched DRIVE_RUN_TYPES=%s or DRIVE_RUN_NAME_PATTERNS. "
                 "pid types seen on the board: %s — a run filed under another "
                 "type/name needs adding to those settings in .env.",
                 DRIVE_RUN_TYPES, sorted(seen_types))
