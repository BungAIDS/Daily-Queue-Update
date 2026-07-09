"""Read scanned (image-only) quote-run PDFs with Claude vision.

The quote-run sweep flags runs whose PDF has no text layer as
"PDF (no text layer)" — pdfplumber gets nothing out of them. This script picks
up exactly those runs, renders the first page(s) to images (pypdfium2, already
installed as a pdfplumber dependency), and asks Claude to classify + extract in
one call:

  - a scanned selection-program Qt Run / quote form  -> the fields, same names
    the text parser uses, status OK, template "pdf_vision"
  - an engineering drawing (nothing to extract)      -> status DRAWING, so it
    stops showing up as "needs attention" forever
  - unreadable / something else                      -> stays flagged

Every read also stores a FULL TRANSCRIPTION of the document (run["vision"]
["transcript"] in the progress store), so future field additions can re-parse
the stored text for free instead of re-paying the API.

Results are written back into the quote-run progress store (and from there the
workbook + live_master.json via master_sync), and each PDF is only ever paid
for once — a run that has a vision result is skipped unless --redo is passed.

QUALITY LOOP (so a re-read is worth paying for):
  - After each read, apply_vision_qc validates the fields (numeric plausibility,
    arrangement whitelist, model-vs-transcript disagreement). Garbled values are
    REPAIRED from the clean transcript parse where possible.
  - A run that still looks wrong is flagged CHECK VISION and re-read ONCE — but
    the re-read is escalated: a higher-resolution render plus the specific list
    of what looked wrong, so it isn't a coin-flip repeat.
  - The two readings are COMPARED. If they still disagree (or agree on an
    implausible value) after MAX_VISION_ATTEMPTS, the run goes to NEEDS HUMAN
    (terminal — never auto-re-read/re-paid again; orange in the workbook, with a
    reason citing the conflict). A later parser fix that makes it clean clears
    it back to OK.

Cost: ~1-2k input tokens per page + a few hundred output tokens. On the default
Haiku model that is well under a cent per document; the whole 100+ document
backlog is on the order of a dollar, one time.

Usage:
    python pdf_vision.py --limit 5        # TRIAL: read 5, eyeball the fields
    python pdf_vision.py                  # read every flagged PDF
    python pdf_vision.py 406244           # just this job's flagged PDFs
    python pdf_vision.py --redo --limit 3 # re-read PDFs that already have a result
    python pdf_vision.py --model claude-opus-4-8   # heavier model for bad scans

Needs ANTHROPIC_API_KEY in .env (the daily brief already uses it). Model comes
from PDF_VISION_MODEL (defaults to CLAUDE_MODEL). Never run two copies at once
(shared progress store).
"""
from __future__ import annotations

import argparse
import base64
import io
import json
import logging
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from config import ANTHROPIC_API_KEY, PDF_VISION_MODEL, PDF_VISION_MAX_PAGES
from templates import summarize

log = logging.getLogger("pdf-vision")

NO_TEXT_STATUS = "PDF (no text layer)"
DRAWING_STATUS = "DRAWING"

# The long edge images are downscaled to before upload. Keeps a page at
# ~1.5-2k input tokens; plenty of resolution to read a typed form.
_LONG_EDGE = 1568

# Field names the text parser produces — given to the model so a scanned run
# comes back under the same names and lands in the same workbook columns.
_FIELD_NAMES = [
    "Serial", "Size", "Design", "Class", "Fan Type", "Arrangement", "% Width",
    "Discharge", "Rotation", "Effective Wheel Dia", "CFM", "SP", "BHP", "RPM",
    "Air Temp F", "Max Temp F", "Density", "Max HP", "Max RPM", "Tip Speed FPM",
    "Shaft Dia", "Brg Centers", "Critical Speed RPM",
    "Shaft Length", "OH", "BX", "STB", "TG&P", "STH",
    "Bearing Size", "Bearing Series", "Bearing L10 Hr",
    "Blade Material", "Blade Gauge", "Sideplate Material", "Sideplate Gauge",
    "Backplate Material", "Backplate Gauge", "Liner Material", "Liner Gauge",
    "Wheel Material", "Hub", "Coupling", "Drive",
    "Blades", "Max RPM Wheel Only", "Wheel Resonance CPM",
    "Wheel Weight Lb", "Wheel Thrust Lb", "Wheel WR2",
    "Rotor WR2", "Rotor Max RPM", "Rotor Material",
    "Motor Frame", "Motor Position", "Motor Enclosure", "Motor Weight Lb",
    "Housing Construction", "Fan Outlet Area FT2",
    "Total Weight Lb", "Total Price",
]


def build_prompt(hints: Optional[List[str]] = None) -> str:
    names = ", ".join(_FIELD_NAMES)
    lead = ""
    if hints:
        # A re-read: tell the model exactly what a prior reading got wrong, so a
        # second identical pass isn't just wishful. Focus its attention.
        lead = ("A PREVIOUS automated reading of THIS document looked wrong:\n  - "
                + "\n  - ".join(hints[:8]) + "\n"
                "Look at those areas especially carefully and read the digits/"
                "letters exactly as printed. If a value is genuinely illegible, "
                "OMIT it rather than guessing.\n\n")
    return (
        lead
        + "You are reading a scanned document from a fan manufacturer's job folder.\n"
        "First decide what it is:\n"
        '  - "quote_run": a selection-program Qt Run / quote run / construction run'
        " form (a typed data sheet: SIZE/DESIGN/ARR spec line, CFM/SP/BHP"
        " performance line, a wheel-construction table, shaft/bearing data).\n"
        '  - "drawing": an engineering drawing / dimensional sketch (mostly'
        " geometry and dimension lines, no run data sheet).\n"
        '  - "other": anything else (letter, purchase order, catalog page...).\n'
        "\n"
        "If it is a quote_run, extract every value you can read, using EXACTLY "
        "these field names where they apply (skip fields not on the document; "
        "add extra clearly-labeled values under their own label): " + names + ".\n"
        "Copy values verbatim (keep fractions like '2 15/16' and gauges like "
        "'0.048 (18)'). Do not guess unreadable values — omit them.\n"
        "\n"
        "Also transcribe the document: every piece of readable text, top to "
        "bottom, preserving line breaks with \\n (for a drawing, transcribe the "
        "title block and any labels/notes). Do not invent unreadable text.\n"
        "\n"
        "Reply with ONLY a JSON object, no other text:\n"
        '{"doc_type": "quote_run" | "drawing" | "other",\n'
        ' "fields": {"<Field Name>": "<value>", ...},\n'
        ' "note": "<one short line: what the document is>",\n'
        ' "transcript": "<full transcription>"}\n'
        'For a drawing or other, "fields" must be {}.'
    )


# --------------------------------------------------------------------------- #
# Pure logic (no I/O — unit-tested)                                           #
# --------------------------------------------------------------------------- #
def parse_vision_response(text: str) -> Dict[str, Any]:
    """The model's reply -> {doc_type, fields, note}. Tolerates code fences and
    stray prose; anything unusable comes back as doc_type 'error' (never raises)."""
    raw = (text or "").strip()
    raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw)
    start, end = raw.find("{"), raw.rfind("}")
    if start < 0 or end <= start:
        return {"doc_type": "error", "fields": {}, "note": "no JSON in reply"}
    try:
        data = json.loads(raw[start:end + 1])
    except json.JSONDecodeError as e:
        return {"doc_type": "error", "fields": {}, "note": f"bad JSON ({e})"}
    doc_type = str(data.get("doc_type", "")).strip().lower()
    if doc_type not in ("quote_run", "drawing", "other"):
        doc_type = "other"
    fields: Dict[str, str] = {}
    raw_fields = data.get("fields")
    if isinstance(raw_fields, dict):
        for k, v in list(raw_fields.items())[:60]:      # sanity cap
            key = str(k).strip()
            val = re.sub(r"\s{2,}", " ", str(v).strip())[:200]
            if key and val:
                fields[key] = val
    return {"doc_type": doc_type, "fields": fields,
            "note": str(data.get("note", "")).strip()[:300],
            # Generous cap (a runaway backstop, not a real limit) so a long
            # scanned doc's transcript isn't clipped in the corpus.
            "transcript": str(data.get("transcript", "")).strip()[:80000]}


def apply_vision_result(run: Dict[str, Any], parsed: Dict[str, Any], model: str) -> bool:
    """Fold a parsed vision result into a run record (in place). Returns True
    when the run was updated (a usable classification came back). Tracks the
    attempt count and stashes the PRIOR reading's fields so a re-read can be
    compared against the reading it's replacing (see escalate_to_human)."""
    doc_type = parsed.get("doc_type")
    if doc_type == "error":
        return False                       # leave the run flagged; retried next time
    prior = run.get("vision") or {}
    run["vision"] = {"model": model,
                     "at": datetime.now().isoformat(timespec="seconds"),
                     "doc_type": doc_type, "note": parsed.get("note", ""),
                     # Full transcription — kept so new fields can be re-parsed
                     # from the stored text later without re-paying the API.
                     "transcript": parsed.get("transcript", ""),
                     "attempts": int(prior.get("attempts", 0)) + 1,
                     # The immediately-prior reading, so a re-read can compare.
                     "prior_fields": dict(run.get("fields") or {}) if prior else None}
    run["template"] = "pdf_vision"
    if doc_type == "drawing":
        run["fields"] = {}
        run["status"] = DRAWING_STATUS
        run["summary"] = parsed.get("note") or "engineering drawing (no run data)"
    elif parsed.get("fields"):
        run["fields"] = parsed["fields"]
        run["status"] = "OK"
        run["summary"] = summarize(parsed["fields"])
    else:                                  # readable but nothing extractable
        run["fields"] = {}
        run["status"] = "NO FIELDS"
        run["summary"] = parsed.get("note", "")
    return True


# --------------------------------------------------------------------------- #
# Vision quality control (pure — no API)                                       #
# --------------------------------------------------------------------------- #
CHECK_STATUS = "CHECK VISION"        # extraction looks wrong; re-read (escalated)
NEEDS_HUMAN = "NEEDS HUMAN"          # re-reads exhausted; a person must eyeball it
MAX_VISION_ATTEMPTS = 2             # after this many reads, stop re-paying — go human

# Fields that must be a plain number (commas ok). A slash or stray letters in
# one of these is the classic OCR garble ("4/100 CFM", "20.000 NON-STD").
_INT_FIELDS = {"CFM", "RPM", "Max RPM", "Tip Speed FPM", "Critical Speed RPM",
               "Blades", "Bearing L10 Hr", "Wheel Weight Lb", "Motor Weight Lb",
               "Total Weight Lb", "Total Price", "Wheel Resonance CPM"}
_DEC_FIELDS = {"SP", "BHP", "Density", "Max HP", "Air Temp F", "Max Temp F",
               "% Width", "Stress Ratio at Hub", "Stress Ratio at Bearing",
               "Fan Outlet Area FT2", "Size"}
# Real arrangements: digit family + optional letter suffix (+ digit after the
# letter): 4, 4S, 4S1, 8S, 9H, 3D, 7S1... "781"/"88"/"48" are OCR garbles.
_ARR_OK = re.compile(r"^[1-9](?:[A-Z]{1,2}\d?)?$")


def _implausible(key: str, val: str) -> bool:
    v = (val or "").replace(",", "").strip()
    if key in _INT_FIELDS:
        return not re.fullmatch(r"\d+", v)
    if key in _DEC_FIELDS:
        return not re.fullmatch(r"\d+(?:\.\d+)?", v)
    return False


def vision_qc(run: Dict[str, Any]) -> List[str]:
    """Reasons a vision run's extraction looks WRONG and the PDF should be
    re-read (or hand-checked). Empty list = passes. Where the transcript's
    targeted parse is clean and the model value is garbled, the field is
    REPAIRED in place instead of flagged (build what we can from what we know)."""
    from templates import _parse_chicago_blower
    vision = run.get("vision") or {}
    reasons: List[str] = []
    if not vision:
        return reasons
    transcript = vision.get("transcript", "")
    if not transcript:
        reasons.append("no transcript saved (trial batch) — cheap to re-read")
    fields = run.get("fields") or {}
    parsed = _parse_chicago_blower(transcript) if transcript else {}
    for k, v in list(fields.items()):
        if not _implausible(k, str(v)):
            continue
        alt = parsed.get(k, "")
        if alt and not _implausible(k, alt):
            fields[k] = alt              # repair: pattern value is clean
            reasons.append(f"repaired {k}: {v!r} -> {alt!r}")
        else:
            reasons.append(f"implausible {k}={v!r}")
    arr = str(fields.get("Arrangement", ""))
    if arr and not _ARR_OK.fullmatch(arr):
        alt = str(parsed.get("Arrangement", ""))
        if alt and _ARR_OK.fullmatch(alt):
            fields["Arrangement"] = alt   # model slop; the transcript line is clean
            reasons.append(f"repaired Arrangement: {arr!r} -> {alt!r}")
        else:
            reasons.append(f"odd Arrangement {arr!r} (OCR garble?)")
    # Model vs clean transcript disagreement on hard numbers (>1% apart).
    for k in ("CFM", "RPM", "BHP", "SP"):
        a, b = str(fields.get(k, "")).replace(",", ""), str(parsed.get(k, "")).replace(",", "")
        try:
            fa, fb = float(a), float(b)
        except ValueError:
            continue
        if fa and fb and abs(fa - fb) / max(fa, fb) > 0.01:
            reasons.append(f"{k}: model read {a}, transcript says {b}")
    return reasons


def apply_vision_qc(run: Dict[str, Any]) -> List[str]:
    """Run QC on one vision run; record the verdict. A run with real (non-repair)
    findings is flagged CHECK VISION so it shows amber and gets re-read — UNLESS
    it already exhausted its re-reads (NEEDS HUMAN stays terminal). A run that is
    now clean (e.g. a later pattern repaired the field) is cleared to OK."""
    reasons = vision_qc(run)
    if not run.get("vision"):
        return []
    run["vision"]["suspect"] = [r for r in reasons if not r.startswith("repaired ")]
    if run["vision"]["suspect"]:
        if run.get("status") != NEEDS_HUMAN:   # don't re-open an exhausted run
            run["status"] = CHECK_STATUS
    elif run.get("status") in (CHECK_STATUS, NEEDS_HUMAN):   # now clean -> clear it
        run["status"] = "OK" if run.get("fields") else run.get("status")
    return reasons


_CMP_FIELDS = _INT_FIELDS | _DEC_FIELDS | {"Arrangement"}


def compare_readings(prior: Dict[str, Any], new: Dict[str, Any]) -> List[str]:
    """Hard-number fields where two vision readings of the same PDF disagree —
    the evidence that a re-read didn't converge and a human is needed."""
    prior, new = prior or {}, new or {}
    diffs = []
    for k in sorted(set(prior) | set(new)):
        if k not in _CMP_FIELDS:
            continue
        pv, nv = str(prior.get(k, "")).strip(), str(new.get(k, "")).strip()
        if pv and nv and pv != nv:
            diffs.append(f"{k}: {pv} vs {nv}")
    return diffs


def escalate_to_human(run: Dict[str, Any]) -> None:
    """Called when a re-read still fails QC and attempts are exhausted: mark the
    run NEEDS HUMAN (terminal — never auto-re-read again) with a reason that
    says whether the two reads DISAGREED (compare them) or AGREED-but-implausible."""
    vision = run.get("vision") or {}
    suspect = vision.get("suspect", [])
    diffs = compare_readings(vision.get("prior_fields") or {}, run.get("fields") or {})
    if diffs:
        reason = "two reads disagree — " + "; ".join(diffs[:4])
    else:
        reason = "two reads agree but values look wrong — " + "; ".join(suspect[:4])
    vision["human_reason"] = reason
    run["status"] = NEEDS_HUMAN


def candidate_runs(records: Dict[str, Dict[str, Any]],
                   jobs: Optional[List[str]] = None,
                   redo: bool = False) -> List[Tuple[str, Dict[str, Any]]]:
    """(job, run) pairs to read: runs still flagged 'PDF (no text layer)', plus —
    with redo — runs that already have a vision result. Sorted by job."""
    wanted = {str(j).strip() for j in jobs} if jobs else None
    out: List[Tuple[str, Dict[str, Any]]] = []
    for rec in sorted(records.values(), key=lambda r: r.get("job", "")):
        job = rec.get("job", "")
        if wanted is not None and job not in wanted:
            continue
        for run in rec.get("runs", []):
            if not str(run.get("path", "")).lower().endswith(".pdf"):
                continue
            if run.get("status") == NO_TEXT_STATUS and not run.get("vision"):
                out.append((job, run))
            elif run.get("status") == CHECK_STATUS:
                out.append((job, run))       # QC-flagged: re-read is the fix
            elif redo and run.get("vision"):
                out.append((job, run))
    return out


# --------------------------------------------------------------------------- #
# PDF rendering + the API call                                                #
# --------------------------------------------------------------------------- #
def render_pdf_images(path: Path, max_pages: int = PDF_VISION_MAX_PAGES,
                      long_edge: int = _LONG_EDGE, scale: float = 2.0) -> List[bytes]:
    """First page(s) of the PDF as PNG bytes, downscaled so a page stays at a
    sane token cost. pypdfium2 ships with pdfplumber, so no new dependency.
    A re-read passes a bigger `long_edge`/`scale` to give the model more to work
    with (2576px is the model's max useful resolution)."""
    import pypdfium2 as pdfium
    pdf = pdfium.PdfDocument(str(path))
    try:
        images: List[bytes] = []
        for i in range(min(len(pdf), max_pages)):
            pil = pdf[i].render(scale=scale).to_pil()
            if max(pil.size) > long_edge:
                ratio = long_edge / max(pil.size)
                pil = pil.resize((max(1, int(pil.width * ratio)),
                                  max(1, int(pil.height * ratio))))
            buf = io.BytesIO()
            pil.save(buf, format="PNG")
            images.append(buf.getvalue())
        return images
    finally:
        pdf.close()


def read_scanned_pdf(path: Path, model: str = PDF_VISION_MODEL,
                     max_pages: int = PDF_VISION_MAX_PAGES,
                     hints: Optional[List[str]] = None,
                     hi_res: bool = False) -> Dict[str, Any]:
    """Render one PDF and ask Claude to classify + extract. Returns the parsed
    {doc_type, fields, note} dict; doc_type 'error' on any failure (never raises).
    A re-read passes `hints` (what looked wrong before) and `hi_res=True` (render
    at the model's max useful resolution) so the second pass is a real retry,
    not a coin-flip repeat."""
    try:
        if hi_res:
            images = render_pdf_images(path, max_pages=max_pages, long_edge=2576, scale=3.0)
        else:
            images = render_pdf_images(path, max_pages=max_pages)
    except Exception as e:  # noqa: BLE001 - a corrupt PDF must not kill the batch
        return {"doc_type": "error", "fields": {}, "note": f"render failed ({e})"}
    if not images:
        return {"doc_type": "error", "fields": {}, "note": "no pages rendered"}

    import anthropic
    content: List[Dict[str, Any]] = [
        {"type": "image",
         "source": {"type": "base64", "media_type": "image/png",
                    "data": base64.standard_b64encode(img).decode("ascii")}}
        for img in images
    ]
    content.append({"type": "text", "text": build_prompt(hints)})
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    try:
        response = client.messages.create(
            # Room for fields + a full transcription of even a long doc. Output
            # tokens are billed only for what's generated, so short docs (most)
            # cost nothing extra; only a genuinely long scan spends more.
            model=model, max_tokens=12000,
            messages=[{"role": "user", "content": content}],
        )
    except anthropic.APIError as e:
        return {"doc_type": "error", "fields": {}, "note": f"API error ({e})"}
    if getattr(response, "stop_reason", None) == "refusal":
        return {"doc_type": "error", "fields": {}, "note": "model refused"}
    text = next((b.text for b in response.content if b.type == "text"), "")
    return parse_vision_response(text)


# --------------------------------------------------------------------------- #
# CLI                                                                         #
# --------------------------------------------------------------------------- #
def _parse_args(argv: List[str]) -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Read scanned (no-text-layer) quote-run PDFs with Claude vision.")
    ap.add_argument("jobs", nargs="*", help="Only these job numbers (default: all flagged).")
    ap.add_argument("--limit", type=int, default=0,
                    help="Read at most N PDFs this run (0 = all). Use --limit 5 as a trial.")
    ap.add_argument("--model", default=PDF_VISION_MODEL,
                    help=f"Claude model (default {PDF_VISION_MODEL}).")
    ap.add_argument("--max-pages", type=int, default=PDF_VISION_MAX_PAGES,
                    help=f"Pages per PDF sent to the model (default {PDF_VISION_MAX_PAGES}).")
    ap.add_argument("--redo", action="store_true",
                    help="Also re-read PDFs that already have a vision result.")
    return ap.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    args = _parse_args(sys.argv[1:] if argv is None else argv)
    if not ANTHROPIC_API_KEY:
        log.error("ANTHROPIC_API_KEY is not set in .env — cannot call the vision model.")
        return 1

    # The progress store belongs to the sweep; lazy import keeps this module light.
    from quote_run_scan import load_progress, save_progress, write_workbook, WORKBOOK_PATH

    records = load_progress()
    todo = candidate_runs(records, jobs=args.jobs, redo=args.redo)
    if args.limit:
        todo = todo[:args.limit]
    if not todo:
        log.info("Nothing to read — no runs flagged '%s'%s.", NO_TEXT_STATUS,
                 "" if args.redo else " without a vision result")
        return 0

    n_reread = sum(1 for _, r in todo if r.get("status") == CHECK_STATUS and r.get("vision"))
    log.info("Reading %d scanned PDF(s) with %s (~%d page(s) each; %d of them are "
             "escalated re-reads). Rough cost: well under a cent per document on Haiku.",
             len(todo), args.model, args.max_pages, n_reread)
    t0 = time.monotonic()
    counts = {"quote_run": 0, "drawing": 0, "other": 0, "error": 0}
    for n, (job, run) in enumerate(todo, start=1):
        path = Path(run.get("path", ""))
        # A re-read: a prior reading looked wrong. Give the model the specific
        # complaints and a higher-resolution render, so this pass is a real retry.
        prior = run.get("vision") or {}
        is_reread = bool(prior) and run.get("status") == CHECK_STATUS
        hints = prior.get("suspect") if is_reread else None
        parsed = read_scanned_pdf(path, model=args.model, max_pages=args.max_pages,
                                  hints=hints, hi_res=is_reread)
        doc_type = parsed.get("doc_type", "error")
        counts[doc_type] = counts.get(doc_type, 0) + 1
        if apply_vision_result(run, parsed, args.model):
            apply_vision_qc(run)          # repair what we can; re-flag the rest
            # Still bad after enough attempts? Compare the two reads and hand off.
            if run.get("status") == CHECK_STATUS and \
                    run["vision"].get("attempts", 1) >= MAX_VISION_ATTEMPTS:
                escalate_to_human(run)
            save_progress(records)        # each answer costs money — never lose one
            extra = ""
            if run.get("status") == NEEDS_HUMAN:
                extra = f"  [{run['vision'].get('human_reason', '')}]"
            elif run.get("fields"):
                extra = f" ({len(run['fields'])} fields)"
            log.info("  [%d/%d] %s %s -> %s%s", n, len(todo), job, path.name,
                     run["status"], extra)
        else:
            log.warning("  [%d/%d] %s %s -> %s (left flagged; will retry next run)",
                        n, len(todo), job, path.name, parsed.get("note", "error"))

    save_progress(records)
    try:   # fold into the master store, same as the sweep does
        import master_sync
        master_sync.run("quote_runs")
    except Exception as e:  # noqa: BLE001
        log.warning("Could not sync to the live master (%s)", e)
    out = write_workbook(records, WORKBOOK_PATH)
    log.info("Done in %.1fs: %s. Wrote %s", time.monotonic() - t0,
             ", ".join(f"{k}={v}" for k, v in counts.items() if v), out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
