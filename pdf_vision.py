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
]


def build_prompt() -> str:
    names = ", ".join(_FIELD_NAMES)
    return (
        "You are reading a scanned document from a fan manufacturer's job folder.\n"
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
            "transcript": str(data.get("transcript", "")).strip()[:20000]}


def apply_vision_result(run: Dict[str, Any], parsed: Dict[str, Any], model: str) -> bool:
    """Fold a parsed vision result into a run record (in place). Returns True
    when the run was updated (a usable classification came back)."""
    doc_type = parsed.get("doc_type")
    if doc_type == "error":
        return False                       # leave the run flagged; retried next time
    run["vision"] = {"model": model,
                     "at": datetime.now().isoformat(timespec="seconds"),
                     "doc_type": doc_type, "note": parsed.get("note", ""),
                     # Full transcription — kept so new fields can be re-parsed
                     # from the stored text later without re-paying the API.
                     "transcript": parsed.get("transcript", "")}
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
            elif redo and run.get("vision"):
                out.append((job, run))
    return out


# --------------------------------------------------------------------------- #
# PDF rendering + the API call                                                #
# --------------------------------------------------------------------------- #
def render_pdf_images(path: Path, max_pages: int = PDF_VISION_MAX_PAGES,
                      long_edge: int = _LONG_EDGE) -> List[bytes]:
    """First page(s) of the PDF as PNG bytes, downscaled so a page stays at a
    sane token cost. pypdfium2 ships with pdfplumber, so no new dependency."""
    import pypdfium2 as pdfium
    pdf = pdfium.PdfDocument(str(path))
    try:
        images: List[bytes] = []
        for i in range(min(len(pdf), max_pages)):
            pil = pdf[i].render(scale=2.0).to_pil()
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
                     max_pages: int = PDF_VISION_MAX_PAGES) -> Dict[str, Any]:
    """Render one PDF and ask Claude to classify + extract. Returns the parsed
    {doc_type, fields, note} dict; doc_type 'error' on any failure (never raises)."""
    try:
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
    content.append({"type": "text", "text": build_prompt()})
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    try:
        response = client.messages.create(
            model=model, max_tokens=6000,   # fields + the full transcription
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

    log.info("Reading %d scanned PDF(s) with %s (~%d page(s) each). "
             "Rough cost: well under a cent per document on Haiku.",
             len(todo), args.model, args.max_pages)
    t0 = time.monotonic()
    counts = {"quote_run": 0, "drawing": 0, "other": 0, "error": 0}
    for n, (job, run) in enumerate(todo, start=1):
        path = Path(run.get("path", ""))
        parsed = read_scanned_pdf(path, model=args.model, max_pages=args.max_pages)
        doc_type = parsed.get("doc_type", "error")
        counts[doc_type] = counts.get(doc_type, 0) + 1
        if apply_vision_result(run, parsed, args.model):
            save_progress(records)        # each answer costs money — never lose one
            log.info("  [%d/%d] %s %s -> %s%s", n, len(todo), job, path.name,
                     run["status"],
                     f" ({len(run.get('fields') or {})} fields)" if run.get("fields") else "")
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
