"""Construction / drive-run ("CBC_DriveRun") parsing.

A drive run is the construction document for a highly-custom fan. Only a small
fraction of orders have one, so its mere presence is a useful signal ("this fan
is highly custom"). `sales_orders.py` finds and downloads it the same way it
does the Sales Order — by the document's pid *type* prefix — and hands the PDF
here.

The exact field layout isn't known until we dump a real one (see
`discover_documents.py` + `dump_pdf.py`), so this does resilient best-effort
extraction now and gives a single place to pin down specific fields afterward:

    parse_drive_run_pdf(path) -> {
        "fields": {label: value, ...},   # every "Label: value" pair found
        "raw_lines": [...],              # first lines, space-correct
        "summary": "k1=v1; k2=v2; ...",  # compact one-liner for the report
    }

>>> AFTER DISCOVERY: add the real fields to FIELDS_OF_INTEREST below (and, if
    they deserve their own Excel columns, surface them in excel_writer). Until
    then the report shows a YES "Drive Run" flag (highly custom) and this
    summary, which already captures whatever labels the PDF carries.
"""
from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any, Dict, List

log = logging.getLogger(__name__)

# Labels we most expect a construction run to carry. This is a STARTING guess
# used only to order/prioritize the compact summary — extraction itself is
# generic, so unknown labels are still captured in "fields". Refine after the
# discovery dump shows the real headings.
FIELDS_OF_INTEREST = [
    "material", "construction", "gauge", "weld", "welder", "coating", "paint",
    "finish", "bearing", "shaft", "seal", "flange", "spark", "duty", "service",
]

# A "Label: value" or "Label : value" pair on one line. Label is short-ish text
# (so prose sentences with a stray colon don't masquerade as fields).
KV_RE = re.compile(r"^\s*([A-Za-z][A-Za-z0-9 /%#.\-]{1,34}?)\s*[:=]\s*(.+?)\s*$")


def _recon_lines(page, x_tol: float = 1.5) -> List[str]:
    """Rebuild text lines from word positions so spaces survive (mirrors the
    same helper in sales_orders/extract_so — plain extraction glues words)."""
    words = page.extract_words(x_tolerance=x_tol, keep_blank_chars=False, use_text_flow=False)
    rows: Dict[int, list] = {}
    for w in words:
        rows.setdefault(round(w["top"]), []).append(w)
    out = []
    for top in sorted(rows):
        ws = sorted(rows[top], key=lambda w: w["x0"])
        out.append(" ".join(w["text"] for w in ws))
    return out


def _kv_from_lines(lines: List[str]) -> Dict[str, str]:
    fields: Dict[str, str] = {}
    for ln in lines:
        m = KV_RE.match(ln)
        if m:
            label, val = m.group(1).strip(), m.group(2).strip()
            if label and val and label not in fields:  # first occurrence wins
                fields[label] = val
    return fields


def _kv_from_tables(tables) -> Dict[str, str]:
    """Two-column table rows (Label | value) are the other common layout."""
    fields: Dict[str, str] = {}
    for table in tables or []:
        for row in table:
            cells = [(c or "").replace("\n", " ").strip() for c in row]
            cells = [c for c in cells if c]
            if len(cells) == 2 and re.match(r"^[A-Za-z]", cells[0]) and len(cells[0]) <= 34:
                fields.setdefault(cells[0].rstrip(":"), cells[1])
    return fields


def _summarize(fields: Dict[str, str]) -> str:
    """Compact one-liner for the report — fields of interest first, then the
    rest, capped so it stays readable in a cell."""
    def rank(label: str) -> int:
        low = label.lower()
        for i, key in enumerate(FIELDS_OF_INTEREST):
            if key in low:
                return i
        return len(FIELDS_OF_INTEREST)
    ordered = sorted(fields.items(), key=lambda kv: (rank(kv[0]), kv[0]))
    return "; ".join(f"{k}={v}" for k, v in ordered[:12])


def parse_drive_run_pdf(path: str | Path) -> Dict[str, Any]:
    """Best-effort, never-raises extraction of a drive-run PDF.

    Returns generic key/value `fields`, the first `raw_lines`, a `summary`, and
    `text` — the FULL reconstructed text, so a caller can recognize a structured
    layout (e.g. the selection-program Qt Run) and parse it properly."""
    res: Dict[str, Any] = {"fields": {}, "raw_lines": [], "summary": "", "text": ""}
    try:
        import pdfplumber
    except ImportError:
        log.warning("pdfplumber not installed; cannot parse drive-run pdfs (pip install pdfplumber)")
        return res
    try:
        fields: Dict[str, str] = {}
        all_lines: List[str] = []
        with pdfplumber.open(str(path)) as pdf:
            for page in pdf.pages:
                lines = _recon_lines(page)
                all_lines.extend(lines)
                for k, v in _kv_from_lines(lines).items():
                    fields.setdefault(k, v)
                for k, v in _kv_from_tables(page.extract_tables()).items():
                    fields.setdefault(k, v)
        res["fields"] = fields
        res["raw_lines"] = all_lines[:40]
        res["text"] = "\n".join(all_lines)
        res["summary"] = _summarize(fields)
    except Exception as e:  # noqa: BLE001 - never let a bad pdf fail the run
        log.warning("Could not parse drive-run pdf %s: %s", path, e)
    return res


def main() -> None:
    """`python drive_run.py <path-or-job#>` — show what we can pull from a drive
    run, so you can confirm the fields before they're wired into the report."""
    import sys
    if len(sys.argv) < 2:
        raise SystemExit("Usage: python drive_run.py <path to drive-run .pdf | job#>")
    arg = sys.argv[1]
    path = Path(arg)
    if not path.is_file():
        from config import DRIVE_RUN_DIR
        pdfs = sorted((DRIVE_RUN_DIR / arg).glob("*.pdf")) if (DRIVE_RUN_DIR / arg).is_dir() else []
        if not pdfs:
            raise SystemExit(f"No drive-run PDF found for {arg!r} (checked that path and {DRIVE_RUN_DIR / arg}).")
        path = pdfs[-1]
    r = parse_drive_run_pdf(path)
    print(f"\n{'=' * 72}\n{path}\n{'=' * 72}")
    print(f"Fields found ({len(r['fields'])}):")
    for k, v in r["fields"].items():
        print(f"  {k:<28} {v}")
    print(f"\nSummary: {r['summary']}")
    print("\n--- first reconstructed lines (for spotting fields to add) ---")
    for ln in r["raw_lines"]:
        print(f"  {ln}")
    print("\nPaste this back so we can pin down which drive-run fields to pull.")


if __name__ == "__main__":
    main()
