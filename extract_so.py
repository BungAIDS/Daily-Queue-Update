"""Extract the fields we want from a downloaded Sales Order PDF — test harness.

    python extract_so.py 421314 421388 421572     # by job number(s)
    python extract_so.py "C:\\path\\to\\file.pdf"

Pulls Design (+ type), Size, Arrangement, and the change-order history from
Additional Features/Notes. The Notes text loses its spaces in plain extraction
(CORRECTEDTOTALBILLING), so this reconstructs lines from word positions to put
the spaces back — and prints RAW vs SPACED so we can confirm the parse before
wiring it into the report + AI briefing.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

# A change-order note line: starts with CO#, C/O#, CO #, etc. then a digit.
CO_START = re.compile(r"^\s*C\s*/?\s*O\s*#?\s*\d", re.I)
DESIGN_HDR = re.compile(r"^\s*Design\s+(\S+)\s*(.*)$")
SPEC_CELL = re.compile(r"^(Design|Size|Arrangement)\b\s*(.*)$", re.I)


def _resolve(arg: str) -> list[Path]:
    p = Path(arg)
    if p.is_file():
        return [p]
    from config import SALES_ORDER_DIR
    folder = SALES_ORDER_DIR / arg
    if folder.is_dir():
        pdfs = sorted(folder.glob("*.pdf"))
        if pdfs:
            return pdfs
    raise SystemExit(f"No PDF found for {arg!r} (checked that path and {folder}).")


def _recon_lines(page, x_tol: float = 1.5) -> list[str]:
    """Rebuild text lines from word positions so spaces survive. Words are
    grouped into the same line by their vertical position, then ordered
    left-to-right and joined with single spaces."""
    words = page.extract_words(x_tolerance=x_tol, keep_blank_chars=False, use_text_flow=False)
    rows: dict[int, list] = {}
    for w in words:
        rows.setdefault(round(w["top"]), []).append(w)
    lines = []
    for top in sorted(rows):
        ws = sorted(rows[top], key=lambda w: w["x0"])
        lines.append(" ".join(w["text"] for w in ws))
    return lines


def _respace_value(value: str, recon_text: str) -> str:
    """Re-insert spaces table extraction glued out of a value, using the page's
    word-position reconstruction. Unmatched tokens are left as-is."""
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


def _spec_from_tables(tables) -> dict:
    fields: dict = {}
    for table in tables or []:
        for row in table:
            for cell in row:
                if not cell:
                    continue
                m = SPEC_CELL.match(cell.replace("\n", " ").strip())
                if m:
                    label, val = m.group(1).title(), m.group(2).strip()
                    if label not in fields and val:   # first wins (vaneaxial has a 2nd "Design")
                        fields[label] = val
    return fields


def parse(path: Path) -> dict:
    import pdfplumber
    res = {"design": "", "design_desc": "", "size": "", "arrangement": "",
           "header_co": None, "change_orders_raw": [], "change_orders_spaced": []}
    with pdfplumber.open(str(path)) as pdf:
        p1 = pdf.pages[0]
        for ln in (p1.extract_text() or "").splitlines()[:8]:
            if res["header_co"] is None:
                m = re.search(r"CO\s*#\s*(\d+)", ln)
                if m:
                    res["header_co"] = int(m.group(1))
            d = DESIGN_HDR.match(ln)
            if d and not res["design"]:
                res["design"], res["design_desc"] = d.group(1), d.group(2).strip()

        # Spec row can be pushed to a later page by a long Tag section — scan all.
        spec, recon = {}, ""
        for page in pdf.pages:
            found = _spec_from_tables(page.extract_tables())
            if found.get("Size") or found.get("Arrangement"):
                spec = found
                recon = "\n".join(_recon_lines(page))
                break
        res["size"] = _respace_value(spec.get("Size", ""), recon)
        res["arrangement"] = _respace_value(spec.get("Arrangement", "") or "N/A", recon)
        if not res["design"]:
            res["design"] = spec.get("Design", "")

        for page in pdf.pages:
            for ln in (page.extract_text() or "").splitlines():
                if CO_START.match(ln):
                    res["change_orders_raw"].append(ln.strip())
            for ln in _recon_lines(page):
                if CO_START.match(ln):
                    res["change_orders_spaced"].append(ln.strip())
    return res


def main() -> None:
    if len(sys.argv) < 2:
        raise SystemExit("Usage: python extract_so.py <jobnumber> [more...]   or a .pdf path")
    try:
        import pdfplumber  # noqa: F401
    except ImportError:
        raise SystemExit("pdfplumber isn't installed yet. Run:  pip install pdfplumber")

    paths: list[Path] = []
    for arg in sys.argv[1:]:
        paths.extend(_resolve(arg))

    for path in paths:
        r = parse(path)
        print(f"\n{'=' * 72}\n{path.name}\n{'=' * 72}")
        print(f"  Design       : {r['design']}  ({r['design_desc']})")
        print(f"  Size         : {r['size']}")
        print(f"  Arrangement  : {r['arrangement']}")
        print(f"  Header CO#   : {r['header_co']}")
        print(f"  Change orders ({len(r['change_orders_spaced'])}):")
        print("    -- RAW (no spaces) --")
        for ln in r["change_orders_raw"]:
            print(f"      {ln}")
        print("    -- SPACED (reconstructed) --")
        for ln in r["change_orders_spaced"]:
            print(f"      {ln}")


if __name__ == "__main__":
    main()
