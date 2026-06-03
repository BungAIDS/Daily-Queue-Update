"""Dump a Sales Order PDF's text + tables so we can see what's extractable.

    pip install pdfplumber
    python dump_pdf.py 421388                 # by job number (looks in SALES_ORDER_DIR)
    python dump_pdf.py "C:\\path\\to\\file.pdf"  # or an explicit path

Prints each page's text and any detected tables. Paste the output back so we
can pick which fields to pull. Nothing leaves your machine except what you
choose to paste. If a page prints "(no extractable text)" the PDF is a scanned
image and we'd need a screenshot / OCR instead.
"""
from __future__ import annotations

import sys
from pathlib import Path


def _resolve(arg: str) -> list[Path]:
    """Accept an explicit file path, or a job number to look up under
    SALES_ORDER_DIR/<job>/ (handy for comparing several orders)."""
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


def main() -> None:
    if len(sys.argv) < 2:
        raise SystemExit('Usage: python dump_pdf.py <jobnumber>   OR   python dump_pdf.py "<path to .pdf>"')
    try:
        import pdfplumber
    except ImportError:
        raise SystemExit("pdfplumber isn't installed yet. Run:  pip install pdfplumber")

    for path in _resolve(sys.argv[1]):
        print(f"\n############################ {path}")
        with pdfplumber.open(str(path)) as pdf:
            print(f"############################ {len(pdf.pages)} page(s)")
            for i, page in enumerate(pdf.pages, 1):
                print(f"\n{'=' * 72}\nPAGE {i} — TEXT\n{'=' * 72}")
                print(page.extract_text() or "(no extractable text — likely a scanned image)")
                for t_i, table in enumerate(page.extract_tables(), 1):
                    print(f"\n--- PAGE {i} TABLE {t_i} ({len(table)} rows) ---")
                    for row in table:
                        print(" | ".join((c or "").replace("\n", " ").strip() for c in row))


if __name__ == "__main__":
    main()
