"""Dump a PDF's text + tables so we can see exactly what's extractable.

    pip install pdfplumber
    python dump_pdf.py "C:\\Users\\dgroth\\Documents\\DailyQueue\\so_discovery\\421314_sales_order.pdf"

Prints each page's text and any detected tables. Paste the output back so we
can pick which fields to pull. Nothing leaves your machine except what you
choose to paste. If a page prints "(no extractable text)" the PDF is a scanned
image and we'd need a screenshot / OCR instead.
"""
from __future__ import annotations

import sys


def main() -> None:
    if len(sys.argv) < 2:
        raise SystemExit('Usage: python dump_pdf.py "<path to the .pdf>"')
    path = sys.argv[1]
    try:
        import pdfplumber
    except ImportError:
        raise SystemExit("pdfplumber isn't installed yet. Run:  pip install pdfplumber")

    with pdfplumber.open(path) as pdf:
        print(f"=== {path}")
        print(f"=== {len(pdf.pages)} page(s)")
        for i, page in enumerate(pdf.pages, 1):
            print(f"\n{'=' * 72}\nPAGE {i} — TEXT\n{'=' * 72}")
            print(page.extract_text() or "(no extractable text — likely a scanned image)")
            for t_i, table in enumerate(page.extract_tables(), 1):
                print(f"\n--- PAGE {i} TABLE {t_i} ({len(table)} rows) ---")
                for row in table:
                    print(" | ".join((c or "").replace("\n", " ").strip() for c in row))


if __name__ == "__main__":
    main()
