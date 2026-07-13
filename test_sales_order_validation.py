from __future__ import annotations

import asyncio
import inspect
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, patch

from sales_order_validation import (
    SalesOrderAcceptance,
    SalesOrderValidation,
    accept_existing,
    extract_internal_order,
    finalize_candidate,
    modal_text_matches_job,
    validate_sales_order_pdf,
)


def _mini_pdf(lines: list[str], path: Path) -> None:
    content = ["BT", "/F1 10 Tf"]
    y = 760
    for line in lines:
        escaped = str(line).replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
        content.append(f"1 0 0 1 40 {y} Tm ({escaped}) Tj")
        y -= 14
    content.append("ET")
    stream = "\n".join(content).encode("latin-1", "replace")
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R "
        b"/Resources << /Font << /F1 5 0 R >> >> >>",
        b"<< /Length " + str(len(stream)).encode() + b" >>\nstream\n" + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    output = bytearray(b"%PDF-1.4\n")
    offsets = []
    for index, body in enumerate(objects, start=1):
        offsets.append(len(output))
        output += f"{index} 0 obj\n".encode() + body + b"\nendobj\n"
    xref = len(output)
    output += f"xref\n0 {len(objects) + 1}\n".encode() + b"0000000000 65535 f \n"
    for offset in offsets:
        output += f"{offset:010d} 00000 n \n".encode()
    output += (
        f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\n"
        f"startxref\n{xref}\n%%EOF"
    ).encode()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(bytes(output))


def test_extracts_standard_and_order_verification_layouts():
    assert extract_internal_order("Order# RepRef#\n421314 987") == ("421314", "header")
    hdx = (
        "Order# Job# Rep. Ref. # Customer P.O. # Fan S/N\n"
        "420402 251535 329-6-1005, rev.3 L227508"
    )
    assert extract_internal_order(hdx) == ("420402", "order-job-table")
    verification = (
        "Order Cust PO Ship Via Package Prepaid Date Order Terms Verification Date\n"
        "421899 CBC PO UPS"
    )
    assert extract_internal_order(verification) == ("421899", "order-verification-table")


def test_modal_text_requires_the_requested_job():
    assert modal_text_matches_job("WORK CENTER DETAIL (417467-0) Documents", "417467")
    assert modal_text_matches_job("417467 - Sales Order.pdf", "417467")
    assert not modal_text_matches_job("WORK CENTER DETAIL (421851-0)", "400083")
    assert not modal_text_matches_job("Order 1417467", "417467")


def test_validate_pdf_matches_internal_order(tmp_path: Path):
    pdf = tmp_path / "421314 - Sales Order.pdf"
    _mini_pdf(["Order# RepRef#", "421314 987"], pdf)
    validation = validate_sales_order_pdf(pdf, "421314")
    assert validation.matched
    assert validation.internal_order == "421314"


def test_valid_staged_pdf_is_promoted(tmp_path: Path):
    root = tmp_path / "SALES ORDERS FOR DAILY QUEUE"
    destination = root / "421314" / "421314 - Sales Order (original).pdf"
    candidate = tmp_path / "pending" / destination.name
    _mini_pdf(["Order# RepRef#", "421314 987"], candidate)

    accepted = finalize_candidate(candidate, destination, "421314")

    assert accepted.path == str(destination)
    assert accepted.validation.matched
    assert destination.exists()
    assert not candidate.exists()


def test_mismatched_pdf_is_quarantined_not_deleted(tmp_path: Path):
    root = tmp_path / "SALES ORDERS FOR DAILY QUEUE"
    destination = root / "400083" / "400083 - Sales Order CO1.pdf"
    candidate = tmp_path / "pending" / destination.name
    _mini_pdf(["Order# RepRef#", "421851 987"], candidate)

    accepted = finalize_candidate(candidate, destination, "400083")

    assert accepted.path is None
    assert accepted.validation.status == "MISMATCH"
    quarantined = Path(accepted.quarantine_path)
    assert quarantined.exists()
    assert quarantined.with_suffix(quarantined.suffix + ".validation.json").exists()
    assert not candidate.exists()
    assert not destination.exists()


def test_bad_existing_pdf_leaves_active_bank(tmp_path: Path):
    root = tmp_path / "SALES ORDERS FOR DAILY QUEUE"
    destination = root / "400083" / "400083 - Sales Order CO1.pdf"
    _mini_pdf(["Order# RepRef#", "421851 987"], destination)

    accepted = accept_existing(destination, "400083")

    assert accepted is not None and accepted.path is None
    assert Path(accepted.quarantine_path).exists()
    assert not destination.exists()


def test_line_item_scan_uses_latest_verified_sales_order(tmp_path: Path):
    from line_items_scan import _latest_so_pdf

    folder = tmp_path / "421314"
    _mini_pdf(["Order# RepRef#", "421314 1"], folder / "421314 - Sales Order (original).pdf")
    _mini_pdf(["Order# RepRef#", "421314 1"], folder / "421314 - Sales Order CO2.pdf")
    _mini_pdf(["Order# RepRef#", "421999 1"], folder / "421314 - Sales Order CO3.pdf")
    _mini_pdf(["Order# RepRef#", "421314 1"], folder / "unrelated.pdf")

    path, co_number = _latest_so_pdf(folder)
    assert path.name == "421314 - Sales Order CO2.pdf"
    assert co_number == 2


def test_async_backfill_rejection_cannot_save_co_or_so_data():
    import backfill_orders

    rejected = SalesOrderAcceptance(
        None,
        SalesOrderValidation("400083", "421851", "MISMATCH", "header"),
        r"Z:\DAG\SALES ORDERS FOR DAILY QUEUE QUARANTINE\rejected.pdf",
    )
    docs = [("/download", {"type": backfill_orders.SO_TYPE, "rev": 2})]

    async def run():
        with (
            patch("backfill_orders.open_order_detail_async", new=AsyncMock(return_value=True)),
            patch("backfill_orders._collect_docs_async", new=AsyncMock(return_value=docs)),
            patch("backfill_orders._download_sales_order_async", new=AsyncMock(return_value=rejected)),
            patch("backfill_orders._close_modal_async", new=AsyncMock()),
        ):
            return await backfill_orders.process_one_async(
                SimpleNamespace(url="https://example.invalid"), object(), "400083"
            )

    record = asyncio.run(run())
    assert record["status"] == "error"
    assert record["so_validation"] == "MISMATCH"
    assert record["so_internal_order"] == "421851"
    assert "co_number" not in record
    assert "so_pdf" not in record
    assert "line_item_count" not in record


def main() -> int:
    passed = 0
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as directory:
        temp_root = Path(directory)
        for name, function in sorted(globals().items()):
            if not name.startswith("test_") or not callable(function):
                continue
            parameters = inspect.signature(function).parameters
            function(temp_root / name) if "tmp_path" in parameters else function()
            print(f"  ok  {name}")
            passed += 1
    print(f"\n{passed} tests passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
