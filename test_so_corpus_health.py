from __future__ import annotations

import json
from pathlib import Path
import sys
from unittest.mock import patch

import line_items
from sales_order_validation import DOCUMENT_KIND_SALES_ORDER, SalesOrderValidation
import so_corpus_health


def _record(pdf: Path, count: int = 1) -> dict:
    return {
        "job": "421430",
        "status": "ok",
        "so_pdf": str(pdf),
        "so_document_kind": DOCUMENT_KIND_SALES_ORDER,
        "so_source_type": "CBC_SalesOrder",
        "so_validation": "MATCH",
        "so_arrangement": "A/4",
        "line_item_count": count,
    }


def test_audit_reports_ok_job_missing_from_store(tmp: Path):
    pdf = tmp / "421430.pdf"
    pdf.write_bytes(b"pdf")
    store = tmp / "items.json"
    line_items.save_store({"jobs": {}, "ai_tags": {}}, store)
    audit = so_corpus_health.audit_corpus({"421430": _record(pdf)}, store_path=store)
    assert audit["missing_store_jobs"] == ["421430"]
    assert audit["missing_pdf_jobs"] == []
    assert not audit["healthy"]


def test_audit_reports_inaccessible_pdf_without_crashing(tmp: Path):
    pdf = tmp / "421430.pdf"
    store = tmp / "items.json"
    line_items.save_store({"jobs": {"421430": {"items": []}}, "ai_tags": {}}, store)
    records = {"421430": _record(pdf)}
    with patch("so_corpus_health.Path.is_file", side_effect=PermissionError("offline")):
        audit = so_corpus_health.audit_corpus(records, store_path=store)
    assert audit["missing_pdf_jobs"] == []
    assert audit["inaccessible_pdf_jobs"] == ["421430"]
    assert not audit["healthy"]
    assert any("share access" in action for action in audit["recommended_actions"])


def test_repair_uses_validated_archive_and_preserves_context(tmp: Path):
    pdf = tmp / "421430.pdf"
    pdf.write_bytes(b"a real archived source for hashing")
    store = tmp / "items.json"
    progress = tmp / "progress.json"
    line_items.save_store({"jobs": {}, "ai_tags": {}}, store)
    records = {"421430": _record(pdf)}
    parsed = {
        "line_items": [{"raw": "Motor", "norm": "MOTOR", "tags": ["MOTOR"]}],
        "arrangement": "A/9",
        "parts_only": True,
        "job_number": "421430-1",
    }
    validation = SalesOrderValidation(
        "421430", "421430", "MATCH", "header",
        document_kind=DOCUMENT_KIND_SALES_ORDER,
    )
    with (
        patch("so_corpus_health.validate_sales_order_pdf", return_value=validation),
        patch("so_corpus_health.parse_sales_order_pdf", return_value=parsed),
    ):
        result = so_corpus_health.repair_missing_line_items(
            records, destination=store, progress_path=progress)

    assert result["repaired"] == ["421430"]
    saved = line_items.load_store(store)["jobs"]["421430"]
    assert saved["arrangement"] == "A/9"
    assert saved["parts_only"] is True
    assert saved["job_number"] == "421430-1"
    assert saved["items"][0]["norm"] == "MOTOR"
    persisted = json.loads(progress.read_text(encoding="utf-8"))["421430"]
    assert persisted["line_item_count"] == 1
    assert len(persisted["so_pdf_sha256"]) == 64
    assert persisted["line_item_store_repaired_at"]


def test_repair_refuses_wrong_document(tmp: Path):
    pdf = tmp / "421430.pdf"
    pdf.write_bytes(b"wrong")
    store = tmp / "items.json"
    line_items.save_store({"jobs": {}, "ai_tags": {}}, store)
    records = {"421430": _record(pdf)}
    validation = SalesOrderValidation(
        "421430", "421999", "MISMATCH", "header",
        document_kind=DOCUMENT_KIND_SALES_ORDER,
    )
    with patch("so_corpus_health.validate_sales_order_pdf", return_value=validation):
        result = so_corpus_health.repair_missing_line_items(records, destination=store)
    assert result["repaired"] == []
    assert result["remaining_missing"] == ["421430"]
    assert "MISMATCH" in result["failures"][0]["reason"]


def test_quality_metrics_expose_parser_coverage_and_review_patterns(tmp: Path):
    store = {"jobs": {
        "421001": {"parser_version": "v2", "items": [
            {"raw": "Motor", "norm": "MOTOR", "tags": ["MOTOR"],
             "attributes": {"component": "MOTOR", "hp": "100"},
             "source": {"page": 1, "row": 2}},
            {"raw": "Mystery NaN wording", "norm": "MYSTERY WORDING", "tags": [],
             "attributes": {}, "review_flags": ["Uncategorized item"]},
        ]},
        "421002": {"items": []},
    }}
    quality = so_corpus_health.corpus_quality(store)
    assert quality["items"] == 2
    assert quality["tagged_rate"] == 0.5
    assert quality["structured_attribute_rate"] == 0.5
    assert quality["review_flagged_items"] == 1
    assert quality["nan_artifact_items"] == 1
    assert quality["jobs_with_source"] == 1
    assert quality["parser_versions"] == {"v2": 1, "legacy/unknown": 1}


def test_consolidation_keeps_freshest_record_and_fills_context(tmp: Path):
    main_path = tmp / "main.json"
    overlay_path = tmp / "overlay.json"
    main = {"jobs": {
        "421001": {"scanned_at": "2026-01-01T00:00:00", "items": [
            {"raw": "T Rails", "details": [], "tags": [], "attributes": {}},
        ]},
        "421002": {"scanned_at": "2026-01-03T00:00:00", "items": []},
    }, "ai_tags": {}}
    overlay = {"jobs": {
        "421001": {"scanned_at": "2026-01-02T00:00:00", "items": [
            {"raw": "Filter NaN Box L 100.00", "details": [], "tags": [], "attributes": {}},
        ]},
        "421003": {"scanned_at": "2026-01-02T00:00:00", "items": []},
    }, "ai_tags": {}}
    line_items.save_store(main, main_path)
    line_items.save_store(overlay, overlay_path)
    records = {"421001": {"so_arrangement": "A/9", "so_size": "270"}}

    result = so_corpus_health.consolidate_stores(
        records, main_path=main_path, overlay_path=overlay_path,
    )

    assert result["overlap_jobs"] == 1
    assert result["consolidated_jobs"] == 3
    saved = line_items.load_store(main_path)
    assert "NAN" not in saved["jobs"]["421001"]["items"][0]["raw"].upper()
    assert saved["jobs"]["421001"]["arrangement"] == "A/9"
    assert saved["jobs"]["421001"]["so_size"] == "270"
    assert saved["jobs"]["421001"]["normalizer_version"] == line_items.NORMALIZER_VERSION
    assert line_items.load_store(overlay_path)["jobs"] == {}
    assert len(result["backups"]) == 2


def main() -> int:
    passed = 0
    root = Path.cwd() / ".tmp_so_corpus_health_tests"
    root.mkdir(exist_ok=True)
    for name, function in sorted(globals().items()):
        if not name.startswith("test_") or not callable(function):
            continue
        target = root / name
        target.mkdir(exist_ok=True)
        function(target)
        print(f"  ok  {name}")
        passed += 1
    print(f"\n{passed} tests passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
