"""Regression coverage for the no-Order-Verification-Report invariant."""
from __future__ import annotations

import json
import os
import sys
import tempfile
from pathlib import Path
from unittest.mock import patch

import line_items
import live_master
import order_verification_cleanup as cleanup
from test_sales_order_validation import _mini_pdf


def _report(path: Path, job: str) -> None:
    _mini_pdf([
        "Order Verification Report",
        "Order Cust PO Ship Via Package Prepaid Date Order Terms Verification Date",
        f"{job} CBC PO UPS",
    ], path)


def _sales_order(path: Path, job: str) -> None:
    _mini_pdf([
        "Chicago Blower Corporation Sales Order",
        "Order# RepRef#",
        f"{job} 987",
    ], path)


def test_cleanup_quarantines_reports_and_removes_every_derived_record(tmp: Path):
    active = tmp / "SALES ORDERS FOR DAILY QUEUE"
    backlog = tmp / "backlog"
    snapshots = tmp / "snapshots"
    report_pdf = active / "422027" / "422027 - Sales Order (original).pdf"
    true_pdf = active / "422018" / "422018 - Sales Order (original).pdf"
    _report(report_pdf, "422027")
    _sales_order(true_pdf, "422018")

    progress_path = backlog / "backfill_progress.json"
    main_items_path = backlog / "line_items.json"
    overlay_path = backlog / "backfill_line_items.json"
    master_path = snapshots / "live_master.json"
    state_path = snapshots / "live_state_2026-07-13.json"
    queue_path = snapshots / "queue_2026-07-13.json"
    audit_path = backlog / "order_verification_cleanup.json"
    backlog.mkdir(parents=True)
    snapshots.mkdir(parents=True)

    report_data = {
        "job": "422027",
        "status": "ok",
        "so_pdf": str(report_pdf),
        "so_document_kind": "ORDER_VERIFICATION",
        "so_source_type": "CS_SalesOrder",
        "co_number": 2,
        "so_size": "270",
        "line_items": [{"raw": "wrong"}],
        "line_item_tags": "MOTOR",
        "dwg_reuse": [{"job": "400001"}],
        "drive_run_pdf": "keep-run.pdf",
        "so_imi": "keep-imi.pdf",
        "backfill_attempts": 4,
    }
    true_data = {
        "job": "422018",
        "status": "ok",
        "so_pdf": str(true_pdf),
        "so_document_kind": "SALES_ORDER",
        "so_source_type": "CBC_SalesOrder",
        "co_number": 0,
        "so_size": "33",
        "line_items": [{"raw": "good"}],
    }
    progress_path.write_text(json.dumps({
        "422027": dict(report_data),
        "422018": dict(true_data),
    }), encoding="utf-8")
    for path in (main_items_path, overlay_path):
        line_items.save_store({
            "jobs": {
                "422027": {"so_pdf": str(report_pdf), "items": [{"raw": "wrong"}]},
                "422018": {"so_pdf": str(true_pdf), "items": [{"raw": "good"}]},
            },
            "ai_tags": {},
        }, path)
    master_path.write_text(json.dumps({
        "orders": {
            "422027": {"on_queue": True, "job": dict(report_data)},
            "422018": {"on_queue": False, "job": dict(true_data)},
        },
    }), encoding="utf-8")
    state_path.write_text(json.dumps({
        "422027": {"present": True, "enriched": True, "job": dict(report_data)},
    }), encoding="utf-8")
    queue_path.write_text(json.dumps([dict(report_data)]), encoding="utf-8")

    with (
        patch.object(cleanup, "SALES_ORDER_DIR", active),
        patch.object(cleanup, "SNAPSHOT_DIR", snapshots),
        patch.object(cleanup, "PROGRESS_PATH", progress_path),
        patch.object(cleanup, "CLEANUP_AUDIT_PATH", audit_path),
        patch.object(cleanup, "SCAN_CACHE_PATH", backlog / "scan_cache.json"),
        patch.object(live_master, "MASTER_PATH", master_path),
        patch("line_items.store_path", return_value=main_items_path),
        patch("line_items.backfill_store_path", return_value=overlay_path),
    ):
        counts = cleanup.run(max_workers=1)

    assert counts == {
        "reports_quarantined": 1,
        "progress_invalidated": 1,
        "line_items_removed": 2,
        "master_invalidated": 1,
        "states_invalidated": 1,
        "snapshots_invalidated": 1,
    }
    assert not report_pdf.exists()
    assert true_pdf.exists()
    quarantined = list(
        (tmp / "SALES ORDERS FOR DAILY QUEUE QUARANTINE").rglob(report_pdf.name)
    )
    assert len(quarantined) == 1

    progress = json.loads(progress_path.read_text(encoding="utf-8"))
    rejected = progress["422027"]
    assert rejected["status"] == "order-verification-removed"
    assert "so_pdf" not in rejected and "co_number" not in rejected
    assert "line_items" not in rejected and "dwg_reuse" not in rejected
    assert "backfill_attempts" not in rejected
    assert rejected["drive_run_pdf"] == "keep-run.pdf"
    assert rejected["so_imi"] == "keep-imi.pdf"
    assert progress["422018"]["so_pdf"] == str(true_pdf)

    for path in (main_items_path, overlay_path):
        jobs = line_items.load_store(path)["jobs"]
        assert set(jobs) == {"422018"}

    master = json.loads(master_path.read_text(encoding="utf-8"))
    master_rejected = master["orders"]["422027"]["job"]
    assert "so_pdf" not in master_rejected and "line_items" not in master_rejected
    assert master_rejected["drive_run_pdf"] == "keep-run.pdf"
    assert master["orders"]["422018"]["job"]["so_pdf"] == str(true_pdf)

    state = json.loads(state_path.read_text(encoding="utf-8"))
    assert state["422027"]["enriched"] is False
    assert "so_pdf" not in state["422027"]["job"]
    queue = json.loads(queue_path.read_text(encoding="utf-8"))
    assert "so_pdf" not in queue[0]
    assert audit_path.exists()


def test_scan_cache_skips_unchanged_clean_pdfs(tmp: Path):
    # The sweep re-parses a PDF only when it is new or changed: a clean file
    # whose mtime+size match the cache from the last pass is skipped, so only
    # the first pass sweeps the whole bank (the pass that used to hold the
    # cleanup lock for the length of a full Z:-drive re-validation).
    active = tmp / "SALES ORDERS FOR DAILY QUEUE"
    cache_path = tmp / "backlog" / "scan_cache.json"
    true_pdf = active / "422018" / "422018 - Sales Order (original).pdf"
    _sales_order(true_pdf, "422018")

    calls = []
    real = cleanup.validate_sales_order_pdf

    def counting(path, job):
        calls.append(str(path))
        return real(path, job)

    with (
        patch.object(cleanup, "SALES_ORDER_DIR", active),
        patch.object(cleanup, "SCAN_CACHE_PATH", cache_path),
        patch.object(cleanup, "validate_sales_order_pdf", counting),
    ):
        assert cleanup.quarantine_active_reports(max_workers=1) == {}
        assert len(calls) == 1                     # first pass validates it
        assert cleanup.quarantine_active_reports(max_workers=1) == {}
        assert len(calls) == 1                     # unchanged -> cache hit, no re-parse
        os.utime(true_pdf, ns=(1, 1))              # changed -> re-validated
        assert cleanup.quarantine_active_reports(max_workers=1) == {}
        assert len(calls) == 2
    cache = json.loads(cache_path.read_text(encoding="utf-8"))
    assert len(cache) == 1                         # the clean survivor is cached


def test_run_with_zero_lock_timeout_raises_instead_of_waiting(tmp: Path):
    # The watcher passes lock_timeout=0: when another process is mid-cleanup it
    # must get TimeoutError immediately (and skip), never sit on the lock.
    from process_lock import exclusive_file_lock

    audit_path = tmp / "backlog" / "order_verification_cleanup.json"
    audit_path.parent.mkdir(parents=True, exist_ok=True)
    lock_path = audit_path.parent / ".process_locks" / (audit_path.name + ".lock")
    with patch.object(cleanup, "CLEANUP_AUDIT_PATH", audit_path):
        with exclusive_file_lock(lock_path, label="test holder"):
            try:
                cleanup.run(max_workers=1, lock_timeout=0)
            except TimeoutError:
                pass
            else:
                raise AssertionError("expected TimeoutError while the lock is held")


def main() -> int:
    passed = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as directory:
                fn(Path(directory))
            print(f"  ok  {name}")
            passed += 1
    print(f"\n{passed} tests passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
