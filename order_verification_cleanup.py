"""Quarantine Order Verification Reports and remove their derived data.

CBC exposes ``CS_SalesOrder`` beside the real ``CBC_SalesOrder`` document.
The former is an Order Verification Report, never a Sales Order.  This module
repairs files and stores created before that distinction was enforced.  PDFs
are moved to quarantine, never deleted.
"""
from __future__ import annotations

import json
import logging
import re
from concurrent.futures import ThreadPoolExecutor
from dataclasses import replace
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable

import line_items
import live_master
from config import BACKLOG_DIR, SALES_ORDER_DIR, SNAPSHOT_DIR
from process_lock import data_file_lock
from sales_order_validation import (
    DOCUMENT_KIND_ORDER_VERIFICATION,
    clear_sales_order_data,
    is_order_verification_record,
    normalize_order,
    quarantine_candidate,
    validate_sales_order_pdf,
)

log = logging.getLogger(__name__)

PROGRESS_PATH = BACKLOG_DIR / "backfill_progress.json"
CLEANUP_AUDIT_PATH = BACKLOG_DIR / "order_verification_cleanup.json"
_JOB_FOLDER_RE = re.compile(r"^[0-9]{4,7}[A-Za-z]?$")


def _read_json(path: Path) -> Any:
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as exc:
        log.warning("Could not inspect %s during report cleanup (%s)", path, exc)
        return None


def _save_json_unlocked(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(value, indent=2, default=str), encoding="utf-8")
    tmp.replace(path)


def _path_key(value: str | Path) -> str:
    return str(value or "").strip().replace("/", "\\").casefold()


def _metadata_says_report(record: Dict[str, Any]) -> bool:
    return is_order_verification_record(record)


def _record_is_report_backed(record: Dict[str, Any], report_paths: set[str]) -> bool:
    pdf = _path_key(record.get("so_pdf") or "")
    return _metadata_says_report(record) or bool(pdf and pdf in report_paths)


def _collect_metadata_paths(
    value: Any, found: set[str], report_jobs: set[str]
) -> None:
    if isinstance(value, dict):
        if _metadata_says_report(value):
            path = _path_key(value.get("so_pdf") or "")
            if path:
                found.add(path)
            job = normalize_order(value.get("job") or "")
            if job:
                report_jobs.add(job)
        for child in value.values():
            _collect_metadata_paths(child, found, report_jobs)
    elif isinstance(value, list):
        for child in value:
            _collect_metadata_paths(child, found, report_jobs)


def _source_json_paths() -> list[Path]:
    paths = [PROGRESS_PATH, live_master.MASTER_PATH]
    try:
        paths.extend(SNAPSHOT_DIR.glob("live_state_*.json"))
        paths.extend(SNAPSHOT_DIR.glob("queue_*.json"))
    except OSError:
        pass
    seen: set[str] = set()
    result = []
    for path in paths:
        key = _path_key(path)
        if key not in seen:
            seen.add(key)
            result.append(path)
    return result


def _active_pdf_rows(root: Path) -> list[tuple[Path, str]]:
    rows: list[tuple[Path, str]] = []
    if not root.is_dir():
        log.warning("Sales Order bank is not reachable for report cleanup: %s", root)
        return rows
    try:
        folders = list(root.iterdir())
    except OSError as exc:
        log.warning("Could not list Sales Order bank %s (%s)", root, exc)
        return rows
    for folder in folders:
        if not folder.is_dir() or not _JOB_FOLDER_RE.fullmatch(folder.name.strip()):
            continue
        job = normalize_order(folder.name)
        try:
            rows.extend((pdf, job) for pdf in folder.glob("*.pdf"))
        except OSError:
            continue
    return rows


def _quarantined_report_evidence(root: Path) -> tuple[set[str], set[str]]:
    """Original paths recorded when earlier runs quarantined report PDFs."""
    quarantine_root = root.parent / f"{root.name} QUARANTINE"
    if not quarantine_root.is_dir():
        return set(), set()
    found: set[str] = set()
    jobs: set[str] = set()
    try:
        metadata_files = quarantine_root.rglob("*.validation.json")
        for path in metadata_files:
            metadata = _read_json(path)
            if not isinstance(metadata, dict):
                continue
            if (
                str(metadata.get("document_kind") or "").upper()
                != DOCUMENT_KIND_ORDER_VERIFICATION
            ):
                continue
            original = _path_key(metadata.get("original_path") or "")
            if original:
                found.add(original)
            job = normalize_order(metadata.get("expected_order") or "")
            if job:
                jobs.add(job)
    except OSError as exc:
        log.warning("Could not inspect Sales Order quarantine receipts (%s)", exc)
    return found, jobs


def quarantine_active_reports(
    root: Path | None = None,
    max_workers: int = 8,
    known_report_paths: set[str] | None = None,
) -> Dict[str, Dict[str, str]]:
    """Move every readable Order Verification Report out of the active bank."""
    root = Path(root or SALES_ORDER_DIR)
    rows = _active_pdf_rows(root)
    if not rows:
        return {}

    def classify(row: tuple[Path, str]):
        path, job = row
        return row, validate_sales_order_pdf(path, job)

    workers = max(1, min(max_workers, len(rows)))
    reports = []
    with ThreadPoolExecutor(max_workers=workers) as pool:
        for (path, job), validation in pool.map(classify, rows):
            known = _path_key(path) in (known_report_paths or set())
            named = "orderverificationreportviewer" in path.name.casefold()
            if (
                validation.document_kind == DOCUMENT_KIND_ORDER_VERIFICATION
                or known
                or named
            ):
                if validation.document_kind != DOCUMENT_KIND_ORDER_VERIFICATION:
                    validation = replace(
                        validation,
                        document_kind=DOCUMENT_KIND_ORDER_VERIFICATION,
                    )
                reports.append((path, job, validation))

    moved: Dict[str, Dict[str, str]] = {}
    for path, job, validation in reports:
        original = str(path)
        try:
            target = quarantine_candidate(
                path,
                path,
                validation,
                bucket="order-verification-reports",
            )
        except OSError as exc:
            log.warning("Could not quarantine Order Verification Report %s (%s)", path, exc)
            continue
        moved[_path_key(original)] = {
            "job": job,
            "original_path": original,
            "quarantine_path": str(target),
        }
    return moved


def _clean_progress(report_paths: set[str], when: str) -> int:
    if not PROGRESS_PATH.exists():
        return 0
    with data_file_lock(PROGRESS_PATH, label="backfill report cleanup"):
        records = _read_json(PROGRESS_PATH)
        if not isinstance(records, dict):
            return 0
        changed = 0
        for job, record in records.items():
            if not isinstance(record, dict) or not _record_is_report_backed(record, report_paths):
                continue
            clear_sales_order_data(record, when)
            record["job"] = str(record.get("job") or job)
            record["status"] = "order-verification-removed"
            record["retry_reason"] = "Order Verification Report removed; fetch a true Sales Order."
            record["scanned_at"] = when
            record.pop("backfill_attempts", None)
            changed += 1
        if changed:
            _save_json_unlocked(PROGRESS_PATH, records)
        return changed


def _clean_line_item_store(
    path: Path, report_paths: set[str], report_jobs: set[str]
) -> int:
    if not path.exists():
        return 0
    with data_file_lock(path, label="line-items report cleanup"):
        store = line_items.load_store(path)
        jobs = store.get("jobs") or {}
        remove = []
        for job, record in jobs.items():
            pdf = _path_key((record or {}).get("so_pdf") or "")
            if pdf in report_paths or (not pdf and normalize_order(job) in report_jobs):
                remove.append(job)
        for job in remove:
            jobs.pop(job, None)
        if remove:
            line_items.save_store(store, path)
        return len(remove)


def _clean_master(report_paths: set[str], when: str) -> int:
    path = live_master.MASTER_PATH
    if not path.exists():
        return 0
    with data_file_lock(path, label="live master report cleanup"):
        master = live_master.load_master()
        changed = 0
        for entry in (master.get("orders") or {}).values():
            record = entry.get("job") or {}
            if _record_is_report_backed(record, report_paths):
                clear_sales_order_data(record, when)
                entry["job"] = record
                changed += 1
        if changed:
            live_master._save_master_unlocked(master)
        return changed


def _clean_live_state(path: Path, report_paths: set[str], when: str) -> int:
    with data_file_lock(path, label="live-state report cleanup"):
        state = _read_json(path)
        if not isinstance(state, dict):
            return 0
        changed = 0
        for entry in state.values():
            if not isinstance(entry, dict):
                continue
            record = entry.get("job") or {}
            if _record_is_report_backed(record, report_paths):
                clear_sales_order_data(record, when)
                entry["job"] = record
                entry["enriched"] = False
                changed += 1
        if changed:
            _save_json_unlocked(path, state)
        return changed


def _clean_queue_snapshot(path: Path, report_paths: set[str], when: str) -> int:
    with data_file_lock(path, label="queue snapshot report cleanup"):
        jobs = _read_json(path)
        if not isinstance(jobs, list):
            return 0
        changed = 0
        for record in jobs:
            if isinstance(record, dict) and _record_is_report_backed(record, report_paths):
                clear_sales_order_data(record, when)
                changed += 1
        if changed:
            _save_json_unlocked(path, jobs)
        return changed


def _unique_paths(paths: Iterable[Path]) -> list[Path]:
    seen: set[str] = set()
    result = []
    for path in paths:
        key = _path_key(path)
        if key not in seen:
            seen.add(key)
            result.append(path)
    return result


def run(max_workers: int = 8) -> Dict[str, int]:
    """Enforce the no-verification-report invariant across files and stores."""
    with data_file_lock(CLEANUP_AUDIT_PATH, label="Order Verification cleanup"):
        source_paths = _source_json_paths()
        report_paths: set[str] = set()
        report_jobs: set[str] = set()
        for path in source_paths:
            _collect_metadata_paths(_read_json(path), report_paths, report_jobs)

        quarantined_paths, quarantined_jobs = _quarantined_report_evidence(
            Path(SALES_ORDER_DIR)
        )
        report_paths.update(quarantined_paths)
        report_jobs.update(quarantined_jobs)

        moved = quarantine_active_reports(
            max_workers=max_workers,
            known_report_paths=report_paths,
        )
        report_paths.update(moved)
        report_jobs.update(row["job"] for row in moved.values())
        when = datetime.now().isoformat(timespec="seconds")

        counts = {
            "reports_quarantined": len(moved),
            "progress_invalidated": _clean_progress(report_paths, when),
            "line_items_removed": 0,
            "master_invalidated": _clean_master(report_paths, when),
            "states_invalidated": 0,
            "snapshots_invalidated": 0,
        }

        stores = _unique_paths([
            line_items.store_path(),
            line_items.backfill_store_path(),
        ])
        counts["line_items_removed"] = sum(
            _clean_line_item_store(path, report_paths, report_jobs) for path in stores
        )

        for path in source_paths:
            if path.name.startswith("live_state_"):
                counts["states_invalidated"] += _clean_live_state(path, report_paths, when)
            elif path.name.startswith("queue_"):
                counts["snapshots_invalidated"] += _clean_queue_snapshot(
                    path, report_paths, when
                )

        if any(counts.values()):
            audit = {
                "cleaned_at": when,
                **counts,
                "quarantined_files": list(moved.values()),
            }
            _save_json_unlocked(CLEANUP_AUDIT_PATH, audit)
            log.warning(
                "Order Verification cleanup: %d PDF(s) quarantined; %d progress, "
                "%d line-item, %d master, %d state/snapshot record(s) invalidated.",
                counts["reports_quarantined"],
                counts["progress_invalidated"],
                counts["line_items_removed"],
                counts["master_invalidated"],
                counts["states_invalidated"] + counts["snapshots_invalidated"],
            )
        return counts


def changed(counts: Dict[str, int]) -> bool:
    return any(int(value or 0) for value in counts.values())
