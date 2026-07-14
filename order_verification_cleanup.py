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
    DOCUMENT_KIND_SALES_ORDER,
    clear_sales_order_data,
    is_order_verification_record,
    is_true_sales_order_record,
    normalize_order,
    quarantine_candidate,
    validate_sales_order_pdf,
)

log = logging.getLogger(__name__)

PROGRESS_PATH = BACKLOG_DIR / "backfill_progress.json"
CLEANUP_AUDIT_PATH = BACKLOG_DIR / "order_verification_cleanup.json"
# Per-machine scan cache: {path key: [mtime_ns, size]} for every bank PDF that
# validated CLEAN on the last pass. Unchanged entries skip the (expensive,
# pdfplumber-over-Z:) re-validation, so only the FIRST pass sweeps the whole
# bank; later passes touch just new/changed files. Deliberately not published
# by data_push — it's local file-state, meaningless on another machine.
SCAN_CACHE_PATH = BACKLOG_DIR / "order_verification_scan_cache.json"
# Written when a full pass finishes with NOTHING to fix, removed when a pass
# finds work: the signal that the historical repair (the migration this module
# exists for) is complete on this machine. startup_check() skips the sweep
# while it exists — a NEW document never needs a database sweep to be trusted,
# because every fetch validates its own PDF (sales_order_validation) before
# saving/parsing it.
CLEAN_MARKER_PATH = BACKLOG_DIR / "order_verification_clean_marker.json"
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


def _metadata_says_verified_sales_order(record: Dict[str, Any]) -> bool:
    return (
        is_true_sales_order_record(record)
        and str(record.get("so_validation") or "").upper() == "MATCH"
        and bool(str(record.get("so_verified_at") or "").strip())
    )


def _record_is_report_backed(record: Dict[str, Any], report_paths: set[str]) -> bool:
    pdf = _path_key(record.get("so_pdf") or "")
    return _metadata_says_report(record) or bool(pdf and pdf in report_paths)


def _collect_metadata_paths(
    value: Any,
    found: set[str],
    report_jobs: set[str],
    trusted_sales_order_paths: set[str] | None = None,
) -> None:
    if isinstance(value, dict):
        if _metadata_says_report(value):
            path = _path_key(value.get("so_pdf") or "")
            if path:
                found.add(path)
            job = normalize_order(value.get("job") or "")
            if job:
                report_jobs.add(job)
        elif _metadata_says_verified_sales_order(value):
            path = _path_key(value.get("so_pdf") or "")
            if path and trusted_sales_order_paths is not None:
                trusted_sales_order_paths.add(path)
        for child in value.values():
            _collect_metadata_paths(
                child, found, report_jobs, trusted_sales_order_paths
            )
    elif isinstance(value, list):
        for child in value:
            _collect_metadata_paths(
                child, found, report_jobs, trusted_sales_order_paths
            )


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


def _pdf_signature(path: Path) -> "list | None":
    """A cheap change signature for a bank PDF ([mtime_ns, size]); None when it
    can't be stat'ed (treated as changed, so it gets re-validated)."""
    try:
        st = path.stat()
        return [st.st_mtime_ns, st.st_size]
    except OSError:
        return None


def quarantine_active_reports(
    root: Path | None = None,
    max_workers: int = 8,
    known_report_paths: set[str] | None = None,
) -> Dict[str, Dict[str, str]]:
    """Move every readable Order Verification Report out of the active bank,
    while allowing a corrected Sales Order at an old report path (current PDF
    contents outrank stale quarantine history).

    Validation is a pdfplumber parse per file — expensive over the Z: drive —
    so it only touches NEW or CHANGED PDFs: a file whose mtime+size still match
    the scan cache validated clean on an earlier pass and is skipped. Only the
    first pass (or a wiped cache) sweeps the whole bank."""
    root = Path(root or SALES_ORDER_DIR)
    rows = _active_pdf_rows(root)
    if not rows:
        return {}
    cache = _read_json(SCAN_CACHE_PATH)
    cache = cache if isinstance(cache, dict) else {}
    known = known_report_paths or set()

    survivors: Dict[str, list] = {}          # clean PDFs -> the next pass's cache
    pending: list = []                       # (path, job, signature) to validate
    for path, job in rows:
        key = _path_key(path)
        sig = _pdf_signature(path)
        named = "orderverificationreportviewer" in path.name.casefold()
        if sig is not None and not named and key not in known and cache.get(key) == sig:
            survivors[key] = sig             # validated clean before, unchanged since
            continue
        pending.append((path, job, sig))
    if len(pending) > 25:
        log.info(
            "Order Verification sweep: validating %d Sales Order PDF(s)%s",
            len(pending),
            " — first pass over the whole bank; this can take a while."
            if not cache else ".",
        )

    def classify(row):
        path, job, _sig = row
        return row, validate_sales_order_pdf(path, job)

    reports = []
    if pending:
        workers = max(1, min(max_workers, len(pending)))
        with ThreadPoolExecutor(max_workers=workers) as pool:
            for (path, job, sig), validation in pool.map(classify, pending):
                path_key = _path_key(path)
                named = "orderverificationreportviewer" in path.name.casefold()
                # Current PDF contents outrank stale quarantine history. CBC
                # reuses the same destination name when a true Sales Order
                # replaces a report, so path history alone must never
                # quarantine that repair. A path with report history (known/
                # named) stays OUT of the scan cache, so it is re-checked
                # every pass rather than trusted from one good read.
                if validation.document_kind == DOCUMENT_KIND_SALES_ORDER:
                    if sig is not None and path_key not in known and not named:
                        survivors[path_key] = sig
                    continue
                if (
                    validation.document_kind == DOCUMENT_KIND_ORDER_VERIFICATION
                    or path_key in known
                    or named
                ):
                    if validation.document_kind != DOCUMENT_KIND_ORDER_VERIFICATION:
                        validation = replace(
                            validation,
                            document_kind=DOCUMENT_KIND_ORDER_VERIFICATION,
                        )
                    reports.append((path, job, validation))
                elif sig is not None:
                    survivors[path_key] = sig   # validated clean this pass

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
    # Survivors only: quarantined/vanished files fall out, a failed quarantine
    # stays uncached so the next pass retries it. Best-effort — a cache that
    # can't be written just means the next pass validates more.
    try:
        _save_json_unlocked(SCAN_CACHE_PATH, survivors)
    except OSError as exc:
        log.debug("Order Verification scan cache not saved (%s)", exc)
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


def run(max_workers: int = 8, lock_timeout: float = 900.0) -> Dict[str, int]:
    """Enforce the no-verification-report invariant across files and stores.

    One process at a time (the cleanup lock). `lock_timeout` is how long to
    wait for another process's run to finish; pass 0 to raise TimeoutError
    immediately instead — callers that merely piggyback the cleanup onto their
    startup (the watcher) skip it when someone else is already sweeping."""
    with data_file_lock(CLEANUP_AUDIT_PATH, label="Order Verification cleanup",
                        timeout=lock_timeout):
        source_paths = _source_json_paths()
        report_paths: set[str] = set()
        report_jobs: set[str] = set()
        trusted_sales_order_paths: set[str] = set()
        for path in source_paths:
            _collect_metadata_paths(
                _read_json(path),
                report_paths,
                report_jobs,
                trusted_sales_order_paths,
            )

        quarantined_paths, quarantined_jobs = _quarantined_report_evidence(
            Path(SALES_ORDER_DIR)
        )
        report_paths.update(quarantined_paths)
        report_jobs.update(quarantined_jobs)

        moved = quarantine_active_reports(
            max_workers=max_workers,
            known_report_paths=report_paths,
        )
        # A successful later fetch supersedes stale path-only evidence from an
        # old report receipt. Actual reports found in the bank are added last,
        # so current PDF contents always win in either direction.
        report_paths.difference_update(trusted_sales_order_paths)
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
            # Work was found — the repair isn't provably done; the next
            # startup sweeps again until a pass comes back clean.
            try:
                CLEAN_MARKER_PATH.unlink(missing_ok=True)
            except OSError:
                pass
        else:
            _save_json_unlocked(CLEAN_MARKER_PATH, {"clean_at": when})
        return counts


def startup_check(max_workers: int = 8, lock_timeout: float = 0):
    """The interactive-tool gate (the watcher calls this, not run()): sweep
    only while there may be repair work left — i.e. no pass has come back
    clean yet on this machine, or the last pass found something. Once a pass
    ends clean, startups skip entirely: a NEW document doesn't need a database
    sweep to be trusted (every fetch validates its own PDF and quarantines a
    verification report before anything parses it); the overnight batch tools
    keep running the full run() as the periodic safety net. Returns run()'s
    counts, or None when skipped."""
    if isinstance(_read_json(CLEAN_MARKER_PATH), dict):
        log.debug("Order Verification cleanup already completed a clean pass; "
                  "skipping the startup sweep (fetch-time validation guards "
                  "new documents).")
        return None
    return run(max_workers=max_workers, lock_timeout=lock_timeout)


def changed(counts: "Dict[str, int] | None") -> bool:
    return any(int(value or 0) for value in (counts or {}).values())
