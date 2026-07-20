"""Audit and repair the contract between verified SO progress and item stores.

This module never searches CBC Insider. Missing parsed records are rebuilt only
from already-archived PDFs that still validate as the requested Sales Order.
"""
from __future__ import annotations

import argparse
import json
import logging
import re
import shutil
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, Optional

from config import BACKLOG_DIR, SALES_ORDER_DIR
import line_items
from process_lock import data_file_lock
from sales_order_validation import (
    DOCUMENT_KIND_ORDER_VERIFICATION,
    DOCUMENT_KIND_SALES_ORDER,
    sales_order_sha256,
    validate_sales_order_pdf,
)
from sales_orders import parse_sales_order_pdf


log = logging.getLogger("so-corpus-health")
DEFAULT_PROGRESS = BACKLOG_DIR / "backfill_progress.json"
DEFAULT_REPORT = BACKLOG_DIR / "so_corpus_health.json"


def load_progress(path: Path = DEFAULT_PROGRESS) -> Dict[str, Dict[str, Any]]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return value if isinstance(value, dict) else {}


def save_progress(records: Dict[str, Dict[str, Any]], path: Path = DEFAULT_PROGRESS) -> None:
    with data_file_lock(path, label="Sales Order corpus progress update"):
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = path.with_suffix(path.suffix + ".tmp")
        temporary.write_text(json.dumps(records, indent=2), encoding="utf-8")
        temporary.replace(path)


def _trusted_ok(rec: Dict[str, Any]) -> bool:
    if rec.get("status") != "ok":
        return False
    if str(rec.get("so_document_kind") or "").upper() == DOCUMENT_KIND_ORDER_VERIFICATION:
        return False
    if str(rec.get("so_source_type") or "").casefold() == "cs_salesorder":
        return False
    return bool(str(rec.get("so_pdf") or "").strip())


def _store_for_audit(store_path: Optional[Path] = None) -> Dict[str, Any]:
    return line_items.load_store(store_path) if store_path is not None else line_items.load_store()


def _rate(value: int, total: int) -> float:
    return round(value / total, 4) if total else 0.0


def _file_access_status(path: Path) -> str:
    if not str(path):
        return "missing"
    try:
        return "present" if path.is_file() else "missing"
    except OSError:
        return "inaccessible"


def corpus_quality(store: Dict[str, Any]) -> Dict[str, Any]:
    """Summarize parser coverage and the remaining human-review workload."""
    jobs = store.get("jobs") or {}
    total_items = tagged = structured = components = flagged = sourced = nan_items = 0
    jobs_with_review = jobs_with_source = jobs_without_items = 0
    review_flags: Counter[str] = Counter()
    unclassified: Counter[str] = Counter()
    component_counts: Counter[str] = Counter()
    parser_versions: Counter[str] = Counter()
    nan_pattern = re.compile(r"(?i)(?<![A-Za-z0-9])nan(?![A-Za-z0-9])")

    for record in jobs.values():
        record = record or {}
        items = record.get("items") or []
        parser_versions[str(record.get("parser_version") or "legacy/unknown")] += 1
        jobs_without_items += int(not items)
        job_review = False
        job_source = False
        for item in items:
            total_items += 1
            tags = item.get("tags") or []
            attrs = item.get("attributes") if isinstance(item.get("attributes"), dict) else {}
            component = str((attrs or {}).get("component") or "").strip()
            flags = [str(flag).strip() for flag in item.get("review_flags") or []
                     if str(flag).strip()]
            tagged += int(bool(tags))
            components += int(bool(component))
            structured += int(any(key != "component" and value not in (None, "", [], {})
                                  for key, value in (attrs or {}).items()))
            flagged += int(bool(flags))
            sourced += int(bool(item.get("source")))
            job_review = job_review or bool(flags)
            job_source = job_source or bool(item.get("source"))
            review_flags.update(flags)
            if component:
                component_counts[component] += 1
            if not tags and not component:
                norm = str(item.get("norm") or item.get("raw") or "").strip()
                if norm:
                    unclassified[norm] += 1
            searchable = " ".join(
                [str(item.get("raw") or "")]
                + [str(detail) for detail in item.get("details") or []]
            )
            nan_items += int(bool(nan_pattern.search(searchable)))
        jobs_with_review += int(job_review)
        jobs_with_source += int(job_source)

    return {
        "jobs": len(jobs),
        "jobs_without_items": jobs_without_items,
        "items": total_items,
        "tagged_items": tagged,
        "tagged_rate": _rate(tagged, total_items),
        "component_items": components,
        "component_rate": _rate(components, total_items),
        "structured_attribute_items": structured,
        "structured_attribute_rate": _rate(structured, total_items),
        "review_flagged_items": flagged,
        "review_flagged_rate": _rate(flagged, total_items),
        "jobs_with_review": jobs_with_review,
        "sourced_items": sourced,
        "source_rate": _rate(sourced, total_items),
        "jobs_with_source": jobs_with_source,
        "nan_artifact_items": nan_items,
        "parser_versions": dict(parser_versions.most_common()),
        "top_components": dict(component_counts.most_common(20)),
        "top_review_flags": dict(review_flags.most_common(20)),
        "top_unclassified_wordings": dict(unclassified.most_common(20)),
    }


def recommended_actions(audit: Dict[str, Any]) -> list[str]:
    quality = audit.get("quality") or {}
    actions = []
    if audit.get("missing_store_jobs"):
        actions.append("Repair trusted Sales Orders missing from the line-item store.")
    if audit.get("inaccessible_pdf_jobs"):
        actions.append("Restore archived Sales Order share access and rerun the corpus audit.")
    if quality.get("nan_artifact_items"):
        actions.append("Incrementally reparse records containing literal NaN cell artifacts.")
    if quality.get("source_rate", 0.0) < 0.95:
        actions.append("Incrementally reparse legacy records to add PDF page/row provenance.")
    if quality.get("review_flagged_rate", 0.0) > 0.03:
        actions.append("Review the highest-frequency unresolved wording templates in batches.")
    if (audit.get("missing_metadata") or {}).get("arrangement", 0):
        actions.append("Backfill order context so arrangement participates in similar-order ranking.")
    return actions


def save_audit(audit: Dict[str, Any], path: Path = DEFAULT_REPORT) -> None:
    with data_file_lock(path, label="Sales Order corpus health report"):
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = path.with_suffix(path.suffix + ".tmp")
        temporary.write_text(json.dumps(audit, indent=2), encoding="utf-8")
        temporary.replace(path)


def _file_signature(path: Path) -> tuple[int, int] | None:
    try:
        stat = path.stat()
    except OSError:
        return None
    return stat.st_size, stat.st_mtime_ns


def _progress_value(record: Dict[str, Any], *keys: str) -> Any:
    for key in keys:
        value = record.get(key)
        if value not in (None, "", [], {}):
            return value
    return None


def apply_progress_metadata(store: Dict[str, Any],
                            records: Dict[str, Dict[str, Any]]) -> int:
    """Fill ranking/provenance context already known by the backfill ledger."""
    changed_jobs = 0
    for job, target in (store.get("jobs") or {}).items():
        source = records.get(job) or {}
        if not source:
            continue
        before = {key: target.get(key) for key in (
            "customer", "co_number", "so_pdf", "arrangement", "job_number",
            "parts_only", *line_items.ORDER_CONTEXT_FIELDS,
        )}
        values = {
            "customer": _progress_value(source, "customer"),
            "co_number": _progress_value(source, "co_number"),
            "so_pdf": _progress_value(source, "so_pdf"),
            "arrangement": _progress_value(source, "so_arrangement", "arrangement"),
            "job_number": _progress_value(source, "job_number"),
            "so_design_desc": _progress_value(source, "so_design_desc"),
            "so_size": _progress_value(source, "so_size"),
            "so_arrangement": _progress_value(source, "so_arrangement", "arrangement"),
            "so_motor_pos": _progress_value(source, "so_motor_pos"),
            "so_class": _progress_value(source, "so_class"),
            "so_rotation": _progress_value(source, "so_rotation"),
            "so_discharge": _progress_value(source, "so_discharge"),
            "so_pct_width": _progress_value(source, "so_pct_width"),
            "so_wheel_type": _progress_value(source, "so_wheel_type"),
            "so_design_temp": _progress_value(source, "so_design_temp"),
            "so_max_temp": _progress_value(source, "so_max_temp"),
            "so_special_temp": _progress_value(source, "so_special_temp"),
            "source_pdf_sha256": _progress_value(source, "so_pdf_sha256"),
        }
        for key, value in values.items():
            if value not in (None, "", [], {}):
                target[key] = value
        if "parts_only" in source:
            target["parts_only"] = bool(source.get("parts_only"))
        after = {key: target.get(key) for key in before}
        changed_jobs += int(after != before)
    return changed_jobs


def _backup_name(path: Path, stamp: str) -> Path:
    return path.with_name(f"{path.stem}.bak-{stamp}{path.suffix}")


def consolidate_stores(
    records: Dict[str, Dict[str, Any]],
    *,
    main_path: Optional[Path] = None,
    overlay_path: Optional[Path] = None,
    dry_run: bool = False,
    create_backups: bool = True,
    max_attempts: int = 3,
) -> Dict[str, Any]:
    """Normalize and consolidate main+overlay without stopping the watcher.

    Parsing happens outside file locks. Immediately before committing, both
    source signatures are checked under their normal process locks. If watch.py
    changed either file meanwhile, the operation retries from fresh snapshots.
    """
    main_path = Path(main_path or line_items.store_path())
    overlay_path = Path(overlay_path or line_items.backfill_store_path())
    if main_path == overlay_path:
        raise ValueError("Main and overlay stores must be different files")

    for attempt in range(1, max(1, max_attempts) + 1):
        before_main = _file_signature(main_path)
        before_overlay = _file_signature(overlay_path)
        main_store = line_items.load_store(main_path)
        overlay_store = line_items.load_store(overlay_path)
        if (before_main != _file_signature(main_path)
                or before_overlay != _file_signature(overlay_path)):
            continue

        main_jobs = set(main_store.get("jobs") or {})
        overlay_jobs = set(overlay_store.get("jobs") or {})
        merged = line_items.merge_stores(main_store, overlay_store)
        item_count = line_items.renormalize_store(merged)
        metadata_jobs = apply_progress_metadata(merged, records)
        result = {
            "attempt": attempt,
            "main_jobs_before": len(main_jobs),
            "overlay_jobs_before": len(overlay_jobs),
            "overlap_jobs": len(main_jobs & overlay_jobs),
            "main_only_jobs": len(main_jobs - overlay_jobs),
            "overlay_only_jobs": len(overlay_jobs - main_jobs),
            "consolidated_jobs": len(merged.get("jobs") or {}),
            "normalized_items": item_count,
            "metadata_jobs_updated": metadata_jobs,
            "dry_run": dry_run,
            "backups": [],
        }
        if dry_run:
            return result

        with data_file_lock(main_path, label="main line-items consolidation"):
            with data_file_lock(overlay_path, label="backfill overlay consolidation"):
                if (before_main != _file_signature(main_path)
                        or before_overlay != _file_signature(overlay_path)):
                    continue
                if create_backups:
                    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    for source in (main_path, overlay_path):
                        if source.is_file():
                            backup = _backup_name(source, stamp)
                            shutil.copy2(source, backup)
                            result["backups"].append(str(backup))
                merged["consolidated_at"] = datetime.now().isoformat(timespec="seconds")
                line_items.save_store(merged, main_path)
                line_items.save_store({
                    "jobs": {},
                    "ai_tags": {},
                    "consolidated_at": merged["consolidated_at"],
                    "consolidated_into": str(main_path),
                }, overlay_path)
                return result
    raise RuntimeError(
        f"Stores changed during all {max_attempts} consolidation attempt(s); retry later."
    )


def audit_corpus(
    records: Dict[str, Dict[str, Any]],
    *,
    store_path: Optional[Path] = None,
) -> Dict[str, Any]:
    store = _store_for_audit(store_path)
    jobs = store.get("jobs") or {}
    trusted = {job for job, rec in records.items() if _trusted_ok(rec)}
    missing_store = sorted(trusted - set(jobs))
    pdf_status = {
        job: _file_access_status(Path(str(records[job].get("so_pdf") or "")))
        for job in trusted
    }
    missing_pdf = sorted(job for job, status in pdf_status.items() if status == "missing")
    inaccessible_pdf = sorted(
        job for job, status in pdf_status.items() if status == "inaccessible"
    )
    count_mismatches = []
    for job in sorted(trusted & set(jobs)):
        expected = records[job].get("line_item_count")
        if expected in (None, ""):
            continue
        try:
            expected_count = int(expected)
        except (TypeError, ValueError):
            continue
        stored_count = len((jobs[job] or {}).get("items") or [])
        if stored_count != expected_count:
            count_mismatches.append({
                "job": job,
                "progress_count": expected_count,
                "stored_count": stored_count,
                "difference": stored_count - expected_count,
            })
    metadata = {
        key: sum(not str((jobs[job] or {}).get(key) or "").strip()
                 for job in trusted & set(jobs))
        for key in ("arrangement", "customer", "job_number", "so_pdf")
    }
    audit = {
        "checked_at": datetime.now().isoformat(timespec="seconds"),
        "progress_jobs": len(records),
        "trusted_ok_jobs": len(trusted),
        "stored_jobs": len(jobs),
        "trusted_with_store": len(trusted & set(jobs)),
        "missing_store_jobs": missing_store,
        "missing_pdf_jobs": missing_pdf,
        "inaccessible_pdf_jobs": inaccessible_pdf,
        "item_count_mismatches": count_mismatches,
        "missing_metadata": metadata,
        "quality": corpus_quality(store),
        "healthy": not missing_store and not missing_pdf and not inaccessible_pdf,
    }
    audit["recommended_actions"] = recommended_actions(audit)
    return audit


def _archive_candidates(job: str, rec: Dict[str, Any]) -> Iterable[Path]:
    preferred = Path(str(rec.get("so_pdf") or ""))
    seen = set()
    try:
        preferred_exists = bool(str(preferred) and preferred.is_file())
    except OSError:
        preferred_exists = False
    if preferred_exists:
        seen.add(str(preferred).casefold())
        yield preferred
    folder = SALES_ORDER_DIR / job
    try:
        if not folder.is_dir():
            return
        candidates = list(folder.glob("*.pdf"))
    except OSError:
        return

    def modified(path: Path) -> float:
        try:
            return path.stat().st_mtime
        except OSError:
            return 0.0

    for candidate in sorted(candidates, key=modified, reverse=True):
        key = str(candidate).casefold()
        if key not in seen:
            seen.add(key)
            yield candidate


def _validated_archive_pdf(job: str, rec: Dict[str, Any]) -> tuple[Optional[Path], str]:
    reasons = []
    for candidate in _archive_candidates(job, rec):
        validation = validate_sales_order_pdf(
            candidate,
            job,
            required_document_kind=DOCUMENT_KIND_SALES_ORDER,
        )
        if validation.matched and validation.document_kind == DOCUMENT_KIND_SALES_ORDER:
            return candidate, ""
        reasons.append(f"{candidate.name}: {validation.status}/{validation.document_kind}")
    return None, "; ".join(reasons) or "no archived PDF exists"


def repair_missing_line_items(
    records: Dict[str, Dict[str, Any]],
    *,
    destination: Optional[Path] = None,
    only_jobs: Optional[Iterable[str]] = None,
    progress_path: Optional[Path] = None,
    dry_run: bool = False,
) -> Dict[str, Any]:
    """Rebuild missing trusted jobs from their validated archived PDFs."""
    destination = destination or line_items.backfill_store_path()
    current = _store_for_audit(destination if destination != line_items.backfill_store_path() else None)
    stored_jobs = set(current.get("jobs") or {})
    wanted = {str(job) for job in only_jobs} if only_jobs is not None else None
    missing = [
        job for job, rec in sorted(records.items())
        if _trusted_ok(rec) and job not in stored_jobs and (wanted is None or job in wanted)
    ]
    updates = []
    parsed_rows = []
    failures = []
    for job in missing:
        rec = records[job]
        pdf, error = _validated_archive_pdf(job, rec)
        if pdf is None:
            failures.append({"job": job, "reason": error})
            continue
        parsed = parse_sales_order_pdf(pdf)
        items = parsed.get("line_items") or []
        try:
            expected = int(rec.get("line_item_count") or 0)
        except (TypeError, ValueError):
            expected = 0
        if expected and not items:
            failures.append({"job": job, "reason": "validated PDF reparsed to zero items"})
            continue
        update = {
            "job": job,
            "items": items,
            "customer": str(rec.get("customer") or ""),
            "co_number": rec.get("co_number"),
            "so_pdf": str(pdf),
            "arrangement": str(parsed.get("arrangement") or rec.get("so_arrangement") or ""),
            "parts_only": bool(parsed.get("parts_only", rec.get("parts_only", False))),
            "job_number": str(parsed.get("job_number") or rec.get("job_number") or ""),
            **{key: rec.get(key) for key in line_items.ORDER_CONTEXT_FIELDS},
            "source_pdf_sha256": sales_order_sha256(pdf),
        }
        updates.append(update)
        parsed_rows.append((job, rec, pdf, parsed, len(items)))

    if dry_run:
        return {
            "missing_before": missing,
            "repairable": [row[0] for row in parsed_rows],
            "repaired": [],
            "failures": failures,
            "remaining_missing": missing,
        }

    if updates:
        line_items.record_jobs_atomic(updates, destination)
        repaired_at = datetime.now().isoformat(timespec="seconds")
        for job, rec, pdf, parsed, item_count in parsed_rows:
            rec["so_pdf"] = str(pdf)
            rec["so_pdf_sha256"] = sales_order_sha256(pdf)
            rec["so_arrangement"] = str(parsed.get("arrangement") or rec.get("so_arrangement") or "")
            rec["parts_only"] = bool(parsed.get("parts_only", rec.get("parts_only", False)))
            rec["job_number"] = str(parsed.get("job_number") or rec.get("job_number") or "")
            rec["line_item_count"] = item_count
            rec["line_item_store_committed_at"] = repaired_at
            rec["line_item_store_repaired_at"] = repaired_at
        if progress_path is not None:
            save_progress(records, progress_path)

    verified = line_items.load_store(destination)
    remaining = [job for job in missing if job not in (verified.get("jobs") or {})]
    repaired = [job for job in missing if job not in remaining]
    return {
        "missing_before": missing,
        "repairable": [row[0] for row in parsed_rows],
        "repaired": repaired,
        "failures": failures,
        "remaining_missing": remaining,
    }


def _print_summary(audit: Dict[str, Any]) -> None:
    print(
        f"Trusted Sales Orders: {audit['trusted_ok_jobs']}; "
        f"with parsed store records: {audit['trusted_with_store']}."
    )
    print(
        f"Missing store: {len(audit['missing_store_jobs'])}; "
        f"missing archived PDF: {len(audit['missing_pdf_jobs'])}; "
        f"inaccessible archived PDF: {len(audit.get('inaccessible_pdf_jobs') or [])}; "
        f"item-count differences: {len(audit['item_count_mismatches'])}."
    )
    quality = audit.get("quality") or {}
    print(
        f"Items: {quality.get('items', 0)}; tagged {quality.get('tagged_rate', 0):.1%}; "
        f"structured {quality.get('structured_attribute_rate', 0):.1%}; "
        f"review {quality.get('review_flagged_rate', 0):.1%}; "
        f"source-linked {quality.get('source_rate', 0):.1%}."
    )
    if audit["missing_store_jobs"]:
        print("Missing store jobs: " + ", ".join(audit["missing_store_jobs"]))


def main(argv: Optional[list[str]] = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    parser = argparse.ArgumentParser(description="Audit or repair the Sales Order corpus stores.")
    parser.add_argument("--progress", type=Path, default=DEFAULT_PROGRESS)
    parser.add_argument("--store", type=Path, help="Explicit item store (default: merged live stores).")
    parser.add_argument("--repair-missing", action="store_true",
                        help="Reparse missing trusted jobs from validated archived PDFs.")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--job", action="append", default=[], help="Limit repair to a job; repeatable.")
    parser.add_argument("--json", action="store_true", help="Print machine-readable JSON.")
    parser.add_argument("--save-report", nargs="?", const=DEFAULT_REPORT, type=Path,
                        help="Write the audit JSON (default: backlog/so_corpus_health.json).")
    parser.add_argument("--consolidate-stores", action="store_true",
                        help="Normalize and merge the backfill overlay into the main item store.")
    args = parser.parse_args(sys.argv[1:] if argv is None else argv)

    records = load_progress(args.progress)
    if args.repair_missing:
        result = repair_missing_line_items(
            records,
            destination=args.store,
            only_jobs=args.job or None,
            progress_path=(None if args.dry_run else args.progress),
            dry_run=args.dry_run,
        )
        if args.json:
            print(json.dumps(result, indent=2))
        elif args.dry_run:
            print(
                f"Dry run: {len(result['repairable'])} job(s) validated and reparsed; "
                f"{len(result['failures'])} failed validation or parsing."
            )
        else:
            print(
                f"Repaired {len(result['repaired'])} job(s); "
                f"{len(result['remaining_missing'])} remain missing."
            )

    if args.consolidate_stores:
        if args.store:
            parser.error("--consolidate-stores cannot be combined with --store")
        result = consolidate_stores(records, dry_run=args.dry_run)
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            mode = "Would consolidate" if args.dry_run else "Consolidated"
            print(
                f"{mode} {result['consolidated_jobs']} job(s), including "
                f"{result['overlap_jobs']} duplicate main/overlay records; "
                f"normalized {result['normalized_items']} item(s) and updated "
                f"context on {result['metadata_jobs_updated']} job(s)."
            )

    audit = audit_corpus(records, store_path=args.store)
    if args.save_report:
        save_audit(audit, args.save_report)
    if args.json:
        print(json.dumps(audit, indent=2))
    else:
        _print_summary(audit)
    return 0 if audit["healthy"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
