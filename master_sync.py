"""Fold every helper's per-order findings into the one master store.

`live_master.json` is the single source of truth for everything we know about an
order. The live daily/watcher enrichment already writes into it; this module
brings in the *offline* helpers' stores so their data lands there too:

    autocad_scan.py       custom AutoCAD DWGs (+ job type/folder)
    quote_run_scan.py     parsed quote/construction runs
    line_items_scan.py    Sales-Order line items + feature tags
    backfill_orders.py    full SO spec + drive run for old orders

Each helper calls `master_sync.run("<source>")` at the end of its run, so any
time a helper executes, what it gathered is merged into the master. You can also
consolidate everything on demand:

    python master_sync.py            # merge every helper store into the master
    python master_sync.py autocad    # just one source

Reads each store straight off disk (by path) so there's no heavy import coupling.
Concurrency note: this does a load-merge-save of live_master.json, so prefer
running the big backlog syncs when the watcher isn't also writing the master.
"""
from __future__ import annotations

import json
import logging
import sys
from pathlib import Path
from typing import Any, Callable, Dict, List

import live_master
from config import BACKLOG_DIR, DATA_PUSH_BRANCH, DATA_PUSH_ON_CHANGE

log = logging.getLogger("master-sync")


def _publish_data_if_enabled() -> None:
    """Republish the order data after a sync changed the master, when auto-publish
    is enabled. Best-effort: a failed/absent push must never sink a scan."""
    if not (DATA_PUSH_ON_CHANGE and DATA_PUSH_BRANCH):
        return
    try:
        import data_push
        data_push.push_data()
    except Exception as e:  # noqa: BLE001 - publishing must never break a sync
        log.debug("auto data-push after sync skipped (%s)", e)

_AUTOCAD = BACKLOG_DIR / "autocad_scan_progress.json"
_QUOTE_RUNS = BACKLOG_DIR / "quote_run_scan_progress.json"
_BACKFILL = BACKLOG_DIR / "backfill_progress.json"


def _read_json(path: Path) -> Any:
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as e:
        log.warning("Could not read %s (%s)", path, e)
        return None


# --- per-source mergers (each returns how many orders it changed) ----------- #
def merge_autocad(master: Dict[str, Any]) -> int:
    recs = _read_json(_AUTOCAD) or {}
    n = 0
    for job, rec in recs.items():
        if live_master.merge_order(master, job, {
            "dwg_extras": rec.get("extras") or {},
            "dwg_missing_std": rec.get("missing_std", False),
            "job_type": rec.get("type", ""),
            "job_folder": rec.get("folder", ""),
        }):
            n += 1
    return n


def merge_quote_runs(master: Dict[str, Any]) -> int:
    recs = _read_json(_QUOTE_RUNS) or {}
    n = 0
    for job, rec in recs.items():
        runs = rec.get("runs") or []
        if not runs:
            continue
        r0 = runs[0]
        if live_master.merge_order(master, job, {
            "has_drive_run": True,
            "drive_run": r0.get("fields") or {},
            "drive_run_summary": r0.get("summary", ""),
            "drive_run_template": r0.get("template", ""),
            "drive_run_pdf": r0.get("path", ""),
            "drive_run_count": len(runs),
            "job_type": rec.get("type", ""),
            "job_folder": rec.get("folder", ""),
        }):
            n += 1
    return n


def merge_line_items(master: Dict[str, Any]) -> int:
    import line_items
    store = line_items.load_store()
    n = 0
    for job, rec in (store.get("jobs") or {}).items():
        items = rec.get("items") or []
        if live_master.merge_order(master, job, {
            "line_items": items,
            "line_item_tags": line_items.tags_label(items),
            "customer": rec.get("customer", ""),
            "co_number": rec.get("co_number"),
            "so_pdf": rec.get("so_pdf", ""),
        }):
            n += 1
    return n


def merge_backfill(master: Dict[str, Any]) -> int:
    recs = _read_json(_BACKFILL) or {}
    skip = {"job", "status", "scanned_at", "line_item_count"}
    n = 0
    for job, rec in recs.items():
        fields = {k: v for k, v in rec.items() if k not in skip}
        if fields and live_master.merge_order(master, job, fields):
            n += 1
    return n


_MERGERS: Dict[str, Callable[[Dict[str, Any]], int]] = {
    "autocad": merge_autocad,
    "quote_runs": merge_quote_runs,
    "line_items": merge_line_items,
    "backfill": merge_backfill,
}
ALL = list(_MERGERS)


def run(*sources: str) -> Dict[str, int]:
    """Load the master, merge the given sources (all by default), save it back.
    Best-effort and self-contained so a helper can call it in a try/except."""
    chosen = [s for s in (sources or ALL) if s in _MERGERS]
    master = live_master.load_master()
    counts: Dict[str, int] = {}
    for s in chosen:
        try:
            counts[s] = _MERGERS[s](master)
        except Exception as e:  # noqa: BLE001 - one bad source shouldn't sink the rest
            log.warning("master sync from %s failed (%s)", s, e)
            counts[s] = 0
    if any(counts.values()):
        live_master.save_master(master)
        _publish_data_if_enabled()   # keep the remote snapshot in step with changes
    return counts


def main(argv: List[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
    args = [a for a in (argv if argv is not None else sys.argv[1:]) if a in _MERGERS] or ALL
    counts = run(*args)
    total_orders = len(live_master.load_master().get("orders", {}))
    print("Merged into the master log:")
    for s in args:
        print(f"  {s:12} {counts.get(s, 0):6d} order(s) updated")
    print(f"\nMaster now holds {total_orders} order(s) -> {live_master.MASTER_PATH}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
