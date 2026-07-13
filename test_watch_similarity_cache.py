"""Regression tests for throttling heavy Similar Data repaints.

    python test_watch_similarity_cache.py
"""
from __future__ import annotations

import sys
from types import SimpleNamespace
from unittest import mock

import autocad_scan
import find_orders
import watch


class _ChangingPath:
    def __init__(self, mtime: int):
        self.mtime = mtime

    def stat(self):
        return SimpleNamespace(st_mtime_ns=self.mtime)


def test_backfill_store_churn_is_throttled_but_queue_changes_are_immediate():
    main_store = _ChangingPath(1)
    backfill_store = _ChangingPath(1)
    dwg_store = _ChangingPath(1)
    jobs = [{"job": "421000", "line_items": [{"raw": "wheel"}]}]
    match = {
        "job": "400001", "customer": "TEST", "score": 2.345,
        "dwg_extras": [], "shared_lines": ["wheel"], "shared_tags": [],
        "dwg_folder": "",
    }
    original_cache = dict(watch._SIM_CACHE)
    watch._SIM_CACHE.clear()
    watch._SIM_CACHE.update(key=None, rows=[], queue_ids=(), refreshed_at=0.0)

    try:
        with mock.patch.object(watch, "SIMILAR_REFRESH_INTERVAL_SECONDS", 900), \
                mock.patch.object(watch.line_items, "store_path", return_value=main_store), \
                mock.patch.object(watch.line_items, "backfill_store_path", return_value=backfill_store), \
                mock.patch.object(watch.line_items, "load_store", return_value={"jobs": {}}), \
                mock.patch.object(autocad_scan, "PROGRESS_PATH", dwg_store), \
                mock.patch.object(autocad_scan, "load_progress", return_value={}), \
                mock.patch.object(find_orders, "build_index", return_value={}) as build, \
                mock.patch.object(find_orders, "similar_to_items", return_value=[match]), \
                mock.patch.object(find_orders, "_dwg_label", return_value=""), \
                mock.patch.object(watch.time, "monotonic",
                                  side_effect=[100.0, 101.0, 102.0, 103.0, 1004.0]):
            first = watch._similar_orders_rows(jobs)
            assert first and build.call_count == 1

            backfill_store.mtime += 1
            assert watch._similar_orders_rows(jobs) == first
            assert build.call_count == 1, "store churn inside the window must be deferred"

            jobs_with_new_order = jobs + [
                {"job": "421001", "line_items": [{"raw": "shaft"}]}
            ]
            watch._similar_orders_rows(jobs_with_new_order)
            assert build.call_count == 2, "queue membership changes must refresh immediately"

            backfill_store.mtime += 1
            watch._similar_orders_rows(jobs_with_new_order)
            assert build.call_count == 2

            watch._similar_orders_rows(jobs_with_new_order)
            assert build.call_count == 3, "store changes refresh after the interval expires"
    finally:
        watch._SIM_CACHE.clear()
        watch._SIM_CACHE.update(original_cache)


def main() -> int:
    passed = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
            print(f"  ok  {name}")
            passed += 1
    print(f"\n{passed} tests passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
