"""Tests for the pure search/similarity helpers in find_orders.py — the
rarity-weighted --like ranking and the custom-DWG join. Store dicts are built
inline; nothing touches disk.

    python test_find_orders.py
"""
from __future__ import annotations

import sys
from types import SimpleNamespace
from unittest import mock

import find_orders as fo


def _store():
    """4 jobs: 421000/421001 share a rare identical line + tag, everyone but
    421003 has the ubiquitous MOTOR tag, 421003 shares nothing with 421000."""
    return {"jobs": {
        "421000": {"customer": "ACME", "items": [
            {"raw": "TEFLON SHAFT SEAL", "norm": "TEFLON SHAFT SEAL", "tags": ["SHAFT SEAL"]},
            {"raw": "MOTOR 10HP TEFC", "norm": "MOTOR", "tags": ["MOTOR"]},
        ]},
        "421001": {"customer": "BRAVO", "co_number": 1, "items": [
            {"raw": "TEFLON SHAFT SEAL", "norm": "TEFLON SHAFT SEAL", "tags": ["SHAFT SEAL"]},
            {"raw": "MOTOR 5HP", "norm": "MOTOR", "tags": ["MOTOR"]},
        ]},
        "421002": {"customer": "CHARLIE", "items": [
            {"raw": "MOTOR 5HP", "norm": "MOTOR", "tags": ["MOTOR"]},
        ]},
        "421003": {"customer": "DELTA", "items": [
            {"raw": "SS SHAFT SLEEVE", "norm": "STAINLESS SHAFT SLEEVE", "tags": ["SHAFT SLEEVE"]},
        ]},
    }}


def test_similar_ranks_identical_rare_line_first():
    res = fo.similar_jobs(_store(), "421000", top=0)
    assert [r["job"] for r in res] == ["421001", "421002"]  # DELTA shares nothing
    assert res[0]["score"] > res[1]["score"]
    assert "TEFLON SHAFT SEAL" in res[0]["shared_lines"]
    assert res[0]["shared_tags"] == ["SHAFT SEAL", "MOTOR"]  # rarest tag first
    # The common MOTOR overlap alone scores 1/3-tag + 2/3-line:
    assert abs(res[1]["score"] - 1.0) < 1e-9


def test_similar_require_dwg_keeps_only_drawn_jobs():
    dwg = {"421002": {"extras": {"07": "DWG"}, "folder": "Z:\\JOBS\\421002"}}
    res = fo.similar_jobs(_store(), "421000", dwg=dwg, top=0, require_dwg=True)
    assert [r["job"] for r in res] == ["421002"]
    assert res[0]["dwg_extras"] == {"07": "DWG"}
    assert res[0]["dwg_folder"] == "Z:\\JOBS\\421002"


def test_similar_top_limits_and_unknown_job_is_none():
    assert len(fo.similar_jobs(_store(), "421000", top=1)) == 1
    assert fo.similar_jobs(_store(), "999999") is None
    assert fo.similar_jobs({"jobs": {}}, "421000") is None


def test_dwg_label():
    assert fo._dwg_label({"07": "DWG", "51": "PDF+DWG"}) == "-07 (DWG), -51 (PDF+DWG)"
    assert fo._dwg_label({}) == ""
    assert fo._dwg_label(None) == ""


def test_similar_to_items_matches_unstored_order():
    # A brand-new order (not in the store) still gets ranked by its items —
    # the watcher path: enrich parses the SO, then asks for reuse candidates.
    items = [{"raw": "TEFLON SHAFT SEAL", "norm": "TEFLON SHAFT SEAL", "tags": ["SHAFT SEAL"]}]
    idx = fo.build_index(_store())
    res = fo.similar_to_items(idx, items, exclude_job="999999", top=0)
    assert [r["job"] for r in res][:2] == ["421001", "421000"]  # both carry the line
    assert res[0]["shared_lines"] == ["TEFLON SHAFT SEAL"]


def test_reuse_suggestions_threshold_and_trim():
    dwg = {"421001": {"extras": {"07": "DWG"}, "folder": "Z:\\JOBS\\421001"},
           "421002": {"extras": {"12": "PDF"}, "folder": "Z:\\JOBS\\421002"}}
    idx = fo.build_index(_store(), dwg=dwg)
    items = _store()["jobs"]["421000"]["items"]
    sugg = fo.reuse_suggestions(idx, items, exclude_job="421000", min_score=0.5, top=3)
    # 421001 shares the rare identical line (scores well above 0.5);
    # 421002 only shares ubiquitous MOTOR (1.0 with 3 carriers -> 1/3 + 2/3 = 1.0)...
    jobs = [r["job"] for r in sugg]
    assert jobs[0] == "421001"
    assert sugg[0]["suffixes"] == ["07"] and sugg[0]["folder"] == "Z:\\JOBS\\421001"
    assert sugg[0]["lines"] and sugg[0]["score"] >= 0.5
    # ...and a high threshold silences everything.
    assert fo.reuse_suggestions(idx, items, exclude_job="421000", min_score=99) == []


def test_reuse_label_and_note():
    sugg = [{"job": "421001", "customer": "BRAVO", "score": 1.5, "suffixes": ["07", "51"],
             "dwg": "-07 (DWG), -51 (PDF)", "folder": "Z:\\JOBS\\421001",
             "lines": ["TEFLON SHAFT SEAL"], "tags": ["SHAFT SEAL"]},
            {"job": "421002", "customer": "", "score": 0.6, "suffixes": ["12"],
             "dwg": "-12 (PDF)", "folder": "", "lines": [], "tags": ["MOTOR"]}]
    assert fo.reuse_label(sugg) == "421001 (-07,-51) +1"
    assert fo.reuse_label(sugg[:1]) == "421001 (-07,-51)"
    assert fo.reuse_label([]) == ""
    note = fo.reuse_note(sugg)
    assert "421001  BRAVO — score 1.50" in note
    assert "= TEFLON SHAFT SEAL" in note and "Z:\\JOBS\\421001" in note
    assert fo.reuse_note([]) == ""


def test_attach_dwg_annotates_hits():
    hits = [{"job": "421000"}, {"job": "421001"}, {"job": "421002"}]
    dwg = {"421000": {"extras": {"07": "DWG"}, "folder": "Z:\\JOBS\\421000"},
           "421001": {"extras": {}, "folder": "Z:\\JOBS\\421001"}}
    fo.attach_dwg(hits, dwg)
    assert hits[0]["dwg_extras"] == {"07": "DWG"} and hits[0]["dwg_scanned"]
    assert hits[1]["dwg_extras"] == {} and hits[1]["dwg_scanned"]   # scanned, all standard
    assert hits[2]["dwg_extras"] == {} and not hits[2]["dwg_scanned"]  # never scanned


def test_watch_similarity_cache_tracks_backfill_overlay_updates():
    import autocad_scan
    import watch

    class FakePath:
        def __init__(self, mtime: int):
            self.mtime = mtime

        def stat(self):
            return SimpleNamespace(st_mtime_ns=self.mtime)

    base = FakePath(10)
    overlay = FakePath(20)
    dwg = FakePath(30)
    jobs = [{"job": "421000", "line_items": []}]
    old_cache = dict(watch._SIM_CACHE)
    try:
        watch._SIM_CACHE.update(key=None, rows=[])
        with (
            mock.patch.object(watch.line_items, "store_path", return_value=base),
            mock.patch.object(watch.line_items, "backfill_store_path", return_value=overlay),
            mock.patch.object(watch.line_items, "load_store", return_value={"jobs": {}}) as load,
            mock.patch.object(autocad_scan, "PROGRESS_PATH", dwg),
            mock.patch.object(autocad_scan, "load_progress", return_value={}),
            mock.patch.object(fo, "build_index", return_value={}),
            mock.patch.object(fo, "similar_to_items", return_value=[]),
        ):
            watch._similar_orders_rows(jobs)
            watch._similar_orders_rows(jobs)
            assert load.call_count == 1

            overlay.mtime += 1
            watch._similar_orders_rows(jobs)
            assert load.call_count == 2
    finally:
        watch._SIM_CACHE.clear()
        watch._SIM_CACHE.update(old_cache)


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
