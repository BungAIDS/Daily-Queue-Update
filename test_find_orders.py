"""Tests for the pure search/similarity helpers in find_orders.py — the
rarity-weighted --like ranking and the custom-DWG join. Store dicts are built
inline; nothing touches disk.

    python test_find_orders.py
"""
from __future__ import annotations

import sys

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


def test_attach_dwg_annotates_hits():
    hits = [{"job": "421000"}, {"job": "421001"}, {"job": "421002"}]
    dwg = {"421000": {"extras": {"07": "DWG"}, "folder": "Z:\\JOBS\\421000"},
           "421001": {"extras": {}, "folder": "Z:\\JOBS\\421001"}}
    fo.attach_dwg(hits, dwg)
    assert hits[0]["dwg_extras"] == {"07": "DWG"} and hits[0]["dwg_scanned"]
    assert hits[1]["dwg_extras"] == {} and hits[1]["dwg_scanned"]   # scanned, all standard
    assert hits[2]["dwg_extras"] == {} and not hits[2]["dwg_scanned"]  # never scanned


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
