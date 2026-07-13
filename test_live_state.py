"""Tests for the live watcher's intraday state (live_state.py).

No pytest needed — run it directly:

    python test_live_state.py

Covers the pure dict logic the watcher relies on: detecting new orders, stamping
a stable first-seen time, refreshing volatile board fields without clobbering
enrichment, baseline vs announced polls, removed/returning transitions, seeding
from the morning snapshot, and the newest-first present-orders ordering.
"""
from __future__ import annotations

import json
import sys
import tempfile
from datetime import date, datetime
from pathlib import Path
from unittest.mock import patch

import live_state


def _job(num, **kw):
    j = {"job": num, "item": "47-0-0000", "customer": "ACME", "end_date": "06/20/2026",
         "total_price": "$1,000.00"}
    j.update(kw)
    return j


T0 = datetime(2026, 6, 16, 9, 0, 0)
T1 = datetime(2026, 6, 16, 9, 2, 0)
T2 = datetime(2026, 6, 16, 9, 4, 0)


def test_new_job_numbers():
    state = {"100": {"job": {"job": "100"}}}
    board = [_job("100"), _job("200"), _job("300"), _job("200")]
    assert live_state.new_job_numbers(state, board) == ["200", "300"]


def test_baseline_poll_marks_carried_over_not_announced():
    state = {}
    deltas = live_state.record_poll(state, [_job("100"), _job("200")], T0, baseline=True)
    # Returned as "new" so the caller enriches them, but flagged carried_over so
    # they aren't announced and have no precise add time.
    assert set(deltas["new"]) == {"100", "200"}
    assert state["100"]["carried_over"] is True
    assert state["100"]["enriched"] is False
    assert state["100"]["present"] is True


def test_new_order_on_later_poll_is_announced():
    state = {}
    live_state.record_poll(state, [_job("100")], T0, baseline=True)
    deltas = live_state.record_poll(state, [_job("100"), _job("200")], T1)
    assert deltas["new"] == ["200"]
    assert state["200"]["carried_over"] is False
    assert state["200"]["first_seen"] == T1.isoformat(timespec="seconds")


def test_first_seen_is_stable_and_board_fields_refresh():
    state = {}
    live_state.record_poll(state, [_job("200", end_date="06/20/2026")], T0, baseline=True)
    state["200"]["enriched"] = True
    state["200"]["job"]["so_design_desc"] = "Vaneaxial"   # enrichment field
    first_seen = state["200"]["first_seen"]
    # Next poll: the End Date changed on the board.
    live_state.record_poll(state, [_job("200", end_date="06/25/2026")], T1)
    assert state["200"]["first_seen"] == first_seen            # never moves
    assert state["200"]["job"]["end_date"] == "06/25/2026"     # volatile field refreshed
    assert state["200"]["job"]["so_design_desc"] == "Vaneaxial"  # enrichment preserved


def test_removed_then_returning():
    state = {}
    live_state.record_poll(state, [_job("100"), _job("200")], T0, baseline=True)
    d1 = live_state.record_poll(state, [_job("100")], T1)       # 200 drops off
    assert d1["removed"] == ["200"]
    assert state["200"]["present"] is False
    d2 = live_state.record_poll(state, [_job("100"), _job("200")], T2)  # 200 is back
    assert d2["returning"] == ["200"]
    assert d2["new"] == []                                      # not a brand-new order
    assert state["200"]["present"] is True


def test_seed_from_snapshot():
    state = {}
    n = live_state.seed_from_snapshot(state, [_job("100"), _job("200")], T0.isoformat())
    assert n == 2
    assert state["100"]["enriched"] is True
    assert state["100"]["carried_over"] is True
    # Already-present entries are never overwritten by a re-seed (restart-safe).
    state["100"]["enriched"] = False
    assert live_state.seed_from_snapshot(state, [_job("100")], T1.isoformat()) == 0
    assert state["100"]["enriched"] is False


def test_present_jobs_orders_newest_first_carried_last():
    state = {}
    live_state.record_poll(state, [_job("100"), _job("200")], T0, baseline=True)  # carried
    live_state.record_poll(state, [_job("100"), _job("200"), _job("300")], T1)    # 300 new
    live_state.record_poll(state, [_job("100"), _job("200"), _job("300"), _job("400")], T2)  # 400 new
    state["300"]["present"] = state["400"]["present"] = True
    rows = live_state.present_jobs(state)
    order = [r["job"] for r in rows]
    # Newest real arrival first (400, then 300), carried-over baseline last (100, 200).
    assert order[0] == "400"
    assert order[1] == "300"
    assert set(order[2:]) == {"100", "200"}
    assert rows[0]["_carried_over"] is False
    assert rows[-1]["_carried_over"] is True


def test_mark_enriched():
    state = {}
    live_state.record_poll(state, [_job("200")], T1)
    enriched = [_job("200", so_design_desc="Plug Fan")]
    live_state.mark_enriched(state, enriched)
    assert state["200"]["enriched"] is True
    assert state["200"]["job"]["so_design_desc"] == "Plug Fan"


def test_mark_enriched_preserves_parser_gaps_from_newer_change_order():
    state = {}
    live_state.record_poll(state, [_job("200")], T1)
    state["200"]["job"].update({
        "co_number": 1,
        "so_pdf": "Z:/SO/200/200 - Sales Order CO1.pdf",
        "so_design_desc": "SQAD Dual Direct Drive",
        "so_size": "33",
        "so_special_temp": "300",
        "line_items": [{"tags": ["UNITARY BASE"]}],
    })
    enriched = [_job(
        "200", co_number=2, so_pdf="Z:/SO/200/200 - Sales Order CO2.pdf",
        so_design_desc="", so_size="", so_special_temp="0", line_items=[],
    )]
    live_state.mark_enriched(state, enriched)
    job = state["200"]["job"]
    assert job["co_number"] == 2 and job["so_pdf"].endswith("CO2.pdf")
    assert job["so_design_desc"] == "SQAD Dual Direct Drive" and job["so_size"] == "33"
    assert job["so_special_temp"] == "300"
    assert job["line_items"] == [{"tags": ["UNITARY BASE"]}]


def test_mark_enriched_rejects_order_verification_report_data():
    state = {
        "200": {
            "present": True,
            "enriched": False,
            "job": {
                "job": "200",
                "so_invalidated_at": "2026-07-13T10:00:00",
                "drive_run_pdf": "keep-run.pdf",
            },
        },
    }
    report = [{
        "job": "200",
        "co_number": 1,
        "so_pdf": "report.pdf",
        "so_document_kind": "ORDER_VERIFICATION",
        "so_source_type": "CS_SalesOrder",
        "line_items": [{"raw": "wrong"}],
        "drive_run_pdf": "keep-run.pdf",
    }]

    live_state.mark_enriched(state, report)

    job = state["200"]["job"]
    assert "so_pdf" not in job and "co_number" not in job
    assert "line_items" not in job
    assert job["drive_run_pdf"] == "keep-run.pdf"
    assert state["200"]["enriched"] is False


def test_mark_enriched_allows_true_so_to_replace_report_without_reusing_its_items():
    state = {
        "200": {
            "present": True,
            "enriched": False,
            "job": {
                "job": "200",
                "so_pdf": "report.pdf",
                "so_document_kind": "ORDER_VERIFICATION",
                "so_source_type": "CS_SalesOrder",
                "line_items": [{"raw": "wrong"}],
            },
        },
    }
    genuine = [{
        "job": "200",
        "so_pdf": "200 - Sales Order CO1.pdf",
        "so_document_kind": "SALES_ORDER",
        "so_source_type": "CBC_SalesOrder",
        "so_verified_at": "9999-01-01T00:00:00",
        "line_items": [],
    }]

    live_state.mark_enriched(state, genuine)

    job = state["200"]["job"]
    assert job["so_document_kind"] == "SALES_ORDER"
    assert job["so_pdf"].endswith("CO1.pdf")
    assert job["line_items"] == []
    assert state["200"]["enriched"] is True


def test_save_state_cannot_overwrite_external_report_cleanup(tmp: Path):
    path = tmp / "live_state.json"
    external = {
        "200": {
            "present": True,
            "enriched": False,
            "job": {
                "job": "200",
                "so_invalidated_at": "2026-07-13T10:00:00",
            },
        },
    }
    path.write_text(json.dumps(external), encoding="utf-8")
    stale_watcher = {
        "200": {
            "present": True,
            "enriched": True,
            "job": {
                "job": "200",
                "so_pdf": "report.pdf",
                "so_document_kind": "ORDER_VERIFICATION",
                "so_source_type": "CS_SalesOrder",
                "co_number": 1,
                "line_items": [{"raw": "wrong"}],
            },
        },
    }

    with patch("live_state.state_path", return_value=path):
        live_state.save_state(stale_watcher, date(2026, 7, 13))

        saved = json.loads(path.read_text(encoding="utf-8"))
        job = saved["200"]["job"]
        assert "so_pdf" not in job and "line_items" not in job
        assert saved["200"]["enriched"] is False

        # A later strict no-SO check is allowed to finish; the invalidation
        # marker must not force this job through enrichment every poll forever.
        saved["200"]["enriched"] = True
        live_state.save_state(saved, date(2026, 7, 13))
        assert json.loads(path.read_text(encoding="utf-8"))["200"]["enriched"] is True


def main() -> int:
    passed = 0
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as directory:
        tmp = Path(directory)
        for name, fn in sorted(globals().items()):
            if name.startswith("test_") and callable(fn):
                fn(tmp) if "tmp" in fn.__code__.co_varnames else fn()
                print(f"  ok  {name}")
                passed += 1
    print(f"\n{passed} tests passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
