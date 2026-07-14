"""Regression tests for keeping the live Excel workbook responsive.

    python test_live_excel_performance.py

These tests exercise the pure scheduling/guard pieces. They deliberately avoid
starting Excel so they can run in CI and while the real workbook is open.
"""
from __future__ import annotations

import sys
import threading
import time
from unittest import mock

import live_excel
from live_sheets import Cell, Sheet


def _clear_render_state() -> None:
    live_excel._HEADER_DONE.clear()
    live_excel._ORDER_HISTORY_READY.clear()
    live_excel._SEARCH_CF_DONE.clear()
    live_excel._POS_LAST.clear()
    live_excel._BELOW_LAST.clear()
    live_excel._RENDER_CACHE.clear()


def test_noop_poll_never_opens_excel():
    _clear_render_state()
    below = {"title": "Removed", "rows": []}
    positions = {"421000": 1}
    lq = {
        "name": "Live Queue", "headers": ["Job #"], "ops": [],
        "key_col": 1, "allow_delete": True, "positions": positions,
        "below": below, "search": True,
    }
    oh = {
        "name": "Order History", "spec": {}, "ops": [],
        "key_col": 1, "rebuild": False,
    }
    changes = Sheet("Changes")
    changes.row([Cell("unchanged")])

    live_excel._HEADER_DONE.add("Live Queue")
    live_excel._ORDER_HISTORY_READY.add("Order History")
    live_excel._SEARCH_CF_DONE.add("Live Queue")
    live_excel._POS_LAST["Live Queue"] = {"421000": 1}
    live_excel._BELOW_LAST["Live Queue"] = below
    live_excel._RENDER_CACHE["Changes"] = live_excel._fingerprint(changes)

    with mock.patch.object(
        live_excel, "_get_excel", side_effect=AssertionError("Excel must not open")
    ):
        rendered = live_excel._update_master_workbook_impl(
            "unused.xlsx", lq, oh, changes_sheet=changes
        )
    assert rendered == {"Live Queue", "Order History"}


def test_gated_order_history_is_none_and_never_opens_excel():
    # The watcher passes oh_payload=None when the Order History inputs were
    # unchanged (it skipped the ~9s rebuild). The Excel layer must treat that as
    # "nothing to render" — not crash on oh_payload["name"] — and, with Live
    # Queue also idle, never open Excel.
    _clear_render_state()
    assert not live_excel._history_needs_excel(None)
    lq = {
        "name": "Live Queue", "headers": ["Job #"], "ops": [],
        "key_col": 1, "allow_delete": True, "positions": {}, "search": False,
    }
    live_excel._HEADER_DONE.add("Live Queue")
    live_excel._POS_LAST["Live Queue"] = {}     # positions unchanged -> Live Queue idle
    with mock.patch.object(
        live_excel, "_get_excel", side_effect=AssertionError("Excel must not open")
    ):
        rendered = live_excel._update_master_workbook_impl("unused.xlsx", lq, None)
    assert rendered == {"Live Queue"}     # OH absent, no crash on None


def test_upsert_preflight_detects_real_work_only():
    _clear_render_state()
    payload = {
        "name": "Live Queue", "ops": [], "positions": {"421000": 1},
        "below": {"rows": []}, "search": True,
    }
    assert live_excel._upsert_needs_excel(payload)

    live_excel._HEADER_DONE.add("Live Queue")
    live_excel._SEARCH_CF_DONE.add("Live Queue")
    live_excel._POS_LAST["Live Queue"] = {"421000": 1}
    live_excel._BELOW_LAST["Live Queue"] = {"rows": []}
    assert not live_excel._upsert_needs_excel(payload)

    payload["ops"] = [("update", "421000", [])]
    assert live_excel._upsert_needs_excel(payload)


def test_tuning_does_not_force_application_wide_calculation():
    class FakeApp:
        def __init__(self):
            object.__setattr__(self, "changes", [])
            object.__setattr__(self, "calculate_calls", 0)
            object.__setattr__(self, "ScreenUpdating", True)
            object.__setattr__(self, "EnableEvents", True)
            object.__setattr__(self, "DisplayAlerts", True)
            object.__setattr__(self, "Calculation", live_excel._XL_CALC_AUTOMATIC)

        def __setattr__(self, name, value):
            if name in {"ScreenUpdating", "EnableEvents", "DisplayAlerts", "Calculation"}:
                self.changes.append((name, value))
            object.__setattr__(self, name, value)

        def Calculate(self):
            self.calculate_calls += 1

    app = FakeApp()
    with live_excel._tuned(app):
        assert app.Calculation == live_excel._XL_CALC_MANUAL
    assert app.calculate_calls == 0
    assert app.Calculation == live_excel._XL_CALC_AUTOMATIC
    assert app.ScreenUpdating is True


def test_history_rows_are_batched_by_contiguous_range():
    class FakeRange:
        def __init__(self, owner, start, end):
            self.owner = owner
            self.start = start
            self.end = end

        @property
        def Value(self):
            return None

        @Value.setter
        def Value(self, value):
            self.owner.writes.append((self.start, self.end, value))

    class FakeSheet:
        def __init__(self):
            self.writes = []

        def Cells(self, row, col):
            return row, col

        def Range(self, start, end):
            return FakeRange(self, start, end)

    ws = FakeSheet()
    rows = [(2, [Cell("a")]), (3, [Cell("b")]), (4, [Cell("c")]),
            (10, [Cell("d")]), (11, [Cell("e")])]
    writes = live_excel._write_oh_rows(ws, rows, ncols=2, chunk_size=2)

    assert writes == 3
    assert [(a[0], b[0]) for a, b, _grid in ws.writes] == [(2, 3), (4, 4), (10, 11)]
    assert ws.writes[0][2] == [["a", ""], ["b", ""]]


def test_timed_out_writer_blocks_overlap_until_it_finishes():
    live_excel._excel_write_active[0] = None
    started = threading.Event()
    release = threading.Event()
    calls = []

    def slow_write():
        calls.append("slow")
        started.set()
        release.wait(2)
        return "slow-done"

    with mock.patch.object(live_excel, "_excel_write_timeout", return_value=0.01):
        first = live_excel._run_excel_guarded("test", slow_write, "timeout")
        assert first == "timeout"
        assert started.is_set()

        second = live_excel._run_excel_guarded(
            "test", lambda: calls.append("overlap") or "bad", "skipped"
        )
        assert second == "skipped"
        assert calls == ["slow"]

        release.set()
        deadline = time.monotonic() + 2
        while live_excel._excel_write_active[0] is not None and time.monotonic() < deadline:
            time.sleep(0.01)
        assert live_excel._excel_write_active[0] is None
        assert live_excel._run_excel_guarded("test", lambda: "third", "bad") == "third"


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
