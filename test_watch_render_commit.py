"""Regression test: a failed Excel write must not lose Live Queue / Order History rows.

    python test_watch_render_commit.py

The live tabs are drawn incrementally — each poll only writes the rows whose
content changed since the last poll, tracked by a per-row signature store
(master['lq_sigs'] / ['oh_sigs']). The bug ('all the orders vanished'): the store
was updated as soon as the ops were *planned*, before the Excel write. So when a
write failed (Excel busy / OLE error 0x800ac472), the store believed the rows
were on the sheet, the next poll planned no op for them, and they stayed missing
until a manual restart.

The fix: commit a tab's signatures only when update_master_workbook reports that
tab rendered without error. This test drives _render_master with the workbook
write mocked to fail, then to succeed, and checks the signature store.
"""
from __future__ import annotations

import sys
import tempfile
from datetime import datetime
from pathlib import Path
from unittest import mock

import config

config.SNAPSHOT_DIR = Path(tempfile.mkdtemp())

import watch  # noqa: E402

NOW = datetime(2026, 6, 25, 9, 0, 0)


def _master_one_order() -> dict:
    return {"orders": {"420734": {
        "added": "2026-06-24T05:00:00", "last_in": "2026-06-24T05:00:00",
        "last_out": None, "left": None, "on_queue": True, "seen_on_queue": True,
        "added_known": True,
        "job": {"job": "420734", "customer": "ABTEC FILTERS", "oper": "53",
                "end_date": "07/01/2026", "co_number": 0},
    }}}


def _render(master: dict, rendered: set) -> None:
    """Run _render_master with the Excel write mocked to report `rendered` as the
    tabs that succeeded, and disk/IO side effects stubbed."""
    with mock.patch.object(watch, "update_master_workbook", return_value=rendered), \
            mock.patch.object(watch.order_explorer, "maybe_write", return_value=None), \
            mock.patch.object(watch.line_items, "load_store", return_value={"jobs": {}}):
        watch._render_master(master, NOW, board_order=["420734"])


def test_failed_write_leaves_rows_uncommitted_then_redraws():
    master = _master_one_order()

    # Poll 1: Excel is busy -> the write fails for every tab (empty set).
    _render(master, set())
    assert master.get("lq_sigs") == {}, \
        "a failed Live Queue write must not commit row signatures"
    assert master.get("oh_sigs") == {}, \
        "a failed Order History write must not commit row signatures"

    # Poll 2: Excel recovers -> both tabs render. The same rows are re-planned
    # (because nothing was committed last time) and now get committed.
    captured = {}

    real_plan = watch._plan

    def spy_plan(records, sig_store, allow_delete):
        ops, commit = real_plan(records, sig_store, allow_delete)
        # remember the op count for whichever store this is
        captured[id(sig_store)] = len(ops)
        return ops, commit

    with mock.patch.object(watch, "_plan", side_effect=spy_plan):
        _render(master, {"Live Queue", "Order History"})

    assert "420734" in master["lq_sigs"], \
        "a successful Live Queue write must commit the row signature"
    # The order is still re-planned on poll 2 (it was never committed on poll 1),
    # proving the row would have been redrawn rather than silently dropped.
    assert any(n >= 1 for n in captured.values()), \
        "rows uncommitted after a failed write must be re-planned next poll"


def test_successful_write_commits_then_skips_unchanged():
    master = _master_one_order()
    _render(master, {"Live Queue", "Order History"})
    assert "420734" in master["lq_sigs"]

    # Nothing changed -> the next render plans no Live Queue op for the order.
    seen = {}

    real_plan = watch._plan

    def spy_plan(records, sig_store, allow_delete):
        ops, commit = real_plan(records, sig_store, allow_delete)
        if any(k == "420734" for _, k, _ in ops):
            seen["lq_op_for_order"] = True
        return ops, commit

    with mock.patch.object(watch, "_plan", side_effect=spy_plan):
        _render(master, {"Live Queue", "Order History"})
    assert "lq_op_for_order" not in seen, \
        "an unchanged, already-committed row must not be rewritten"


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
