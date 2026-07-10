"""Focused regression tests for serial, resumable backlog operation."""
from __future__ import annotations

import asyncio
import contextlib
import json
import subprocess
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, patch

import backfill_orders
import line_items
import live_master
import process_lock
import sales_orders


def test_default_search_population_starts_at_401000():
    minimum = backfill_orders.DEFAULT_CBC_SEARCH_MIN_JOB
    assert minimum == 401000
    assert not backfill_orders._inside_job_caps("400999", minimum)
    assert backfill_orders._inside_job_caps("401000", minimum)


def test_resume_trusts_only_current_serial_misses():
    version = backfill_orders.BACKFILL_SCAN_VERSION
    assert backfill_orders._is_done({"status": "ok"})
    assert not backfill_orders._is_done({"status": "error"})
    assert not backfill_orders._is_done({"status": "needs-retry-wrong-SO-quarantined"})
    assert not backfill_orders._is_done({"status": "not-found"})
    assert not backfill_orders._is_done({"status": "no-SO"})
    assert not backfill_orders._is_done({
        "status": "not-found", "backfill_scan_version": version,
        "backfill_attempts": 1})
    assert backfill_orders._is_done({
        "status": "not-found", "backfill_scan_version": version,
        "backfill_attempts": 2})
    assert backfill_orders._is_done({
        "status": "no-SO", "backfill_scan_version": version,
        "backfill_attempts": 2})
    assert not backfill_orders._is_done(
        {"status": "not-found", "backfill_scan_version": version,
         "backfill_attempts": 2},
        retry_not_found=True,
    )


def test_serial_pass_has_one_order_in_flight_and_saves_each_job():
    active = 0
    max_active = 0
    order = []
    saves = []
    publishes = []
    page = SimpleNamespace(close=AsyncMock())

    async def process(_page, _context, job, _folder, **_kwargs):
        nonlocal active, max_active
        active += 1
        max_active = max(max_active, active)
        order.append(job)
        await asyncio.sleep(0)
        active -= 1
        return {
            "job": job,
            "status": "ok",
            "scanned_at": "now",
            "backfill_scan_version": backfill_orders.BACKFILL_SCAN_VERSION,
        }

    records = {}
    state = {"processed": 0, "publish_every": 2}
    with (
        patch("backfill_orders._open_backfill_page", new=AsyncMock(return_value=page)),
        patch("backfill_orders.process_one_async", new=process),
        patch("backfill_orders.cbc_fetch_lock", side_effect=lambda: contextlib.nullcontext()),
        patch("backfill_orders.save_progress", side_effect=lambda value: saves.append(dict(value))),
        patch("backfill_orders._publish_checkpoint",
              side_effect=lambda: publishes.append(state["processed"]) or True),
    ):
        completed = asyncio.run(backfill_orders._run_serial_pass(
            object(), ["401001", "401002", "401003"], records, {}, 0, 1, 1, 1, state))

    assert completed == 3
    assert state["processed"] == 3
    assert order == ["401001", "401002", "401003"]
    assert max_active == 1
    assert len(saves) == 3
    assert publishes == [2]
    assert state["last_published"] == 2
    assert set(saves[-1]) == set(order)
    page.close.assert_awaited_once()


def test_atomic_line_item_update_preserves_existing_jobs(tmp: Path):
    path = tmp / "line_items.json"
    line_items.save_store({
        "jobs": {"401000": {"items": [{"raw": "old"}], "customer": "Existing"}},
        "ai_tags": {},
    }, path)

    items = [{"raw": "Outlet Damper L 1.00", "norm": "OUTLET DAMPER", "tags": ["DAMPER"]}]
    assert line_items.record_jobs_atomic([{
        "job": "401001", "items": items, "customer": "New", "co_number": 2,
        "so_pdf": "401001.pdf",
    }], path) == 1

    stored = line_items.load_store(path)
    assert set(stored["jobs"]) == {"401000", "401001"}
    assert stored["jobs"]["401000"]["customer"] == "Existing"
    assert stored["jobs"]["401001"]["co_number"] == 2


def test_backfill_overlay_survives_and_merges_with_main_store(tmp: Path):
    base = tmp / "line_items.json"
    overlay = tmp / "backfill_line_items.json"
    line_items.save_store({
        "jobs": {
            "401000": {"scanned_at": "2026-01-01", "items": [{"raw": "base only"}]},
            "401001": {"scanned_at": "2026-01-01", "items": [{"raw": "stale"}]},
        },
        "ai_tags": {},
    }, base)
    line_items.save_store({
        "jobs": {
            "401001": {"scanned_at": "2026-02-01", "items": [{"raw": "fresh"}]},
            "401002": {"scanned_at": "2026-02-01", "items": [{"raw": "overlay only"}]},
        },
        "ai_tags": {},
    }, overlay)

    with (
        patch("line_items.store_path", return_value=base),
        patch("line_items.backfill_store_path", return_value=overlay),
    ):
        merged = line_items.load_store()

    assert set(merged["jobs"]) == {"401000", "401001", "401002"}
    assert merged["jobs"]["401001"]["items"][0]["raw"] == "fresh"


def test_dead_session_probe_failure_warns_but_never_stops():
    page = SimpleNamespace(close=AsyncMock())

    async def all_misses(_page, _context, job, _folder, **_kwargs):
        return {"job": job, "status": "not-found", "scanned_at": "now",
                "backfill_scan_version": backfill_orders.BACKFILL_SCAN_VERSION}

    jobs = [str(401100 + i) for i in range(backfill_orders.DEAD_SESSION_STREAK * 2 + 3)]
    records = {"401001": {"status": "ok",
                          "backfill_scan_version": backfill_orders.BACKFILL_SCAN_VERSION}}
    state = {"processed": 0, "publish_every": 0}
    probes = []

    async def probe_fails(_page, job, *_a, **_k):
        probes.append(job)
        return False

    with (
        patch("backfill_orders._open_backfill_page", new=AsyncMock(return_value=page)),
        patch("backfill_orders.process_one_async", new=all_misses),
        patch("backfill_orders.open_order_detail_async", new=probe_fails),
        patch("backfill_orders._close_modal_async", new=AsyncMock()),
        patch("backfill_orders.cbc_fetch_lock", side_effect=lambda: contextlib.nullcontext()),
        patch("backfill_orders.save_progress", side_effect=lambda _v: None),
    ):
        completed = asyncio.run(backfill_orders._run_serial_pass(
            object(), jobs, records, {}, 0, 1, 1, 1, state))

    assert completed == len(jobs)               # LOG-ONLY: the run never stops
    assert state.get("dead_session_warned") is True
    assert probes == ["401001", "401001"]       # one probe per streak window


def test_dead_session_probe_success_resets_and_continues():
    page = SimpleNamespace(close=AsyncMock())

    async def all_misses(_page, _context, job, _folder, **_kwargs):
        return {"job": job, "status": "not-found", "scanned_at": "now",
                "backfill_scan_version": backfill_orders.BACKFILL_SCAN_VERSION}

    jobs = [str(401100 + i) for i in range(backfill_orders.DEAD_SESSION_STREAK + 3)]
    records = {"401001": {"status": "ok",
                          "backfill_scan_version": backfill_orders.BACKFILL_SCAN_VERSION}}
    state = {"processed": 0, "publish_every": 0}

    with (
        patch("backfill_orders._open_backfill_page", new=AsyncMock(return_value=page)),
        patch("backfill_orders.process_one_async", new=all_misses),
        patch("backfill_orders.open_order_detail_async", new=AsyncMock(return_value=True)),
        patch("backfill_orders._close_modal_async", new=AsyncMock()),
        patch("backfill_orders.cbc_fetch_lock", side_effect=lambda: contextlib.nullcontext()),
        patch("backfill_orders.save_progress", side_effect=lambda _v: None),
    ):
        completed = asyncio.run(backfill_orders._run_serial_pass(
            object(), jobs, records, {}, 0, 1, 1, 1, state))

    assert completed == len(jobs)               # thin stretch: the pass finishes
    assert not state.get("dead_session")


def test_watcher_fetch_uses_the_shared_cbc_lock():
    events = []

    @contextlib.contextmanager
    def lock():
        events.append("enter")
        try:
            yield
        finally:
            events.append("exit")

    async def run():
        with (
            patch("sales_orders.cbc_fetch_lock", side_effect=lock),
            patch("sales_orders._afetch_all_unlocked", new=AsyncMock(return_value={"401001": {}})),
        ):
            return await sales_orders._afetch_all(["401001"])

    assert asyncio.run(run()) == {"401001": {}}
    assert events == ["enter", "exit"]


def test_live_master_save_preserves_external_backfill_rows(tmp: Path):
    path = tmp / "live_master.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    external = {
        "orders": {
            "401000": {"added": "old", "on_queue": False, "job": {"job": "401000"}},
            "401001": {"added": "old", "left": "stale departure", "on_queue": True,
                       "job": {"job": "401001", "so_pdf": "verified.pdf"}},
        }
    }
    path.write_text(json.dumps(external), encoding="utf-8")
    watcher = {
        "orders": {
            "401001": {"added": "old", "left": None, "on_queue": True,
                       "job": {"job": "401001", "status_note": "live"}},
        }
    }

    with (
        patch.object(live_master, "MASTER_PATH", path),
        patch("live_master.data_file_lock", side_effect=lambda *_a, **_k: contextlib.nullcontext()),
    ):
        live_master.save_master(watcher)

    saved = json.loads(path.read_text(encoding="utf-8"))
    assert set(saved["orders"]) == {"401000", "401001"}
    assert saved["orders"]["401001"]["job"]["status_note"] == "live"
    assert saved["orders"]["401001"]["job"]["so_pdf"] == "verified.pdf"
    assert saved["orders"]["401001"]["left"] is None
    assert "401000" in watcher["orders"]


def test_kernel_lock_excludes_a_second_process(tmp: Path):
    lock_path = tmp / "shared.lock"
    child = (
        "import sys\n"
        "from pathlib import Path\n"
        "from process_lock import exclusive_file_lock\n"
        "try:\n"
        "    with exclusive_file_lock(Path(sys.argv[1]), label='test', timeout=0.2, poll=0.02):\n"
        "        pass\n"
        "except TimeoutError:\n"
        "    raise SystemExit(3)\n"
    )
    with process_lock.exclusive_file_lock(lock_path, label="parent test lock"):
        result = subprocess.run(
            [sys.executable, "-c", child, str(lock_path)],
            cwd=Path(__file__).parent,
            capture_output=True,
            text=True,
            timeout=5,
        )
    assert result.returncode == 3, (result.stdout, result.stderr)


def main() -> int:
    passed = 0
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as directory:
        tmp = Path(directory)
        for name, function in sorted(globals().items()):
            if not name.startswith("test_") or not callable(function):
                continue
            function(tmp / name) if "tmp" in function.__code__.co_varnames else function()
            print(f"  ok  {name}")
            passed += 1
    print(f"\n{passed} tests passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
