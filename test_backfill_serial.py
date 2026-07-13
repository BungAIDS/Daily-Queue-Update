"""Focused regression tests for serial, resumable backlog operation."""
from __future__ import annotations

import asyncio
import contextlib
import json
import subprocess
import sys
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


def test_dead_session_probe_has_a_bootstrap_known_good_order():
    assert backfill_orders._probe_job({}) == "401468"


def test_resume_trusts_only_current_serial_misses():
    version = backfill_orders.BACKFILL_SCAN_VERSION
    assert backfill_orders._is_done({"status": "ok"})
    assert not backfill_orders._is_done({"status": "error"})
    assert not backfill_orders._is_done({"status": "needs-retry-wrong-SO-quarantined"})
    assert not backfill_orders._is_done({"status": "search-state-retry"})
    assert not backfill_orders._is_done({"status": "session-failed"})
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


def test_missing_search_input_is_not_reported_as_job_not_found():
    page = SimpleNamespace()
    with (
        patch("backfill_orders._close_modal_async", new=AsyncMock()),
        patch("backfill_orders.find_search_box_async", new=AsyncMock(return_value=None)),
    ):
        try:
            asyncio.run(backfill_orders.open_order_detail_async(page, "413393", 1, 1))
            assert False, "a broken search page must not become a not-found result"
        except backfill_orders.SearchPageError:
            pass


def test_process_one_propagates_a_broken_search_page_to_the_runner():
    page = SimpleNamespace()
    with (
        patch("backfill_orders.open_order_detail_async",
              new=AsyncMock(side_effect=backfill_orders.SearchPageError("stale input"))),
        patch("backfill_orders._close_modal_async", new=AsyncMock()),
    ):
        try:
            asyncio.run(backfill_orders.process_one_async(page, object(), "413393"))
            assert False, "the serial runner must receive the broken-page signal"
        except backfill_orders.SearchPageError:
            pass


def test_serial_pass_reopens_a_broken_search_page_and_retries_same_job():
    page = SimpleNamespace(close=AsyncMock())
    fresh_page = SimpleNamespace(close=AsyncMock())
    calls = []

    async def process(active_page, _context, job, _folder, **_kwargs):
        calls.append((active_page, job))
        if active_page is page:
            raise backfill_orders.SearchPageError("detached search input")
        return {
            "job": job,
            "status": "ok",
            "scanned_at": "now",
            "backfill_scan_version": backfill_orders.BACKFILL_SCAN_VERSION,
        }

    records = {}
    state = {"processed": 0, "publish_every": 0}
    with (
        patch("backfill_orders._open_backfill_page",
              new=AsyncMock(side_effect=[page, fresh_page])),
        patch("backfill_orders.process_one_async", new=process),
        patch("backfill_orders.cbc_fetch_lock", side_effect=lambda: contextlib.nullcontext()),
        patch("backfill_orders.save_progress", side_effect=lambda _v: None),
    ):
        completed = asyncio.run(backfill_orders._run_serial_pass(
            object(), ["413393"], records, {}, 0, 1, 1, 1, state))

    assert completed == 1
    assert records["413393"]["status"] == "ok"
    assert records["413393"]["backfill_attempts"] == 1
    assert calls == [(page, "413393"), (fresh_page, "413393")]
    page.close.assert_awaited_once()
    fresh_page.close.assert_awaited_once()


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


def test_dead_session_probe_failure_invalidates_streak_and_stops():
    page = SimpleNamespace(close=AsyncMock())
    fresh_page = SimpleNamespace(close=AsyncMock())

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
        patch("backfill_orders._open_backfill_page",
              new=AsyncMock(side_effect=[page, fresh_page])),
        patch("backfill_orders.process_one_async", new=all_misses),
        patch("backfill_orders.open_order_detail_async", new=probe_fails),
        patch("backfill_orders._close_modal_async", new=AsyncMock()),
        patch("backfill_orders.cbc_fetch_lock", side_effect=lambda: contextlib.nullcontext()),
        patch("backfill_orders.save_progress", side_effect=lambda _v: None),
    ):
        try:
            asyncio.run(backfill_orders._run_serial_pass(
                object(), jobs, records, {}, 0, 1, 1, 1, state))
            assert False, "dead CBC session should stop the serial pass"
        except backfill_orders.DeadSessionError:
            pass

    assert state["processed"] == backfill_orders.DEAD_SESSION_STREAK
    assert state.get("dead_session") == 1
    assert probes == ["401001", "401001"]       # current page, then fresh page
    assert all(records[j]["status"] == "session-failed"
               for j in jobs[:backfill_orders.DEAD_SESSION_STREAK])
    assert all(not backfill_orders._is_done(records[j])
               for j in jobs[:backfill_orders.DEAD_SESSION_STREAK])
    page.close.assert_awaited_once()
    fresh_page.close.assert_awaited_once()


def test_dead_session_fresh_page_recovery_marks_streak_for_retry():
    page = SimpleNamespace(close=AsyncMock())
    fresh_page = SimpleNamespace(close=AsyncMock())

    async def all_misses(_page, _context, job, _folder, **_kwargs):
        return {"job": job, "status": "not-found", "scanned_at": "now",
                "backfill_scan_version": backfill_orders.BACKFILL_SCAN_VERSION}

    jobs = [str(401100 + i) for i in range(backfill_orders.DEAD_SESSION_STREAK + 3)]
    records = {"401001": {"status": "ok",
                          "backfill_scan_version": backfill_orders.BACKFILL_SCAN_VERSION}}
    state = {"processed": 0, "publish_every": 0}

    with (
        patch("backfill_orders._open_backfill_page",
              new=AsyncMock(side_effect=[page, fresh_page])),
        patch("backfill_orders.process_one_async", new=all_misses),
        patch("backfill_orders.open_order_detail_async",
              new=AsyncMock(side_effect=[False, True])),
        patch("backfill_orders._close_modal_async", new=AsyncMock()),
        patch("backfill_orders.cbc_fetch_lock", side_effect=lambda: contextlib.nullcontext()),
        patch("backfill_orders.save_progress", side_effect=lambda _v: None),
    ):
        completed = asyncio.run(backfill_orders._run_serial_pass(
            object(), jobs, records, {}, 0, 1, 1, 1, state))

    assert completed == len(jobs)
    assert all(records[j]["status"] == "search-state-retry"
               for j in jobs[:backfill_orders.DEAD_SESSION_STREAK])
    assert all(records[j]["status"] == "not-found"
               for j in jobs[backfill_orders.DEAD_SESSION_STREAK:])
    assert not state.get("dead_session")


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


def test_excel_safe_value_removes_only_illegal_control_characters():
    assert backfill_orders._excel_safe_value("A\x00B\x07C\nD\tE") == "ABC\nD\tE"
    assert backfill_orders._excel_safe_value(42) == 42


def test_quote_run_parse_gets_a_freshness_timestamp():
    rec = {}
    with patch("backfill_orders.parse_quote_run", return_value={
        "fields": {"Size": "27"}, "summary": "Size=27", "template": "pdf",
    }):
        backfill_orders._attach_quote_run_parse(rec, "run-without-an-extension")

    assert rec["drive_run"] == {"Size": "27"}
    assert rec["drive_run_template"] == "pdf"
    assert rec["drive_run_parsed_at"].startswith("20")


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


def test_live_master_save_accepts_a_fresher_external_drive_run_parse(tmp: Path):
    path = tmp / "live_master.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    external = {
        "orders": {
            "401001": {
                "added": "old",
                "on_queue": False,
                "job": {
                    "job": "401001",
                    "drive_run": {},
                    "drive_run_summary": "",
                    "drive_run_template": "unknown",
                    "drive_run_parsed_at": "2026-07-13T08:00:00",
                },
            },
        },
    }
    path.write_text(json.dumps(external), encoding="utf-8")
    watcher = {
        "orders": {
            "401001": {
                "added": "old",
                "on_queue": False,
                "job": {
                    "job": "401001",
                    "status_note": "live",
                    "drive_run": {"binary": "garbage"},
                    "drive_run_summary": "bad bytes",
                    "drive_run_template": "generic_text",
                },
            },
        },
    }

    with (
        patch.object(live_master, "MASTER_PATH", path),
        patch("live_master.data_file_lock", side_effect=lambda *_a, **_k: contextlib.nullcontext()),
    ):
        live_master.save_master(watcher)

    saved = json.loads(path.read_text(encoding="utf-8"))
    job = saved["orders"]["401001"]["job"]
    assert job["status_note"] == "live"
    assert job["drive_run"] == {}
    assert job["drive_run_summary"] == ""
    assert job["drive_run_template"] == "unknown"
    assert job["drive_run_parsed_at"] == "2026-07-13T08:00:00"


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
    tmp = Path.cwd() / ".tmp_backfill_serial_tests"
    tmp.mkdir(exist_ok=True)
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
