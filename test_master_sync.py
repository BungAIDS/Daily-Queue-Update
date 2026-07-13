"""Tests for the per-source mergers that fold helper stores into the master
(master_sync.py). Writes tiny temp store files to the real backlog paths and
cleans them up.

    python test_master_sync.py
"""
from __future__ import annotations

import json
import sys

import master_sync


def _write(path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj), encoding="utf-8")


def test_merge_autocad():
    p = master_sync._AUTOCAD
    existed = p.exists()
    backup = p.read_text() if existed else None
    try:
        _write(p, {"421000": {"job": "421000", "type": "GENERAL LINE",
                              "folder": "Z:\\x", "extras": {"51": "DWG"},
                              "missing_std": False}})
        m = {"orders": {}}
        assert master_sync.merge_autocad(m) == 1
        job = m["orders"]["421000"]["job"]
        assert job["dwg_extras"] == {"51": "DWG"} and job["job_type"] == "GENERAL LINE"
    finally:
        if backup is not None:
            p.write_text(backup)
        elif p.exists():
            p.unlink()


def test_merge_backfill():
    p = master_sync._BACKFILL
    existed = p.exists()
    backup = p.read_text() if existed else None
    try:
        _write(p, {
            "421000": {"job": "421000", "status": "ok", "scanned_at": "t",
                       "backfill_scan_version": "serial-verified-v1",
                       "backfill_attempts": 2,
                       "so_size": "27", "so_arrangement": "9", "co_number": 2},
            "421999": {"job": "421999", "status": "needs-retry-wrong-SO-quarantined",
                       "so_pdf": "wrong.pdf", "so_size": "WRONG"},
        })
        m = {"orders": {}}
        assert master_sync.merge_backfill(m) == 1
        job = m["orders"]["421000"]["job"]
        assert job["so_size"] == "27" and job["co_number"] == 2
        assert "status" not in job and "scanned_at" not in job   # metadata skipped
        assert "backfill_scan_version" not in job
        assert "backfill_attempts" not in job
        assert "421999" not in m["orders"]
    finally:
        if backup is not None:
            p.write_text(backup)
        elif p.exists():
            p.unlink()


def test_merge_backfill_skips_scans_older_than_live_verification():
    p = master_sync._BACKFILL
    existed = p.exists()
    backup = p.read_text() if existed else None
    try:
        _write(p, {
            "421000": {"job": "421000", "status": "ok",
                       "scanned_at": "2026-07-12T09:00:00", "so_size": "27"},
            "421001": {"job": "421001", "status": "ok",
                       "scanned_at": "2026-07-12T09:00:00", "so_size": "31"},
            "421002": {"job": "421002", "status": "ok",
                       "scanned_at": "2026-07-12T09:00:00", "so_size": "44"},
        })
        m = {"orders": {
            # Live-verified AFTER the scan -> the older scan must not clobber it.
            "421000": {"job": {"job": "421000", "so_size": "245",
                               "so_verified_at": "2026-07-13T06:30:00"}},
            # Live-verified BEFORE the scan -> the scan is fresher and merges.
            "421001": {"job": {"job": "421001", "so_size": "30",
                               "so_verified_at": "2026-07-11T12:00:00"}},
            # No verification stamp -> merges as before.
            "421002": {"job": {"job": "421002", "so_size": "43"}},
        }}
        assert master_sync.merge_backfill(m) == 2
        assert m["orders"]["421000"]["job"]["so_size"] == "245"   # untouched
        assert m["orders"]["421001"]["job"]["so_size"] == "31"    # updated
        assert m["orders"]["421002"]["job"]["so_size"] == "44"    # updated
    finally:
        if backup is not None:
            p.write_text(backup)
        elif p.exists():
            p.unlink()


def test_merge_quote_runs():
    p = master_sync._QUOTE_RUNS
    existed = p.exists()
    backup = p.read_text() if existed else None
    try:
        _write(p, {"421000": {"job": "421000", "type": "GL", "folder": "Z:\\x",
                              "runs": [{"template": "cbc_qt_run_text", "summary": "Size=37",
                                        "fields": {"Size": "37"}, "path": "Z:\\x\\run.txt"}]}})
        m = {"orders": {}}
        assert master_sync.merge_quote_runs(m) == 1
        job = m["orders"]["421000"]["job"]
        assert job["has_drive_run"] is True and job["drive_run_template"] == "cbc_qt_run_text"
        assert job["drive_run"] == {"Size": "37"}
    finally:
        if backup is not None:
            p.write_text(backup)
        elif p.exists():
            p.unlink()


def test_merge_quote_runs_ranks_revisions_and_keeps_history():
    p = master_sync._QUOTE_RUNS
    existed = p.exists()
    backup = p.read_text() if existed else None
    try:
        # Alphabetically the 2021 base run comes first; CO#1 must lead anyway.
        _write(p, {"400567": {"job": "400567", "type": "GL", "folder": "Z:\\x", "runs": [
            {"file": "QT RUN 8-11-21.pdf", "path": "Z:\\x\\QT RUN 8-11-21.pdf",
             "template": "pdf_vision", "summary": "old", "status": "OK",
             "fields": {"Size": "20", "RPM": "1770"}, "mtime": 1_600_000_000},
            {"file": "Qt Run CO#1.txt", "path": "Z:\\x\\Qt Run CO#1.txt",
             "template": "cbc_qt_run_text", "summary": "new", "status": "OK",
             "fields": {"Size": "20", "RPM": "1850"}, "mtime": 1_650_000_000},
        ]}})
        m = {"orders": {}}
        assert master_sync.merge_quote_runs(m) == 1
        job = m["orders"]["400567"]["job"]
        assert job["drive_run"]["RPM"] == "1850"              # CO#1 leads, not the 2021 run
        assert job["drive_run_pdf"].endswith("CO#1.txt")
        assert job["drive_run_count"] == 2
        hist = job["drive_runs"]                              # full history kept, ranked
        assert [h["file"] for h in hist] == ["Qt Run CO#1.txt", "QT RUN 8-11-21.pdf"]
        assert hist[1]["fields"]["RPM"] == "1770"             # superseded values queryable
    finally:
        if backup is not None:
            p.write_text(backup)
        elif p.exists():
            p.unlink()


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
