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
        _write(p, {"421000": {"job": "421000", "status": "ok", "scanned_at": "t",
                              "so_size": "27", "so_arrangement": "9", "co_number": 2}})
        m = {"orders": {}}
        assert master_sync.merge_backfill(m) == 1
        job = m["orders"]["421000"]["job"]
        assert job["so_size"] == "27" and job["co_number"] == 2
        assert "status" not in job and "scanned_at" not in job   # metadata skipped
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
