"""Tests for the launcher<->script stop signal (stop_signal.py).

    python test_stop_signal.py
"""
from __future__ import annotations

import os
import sys

import stop_signal


def test_request_check_clear_roundtrip():
    pid = 999_000 + (os.getpid() % 1000)  # unlikely to collide with a real process
    stop_signal.clear_stop(pid)
    try:
        assert stop_signal.stop_requested(pid) is False
        assert stop_signal.request_stop(pid) is True
        assert stop_signal.stop_requested(pid) is True
        stop_signal.clear_stop(pid)
        assert stop_signal.stop_requested(pid) is False
    finally:
        stop_signal.clear_stop(pid)


def test_paths_are_pid_specific():
    assert stop_signal.stop_path(111) != stop_signal.stop_path(222)
    assert stop_signal.stop_path(111).name == "111.stop"


def test_clear_is_safe_when_absent():
    pid = 999_777
    stop_signal.clear_stop(pid)        # no file yet
    stop_signal.clear_stop(pid)        # still fine


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
