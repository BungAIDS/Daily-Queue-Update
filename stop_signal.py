"""File-based stop signal between the launcher and the scripts it runs.

The desktop launcher runs under ``pythonw`` (a windowless GUI), so it has no
console and cannot deliver Ctrl+C / Ctrl+Break to a child process to trigger a
graceful shutdown. Instead it drops a tiny flag file here; a long-running script
(e.g. ``watch.py``) watches for the flag and runs its own clean-exit path
(finish the current poll, save state, publish logs, exit).

Import-light (standard library only) so both sides can use it and it is easy to
unit-test. The flag is keyed by the child's PID so multiple runs never collide.
"""
from __future__ import annotations

import os
from pathlib import Path


STOP_DIR = Path(__file__).resolve().parent / ".launcher_stops"


def stop_path(pid: int) -> Path:
    """Path of the stop-flag file for a given process id."""
    return STOP_DIR / f"{pid}.stop"


def request_stop(pid: int) -> bool:
    """Ask the process with this PID to stop. Returns True if the flag was written."""
    try:
        STOP_DIR.mkdir(parents=True, exist_ok=True)
        stop_path(pid).write_text("stop", encoding="utf-8")
        return True
    except OSError:
        return False


def stop_requested(pid: int) -> bool:
    """True if a stop has been requested for this PID."""
    return stop_path(pid).exists()


def clear_stop(pid: int) -> None:
    """Remove the stop-flag file (best effort) — call on startup and on exit."""
    try:
        stop_path(pid).unlink()
    except OSError:
        pass
