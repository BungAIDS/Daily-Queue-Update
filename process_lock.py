"""Small cross-process locks for the watcher and backlog helpers.

The lock files are permanent markers; the operating system owns the actual
lock and releases it automatically if a process exits or is killed.
"""
from __future__ import annotations

import contextlib
import logging
import os
import time
from pathlib import Path
from typing import Iterator

from config import BACKLOG_DIR


_LOCK_DIR = BACKLOG_DIR / ".process_locks"
_CBC_FETCH_LOCK = _LOCK_DIR / "cbc_sales_order_fetch.lock"


def _lock_byte(handle) -> None:
    handle.seek(0)
    if os.name == "nt":
        import msvcrt

        msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
    else:
        import fcntl

        fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)


def _unlock_byte(handle) -> None:
    handle.seek(0)
    if os.name == "nt":
        import msvcrt

        msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
    else:
        import fcntl

        fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


@contextlib.contextmanager
def exclusive_file_lock(lock_path: Path, *, label: str, timeout: float = 900.0,
                        poll: float = 0.25) -> Iterator[None]:
    """Hold one kernel-backed byte lock, waiting up to ``timeout`` seconds."""
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    with lock_path.open("a+b") as handle:
        handle.seek(0, os.SEEK_END)
        if handle.tell() == 0:
            handle.write(b"\0")
            handle.flush()

        deadline = time.monotonic() + timeout
        announced = False
        while True:
            try:
                _lock_byte(handle)
                break
            except OSError:
                if time.monotonic() >= deadline:
                    raise TimeoutError(f"Timed out waiting for {label}")
                if not announced:
                    logging.getLogger(__name__).info(
                        "Waiting for %s; the other process can keep running.", label)
                    announced = True
                time.sleep(poll)
        try:
            yield
        finally:
            _unlock_byte(handle)


def data_file_lock(path: Path, *, label: str, timeout: float = 900.0):
    """Return a lock dedicated to a load/modify/save transaction for ``path``."""
    lock_path = path.parent / ".process_locks" / (path.name + ".lock")
    return exclusive_file_lock(lock_path, label=label, timeout=timeout)


def cbc_fetch_lock(*, timeout: float = 900.0):
    """Serialize CBC Sales Order retrieval between watch.py and the backfill."""
    return exclusive_file_lock(
        _CBC_FETCH_LOCK,
        label="CBC Sales Order access (watch.py/backfill_orders.py)",
        timeout=timeout,
    )
