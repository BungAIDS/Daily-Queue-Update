"""Handle a GL Queue Explorer ``glqtransmittal:`` order link.

The Explorer is a static file page, so Windows needs a small per-user protocol
handler to bridge an explicit button click to the local, review-only
``fill_transmittal_insider.py`` workflow.  The URI payload is deliberately
digits-only and is never interpreted as a command.  This module uses only the
standard library so it is safe to import from the Explorer/launcher path.
"""
from __future__ import annotations

import contextlib
import json
import os
import re
import subprocess
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import Callable, TextIO


ROOT = Path(__file__).resolve().parent
SCHEME = "glqtransmittal"
_ORDER_RE = re.compile(r"\d{4,12}")
_MUTEX_NAME = r"Local\CBCDailyQueuePrepareTransmittal"


def parse_order_uri(uri: str) -> str:
    """Return the numeric order from a strict ``glqtransmittal:<order>`` URI."""
    value = str(uri).strip()
    prefix = f"{SCHEME}:"
    if value[:len(prefix)].lower() != prefix:
        raise ValueError(f"expected a {prefix}<order> link")
    order = value[len(prefix):]
    if not _ORDER_RE.fullmatch(order):
        raise ValueError("the transmittal order must contain digits only")
    return order


def console_python_path(root: Path = ROOT, executable: str | Path | None = None) -> Path:
    """Prefer this checkout's console Python, never a windowless pythonw.exe."""
    for candidate in (
        root / "venv" / "Scripts" / "python.exe",
        root / ".venv" / "Scripts" / "python.exe",
    ):
        if candidate.exists():
            return candidate
    current = Path(executable or sys.executable)
    if current.name.lower() == "pythonw.exe":
        console = current.with_name("python.exe")
        if console.exists():
            return console
    return current


def protocol_command(
    root: Path = ROOT,
    executable: str | Path | None = None,
) -> str:
    """Windows registry command line for this handler; ``%1`` is the URI."""
    return subprocess.list2cmdline([
        str(console_python_path(root, executable)),
        str(root / Path(__file__).name),
        "%1",
    ])


def saved_initials(root: Path = ROOT) -> str:
    """Reuse the launcher's saved Email Drawings signature, when it is safe."""
    try:
        state = json.loads((root / ".launcher_state.json").read_text(encoding="utf-8"))
        value = str(state["options"]["email_drawings"].get("initials") or "").strip()
    except (OSError, KeyError, TypeError, ValueError, json.JSONDecodeError):
        return ""
    return value if re.fullmatch(r"[A-Za-z]{1,10}", value) else ""


@contextlib.contextmanager
def transmittal_instance_lock():
    """A crash-safe Windows single-instance gate for Word/browser preparation."""
    if os.name != "nt":
        yield True
        return

    import ctypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateMutexW.argtypes = (ctypes.c_void_p, ctypes.c_bool, ctypes.c_wchar_p)
    kernel32.CreateMutexW.restype = ctypes.c_void_p
    kernel32.CloseHandle.argtypes = (ctypes.c_void_p,)
    kernel32.CloseHandle.restype = ctypes.c_bool
    handle = kernel32.CreateMutexW(None, False, _MUTEX_NAME)
    if not handle:
        raise ctypes.WinError(ctypes.get_last_error())
    if ctypes.get_last_error() == 183:  # ERROR_ALREADY_EXISTS
        kernel32.CloseHandle(handle)
        yield False
        return
    try:
        yield True
    finally:
        kernel32.CloseHandle(handle)


class _Tee:
    """Write the run to its visible console and the launcher's normal log."""

    def __init__(self, *streams: TextIO) -> None:
        self.streams = streams

    def write(self, value: str) -> int:
        for stream in self.streams:
            stream.write(value)
        return len(value)

    def flush(self) -> None:
        for stream in self.streams:
            stream.flush()

    def isatty(self) -> bool:
        return any(getattr(stream, "isatty", lambda: False)() for stream in self.streams)

    @property
    def encoding(self) -> str:
        return getattr(self.streams[0], "encoding", None) or "utf-8"


def run_transmittal(
    order: str,
    *,
    root: Path = ROOT,
    fill_main: Callable[[list[str]], int] | None = None,
) -> int:
    """Run the existing review form and mirror its output to an Email Drawings log."""
    root = root.resolve()
    log_dir = root / "launcher_logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    log_path = log_dir / f"{stamp}_{os.getpid()}_email_drawings.log"
    error = None
    result = 1
    previous_cwd = Path.cwd()
    try:
        # Match launcher.py's cwd=ROOT behavior. Config paths such as the saved
        # CBC Insider session are intentionally repo-relative.
        os.chdir(root)
        with log_path.open("w", encoding="utf-8", errors="replace") as log_file:
            stdout = _Tee(sys.stdout, log_file)
            stderr = _Tee(sys.stderr, log_file)
            with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
                fill_args = [order]
                initials = saved_initials(root)
                if initials:
                    fill_args.extend(["--initials", initials])
                command = [
                    str(console_python_path(root)),
                    str(root / "fill_transmittal_insider.py"),
                    *fill_args,
                ]
                print(f"$ {subprocess.list2cmdline(command)}\n")
                print(f"[order page] Preparing transmittal for order {order}.")
                try:
                    if fill_main is None:
                        from fill_transmittal_insider import main as fill_main
                    result = int(fill_main(fill_args) or 0)
                except Exception:  # keep the failure visible and in the run log
                    error = sys.exc_info()
                    traceback.print_exc()
    finally:
        os.chdir(previous_cwd)
    if error is not None:
        raise error[1].with_traceback(error[2])
    return result


def _pause_after_error() -> None:
    try:
        if sys.stdin and sys.stdin.isatty():
            input("Press Enter to close this window...")
    except (EOFError, OSError):
        pass


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if len(args) != 1:
        print(f"Usage: {Path(__file__).name} {SCHEME}:<order>")
        _pause_after_error()
        return 2
    try:
        order = parse_order_uri(args[0])
        with transmittal_instance_lock() as acquired:
            if not acquired:
                print("A transmittal is already being prepared. Finish or close that "
                      "review before starting another one.")
                _pause_after_error()
                return 3
            return run_transmittal(order)
    except Exception as exc:  # the detailed traceback is already in the run log
        print(f"\nCould not prepare the transmittal: {exc}")
        _pause_after_error()
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
