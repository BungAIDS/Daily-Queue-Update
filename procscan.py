"""Parsing for the launcher's external-process detection.

Import-light (standard library only, no tkinter) so it can be unit-tested. The
actual `wmic` / PowerShell / `ps` subprocess calls live in `launcher.py`; this
module only turns their output into `(pid, command_line)` pairs and matches a
command line against a script name.
"""
from __future__ import annotations

import re
from pathlib import Path


def process_line_matches_script(line: str, script: str) -> bool:
    """True when a process command line refers to ``script`` (by file name)."""
    script_name = re.escape(Path(script).name.lower())
    pattern = rf'(?:^|[\s"\'/\\]){script_name}(?:$|[\s"\'])'
    return bool(re.search(pattern, line.lower()))


def parse_scan_line(line: str, method: str) -> tuple[int | None, str]:
    """Split one scan output line into ``(pid, command_line)``.

    The PID is leading for the ``powershell-cim`` and ``ps`` formats (we ask for
    it first) and trailing for ``wmic`` (its ``ProcessId`` column sorts after
    ``CommandLine``). Returns ``(None, …)`` for header/blank/unparseable lines.
    """
    stripped = line.strip()
    if not stripped:
        return None, ""
    if method in ("powershell-cim", "ps"):
        head, _, rest = stripped.partition(" ")
        if head.isdigit():
            return int(head), rest.strip()
        return None, stripped
    if method == "wmic":
        match = re.search(r"(\d+)\s*$", line)
        if match:
            return int(match.group(1)), line[: match.start()].strip()
        return None, stripped
    return None, stripped


def parse_scan_output(
    raw_lines: list[str], method: str, *, keyword: str = "python"
) -> list[tuple[int | None, str]]:
    """Parse scan output into ``(pid, command_line)`` pairs for matching lines.

    Only lines whose command line contains ``keyword`` (default ``python``) are
    kept, since those are the ones that could be one of our scripts.
    """
    out: list[tuple[int | None, str]] = []
    for line in raw_lines:
        pid, cmd = parse_scan_line(line, method)
        if cmd and keyword in cmd.lower():
            out.append((pid, cmd))
    return out
