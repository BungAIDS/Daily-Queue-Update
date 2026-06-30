"""Tests for the launcher's external-process parsing (procscan.py).

    python test_procscan.py
"""
from __future__ import annotations

import sys

import procscan


def test_matches_script_by_name():
    cmd = r'"C:\proj\venv\Scripts\python.exe" "C:\proj\watch.py" --now'
    assert procscan.process_line_matches_script(cmd, "watch.py") is True
    assert procscan.process_line_matches_script(cmd, "send.py") is False
    # A different script that merely contains the name should not match.
    assert procscan.process_line_matches_script("python watchdog.py", "watch.py") is False


def test_parse_powershell_line_pid_first():
    pid, cmd = procscan.parse_scan_line(
        r'4321 "C:\proj\python.exe" "C:\proj\watch.py" --now', "powershell-cim"
    )
    assert pid == 4321
    assert cmd.endswith("watch.py\" --now")


def test_parse_ps_line_pid_first():
    pid, cmd = procscan.parse_scan_line("  1234 /usr/bin/python /proj/watch.py", "ps")
    assert pid == 1234
    assert cmd == "/usr/bin/python /proj/watch.py"


def test_parse_wmic_line_pid_trailing():
    # wmic prints CommandLine first, then the ProcessId column.
    pid, cmd = procscan.parse_scan_line(
        r'"C:\proj\python.exe" "C:\proj\watch.py" 421473        5678', "wmic"
    )
    assert pid == 5678  # the trailing integer is the PID, not the job number in the args
    assert "watch.py" in cmd and "421473" in cmd


def test_parse_header_and_blank_lines_are_skipped():
    assert procscan.parse_scan_line("CommandLine                 ProcessId", "wmic") == (None, "CommandLine                 ProcessId")
    assert procscan.parse_scan_line("   ", "powershell-cim") == (None, "")


def test_parse_scan_output_filters_and_pairs():
    raw = [
        r'10 "C:\python.exe" "C:\watch.py" --now',
        r'11 "C:\python.exe" "C:\send.py"',
        r'12 C:\Windows\explorer.exe',          # no python -> dropped
        "   ",                                     # blank -> dropped
    ]
    pairs = procscan.parse_scan_output(raw, "powershell-cim")
    assert pairs == [(10, '"C:\\python.exe" "C:\\watch.py" --now'), (11, '"C:\\python.exe" "C:\\send.py"')]


def test_parse_scan_output_keeps_unknown_pid():
    # A python line we couldn't get a PID for is still returned (pid None).
    raw = ["CommandLine python watch.py ProcessId"]  # wmic header-ish, no trailing int
    pairs = procscan.parse_scan_output(raw, "wmic")
    assert pairs == [(None, "CommandLine python watch.py ProcessId")]


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
