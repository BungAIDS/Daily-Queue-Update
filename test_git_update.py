"""Tests for the launcher's Git helpers (git_update.py).

These cover the pure parsing/command-building logic only; they do not touch
the network or a real repository. Run directly:

    python test_git_update.py
"""
from __future__ import annotations

import sys

import git_update


def test_normalize_strips_markers_and_remote():
    assert git_update.normalize_branch_name("* main") == "main"
    assert git_update.normalize_branch_name("  feature/x  ") == "feature/x"
    assert git_update.normalize_branch_name("  origin/dev") == "dev"
    assert git_update.normalize_branch_name("  remotes/origin/dev") == "dev"
    assert git_update.normalize_branch_name("+ wt-branch") == "wt-branch"


def test_normalize_skips_pointers_and_detached():
    assert git_update.normalize_branch_name("  origin/HEAD -> origin/main") is None
    assert git_update.normalize_branch_name("* (HEAD detached at abc1234)") is None
    assert git_update.normalize_branch_name("   ") is None
    assert git_update.normalize_branch_name("") is None


def test_normalize_respects_custom_remote():
    assert git_update.normalize_branch_name("upstream/dev", remote="upstream") == "dev"
    # The default remote prefix is left intact when a different remote is named.
    assert git_update.normalize_branch_name("origin/dev", remote="upstream") == "origin/dev"


def test_parse_branches_dedupes_and_sorts():
    local = "* main\n  feature/a\n"
    remote = "  origin/HEAD -> origin/main\n  origin/main\n  origin/feature/b\n"
    assert git_update.parse_branches(local, remote) == ["feature/a", "feature/b", "main"]


def test_build_pull_steps_same_branch_skips_checkout():
    steps = git_update.build_pull_steps("main", "main", switch=True)
    assert steps == [["fetch", "origin", "main"], ["pull", "origin", "main"]]


def test_build_pull_steps_switch_adds_checkout():
    steps = git_update.build_pull_steps("dev", "main", switch=True)
    assert steps == [
        ["fetch", "origin", "dev"],
        ["checkout", "dev"],
        ["pull", "origin", "dev"],
    ]


def test_build_pull_steps_no_switch_omits_checkout():
    steps = git_update.build_pull_steps("dev", "main", switch=False)
    assert steps == [["fetch", "origin", "dev"], ["pull", "origin", "dev"]]


def test_build_pull_steps_requires_a_branch():
    for bad in ("", "   "):
        try:
            git_update.build_pull_steps(bad, "main")
        except ValueError:
            pass
        else:
            raise AssertionError(f"expected ValueError for branch {bad!r}")


def test_launcher_needs_restart_detects_program_files():
    assert git_update.launcher_needs_restart(["launcher.py"]) is True
    assert git_update.launcher_needs_restart(["git_update.py"]) is True
    # Match on the file name even when git reports a path.
    assert git_update.launcher_needs_restart(["subdir/launcher.py", "watch.py"]) is True


def test_launcher_needs_restart_ignores_other_files():
    assert git_update.launcher_needs_restart([]) is False
    assert git_update.launcher_needs_restart(["watch.py", "config.py", "README.md"]) is False


def test_run_pull_steps_stops_at_first_failure():
    calls: list[list[str]] = []

    def fake_stream(args, on_line):
        calls.append(args)
        on_line(f"ran {' '.join(args)}\n")
        return 0 if args[0] == "fetch" else 1  # checkout fails

    original = git_update.stream_git
    git_update.stream_git = fake_stream
    try:
        lines: list[str] = []
        code = git_update.run_pull_steps(
            [["fetch", "origin", "dev"], ["checkout", "dev"], ["pull", "origin", "dev"]],
            lines.append,
        )
    finally:
        git_update.stream_git = original

    assert code == 1
    # fetch ran, checkout ran and failed, pull was never reached.
    assert calls == [["fetch", "origin", "dev"], ["checkout", "dev"]]


def test_run_pull_steps_runs_all_on_success():
    calls: list[list[str]] = []

    def fake_stream(args, on_line):
        calls.append(args)
        return 0

    original = git_update.stream_git
    git_update.stream_git = fake_stream
    try:
        code = git_update.run_pull_steps(
            [["fetch", "origin", "dev"], ["checkout", "dev"], ["pull", "origin", "dev"]],
            lambda _line: None,
        )
    finally:
        git_update.stream_git = original

    assert code == 0
    assert len(calls) == 3


class _FakeProc:
    def __init__(self, returncode=0, stdout="", stderr=""):
        self.returncode = returncode
        self.stdout = stdout
        self.stderr = stderr


def _run_publish_with(fake):
    calls: list[list[str]] = []

    def wrapper(args, *, input_text=None, timeout=None, env_extra=None):
        calls.append(args)
        return fake(args)

    original = git_update._git_text
    git_update._git_text = wrapper
    try:
        ok, detail = git_update.publish_report("report body\n", branch="debug/launcher")
    finally:
        git_update._git_text = original
    return ok, detail, calls


def test_publish_report_existing_branch_fast_forwards():
    def fake(args):
        if args[0] == "fetch":
            return _FakeProc(0)
        if args[:2] == ["rev-parse", "FETCH_HEAD"]:
            return _FakeProc(0, "PARENTSHA\n")
        if args[0] == "hash-object":
            return _FakeProc(0, "BLOBSHA\n")
        if args[0] == "write-tree":
            return _FakeProc(0, "TREESHA\n")
        if args[0] == "commit-tree":
            return _FakeProc(0, "COMMITSHA\n")
        return _FakeProc(0)

    ok, detail, calls = _run_publish_with(fake)
    assert ok, detail
    commit = next(c for c in calls if c[0] == "commit-tree")
    assert "-p" in commit and "PARENTSHA" in commit  # builds on the existing tip
    assert any(c[0] == "read-tree" for c in calls)    # preserves the branch's tree
    push = next(c for c in calls if c[0] == "push")
    assert push[-1] == "COMMITSHA:refs/heads/debug/launcher"


def test_publish_report_new_branch_has_no_parent():
    def fake(args):
        if args[0] == "fetch":
            return _FakeProc(1, stderr="couldn't find remote ref")
        if args[0] == "hash-object":
            return _FakeProc(0, "BLOB\n")
        if args[0] == "write-tree":
            return _FakeProc(0, "TREE\n")
        if args[0] == "commit-tree":
            return _FakeProc(0, "COMMIT\n")
        return _FakeProc(0)

    ok, detail, calls = _run_publish_with(fake)
    assert ok, detail
    commit = next(c for c in calls if c[0] == "commit-tree")
    assert "-p" not in commit                          # orphan first commit
    assert not any(c[0] == "read-tree" for c in calls)  # nothing to base on


def test_publish_report_reports_push_failure():
    def fake(args):
        if args[0] == "fetch":
            return _FakeProc(0)
        if args[:2] == ["rev-parse", "FETCH_HEAD"]:
            return _FakeProc(0, "P\n")
        if args[0] == "hash-object":
            return _FakeProc(0, "B\n")
        if args[0] == "write-tree":
            return _FakeProc(0, "T\n")
        if args[0] == "commit-tree":
            return _FakeProc(0, "C\n")
        if args[0] == "push":
            return _FakeProc(1, stderr="permission denied")
        return _FakeProc(0)

    ok, detail, _calls = _run_publish_with(fake)
    assert ok is False
    assert "push failed" in detail and "permission denied" in detail


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
