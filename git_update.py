"""Git helpers for the desktop launcher's "Git Update" feature.

Kept deliberately import-light (standard library only, no tkinter) so the
branch/pull logic can be unit-tested on any machine and exercised in CI without
a display. The Tk dialog that drives these helpers lives in ``launcher.py``.
"""
from __future__ import annotations

import os
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Callable


ROOT = Path(__file__).resolve().parent
DEFAULT_REMOTE = "origin"

# Files that make up the launcher program itself. A running launcher loads these
# into memory at startup, so a pull that changes them only takes effect after the
# launcher is closed and reopened.
LAUNCHER_FILES = ("launcher.py", "git_update.py")


def _hidden_console_kwargs() -> dict[str, object]:
    """On Windows, keep git from flashing its own console window."""
    if os.name != "nt":
        return {}
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = 0
    return {
        "creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0),
        "startupinfo": startupinfo,
    }


def run_git(args: list[str], *, timeout: float | None = None) -> subprocess.CompletedProcess:
    """Run a git command in the project root and capture its output."""
    return subprocess.run(
        ["git", *args],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        **_hidden_console_kwargs(),
    )


def git_available() -> bool:
    """True when a usable ``git`` executable is on PATH."""
    try:
        result = run_git(["--version"], timeout=5)
    except (OSError, subprocess.SubprocessError):
        return False
    return result.returncode == 0


def is_git_repo() -> bool:
    """True when the project root is inside a git work tree."""
    try:
        result = run_git(["rev-parse", "--is-inside-work-tree"], timeout=5)
    except (OSError, subprocess.SubprocessError):
        return False
    return result.returncode == 0 and result.stdout.strip() == "true"


def current_branch() -> str:
    """The currently checked-out branch name (empty string if detached/unknown)."""
    try:
        result = run_git(["rev-parse", "--abbrev-ref", "HEAD"], timeout=5)
    except (OSError, subprocess.SubprocessError):
        return ""
    name = result.stdout.strip()
    return "" if name in ("", "HEAD") else name


def normalize_branch_name(raw: str, *, remote: str = DEFAULT_REMOTE) -> str | None:
    """Turn one line of ``git branch`` / ``git branch -r`` into a plain name.

    Returns ``None`` for lines that are not real branches (detached HEAD
    markers, the ``origin/HEAD -> origin/main`` pointer, blank lines).
    """
    name = raw.strip()
    if not name:
        return None
    if name.startswith(("* ", "+ ")):  # current branch / worktree markers
        name = name[2:].strip()
    if name.startswith("("):  # e.g. "(HEAD detached at abc1234)"
        return None
    if "->" in name:  # e.g. "origin/HEAD -> origin/main"
        return None
    if name.startswith("remotes/"):
        name = name[len("remotes/"):]
    prefix = f"{remote}/"
    if name.startswith(prefix):
        name = name[len(prefix):]
    return name or None


def parse_branches(*outputs: str, remote: str = DEFAULT_REMOTE) -> list[str]:
    """Merge git branch listings into a sorted, de-duplicated list of names."""
    seen: set[str] = set()
    for block in outputs:
        for line in block.splitlines():
            name = normalize_branch_name(line, remote=remote)
            if name:
                seen.add(name)
    return sorted(seen)


def list_branches(*, remote: str = DEFAULT_REMOTE) -> list[str]:
    """List local and remote-tracking branches known to the repo (no network)."""
    local = run_git(["branch"], timeout=10)
    remote_out = run_git(["branch", "-r"], timeout=10)
    return parse_branches(local.stdout, remote_out.stdout, remote=remote)


def head_rev() -> str:
    """The current HEAD commit hash, or empty string if it cannot be read."""
    try:
        result = run_git(["rev-parse", "HEAD"], timeout=5)
    except (OSError, subprocess.SubprocessError):
        return ""
    return result.stdout.strip() if result.returncode == 0 else ""


def changed_files(old_rev: str, new_rev: str) -> list[str]:
    """Repo-relative paths that differ between two commits (empty on any problem)."""
    if not old_rev or not new_rev or old_rev == new_rev:
        return []
    try:
        result = run_git(["diff", "--name-only", old_rev, new_rev], timeout=15)
    except (OSError, subprocess.SubprocessError):
        return []
    if result.returncode != 0:
        return []
    return [line.strip() for line in result.stdout.splitlines() if line.strip()]


def launcher_needs_restart(paths: list[str]) -> bool:
    """True when any changed path is one of the launcher's own program files."""
    return any(Path(p).name in LAUNCHER_FILES for p in paths)


def build_pull_steps(
    branch: str,
    current: str,
    *,
    switch: bool = True,
    remote: str = DEFAULT_REMOTE,
) -> list[list[str]]:
    """The ordered git commands that update the work tree from ``branch``.

    Always fetches first. When ``switch`` is set and the chosen branch differs
    from the current one, it checks the branch out before pulling so the user
    ends up *on* the branch they picked rather than merging it sideways.
    """
    branch = branch.strip()
    if not branch:
        raise ValueError("Choose a branch to pull.")
    steps: list[list[str]] = [["fetch", remote, branch]]
    if switch and branch != current:
        steps.append(["checkout", branch])
    steps.append(["pull", remote, branch])
    return steps


def stream_git(args: list[str], on_line: Callable[[str], None]) -> int:
    """Run one git command, sending each output line to ``on_line``.

    stdout and stderr are merged so progress and errors appear in order.
    stdin is closed so git fails fast instead of hanging on a credential
    prompt that has nowhere to show. Returns the process exit code.
    """
    on_line("$ git " + " ".join(args) + "\n")
    try:
        process = subprocess.Popen(
            ["git", *args],
            cwd=str(ROOT),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.DEVNULL,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            **_hidden_console_kwargs(),
        )
    except OSError as exc:
        on_line(f"[git] Could not start git: {exc}\n")
        return -1
    assert process.stdout is not None
    for line in process.stdout:
        on_line(line)
    return process.wait()


def _git_text(
    args: list[str],
    *,
    input_text: str | None = None,
    timeout: float | None = None,
    env_extra: dict[str, str] | None = None,
) -> subprocess.CompletedProcess:
    """Run git capturing output, never blocking on a credential prompt."""
    env = os.environ.copy()
    env["GIT_TERMINAL_PROMPT"] = "0"  # fail fast instead of hanging the GUI on auth
    if env_extra:
        env.update(env_extra)
    return subprocess.run(
        ["git", *args],
        cwd=str(ROOT),
        input=input_text,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        env=env,
        **_hidden_console_kwargs(),
    )


def publish_report(
    content: str,
    *,
    branch: str,
    rel_path: str = "diagnostics/launcher_report.txt",
    remote: str = DEFAULT_REMOTE,
    message: str | None = None,
    timeout: float = 90,
) -> tuple[bool, str]:
    """Commit ``content`` as ``rel_path`` onto ``branch`` and push it.

    Uses git plumbing so the working tree and current checkout are never
    touched: it writes a blob, builds a tree on top of the branch's current tip
    (preserving the branch's other files), commits it, and pushes the new commit
    straight to ``refs/heads/<branch>`` on ``remote``. Returns ``(ok, detail)``.
    """
    message = message or f"Launcher debug report {datetime.now():%Y-%m-%d %H:%M:%S}"

    # Latest tip of the branch, if it already exists on the remote.
    parent = ""
    fetch = _git_text(["fetch", remote, branch], timeout=timeout)
    if fetch.returncode == 0:
        rev = _git_text(["rev-parse", "FETCH_HEAD"], timeout=15)
        if rev.returncode == 0:
            parent = rev.stdout.strip()

    blob_res = _git_text(["hash-object", "-w", "--stdin"], input_text=content, timeout=15)
    if blob_res.returncode != 0:
        return False, f"hash-object failed: {blob_res.stderr.strip()}"
    blob = blob_res.stdout.strip()

    with tempfile.TemporaryDirectory() as td:
        env_idx = {"GIT_INDEX_FILE": str(Path(td) / "index")}
        if parent:
            read = _git_text(["read-tree", f"{parent}^{{tree}}"], env_extra=env_idx, timeout=15)
            if read.returncode != 0:
                return False, f"read-tree failed: {read.stderr.strip()}"
        upd = _git_text(
            ["update-index", "--add", "--cacheinfo", f"100644,{blob},{rel_path}"],
            env_extra=env_idx,
            timeout=15,
        )
        if upd.returncode != 0:
            return False, f"update-index failed: {upd.stderr.strip()}"
        wtree = _git_text(["write-tree"], env_extra=env_idx, timeout=15)
        if wtree.returncode != 0:
            return False, f"write-tree failed: {wtree.stderr.strip()}"
        tree = wtree.stdout.strip()

    commit_args = ["commit-tree", tree, "-m", message]
    if parent:
        commit_args[2:2] = ["-p", parent]
    # A fallback identity so commit-tree never fails on an unconfigured git.
    identity = {
        "GIT_AUTHOR_NAME": "Daily Queue Launcher",
        "GIT_AUTHOR_EMAIL": "launcher@local",
        "GIT_COMMITTER_NAME": "Daily Queue Launcher",
        "GIT_COMMITTER_EMAIL": "launcher@local",
    }
    commit_res = _git_text(commit_args, env_extra=identity, timeout=15)
    if commit_res.returncode != 0:
        return False, f"commit-tree failed: {commit_res.stderr.strip()}"
    commit = commit_res.stdout.strip()

    push = _git_text(["push", remote, f"{commit}:refs/heads/{branch}"], timeout=timeout)
    if push.returncode != 0:
        return False, f"push failed: {push.stderr.strip() or push.stdout.strip()}"
    return True, f"pushed {commit[:10]} to {remote}/{branch}"


def run_pull_steps(steps: list[list[str]], on_line: Callable[[str], None]) -> int:
    """Run each git step in order, stopping at the first failure.

    Returns the exit code of the last command run (0 means every step
    succeeded).
    """
    code = 0
    for step in steps:
        code = stream_git(step, on_line)
        if code != 0:
            on_line(f"[git] Step failed with exit code {code}; stopping.\n")
            break
    return code
