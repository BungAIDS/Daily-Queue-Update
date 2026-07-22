"""Publish order data as a replaceable, orphaned Git snapshot.

Large job-keyed JSON stores are published as deterministic job-range shards
with a manifest.  Small files and JSON that is not a job store retain their
existing names and contents for backwards compatibility.
"""
from __future__ import annotations

import json
import logging
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, List, Optional, Sequence

from config import BACKLOG_DIR, SNAPSHOT_DIR, LINE_ITEMS_STORE, DATA_PUSH_BRANCH

log = logging.getLogger(__name__)

_REPO = Path(__file__).resolve().parent
_IDENTITY = {
    "GIT_AUTHOR_NAME": "queue-data-publisher", "GIT_AUTHOR_EMAIL": "queue-data@local",
    "GIT_COMMITTER_NAME": "queue-data-publisher", "GIT_COMMITTER_EMAIL": "queue-data@local",
}

# GitHub starts warning at 50 MiB and rejects a blob at 100 MiB.  Forty MiB
# leaves room for growth between scans without flirting with either boundary.
MAX_JSON_BYTES = int(os.environ.get("DATA_PUSH_MAX_JSON_BYTES", 40 * 1024 * 1024))
JOB_RANGE_SIZE = max(1, int(os.environ.get("DATA_PUSH_JOB_RANGE", "1000")))
_JOB_KEY = re.compile(r"^(?P<number>\d{3,})(?P<suffix>[A-Za-z0-9-]*)$")


@dataclass(frozen=True)
class _PublishedFile:
    name: str
    path: Optional[Path] = None
    data: Optional[bytes] = None


def data_files() -> List[Path]:
    """Return existing order-data files in stable publication order."""
    line_items = LINE_ITEMS_STORE if LINE_ITEMS_STORE else BACKLOG_DIR / "line_items.json"
    today = date.today()
    candidates = [
        SNAPSHOT_DIR / "live_master.json",
        # Today's + yesterday's field-change logs: the Changes tab's source, so
        # a page/report built from this branch shows the day's activity too.
        SNAPSHOT_DIR / f"change_log_{today.isoformat()}.json",
        SNAPSHOT_DIR / f"change_log_{(today - timedelta(days=1)).isoformat()}.json",
        BACKLOG_DIR / "quote_run_scan_progress.json",
        line_items,
        BACKLOG_DIR / "backfill_line_items.json",
        BACKLOG_DIR / "backfill_progress.json",
        BACKLOG_DIR / "autocad_scan_progress.json",
        BACKLOG_DIR / "solidworks_scan.json",        # which jobs have 3D data
        BACKLOG_DIR / "so_review_notes.json",
        BACKLOG_DIR / "so_review_parser_metrics.json",
        BACKLOG_DIR / "so_corpus_health.json",
        BACKLOG_DIR / "quote_run_review_notes.json",
        BACKLOG_DIR / "quote_runs.xlsx",
        BACKLOG_DIR / "backlog.xlsx",
        BACKLOG_DIR / "line_items.xlsx",
        BACKLOG_DIR / "autocad_dwgs.xlsx",
    ]
    cleanup_audit = BACKLOG_DIR / "order_verification_cleanup.json"
    try:
        cleanup_mtime = cleanup_audit.stat().st_mtime
    except OSError:
        cleanup_mtime = 0.0

    seen, out = set(), []
    for p in candidates:
        rp = Path(p)
        if rp in seen:
            continue
        seen.add(rp)
        if rp.is_file():
            if rp.name in {"backlog.xlsx", "line_items.xlsx"} and cleanup_mtime:
                try:
                    if rp.stat().st_mtime < cleanup_mtime:
                        continue
                except OSError:
                    continue
            out.append(rp)
    return out


def _git(args: List[str]) -> subprocess.CompletedProcess:
    env = {**os.environ, **_IDENTITY, "GIT_TERMINAL_PROMPT": "0"}
    return subprocess.run(["git", *args], cwd=_REPO, env=env,
                          capture_output=True, text=True, timeout=120)


def _git_stdin(args: List[str], data: bytes) -> subprocess.CompletedProcess:
    env = {**os.environ, **_IDENTITY, "GIT_TERMINAL_PROMPT": "0"}
    return subprocess.run(["git", *args], cwd=_REPO, env=env, input=data,
                          capture_output=True, timeout=120)


def _job_records(payload: Any) -> tuple[Optional[str], Optional[dict[str, Any]]]:
    """Find a job-keyed mapping, returning its optional wrapper key."""
    if not isinstance(payload, dict):
        return None, None
    candidates: list[tuple[Optional[str], Any]] = [(None, payload)]
    candidates.extend((key, payload.get(key)) for key in ("orders", "records", "items", "jobs"))
    for wrapper, value in candidates:
        if isinstance(value, dict) and value and all(_JOB_KEY.fullmatch(str(k)) for k in value):
            return wrapper, value
    return None, None


def _job_sort_key(key: str) -> tuple[int, str]:
    match = _JOB_KEY.fullmatch(str(key))
    if not match:
        raise ValueError(f"Not a job key: {key!r}")
    return int(match.group("number")), str(key)


def _json_bytes(value: Any) -> bytes:
    return (json.dumps(value, indent=2, ensure_ascii=False) + "\n").encode("utf-8")


def _split_oversized_group(entries: list[tuple[str, Any]],
                           wrapper: Optional[str]) -> list[list[tuple[str, Any]]]:
    """Split an unusually dense numeric range until every shard fits."""
    payload = {wrapper: dict(entries)} if wrapper else dict(entries)
    if len(_json_bytes(payload)) <= MAX_JSON_BYTES or len(entries) <= 1:
        return [entries]
    midpoint = len(entries) // 2
    return (_split_oversized_group(entries[:midpoint], wrapper)
            + _split_oversized_group(entries[midpoint:], wrapper))


def _shard_json(path: Path) -> Optional[List[_PublishedFile]]:
    """Shard an oversized job store, or return None when it is not shardable."""
    try:
        if path.stat().st_size <= MAX_JSON_BYTES or path.suffix.lower() != ".json":
            return None
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return None

    wrapper, records = _job_records(payload)
    if records is None:
        log.warning("Oversized JSON is not a job-keyed store; publishing it unchanged: %s", path)
        return None

    ordered = [(str(key), value) for key, value in
               sorted(records.items(), key=lambda item: _job_sort_key(str(item[0])))]
    ranges: dict[int, list[tuple[str, Any]]] = {}
    for key, value in ordered:
        number, _ = _job_sort_key(key)
        range_start = (number // JOB_RANGE_SIZE) * JOB_RANGE_SIZE
        ranges.setdefault(range_start, []).append((key, value))

    chunks: list[tuple[int, list[tuple[str, Any]]]] = []
    for range_start, entries in sorted(ranges.items()):
        for split in _split_oversized_group(entries, wrapper):
            chunks.append((range_start, split))

    shard_files: list[_PublishedFile] = []
    shard_names: list[dict[str, Any]] = []
    stem = path.stem
    parts_per_range: dict[int, int] = {}
    for range_start, _chunk in chunks:
        parts_per_range[range_start] = parts_per_range.get(range_start, 0) + 1
    part_index: dict[int, int] = {}

    for range_start, chunk in chunks:
        range_end = range_start + JOB_RANGE_SIZE - 1
        part_index[range_start] = part_index.get(range_start, 0) + 1
        part_suffix = (f".part-{part_index[range_start]:03d}"
                       if parts_per_range[range_start] > 1 else "")
        name = f"{stem}.jobs-{range_start}-{range_end}{part_suffix}.json"
        shard_payload = {wrapper: dict(chunk)} if wrapper else dict(chunk)
        data = _json_bytes(shard_payload)
        shard_files.append(_PublishedFile(name=name, data=data))
        first_job, _ = _job_sort_key(chunk[0][0])
        last_job, _ = _job_sort_key(chunk[-1][0])
        shard_names.append({"file": name, "range_start": range_start,
                            "range_end": range_end, "first_job": first_job,
                            "last_job": last_job, "jobs": len(chunk),
                            "bytes": len(data)})

    metadata = ({key: value for key, value in payload.items() if key != wrapper}
                if wrapper else {})
    manifest = {
        "format": "order-data-shards-v1",
        "source": path.name,
        "wrapper": wrapper,
        "metadata": metadata,
        "max_shard_bytes": MAX_JSON_BYTES,
        "job_range_size": JOB_RANGE_SIZE,
        "job_count": len(ordered),
        "shards": shard_names,
    }
    manifest_name = f"{stem}.manifest.json"
    shard_files.append(_PublishedFile(name=manifest_name, data=_json_bytes(manifest)))
    return shard_files


def prepare_publish_files(files: Sequence[Path]) -> List[_PublishedFile]:
    """Expand only oversized job stores; preserve all other files as-is."""
    out: list[_PublishedFile] = []
    for path in files:
        shards = _shard_json(Path(path))
        if shards:
            out.extend(shards)
        else:
            out.append(_PublishedFile(name=Path(path).name, path=Path(path)))
    return out


def _hash_published_file(file: _PublishedFile) -> Optional[str]:
    if file.data is not None:
        result = _git_stdin(["hash-object", "-w", "--stdin"], file.data)
    else:
        result = _git(["hash-object", "-w", str(file.path)])
    if result.returncode != 0:
        detail = result.stderr.decode(errors="replace") if isinstance(result.stderr, bytes) else result.stderr
        log.debug("data push: hash-object failed for %s (%s)", file.name, detail.strip())
        return None
    stdout = result.stdout.decode(errors="replace") if isinstance(result.stdout, bytes) else result.stdout
    return stdout.strip()


def _build_snapshot_commit_prepared(files: Sequence[_PublishedFile],
                                    message: str) -> Optional[str]:
    if not files:
        return None
    try:
        entries = []
        for file in files:
            sha = _hash_published_file(file)
            if not sha:
                return None
            name = file.name.replace("\r", "").replace("\n", "")
            entries.append(f"100644 blob {sha}\t{name}")
        result = _git_stdin(["mktree"], ("\n".join(entries) + "\n").encode("utf-8"))
        if result.returncode != 0:
            log.debug("data push: mktree failed (%s)", result.stderr.decode(errors="replace").strip())
            return None
        tree = result.stdout.decode().strip()
        result = _git(["commit-tree", tree, "-m", message])
        if result.returncode != 0:
            log.debug("data push: commit-tree failed (%s)", result.stderr.strip())
            return None
        return result.stdout.strip()
    except (subprocess.SubprocessError, OSError) as e:  # noqa: BLE001
        log.debug("data push: snapshot build skipped (%s)", e)
        return None


def build_snapshot_commit(files: Sequence[Path], message: str) -> Optional[str]:
    """Create one orphan commit containing files, without touching the worktree."""
    return _build_snapshot_commit_prepared(prepare_publish_files(files), message)


def _failure_detail(result: subprocess.CompletedProcess) -> str:
    parts = []
    for value in (result.stderr, result.stdout):
        if value:
            parts.append(value.decode(errors="replace") if isinstance(value, bytes) else value)
    detail = "\n".join(parts).strip()
    return detail[-4000:] if len(detail) > 4000 else detail


def push_data(branch: Optional[str] = None) -> bool:
    """Force-push the current data as one orphan snapshot and verify its SHA."""
    branch = (branch if branch is not None else DATA_PUSH_BRANCH) or ""
    if not branch:
        log.info("DATA_PUSH_BRANCH is empty — order-data publishing is disabled.")
        return False
    files = data_files()
    if not files:
        log.warning("No order-data files found to publish (looked under %s / %s).",
                    SNAPSHOT_DIR, BACKLOG_DIR)
        return False
    published = prepare_publish_files(files)
    msg = (f"order data @ {datetime.now().isoformat(timespec='seconds')} "
           f"({len(published)} published files from {len(files)} sources)")
    commit = _build_snapshot_commit_prepared(published, msg)
    if not commit:
        log.warning("Could not build the order-data snapshot commit.")
        return False
    result = _git(["push", "--force", "origin", f"{commit}:refs/heads/{branch}"])
    if result.returncode != 0:
        log.warning("Order-data push to '%s' failed:\n%s", branch, _failure_detail(result))
        return False

    remote = _git(["ls-remote", "origin", f"refs/heads/{branch}"])
    if remote.returncode == 0:
        remote_sha = remote.stdout.strip().split()[0] if remote.stdout.strip() else ""
        if remote_sha != commit:
            log.warning("Order-data push to '%s' reached remote SHA %s, expected %s.",
                        branch, remote_sha or "<none>", commit)
            return False
        log.info("Verified remote order-data branch '%s' at %s.", branch, remote_sha)
    else:
        log.warning("Order-data push succeeded but remote verification failed:\n%s",
                    _failure_detail(remote))
        return False
    log.info("Published %d order-data file(s) to '%s': %s",
             len(published), branch, ", ".join(p.name for p in published))
    return True


def main(argv: Optional[List[str]] = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    args = sys.argv[1:] if argv is None else argv
    branch = args[0].strip() if args and args[0].strip() else None
    return 0 if push_data(branch) else 1


if __name__ == "__main__":
    raise SystemExit(main())
