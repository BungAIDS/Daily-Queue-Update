"""Find a job's AutoCAD folder (and thus its type) by searching the Z: tree.

A job appears under exactly one  <type>\\<intermediate>\\<job>  path, so locating
it gives BOTH the Excel link target and the job type in one shot.

    python find_job_folder.py 421314 421388 421572

Read-only. Prints the type, the full folder path, and how long the search took
(so we know the network traversal is quick enough to run per job).
"""
from __future__ import annotations

import sys
import time

from config import AUTOCAD_JOBS_DIR


def main() -> None:
    if len(sys.argv) < 2:
        raise SystemExit("Usage: python find_job_folder.py <jobnumber> [more...]")
    root = AUTOCAD_JOBS_DIR
    print(f"Searching under: {root}")
    if not root.exists():
        raise SystemExit(f"Not reachable: {root} — check the path / that Z: is mapped.")

    for job in sys.argv[1:]:
        t0 = time.monotonic()
        # <type>/<intermediate>/<job> — exact first, then a prefix match in case
        # the folder name carries a suffix.
        matches = list(root.glob(f"*/*/{job}")) or list(root.glob(f"*/*/{job}*"))
        dt = time.monotonic() - t0
        if not matches:
            print(f"  {job}: NOT FOUND  ({dt:.1f}s)")
            continue
        for m in matches:
            parts = m.relative_to(root).parts
            print(f"  {job}: type={parts[0]!r}  folder={m}  ({dt:.1f}s)")


if __name__ == "__main__":
    main()
