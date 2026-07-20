"""Scan the SolidWorks job folders — which jobs have real 3D data?

    python solidworks_scan.py                 # full sweep of SOLIDWORKS_JOBS_DIR
    python solidworks_scan.py --root "Z:\\SW" # sweep a different root
    python solidworks_scan.py --job 421966    # re-check one job
    python solidworks_scan.py --limit 50      # stop after N job folders (testing)

A job "has 3D" when a folder named after it exists under the SolidWorks root
AND that folder holds at least one SolidWorks file — a part (.sldprt),
assembly (.sldasm), or drawing (.slddrw). Results land in one JSON store
(backlog/solidworks_scan.json), read by the GL Queue Explorer's "Has 3D"
filter and published with the other order data.

The share's layout is Z:\\Solidworks\\Current\\JOBS\\<type>\\<intermediate>\\<job>:

    GENERAL LINE   range folder over the first 3 digits ("416-420")
    HD-PFD         first two digits + XXXX ("42XXXX")   [AutoCAD: HD-PFD-IAF]
    HDX            range folder over the first 3 digits ("416-420")
    AXIAL          no intermediate — the job sits right under AXIAL\\

Range folders group the 3-digit prefix in fives: n = ceil(prefix/5), so
prefix 420 -> "416-420"; the one irregular folder is "400-405" (in place of
401-405). `sw_candidates` derives these paths directly, which is how --job
finds a folder without walking; the full sweep stays a NAME-based walk (any
directory named like a job, up to three levels down), so a new type or a
misfiled job still gets found.

Jobs are recorded whether or not 3D files were found — "scanned, nothing
there" and "never scanned" stay distinguishable. Re-running refreshes the
whole store; --job refreshes one entry in place. Import-light (config +
stdlib), so the pure logic is CI-testable (test_solidworks_scan.py).
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

from config import BACKLOG_DIR, SOLIDWORKS_JOBS_DIR

log = logging.getLogger("solidworks-scan")

SW_EXTS = {".sldprt", ".sldasm", ".slddrw"}
PROGRESS_PATH = BACKLOG_DIR / "solidworks_scan.json"
# A job folder is 6 digits with an optional trailing letter (421966, 169979C).
_JOB_DIR_RE = re.compile(r"^\d{6}[A-Za-z]?$")
_WALK_DEPTH = 3          # how deep below the root job folders may sit
_INSIDE_DEPTH = 2        # how deep inside a job folder to look for SW files


def load_progress() -> Dict[str, Dict[str, Any]]:
    if not PROGRESS_PATH.exists():
        return {}
    try:
        return json.loads(PROGRESS_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as e:
        log.warning("Could not read %s (%s); starting fresh", PROGRESS_PATH, e)
        return {}


def save_progress(records: Dict[str, Dict[str, Any]]) -> None:
    PROGRESS_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp = PROGRESS_PATH.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(records, indent=2), encoding="utf-8")
    tmp.replace(PROGRESS_PATH)   # atomic: a crash mid-write never corrupts the store


def has_3d(store: Dict[str, Dict[str, Any]], job: str) -> bool:
    rec = store.get(str(job))
    return bool(rec and rec.get("has_sw"))


# The AutoCAD tree's name for the one type the two shares spell differently.
_SW_TYPE_ALIASES = {"HD-PFD-IAF": "HD-PFD"}
_SW_TYPES = ("GENERAL LINE", "HD-PFD", "HDX", "AXIAL")


def range_folder(prefix: int) -> str:
    """The share's 5-wide range folder for a 3-digit job prefix: 420 ->
    '416-420' (n = ceil(prefix/5)). The share's one irregular folder is
    '400-405', which stands in for the formula's 401-405."""
    import math
    n = math.ceil(prefix / 5)
    start, end = 5 * (n - 1) + 1, 5 * n
    if (start, end) == (401, 405):
        return "400-405"
    return f"{start}-{end}"


def sw_candidates(root: Path, job: str, jtype: str = "") -> "list[Path]":
    """The expected folder location(s) for `job` under the known layout. With
    a job type (the AutoCAD scan's, aliases handled) one path comes back;
    without, one per type — callers try them in order."""
    job = str(job)
    m = re.match(r"^(\d{3})\d{3}", job)
    if not m:
        return []
    prefix = int(m.group(1))
    types = ([_SW_TYPE_ALIASES.get(jtype.upper(), jtype.upper())]
             if jtype else list(_SW_TYPES))
    out = []
    for t in types:
        if t == "AXIAL":
            out.append(root / "AXIAL" / job)
        elif t == "HD-PFD":
            out.append(root / "HD-PFD" / f"{job[:2]}XXXX" / job)
        elif t in ("GENERAL LINE", "HDX"):
            out.append(root / t / range_folder(prefix) / job)
    return out


def job_record(job: str, folder: Path) -> Dict[str, Any]:
    """One job folder's verdict: does it hold any SolidWorks file? Looks a
    couple of levels deep (revision subfolders are common), extensions
    case-insensitive ('QT RUN.SLDASM' counts)."""
    exts: set = set()
    n = 0
    for dirpath, dirnames, filenames in os.walk(folder):
        if len(Path(dirpath).relative_to(folder).parts) >= _INSIDE_DEPTH:
            dirnames[:] = []
        for f in filenames:
            s = Path(f).suffix.lower()
            if s in SW_EXTS:
                exts.add(s)
                n += 1
    return {"job": job, "folder": str(folder), "has_sw": n > 0, "sw_files": n,
            "exts": sorted(exts),
            "scanned_at": datetime.now().isoformat(timespec="seconds")}


def scan_tree(root: Path, limit: int = 0) -> Dict[str, Dict[str, Any]]:
    """One walk of the root: every directory named like a job number (up to
    _WALK_DEPTH levels down) becomes a record. A job appearing twice keeps
    whichever sighting has 3D files."""
    out: Dict[str, Dict[str, Any]] = {}
    for dirpath, dirnames, filenames in os.walk(root):
        p = Path(dirpath)
        depth = len(p.relative_to(root).parts)
        name = p.name
        if depth and _JOB_DIR_RE.match(name):
            rec = job_record(name, p)
            old = out.get(name)
            if old is None or (rec["has_sw"] and not old.get("has_sw")):
                out[name] = rec
            dirnames[:] = []            # job_record already looked inside
            if limit and len(out) >= limit:
                break
            continue
        if depth >= _WALK_DEPTH:
            dirnames[:] = []
    return out


def find_job_folder(root: Path, job: str, jtype: str = "") -> Path | None:
    """Locate one job's folder: the layout-derived candidates first (a couple
    of stats), then the depth-limited walk as the safety net."""
    for cand in sw_candidates(root, job, jtype):
        if cand.is_dir():
            return cand
    direct = root / str(job)
    if direct.is_dir():
        return direct
    for dirpath, dirnames, _ in os.walk(root):
        p = Path(dirpath)
        if len(p.relative_to(root).parts) >= _WALK_DEPTH:
            dirnames[:] = []
        if p.name == str(job):
            return p
    return None


def main(argv: list | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--root", type=Path, default=None,
                    help=f"SolidWorks jobs root (default: {SOLIDWORKS_JOBS_DIR})")
    ap.add_argument("--job", default="",
                    help="Re-check a single job number instead of sweeping.")
    ap.add_argument("--limit", type=int, default=0,
                    help="Stop after this many job folders (0 = no limit).")
    args = ap.parse_args(argv)

    root = args.root or SOLIDWORKS_JOBS_DIR
    if not root.is_dir():
        print(f"SolidWorks root not found: {root}\n"
              "Set SOLIDWORKS_JOBS_DIR in .env to the folder that holds the "
              "per-job SolidWorks directories, then run this again.")
        return 1

    if args.job:
        store = load_progress()
        # The AutoCAD scan already knows the job's type — use it to derive the
        # exact SolidWorks path instead of walking the share.
        jtype = ""
        try:
            import autocad_scan
            jtype = (autocad_scan.load_progress().get(str(args.job)) or {}).get("type", "")
        except Exception:  # noqa: BLE001 - the walk fallback covers us
            pass
        folder = find_job_folder(root, str(args.job), jtype)
        if folder is None:
            store.pop(str(args.job), None)
            save_progress(store)
            print(f"{args.job}: no SolidWorks folder under {root} (entry cleared).")
            return 0
        rec = job_record(str(args.job), folder)
        store[str(args.job)] = rec
        save_progress(store)
        state = "HAS 3D" if rec["has_sw"] else "no SolidWorks files"
        print(f"{args.job}: {state}  ({rec['sw_files']} file(s); {folder})")
        return 0

    records = scan_tree(root, limit=args.limit)
    save_progress(records)
    n_3d = sum(1 for r in records.values() if r.get("has_sw"))
    print(f"Scanned {len(records)} job folder(s) under {root}: "
          f"{n_3d} with SolidWorks files, {len(records) - n_3d} without "
          f"-> {PROGRESS_PATH}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
