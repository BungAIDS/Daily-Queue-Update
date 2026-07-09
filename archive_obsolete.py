"""Tuck non-active quote runs and sales orders into an OBSOLETE\\ subfolder.

Each job accumulates duplicates and superseded copies: a run saved as both .txt
and .pdf, an old CO#/REV run beside the current one, older Sales-Order revisions
kept on disk. This moves everything EXCEPT the current one into
`<job folder>\\OBSOLETE\\`, so the live folder holds only what's active. It runs
against:
  - DRIVE_RUN_DIR\\<job>\\   (the quote-run archive the report's YES(X) opens)
  - SALES_ORDER_DIR\\<job>\\ (Sales Order PDFs, older revisions kept)

"Active" = the most-current file by run_rank (CO# > REV > file mtime). Ties and
single-file folders are left alone.

SAFE BY DESIGN — it touches your real Z: files, so:
  - DRY RUN by default: prints what WOULD move, moves nothing. Add --apply to
    actually move.
  - It MOVES, never deletes (into OBSOLETE\\; a name clash gets a numbered
    suffix, never an overwrite).
  - Every move is written to a manifest under BACKLOG_DIR, and
    `--undo <manifest.json>` puts everything back.
  - Best-effort per folder: one problem folder never sinks the rest.

Usage:
    python archive_obsolete.py                 # DRY RUN over runs + sales orders
    python archive_obsolete.py --apply         # actually move
    python archive_obsolete.py --runs          # only the quote-run archive
    python archive_obsolete.py --sales         # only sales orders
    python archive_obsolete.py 421311 421457   # only these jobs
    python archive_obsolete.py --undo C:\\...\\obsolete_manifest_2026....json
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from config import BACKLOG_DIR, DRIVE_RUN_DIR, SALES_ORDER_DIR
from run_rank import rank_paths

log = logging.getLogger("archive-obsolete")

OBSOLETE = "OBSOLETE"


def _job_files(job_dir: Path) -> List[Path]:
    """Real files directly in a job folder (not the OBSOLETE subfolder, not
    Office lock/temp files)."""
    try:
        return sorted(f for f in job_dir.iterdir()
                      if f.is_file() and not f.name.startswith("~$"))
    except OSError:
        return []


def plan_folder(job_dir: Path) -> Tuple[Optional[Path], List[Path]]:
    """(active file, [obsolete files]) for one job folder. Active = the most
    current by run_rank; everything else is obsolete. (None, []) when there's
    nothing to move (0 or 1 file)."""
    files = _job_files(job_dir)
    if len(files) <= 1:
        return (files[0] if files else None), []
    ranked = rank_paths(files)
    return ranked[0], ranked[1:]


def _unique_dest(dest_dir: Path, name: str) -> Path:
    """A destination path in dest_dir that doesn't clobber an existing file."""
    dest = dest_dir / name
    if not dest.exists():
        return dest
    stem, suf = Path(name).stem, Path(name).suffix
    i = 1
    while (dest_dir / f"{stem} ({i}){suf}").exists():
        i += 1
    return dest_dir / f"{stem} ({i}){suf}"


def archive_folder(job_dir: Path, apply: bool) -> List[Dict[str, str]]:
    """Move (or, in dry run, just report) the obsolete files of one job folder
    into its OBSOLETE\\ subfolder. Returns [{src, dst}] move records."""
    active, obsolete = plan_folder(job_dir)
    if not obsolete:
        return []
    moves: List[Dict[str, str]] = []
    dest_dir = job_dir / OBSOLETE
    for src in obsolete:
        dst = _unique_dest(dest_dir, src.name)
        moves.append({"src": str(src), "dst": str(dst)})
        if apply:
            try:
                dest_dir.mkdir(exist_ok=True)
                src.rename(dst)                       # move, never delete
            except OSError as e:
                log.warning("  could not move %s (%s)", src.name, e)
                moves.pop()
    verb = "moved" if apply else "would move"
    log.info("%s  active=%s  %s %d -> OBSOLETE\\", job_dir.name,
             active.name if active else "-", verb, len(moves))
    return moves


def _iter_job_dirs(roots: List[Path], jobs: Optional[List[str]]) -> List[Path]:
    wanted = {str(j).strip() for j in jobs} if jobs else None
    out: List[Path] = []
    for root in roots:
        if not root.exists():
            log.warning("root not reachable: %s (is Z: mapped?)", root)
            continue
        for d in sorted(root.iterdir()):
            if d.is_dir() and d.name != OBSOLETE and (wanted is None or d.name in wanted):
                out.append(d)
    return out


def undo(manifest_path: Path) -> int:
    """Reverse a prior --apply: move every file back from OBSOLETE to where it
    was. Files that already moved on are skipped, not clobbered."""
    try:
        moves = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as e:
        log.error("could not read manifest %s (%s)", manifest_path, e)
        return 1
    restored = 0
    for mv in reversed(moves):                        # undo in reverse order
        dst, src = Path(mv["dst"]), Path(mv["src"])
        if dst.exists() and not src.exists():
            try:
                dst.rename(src)
                restored += 1
            except OSError as e:
                log.warning("  could not restore %s (%s)", dst, e)
    log.info("Restored %d file(s) from the manifest.", restored)
    return 0


def _parse_args(argv: List[str]) -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Move non-active quote runs / sales orders into OBSOLETE\\ subfolders.")
    ap.add_argument("jobs", nargs="*", help="Only these job numbers (default: all).")
    ap.add_argument("--apply", action="store_true",
                    help="Actually move files (default is a dry run that moves nothing).")
    ap.add_argument("--runs", action="store_true", help="Only the quote-run archive.")
    ap.add_argument("--sales", action="store_true", help="Only sales-order folders.")
    ap.add_argument("--undo", metavar="MANIFEST", help="Reverse a prior --apply from its manifest.")
    return ap.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    args = _parse_args(sys.argv[1:] if argv is None else argv)
    if args.undo:
        return undo(Path(args.undo))

    roots: List[Path] = []
    if args.runs or not args.sales:
        roots.append(DRIVE_RUN_DIR)
    if args.sales or not args.runs:
        roots.append(SALES_ORDER_DIR)

    job_dirs = _iter_job_dirs(roots, args.jobs)
    if not job_dirs:
        log.info("No job folders found under %s.", " / ".join(str(r) for r in roots))
        return 0

    if not args.apply:
        log.info("DRY RUN — nothing will move. Re-run with --apply once the list looks right.")
    all_moves: List[Dict[str, str]] = []
    for d in job_dirs:
        all_moves.extend(archive_folder(d, apply=args.apply))

    if args.apply and all_moves:
        BACKLOG_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        manifest = BACKLOG_DIR / f"obsolete_manifest_{stamp}.json"
        manifest.write_text(json.dumps(all_moves, indent=2), encoding="utf-8")
        log.info("Moved %d file(s). Undo with:  python archive_obsolete.py --undo %s",
                 len(all_moves), manifest)
    elif not args.apply:
        log.info("Would move %d file(s) across %d folder(s). Add --apply to do it.",
                 len(all_moves), len(job_dirs))
    else:
        log.info("Nothing to move — every folder already holds only its active file.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
