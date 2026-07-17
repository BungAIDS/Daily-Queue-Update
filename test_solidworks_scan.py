"""Tests for solidworks_scan.py — the per-job SolidWorks 3D folder sweep.

Run:  python test_solidworks_scan.py
Pure filesystem logic on a temp tree; no Z: drive, no network — CI-safe.
"""
from __future__ import annotations

import tempfile
from pathlib import Path

import solidworks_scan as sws


def _tree(td: Path) -> Path:
    """A miniature SolidWorks share:
       root/GENERAL LINE/421/421966/QT RUN.SLDASM      job w/ 3D (uppercase ext)
       root/GENERAL LINE/421/421950/readme.txt          job, no SW files
       root/AXIAL/420001/rev2/wheel.sldprt              nested SW file one level in
       root/169979C/part.slddrw                         flat job at the root
       root/notes/scratch.sldasm                        NOT a job dir -> ignored
    """
    root = td / "SW"
    (root / "GENERAL LINE" / "421" / "421966").mkdir(parents=True)
    (root / "GENERAL LINE" / "421" / "421966" / "QT RUN.SLDASM").write_text("x")
    (root / "GENERAL LINE" / "421" / "421950").mkdir(parents=True)
    (root / "GENERAL LINE" / "421" / "421950" / "readme.txt").write_text("x")
    (root / "AXIAL" / "420001" / "rev2").mkdir(parents=True)
    (root / "AXIAL" / "420001" / "rev2" / "wheel.sldprt").write_text("x")
    (root / "169979C").mkdir()
    (root / "169979C" / "part.slddrw").write_text("x")
    (root / "notes").mkdir()
    (root / "notes" / "scratch.sldasm").write_text("x")
    return root


def test_scan_tree():
    with tempfile.TemporaryDirectory() as td:
        root = _tree(Path(td))
        recs = sws.scan_tree(root)
        assert set(recs) == {"421966", "421950", "420001", "169979C"}, set(recs)
        assert recs["421966"]["has_sw"] and recs["421966"]["exts"] == [".sldasm"]
        assert not recs["421950"]["has_sw"] and recs["421950"]["sw_files"] == 0
        assert recs["420001"]["has_sw"], "nested rev-folder SW file must count"
        assert recs["169979C"]["has_sw"] and recs["169979C"]["exts"] == [".slddrw"]
        assert recs["421966"]["folder"].endswith("421966")
    print("  scan_tree OK")


def test_has_3d_and_find():
    with tempfile.TemporaryDirectory() as td:
        root = _tree(Path(td))
        recs = sws.scan_tree(root)
        assert sws.has_3d(recs, "421966") and sws.has_3d(recs, 169979) is False
        assert not sws.has_3d(recs, "421950")      # scanned, nothing there
        assert not sws.has_3d(recs, "999999")      # never scanned
        assert sws.find_job_folder(root, "169979C") == root / "169979C"
        nested = sws.find_job_folder(root, "421950")
        assert nested is not None and nested.name == "421950"
        assert sws.find_job_folder(root, "999999") is None
    print("  has_3d / find_job_folder OK")


def test_job_dir_pattern():
    ok = ["421966", "169979C", "420001"]
    bad = ["42196", "4219667", "421966CD", "notes", "421-966"]
    for s in ok:
        assert sws._JOB_DIR_RE.match(s), s
    for s in bad:
        assert not sws._JOB_DIR_RE.match(s), s
    print("  job-dir pattern OK")


def main() -> int:
    test_scan_tree()
    test_has_3d_and_find()
    test_job_dir_pattern()
    print("All solidworks_scan tests passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
