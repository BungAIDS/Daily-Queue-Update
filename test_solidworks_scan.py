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


def test_range_folder():
    cases = {420: "416-420", 121: "121-125", 125: "121-125", 123: "121-125",
             401: "400-405", 402: "400-405", 405: "400-405",
             416: "416-420", 421: "421-425", 400: "396-400"}
    for prefix, want in cases.items():
        got = sws.range_folder(prefix)
        assert got == want, f"{prefix}: {got} != {want}"
    print("  range_folder OK")


def test_sw_candidates():
    root = Path(r"Z:\Solidworks\Current\JOBS")
    one = lambda job, t: sws.sw_candidates(root, job, t)[0]
    assert one("420123", "AXIAL") == root / "AXIAL" / "420123"
    assert one("421966", "GENERAL LINE") == root / "GENERAL LINE" / "421-425" / "421966"
    assert one("420500", "HDX") == root / "HDX" / "416-420" / "420500"
    assert one("401234", "HD-PFD") == root / "HD-PFD" / "40XXXX" / "401234"
    # The AutoCAD tree's spelling maps to the SolidWorks one.
    assert one("421234", "HD-PFD-IAF") == root / "HD-PFD" / "42XXXX" / "421234"
    # No type -> one candidate per type; non-numeric job -> none.
    assert len(sws.sw_candidates(root, "421966")) == len(sws._SW_TYPES)
    assert sws.sw_candidates(root, "REPAIR") == []
    print("  sw_candidates OK")


def test_find_derived_first():
    """A job in the real layout is found by derivation (no walk needed)."""
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        target = root / "GENERAL LINE" / "421-425" / "421966"
        target.mkdir(parents=True)
        assert sws.find_job_folder(root, "421966", "GENERAL LINE") == target
        assert sws.find_job_folder(root, "421966") == target   # type unknown
    print("  find derived-first OK")


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
    test_range_folder()
    test_sw_candidates()
    test_find_derived_first()
    test_job_dir_pattern()
    print("All solidworks_scan tests passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
