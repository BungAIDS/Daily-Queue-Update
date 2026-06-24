"""Tests for engineer roster matching / accumulation (engineers.py).

    python test_engineers.py
"""
import sys

import engineers


def _with_roster(roster):
    """Swap in a test roster and recompile; restore is the caller's job via the
    fixture-less try/finally pattern below."""
    engineers.ROSTER = roster
    engineers.reload_roster()


def setup_function(_):
    _with_roster({
        "John Doe": ["John", "JD", "J.D.", "Doe"],
        "Maria Ruiz": ["Maria", "MR", "Ruiz"],
    })


def teardown_function(_):
    _with_roster({})


def test_detect_in_assigned_to():
    assert engineers.detect({"assigned_to": "John"}) == ["John Doe"]


def test_detect_in_checker():
    assert engineers.detect({"checker": "Ruiz"}) == ["Maria Ruiz"]


def test_detect_in_note():
    assert engineers.detect({"status_note": "waiting on Maria for sign-off"}) == ["Maria Ruiz"]


def test_initials_matched_in_note():
    # Project decision: initials are matched everywhere, including the Note.
    assert engineers.detect({"status_note": "checked by JD"}) == ["John Doe"]


def test_dotted_initials():
    assert engineers.detect({"assigned_to": "J.D."}) == ["John Doe"]


def test_multiple_engineers_one_order():
    job = {"assigned_to": "John", "checker": "MR"}
    assert engineers.detect(job) == ["John Doe", "Maria Ruiz"]


def test_no_match_inside_a_word():
    # "JD" must not fire inside "AJDX", nor "Doe" inside "Doendorf".
    assert engineers.detect({"status_note": "AJDX Doendorf"}) == []


def test_case_insensitive():
    assert engineers.detect({"checker": "DOE"}) == ["John Doe"]


def test_empty_when_no_fields():
    assert engineers.detect({"customer": "John's Welding"}) == []  # not a scanned field


def test_empty_roster_returns_empty():
    _with_roster({})
    assert engineers.detect({"assigned_to": "John"}) == []


def test_merge_is_cumulative_union():
    assert engineers.merge(["John Doe"], ["Maria Ruiz"]) == ["John Doe", "Maria Ruiz"]
    assert engineers.merge(["John Doe"], ["John Doe"]) == ["John Doe"]
    assert engineers.merge(None, ["John Doe"]) == ["John Doe"]


def test_cell_text_uses_stored_list():
    assert engineers.cell_text({"engineers": ["John Doe", "Maria Ruiz"]}) == "John Doe, Maria Ruiz"
    assert engineers.cell_text({}) == ""


def test_cell_text_falls_back_to_detection():
    # No stored list (e.g. a daily-report snapshot job) -> detect on the fly.
    assert engineers.cell_text({"assigned_to": "John"}) == "John Doe"


def test_backfill_tags_and_preserves():
    master = {"orders": {
        "100": {"job": {"assigned_to": "John", "checker": "", "status_note": ""}},
        "200": {"job": {"assigned_to": "", "checker": "MR", "status_note": "",
                        "engineers": ["John Doe"]}},  # keep prior + add new (union)
        "300": {"job": {"assigned_to": "nobody"}},
    }}
    changed = engineers.backfill(master)
    assert master["orders"]["100"]["job"]["engineers"] == ["John Doe"]
    assert master["orders"]["200"]["job"]["engineers"] == ["John Doe", "Maria Ruiz"]
    assert master["orders"]["300"]["job"].get("engineers") in (None, [])
    assert changed == 2
    # Idempotent: a second pass changes nothing.
    assert engineers.backfill(master) == 0


def main() -> int:
    # Mirror pytest's per-test setup/teardown so the suite also runs as a plain
    # script (how CI invokes it: `python test_engineers.py`).
    passed = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            setup_function(fn)
            try:
                fn()
            finally:
                teardown_function(fn)
            print(f"  ok  {name}")
            passed += 1
    print(f"\n{passed} tests passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
