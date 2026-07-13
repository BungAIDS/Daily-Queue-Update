"""Tests for the all-time master log (live_master.py).

    python test_live_master.py

Covers the upsert: first-seen 'added' is set once and stable, returning orders
clear 'left' and flip on_queue back on, departing orders get stamped 'left' and
on_queue=False, and ordered() is chronological by added.
"""
from __future__ import annotations

import sys
from datetime import datetime

import live_master as lm

T0 = datetime(2026, 6, 16, 9, 0, 0)
T1 = datetime(2026, 6, 16, 9, 2, 0)
T2 = datetime(2026, 6, 16, 9, 4, 0)


def _job(num, **kw):
    j = {"job": num, "customer": "ACME", "end_date": "06/20/2026"}
    j.update(kw)
    return j


def test_append_and_added_stable():
    m = {"orders": {}}
    lm.update(m, [_job("100", _first_seen="2026-06-16T09:00:00")], T0)
    assert m["orders"]["100"]["added"] == "2026-06-16T09:00:00"
    assert m["orders"]["100"]["on_queue"] is True and m["orders"]["100"]["left"] is None
    # A later poll with changed data keeps 'added' fixed.
    lm.update(m, [_job("100", end_date="06/25/2026")], T1)
    assert m["orders"]["100"]["added"] == "2026-06-16T09:00:00"
    assert m["orders"]["100"]["job"]["end_date"] == "06/25/2026"


def test_leave_then_return():
    m = {"orders": {}}
    lm.update(m, [_job("100"), _job("200")], T0)
    lm.update(m, [_job("100")], T1)                  # 200 drops off
    assert m["orders"]["200"]["on_queue"] is False
    assert m["orders"]["200"]["left"] == T1.isoformat(timespec="seconds")
    lm.update(m, [_job("100"), _job("200")], T2)     # 200 returns
    assert m["orders"]["200"]["on_queue"] is True
    assert m["orders"]["200"]["left"] is None


def test_history_tracks_every_in_and_out():
    m = {"orders": {}}
    lm.update(m, [_job("100", _first_seen=T0.isoformat(timespec="seconds"))], T0)  # in
    lm.update(m, [], T1)                                                            # out
    lm.update(m, [_job("100")], T2)                                                 # in (return)
    hist = m["orders"]["100"]["history"]
    assert hist == [
        {"event": "in", "time": T0.isoformat(timespec="seconds")},
        {"event": "out", "time": T1.isoformat(timespec="seconds")},
        {"event": "in", "time": T2.isoformat(timespec="seconds")},
    ]


def test_history_seeds_legacy_entries_from_added_and_last_out():
    # An entry created before history tracking is seeded from added (+ prior
    # last_out), then new transitions are appended continuously.
    m = {"orders": {"600": {"added": "2026-06-10T09:00:00", "last_in": "2026-06-12T09:00:00",
                            "last_out": "2026-06-11T17:00:00", "on_queue": False,
                            "seen_on_queue": True, "left": "2026-06-11T17:00:00",
                            "job": {"job": "600"}}}}
    lm.update(m, [_job("600")], T0)            # legacy order returns
    assert m["orders"]["600"]["history"] == [
        {"event": "in", "time": "2026-06-10T09:00:00"},
        {"event": "out", "time": "2026-06-11T17:00:00"},
        {"event": "in", "time": T0.isoformat(timespec="seconds")},
    ]


def test_last_in_updates_on_return_added_stays():
    m = {"orders": {}}
    lm.update(m, [_job("100", _first_seen="2026-06-16T09:00:00")], T0)
    e = m["orders"]["100"]
    assert e["last_in"] == T0.isoformat(timespec="seconds") and e["last_out"] is None
    # Drops off -> last_out stamped, last_in unchanged.
    lm.update(m, [], T1)
    assert e["on_queue"] is False and e["last_out"] == T1.isoformat(timespec="seconds")
    assert e["last_in"] == T0.isoformat(timespec="seconds")
    # Returns -> last_in moves to the re-entry; all-time 'added' stays the first
    # sight, and on_queue() shows last_in as what the Live Queue 'Added' uses.
    lm.update(m, [_job("100")], T2)
    assert e["last_in"] == T2.isoformat(timespec="seconds")
    assert e["added"] == "2026-06-16T09:00:00"
    assert lm.on_queue(m)[0]["_added_iso"] == T2.isoformat(timespec="seconds")


def test_last_in_stable_while_continuously_present():
    # A poll (e.g. after a watch.py restart) on an order that never left must not
    # move its last_in — the add time it shows survives restarts.
    m = {"orders": {}}
    lm.update(m, [_job("100", _first_seen="2026-06-16T09:00:00")], T0)
    lm.update(m, [_job("100", end_date="06/30/2026")], T1)
    assert m["orders"]["100"]["last_in"] == T0.isoformat(timespec="seconds")


def test_on_queue_migrates_missing_last_in_from_added():
    # An entry persisted before last_in existed: on_queue falls back to 'added',
    # and the next present poll backfills last_in from 'added' (not 'now').
    m = {"orders": {"100": {"added": "2026-06-16T09:00:00", "on_queue": True,
                            "job": {"job": "100"}}}}
    assert lm.on_queue(m)[0]["_added_iso"] == "2026-06-16T09:00:00"
    lm.update(m, [_job("100")], T2)
    assert m["orders"]["100"]["last_in"] == "2026-06-16T09:00:00"


def test_ordered_is_chronological_and_on_queue_filter():
    m = {"orders": {}}
    lm.update(m, [_job("100", _first_seen="2026-06-16T09:00:00")], T0)
    lm.update(m, [_job("100"), _job("300", _first_seen="2026-06-16T09:02:00")], T1)
    lm.update(m, [_job("300")], T2)                  # 100 leaves
    keys = [k for k, _ in lm.ordered(m)]
    assert keys == ["100", "300"]                    # oldest-added first
    assert [j["job"] for j in lm.on_queue(m)] == ["300"]


def test_update_logs_field_modifications():
    m = {"orders": {}}
    lm.update(m, [_job("100", end_date="06/20/2026", total_price="$1,000.00")], T0)
    # No events on first sight (it's new, not modified).
    ev = lm.update(m, [_job("100", end_date="06/25/2026", total_price="$1,000.00")], T1)
    assert len(ev) == 1
    e = ev[0]
    assert e["job"] == "100" and e["field"] == "End Date"
    assert e["old"] == "06/20/2026" and e["new"] == "06/25/2026"
    assert e["time"] == T1.isoformat(timespec="seconds")


def test_update_tracks_co_and_skips_initial_population():
    m = {"orders": {}}
    # First sight has no SO size yet; later it's enriched (''-> value): NOT a change.
    lm.update(m, [_job("100", co_number=0)], T0)
    ev = lm.update(m, [_job("100", co_number=0, so_size="M2")], T1)
    assert not any(x["field"] == "Size" for x in ev)        # initial population skipped
    # A real CO# bump (0 -> 1) is logged.
    ev2 = lm.update(m, [_job("100", co_number=1, so_size="M2")], T2)
    co = [x for x in ev2 if x["field"] == "CO#"]
    assert co and co[0]["old"] == "0" and co[0]["new"] == "1"
    # And a real Size modification (M2 -> M3) is logged.
    ev3 = lm.update(m, [_job("100", co_number=1, so_size="M3")], T2)
    assert any(x["field"] == "Size" and x["old"] == "M2" and x["new"] == "M3" for x in ev3)


def test_sparse_source_keeps_unknown_fields_and_logs_no_blanks():
    # A job dict with NO KEY for a field (a board-only seed from a raw morning
    # snapshot folding over a backfill-enriched entry) must neither wipe the
    # field nor log a '-> blank' change. Stripping here used to flip-flop with
    # _merge_external_before_save reviving the fields from disk on every save,
    # re-logging the same phantom changes poll after poll.
    m = {"orders": {}}
    lm.update(m, [_job("100", co_number=2, so_pdf="Z:/SO/100/CO#2.pdf",
                       so_size="245", so_wheel_type="RTF",
                       line_items=[{"tags": ["stainless"]}])], T0)
    ev = lm.update(m, [_job("100")], T1)      # board fields only, nothing changed
    assert ev == []
    job = m["orders"]["100"]["job"]
    assert job["so_size"] == "245" and job["so_wheel_type"] == "RTF"
    assert job["line_items"] == [{"tags": ["stainless"]}]
    # A genuine board change still lands and is still the only event.
    ev2 = lm.update(m, [_job("100", end_date="06/25/2026")], T2)
    assert [(e["field"], e["old"], e["new"]) for e in ev2] == \
        [("End Date", "06/20/2026", "06/25/2026")]
    assert m["orders"]["100"]["job"]["so_size"] == "245"


def test_present_but_empty_key_still_blanks_and_logs():
    # A source that DOES carry the key with an empty value is a real blanking
    # (a Note cleared on the board) — applied and logged as before.
    m = {"orders": {}}
    lm.update(m, [_job("100", status_note="CREDIT HOLD")], T0)
    ev = lm.update(m, [_job("100", status_note="")], T1)
    assert [(e["field"], e["old"], e["new"]) for e in ev] == \
        [("Note", "CREDIT HOLD", "")]
    assert m["orders"]["100"]["job"]["status_note"] == ""


def test_failed_refetch_does_not_regress_change_order():
    m = {"orders": {}}
    # A real change order with its SO link + spec.
    lm.update(m, [_job("100", co_number=2, so_pdf="Z:/SO/100/CO#2.pdf", so_size="27")], T0)
    assert m["orders"]["100"]["job"]["co_number"] == 2
    # A later poll where a re-fetch FAILED (co dropped to 0, link + spec blanked)
    # must NOT wipe the known change order, and must log NO CO# change.
    ev = lm.update(m, [_job("100", co_number=0, so_pdf="", so_size="")], T1)
    job = m["orders"]["100"]["job"]
    assert job["co_number"] == 2 and job["so_pdf"] == "Z:/SO/100/CO#2.pdf" and job["so_size"] == "27"
    assert not any(e["field"] == "CO#" for e in ev)
    # A genuine advance (CO#2 -> CO#3) is still accepted and logged.
    ev2 = lm.update(m, [_job("100", co_number=3, so_pdf="Z:/SO/100/CO#3.pdf", so_size="27")], T2)
    assert m["orders"]["100"]["job"]["co_number"] == 3
    co = [e for e in ev2 if e["field"] == "CO#"]
    assert co and co[0]["old"] == "2" and co[0]["new"] == "3"


def test_blank_so_link_does_not_overwrite_known_link():
    m = {"orders": {}}
    lm.update(m, [_job("200", co_number=1, so_pdf="Z:/SO/200/CO#1.pdf")], T0)
    # Same CO#, but the link came back blank — keep the one we had.
    lm.update(m, [_job("200", co_number=1, so_pdf="")], T1)
    assert m["orders"]["200"]["job"]["so_pdf"] == "Z:/SO/200/CO#1.pdf"


def test_merge_order_adds_and_never_regresses():
    m = {"orders": {}}
    # A backlog order we've never seen on the board is created off-queue.
    assert lm.merge_order(m, "900", {"so_size": "M2", "dwg_extras": {"51": "x"}}) is True
    e = m["orders"]["900"]
    assert e["on_queue"] is False and e["job"]["so_size"] == "M2"
    assert e["job"]["dwg_extras"] == {"51": "x"}
    # A sparse source must not wipe an existing value with an empty one.
    assert lm.merge_order(m, "900", {"so_size": "", "customer": "ACME"}) is True
    assert e["job"]["so_size"] == "M2" and e["job"]["customer"] == "ACME"
    # No change -> returns False.
    assert lm.merge_order(m, "900", {"so_size": "M2"}) is False
    # Merging onto a live (on-queue) order keeps it on-queue.
    lm.update(m, [_job("900")], T0)
    assert m["orders"]["900"]["on_queue"] is True
    lm.merge_order(m, "900", {"so_arrangement": "9"})
    assert m["orders"]["900"]["on_queue"] is True
    assert m["orders"]["900"]["job"]["so_arrangement"] == "9"


def test_merge_drive_run_clears_stale_binary_parse_at_same_timestamp():
    master = {"orders": {
        "100": {"job": {
            "job": "100",
            "drive_run": {"binary": "garbage"},
            "drive_run_summary": "bad bytes",
            "drive_run_template": "generic_text",
            "drive_run_parsed_at": "2026-07-13T08:00:00",
        }},
    }}
    changed = lm.merge_drive_run(master, "100", {
        "drive_run": {},
        "drive_run_summary": "",
        "drive_run_template": "unknown",
        "drive_run_parsed_at": "2026-07-13T08:00:00",
    })
    job = master["orders"]["100"]["job"]
    assert changed
    assert job["drive_run"] == {}
    assert job["drive_run_summary"] == ""
    assert job["drive_run_template"] == "unknown"


def test_merged_backlog_is_not_a_removal():
    m = {"orders": {}}
    # A merged backlog order was never on the board: off-queue, no left, not seen.
    lm.merge_order(m, "900", {"so_size": "M2"})
    e = m["orders"]["900"]
    assert e["on_queue"] is False and e["left"] is None and not e.get("seen_on_queue")
    # An order the watcher saw, then it leaves: seen_on_queue True, left stamped.
    lm.update(m, [_job("800")], T0)
    assert m["orders"]["800"]["seen_on_queue"] is True
    lm.update(m, [], T1)                      # 800 drops off the board
    assert m["orders"]["800"]["on_queue"] is False and m["orders"]["800"]["left"]
    assert m["orders"]["800"]["seen_on_queue"] is True


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
