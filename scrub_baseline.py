"""One-time cleanup: drop a day's start-of-day baseline change events.

    python scrub_baseline.py            # clean today's change log
    python scrub_baseline.py 2026-06-25 # clean a specific day

Before the baseline-poll fix landed (watch.poll_once now skips logging on the
baseline poll), the first poll of the day wrote its whole diff-vs-yesterday into
the change log, so the Changes tab showed a grey "changed today" row under nearly
every order, all stamped at the first poll's time (e.g. 5:00 AM). This removes
exactly those events from a day already logged by the old code, so the tab reads
right today without waiting for tomorrow's clean baseline.

It is surgical and safe: it removes only the events from the single EARLIEST
timestamp in the day (the baseline poll — the first time the board was seen, when
no intraday change is possible yet), and only when that batch is baseline-sized —
a flood across the queue, not the handful a normal poll logs. That makes it
idempotent: once the flood is gone, the new earliest batch is small and a second
run does nothing. Run it with the watcher stopped so a poll can't append while it
rewrites the file.
"""
from __future__ import annotations

import sys
from datetime import date, datetime

import change_log

# A baseline poll diffs the whole board against yesterday, so it floods the log
# with many events at one instant; a normal poll logs only the few fields that
# actually moved. Only strip the earliest batch when it's at least this big, so
# we never mistake a genuine cluster of intraday changes for the baseline.
BASELINE_MIN_BATCH = 8


def scrub(d: date) -> int:
    events = change_log.load(d)
    if not events:
        print(f"No change log for {d.isoformat()} — nothing to do.")
        return 0
    t0 = min(e.get("time", "") for e in events)
    batch = [e for e in events if e.get("time", "") == t0]
    if len(batch) < BASELINE_MIN_BATCH:
        print(f"{d.isoformat()}: earliest batch ({t0}) has only {len(batch)} "
              f"event(s) — not a baseline flood, leaving it. Nothing removed.")
        return 0
    kept = [e for e in events if e.get("time", "") != t0]
    change_log.save(d, kept)
    print(f"{d.isoformat()}: removed {len(batch)} baseline event(s) stamped {t0}; "
          f"kept {len(kept)} genuine intraday change(s).")
    return len(batch)


def main() -> int:
    arg = sys.argv[1] if len(sys.argv) > 1 else ""
    try:
        d = datetime.strptime(arg, "%Y-%m-%d").date() if arg else date.today()
    except ValueError:
        print(f"Bad date {arg!r} — use YYYY-MM-DD.", file=sys.stderr)
        return 2
    scrub(d)
    return 0


if __name__ == "__main__":
    sys.exit(main())
