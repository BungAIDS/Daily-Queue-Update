#!/usr/bin/env python3
"""Show how one order's line items CURRENTLY parse, from the published corpus.

A review note is a comment on how a row was captured, so you cannot judge it
from the note text alone — you need the real line item: its verbatim `raw`, the
tags/component the parser assigned, and the attributes. This reads that order's
record from the sharded line-items store on the `order-data` branch (main store
first, then the backfill overlay) so you can see exactly what the note is
reacting to before touching line_items.py.

    python .claude/skills/apply-so-review-notes/scripts/show_order.py 422029
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[4]
BRANCH = "order-data"


def _show(path: str) -> str | None:
    r = subprocess.run(["git", "-C", str(REPO), "show", f"origin/{BRANCH}:{path}"],
                       capture_output=True, text=True)
    return r.stdout if r.returncode == 0 else None


def _shard_paths(order: str) -> list[str]:
    m = re.match(r"(\d{3,})", str(order))
    if not m:
        return []
    lo = (int(m.group(1)) // 1000) * 1000
    span = f"jobs-{lo}-{lo + 999}.json"
    return [f"line_items.{span}", f"backfill_line_items.{span}"]


def main() -> int:
    if len(sys.argv) != 2:
        print("usage: show_order.py <order-number>", file=sys.stderr)
        return 2
    order = sys.argv[1].strip()
    subprocess.run(["git", "-C", str(REPO), "fetch", "origin", BRANCH],
                   capture_output=True, text=True)

    job = None
    for path in _shard_paths(order):
        blob = _show(path)
        if not blob:
            continue
        job = (json.loads(blob).get("jobs") or {}).get(order)
        if job:
            print(f"# {order} — from {path}")
            break
    if not job:
        print(f"Order {order} not found in the published corpus shards. It may "
              f"predate the store or not be published yet.", file=sys.stderr)
        return 1

    print(f"customer={job.get('customer','')!r}  co={job.get('co_number')}  "
          f"arrangement={job.get('arrangement','')!r}  "
          f"parts_only={job.get('parts_only')}")
    for i, it in enumerate(job.get("items") or [], 1):
        attrs = it.get("attributes") if isinstance(it.get("attributes"), dict) else {}
        comp = attrs.get("component", "")
        extra = {k: v for k, v in attrs.items() if k != "component"}
        print(f"\n[{i}] {it.get('raw','')}")
        print(f"    norm     : {it.get('norm','')}")
        print(f"    tags     : {it.get('tags') or []}")
        print(f"    component: {comp!r}")
        if extra:
            print(f"    attrs    : {json.dumps(extra, ensure_ascii=False)}")
        if it.get("details"):
            print(f"    details  : {it.get('details')}")
        if it.get("review_flags"):
            print(f"    review   : {it.get('review_flags')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
