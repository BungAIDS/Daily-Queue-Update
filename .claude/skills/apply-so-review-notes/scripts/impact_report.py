#!/usr/bin/env python3
"""Measure what a parser change did to the corpus — the "did this help?" report.

Re-derives every stored line item from its verbatim `raw` text with the CURRENT
rules (exactly what `line_items_scan.py --renorm` will do on the user's machine),
then diffs that against how the corpus is stored today. Reports how many line
items changed, how many fewer MARKED-FOR-REVIEW rows remain, and how much closer
the parser now gets to fully capturing what's on the sales orders.

Run it right after applying a change, before it's deployed, to quantify the win:

    python .claude/skills/apply-so-review-notes/scripts/impact_report.py
    python .claude/skills/apply-so-review-notes/scripts/impact_report.py --store <line_items.json>

With no --store it reads the published corpus (all shards) from the order-data
branch; on the user's machine, point --store at the local line-items store.
"""
from __future__ import annotations

import argparse
import copy
import json
import subprocess
import sys
from collections import Counter
from pathlib import Path

REPO = Path(__file__).resolve().parents[4]
sys.path.insert(0, str(REPO))

import line_items  # noqa: E402
import so_review  # noqa: E402  (parser_review_metrics: the workbook's red-row count)


import re  # noqa: E402


def _shard_span(order: str) -> str | None:
    m = re.match(r"(\d{3,})", str(order))
    if not m:
        return None
    lo = (int(m.group(1)) // 1000) * 1000
    return f"jobs-{lo}-{lo + 999}.json"


def _load_corpus(store_arg: str | None, jobs: list[str] | None) -> dict:
    if store_arg:
        store = line_items.load_store(Path(store_arg))
        if jobs:
            wanted = set(jobs)
            store = {"jobs": {j: r for j, r in (store.get("jobs") or {}).items() if j in wanted}}
        return store
    subprocess.run(["git", "-C", str(REPO), "fetch", "origin", "order-data"],
                   capture_output=True, text=True)
    if jobs:
        # Scoped: read only the shards those orders live in — fast, no full sweep.
        spans = {s for o in jobs if (s := _shard_span(o))}
        names = [f"{pre}.{s}" for s in spans for pre in ("backfill_line_items", "line_items")]
        wanted: set | None = set(jobs)
    else:
        ls = subprocess.run(["git", "-C", str(REPO), "ls-tree", "--name-only", "origin/order-data"],
                            capture_output=True, text=True)
        if ls.returncode != 0:
            sys.exit("Could not read origin/order-data — has the user published order data?")
        names = [n for n in ls.stdout.split() if n.endswith(".json")
                 and (n.startswith("line_items.jobs-") or n.startswith("backfill_line_items.jobs-"))]
        wanted = None
    # backfill first so the authoritative main shards win on any overlap.
    names.sort(key=lambda n: (not n.startswith("backfill_"), n))
    store: dict = {"jobs": {}}
    total = len(names)
    for i, name in enumerate(names, 1):
        if not wanted:
            print(f"  loading shard {i}/{total}: {name}  "
                  f"({len(store['jobs']):,} orders so far)", file=sys.stderr, flush=True)
        blob = subprocess.run(["git", "-C", str(REPO), "show", f"origin/order-data:{name}"],
                              capture_output=True, text=True)
        if blob.returncode == 0:
            jobs_in = (json.loads(blob.stdout).get("jobs") or {})
            if wanted is not None:
                jobs_in = {j: r for j, r in jobs_in.items() if j in wanted}
            store["jobs"].update(jobs_in)
    return store


def _renorm_with_progress(store: dict) -> None:
    """Re-derive the corpus in batches, printing progress to stderr so a long
    full-corpus run visibly stays alive. Batching is safe: renormalize_store
    processes each order independently, and the batch dicts hold the same order
    records as `store`, so the in-place re-derivation lands back in `store`."""
    jobs = list((store.get("jobs") or {}).items())
    total = len(jobs)
    total_items = sum(len((r or {}).get("items") or []) for _, r in jobs)
    ai_tags = store.get("ai_tags") or {}
    print(f"Re-deriving {total_items:,} line items across {total:,} orders with the "
          f"current rules…", file=sys.stderr, flush=True)
    batch = 300
    for i in range(0, total, batch):
        line_items.renormalize_store({"jobs": dict(jobs[i:i + batch]), "ai_tags": ai_tags})
        done = min(i + batch, total)
        pct = done / total if total else 1.0
        print(f"  … {done:,}/{total:,} orders re-derived ({pct:.0%})",
              file=sys.stderr, flush=True)
    print("Done re-deriving; computing the diff…", file=sys.stderr, flush=True)


def _coverage(store: dict) -> dict:
    items = comp = tagged = structured = flagged = 0
    for rec in (store.get("jobs") or {}).values():
        for it in (rec or {}).get("items") or []:
            items += 1
            attrs = it.get("attributes") if isinstance(it.get("attributes"), dict) else {}
            comp += bool(attrs.get("component"))
            tagged += bool(it.get("tags"))
            structured += any(k != "component" and v not in (None, "", [], {})
                              for k, v in (attrs or {}).items())
            flagged += bool(it.get("review_flags"))
    return {"items": items, "component": comp, "tagged": tagged,
            "structured": structured, "flagged": flagged}


def _sig(it: dict) -> tuple:
    attrs = it.get("attributes") if isinstance(it.get("attributes"), dict) else {}
    return (tuple(sorted(it.get("tags") or [])), attrs.get("component", ""),
            bool(it.get("review_flags")),
            json.dumps(attrs, sort_keys=True, ensure_ascii=False))


def _by_raw(items: list) -> dict:
    """Key items by their verbatim raw text + occurrence, so a renorm that drops
    a now-skipped item (or reclassifies one) lines up against its old self."""
    out, seen = {}, Counter()
    for it in items or []:
        raw = str(it.get("raw", ""))
        out[(raw, seen[raw])] = it
        seen[raw] += 1
    return out


def _component(it: dict | None) -> str:
    if not it:
        return ""
    attrs = it.get("attributes") if isinstance(it.get("attributes"), dict) else {}
    return attrs.get("component", "")


def _pct(n: int, d: int) -> float:
    return (n / d) if d else 0.0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--store", help="line-items store to read (default: order-data shards)")
    ap.add_argument("--jobs", nargs="+", metavar="ORDER",
                    help="scope to specific orders (fast); default is the whole "
                         "corpus, which re-derives everything and takes minutes")
    ap.add_argument("--json", action="store_true", help="machine-readable output")
    args = ap.parse_args()

    before = _load_corpus(args.store, args.jobs)
    scope = f"{len(before.get('jobs') or {})} order(s) you changed" if args.jobs \
        else f"{len(before.get('jobs') or {}):,} orders (whole corpus)"
    after = copy.deepcopy(before)
    if args.jobs:
        line_items.renormalize_store(after)   # scoped: fast, no progress needed
    else:
        _renorm_with_progress(after)

    cb, ca = _coverage(before), _coverage(after)
    rb = so_review.parser_review_metrics(before)
    ra = so_review.parser_review_metrics(after)

    changed = dropped = added = newly_comp = reclass = review_cleared = review_added = 0
    comp_gained: Counter = Counter()
    for job in set(before.get("jobs") or {}) | set(after.get("jobs") or {}):
        b = _by_raw((before["jobs"].get(job) or {}).get("items"))
        a = _by_raw((after["jobs"].get(job) or {}).get("items"))
        for key in set(b) | set(a):
            bi, ai = b.get(key), a.get(key)
            if bi and not ai:
                dropped += 1; changed += 1; continue
            if ai and not bi:
                added += 1; changed += 1
                if _component(ai):
                    comp_gained[_component(ai)] += 1
                continue
            if _sig(bi) == _sig(ai):
                continue
            changed += 1
            bc, ac = _component(bi), _component(ai)
            if not bc and ac:
                newly_comp += 1; comp_gained[ac] += 1
            elif bc and ac and bc != ac:
                reclass += 1; comp_gained[ac] += 1
            if bool(bi.get("review_flags")) and not bool(ai.get("review_flags")):
                review_cleared += 1
            if bool(ai.get("review_flags")) and not bool(bi.get("review_flags")):
                review_added += 1

    review_delta = rb["review_rows"] - ra["review_rows"]
    comp_delta = ca["component"] - cb["component"]
    result = {
        "orders": len(before.get("jobs") or {}),
        "items": cb["items"],
        "items_changed": changed,
        "newly_componentized": newly_comp,
        "reclassified": reclass,
        "dropped_from_capture": dropped,
        "newly_captured": added,
        "review_rows_before": rb["review_rows"],
        "review_rows_after": ra["review_rows"],
        "review_rows_fewer": review_delta,
        "flagged_items_before": cb["flagged"],
        "flagged_items_after": ca["flagged"],
        "component_rate_before": round(_pct(cb["component"], cb["items"]), 4),
        "component_rate_after": round(_pct(ca["component"], ca["items"]), 4),
        "tagged_rate_before": round(_pct(cb["tagged"], cb["items"]), 4),
        "tagged_rate_after": round(_pct(ca["tagged"], ca["items"]), 4),
        "component_items_gained": comp_delta,
        "top_components_gained": dict(comp_gained.most_common(10)),
    }

    if args.json:
        print(json.dumps(result, indent=2))
        return 0

    print(f"PARSER-CHANGE IMPACT  ·  {scope}, {result['items']:,} line items")
    print("-" * 64)
    print(f"Line items changed:        {changed:,}  "
          f"({_pct(changed, cb['items']):.1%} of the {cb['items']:,} examined)")
    print(f"  newly given a component: {newly_comp:,}")
    print(f"  reclassified:            {reclass:,}")
    if dropped:
        print(f"  dropped from capture:    {dropped:,}  (now matched by a skip rule)")
    if added:
        print(f"  newly captured:          {added:,}")
    print()
    _dir = "fewer" if review_delta >= 0 else "more"
    print(f"MARKED-FOR-REVIEW rows:    {rb['review_rows']:,} -> {ra['review_rows']:,}  "
          f"({abs(review_delta):,} {_dir}; {review_cleared:,} item flags cleared)")
    print(f"Flagged line items:        {cb['flagged']:,} -> {ca['flagged']:,}")
    print()
    print("Capture coverage (share of line items the parser resolves for you):")
    print(f"  component-classified:    {_pct(cb['component'], cb['items']):.1%} -> "
          f"{_pct(ca['component'], ca['items']):.1%}  ({comp_delta:+,} items)")
    print(f"  any canonical tag:       {_pct(cb['tagged'], cb['items']):.1%} -> "
          f"{_pct(ca['tagged'], ca['items']):.1%}")
    if comp_gained:
        gained = ", ".join(f"{c} (+{n})" for c, n in comp_gained.most_common(6))
        print(f"\nComponents gained: {gained}")
    print("\nEfficacy: this update changed "
          f"{changed:,} line item(s) — {newly_comp:,} now resolve to a named "
          f"component and {abs(review_delta):,} {_dir} MARKED-FOR-REVIEW row(s) "
          f"for a human to hand-categorize, moving component capture "
          f"{_pct(cb['component'], cb['items']):.1%} -> "
          f"{_pct(ca['component'], ca['items']):.1%} of everything on the sales orders.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
