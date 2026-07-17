"""Build the GL Queue Explorer — a self-contained HTML search page over the
line-items store, written next to the live workbook on the shared drive.

    python order_explorer.py                  # from the real stores (config paths)
    python order_explorer.py --out page.html  # write somewhere else
    python order_explorer.py --store s.json --master m.json --dwg d.json
                                              # from explicit files (dev/testing)

The page is ONE file with everything embedded (gzip+base64 payload, no server,
no internet, no install): coworkers double-click it — or the app-mode launcher
`Open GL Queue Explorer.bat` written beside it — and get click-driven search
the no-macro Excel tabs can't do:

  - search any job # (or its last digits) or free text over every captured
    Sales-Order line (normalized text + canonical tags) across the whole store;
  - a job's parsed spec, CO history, Open-PDF link, and its component
    hierarchy (so_hierarchy's rollup — the Sales Order tab's tree);
  - click a component -> every other order sharing it, most relevant first,
    each with the line items that made it a match, attribute agreement chips,
    its custom DWGs and CAD folder — find_orders' rarity-weighted scoring
    (shared tag = 1/df, identical normalized line = 2/df) computed on click;
  - click an attribute -> matches must carry it ([ATTRIBUTES] refinement);
  - or match the whole Sales Order (the Similar Orders view, uncapped).

Import-light on purpose (stdlib + so_hierarchy + config): build_payload /
render_html are pure and unit-tested (test_order_explorer.py); the store/dwg/
master loads happen only in maybe_write() and the CLI. The watcher calls
maybe_write() each poll — it regenerates only when a store file changed, the
board membership changed, or an hour passed, so the usual poll adds nothing.
"""
from __future__ import annotations

import argparse
import base64
import gzip
import json
import logging
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import so_hierarchy
from config import EXPLORER_PATH, LIVE_WORKBOOK_PATH, OUTPUT_DIR

log = logging.getLogger("order-explorer")

HTML_NAME = "GL Queue Explorer.html"
BAT_NAME = "Open GL Queue Explorer.bat"

# The parsed SO spec fields shown in the job header (mirrors the Sales Order
# tab's live_sheets.SO_SUMMARY_COLUMNS, minus the fields the header shows
# elsewhere: customer, CO#, rep).
SPEC_FIELDS = [
    ("Design", "design"), ("Description", "so_design_desc"), ("Size", "so_size"),
    ("Arrangement", "so_arrangement"), ("Class", "so_class"),
    ("Rotation", "so_rotation"), ("Discharge", "so_discharge"),
    ("Motor Pos", "so_motor_pos"), ("% Width", "so_pct_width"),
    ("Wheel", "so_wheel_type"), ("Design Temp", "so_design_temp"),
    ("Max Temp", "so_max_temp"), ("Total Price", "total_price"),
]
_HIST_MAX = 8            # CO-history entries kept per on-board job
_HIST_CLIP = 160         # ...each clipped to this many chars (some run to pages)


def default_output_path() -> Path:
    """EXPLORER_PATH from .env when set; else next to the live workbook (the
    shared location coworkers already know); else the local output folder."""
    if EXPLORER_PATH:
        return EXPLORER_PATH
    base = LIVE_WORKBOOK_PATH.parent if LIVE_WORKBOOK_PATH else OUTPUT_DIR
    return base / HTML_NAME


def _dwg_label(extras: Dict[str, str] | None) -> str:
    """'-07 (DWG), -51 (PDF+DWG)' — same form as find_orders._dwg_label."""
    return ", ".join(f"-{s} ({fmt})" if fmt else f"-{s}"
                     for s, fmt in (extras or {}).items())


def _attr_str(v: Any) -> str:
    if isinstance(v, list):
        return "; ".join(str(x) for x in v)
    return str(v)


def _comp_entry(c: Dict[str, Any]) -> Dict[str, Any]:
    """One so_hierarchy component record -> the page's compact form."""
    return {
        "n": c["name"],
        "k": 1 if c.get("keyed") else 0,
        "p": round(float(c.get("price") or 0), 2),
        "a": {k: _attr_str(v) for k, v in (c.get("attributes") or {}).items()},
        "r": [x["text"] + (f" (#{x['item_no']})" if x.get("item_no") else "")
              for x in c.get("review") or []],
        "i": [s["item_no"] for s in c.get("sources") or []],   # primary first
        "s": [_comp_entry(ch) for ch in c.get("children") or []],
    }


def _item_rows(items: List[Dict[str, Any]]) -> List[List[Any]]:
    """[no, text, price, qty, section, norm, tags] per captured item — the only
    per-item fields the page needs (attributes/review live on the components).
    `text` is so_hierarchy.line_text: the printed line with its leading
    numbering and trailing price columns stripped — price has its own slot."""
    return [[i, so_hierarchy.line_text(it), it.get("price", ""), it.get("qty", ""),
             it.get("section", ""), it.get("norm", ""),
             list(it.get("tags") or [])]
            for i, it in enumerate(items, start=1)]


def build_payload(store: Dict[str, Any],
                  dwg: Dict[str, Dict[str, Any]] | None = None,
                  master_orders: Dict[str, Dict[str, Any]] | None = None,
                  queue_jobs: Dict[str, Dict[str, Any]] | None = None) -> Dict[str, Any]:
    """The page's embedded data: every stored order with its items and derived
    component tree, enriched with DWG-scan and master-log facts where known.

    `queue_jobs` ({job -> job dict}) marks the on-board orders and supplies
    their FRESH line items / spec (the master job dict the watcher carries —
    the store can lag a poll or two behind for brand-new orders). When omitted,
    board membership falls back to the master log's on_queue flags."""
    dwg = dwg or {}
    master_orders = master_orders or {}
    queue_jobs = dict(queue_jobs or {})
    if not queue_jobs:
        for j, rec in master_orders.items():
            if rec.get("on_queue") and isinstance(rec.get("job"), dict):
                queue_jobs[str(j)] = rec["job"]

    jobs: Dict[str, Any] = {}
    all_jns = set(store.get("jobs") or {}) | set(queue_jobs)
    for jn in all_jns:
        rec = (store.get("jobs") or {}).get(jn) or {}
        qjob = queue_jobs.get(jn)
        mjob = qjob or (master_orders.get(jn) or {}).get("job") or {}
        items = (qjob.get("line_items") if qjob else None) or rec.get("items") or []

        entry: Dict[str, Any] = {
            "c": mjob.get("customer") or rec.get("customer") or "",
            "co": (lambda n: f"CO#{n}" if n else "")(
                mjob.get("co_number") or rec.get("co_number")),
            "pdf": (mjob.get("so_pdf") or rec.get("so_pdf") or "").strip(),
            "it": _item_rows(items),
            "cp": [_comp_entry(c) for c in so_hierarchy.components(items)],
        }
        drec = dwg.get(jn) or {}
        if drec.get("extras"):
            entry["d"] = _dwg_label(drec["extras"])
        if drec.get("folder"):
            entry["f"] = drec["folder"]
        if drec.get("type"):
            entry["t"] = drec["type"]
        if qjob:
            entry["q"] = 1
        spec = [[label, str(mjob.get(key)).strip()] for label, key in SPEC_FIELDS
                if str(mjob.get(key) or "").strip() not in ("", "None")]
        if spec:
            entry["sp"] = spec
        if qjob:
            hist = [str(h)[:_HIST_CLIP] for h in (qjob.get("co_history") or [])[:_HIST_MAX]]
            if hist:
                entry["h"] = hist
        jobs[str(jn)] = entry

    return {
        "gen": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "n_jobs": len(jobs),
        "n_items": sum(len(e["it"]) for e in jobs.values()),
        "jobs": jobs,
    }


def render_html(payload: Dict[str, Any]) -> str:
    """The complete page: template + the payload gzip+base64'd into it. Base64
    keeps the embedded data byte-safe inside <script> (no </script>/quoting
    hazards) and ~7x smaller than raw JSON — kinder to the shared drive."""
    raw = json.dumps(payload, separators=(",", ":"), ensure_ascii=False)
    b64 = base64.b64encode(gzip.compress(raw.encode("utf-8"), 9)).decode("ascii")
    return _TEMPLATE.replace("__B64__", b64)


def bat_text(html_name: str = HTML_NAME) -> str:
    """The app-mode launcher written beside the page: opens it in a clean
    Edge/Chrome app window (no tabs/address bar) so it feels like a program,
    not a website. Pure ASCII + CRLF (cmd.exe chokes on anything fancier)."""
    lines = [
        "@echo off",
        "setlocal",
        "rem Opens the GL Queue Explorer in a clean app window (no tabs, no",
        "rem address bar). Auto-generated by order_explorer.py next to the page;",
        "rem it finds the page beside itself, so the pair works from any folder.",
        "",
        f'set "PAGE=%~dp0{html_name}"',
        'if not exist "%PAGE%" (',
        f'    echo Could not find {html_name} next to this launcher.',
        "    pause",
        "    exit /b 1",
        ")",
        "",
        'set "EDGE=%ProgramFiles(x86)%\\Microsoft\\Edge\\Application\\msedge.exe"',
        'if not exist "%EDGE%" set "EDGE=%ProgramFiles%\\Microsoft\\Edge\\Application\\msedge.exe"',
        'if exist "%EDGE%" (',
        '    start "" "%EDGE%" --app="file:///%PAGE%"',
        "    exit /b 0",
        ")",
        "",
        'set "CHROME=%ProgramFiles%\\Google\\Chrome\\Application\\chrome.exe"',
        'if not exist "%CHROME%" set "CHROME=%ProgramFiles(x86)%\\Google\\Chrome\\Application\\chrome.exe"',
        'if exist "%CHROME%" (',
        '    start "" "%CHROME%" --app="file:///%PAGE%"',
        "    exit /b 0",
        ")",
        'start "" msedge --app="file:///%PAGE%"',
        "",
    ]
    return "\r\n".join(lines)


def write_explorer(payload: Dict[str, Any], out: Path | None = None) -> Path:
    """Write the page (atomically) and keep the .bat launcher beside it."""
    out = out or default_output_path()
    out.parent.mkdir(parents=True, exist_ok=True)
    html = render_html(payload)
    tmp = out.with_suffix(out.suffix + ".tmp")
    tmp.write_text(html, encoding="utf-8")
    tmp.replace(out)

    bat = out.parent / BAT_NAME
    text = bat_text(out.name)
    try:
        if not bat.exists() or bat.read_text(encoding="ascii", errors="replace") != text:
            bat.write_bytes(text.encode("ascii"))
    except OSError as e:  # the launcher is a nicety — never fail the page for it
        log.warning("Could not write %s (%s)", bat, e)
    log.info("Explorer written: %s (%d orders, %d items)",
             out, payload.get("n_jobs", 0), payload.get("n_items", 0))
    return out


# --------------------------------------------------------------------------- #
# Watcher hook: regenerate only when something it shows could have changed     #
# --------------------------------------------------------------------------- #
_CACHE: Dict[str, Any] = {"key": None, "at": 0.0}
_MAX_AGE_SECONDS = 3600     # regenerate at least hourly regardless


def maybe_write(master: Dict[str, Any] | None,
                lq_jobs: List[Dict[str, Any]] | None,
                force: bool = False) -> Optional[Path]:
    """Called by the watcher each poll. Cheap unless the line-items/DWG store
    files changed on disk, the board membership changed, or an hour passed —
    then the page is rebuilt from the real stores. Returns the path written,
    or None when the page was already current."""
    import autocad_scan
    import line_items as li

    out = default_output_path()
    def _mtime(p: Path) -> float:
        try:
            return p.stat().st_mtime
        except OSError:
            return 0.0
    queue = {str(j.get("job")): j for j in lq_jobs or [] if j.get("job")}
    key = (_mtime(li.store_path()), _mtime(autocad_scan.PROGRESS_PATH),
           tuple(sorted(queue)))
    now = time.time()
    if (not force and out.exists() and _CACHE["key"] == key
            and now - _CACHE["at"] < _MAX_AGE_SECONDS):
        return None

    payload = build_payload(li.load_store(), autocad_scan.load_progress(),
                            master_orders=(master or {}).get("orders"),
                            queue_jobs=queue)
    path = write_explorer(payload, out)
    _CACHE.update(key=key, at=now)
    return path


def main(argv: List[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--out", type=Path, default=None,
                    help=f"Output page path (default: {default_output_path()})")
    ap.add_argument("--store", type=Path, default=None,
                    help="Line-items store JSON (default: the configured store)")
    ap.add_argument("--master", type=Path, default=None,
                    help="live_master.json (default: the configured snapshot)")
    ap.add_argument("--dwg", type=Path, default=None,
                    help="AutoCAD scan store JSON (default: the configured store)")
    args = ap.parse_args(argv)

    import line_items as li
    store = li.load_store(args.store) if args.store else li.load_store()
    if not store.get("jobs"):
        print("The line-items store is empty — run the daily/backfill first.")
        return 1

    if args.dwg:
        dwg = json.loads(args.dwg.read_text(encoding="utf-8"))
    else:
        import autocad_scan
        dwg = autocad_scan.load_progress()

    if args.master:
        master = json.loads(args.master.read_text(encoding="utf-8"))
    else:
        import live_master
        master = live_master.load_master()

    payload = build_payload(store, dwg, master_orders=master.get("orders"))
    out = write_explorer(payload, args.out)
    n_q = sum(1 for e in payload["jobs"].values() if e.get("q"))
    print(f"Wrote {out}  ({payload['n_jobs']} orders, {payload['n_items']} line "
          f"items, {n_q} on the board)  + {BAT_NAME}")
    return 0


# --------------------------------------------------------------------------- #
# The page. One token: __B64__ (the gzip+base64 payload). No external          #
# resources of any kind — works from file:// on a shared drive, offline.       #
# --------------------------------------------------------------------------- #
_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>GL Queue Explorer</title>
<link rel="icon" href='data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text y=".9em" font-size="90">&%23127744;</text></svg>'>
<style>
  :root {
    --bg: #F4F6F5; --panel: #FFFFFF; --panel-2: #FAFBFB;
    --ink: #1B242D; --muted: #5C6B77; --faint: #8A98A3;
    --line: #DCE2E6; --accent: #C25E10; --accent-ink: #FFFFFF;
    --accent-soft: #F9EADD;
    --good: #1E7F4F; --good-soft: #DFF1E6;
    --bad: #B42318; --bad-soft: #F9E4E1;
    --chip: #EEF1F3; --hit: #FFF3C4;
    --mono: "Cascadia Code", Consolas, "SF Mono", ui-monospace, Menlo, monospace;
    --sans: "Segoe UI", "Avenir Next", system-ui, -apple-system, sans-serif;
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --bg: #12181E; --panel: #1A222A; --panel-2: #161D24;
      --ink: #E6EBEF; --muted: #8DA0AE; --faint: #64747F;
      --line: #2A343E; --accent: #E8873C; --accent-ink: #1A1208;
      --accent-soft: #342414;
      --good: #4CC38A; --good-soft: #173527;
      --bad: #F0907E; --bad-soft: #3A201B;
      --chip: #232E37; --hit: #3B3417;
    }
  }
  * { box-sizing: border-box; }
  html, body { margin: 0; }
  body { background: var(--bg); color: var(--ink); font: 14px/1.5 var(--sans);
         -webkit-font-smoothing: antialiased; }
  button { font: inherit; color: inherit; background: none; border: none;
           padding: 0; cursor: pointer; text-align: left; }
  a { color: var(--accent); font-weight: 600; text-decoration: none; }
  a:hover { text-decoration: underline; }
  button:focus-visible, input:focus-visible, a:focus-visible {
    outline: 2px solid var(--accent); outline-offset: 2px; border-radius: 3px; }

  .wrap { max-width: 1240px; margin: 0 auto; padding: 0 20px 40px; }
  header.top { display: flex; flex-wrap: wrap; align-items: center; gap: 12px 20px;
    padding: 18px 0 14px; border-bottom: 2px solid var(--ink); margin-bottom: 18px; }
  .wordmark { font-family: var(--mono); font-size: 15px; letter-spacing: .14em;
              font-weight: 700; white-space: nowrap; }
  .wordmark .dim { color: var(--muted); font-weight: 400; }
  .searchbox { position: relative; flex: 1 1 340px; max-width: 560px; margin-left: auto; }
  .searchbox input { width: 100%; padding: 9px 12px 9px 36px; font: 13.5px var(--sans);
    color: var(--ink); background: var(--panel); border: 1.5px solid var(--line);
    border-radius: 8px; }
  .searchbox input::placeholder { color: var(--faint); }
  .searchbox .glass { position: absolute; left: 12px; top: 50%;
    transform: translateY(-50%); color: var(--faint); pointer-events: none; }
  .search-drop { position: absolute; z-index: 30; top: calc(100% + 6px); left: 0; right: 0;
    background: var(--panel); border: 1px solid var(--line); border-radius: 10px;
    box-shadow: 0 10px 30px rgba(0,0,0,.14); overflow: hidden; display: none; }
  .search-drop.open { display: block; }
  .search-drop .sd-note { padding: 8px 14px; font-size: 11.5px; color: var(--muted);
    border-bottom: 1px solid var(--line); background: var(--panel-2); }
  .sd-item { display: block; width: 100%; padding: 9px 14px;
    border-bottom: 1px solid var(--line); }
  .sd-item:last-child { border-bottom: none; }
  .sd-item:hover { background: var(--accent-soft); }
  .sd-item .l1 { display: flex; gap: 10px; align-items: baseline; }
  .sd-item .job { font-family: var(--mono); font-weight: 700; }
  .sd-item .cust { color: var(--muted); font-size: 12.5px; overflow: hidden;
    text-overflow: ellipsis; white-space: nowrap; }
  .sd-item .onq { margin-left: auto; font-size: 10.5px; font-family: var(--mono);
    color: var(--good); letter-spacing: .06em; }
  .sd-item .why { font-family: var(--mono); font-size: 11.5px; color: var(--muted);
    margin-top: 2px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .sd-item .why mark { background: var(--hit); color: inherit; padding: 0 1px; }

  main.cols { display: grid; grid-template-columns: minmax(0, 11fr) minmax(0, 9fr);
    gap: 18px; align-items: start; }
  @media (max-width: 920px) { main.cols { grid-template-columns: 1fr; } }
  .panel { background: var(--panel); border: 1px solid var(--line);
    border-radius: 12px; overflow: hidden; }
  .panel-head { display: flex; align-items: baseline; gap: 10px; flex-wrap: wrap;
    padding: 12px 16px; border-bottom: 1px solid var(--line); background: var(--panel-2); }
  .eyebrow { font-family: var(--mono); font-size: 10.5px; letter-spacing: .14em;
    color: var(--muted); text-transform: uppercase; }
  .panel-body { padding: 14px 16px 16px; }
  .empty { padding: 46px 24px; text-align: center; color: var(--muted); }
  .empty .big { font-size: 15px; color: var(--ink); margin-bottom: 6px; }

  .qrow { display: flex; align-items: baseline; gap: 12px; width: 100%;
    padding: 10px 16px; border-bottom: 1px solid var(--line); }
  .qrow:last-child { border-bottom: none; }
  .qrow:hover { background: var(--accent-soft); }
  .qrow .job { font-family: var(--mono); font-weight: 700; font-size: 14px; }
  .qrow .cust { color: var(--muted); font-size: 12.5px; flex: 1; overflow: hidden;
    text-overflow: ellipsis; white-space: nowrap; }
  .qrow .n { font-family: var(--mono); font-size: 11.5px; color: var(--faint);
    white-space: nowrap; }

  .backlink { font-size: 12.5px; color: var(--accent); font-weight: 600; }
  .ohead { display: flex; align-items: baseline; gap: 10px; flex-wrap: wrap; flex: 1; }
  .ohead .job { font-family: var(--mono); font-size: 19px; font-weight: 700; }
  .ohead .cust { font-size: 13px; color: var(--muted); }
  .ohead .co { font-family: var(--mono); font-size: 11.5px; color: var(--muted);
    border: 1px solid var(--line); border-radius: 4px; padding: 1px 6px; }
  .ohead .onq { font-size: 10.5px; font-family: var(--mono); color: var(--good);
    letter-spacing: .06em; }
  .metaline { display: flex; flex-wrap: wrap; gap: 6px 16px; margin-top: 10px;
    font-size: 12px; align-items: baseline; }
  .metaline .dwg { font-family: var(--mono); color: var(--good); font-weight: 600; }
  .metaline .path { font-family: var(--mono); color: var(--muted); }
  .copybtn { font-size: 11px; color: var(--accent); border: 1px solid var(--accent);
    border-radius: 5px; padding: 0 6px; }
  .copybtn:hover { background: var(--accent-soft); }

  .specs { display: grid; grid-template-columns: repeat(auto-fill, minmax(112px, 1fr));
    gap: 8px 14px; margin: 12px 0 4px; }
  .spec .k { font-size: 10px; letter-spacing: .1em; text-transform: uppercase;
    color: var(--faint); font-family: var(--mono); }
  .spec .v { font-size: 12.5px; font-weight: 600; }
  details.hist { margin: 10px 0 0; font-size: 12px; }
  details.hist summary { cursor: pointer; color: var(--muted); font-weight: 600; }
  details.hist div { margin: 6px 0 0 14px; color: var(--muted);
    font-family: var(--mono); font-size: 11.5px; }

  .sectionbar { display: flex; align-items: baseline; gap: 10px; margin: 16px 0 8px;
    flex-wrap: wrap; }
  .sectionbar .hint { font-size: 11.5px; color: var(--muted); }
  .wholebtn { font-size: 12px; font-weight: 600; color: var(--accent);
    border: 1px solid var(--accent); border-radius: 999px; padding: 3px 12px;
    margin-left: auto; }
  .wholebtn:hover, .wholebtn.active { background: var(--accent);
    color: var(--accent-ink); }

  .tree { display: flex; flex-direction: column; gap: 6px; }
  .comp { border: 1px solid var(--line); border-radius: 9px; overflow: hidden;
    background: var(--panel); }
  .comp.active { border-color: var(--accent); box-shadow: 0 0 0 1px var(--accent); }
  .comp-row { display: flex; align-items: center; gap: 10px; width: 100%;
    padding: 8px 12px; background: var(--panel-2); }
  .comp-row:hover { background: var(--accent-soft); }
  .comp.active > .comp-row { background: var(--accent); color: var(--accent-ink); }
  .comp-row .name { font-family: var(--mono); font-weight: 700; font-size: 13px;
    overflow-wrap: anywhere; }
  .comp-row .meta { font-family: var(--mono); font-size: 11px; opacity: .75;
    white-space: nowrap; }
  .comp-row .price { margin-left: auto; font-family: var(--mono); font-size: 12px;
    font-variant-numeric: tabular-nums; white-space: nowrap; }
  .comp-row .go { font-size: 11px; font-family: var(--mono); opacity: .6;
    white-space: nowrap; }
  .comp-kids { padding: 6px 12px 8px 26px; display: flex; flex-direction: column;
    gap: 2px; }
  .attr-row { display: flex; gap: 8px; align-items: baseline; width: 100%;
    padding: 2px 6px; border-radius: 5px; font-size: 12.5px; }
  .attr-row:hover { background: var(--accent-soft); }
  .attr-row .k { color: var(--muted); white-space: nowrap; }
  .attr-row .v { font-weight: 600; overflow-wrap: anywhere; }
  .attr-row .pin { margin-left: auto; font-size: 10.5px; font-family: var(--mono);
    color: var(--faint); white-space: nowrap; }
  .attr-row.pinned { background: var(--accent-soft); }
  .attr-row.pinned .pin { color: var(--accent); font-weight: 700; }
  .rev-row { font-size: 12px; color: var(--bad); font-weight: 600; padding: 2px 6px;
    overflow-wrap: anywhere; }
  .src-row { font-size: 11.5px; color: var(--faint); font-family: var(--mono);
    padding: 1px 6px; overflow-wrap: anywhere; }
  .subwrap { margin: 6px 0 2px; border: 1px solid var(--line); border-radius: 8px;
    overflow: hidden; }

  .m-target { font-family: var(--mono); font-weight: 700; overflow-wrap: anywhere; }
  .m-count { margin-left: auto; font-size: 11.5px; color: var(--muted); }
  .filterbar { display: flex; flex-wrap: wrap; gap: 6px; padding: 10px 16px;
    border-bottom: 1px solid var(--line); background: var(--panel-2); }
  .filterbar .fl { font-size: 11px; color: var(--muted); align-self: center; }
  .fchip { display: inline-flex; gap: 6px; align-items: center; font-size: 11.5px;
    font-weight: 600; padding: 2px 9px; border-radius: 999px;
    background: var(--accent-soft); color: var(--accent);
    border: 1px solid var(--accent); }
  .match { padding: 12px 16px; border-bottom: 1px solid var(--line); }
  .match:last-child { border-bottom: none; }
  .m-head { display: flex; align-items: baseline; gap: 10px; flex-wrap: wrap; }
  .m-rank { font-family: var(--mono); font-size: 11px; color: var(--faint); width: 22px; }
  .m-job { font-family: var(--mono); font-size: 15px; font-weight: 700; }
  .m-job:hover { color: var(--accent); }
  .m-cust { font-size: 12px; color: var(--muted); }
  .m-score { margin-left: auto; font-family: var(--mono); font-size: 11.5px;
    font-variant-numeric: tabular-nums; color: var(--accent); font-weight: 700; }
  .m-scorebar { height: 3px; border-radius: 2px; background: var(--chip);
    margin: 6px 0 8px 32px; overflow: hidden; }
  .m-scorebar i { display: block; height: 100%; background: var(--accent); }
  .m-lines { margin-left: 32px; display: flex; flex-direction: column; gap: 3px; }
  .m-line { font-family: var(--mono); font-size: 12px; display: flex; gap: 8px;
    align-items: baseline; }
  .m-line .eq { color: var(--good); font-weight: 700; }
  .m-line .txt { overflow-wrap: anywhere; }
  .m-line .price { margin-left: auto; color: var(--muted);
    font-variant-numeric: tabular-nums; white-space: nowrap; }
  .m-more { font-size: 11px; color: var(--faint); font-family: var(--mono); }
  .m-chips { margin: 8px 0 0 32px; display: flex; flex-wrap: wrap; gap: 5px; }
  .chip { font-size: 11px; padding: 2px 8px; border-radius: 999px;
    background: var(--chip); color: var(--muted); }
  .chip.same { background: var(--good-soft); color: var(--good); font-weight: 600; }
  .chip.diff { background: var(--bad-soft); color: var(--bad); }
  .m-foot { margin: 8px 0 0 32px; display: flex; flex-wrap: wrap; gap: 6px 14px;
    font-size: 11.5px; align-items: baseline; }
  .m-foot .dwg { font-family: var(--mono); color: var(--good); font-weight: 600; }
  .m-foot .nodwg { font-family: var(--mono); color: var(--faint); }
  .m-foot .path { font-family: var(--mono); color: var(--muted);
    overflow-wrap: anywhere; }
  .tailnote { padding: 10px 16px; font-size: 11.5px; color: var(--faint);
    font-family: var(--mono); }

  footer.note { margin-top: 26px; padding-top: 14px; border-top: 1px solid var(--line);
    font-size: 12px; color: var(--muted); display: flex; gap: 8px; flex-wrap: wrap;
    justify-content: space-between; }
  footer.note .mono { font-family: var(--mono); }
  #boot { padding: 60px 24px; text-align: center; color: var(--muted); }
</style>
</head>
<body>
<div class="wrap">
  <header class="top">
    <span class="wordmark">GL QUEUE <span class="dim">/</span> ORDER EXPLORER</span>
    <div class="searchbox">
      <span class="glass">&#8981;</span>
      <input id="q" type="search" autocomplete="off" spellcheck="false"
             placeholder="Job # (or last digits) &mdash; or a feature: teflon, low leak, stainless&hellip;"
             aria-label="Search jobs or features">
      <div class="search-drop" id="drop"></div>
    </div>
  </header>
  <div id="boot">Loading the order data&hellip;</div>
  <main class="cols" style="display:none">
    <section class="panel" id="left"></section>
    <section class="panel" id="right"></section>
  </main>
  <footer class="note" style="display:none">
    <span>One file, no install, no internet &mdash; regenerated by the queue watcher.
      Scores: rare shared features count highest, identical Sales-Order lines double.</span>
    <span class="mono" id="stamp"></span>
  </footer>
</div>

<script>
"use strict";
const PAYLOAD_B64 = "__B64__";

let DB = null;                 // {gen, n_jobs, n_items, jobs: {job -> entry}}
let IDX = null;                // {tagDF, normDF, sets: {job -> {t:Set, n:Set}}}
const state = { job: null, path: null, whole: false, pinned: new Set() };

const $ = id => document.getElementById(id);
const esc = s => String(s).replace(/[&<>"']/g, c =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const money = n => (+n).toLocaleString("en-US",
  { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const jobNum = j => /^\d+/.test(j) ? parseInt(j, 10) : -1;
const fileUrl = p => "file:///" + encodeURI(String(p).replace(/\\/g, "/"))
  .replace(/#/g, "%23");

/* item row indices: [no, raw, price, qty, section, norm, tags] */
const IT = { NO: 0, RAW: 1, PRICE: 2, QTY: 3, SECTION: 4, NORM: 5, TAGS: 6 };

async function boot() {
  const bytes = Uint8Array.from(atob(PAYLOAD_B64), c => c.charCodeAt(0));
  const stream = new Blob([bytes]).stream()
    .pipeThrough(new DecompressionStream("gzip"));
  DB = JSON.parse(await new Response(stream).text());
  $("boot").style.display = "none";
  document.querySelector("main").style.display = "";
  document.querySelector("footer").style.display = "";
  $("stamp").textContent = "generated " + DB.gen + " · " + DB.n_jobs
    + " orders · " + DB.n_items + " line items";
  render();
  setTimeout(ensureIndex, 50);        // warm the match index off the first paint
}

function ensureIndex() {
  if (IDX) return IDX;
  const tagDF = {}, normDF = {}, sets = {};
  for (const j in DB.jobs) {
    const t = new Set(), n = new Set();
    for (const row of DB.jobs[j].it) {
      for (const tg of row[IT.TAGS]) t.add(tg);
      if (row[IT.NORM]) n.add(row[IT.NORM]);
    }
    sets[j] = { t, n };
    for (const x of t) tagDF[x] = (tagDF[x] || 0) + 1;
    for (const x of n) normDF[x] = (normDF[x] || 0) + 1;
  }
  IDX = { tagDF, normDF, sets };
  return IDX;
}

/* ---- component helpers (cp entries: {n,k,p,a,r,i,s}) ---------------------- */
function compAt(entry, path) {
  let list = entry.cp, c = null;
  for (const ix of path.split(".")) { c = list[+ix]; if (!c) return null; list = c.s || []; }
  return c;
}
function compParent(entry, path) {
  const parts = path.split(".");
  return parts.length > 1 ? compAt(entry, parts.slice(0, -1).join(".")) : null;
}
function itemByNo(entry, no) {
  return entry.it[no - 1] && entry.it[no - 1][IT.NO] === no
    ? entry.it[no - 1] : entry.it.find(r => r[IT.NO] === no);
}
function compItems(entry, path) {
  const c = compAt(entry, path);
  let nos = c && c.i.length ? c.i : [];
  if (!nos.length) {                      // synthetic child: fall back to parent's lines
    const p = compParent(entry, path);
    if (p) nos = p.i;
  }
  return nos.map(no => itemByNo(entry, no)).filter(Boolean);
}
function findCompByName(entry, name) {
  const walk = list => {
    for (const c of list || []) {
      if (c.n === name) return c;
      const hit = walk(c.s);
      if (hit) return hit;
    }
    return null;
  };
  return walk(entry.cp);
}
function anyAttrMatch(entry, key, val) {
  const walk = list => (list || []).some(c => {
    const have = c.a[key];
    if (have && (have === val || have.split(" | ").includes(val))) return true;
    return walk(c.s);
  });
  return walk(entry.cp);
}

/* ---- matching: rarity-weighted overlap (find_orders.similar_to_items) ----- */
function rankMatches(srcJob, items) {
  const { tagDF, normDF, sets } = ensureIndex();
  const tTags = new Set(), tNorms = new Set();
  for (const row of items) {
    for (const tg of row[IT.TAGS]) tTags.add(tg);
    if (row[IT.NORM]) tNorms.add(row[IT.NORM]);
  }
  const out = [];
  for (const j in DB.jobs) {
    if (j === srcJob) continue;
    const s = sets[j];
    let score = 0;
    const sharedNorms = [];
    for (const n of tNorms) if (s.n.has(n)) { score += 2 / normDF[n]; sharedNorms.push(n); }
    let anyTag = false;
    for (const t of tTags) if (s.t.has(t)) { score += 1 / tagDF[t]; anyTag = true; }
    if (!score) continue;
    if (state.pinned.size) {
      let ok = true;
      for (const p of state.pinned) {
        const ix = p.indexOf("=");
        if (!anyAttrMatch(DB.jobs[j], p.slice(0, ix), p.slice(ix + 1))) { ok = false; break; }
      }
      if (!ok) continue;
    }
    out.push({ j, score, sharedNorms: new Set(sharedNorms), tTags, anyTag });
  }
  out.sort((a, b) => b.score - a.score || jobNum(b.j) - jobNum(a.j));
  return out;
}

/* ------------------------------- rendering --------------------------------- */
function render() { renderLeft(); renderRight(); }
function selectJob(j) {
  state.job = j; state.path = null; state.whole = false; state.pinned.clear();
  render();
  window.scrollTo(0, 0);
}

function copyBtn(path) {
  return '<button class="copybtn" data-copy="' + esc(path)
    + '" title="Copy this path — paste it into File Explorer">copy path</button>';
}
function wireCopy(root) {
  root.querySelectorAll("[data-copy]").forEach(b => b.onclick = () => {
    const t = b.dataset.copy;
    const done = () => { b.textContent = "copied ✓";
      setTimeout(() => { b.textContent = "copy path"; }, 1200); };
    if (navigator.clipboard && navigator.clipboard.writeText)
      navigator.clipboard.writeText(t).then(done, () => fallbackCopy(t, done));
    else fallbackCopy(t, done);
  });
}
function fallbackCopy(t, done) {
  const ta = document.createElement("textarea");
  ta.value = t; document.body.appendChild(ta); ta.select();
  try { document.execCommand("copy"); done(); } catch (e) {}
  ta.remove();
}

function renderLeft() {
  const el = $("left");
  if (!state.job) {
    const onq = Object.keys(DB.jobs).filter(j => DB.jobs[j].q)
      .sort((a, b) => jobNum(a) - jobNum(b));
    el.innerHTML = '<div class="panel-head"><span class="eyebrow">On the board now</span>'
      + '<span class="m-count">' + onq.length
      + ' orders · click one, or search all ' + DB.n_jobs + ' above</span></div>'
      + onq.map(j => {
          const e = DB.jobs[j];
          return '<button class="qrow" data-job="' + esc(j) + '">'
            + '<span class="job">' + esc(j) + '</span>'
            + '<span class="cust">' + esc(e.c) + '</span>'
            + '<span class="n">' + e.it.length + ' items</span></button>';
        }).join("");
    el.querySelectorAll(".qrow").forEach(b => b.onclick = () => selectJob(b.dataset.job));
    return;
  }
  const j = state.job, e = DB.jobs[j];
  const specs = (e.sp || []).map(kv =>
    '<div class="spec"><div class="k">' + esc(kv[0]) + '</div><div class="v">'
    + esc(kv[1]) + '</div></div>').join("");
  const meta = [];
  if (e.t) meta.push('<span class="path">' + esc(e.t) + '</span>');
  if (e.d) meta.push('<span class="dwg">custom DWGs: ' + esc(e.d) + '</span>');
  if (e.f) meta.push('<span class="path">' + esc(e.f) + '</span>' + copyBtn(e.f));
  const hist = e.h ? '<details class="hist"><summary>CO history ('
    + e.h.length + ')</summary>'
    + e.h.map(x => "<div>" + esc(x) + "</div>").join("") + "</details>" : "";

  const compCard = (c, path) => {
    const active = !state.whole && state.path === path;
    const items = c.i.map(no => itemByNo(e, no)).filter(Boolean);
    const attrs = Object.entries(c.a).map(([k, v]) => {
      const key = k + "=" + v, pin = state.pinned.has(key);
      return '<button class="attr-row' + (pin ? " pinned" : "")
        + '" data-a="' + esc(key)
        + '" title="Click to require this attribute on every match">'
        + '<span class="k">' + esc(k.replace(/_/g, " ")) + ':</span>'
        + '<span class="v">' + esc(v) + '</span>'
        + '<span class="pin">' + (pin ? "✕ required" : "filter") + "</span></button>";
    }).join("");
    const revs = (c.r || []).map(x => '<div class="rev-row">' + esc(x) + "</div>").join("");
    const srcs = items.length > 1 ? items.map((row, i) =>
      '<div class="src-row">' + (i ? "+ " : "") + "#" + row[IT.NO] + " "
      + esc(row[IT.RAW]) + "</div>").join("") : "";
    const subs = (c.s || []).map((ch, ix) =>
      '<div class="subwrap">' + compCard(ch, path + "." + ix) + "</div>").join("");
    return '<div class="comp' + (active ? " active" : "") + '">'
      + '<button class="comp-row" data-c="' + path + '">'
      + '<span class="name">' + (c.k ? "[" + esc(c.n) + "]" : esc(c.n)) + "</span>"
      + '<span class="meta">' + (items.length || "") + (items.length > 1 ? " lines"
        : items.length === 1 ? " line" : "") + "</span>"
      + '<span class="price">' + (c.p ? money(c.p) : "") + "</span>"
      + '<span class="go">find matches ▸</span></button>'
      + '<div class="comp-kids">' + attrs + revs + subs + srcs + "</div></div>";
  };
  const tree = e.cp.length
    ? e.cp.map((c, ix) => compCard(c, String(ix))).join("")
    : '<div class="empty">No line items captured for this order yet.</div>';

  el.innerHTML = '<div class="panel-head">'
    + '<button class="backlink" id="back">← board</button>'
    + '<div class="ohead"><span class="job">' + esc(j) + '</span>'
    + '<span class="cust">' + esc(e.c) + "</span>"
    + (e.co ? '<span class="co">' + esc(e.co) + "</span>" : "")
    + (e.q ? '<span class="onq">ON BOARD</span>' : "")
    + (e.pdf ? '<a href="' + esc(fileUrl(e.pdf)) + '" target="_blank" title="'
        + esc(e.pdf) + '">Open SO PDF</a>' : "")
    + "</div></div>"
    + '<div class="panel-body">'
    + (specs ? '<div class="specs">' + specs + "</div>" : "")
    + (meta.length ? '<div class="metaline">' + meta.join(" ") + "</div>" : "")
    + hist
    + '<div class="sectionbar"><span class="eyebrow">Components</span>'
    + '<span class="hint">click one to rank past orders that share it</span>'
    + '<button class="wholebtn' + (state.whole ? " active" : "")
    + '" id="whole">match whole order</button></div>'
    + '<div class="tree">' + tree + "</div></div>";

  $("back").onclick = () => { state.job = null; state.path = null;
    state.whole = false; state.pinned.clear(); render(); };
  $("whole").onclick = () => { state.whole = !state.whole;
    if (state.whole) state.path = null; render(); };
  el.querySelectorAll(".comp-row").forEach(b => b.onclick = () => {
    state.whole = false; state.path = b.dataset.c; state.pinned.clear(); render();
  });
  el.querySelectorAll(".attr-row").forEach(b => b.onclick = () => {
    const k = b.dataset.a;
    state.pinned.has(k) ? state.pinned.delete(k) : state.pinned.add(k);
    render();
  });
  wireCopy(el);
}

function renderRight() {
  const el = $("right");
  const ready = state.job && (state.whole || state.path !== null);
  if (!ready) {
    el.innerHTML = '<div class="panel-head"><span class="eyebrow">Matching orders</span></div>'
      + '<div class="empty"><div class="big">'
      + (state.job ? "Click a component on the left" : "Pick an order first")
      + "</div>Past orders sharing it appear here, most relevant first, each with "
      + "the line items that made it a match.</div>";
    return;
  }
  const j = state.job, e = DB.jobs[j];
  const items = state.whole ? e.it : compItems(e, state.path);
  const target = state.whole ? "whole order" : (() => {
    const c = compAt(e, state.path);
    return c ? (c.k ? "[" + c.n + "]" : c.n) : "?";
  })();
  const res = rankMatches(j, items);
  const shown = res.slice(0, 25);
  const max = res.length ? res[0].score : 1;
  const chips = [...state.pinned].map(p => {
    const ix = p.indexOf("=");
    return '<button class="fchip" data-a="' + esc(p) + '" title="Remove this filter">'
      + esc(p.slice(0, ix).replace(/_/g, " ")) + ": " + esc(p.slice(ix + 1))
      + " <span>✕</span></button>";
  }).join("");

  const targetComp = state.whole ? null : compAt(e, state.path);
  const cards = shown.map((r, i) => {
    const o = DB.jobs[r.j];
    const lines = o.it.filter(row => r.sharedNorms.has(row[IT.NORM])
      || row[IT.TAGS].some(t => r.tTags.has(t)));
    const head = lines.slice(0, 4).map(row =>
      '<div class="m-line"><span class="eq">=</span><span class="txt">'
      + esc(row[IT.RAW]) + '</span><span class="price">'
      + esc(row[IT.PRICE]) + "</span></div>").join("");
    const more = lines.length > 4
      ? '<div class="m-more">+ ' + (lines.length - 4) + " more shared lines</div>" : "";
    let chipsHtml = "";
    if (targetComp) {
      const theirs = findCompByName(o, targetComp.n);
      if (theirs) {
        const same = [], diff = [];
        for (const [k, v] of Object.entries(targetComp.a)) {
          if (!(k in theirs.a)) continue;
          if (theirs.a[k] === v) same.push(k.replace(/_/g, " ") + ": " + v);
          else diff.push(k.replace(/_/g, " ") + ": theirs " + theirs.a[k]);
        }
        if (same.length || diff.length)
          chipsHtml = '<div class="m-chips">'
            + same.map(a => '<span class="chip same">✓ ' + esc(a) + "</span>").join("")
            + diff.map(a => '<span class="chip diff">≠ ' + esc(a) + "</span>").join("")
            + "</div>";
      }
    }
    const foot = ['<span class="' + (o.d ? "dwg" : "nodwg") + '">'
      + (o.d ? "custom DWGs: " + esc(o.d) : "no custom DWGs") + "</span>"];
    if (o.f) foot.push('<span class="path">' + esc(o.f) + "</span>" + copyBtn(o.f));
    if (o.pdf) foot.push('<a href="' + esc(fileUrl(o.pdf))
      + '" target="_blank" title="' + esc(o.pdf) + '">SO PDF</a>');
    return '<div class="match"><div class="m-head">'
      + '<span class="m-rank">' + (i + 1) + ".</span>"
      + '<button class="m-job" data-job="' + esc(r.j) + '" title="Open this order">'
      + esc(r.j) + "</button>"
      + '<span class="m-cust">' + esc(o.c) + (o.co ? " · " + esc(o.co) : "")
      + (o.q ? " · ON BOARD" : "") + "</span>"
      + '<span class="m-score">score ' + r.score.toFixed(2) + "</span></div>"
      + '<div class="m-scorebar"><i style="width:'
      + Math.max(6, 100 * r.score / max) + '%"></i></div>'
      + '<div class="m-lines">' + head + more + "</div>"
      + chipsHtml + '<div class="m-foot">' + foot.join(" ") + "</div></div>";
  }).join("");

  el.innerHTML = '<div class="panel-head"><span class="eyebrow">Matching orders</span>'
    + '<span class="m-target">' + esc(target) + '</span>'
    + '<span class="m-cust">on ' + esc(j) + "</span>"
    + '<span class="m-count">' + res.length + " match"
    + (res.length === 1 ? "" : "es") + "</span></div>"
    + (state.pinned.size
        ? '<div class="filterbar"><span class="fl">Required attributes:</span>'
          + chips + "</div>" : "")
    + (res.length ? cards : '<div class="empty"><div class="big">No orders match</div>'
        + (state.pinned.size ? "Try removing an attribute filter."
           : "Nothing in the store shares this yet.") + "</div>")
    + (res.length > 25 ? '<div class="tailnote">…and ' + (res.length - 25)
        + " more — narrow it with an attribute filter</div>" : "");
  el.querySelectorAll(".fchip").forEach(b => b.onclick = () => {
    state.pinned.delete(b.dataset.a); render(); });
  el.querySelectorAll(".m-job").forEach(b => b.onclick = () => selectJob(b.dataset.job));
  wireCopy(el);
}

/* ------------------------------- search ------------------------------------ */
const q = $("q"), drop = $("drop");
let searchTimer = null;
function doSearch() {
  const v = q.value.trim();
  if (!v || !DB) { drop.classList.remove("open"); return; }
  const hits = [];
  let total = 0;
  if (/^\d+$/.test(v)) {
    const all = Object.keys(DB.jobs).sort((a, b) => jobNum(b) - jobNum(a));
    for (const j of all) {
      if (j.endsWith(v) || j.startsWith(v)) {
        total++;
        if (hits.length < 10) hits.push({ j, why: "" });
      }
    }
  } else {
    const needle = v.toUpperCase().replace(/[^A-Z0-9 ]/g, " ")
      .replace(/\s+/g, " ").trim();
    if (!needle) { drop.classList.remove("open"); return; }
    const all = Object.keys(DB.jobs).sort((a, b) => jobNum(b) - jobNum(a));
    for (const j of all) {
      const e = DB.jobs[j];
      let why = "";
      for (const row of e.it) {
        const at = row[IT.NORM].indexOf(needle);
        if (at >= 0) {
          why = esc(row[IT.NORM].slice(0, at)) + "<mark>"
            + esc(row[IT.NORM].slice(at, at + needle.length)) + "</mark>"
            + esc(row[IT.NORM].slice(at + needle.length));
          break;
        }
        if (row[IT.TAGS].some(t => t.includes(needle))) {
          why = "tag: " + esc(row[IT.TAGS].find(t => t.includes(needle)));
          break;
        }
      }
      if (why) {
        total++;
        if (hits.length < 10) hits.push({ j, why });
      }
    }
  }
  const note = /^\d+$/.test(v)
    ? "Job-number match — full number or just the last few digits"
    : total + " order(s) with a Sales-Order line matching “" + esc(v)
      + "”" + (total > 10 ? " — first 10 shown" : "");
  drop.innerHTML = '<div class="sd-note">' + note + "</div>"
    + (hits.length ? hits.map(h => {
        const e = DB.jobs[h.j];
        return '<button class="sd-item" data-job="' + esc(h.j) + '">'
          + '<span class="l1"><span class="job">' + esc(h.j) + "</span>"
          + '<span class="cust">' + esc(e.c) + "</span>"
          + (e.q ? '<span class="onq">ON BOARD</span>' : "") + "</span>"
          + (h.why ? '<span class="why">= ' + h.why + "</span>" : "")
          + "</button>";
      }).join("")
      : '<div class="sd-note">No order matches — try another spelling, '
        + "a tag like “shaft seal”, or a job #.</div>");
  drop.classList.add("open");
  drop.querySelectorAll(".sd-item").forEach(b => b.onclick = () => {
    selectJob(b.dataset.job); drop.classList.remove("open"); q.value = "";
  });
}
q.addEventListener("input", () => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(doSearch, 120);
});
q.addEventListener("keydown", ev => {
  if (ev.key === "Enter") { const f = drop.querySelector(".sd-item"); if (f) f.click(); }
  if (ev.key === "Escape") drop.classList.remove("open");
});
document.addEventListener("click", ev => {
  if (!ev.target.closest(".searchbox")) drop.classList.remove("open");
});

boot().catch(err => {
  $("boot").textContent = "Could not load the embedded data (" + err
    + "). This page needs a Chromium browser (Edge or Chrome).";
});
</script>
</body>
</html>
"""


if __name__ == "__main__":
    sys.exit(main())
