"""Build the GL Queue Explorer — a self-contained HTML companion to the live
workbook, written next to it on the shared drive.

    python order_explorer.py                  # from the real stores (config paths)
    python order_explorer.py --out page.html  # write somewhere else
    python order_explorer.py --store s.json --master m.json --dwg d.json
                                              # from explicit files (dev/testing)

The page is ONE file with everything embedded (gzip+base64 payload, no server,
no internet, no install): coworkers double-click it — or the app-mode launcher
`Open GL Queue Explorer.bat` written beside it — and get four views:

  Board          the Live Queue with the workbook's exact color coding
                 (end-date urgency reds/orange/gold, grey new-today rows and
                 their blended urgent variants, red text on a CO landing,
                 orange Quote Run marks) and the totals footer.
  Changes        today's activity log, mirroring the Changes tab's sections:
                 new orders, change orders (with the CO description resolved
                 from co_history), the field-modification log (grey ladder per
                 later instance), and removals.
  Order History  every order the master log has ever seen (stable spec columns
                 + On Queue/Added/Left), filterable; the green-✓/red custom-DWG
                 and feature matrices appear once the filter is narrow enough
                 for a browser to draw them (Excel virtualizes 13K x 150 cells;
                 a DOM can't).
  Job view       click any job # anywhere: its parsed spec, CO history, and
                 component hierarchy (so_hierarchy's rollup). Click a component
                 to rank every other order sharing it (find_orders' rarity-
                 weighted scoring — shared tag 1/df, identical line 2/df —
                 computed on click), pin attributes to filter, or match the
                 whole order (the Similar Orders view, uncapped).

AUTO-REFRESH: every write also updates `gl_queue_explorer_version.js` beside
the page. Open pages poll that stamp each minute via a <script> tag (the only
cross-file read a file:// page is allowed) and, when it moves, reload
themselves — restoring the view they were on. Delete the .js and pages simply
stop auto-refreshing; nothing breaks.

build_payload / render_html are pure and unit-tested (test_order_explorer.py);
store loads live only in maybe_write() and the CLI. The watcher calls
maybe_write() each poll — it regenerates only when a store file, the day's
change log, or the board membership changed (hourly floor), so the usual poll
adds nothing. The scoring formula is duplicated in the page's JS by necessity
(no Python inside a browser) — if find_orders.similar_to_items' weights ever
change, update the JS to match.
"""
from __future__ import annotations

import argparse
import base64
import gzip
import json
import logging
import sys
import time
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import engineers
import so_hierarchy
from config import EXPLORER_PATH, LIVE_WORKBOOK_PATH, OUTPUT_DIR
# Pure display helpers shared with the Excel tabs, so CO descriptions and the
# tidy Size / Arrangement forms never diverge between the two.
from excel_writer import split_arrangement, split_size
from live_sheets import _co_change_desc

log = logging.getLogger("order-explorer")

HTML_NAME = "GL Queue Explorer.html"
BAT_NAME = "Open GL Queue Explorer.bat"
VERSION_NAME = "gl_queue_explorer_version.js"
VBS_NAME = "glq_open.vbs"
ENABLE_NAME = "Enable Folder Links.bat"

# The parsed SO spec fields shown in the job header and the Order History
# columns (mirrors live_sheets.SO_SUMMARY_COLUMNS / OH_DATA_COLUMNS).
SPEC_FIELDS = [
    ("Design", "design"), ("Description", "so_design_desc"), ("Size", "so_size"),
    ("Arrangement", "so_arrangement"), ("Class", "so_class"),
    ("Rotation", "so_rotation"), ("Discharge", "so_discharge"),
    ("Motor Pos", "so_motor_pos"), ("% Width", "so_pct_width"),
    ("Wheel", "so_wheel_type"), ("Design Temp", "so_design_temp"),
    ("Max Temp", "so_max_temp"), ("Special Temp", "so_special_temp"),
    ("Total Price", "total_price"), ("Primary Rep", "primary_rep"),
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


def _spec_value(key: str, v: str) -> str:
    """Tidy display forms, identical to the workbook's columns: the size keeps
    only its leading number ('2412 (3600 RPM or less)' -> '2412') and the
    arrangement its short code ('Arrangement 4' -> 'A/4', 'A/4V C-Face Flange
    mount (no motor base)' -> 'A/4V')."""
    if key == "so_size":
        return split_size(v)[0]
    if key == "so_arrangement":
        return split_arrangement(v)[0]
    return v


def _board_fields(j: Dict[str, Any], added: str = "") -> Dict[str, Any]:
    """The churny Live Queue columns for one on-board order, keys matched to
    the page's Board table. `added` is the master entry's arrival timestamp —
    the same value the workbook's 'Added' column reflects — used when the job
    dict wasn't stamped by a live poll (snapshot builds)."""
    bd = {
        "st": j.get("status", ""), "op": j.get("oper", ""),
        "as": j.get("assigned_to", ""), "ck": j.get("checker", ""),
        "no": j.get("status_note", ""), "sd": j.get("start_date", ""),
        "ed": j.get("end_date", ""), "hr": j.get("plan_hrs", ""),
        "fn": j.get("fannet_date", ""), "pr": j.get("total_price", ""),
        "ru": j.get("dwg_reuse_label", ""),
        "ai": j.get("_added_iso") or added or j.get("_first_seen") or "",
    }
    if j.get("has_drive_run"):
        bd["dr"] = 1
    pos = j.get("_cbc_pos")
    if isinstance(pos, int):
        bd["ps"] = pos
    return {k: v for k, v in bd.items() if v not in ("", None)}


def _events_payload(events: List[Dict[str, Any]],
                    master_orders: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Today's change-log entries, newest first, with the change-order
    description resolved from co_history at build time (same helper as the
    Excel Changes tab)."""
    out = []
    for e in sorted(events or [], key=lambda x: x.get("time", ""), reverse=True):
        row = {"t": e.get("time", ""), "j": str(e.get("job", "")),
               "c": e.get("customer", ""), "f": e.get("field", ""),
               "o": str(e.get("old", "")), "n": str(e.get("new", ""))}
        if row["f"] == "CO#":
            order = (master_orders.get(row["j"]) or {}).get("job") or {}
            desc = _co_change_desc(order, e.get("new"))
            if desc:
                row["d"] = desc[:400]
        out.append(row)
    return out


def build_payload(store: Dict[str, Any],
                  dwg: Dict[str, Dict[str, Any]] | None = None,
                  master_orders: Dict[str, Dict[str, Any]] | None = None,
                  queue_jobs: Dict[str, Dict[str, Any]] | None = None,
                  events: List[Dict[str, Any]] | None = None,
                  today: date | None = None,
                  new_ids: set | None = None,
                  sw: Dict[str, Dict[str, Any]] | None = None) -> Dict[str, Any]:
    """The page's embedded data: every order the store OR the master log knows
    (the master-only ones carry no line items but fill Order History), with
    items + derived component tree, DWG-scan facts, board fields for on-board
    orders, today's change events, and today's removals.

    `queue_jobs` ({job -> job dict}) marks the on-board orders and supplies
    their FRESH line items / spec (the master job dict the watcher carries —
    the store can lag a poll or two behind for brand-new orders). When omitted,
    board membership falls back to the master log's on_queue flags."""
    dwg = dwg or {}
    master_orders = master_orders or {}
    queue_jobs = dict(queue_jobs or {})
    today = today or date.today()
    if not queue_jobs:
        for j, rec in master_orders.items():
            if rec.get("on_queue") and isinstance(rec.get("job"), dict):
                queue_jobs[str(j)] = rec["job"]

    jobs: Dict[str, Any] = {}
    all_jns = (set(store.get("jobs") or {}) | set(queue_jobs)
               | {str(j) for j in master_orders})
    for jn in all_jns:
        rec = (store.get("jobs") or {}).get(jn) or {}
        ment = master_orders.get(jn) or {}
        qjob = queue_jobs.get(jn)
        mjob = qjob or ment.get("job") or {}
        items = (qjob.get("line_items") if qjob else None) or rec.get("items") or []

        entry: Dict[str, Any] = {
            "c": mjob.get("customer") or rec.get("customer") or "",
            "co": (lambda n: f"CO#{n}" if n else "")(
                mjob.get("co_number") or rec.get("co_number")),
            "pdf": (mjob.get("so_pdf") or rec.get("so_pdf") or "").strip(),
            "it": _item_rows(items),
            "cp": [_comp_entry(c) for c in so_hierarchy.components(items)],
        }
        qr = (mjob.get("drive_run_pdf") or "").strip()
        if qr:
            entry["qr"] = qr
        swrec = (sw or {}).get(jn) or {}
        if swrec.get("has_sw") and swrec.get("folder"):
            entry["sw"] = swrec["folder"]
        drec = dwg.get(jn) or {}
        extras = drec.get("extras") or mjob.get("dwg_extras") or {}
        if extras:
            entry["d"] = _dwg_label(extras)
            entry["dx"] = sorted(extras)
        if drec.get("folder") or mjob.get("job_folder"):
            entry["f"] = drec.get("folder") or mjob.get("job_folder")
        if drec.get("type") or mjob.get("job_type"):
            entry["t"] = drec.get("type") or mjob.get("job_type")
        eng = engineers.cell_text(mjob) if mjob else ""
        if eng:
            entry["e"] = eng
        if mjob.get("item"):
            entry["im"] = mjob["item"]
        spec = [[label, _spec_value(key, str(mjob.get(key)).strip())]
                for label, key in SPEC_FIELDS
                if str(mjob.get(key) or "").strip() not in ("", "None")]
        if spec:
            entry["sp"] = spec
        if ment:      # in the master log -> an Order History row
            entry["oh"] = [1 if ment.get("on_queue") else 0,
                           str(ment.get("added") or ""), str(ment.get("left") or "")]
        if qjob:
            entry["q"] = 1
            entry["bd"] = _board_fields(qjob, added=str(ment.get("added") or ""))
            hist = [str(h)[:_HIST_CLIP] for h in (qjob.get("co_history") or [])[:_HIST_MAX]]
            if hist:
                entry["h"] = hist
        jobs[str(jn)] = entry

    removed = sorted(
        ([str(j), str(e.get("left") or "")] for j, e in master_orders.items()
         if e.get("seen_on_queue") and not e.get("on_queue")
         and str(e.get("left") or "")[:10] == today.isoformat()),
        key=lambda r: r[1], reverse=True)

    payload = {
        "gen": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "today": today.isoformat(),
        "n_jobs": len(jobs),
        "n_items": sum(len(e["it"]) for e in jobs.values()),
        "jobs": jobs,
        "ev": _events_payload(events or [], master_orders),
        "rm": removed,
    }
    # The watcher's own new-today set (snapshot-diff based), so the grey
    # new-order shading matches the workbook EXACTLY. When absent (snapshot
    # builds) the page falls back to comparing arrival dates.
    if new_ids is not None:
        payload["nw"] = sorted(str(x) for x in new_ids)
    return payload


def render_html(payload: Dict[str, Any]) -> str:
    """The complete page: template + the payload gzip+base64'd into it. Base64
    keeps the embedded data byte-safe inside <script> (no </script>/quoting
    hazards) and ~7x smaller than raw JSON — kinder to the shared drive."""
    raw = json.dumps(payload, separators=(",", ":"), ensure_ascii=False)
    b64 = base64.b64encode(gzip.compress(raw.encode("utf-8"), 9)).decode("ascii")
    return _TEMPLATE.replace("__B64__", b64)


def version_js(payload: Dict[str, Any]) -> str:
    """The tiny sibling stamp open pages poll to notice a rewrite."""
    return f'window.__GLQ_VERSION__ = "{payload.get("gen", "")}";\n'


def vbs_text() -> str:
    """The windowless glqueue: protocol handler written beside the page: takes
    the clicked link's URI, strips the scheme, percent-decodes, flips slashes,
    and opens the folder in Windows File Explorer. wscript runs it with no
    console flash. Registered per-user (HKCU, no admin) by the .bat and the
    --open flow. Pure ASCII + CRLF."""
    lines = [
        "' Opens a glqueue: folder link in Windows File Explorer. Registered as",
        "' the per-user glqueue: protocol handler by Open GL Queue Explorer.bat",
        "' (or the launcher's Open button); regenerated by order_explorer.py.",
        "Option Explicit",
        "",
        "Function DecodeUrl(s)",
        "  Dim i, c, h, r",
        '  r = ""',
        "  i = 1",
        "  Do While i <= Len(s)",
        "    c = Mid(s, i, 1)",
        '    If c = "%" And i + 2 <= Len(s) Then',
        "      h = Mid(s, i + 1, 2)",
        "      On Error Resume Next",
        "      Err.Clear",
        '      c = Chr(CLng("&H" & h))',
        "      If Err.Number = 0 Then",
        "        i = i + 3",
        "      Else",
        '        c = "%"',
        "        i = i + 1",
        "      End If",
        "      On Error GoTo 0",
        "    Else",
        "      i = i + 1",
        "    End If",
        "    r = r & c",
        "  Loop",
        "  DecodeUrl = r",
        "End Function",
        "",
        "Dim u",
        "If WScript.Arguments.Count = 0 Then WScript.Quit 1",
        "u = WScript.Arguments(0)",
        'If LCase(Left(u, 8)) = "glqueue:" Then u = Mid(u, 9)',
        'Do While Left(u, 1) = "/"',
        "  u = Mid(u, 2)",
        "Loop",
        'u = Replace(DecodeUrl(u), "/", "\\")',
        'If u <> "" Then',
        '  CreateObject("WScript.Shell").Run "explorer.exe """ & u & """", 1, False',
        "End If",
        "",
    ]
    return "\r\n".join(lines)


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
        "rem Per-user (no admin), idempotent: register the glqueue: handler so",
        "rem the page's folder links open in Windows File Explorer.",
        'if exist "%~dp0glq_open.vbs" (',
        '    reg add "HKCU\\Software\\Classes\\glqueue" /ve /t REG_SZ /d "URL:GL Queue folder link" /f >nul 2>nul',
        '    reg add "HKCU\\Software\\Classes\\glqueue" /v "URL Protocol" /t REG_SZ /d "" /f >nul 2>nul',
        '    reg add "HKCU\\Software\\Classes\\glqueue\\shell\\open\\command" /ve /t REG_SZ /d "wscript.exe \\"%~dp0glq_open.vbs\\" \\"%%1\\"" /f >nul 2>nul',
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
    """Write the page (atomically), bump the version stamp open pages poll,
    and keep the .bat launcher beside it. The stamp is replaced strictly AFTER
    the page, so a page that notices the new stamp always reloads new data."""
    out = out or default_output_path()
    out.parent.mkdir(parents=True, exist_ok=True)
    html = render_html(payload)
    tmp = out.with_suffix(out.suffix + ".tmp")
    tmp.write_text(html, encoding="utf-8")
    tmp.replace(out)

    ver = out.parent / VERSION_NAME
    try:
        vtmp = ver.with_suffix(".js.tmp")
        vtmp.write_text(version_js(payload), encoding="utf-8")
        vtmp.replace(ver)
    except OSError as e:  # auto-refresh is a nicety — never fail the page for it
        log.warning("Could not write %s (%s)", ver, e)

    bat = out.parent / BAT_NAME
    text = bat_text(out.name)
    try:
        if not bat.exists() or bat.read_text(encoding="ascii", errors="replace") != text:
            bat.write_bytes(text.encode("ascii"))
    except OSError as e:
        log.warning("Could not write %s (%s)", bat, e)

    vbs = out.parent / VBS_NAME
    vtext = vbs_text()
    try:
        if not vbs.exists() or vbs.read_text(encoding="ascii", errors="replace") != vtext:
            vbs.write_bytes(vtext.encode("ascii"))
    except OSError as e:
        log.warning("Could not write %s (%s)", vbs, e)
    log.info("Explorer written: %s (%d orders, %d items)",
             out, payload.get("n_jobs", 0), payload.get("n_items", 0))
    return out


def _ensure_folder_link_handler(page_dir: Path) -> None:
    """Register the per-user glqueue: protocol (HKCU — no admin) pointing at
    the glq_open.vbs beside the page, so folder links open in File Explorer.
    Idempotent; mirrors the registration the .bat performs."""
    import os
    import subprocess
    if os.name != "nt":
        return
    vbs = page_dir / VBS_NAME
    if not vbs.exists():
        return
    base = r"HKCU\Software\Classes\glqueue"
    for cmd in (
        ["reg", "add", base, "/ve", "/t", "REG_SZ",
         "/d", "URL:GL Queue folder link", "/f"],
        ["reg", "add", base, "/v", "URL Protocol", "/t", "REG_SZ", "/d", "", "/f"],
        ["reg", "add", base + r"\shell\open\command", "/ve", "/t", "REG_SZ",
         "/d", f'wscript.exe "{vbs}" "%1"', "/f"],
    ):
        subprocess.run(cmd, capture_output=True, check=False)


def open_in_app_window(page: Path) -> bool:
    """Open the page as a clean app window (no tabs/address bar): Edge first
    (ships with Windows), then Chrome, then the OS default handler. The same
    fallback chain as the .bat launcher, for the launcher's Open button."""
    import os
    import subprocess
    if os.name != "nt":
        print(f"Open this file in a Chromium browser: {page}")
        return False
    _ensure_folder_link_handler(page.parent)
    url = "file:///" + str(page).replace("\\", "/")
    candidates = (
        ("ProgramFiles(x86)", r"Microsoft\Edge\Application\msedge.exe"),
        ("ProgramFiles", r"Microsoft\Edge\Application\msedge.exe"),
        ("ProgramFiles", r"Google\Chrome\Application\chrome.exe"),
        ("ProgramFiles(x86)", r"Google\Chrome\Application\chrome.exe"),
    )
    for env, sub in candidates:
        base = os.environ.get(env)
        exe = Path(base) / sub if base else None
        if exe and exe.exists():
            subprocess.Popen([str(exe), f"--app={url}"], close_fds=True)
            return True
    try:
        os.startfile(str(page))          # plain browser tab as a last resort
        return True
    except OSError as e:
        print(f"Could not open {page}: {e}")
        return False


# --------------------------------------------------------------------------- #
# Watcher hook: regenerate only when something it shows could have changed     #
# --------------------------------------------------------------------------- #
_CACHE: Dict[str, Any] = {"key": None, "at": 0.0}
_MAX_AGE_SECONDS = 3600     # regenerate at least hourly regardless


def maybe_write(master: Dict[str, Any] | None,
                lq_jobs: List[Dict[str, Any]] | None,
                new_ids: set | None = None,
                force: bool = False) -> Optional[Path]:
    """Called by the watcher each poll. Cheap unless the line-items/DWG store
    files or today's change log changed on disk, the board membership changed,
    or an hour passed — then the page is rebuilt from the real stores. Returns
    the path written, or None when the page was already current."""
    import autocad_scan
    import change_log
    import line_items as li
    import solidworks_scan

    out = default_output_path()
    def _mtime(p: Path) -> float:
        try:
            return p.stat().st_mtime
        except OSError:
            return 0.0
    today = date.today()
    queue = {str(j.get("job")): j for j in lq_jobs or [] if j.get("job")}
    key = (_mtime(li.store_path()), _mtime(autocad_scan.PROGRESS_PATH),
           _mtime(change_log.log_path(today)), _mtime(solidworks_scan.PROGRESS_PATH),
           _mtime(Path(__file__)),      # a git pull regenerates on the next poll
           today.isoformat(),
           tuple(sorted(queue)), tuple(sorted(str(x) for x in new_ids or ())))
    now = time.time()
    if (not force and out.exists() and _CACHE["key"] == key
            and now - _CACHE["at"] < _MAX_AGE_SECONDS):
        return None

    payload = build_payload(li.load_store(), autocad_scan.load_progress(),
                            master_orders=(master or {}).get("orders"),
                            queue_jobs=queue, events=change_log.load(today),
                            today=today, new_ids=new_ids,
                            sw=solidworks_scan.load_progress())
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
    ap.add_argument("--changes", type=Path, default=None,
                    help="Day change-log JSON for the Changes tab (default: "
                         "today's log from the configured snapshots folder)")
    ap.add_argument("--sw", type=Path, default=None,
                    help="SolidWorks scan store JSON (default: the configured "
                         "solidworks_scan store)")
    ap.add_argument("--open", action="store_true",
                    help="Open the page in an app window (Edge/Chrome --app). "
                         "The page is rebuilt first when it's missing OR older "
                         "than this code (i.e. right after a git pull); "
                         "otherwise it opens as-is — the watcher keeps it fresh.")
    args = ap.parse_args(argv)

    if args.open:
        out = args.out or default_output_path()
        try:
            fresh = (out.exists()
                     and out.stat().st_mtime >= Path(__file__).stat().st_mtime)
        except OSError:
            fresh = False
        if fresh:
            print(f"Opening {out}")
            return 0 if open_in_app_window(out) else 1
        if out.exists():
            print(f"{out} predates the current code (git pull?) — rebuilding "
                  "it first, takes a moment…")
        else:
            print(f"{out} does not exist yet — building it first…")

    import change_log
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

    today = date.today()
    if args.changes:
        events = json.loads(args.changes.read_text(encoding="utf-8"))
        if not isinstance(events, list):
            events = []
    else:
        events = change_log.load(today)

    if args.sw:
        sw = json.loads(args.sw.read_text(encoding="utf-8"))
    else:
        import solidworks_scan
        sw = solidworks_scan.load_progress()

    payload = build_payload(store, dwg, master_orders=master.get("orders"),
                            events=events, today=today, sw=sw)
    out = write_explorer(payload, args.out)
    n_q = sum(1 for e in payload["jobs"].values() if e.get("q"))
    print(f"Wrote {out}  ({payload['n_jobs']} orders, {payload['n_items']} line "
          f"items, {n_q} in the queue)  + {BAT_NAME} + {VERSION_NAME}")
    if args.open:
        return 0 if open_in_app_window(out) else 1
    return 0


# --------------------------------------------------------------------------- #
# The page. One token: __B64__ (the gzip+base64 payload). No external          #
# resources of any kind — works from file:// on a shared drive, offline.       #
# The xf-*/chg*/mx-* classes carry the Excel workbook's exact fill RGBs        #
# (live_excel._FILL_RGB) in BOTH themes, with dark ink on the fills like       #
# Excel, so the color language transfers 1:1.                                  #
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
  input { font: inherit; }
  button:focus-visible, input:focus-visible, a:focus-visible {
    outline: 2px solid var(--accent); outline-offset: 2px; border-radius: 3px; }

  .wrap { max-width: 1360px; margin: 0 auto; padding: 0 20px 40px; }
  header.top { display: flex; flex-wrap: wrap; align-items: center; gap: 10px 18px;
    padding: 16px 0 12px; border-bottom: 2px solid var(--ink); margin-bottom: 16px; }
  .navback { font-size: 16px; font-weight: 700; line-height: 1; color: var(--muted);
    border: 1px solid var(--line); border-radius: 8px; padding: 4px 12px; }
  .navback:hover { color: var(--accent); border-color: var(--accent); }
  .navback[disabled] { opacity: .35; pointer-events: none; }
  .wordmark { font-family: var(--mono); font-size: 15px; letter-spacing: .14em;
              font-weight: 700; white-space: nowrap; }
  .wordmark .dim { color: var(--muted); font-weight: 400; }
  nav.tabs { display: flex; gap: 6px; }
  .tabbtn { font-size: 12.5px; font-weight: 600; padding: 5px 14px;
    border-radius: 999px; border: 1px solid var(--line); color: var(--muted); }
  .tabbtn:hover { border-color: var(--accent); color: var(--accent); }
  .tabbtn.active { background: var(--accent); color: var(--accent-ink);
    border-color: var(--accent); }
  .searchbox { position: relative; flex: 1 1 300px; max-width: 520px; margin-left: auto; }
  .searchbox input[type=search] { width: 100%; padding: 8px 12px 8px 34px;
    font-size: 13.5px; color: var(--ink); background: var(--panel);
    border: 1.5px solid var(--line); border-radius: 8px; }
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
  .sd-item .onq { margin-left: auto; }
  .sd-item .why { font-family: var(--mono); font-size: 11.5px; color: var(--muted);
    margin-top: 2px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .sd-item .why mark { background: var(--hit); color: inherit; padding: 0 1px; }

  .view { display: none; }
  .view.on { display: block; }
  .panel { background: var(--panel); border: 1px solid var(--line);
    border-radius: 12px; overflow: hidden; }
  .panel-head { display: flex; align-items: baseline; gap: 10px; flex-wrap: wrap;
    padding: 12px 16px; border-bottom: 1px solid var(--line); background: var(--panel-2); }
  .eyebrow { font-family: var(--mono); font-size: 10.5px; letter-spacing: .14em;
    color: var(--muted); text-transform: uppercase; }
  .panel-body { padding: 14px 16px 16px; }
  .empty { padding: 46px 24px; text-align: center; color: var(--muted); }
  .empty .big { font-size: 15px; color: var(--ink); margin-bottom: 6px; }

  /* ---- data tables (Board / Changes / History): Excel's color language ---- */
  .tablewrap { overflow: auto; max-height: 78vh; }
  table.gt { border-collapse: collapse; font-size: 12px; width: 100%; }
  .gt th { position: sticky; top: 0; z-index: 2; background: #305496; color: #FFFFFF;
    font-weight: 700; text-align: left; padding: 5px 8px; white-space: nowrap; }
  .gt th.sth { cursor: pointer; user-select: none; }
  .gt th.sth:hover { background: #3E64A8; }
  .gt th.vh { writing-mode: vertical-rl; transform: rotate(180deg);
    text-align: left; vertical-align: bottom; font-weight: 600; font-size: 11px;
    padding: 6px 2px; max-height: 150px; }
  .gt td { padding: 3px 8px; border-bottom: 1px solid var(--line);
    white-space: nowrap; font-variant-numeric: tabular-nums; }
  .gt td.wrapcell { white-space: normal; min-width: 260px; }
  .gt tr.rowbtn:hover td { outline: 1px solid var(--accent); outline-offset: -1px; }
  .gt .jobcell { font-family: var(--mono); font-weight: 700; cursor: pointer; }
  .gt .jobcell:hover { color: var(--accent); text-decoration: underline; }
  .gt td.num { text-align: right; }
  .gt td.ctr { text-align: center; }
  .totrow td { font-weight: 700; border-top: 2px solid var(--ink); border-bottom: none;
    padding-top: 6px; }

  /* Excel fills, exact RGBs from live_excel._FILL_RGB; dark ink like Excel. */
  .xf-ov td, .xf-dt td, .xf-sn td, .xf-nw td,
  .xf-ovn td, .xf-dtn td, .xf-snn td,
  .chg1 td, .chg2 td, .chg3 td, .chg4 td { color: #1B242D; }
  .xf-ov td { background: #FFC7CE; }    /* overdue */
  .xf-dt td { background: #F8CBAD; }    /* due today */
  .xf-sn td { background: #FFEB9C; }    /* due within 3 days */
  .xf-nw td { background: #D9D9D9; }    /* new today */
  .xf-ovn td { background: #F4A5A8; }   /* overdue + new */
  .xf-dtn td { background: #F4B183; }   /* due today + new */
  .xf-snn td { background: #F5D750; }   /* soon + new */
  .chg1 td { background: #D9D9D9; }     /* change instances, darker each time */
  .chg2 td { background: #BFBFBF; }
  .chg3 td { background: #A6A6A6; }
  .chg4 td { background: #8C8C8C; }
  td.mx-y { background: #C6EFCE; color: #1B242D; text-align: center; }  /* has it */
  td.mx-n { background: #FFC7CE; }                                      /* doesn't */
  td.mx-sep { background: #808080; padding: 0 2px; }
  tr.co-red td { color: #C00000; font-weight: 600; }   /* a CO landed today */
  .drun { color: #C55A11; font-weight: 700; }          /* highly-custom quote run */
  .flag-yes { color: #1B242D; background: #C6EFCE; border-radius: 4px;
    padding: 0 6px; font-weight: 700; font-size: 11px; }

  .secttitle { font-weight: 700; font-size: 13px; padding: 14px 16px 6px; }
  .secttitle .cnt { color: var(--muted); font-weight: 400; }
  .sectnote { padding: 0 16px 8px; font-size: 12px; color: var(--muted); }
  .histbar { display: flex; gap: 10px; align-items: center; padding: 10px 16px;
    border-bottom: 1px solid var(--line); background: var(--panel-2); flex-wrap: wrap; }
  .histbar input { padding: 6px 10px; border: 1.5px solid var(--line);
    border-radius: 7px; background: var(--panel); color: var(--ink); width: 300px;
    max-width: 100%; font-size: 13px; }
  .histbar .note { font-size: 11.5px; color: var(--muted); }
  .morebtn { display: block; width: 100%; text-align: center; padding: 9px;
    font-weight: 600; color: var(--accent); border-top: 1px solid var(--line); }
  .morebtn:hover { background: var(--accent-soft); }

  #toast { position: fixed; right: 18px; bottom: 18px; z-index: 90;
    background: var(--ink); color: var(--bg); font-size: 12.5px; font-weight: 600;
    padding: 9px 16px; border-radius: 9px; box-shadow: 0 8px 24px rgba(0,0,0,.25);
    opacity: 0; pointer-events: none; transition: opacity .3s; }
  #toast.show { opacity: 1; }
  @media (prefers-reduced-motion: reduce) { #toast { transition: none; } }

  /* The Order view is a grid ONLY while active — a bare `main.cols`
     display rule would outrank `.view`'s hide and bleed under other tabs. */
  main.cols { grid-template-columns: minmax(0, 11fr) minmax(0, 9fr);
    gap: 18px; align-items: start; }
  main.cols.view.on { display: grid; }
  @media (max-width: 920px) { main.cols { grid-template-columns: 1fr; } }

  .backlink { font-size: 12.5px; color: var(--accent); font-weight: 600; }
  .ohead { display: flex; align-items: baseline; gap: 10px; flex-wrap: wrap; flex: 1; }
  .ohead .job { font-family: var(--mono); font-size: 19px; font-weight: 700; }
  .ohead .cust { font-size: 13px; color: var(--muted); }
  .ohead .co { font-family: var(--mono); font-size: 11.5px; color: var(--muted);
    border: 1px solid var(--line); border-radius: 4px; padding: 1px 6px; }
  /* IN QUEUE badge: filled with the job's own row color from the Live Queue
     (overdue red, due-today orange, soon gold, new grey) — plain green outline
     when its row is unfilled. */
  .onq { font-size: 10px; font-family: var(--mono); letter-spacing: .06em;
    color: var(--good); border: 1px solid var(--good); border-radius: 999px;
    padding: 1px 7px; white-space: nowrap; align-self: center; }
  .onq.qbf { color: #1B242D; border-color: rgba(27, 36, 45, .25); }
  .qb-xf-ov { background: #FFC7CE; }   .qb-xf-dt { background: #F8CBAD; }
  .qb-xf-sn { background: #FFEB9C; }   .qb-xf-nw { background: #D9D9D9; }
  .qb-xf-ovn { background: #F4A5A8; }  .qb-xf-dtn { background: #F4B183; }
  .qb-xf-snn { background: #F5D750; }
  .metaline { display: flex; flex-wrap: wrap; gap: 6px 16px; margin-top: 10px;
    font-size: 12px; align-items: baseline; }
  .metaline .dwg { font-family: var(--mono); color: var(--good); font-weight: 600; }
  .metaline .path { font-family: var(--mono); color: var(--muted); }

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
  .m-spec { margin: 1px 0 0 32px; font-size: 12px; color: var(--muted); }
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
    <button class="navback" id="navback" title="Back (Alt+&#8592;)">&#8592;</button>
    <span class="wordmark">GL QUEUE <span class="dim">/</span> EXPLORER</span>
    <nav class="tabs" id="tabs" style="display:none">
      <button class="tabbtn" data-tab="changes">Changes</button>
      <button class="tabbtn" data-tab="board">Live Queue</button>
      <button class="tabbtn" data-tab="hist">Order History</button>
      <button class="tabbtn" data-tab="job" id="jobtab">Order</button>
    </nav>
    <div class="searchbox">
      <span class="glass">&#8981;</span>
      <input id="q" type="search" autocomplete="off" spellcheck="false"
             placeholder="Job # (or last digits) &mdash; or a feature: teflon, low leak&hellip;"
             aria-label="Search jobs or features">
      <div class="search-drop" id="drop"></div>
    </div>
  </header>
  <div id="boot">Loading the order data&hellip;</div>
  <section class="view panel" id="view-board"></section>
  <section class="view panel" id="view-changes"></section>
  <section class="view panel" id="view-hist"></section>
  <main class="cols view" id="view-job">
    <section class="panel" id="left"></section>
    <section class="panel" id="right"></section>
  </main>
  <footer class="note" style="display:none">
    <span>One file, no install, no internet &mdash; the watcher regenerates it and open
      pages refresh themselves. Scores: rare shared features count highest,
      identical Sales-Order lines double.</span>
    <span class="mono" id="stamp"></span>
  </footer>
  <div id="toast" role="status"></div>
</div>

<script>
"use strict";
const PAYLOAD_B64 = "__B64__";
const VERSION_FILE = "gl_queue_explorer_version.js";
const STATE_KEY = "glq_state";

let DB = null;                 // {gen, today, n_jobs, n_items, jobs, ev, rm, nw?}
let IDX = null;                // {tagDF, normDF, sets: {job -> {t:Set, n:Set}}}
let COSET = new Set();         // jobs a CO# landed on today (red text)
let NWSET = null;              // the watcher's exact new-today set (null = derive)
const state = { tab: "board", job: null, path: null, whole: false,
                pinned: new Set(), histQ: "", histN: 500,
                boardQ: "", boardSort: null, histSort: null, only3d: false };

const $ = id => document.getElementById(id);
const esc = s => String(s).replace(/[&<>"']/g, c =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const money = n => (+n).toLocaleString("en-US",
  { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const jobNum = j => /^\d+/.test(j) ? parseInt(j, 10) : -1;
const fileUrl = p => "file:///" + encodeURI(String(p).replace(/\\/g, "/"))
  .replace(/#/g, "%23");
/* Folder links use the glqueue: protocol (registered by the launcher .bat /
   Open button) so they open in Windows File Explorer, not a browser listing. */
const folderUrl = p => "glqueue:" + encodeURI(String(p).replace(/\\/g, "/"))
  .replace(/#/g, "%23");
const spv = (e, label) => { const kv = (e.sp || []).find(x => x[0] === label);
  return kv ? kv[1] : ""; };

/* item row indices: [no, text, price, qty, section, norm, tags] */
const IT = { NO: 0, RAW: 1, PRICE: 2, QTY: 3, SECTION: 4, NORM: 5, TAGS: 6 };

function pDate(s) { if (!s) return null; const d = new Date(s);
  return isNaN(d) ? null : d; }
function dayFloor(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function sameDay(a, b) { return a && b && dayFloor(a).getTime() === dayFloor(b).getTime(); }
function fmtWhen(iso) {
  const d = pDate(iso); if (!d) return "";
  if (sameDay(d, new Date()))
    return d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  return d.toLocaleDateString([], { month: "short", day: "numeric", year: "numeric" });
}
/* The workbook's row-fill rules (live_sheets._row_fill), against the REAL
   current date so a page left open overnight keeps telling the truth. The
   new-today set is the watcher's own (snapshot-diff based) when embedded —
   identical to the workbook — else derived from the arrival timestamp. */
function isNewToday(j, bd) {
  if (NWSET) return NWSET.has(j);
  return sameDay(pDate(bd.ai), new Date());
}
function fillClass(j, e) {
  const bd = e.bd || {};
  const isNew = isNewToday(j, bd);
  const ed = pDate(bd.ed);
  if (ed) {
    const t0 = dayFloor(new Date());
    const soon = new Date(t0.getFullYear(), t0.getMonth(), t0.getDate() + 3);
    if (ed < t0) return isNew ? "xf-ovn" : "xf-ov";
    if (ed.getTime() === t0.getTime()) return isNew ? "xf-dtn" : "xf-dt";
    if (ed <= soon) return isNew ? "xf-snn" : "xf-sn";
  }
  return isNew ? "xf-nw" : "";
}
/* The IN QUEUE badge, filled with the job's current row color on the queue. */
function queueBadge(j, e) {
  if (!e.q) return "";
  const f = fillClass(j, e);
  return '<span class="onq' + (f ? " qbf qb-" + f : "") + '" title="'
    + (f ? "Filled with its row color on the queue" : "Currently in the queue")
    + '">IN QUEUE</span>';
}
/* One compact fan line: D/design · S/size · A/arr · width% · rot-disch · wheel */
function fanSpec(e) {
  const v = label => {
    const t = spv(e, label);
    return t && t.toUpperCase() !== "N/A" ? t : "";
  };
  const parts = [];
  if (v("Design")) parts.push("D/" + v("Design"));
  if (v("Size")) parts.push("S/" + v("Size"));
  const a = v("Arrangement");
  if (a) parts.push(a.startsWith("A/") ? a : "A/" + a);
  const w = v("% Width");
  if (w) parts.push(w.endsWith("%") ? w : w + "%");
  const rd = [v("Rotation"), v("Discharge")].filter(Boolean).join("-");
  if (rd) parts.push(rd);
  if (v("Wheel")) parts.push(v("Wheel"));
  return parts.join(" · ");
}

async function boot() {
  const bytes = Uint8Array.from(atob(PAYLOAD_B64), c => c.charCodeAt(0));
  const stream = new Blob([bytes]).stream()
    .pipeThrough(new DecompressionStream("gzip"));
  DB = JSON.parse(await new Response(stream).text());
  COSET = new Set(DB.ev.filter(x => x.f === "CO#").map(x => x.j));
  NWSET = Array.isArray(DB.nw) ? new Set(DB.nw) : null;
  $("boot").style.display = "none";
  $("tabs").style.display = "";
  document.querySelector("footer").style.display = "";
  $("stamp").textContent = "generated " + DB.gen + " · " + DB.n_jobs
    + " orders · " + DB.n_items + " line items";
  try {   // hover the stamp to see WHICH copy of the page this is
    $("stamp").title = decodeURIComponent(location.pathname.replace(/^\//, ""))
      .replace(/\//g, "\\");
  } catch (e) {}
  loadPrefs();
  restoreState();
  render();                       // also seeds the first history entry (syncNav)
  $("navback").onclick = () => history.back();
  updateBackBtn();
  setTimeout(ensureIndex, 50);        // warm the match index off the first paint
  setInterval(pollVersion, 60000);
}

/* ---- auto-refresh: poll the sibling stamp; reload when the watcher wrote --- */
function pollVersion() {
  const s = document.createElement("script");
  s.src = VERSION_FILE + "?t=" + Date.now();   // query defeats any file cache
  s.onload = () => { s.remove();
    if (window.__GLQ_VERSION__ && DB && window.__GLQ_VERSION__ !== DB.gen)
      reloadFresh(); };
  s.onerror = () => s.remove();                // stamp missing -> feature off
  document.head.appendChild(s);
}
function reloadFresh() {
  try {
    sessionStorage.setItem(STATE_KEY, JSON.stringify({
      tab: state.tab, job: state.job, path: state.path, whole: state.whole,
      pinned: [...state.pinned], histQ: state.histQ, y: window.scrollY,
      refreshed: 1 }));
  } catch (e) {}
  location.reload();
}
function restoreState() {
  let saved = null;
  try { saved = JSON.parse(sessionStorage.getItem(STATE_KEY) || "null");
        sessionStorage.removeItem(STATE_KEY); } catch (e) {}
  if (!saved) return;
  if (saved.job && DB.jobs[saved.job]) {
    state.job = saved.job; state.path = saved.path; state.whole = !!saved.whole;
    state.pinned = new Set(saved.pinned || []);
  }
  state.tab = ["board", "changes", "hist", "job"].includes(saved.tab)
    ? saved.tab : "board";
  if (state.tab === "job" && !state.job) state.tab = "board";
  state.histQ = saved.histQ || "";
  if (saved.y) setTimeout(() => window.scrollTo(0, saved.y), 60);
  if (saved.refreshed) toast("Refreshed — new data as of " + DB.gen);
}
/* Per-user view preferences: each coworker's browser profile keeps their own
   sort/filter/tab in localStorage, so watcher refreshes and reopens never
   reset anyone's view. The shared file itself stores nothing per user. */
const PREFS_KEY = "glq_prefs";
function savePrefs() {
  try {
    localStorage.setItem(PREFS_KEY, JSON.stringify({
      tab: state.tab === "job" ? "board" : state.tab,
      boardSort: state.boardSort, histSort: state.histSort,
      boardQ: state.boardQ, histQ: state.histQ, only3d: state.only3d }));
  } catch (e) {}
}
function loadPrefs() {
  try {
    const p = JSON.parse(localStorage.getItem(PREFS_KEY) || "null");
    if (!p) return;
    if (["board", "changes", "hist"].includes(p.tab)) state.tab = p.tab;
    if (p.boardSort && typeof p.boardSort.col === "number") state.boardSort = p.boardSort;
    if (p.histSort && typeof p.histSort.col === "number") state.histSort = p.histSort;
    state.boardQ = p.boardQ || "";
    state.histQ = p.histQ || "";
    state.only3d = !!p.only3d;
  } catch (e) {}
}
let toastTimer = null;
function toast(msg) {
  const t = $("toast");
  t.textContent = msg; t.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove("show"), 4000);
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

/* ---- matching: rarity-weighted overlap (find_orders.similar_to_items),
   plus a fan-design bonus — an order of the SAME design that shares content
   outranks an equally-shared order of a different design. ------------------- */
const DESIGN_BONUS = 1.5;
function rankMatches(srcJob, items) {
  const { tagDF, normDF, sets } = ensureIndex();
  const srcDesign = spv(DB.jobs[srcJob], "Design");
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
    for (const t of tTags) if (s.t.has(t)) score += 1 / tagDF[t];
    if (!score) continue;
    if (srcDesign && spv(DB.jobs[j], "Design") === srcDesign) score += DESIGN_BONUS;
    if (state.pinned.size) {
      let ok = true;
      for (const p of state.pinned) {
        const ix = p.indexOf("=");
        if (!anyAttrMatch(DB.jobs[j], p.slice(0, ix), p.slice(ix + 1))) { ok = false; break; }
      }
      if (!ok) continue;
    }
    out.push({ j, score, sharedNorms: new Set(sharedNorms), tTags });
  }
  out.sort((a, b) => b.score - a.score || jobNum(b.j) - jobNum(a.j));
  return out;
}

/* ------------------------------- rendering --------------------------------- */
function render() {
  for (const t of ["board", "changes", "hist", "job"])
    $("view-" + t).classList.toggle("on", state.tab === t);
  document.querySelectorAll(".tabbtn").forEach(b =>
    b.classList.toggle("active", b.dataset.tab === state.tab));
  $("jobtab").textContent = state.job ? "Order · " + state.job : "Order";
  if (state.tab === "board") renderBoard();
  else if (state.tab === "changes") renderChanges();
  else if (state.tab === "hist") renderHist();
  else { renderJobPane(); renderMatches(); }
  syncNav();          // refinements (component/pins) update the current entry
}

/* ---- real browser history: every tab switch and order open is an entry, so
   the header's ← arrow, Alt+Left, and the mouse back button all walk the
   same chain; in-view refinements update the current entry in place. ------ */
let NAV_POS = 0, POPPING = false;
function navSnapshot() {
  return { tab: state.tab, job: state.job, path: state.path,
           whole: state.whole, pinned: [...state.pinned], i: NAV_POS };
}
function pushNav() {
  NAV_POS++;
  try { history.pushState(navSnapshot(), ""); } catch (e) {}
  updateBackBtn();
}
function syncNav() {
  if (POPPING) return;
  try { history.replaceState(navSnapshot(), ""); } catch (e) {}
}
function updateBackBtn() {
  $("navback").disabled = NAV_POS <= 0 && history.length <= 1;
}
window.addEventListener("popstate", ev => {
  const s = ev.state;
  if (!s || !DB) return;
  POPPING = true;
  NAV_POS = s.i || 0;
  state.tab = ["board", "changes", "hist", "job"].includes(s.tab) ? s.tab : "board";
  state.job = s.job && DB.jobs[s.job] ? s.job : null;
  if (state.tab === "job" && !state.job) state.tab = "board";
  state.path = s.path === undefined ? null : s.path;
  state.whole = !!s.whole;
  state.pinned = new Set(s.pinned || []);
  render();
  POPPING = false;
  updateBackBtn();
});

function setTab(t) {
  if (t !== state.tab) { state.tab = t; pushNav(); }
  savePrefs(); render(); window.scrollTo(0, 0);
}
function selectJob(j) {
  /* Opening an order lands on its whole-order matches (the Similar Orders
     view) — a component/attribute click narrows from there. */
  if (state.tab === "job" && state.job === j) { render(); return; }
  state.job = j; state.path = null; state.whole = true; state.pinned.clear();
  state.tab = "job";
  pushNav();
  savePrefs(); render(); window.scrollTo(0, 0);
}
function wireJobCells(root) {
  root.querySelectorAll("[data-job]").forEach(el =>
    el.onclick = () => selectJob(el.dataset.job));
}

/* Folder/file access is plain links now (AutoCAD folder, SolidWorks 3D, SO
   PDF, Quote Run) — a directory link opens the browser's file listing, and
   every link's hover title still shows the full Z: path. */

/* ------------------------------ Live Queue --------------------------------- */
const BOARD_COLS = ["Added", "Job #", "Run", "CO#", "Oper", "Design", "Customer",
  "Size", "Arr.", "Assigned To", "Checker", "Note", "Engineer", "End Date",
  "Start Date", "FanNet", "Price", "DWG Reuse", "#"];

/* Sort value: currency/plain numbers numerically, dates chronologically,
   everything else as text; blanks always sink to the bottom. */
function numish(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return v;
  const s = String(v).replace(/[$,]/g, "");
  if (/^-?\d+(\.\d+)?$/.test(s)) return +s;
  const d = new Date(v);
  return isNaN(d) ? null : d.getTime();
}
function sortRows(rows, sort) {
  const { col, dir } = sort;
  return rows.slice().sort((a, b) => {
    const va = a.c[col].v, vb = b.c[col].v;
    const ea = va === "" || va === null, eb = vb === "" || vb === null;
    if (ea || eb) return ea && eb ? 0 : ea ? 1 : -1;
    const na = numish(va), nb = numish(vb);
    const r = (na !== null && nb !== null) ? na - nb
      : String(va).localeCompare(String(vb));
    return dir * r;
  });
}
function sortableHead(cols, sort) {
  return cols.map((h, i) => '<th class="sth" data-si="' + i
    + '" title="Click to sort">' + esc(h)
    + (sort && sort.col === i ? (sort.dir < 0 ? " ▼" : " ▲") : "")
    + "</th>").join("");
}
function wireSort(el, key) {
  el.querySelectorAll("th.sth").forEach(th => th.onclick = () => {
    const col = +th.dataset.si, cur = state[key];
    state[key] = cur && cur.col === col ? { col, dir: -cur.dir } : { col, dir: 1 };
    savePrefs();
    render();
  });
}
function wireFilter(inp, key, rerender) {
  inp.oninput = () => { state[key] = inp.value;
    clearTimeout(inp._t); inp._t = setTimeout(() => {
      savePrefs();
      const pos = inp.selectionStart; rerender();
      const ni = $(inp.id); ni.focus(); ni.setSelectionRange(pos, pos);
    }, 250); };
}

function renderBoard() {
  const el = $("view-board");
  // Base order: job # ascending, so when the "#" board position is missing
  // (snapshot builds) the stable sort still yields a sensible fixed order.
  const onq = Object.keys(DB.jobs).filter(j => DB.jobs[j].q)
    .sort((a, b) => jobNum(a) - jobNum(b));
  // The "#" board position exists only on watcher-generated pages; hide the
  // column entirely (instead of showing blanks) when this build lacks it.
  const hasPos = onq.some(j => ((DB.jobs[j].bd || {}).ps || 0) > 0);
  const cols = hasPos ? BOARD_COLS : BOARD_COLS.slice(0, -1);
  let sort = state.boardSort;
  if (sort && sort.col >= cols.length) sort = null;
  sort = sort || (hasPos ? { col: 18, dir: 1 }          // cbcinsider board order
                         : { col: 1, dir: 1 });         // else job # order
  let rows = onq.map(j => {
    const e = DB.jobs[j], bd = e.bd || {};
    const t = s => ({ v: s || "", h: esc(s || "") });
    const c = [
      { v: bd.ai || "", h: esc(fmtWhen(bd.ai)) },
      { v: jobNum(j), h: '<span class="jobcell" data-job="' + esc(j) + '">'
          + esc(j) + "</span>" },
      { v: bd.dr ? 1 : "", h: !bd.dr ? "" : e.qr
          ? '<a class="drun" href="' + esc(fileUrl(e.qr)) + '" target="_blank" title="'
            + esc(e.qr) + '">Run</a>'
          : '<span class="drun">Run</span>' },
      t(e.co), t(bd.op), t(spv(e, "Design")), t(e.c), t(spv(e, "Size")),
      t(spv(e, "Arrangement")), t(bd.as), t(bd.ck), t(bd.no), t(e.e),
      t(bd.ed), t(bd.sd), t(bd.fn), t(bd.pr), t(bd.ru),
      { v: bd.ps || "", h: String(bd.ps || "") },
    ];
    if (!hasPos) c.pop();
    return { j, e, bd, c };
  });
  const needle = state.boardQ.trim().toUpperCase();
  if (needle)
    rows = rows.filter(r => [r.j, r.e.c, spv(r.e, "Design"), r.bd.no, r.bd.as,
      r.bd.ck, r.e.e, r.e.co, r.bd.st].some(x =>
        String(x || "").toUpperCase().includes(needle)));
  rows = sortRows(rows, sort);
  let total = 0;
  const body = rows.map(r => {
    total += +String(r.bd.pr || "").replace(/[^0-9.\-]/g, "") || 0;
    const cls = [fillClass(r.j, r.e), COSET.has(r.j) ? "co-red" : "", "rowbtn"]
      .filter(Boolean).join(" ");
    return '<tr class="' + cls + '">' + r.c.map((x, i) =>
      "<td" + (i === 16 ? ' class="num"' : i === 18 ? ' class="ctr"' : "")
      + ">" + x.h + "</td>").join("") + "</tr>";
  }).join("");
  el.innerHTML = '<div class="panel-head"><span class="eyebrow">Live Queue</span>'
    + '<span class="m-count">' + rows.length + " of " + onq.length
    + " orders · $" + money(total) + " shown · click a column to sort · "
    + "colors match the workbook"
    + (hasPos ? "" : " · queue # appears on watcher-generated pages")
    + "</span></div>"
    + '<div class="histbar"><input id="boardq" type="text" placeholder="Filter: '
    + 'job #, customer, design, engineer, note&hellip;" value="'
    + esc(state.boardQ) + '"></div>'
    + '<div class="tablewrap"><table class="gt"><thead><tr>'
    + sortableHead(cols, sort)
    + "</tr></thead><tbody>" + body
    + '<tr class="totrow"><td colspan="6">Total jobs: ' + rows.length + "</td>"
    + '<td colspan="10" style="text-align:right">Total $ in process:</td>'
    + '<td class="num">' + money(total) + "</td>"
    + "<td></td>".repeat(cols.length - 17) + "</tr>"
    + "</tbody></table></div>";
  wireJobCells(el);
  wireSort(el, "boardSort");
  wireFilter($("boardq"), "boardQ", renderBoard);
}

/* ------------------------------- Changes ----------------------------------- */
function miniTable(headers, rows) {
  if (!rows.length) return '<div class="sectnote">(none)</div>';
  return '<div class="tablewrap" style="max-height:none"><table class="gt"><thead><tr>'
    + headers.map(h => "<th>" + esc(h) + "</th>").join("")
    + "</tr></thead><tbody>" + rows.join("") + "</tbody></table></div>";
}
function renderChanges() {
  const el = $("view-changes");
  const newToday = Object.keys(DB.jobs).filter(j => {
    const e = DB.jobs[j];
    return e.q && (NWSET ? NWSET.has(j)
      : String((e.bd || {}).ai || "").slice(0, 10) === DB.today);
  }).sort((a, b) => jobNum(b) - jobNum(a));
  const newRows = newToday.map(j => {
    const e = DB.jobs[j], bd = e.bd || {};
    return '<tr class="rowbtn"><td>' + esc(fmtWhen(bd.ai)) + "</td>"
      + '<td class="jobcell" data-job="' + esc(j) + '">' + esc(j) + "</td>"
      + "<td>" + esc(e.co || "") + "</td>"
      + "<td>" + esc(spv(e, "Design")) + "</td>"
      + "<td>" + esc(e.c) + "</td>"
      + "<td>" + esc(bd.no || "") + "</td>"
      + "<td>" + esc(bd.ed || "") + "</td>"
      + '<td class="num">' + esc(bd.pr || "") + "</td></tr>";
  });

  const coEv = DB.ev.filter(x => x.f === "CO#");
  const coRows = coEv.map(x => {
    const e = DB.jobs[x.j] || {};
    return '<tr class="rowbtn"><td>' + esc(fmtWhen(x.t)) + "</td>"
      + '<td class="jobcell" data-job="' + esc(x.j) + '">' + esc(x.j) + "</td>"
      + '<td style="color:#C00000;font-weight:700">CO#' + esc(x.n) + "</td>"
      + "<td>" + esc(spv(e, "Design")) + "</td>"
      + "<td>" + esc(x.c || e.c || "") + "</td>"
      + '<td class="wrapcell">' + esc(x.d || "") + "</td></tr>";
  });

  const fieldEv = DB.ev.filter(x => x.f !== "CO#");
  const byJob = new Map();
  for (const x of fieldEv) {
    if (!byJob.has(x.j)) byJob.set(x.j, []);
    byJob.get(x.j).push(x);
  }
  const chRows = [];
  for (const [j, evs] of byJob) {
    const e = DB.jobs[j] || {};
    chRows.push('<tr><td colspan="5" style="font-weight:700">'
      + '<span class="jobcell" data-job="' + esc(j) + '">' + esc(j) + "</span>"
      + '&nbsp; <span style="font-weight:400;color:var(--muted)">'
      + esc(evs[0].c || e.c || "") + "</span></td></tr>");
    evs.forEach((x, i) => {
      const shade = "chg" + Math.min(i + 1, 4);   // darker per later instance
      chRows.push('<tr class="' + shade + '"><td>' + esc(fmtWhen(x.t)) + "</td>"
        + "<td>" + esc(x.f) + "</td>"
        + "<td>" + esc(x.o) + "</td><td>&rarr;</td>"
        + "<td style='font-weight:600'>" + esc(x.n) + "</td></tr>");
    });
  }

  const rmRows = DB.rm.map(r => {
    const e = DB.jobs[r[0]] || {};
    return '<tr class="rowbtn"><td>' + esc(fmtWhen(r[1])) + "</td>"
      + '<td class="jobcell" data-job="' + esc(r[0]) + '">' + esc(r[0]) + "</td>"
      + "<td>" + esc(e.co || "") + "</td>"
      + "<td>" + esc(spv(e, "Design")) + "</td>"
      + "<td>" + esc(e.c || "") + "</td></tr>";
  });

  const allEmpty = !newRows.length && !coRows.length && !byJob.size && !rmRows.length;
  el.innerHTML = '<div class="panel-head"><span class="eyebrow">Changes — '
    + esc(DB.today) + '</span><span class="m-count">generated ' + esc(DB.gen)
    + "</span></div>"
    + (allEmpty ? '<div class="sectnote" style="padding-top:12px">Nothing logged '
        + "for " + esc(DB.today) + " yet. This tab fills from the watcher PC&#39;s "
        + "day log as orders arrive, change, and leave — a page built from a "
        + "data snapshot (like a download of this file) starts empty.</div>" : "")
    + '<div class="secttitle">New orders today <span class="cnt">(' + newRows.length
    + ")</span></div>"
    + miniTable(["Time", "Job #", "CO#", "Design", "Customer", "Note", "End Date",
                 "Price"], newRows)
    + '<div class="secttitle">Change orders today <span class="cnt">(' + coRows.length
    + ")</span></div>"
    + miniTable(["Time", "Job #", "CO#", "Design", "Customer", "What changed"], coRows)
    + '<div class="secttitle">Orders that changed today <span class="cnt">('
    + byJob.size + ")</span></div>"
    + (chRows.length
        ? miniTable(["Time", "Field", "Old", "", "New"], chRows)
        : '<div class="sectnote">(none)</div>')
    + '<div class="secttitle">Removed / completed today <span class="cnt">('
    + rmRows.length + ")</span></div>"
    + miniTable(["Time", "Job #", "CO#", "Design", "Customer"], rmRows)
    + '<div style="height:10px"></div>';
  wireJobCells(el);
}

/* ---------------------------- Order History -------------------------------- */
const MATRIX_LIMIT = 300;   // full green-check/red matrices once filtered this narrow
const HIST_COLS = ["Job #", "CO#", "Design", "Description", "Size", "Arr.",
  "Motor Pos", "Class", "Rot.", "Disch.", "% W", "Wheel", "D.Temp", "M.Temp",
  "Customer", "Engineer", "Item", "On Queue", "Added", "Left"];

function histRowsAll() {
  return Object.keys(DB.jobs).filter(j => DB.jobs[j].oh)
    .sort((a, b) => jobNum(b) - jobNum(a));
}
function histFilter(rows, q) {
  if (!q) return rows;
  const n = q.toUpperCase();
  return rows.filter(j => {
    const e = DB.jobs[j];
    return j.includes(n)
      || (e.c || "").toUpperCase().includes(n)
      || spv(e, "Design").toUpperCase().includes(n)
      || (e.e || "").toUpperCase().includes(n)
      || (e.im || "").toUpperCase().includes(n);
  });
}
function jobTags(e) {
  const s = new Set();
  for (const row of e.it) for (const t of row[IT.TAGS]) s.add(t);
  return s;
}
function renderHist() {
  const el = $("view-hist");
  const all = histRowsAll();
  const filtered = histFilter(all, state.histQ.trim());
  const matrix = filtered.length <= MATRIX_LIMIT && filtered.length > 0;
  const sort = state.histSort || { col: 0, dir: -1 };   // newest job first

  let rows = filtered.map(j => {
    const e = DB.jobs[j], oh = e.oh;
    const t = s => ({ v: s || "", h: esc(s || "") });
    const c = [
      { v: jobNum(j), h: '<span class="jobcell" data-job="' + esc(j) + '">'
          + esc(j) + "</span>" },
      t(e.co), t(spv(e, "Design")), t(spv(e, "Description")), t(spv(e, "Size")),
      t(spv(e, "Arrangement")), t(spv(e, "Motor Pos")), t(spv(e, "Class")),
      t(spv(e, "Rotation")), t(spv(e, "Discharge")), t(spv(e, "% Width")),
      t(spv(e, "Wheel")), t(spv(e, "Design Temp")), t(spv(e, "Max Temp")),
      t(e.c), t(e.e), t(e.im),
      { v: oh[0], h: oh[0] ? '<span class="flag-yes">YES</span>' : "NO" },
      { v: oh[1] || "", h: esc(fmtWhen(oh[1])) },
      { v: oh[2] || "", h: esc(fmtWhen(oh[2])) },
    ];
    return { j, e, c };
  });
  rows = sortRows(rows, sort);
  const shown = rows.slice(0, state.histN);

  let sufCols = [], tagCols = [];
  if (matrix) {
    const suf = new Set(), tgs = new Set();
    for (const r of rows) {
      for (const s of r.e.dx || []) suf.add(s);
      for (const t of jobTags(r.e)) tgs.add(t);
    }
    sufCols = [...suf].sort((a, b) => (parseInt(a, 10) || 9e9) - (parseInt(b, 10) || 9e9)
                                      || String(a).localeCompare(String(b)));
    tagCols = [...tgs].sort();
  }

  const head = sortableHead(HIST_COLS, sort)
    + (matrix ? sufCols.map(s => '<th class="vh">-' + esc(s) + "</th>").join("")
        + (sufCols.length || tagCols.length ? '<th class="vh">&#9474;</th>' : "")
        + tagCols.map(t => '<th class="vh">' + esc(t) + "</th>").join("") : "");

  const body = shown.map(r => {
    let tds = r.c.map((x, i) => "<td" + (i === 17 ? ' class="ctr"' : "") + ">"
      + x.h + "</td>").join("");
    if (matrix) {
      const dx = new Set(r.e.dx || []), tg = jobTags(r.e);
      tds += sufCols.map(s => dx.has(s) ? '<td class="mx-y">✓</td>'
                                        : '<td class="mx-n"></td>').join("")
        + (sufCols.length || tagCols.length ? '<td class="mx-sep"></td>' : "")
        + tagCols.map(t => tg.has(t) ? '<td class="mx-y">✓</td>'
                                     : '<td class="mx-n"></td>').join("");
    }
    return '<tr class="rowbtn">' + tds + "</tr>";
  }).join("");

  el.innerHTML = '<div class="panel-head"><span class="eyebrow">Order History</span>'
    + '<span class="m-count">' + rows.length + " of " + all.length
    + " orders · click a column to sort"
    + (matrix ? " · showing the ✓/✗ DWG + feature matrices"
              : " · filter to ≤ " + MATRIX_LIMIT + " to unfold the ✓/✗ matrices")
    + "</span></div>"
    + '<div class="histbar"><input id="histq" type="text" placeholder="Filter: job #, '
    + 'customer, design, engineer, item&hellip;" value="' + esc(state.histQ) + '">'
    + '<span class="note">green ✓ = has that custom DWG / feature, red = doesn&#39;t '
    + "&mdash; same as the workbook</span></div>"
    + '<div class="tablewrap"><table class="gt"><thead><tr>' + head
    + "</tr></thead><tbody>" + body + "</tbody></table></div>"
    + (rows.length > shown.length
        ? '<button class="morebtn" id="histmore">Show '
          + Math.min(500, rows.length - shown.length) + " more (of "
          + (rows.length - shown.length) + " remaining)</button>" : "");
  wireJobCells(el);
  wireSort(el, "histSort");
  wireFilter($("histq"), "histQ", () => { state.histN = 500; renderHist(); });
  const more = $("histmore");
  if (more) more.onclick = () => { state.histN += 500; renderHist(); };
}

/* ------------------------------ Job view ----------------------------------- */
function renderJobPane() {
  const el = $("left");
  if (!state.job) {
    el.innerHTML = '<div class="panel-head"><span class="eyebrow">Order</span></div>'
      + '<div class="empty"><div class="big">No order open yet</div>'
      + "Click any job # on the Changes, Live Queue, or Order History tab — "
      + "or search one above — and it opens here.</div>";
    return;
  }
  const j = state.job, e = DB.jobs[j];
  const specs = (e.sp || []).map(kv =>
    '<div class="spec"><div class="k">' + esc(kv[0]) + '</div><div class="v">'
    + esc(kv[1]) + '</div></div>').join("");
  const meta = [];
  if (e.t) meta.push('<span class="path">' + esc(e.t) + '</span>');
  if (e.d) meta.push('<span class="dwg">custom DWGs: ' + esc(e.d) + '</span>');
  if (e.f) meta.push('<a href="' + esc(folderUrl(e.f)) + '" title="'
    + esc(e.f) + ' — opens in File Explorer">AutoCAD folder</a>');
  if (e.sw) meta.push('<a href="' + esc(folderUrl(e.sw)) + '" title="'
    + esc(e.sw) + ' — opens in File Explorer">SolidWorks 3D</a>');
  const hist = e.h ? '<details class="hist"><summary>CO history ('
    + e.h.length + ')</summary>'
    + e.h.map(x => "<div>" + esc(x) + "</div>").join("") + "</details>" : "";

  const compCard = (c, path) => {
    const active = !state.whole && state.path === path;
    const items = c.i.map(no => itemByNo(e, no)).filter(Boolean);
    const attrs = Object.entries(c.a).map(([k, v]) => {
      const key = k + "=" + v;
      const pin = state.pinned.has(key) && !state.whole && state.path === path;
      return '<button class="attr-row' + (pin ? " pinned" : "")
        + '" data-a="' + esc(key) + '" data-p="' + path
        + '" title="Find orders whose ' + esc(c.n)
        + ' has this attribute — most similar fans first">'
        + '<span class="k">' + esc(k.replace(/_/g, " ")) + ':</span>'
        + '<span class="v">' + esc(v) + '</span>'
        + '<span class="pin">' + (pin ? "✕ required" : "find matches") + "</span></button>";
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
    + '<button class="backlink" id="back">← back</button>'
    + '<div class="ohead"><span class="job">' + esc(j) + '</span>'
    + '<span class="cust">' + esc(e.c) + "</span>"
    + (e.co ? '<span class="co">' + esc(e.co) + "</span>" : "")
    + queueBadge(j, e)
    + (e.pdf ? '<a href="' + esc(fileUrl(e.pdf)) + '" target="_blank" title="'
        + esc(e.pdf) + '">Open SO PDF</a>' : "")
    + (e.qr ? '<a href="' + esc(fileUrl(e.qr)) + '" target="_blank" title="'
        + esc(e.qr) + '">Open Quote Run</a>' : "")
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

  $("back").onclick = () => history.length > 1 ? history.back() : setTab("board");
  $("whole").onclick = () => { state.whole = !state.whole;
    if (state.whole) { state.path = null; state.pinned.clear(); } render(); };
  el.querySelectorAll(".comp-row").forEach(b => b.onclick = () => {
    state.whole = false; state.path = b.dataset.c; state.pinned.clear(); render();
  });
  /* An attribute click works standalone: it targets the component the
     attribute belongs to and requires the attribute on every match. */
  el.querySelectorAll(".attr-row").forEach(b => b.onclick = () => {
    const k = b.dataset.a, p = b.dataset.p;
    if (state.whole || state.path !== p) {
      state.whole = false; state.path = p; state.pinned.clear();
    }
    state.pinned.has(k) ? state.pinned.delete(k) : state.pinned.add(k);
    render();
  });
}

function renderMatches() {
  const el = $("right");
  const ready = state.job && (state.whole || state.path !== null);
  if (!ready) {
    el.innerHTML = '<div class="panel-head"><span class="eyebrow">Matching orders</span></div>'
      + '<div class="empty"><div class="big">'
      + (state.job ? "Click a component on the left" : "No order open yet")
      + "</div>Past orders sharing it appear here, most relevant first, each "
      + "with the line items that made it a match.</div>";
    return;
  }
  const j = state.job, e = DB.jobs[j];
  const srcDesign = spv(e, "Design");
  const items = state.whole ? e.it : compItems(e, state.path);
  const target = state.whole ? "whole order" : (() => {
    const c = compAt(e, state.path);
    return c ? (c.k ? "[" + c.n + "]" : c.n) : "?";
  })();
  const resAll = rankMatches(j, items);
  const res = state.only3d ? resAll.filter(r => DB.jobs[r.j].sw) : resAll;
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
    if (o.f) foot.push('<a href="' + esc(folderUrl(o.f)) + '" title="'
      + esc(o.f) + ' — opens in File Explorer">AutoCAD folder</a>');
    if (o.sw) foot.push('<a href="' + esc(folderUrl(o.sw)) + '" title="'
      + esc(o.sw) + ' — opens in File Explorer">SolidWorks 3D</a>');
    if (o.pdf) foot.push('<a href="' + esc(fileUrl(o.pdf))
      + '" target="_blank" title="' + esc(o.pdf) + '">SO PDF</a>');
    if (o.qr) foot.push('<a href="' + esc(fileUrl(o.qr))
      + '" target="_blank" title="' + esc(o.qr) + '">Quote Run</a>');
    const theirDesign = spv(o, "Design");
    const designChip = srcDesign && theirDesign === srcDesign
      ? '<span class="chip same">✓ design ' + esc(srcDesign) + "</span>" : "";
    const spec = fanSpec(o);
    return '<div class="match"><div class="m-head">'
      + '<span class="m-rank">' + (i + 1) + ".</span>"
      + '<button class="m-job" data-job="' + esc(r.j) + '" title="Open this order">'
      + esc(r.j) + "</button>"
      + '<span class="m-cust">' + esc(o.c) + (o.co ? " · " + esc(o.co) : "")
      + "</span>" + queueBadge(r.j, o) + designChip
      + '<span class="m-score">score ' + r.score.toFixed(2) + "</span></div>"
      + (spec ? '<div class="m-spec">' + esc(spec) + "</div>" : "")
      + '<div class="m-scorebar"><i style="width:'
      + Math.max(6, 100 * r.score / max) + '%"></i></div>'
      + '<div class="m-lines">' + head + more + "</div>"
      + chipsHtml + '<div class="m-foot">' + foot.join(" ") + "</div></div>";
  }).join("");

  el.innerHTML = '<div class="panel-head"><span class="eyebrow">Matching orders</span>'
    + '<span class="m-target">' + esc(target) + '</span>'
    + '<span class="m-cust">on ' + esc(j) + "</span>"
    + '<button class="wholebtn' + (state.only3d ? " active" : "") + '" id="only3d" '
    + 'title="Only orders with SolidWorks 3D files (parts / assemblies / drawings)"'
    + ">Has 3D</button>"
    + '<span class="m-count" style="margin-left:0">' + res.length + " match"
    + (res.length === 1 ? "" : "es")
    + (state.only3d && resAll.length !== res.length
        ? " (of " + resAll.length + ")" : "")
    + (srcDesign ? " · same fan design (" + esc(srcDesign) + ") ranks first" : "")
    + "</span></div>"
    + (state.pinned.size
        ? '<div class="filterbar"><span class="fl">Required attributes:</span>'
          + chips + "</div>" : "")
    + (res.length ? cards : '<div class="empty"><div class="big">No orders match</div>'
        + (state.only3d && resAll.length
            ? "None of the " + resAll.length + " matches has SolidWorks 3D data "
              + "on file — turn off Has 3D, or run the SolidWorks scan if it "
              + "has never been run."
            : state.pinned.size ? "Try removing an attribute filter."
              : "Nothing in the store shares this yet.") + "</div>")
    + (res.length > 25 ? '<div class="tailnote">…and ' + (res.length - 25)
        + " more — narrow it with an attribute filter</div>" : "");
  el.querySelectorAll(".fchip").forEach(b => b.onclick = () => {
    state.pinned.delete(b.dataset.a); render(); });
  el.querySelectorAll(".m-job").forEach(b => b.onclick = () => selectJob(b.dataset.job));
  $("only3d").onclick = () => { state.only3d = !state.only3d; savePrefs(); render(); };
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
          + queueBadge(h.j, e) + "</span>"
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
document.querySelectorAll(".tabbtn").forEach(b =>
  b.onclick = () => setTab(b.dataset.tab));

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
