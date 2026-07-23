"""Tests for order_explorer.py — the GL Queue Explorer page generator.

Run:  python test_order_explorer.py
Pure logic only (payload building, HTML embedding, the .bat launcher); no
browser, no COM, no network — safe for CI.
"""
from __future__ import annotations

import base64
import gzip
import json
import os
import re
import tempfile
from pathlib import Path

import order_explorer as oe
import prepare_transmittal_link as ptl
import prepare_so_review_note_link as pnl


# The docstring example from so_hierarchy: three printed lines, one real IVC.
_IVC_ITEMS = [
    {"raw": "Inlet Volume Control, Low Leak, Automatic 3,531.00",
     "norm": "INLET VOLUME CONTROL LOW LEAK AUTOMATIC", "qty": "", "price": "3,531.00",
     "ptype": "L", "section": "", "details": [], "tags": ["IVC"],
     "attributes": {"used_on": "IVC", "leakage_class": "LOW LEAKAGE",
                    "operation": "Automatic"}},
    {"raw": "Inlet Volume Control Handle Location, Non-Standard 425.00",
     "norm": "INLET VOLUME CONTROL HANDLE LOCATION NON STANDARD", "qty": "",
     "price": "425.00", "ptype": "L", "section": "", "details": [], "tags": ["IVC"],
     "attributes": {"used_on": "IVC", "handle_location": "4:30 (NON-STD)"}},
    {"raw": "Shaft Seal, Teflon 610.00", "norm": "SHAFT SEAL TEFLON", "qty": "",
     "price": "610.00", "ptype": "L", "section": "", "details": [],
     "tags": ["SHAFT SEAL"], "attributes": {"component": "SHAFT SEAL"}},
]


def _store():
    return {"jobs": {
        "421966": {"customer": "Meridian Foundry", "co_number": 2,
                   "so_pdf": r"Z:\SO\421966.pdf", "items": _IVC_ITEMS},
        "421314": {"customer": "Bayside Chemical", "co_number": None,
                   "so_pdf": "", "items": [_IVC_ITEMS[0]]},
    }}


def _dwg():
    return {"421314": {"job": "421314", "type": "BC",
                       "folder": r"Z:\AUTOCAD\CURRENT\JOBS\BC\421\421314",
                       "extras": {"07": "DWG", "51": "PDF+DWG"}}}


def test_payload_components_merge():
    p = oe.build_payload(_store(), _dwg())
    assert isinstance(p["v"], str), "code version stamp missing"
    e = p["jobs"]["421966"]
    assert e["c"] == "Meridian Foundry", e["c"]
    assert e["co"] == "CO#2", e["co"]
    assert len(e["it"]) == 3, e["it"]
    # Two components: the merged [IVC] (2 lines) and the shaft seal.
    names = [c["n"] for c in e["cp"]]
    assert names == ["IVC", "SHAFT SEAL"], names
    ivc = e["cp"][0]
    assert ivc["k"] == 1 and len(ivc["i"]) == 2, ivc
    assert ivc["a"].get("leakage_class") == "LOW LEAKAGE", ivc["a"]
    assert ivc["p"] == 3956.0, ivc["p"]          # 3,531 + 425
    # DWG enrichment lands on the other job.
    e2 = p["jobs"]["421314"]
    assert e2["d"] == "-07, -51", e2.get("d")
    assert e2["f"].endswith("421314"), e2.get("f")
    assert e2["t"] == "BC", e2.get("t")
    # Each custom drawing carries its own [suffix, ext] link pair (PDF when a
    # PDF exists for that suffix, else DWG), numerically ordered.
    assert e2["dl"] == [["07", "dwg"], ["51", "pdf"]], e2.get("dl")
    print("  payload components/enrichment OK")


def test_queue_jobs_take_master_items():
    fresh = [{"raw": "Motor, 40 HP, TEFC 6,890.00", "norm": "MOTOR 40 HP TEFC",
              "qty": "", "price": "6,890.00", "ptype": "L", "section": "",
              "details": [], "tags": ["MOTOR"],
              "attributes": {"component": "MOTOR"}}]
    qjob = {"job": "421966", "customer": "Meridian Foundry", "co_number": 4,
            "so_pdf": r"Z:\SO\421966 CO4.pdf", "line_items": fresh,
            "design": "BC-3660", "so_size": "366",
            "status": "AT WC", "oper": "23", "end_date": "7/16/2026",
            "total_price": "$47,763.00", "has_drive_run": True, "_cbc_pos": 3,
            "_added_iso": "2026-07-17T06:01:00",
            "co_history": ["CO#4 071726 DG - ADDED VFD CONTROLS",
                           "CO#3 " + "x" * 500]}
    p = oe.build_payload(_store(), queue_jobs={"421966": qjob})
    e = p["jobs"]["421966"]
    assert e.get("q") == 1, e
    assert e["co"] == "CO#4", e["co"]                    # master beats store
    assert e["cd"] == "ADDED VFD CONTROLS", e.get("cd")
    assert len(e["it"]) == 1 and e["it"][0][5] == "MOTOR 40 HP TEFC", e["it"]
    assert ["Design", "BC-3660"] in e["sp"] and ["Size", "366"] in e["sp"], e["sp"]
    assert len(e["h"]) == 2 and len(e["h"][1]) <= 160, e["h"]
    # The Board fields ride along for the on-board order.
    bd = e["bd"]
    assert bd["ed"] == "7/16/2026" and bd["pr"] == "$47,763.00", bd
    assert bd["dr"] == 1 and bd["ps"] == 3 and bd["ai"].startswith("2026-07-17"), bd
    # The watcher's exact new-today set is embedded verbatim when given.
    p2 = oe.build_payload(_store(), queue_jobs={"421966": qjob},
                          new_ids={"421966"})
    assert p2["nw"] == ["421966"], p2.get("nw")
    assert "nw" not in p, "nw must be absent when new_ids not passed"
    # A store-only job is still present (the match pool).
    assert "421314" in p["jobs"] and "q" not in p["jobs"]["421314"]
    print("  queue-job override / spec / board fields OK")


def test_events_and_removed():
    master_orders = {
        "421966": {"on_queue": True, "seen_on_queue": True,
                   "job": {"job": "421966", "customer": "Meridian Foundry",
                           "line_items": [],
                           "co_history": ["CO#2 070126 ABC - ADDED SHAFT COOLER"]}},
        "421000": {"on_queue": False, "seen_on_queue": True,
                   "left": "2026-07-17T09:30:00", "job": {"job": "421000"}},
        "420900": {"on_queue": False, "seen_on_queue": True,
                   "left": "2026-07-10T09:30:00", "job": {"job": "420900"}},
    }
    events = [
        {"time": "2026-07-17T08:00:00", "job": "421966", "customer": "Meridian",
         "field": "CO#", "old": "1", "new": "2"},
        {"time": "2026-07-17T09:00:00", "job": "421966", "customer": "Meridian",
         "field": "End Date", "old": "7/10/2026", "new": "7/20/2026"},
        {"time": "2026-07-17T09:00:00", "job": "421966", "customer": "Meridian",
         "field": "Features", "old": "ACCESS DOOR", "new": "ACCESS DOOR, EVASE"},
        {"time": "2026-07-17T09:00:00", "job": "421966", "customer": "Meridian",
         "field": "Line Items", "old": "old raw parser blob", "new": "new raw parser blob"},
    ]
    from datetime import date as _date
    p = oe.build_payload(_store(), master_orders=master_orders, events=events,
                         today=_date(2026, 7, 17))
    assert p["today"] == "2026-07-17"
    # Newest first; the CO event carries its co_history description.
    assert [e["f"] for e in p["ev"]] == ["End Date", "CO#"], p["ev"]
    assert all(e["f"] not in ("Features", "Line Items") for e in p["ev"])
    co = p["ev"][1]
    assert co["d"] == "ADDED SHAFT COOLER", co
    # Only the departure dated today lands in the removed list.
    assert p["rm"] == [["421000", "2026-07-17T09:30:00"]], p["rm"]
    # A master-only order (never in the store) still gets an Order History row.
    assert p["jobs"]["420900"]["oh"] == [0, "", "2026-07-10T09:30:00"]
    assert p["jobs"]["421966"]["oh"][0] == 1
    print("  change events / removals / master-only history OK")


def test_spec_values_tidied():
    """Size/Arrangement show the workbook's short forms (split_size /
    split_arrangement), not the raw SO text."""
    qjob = {"job": "421966", "customer": "X", "line_items": [],
            "so_size": "2412 (3600 RPM or less)", "so_arrangement": "Arrangement 4"}
    p = oe.build_payload(_store(), queue_jobs={"421966": qjob})
    sp = dict(p["jobs"]["421966"]["sp"])
    assert sp["Size"] == "2412", sp["Size"]
    assert sp["Arrangement"] == "A/4", sp["Arrangement"]
    qjob["so_arrangement"] = "A/4V C-Face Flange mount (no motor base)"
    qjob["so_size"] = "3650-C12 Blade-1800"
    p = oe.build_payload(_store(), queue_jobs={"421966": qjob})
    sp = dict(p["jobs"]["421966"]["sp"])
    assert sp["Arrangement"] == "A/4V", sp["Arrangement"]
    assert sp["Size"] == "3650", sp["Size"]
    print("  size/arrangement tidying OK")


def test_quote_run_and_solidworks_links():
    """drive_run_pdf becomes the Run/Quote-Run link target; a SolidWorks scan
    record with 3D files becomes the sw folder link (and only then)."""
    master_orders = {"421966": {"on_queue": True,
                                "job": {"job": "421966", "customer": "X",
                                        "line_items": [],
                                        "drive_run_pdf": r"Z:\A\421966\ENG REF\RUN.txt"}}}
    sw = {"421966": {"has_sw": True, "folder": r"Z:\SW\421966"},
          "421314": {"has_sw": False, "folder": r"Z:\SW\421314"}}
    p = oe.build_payload(_store(), master_orders=master_orders, sw=sw)
    assert p["jobs"]["421966"]["qr"] == r"Z:\A\421966\ENG REF\RUN.txt"
    assert p["jobs"]["421966"]["sw"] == r"Z:\SW\421966"
    # Scanned-but-empty folders must NOT count as having 3D.
    assert "sw" not in p["jobs"]["421314"]
    assert "qr" not in p["jobs"]["421314"]
    print("  quote-run / solidworks links OK")


def test_wheel_component_from_quote_run():
    """A parsed quote run yields the synthetic [WHEEL (QUOTE RUN)] component:
    blade type from the SO spec, one attribute + pseudo-item per tracked wheel
    part, numbering continuing after the real items."""
    qjob = {"job": "421966", "customer": "X", "line_items": _IVC_ITEMS,
            "so_wheel_type": "BC",
            "drive_run": {"Blade Gauge": "0.075 (14)", "Number of Blades": "12",
                          "Backplate Material": "ASTM CQ HRS A36",
                          "Hub": "19-5-1057", "Wheel Weight Lb": "199"}}
    p = oe.build_payload(_store(), queue_jobs={"421966": qjob})
    e = p["jobs"]["421966"]
    comp = e["cp"][0]
    assert comp["n"] == oe.WHEEL_COMP_NAME and comp["k"] == 1 and comp["hs"] == 1
    a = comp["a"]
    assert a["blade type"] == "BC" and a["blades"] == "12", a       # alias merged
    assert a["blade gauge"] == "0.075 (14)" and a["hub"] == "19-5-1057", a
    assert "wheel weight lb" not in a, "performance stats must stay out"
    # Pseudo-items: contiguous numbering after the 3 real items, stable norms.
    wheel_rows = [r for r in e["it"] if r[4] == "QUOTE RUN"]
    assert [r[0] for r in wheel_rows] == [4, 5, 6, 7, 8], wheel_rows
    assert comp["i"] == [4, 5, 6, 7, 8], comp["i"]
    norms = {r[5] for r in wheel_rows}
    assert "WHEEL BLADE TYPE BC" in norms and "WHEEL BLADES 12" in norms, norms
    assert "WHEEL BLADE GAUGE 0 075 14" in norms, norms
    assert all(r[6] == ["WHEEL RUN"] for r in wheel_rows)
    # No run -> no wheel component.
    p2 = oe.build_payload(_store(), queue_jobs={"421966": {"job": "421966",
                          "line_items": _IVC_ITEMS, "so_wheel_type": "BC"}})
    assert all(c["n"] != oe.WHEEL_COMP_NAME for c in p2["jobs"]["421966"]["cp"])
    print("  quote-run wheel component OK")


def test_master_orders_fallback_queue():
    master_orders = {"421314": {"on_queue": True, "added": "2026-07-15T08:30:00",
                                "job": {"job": "421314", "customer": "Bayside",
                                        "line_items": [], "co_number": 1}}}
    p = oe.build_payload(_store(), master_orders=master_orders)
    assert p["jobs"]["421314"].get("q") == 1
    # Empty master line_items falls back to the store's captured items.
    assert len(p["jobs"]["421314"]["it"]) == 1
    # The job dict has no live _added_iso stamp (snapshot build), so the
    # master entry's arrival timestamp — what the workbook shows — is used.
    assert p["jobs"]["421314"]["bd"]["ai"] == "2026-07-15T08:30:00"
    # The watcher persists each order's board position into the master job
    # dict, so an Open-button rebuild (no live scrape) keeps the "#" column
    # and the cbcinsider row order.
    master_orders["421314"]["job"]["_cbc_pos"] = 3
    p = oe.build_payload(_store(), master_orders=master_orders)
    assert p["jobs"]["421314"]["bd"]["ps"] == 3
    print("  master on_queue fallback / added timestamp OK")


def test_render_roundtrip_and_safety():
    store = _store()
    # Hostile strings must survive the embedding: script closers, quotes, unicode.
    store["jobs"]["421966"]["items"][0]["raw"] = 'X </script> "quote" 650°F & <b>'
    p = oe.build_payload(store)
    html = oe.render_html(p)
    assert "</script> \"quote\"" not in html            # never embedded raw
    m = re.search(r'const PAYLOAD_B64 = "([A-Za-z0-9+/=]+)"', html)
    assert m, "payload marker missing"
    back = json.loads(gzip.decompress(base64.b64decode(m.group(1))).decode("utf-8"))
    assert back == p, "payload did not round-trip"
    assert html.count("__B64__") == 0
    assert "__SIMILARITY_JS__" not in html
    assert "GLQSimilarity" in html
    assert "0.000..1.000" in html
    assert "renderOrderPreview" in html
    assert "previewJob" in html
    assert "Back to List" in html
    assert 'rel="icon" type="image/png" sizes="256x256" href="GL Queue Explorer Fan.png"' in html
    assert 'rel="shortcut icon" type="image/x-icon" href="GL Queue Explorer Fan.ico"' in html
    assert 'rel="manifest" href="GL Queue Explorer.webmanifest"' in html
    assert "click order # again to move it to the left" in html
    assert "red = scored construction difference" in html
    assert "preview-relevant" in html
    assert "preview-match" in html
    assert "changeHeaders" in html and "Field #" in html
    assert "chgfield" in html
    assert "green = selected combination match" in html
    assert "combinedFocusedSimilarity" in html
    assert "state.selections" in html
    assert "Searching for (closest first)" in html
    assert "pinHits" in html          # soft-selection miss chips are wired in
    assert "specHits" in html and "specCardChips" in html
    assert "match-actions" in html
    # Fan stats are selectable search pins; Match Base Fan presses them all.
    assert "Match Base Fan" in html and "spec-btn" in html
    assert 'data-spec="' in html and 'data-kind="spec"' in html
    assert "specAssessment" in html and "combinedSearchSimilarity" in html
    assert "baseFanLabels" in html and "state.specSel" in html
    assert "selected components" in html
    assert "alignComponentStarts" in html
    assert 'class="co-tip"' in html
    assert "No change-order description was captured." in html
    assert 'id="leftcomponenttree"' in html
    assert 'id="rightcomponenttree"' in html
    assert "ontoggle = alignComponentStarts" in html
    assert "<title>GL Queue Explorer</title>" in html
    assert html.count("Prepare Transmittal") == 2
    assert 'class="prepare-transmittal"' in html
    assert "const transmittalUrl" in html and "glqtransmittal:" in html
    assert "noteTargetHtml" in html and "wireInlineNotes" in html
    assert "data-note-target" in html and "data-note-id" in html
    assert "state.heldVersion" in html and "refresh paused while note editing" in html
    assert "noteDraftKey" in html and "hasNoteDrafts" in html
    assert "note-del" in html and "action: \"delete\"" in html
    assert "noteSelections" in html and "review|" in html and "data-r=" in html
    assert "deepLinkJob" in html and "#order=" in html and "applyDeepLink" in html
    # In both order headers the action is immediately after that order's customer.
    assert re.search(r"esc\(e\.c\).*?prepare-transmittal.*?transmittalUrl\(j\)",
                     html, re.DOTALL)
    assert re.search(r"esc\(right\.c\).*?prepare-transmittal.*?transmittalUrl\(rightJob\)",
                     html, re.DOTALL)
    print("  render round-trip / embedding safety OK")


def test_bat_launcher():
    text = oe.bat_text("My Page.html")
    text.encode("ascii")                                 # cmd.exe-safe
    assert "\r\n" in text and "\n" == text[-1] or text.endswith("\r\n")
    assert 'set "PAGE=%~dp0My Page.html"' in text
    assert "--app=" in text and "msedge" in text.lower()
    # Registers the per-user glqueue: handler (File Explorer folder links).
    assert text.count("reg add") == 3 and "glq_open.vbs" in text
    assert r"HKCU\Software\Classes\glqueue" in text and "%%1" in text
    print("  bat launcher OK")


def test_vbs_folder_opener():
    text = oe.vbs_text()
    text.encode("ascii")                                 # wscript-safe
    assert text.endswith("\r\n") and "\r\n" in text
    assert "DecodeUrl" in text and "explorer.exe" in text
    assert 'LCase(Left(u, 8)) = "glqueue:"' in text
    # find? mode resolves the real drawing filename (revision letter) at click.
    assert 'LCase(Left(u, 5)) = "find?"' in text
    assert "FindDrawing" in text and "QueryParam" in text
    print("  vbs folder opener OK")


def test_dwg_links_render_as_links():
    p = oe.build_payload(_store(), _dwg())
    html = oe.render_html(p)
    assert "a.dwglink" in html                 # style present
    assert 'glqueue:find?dir=' in html or 'dwgUrl' in html
    assert "dwgListHtml" in html
    # PDF or PDF+DWG suffixes prefer pdf; DWG-only prefers dwg.
    links = oe._dwg_links({"07": "PDF", "51": "PDF+DWG", "12": "DWG"})
    assert links == [["07", "pdf"], ["12", "dwg"], ["51", "pdf"]], links
    print("  dwg links OK")


def test_transmittal_protocol_accepts_only_a_numeric_order():
    assert ptl.parse_order_uri("glqtransmittal:421968") == "421968"
    assert ptl.parse_order_uri("GLQTRANSMITTAL:421968") == "421968"
    for bad in (
        "glqtransmittal:", "glqtransmittal://421968",
        "glqtransmittal:421968&send=1", "glqtransmittal:421968/anything",
        "glqueue:421968", "glqtransmittal:abc",
    ):
        try:
            ptl.parse_order_uri(bad)
        except ValueError:
            pass
        else:
            raise AssertionError(f"unsafe transmittal URI was accepted: {bad}")
    command = ptl.protocol_command(Path(r"C:\CBC INSIDER GL QUEUE CHECKER"),
                                   Path(r"C:\Python\python.exe"))
    assert "prepare_transmittal_link.py" in command and "%1" in command
    print("  transmittal protocol validation OK")



def test_note_protocol_records_review_note(tmp_path: Path):
    uri = "glqnote:?order=421968&item_no=2&item_text=WHEEL&row_key=item%3A2&note=Check%20material"
    parsed = pnl.parse_note_uri(uri)
    assert parsed["order"] == "421968"
    assert parsed["row_key"] == "item:2"
    assert parsed["note"] == "Check material"
    deleted = pnl.parse_note_uri("glqnote:?action=delete&order=421968&note_id=7")
    assert deleted["action"] == "delete" and deleted["note_id"] == "7"
    assert "prepare_so_review_note_link.py" in pnl.protocol_command(Path("/tmp/repo"), "python")

    calls = []
    orig_load, orig_record, orig_save = (
        pnl.so_review.load_store, pnl.so_review.record_note, pnl.so_review.save_store)
    orig_publish = pnl.publish_review_notes
    try:
        pnl.so_review.load_store = lambda: {"notes": []}
        pnl.so_review.record_note = lambda store, order, item, text, note, row_key="": calls.append(
            (order, item, text, note, row_key)) or {"id": 1}
        pnl.so_review.save_store = lambda store: calls.append(("save", len(store["notes"])))
        pnl.publish_review_notes = lambda: calls.append(("push",)) or True
        added, pushed = pnl.record_from_uri(uri)
        store = {"notes": [{"id": 7, "order": "421968", "note": "old"},
                           {"id": 8, "order": "421968", "note": "keep"}]}
        assert pnl.delete_note(store, "421968", "7")
        assert [n["id"] for n in store["notes"]] == [8]
    finally:
        pnl.so_review.load_store = orig_load
        pnl.so_review.record_note = orig_record
        pnl.so_review.save_store = orig_save
        pnl.publish_review_notes = orig_publish
    assert added and pushed
    assert ("421968", "2", "WHEEL", "Check material", "item:2") in calls
    assert ("push",) in calls
    print("  note protocol validation / immediate publish OK")

def test_transmittal_protocol_runs_existing_review_flow_and_logs():
    seen = []
    original_cwd = Path.cwd()

    def fake_fill_main(args):
        seen.append((args, Path.cwd()))
        print("review form prepared; Send remains disabled")
        return 0

    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        (root / ".launcher_state.json").write_text(json.dumps({
            "options": {"email_drawings": {"initials": "JZ"}},
        }), encoding="utf-8")
        assert ptl.run_transmittal("421968", root=root,
                                   fill_main=fake_fill_main) == 0
        logs = list((root / "launcher_logs").glob("*_email_drawings.log"))
        assert len(logs) == 1
        text = logs[0].read_text(encoding="utf-8")
        assert "fill_transmittal_insider.py 421968 --initials JZ" in text
        assert "Preparing transmittal for order 421968" in text
        assert "Send remains disabled" in text
        assert seen == [(["421968", "--initials", "JZ"], root.resolve())]
    assert Path.cwd() == original_cwd
    print("  transmittal protocol review handoff / logging OK")


def test_transmittal_protocol_is_single_instance_on_windows():
    if os.name == "nt":
        with ptl.transmittal_instance_lock() as first:
            assert first is True
            with ptl.transmittal_instance_lock() as second:
                assert second is False
    print("  transmittal protocol single-instance gate OK")


def test_default_output_path_accepts_folder():
    """EXPLORER_PATH may be the page file OR its folder — a folder (or any
    non-.html path) gets the standard page name appended."""
    saved = oe.EXPLORER_PATH
    try:
        folder = Path("/z/GL QUEUE LIVE")        # separator-portable stand-in
        oe.EXPLORER_PATH = folder
        assert oe.default_output_path() == folder / oe.HTML_NAME
        oe.EXPLORER_PATH = folder / "My Page.html"
        assert oe.default_output_path() == folder / "My Page.html"
    finally:
        oe.EXPLORER_PATH = saved
    print("  EXPLORER_PATH folder/file forms OK")


def test_default_output_path_is_coworker_share():
    expected = Path(r"\\gdh-fs02\engineering\DAG\GL QUEUE LIVE") / oe.HTML_NAME
    assert oe.default_output_path() == expected, oe.default_output_path()
    print("  default Explorer path is the canonical coworker share")


def test_write_explorer_files():
    p = oe.build_payload(_store(), _dwg())
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "GL Queue Explorer.html"
        got = oe.write_explorer(p, out)
        assert got == out and out.exists() and out.stat().st_size > 10_000
        bat = Path(td) / oe.BAT_NAME
        assert bat.exists(), "launcher .bat not written"
        assert (Path(td) / oe.VBS_NAME).exists(), "glq_open.vbs not written"
        assert (Path(td) / oe.ICON_PNG_NAME).exists(), "Explorer PNG icon not written"
        assert (Path(td) / oe.ICON_ICO_NAME).exists(), "Explorer ICO icon not written"
        manifest = Path(td) / oe.MANIFEST_NAME
        assert manifest.exists(), "Explorer web app manifest not written"
        assert oe.ICON_PNG_NAME in manifest.read_text(encoding="utf-8")
        shortcut = Path(td) / oe.SHORTCUT_VBS_NAME
        assert shortcut.exists(), "custom-icon shortcut helper not written"
        shortcut_text = shortcut.read_text(encoding="ascii")
        assert oe.ICON_ICO_NAME in shortcut_text and "CreateShortcut" in shortcut_text
        # The auto-refresh stamp open pages poll: the data fingerprint drives
        # the all-tab reload, the board fingerprint the Live-Queue-only reload.
        ver = Path(td) / oe.VERSION_NAME
        assert ver.exists(), "version stamp not written"
        vtext = ver.read_text(encoding="utf-8")
        assert "__GLQ_VERSION__" in vtext and "__GLQ_BOARD_VERSION__" in vtext
        assert oe._fingerprint(p, drop_board=True) in vtext, "content stamp missing"
        assert oe._fingerprint(p, drop_board=False) in vtext, "board stamp missing"
        first = bat.read_bytes()
        oe.write_explorer(p, out)                        # idempotent second write
        assert bat.read_bytes() == first
        assert not list(Path(td).glob("*.tmp")), "temp file left behind"
    print("  write_explorer files OK")


def test_content_fingerprint_ignores_volatile_columns():
    qjob = {"job": "421966", "status": "IN PROCESS", "_cbc_pos": 3,
            "total_price": "10,000", "assigned_to": "JZ"}
    base = oe.build_payload(_store(), queue_jobs={"421966": qjob})
    fp = oe._content_fingerprint
    baseline = fp(base)

    # The build timestamp never affects the fingerprint.
    bumped = json.loads(json.dumps(base))
    bumped["gen"] = "2099-01-01 00:00"
    assert fp(bumped) == baseline

    # Neither do the churny board columns (status / live position / price).
    moved = json.loads(json.dumps(base))
    assert "bd" in moved["jobs"]["421966"], "expected board fields on a queued job"
    moved["jobs"]["421966"]["bd"] = {"st": "SHIPPED", "ps": 99, "pr": "1"}
    assert fp(moved) == baseline

    # But genuinely new line-item data does change it...
    newdata = json.loads(json.dumps(base))
    newdata["jobs"]["421966"]["it"].append(["EVASE", "Discharge Evase"])
    assert fp(newdata) != baseline

    # ...as do a new change event and a queue-membership change.
    ev = json.loads(json.dumps(base))
    ev["ev"] = [{"job": "421966", "kind": "added"}]
    assert fp(ev) != baseline
    left = json.loads(json.dumps(base))
    assert left["jobs"]["421966"].get("q") == 1, "queued job should carry q=1"
    del left["jobs"]["421966"]["q"]
    assert fp(left) != baseline

    # The board fingerprint (drop_board=False) DOES react to a board change, so
    # the Live Queue tab can be told to reload for it.
    board_base = oe._fingerprint(base, drop_board=False)
    board_moved = json.loads(json.dumps(base))
    board_moved["jobs"]["421966"]["bd"] = {"st": "SHIPPED", "ps": 99}
    assert oe._fingerprint(board_moved, drop_board=False) != board_base
    # ...while the content fingerprint stays put (no all-tab reload for it).
    assert oe._content_fingerprint(board_moved) == baseline
    print("  content vs board fingerprint split OK")


def test_board_signature_tracks_column_changes():
    q1 = {"421966": {"job": "421966", "status": "IN PROCESS", "_cbc_pos": 3}}
    q2 = {"421966": {"job": "421966", "status": "SHIPPED", "_cbc_pos": 3}}
    q3 = {"421966": {"job": "421966", "status": "IN PROCESS", "_cbc_pos": 4}}
    assert oe._board_signature(q1) == oe._board_signature(dict(q1))
    assert oe._board_signature(q1) != oe._board_signature(q2), "status change missed"
    assert oe._board_signature(q1) != oe._board_signature(q3), "position change missed"
    assert oe._board_signature({}) == oe._board_signature({})
    print("  board signature tracks status/position OK")


def test_maybe_write_only_publishes_on_new_data():
    """The watcher hook republishes on genuinely new data only: unchanged polls
    and no-op store re-saves never rewrite the page (which would reload every
    open browser)."""
    import line_items as li
    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d)
        out = tmp / "explorer.html"
        store_file = tmp / "line_items.json"
        store_file.write_text("{}", encoding="utf-8")

        saved = {name: getattr(oe, name) for name in (
            "default_output_path", "build_payload", "write_explorer",
            "_ensure_transmittal_link_handler", "_ensure_note_link_handler")}
        saved_store_path, saved_cache = li.store_path, dict(oe._CACHE)
        writes: list = []
        payload_box = {"p": {"gen": "t0", "jobs": {}, "n_items": 0}}

        def _bump_store(seconds: int) -> None:
            st = store_file.stat()
            os.utime(store_file, (st.st_atime, st.st_mtime + seconds))

        try:
            oe.default_output_path = lambda: out
            li.store_path = lambda: store_file
            oe.build_payload = lambda *a, **k: payload_box["p"]
            oe._ensure_transmittal_link_handler = lambda *a, **k: None
            oe._ensure_note_link_handler = lambda *a, **k: None

            def _fake_write(payload, o):
                writes.append(payload)
                Path(o).write_text("page", encoding="utf-8")
                return Path(o)
            oe.write_explorer = _fake_write
            oe._CACHE = {"touch": None, "full": None, "at": 0.0}

            # First poll — nothing published yet, so it writes.
            assert oe.maybe_write(None, []) == out
            assert len(writes) == 1

            # Same inputs, nothing touched -> stage-1 skip, no rebuild.
            assert oe.maybe_write(None, []) is None
            assert len(writes) == 1

            # A store re-save that carried no new data (only a fresh gen):
            # stage 1 proceeds, but the content fingerprint is unchanged.
            payload_box["p"] = {"gen": "t0-resave", "jobs": {}, "n_items": 0}
            _bump_store(100)
            assert oe.maybe_write(None, []) is None
            assert len(writes) == 1

            # Genuinely new data + another re-save -> republishes.
            payload_box["p"] = {"gen": "t1", "jobs": {"9": {"it": [1]}}, "n_items": 1}
            _bump_store(200)
            assert oe.maybe_write(None, []) == out
            assert len(writes) == 2

            # A board-only change (bd) still republishes, so the Live Queue tab
            # can pull the current board — even though it's not "new data".
            payload_box["p"] = {"gen": "t2",
                                "jobs": {"9": {"it": [1], "bd": {"st": "SHIPPED"}}},
                                "n_items": 1}
            _bump_store(300)
            assert oe.maybe_write(None, []) == out
            assert len(writes) == 3
        finally:
            for name, fn in saved.items():
                setattr(oe, name, fn)
            li.store_path, oe._CACHE = saved_store_path, saved_cache
    print("  maybe_write publishes only on new data OK")


def test_explorer_search_supports_multi_part_scored_queries():
    html = oe.render_html({"gen": "now", "jobs": {}, "n_items": 0})
    assert "function searchParts" in html
    assert "function scoreSearchEntry" in html
    assert "complete matches first" in html
    assert "h.score.toFixed(2)" in html
    assert "class=\"missing\">missing:" in html
    assert "function searchPartAliases" in html
    assert "multiple features: D16, S245, access door" in html
    print("  multi-part scored Explorer search UI OK")


def main() -> int:
    test_payload_components_merge()
    test_queue_jobs_take_master_items()
    test_events_and_removed()
    test_spec_values_tidied()
    test_quote_run_and_solidworks_links()
    test_wheel_component_from_quote_run()
    test_master_orders_fallback_queue()
    test_render_roundtrip_and_safety()
    test_explorer_search_supports_multi_part_scored_queries()
    test_bat_launcher()
    test_vbs_folder_opener()
    test_dwg_links_render_as_links()
    test_transmittal_protocol_accepts_only_a_numeric_order()
    test_note_protocol_records_review_note(Path(tempfile.mkdtemp()))
    test_transmittal_protocol_runs_existing_review_flow_and_logs()
    test_transmittal_protocol_is_single_instance_on_windows()
    test_default_output_path_accepts_folder()
    test_default_output_path_is_coworker_share()
    test_write_explorer_files()
    test_content_fingerprint_ignores_volatile_columns()
    test_board_signature_tracks_column_changes()
    test_maybe_write_only_publishes_on_new_data()
    print("All order_explorer tests passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
