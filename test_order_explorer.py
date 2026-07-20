"""Tests for order_explorer.py — the GL Queue Explorer page generator.

Run:  python test_order_explorer.py
Pure logic only (payload building, HTML embedding, the .bat launcher); no
browser, no COM, no network — safe for CI.
"""
from __future__ import annotations

import base64
import gzip
import json
import re
import tempfile
from pathlib import Path

import order_explorer as oe


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
    ]
    from datetime import date as _date
    p = oe.build_payload(_store(), master_orders=master_orders, events=events,
                         today=_date(2026, 7, 17))
    assert p["today"] == "2026-07-17"
    # Newest first; the CO event carries its co_history description.
    assert [e["f"] for e in p["ev"]] == ["End Date", "CO#"], p["ev"]
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
    assert "click order # again to move it to the left" in html
    assert "red = scored construction difference" in html
    assert "preview-relevant" in html
    assert "preview-match" in html
    assert "green = selected combination match" in html
    assert "combinedFocusedSimilarity" in html
    assert "state.selections" in html
    assert "Required combination" in html
    assert "selected components" in html
    assert "alignComponentStarts" in html
    assert 'class="co-tip"' in html
    assert "No change-order description was captured." in html
    assert 'id="leftcomponenttree"' in html
    assert 'id="rightcomponenttree"' in html
    assert "ontoggle = alignComponentStarts" in html
    assert "<title>GL Queue Explorer</title>" in html
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
    print("  vbs folder opener OK")


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


def test_write_explorer_files():
    p = oe.build_payload(_store(), _dwg())
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "GL Queue Explorer.html"
        got = oe.write_explorer(p, out)
        assert got == out and out.exists() and out.stat().st_size > 10_000
        bat = Path(td) / oe.BAT_NAME
        assert bat.exists(), "launcher .bat not written"
        assert (Path(td) / oe.VBS_NAME).exists(), "glq_open.vbs not written"
        # The auto-refresh stamp open pages poll, matching the page's gen.
        ver = Path(td) / oe.VERSION_NAME
        assert ver.exists(), "version stamp not written"
        assert p["gen"] in ver.read_text(encoding="utf-8")
        assert "__GLQ_VERSION__" in ver.read_text(encoding="utf-8")
        first = bat.read_bytes()
        oe.write_explorer(p, out)                        # idempotent second write
        assert bat.read_bytes() == first
        assert not list(Path(td).glob("*.tmp")), "temp file left behind"
    print("  write_explorer files OK")


def main() -> int:
    test_payload_components_merge()
    test_queue_jobs_take_master_items()
    test_events_and_removed()
    test_spec_values_tidied()
    test_quote_run_and_solidworks_links()
    test_wheel_component_from_quote_run()
    test_master_orders_fallback_queue()
    test_render_roundtrip_and_safety()
    test_bat_launcher()
    test_vbs_folder_opener()
    test_default_output_path_accepts_folder()
    test_write_explorer_files()
    print("All order_explorer tests passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
