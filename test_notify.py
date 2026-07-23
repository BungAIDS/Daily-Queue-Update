"""Tests for new-order notification payloads.

Run:  python test_notify.py
"""
from __future__ import annotations

import notify


def test_order_url_deep_links_to_explorer():
    url = notify._order_url("421966")
    assert url.startswith("file:"), url
    assert "GL%20Queue%20Explorer.html" in url, url
    assert url.endswith("#order=421966"), url
    print("  Explorer order deep link OK")


def test_workflow_card_has_open_order_action():
    card = notify._workflow_card("New order", "summary", [{"job": "421966"}])
    body = card["attachments"][0]["content"]["body"]
    actions = [part for part in body if part.get("type") == "ActionSet"]
    assert actions, card
    action = actions[0]["actions"][0]
    assert action["type"] == "Action.OpenUrl"
    assert action["title"] == "Open order in Explorer"
    assert action["url"].endswith("#order=421966")
    assert "Open order in Explorer" in card["text"]
    print("  Teams workflow action OK")


def test_messagecard_has_open_order_action():
    card = notify._messagecard("New order", "summary", [{"job": "421966"}])
    action = card["sections"][0]["potentialAction"][0]
    assert action["@type"] == "OpenUri"
    assert action["name"] == "Open order in Explorer"
    assert action["targets"][0]["uri"].endswith("#order=421966")
    print("  Teams MessageCard action OK")


def main() -> int:
    test_order_url_deep_links_to_explorer()
    test_workflow_card_has_open_order_action()
    test_messagecard_has_open_order_action()
    print("All notify tests passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
