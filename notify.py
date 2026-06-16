"""New-order notifications for the live watcher: a Windows toast on this PC, and
a Microsoft Teams card so coworkers (and their phones) get pinged too.

Both are best-effort: a notification is a nicety, never worth sinking a poll
cycle, so every function swallows its own errors and logs rather than raising.

- Toast: the little corner pop-up (like a new-Outlook-email alert). Fires only on
  the machine running the watcher. Uses the `winotify` package if installed, else
  falls back to a PowerShell BurntToast/legacy balloon, else logs and moves on.
- Teams: posts to TEAMS_WEBHOOK_URL. Supports BOTH webhook flavors, auto-detected
  from the URL — the new **Workflows** (Power Automate) webhook (Adaptive Card),
  which is Microsoft's replacement for the retiring connector, and the legacy
  **Incoming Webhook** connector (MessageCard). Every member of the channel gets
  a desktop + phone notification, nothing to install. Uses urllib (no extra dep).
"""
from __future__ import annotations

import json
import logging
import urllib.error
import urllib.request
from typing import Any, Dict, List

from config import TEAMS_WEBHOOK_URL, LIVE_TOAST_ENABLED

log = logging.getLogger(__name__)

_TEAMS_THEME = "0076D7"


def _toast_winotify(title: str, message: str) -> bool:
    try:
        from winotify import Notification  # type: ignore
    except ImportError:
        return False
    try:
        Notification(app_id="Daily Queue Watcher", title=title, msg=message).show()
        return True
    except Exception as e:  # noqa: BLE001
        log.warning("winotify toast failed (%s)", e)
        return False


def _toast_powershell(title: str, message: str) -> bool:
    """Fallback toast via PowerShell's tray balloon — present on any Windows box,
    no package needed. Not as pretty as a real toast but reaches the same corner."""
    import subprocess

    safe_title = title.replace("'", "''")
    safe_msg = message.replace("'", "''")
    ps = (
        "Add-Type -AssemblyName System.Windows.Forms;"
        "$n = New-Object System.Windows.Forms.NotifyIcon;"
        "$n.Icon = [System.Drawing.SystemIcons]::Information;"
        "$n.BalloonTipTitle = '" + safe_title + "';"
        "$n.BalloonTipText = '" + safe_msg + "';"
        "$n.Visible = $true; $n.ShowBalloonTip(8000); Start-Sleep -Seconds 9; $n.Dispose()"
    )
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
            check=True, capture_output=True, timeout=20,
        )
        return True
    except Exception as e:  # noqa: BLE001
        log.warning("PowerShell toast failed (%s)", e)
        return False


def toast(title: str, message: str) -> None:
    """Pop a Windows toast on this PC (best-effort)."""
    if not LIVE_TOAST_ENABLED:
        return
    if _toast_winotify(title, message):
        return
    if _toast_powershell(title, message):
        return
    log.info("Toast not shown (no winotify and PowerShell unavailable): %s — %s", title, message)


def _order_facts(j: Dict[str, Any]) -> List[tuple]:
    """The (label, value) rows shown for one order — shared by both card formats."""
    facts = [
        ("Order", str(j.get("job", ""))),
        ("Customer", j.get("customer", "") or "—"),
        ("Design", j.get("design", "") or "—"),
        ("Description", j.get("so_design_desc", "") or "—"),
        ("Added", j.get("_added_label", "") or "—"),
    ]
    flags = []
    if j.get("unapproved"):
        flags.append("UNAPPROVED")
    if j.get("credit_hold"):
        flags.append("CREDIT HOLD")
    if j.get("has_drive_run"):
        flags.append("QUOTE RUN")
    if flags:
        facts.append(("Flags", ", ".join(flags)))
    return facts


def _messagecard(title: str, summary: str, jobs: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Legacy MessageCard — the format the (retiring) Incoming Webhook accepts."""
    sections = [
        {"activityTitle": f"Order {j.get('job', '')}",
         "facts": [{"name": n, "value": v} for n, v in _order_facts(j)]}
        for j in jobs
    ]
    return {
        "@type": "MessageCard",
        "@context": "http://schema.org/extensions",
        "themeColor": _TEAMS_THEME,
        "summary": summary,
        "title": title,
        "sections": sections,
    }


def _workflow_card(title: str, jobs: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Adaptive Card wrapped for the Workflows ('Post to a channel when a webhook
    request is received') trigger — Microsoft's replacement for the connector."""
    body: List[Dict[str, Any]] = [
        {"type": "TextBlock", "text": title, "weight": "Bolder", "size": "Medium", "wrap": True}
    ]
    for j in jobs:
        body.append({"type": "TextBlock", "text": f"Order {j.get('job', '')}",
                     "weight": "Bolder", "wrap": True, "spacing": "Medium"})
        body.append({"type": "FactSet",
                     "facts": [{"title": n, "value": v} for n, v in _order_facts(j)]})
    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.4",
        "body": body,
    }
    return {"type": "message",
            "attachments": [{"contentType": "application/vnd.microsoft.card.adaptive",
                             "content": card}]}


def _is_workflow_url(url: str) -> bool:
    """True for a Workflows/Power Automate webhook URL (Adaptive Card), False for
    a legacy Incoming Webhook connector URL (MessageCard). Covers both the
    logic.azure.com and the newer Power Platform (powerplatform.com) hosts."""
    u = (url or "").lower()
    return ("logic.azure.com" in u or "powerautomate" in u or "powerplatform" in u
            or "/workflows/" in u or "azure-apihub" in u)


def teams_post(title: str, summary: str, jobs: List[Dict[str, Any]]) -> None:
    """Post a new-orders card to the configured Teams webhook, in whichever format
    that webhook expects (Workflows Adaptive Card vs legacy MessageCard)."""
    if not TEAMS_WEBHOOK_URL:
        return
    workflow = _is_workflow_url(TEAMS_WEBHOOK_URL)
    card = _workflow_card(title, jobs) if workflow else _messagecard(title, summary, jobs)
    req = urllib.request.Request(
        TEAMS_WEBHOOK_URL, data=json.dumps(card).encode("utf-8"),
        headers={"Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            status = getattr(resp, "status", None) or resp.getcode()
            reply = resp.read().decode("utf-8", "replace").strip()
        # Workflows returns 202 with an empty body; the legacy connector returns
        # 200 with "1". Treat any 2xx as success; only the legacy "not 1" is odd.
        if not workflow and reply and reply != "1":
            log.warning("Teams webhook returned: %s", reply[:200])
        else:
            log.debug("Teams post ok (HTTP %s, %s)", status, "workflow" if workflow else "connector")
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "replace")[:200] if hasattr(e, "read") else ""
        log.warning("Teams notification failed (HTTP %s) %s", e.code, detail)
    except Exception as e:  # noqa: BLE001
        log.warning("Teams notification failed (%s)", e)


def notify_new_orders(jobs: List[Dict[str, Any]]) -> None:
    """Fire both channels for the orders that just appeared this cycle."""
    if not jobs:
        return
    n = len(jobs)
    if n == 1:
        j = jobs[0]
        title = f"New order {j.get('job', '')}"
        line = " — ".join(
            p for p in (j.get("customer", ""), j.get("so_design_desc") or j.get("design", "")) if p
        )
        toast(title, line or "New order on the queue")
    else:
        title = f"{n} new orders"
        toast(title, ", ".join(str(j.get("job", "")) for j in jobs))

    teams_post(
        title=f"{n} new order{'s' if n != 1 else ''} on the queue",
        summary=f"{n} new order(s)",
        jobs=jobs,
    )
