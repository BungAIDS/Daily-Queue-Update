"""New-order notifications for the live watcher: a Windows toast on this PC, and
a Microsoft Teams card so coworkers (and their phones) get pinged too.

Both are best-effort: a notification is a nicety, never worth sinking a poll
cycle, so every function swallows its own errors and logs rather than raising.

- Toast: the little corner pop-up (like a new-Outlook-email alert). Fires only on
  the machine running the watcher. Uses the `winotify` package if installed, else
  falls back to a PowerShell BurntToast/legacy balloon, else logs and moves on.
- Teams: posts a MessageCard to an Incoming Webhook URL (TEAMS_WEBHOOK_URL).
  Every member of that channel gets a desktop + phone notification, with nothing
  to install on their machines. Uses urllib so there's no extra dependency.
"""
from __future__ import annotations

import json
import logging
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


def _teams_card(title: str, summary: str, jobs: List[Dict[str, Any]]) -> Dict[str, Any]:
    """A legacy MessageCard (the format Incoming Webhooks accept) with one
    'section' of facts per new order."""
    sections = []
    for j in jobs:
        facts = [
            {"name": "Order", "value": str(j.get("job", ""))},
            {"name": "Customer", "value": j.get("customer", "") or "—"},
            {"name": "Design", "value": j.get("design", "") or "—"},
            {"name": "Description", "value": j.get("so_design_desc", "") or "—"},
            {"name": "Added", "value": j.get("_added_label", "") or "—"},
        ]
        flags = []
        if j.get("unapproved"):
            flags.append("UNAPPROVED")
        if j.get("credit_hold"):
            flags.append("CREDIT HOLD")
        if j.get("has_drive_run"):
            flags.append("QUOTE RUN")
        if flags:
            facts.append({"name": "Flags", "value": ", ".join(flags)})
        sections.append({"activityTitle": f"Order {j.get('job', '')}", "facts": facts})
    return {
        "@type": "MessageCard",
        "@context": "http://schema.org/extensions",
        "themeColor": _TEAMS_THEME,
        "summary": summary,
        "title": title,
        "sections": sections,
    }


def teams_post(title: str, summary: str, jobs: List[Dict[str, Any]]) -> None:
    """Post a new-orders card to the configured Teams Incoming Webhook."""
    if not TEAMS_WEBHOOK_URL:
        return
    payload = json.dumps(_teams_card(title, summary, jobs)).encode("utf-8")
    req = urllib.request.Request(
        TEAMS_WEBHOOK_URL, data=payload, headers={"Content-Type": "application/json"}
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            body = resp.read().decode("utf-8", "replace").strip()
            # Teams replies "1" on success; anything else is worth logging.
            if body and body != "1":
                log.warning("Teams webhook returned: %s", body[:200])
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
