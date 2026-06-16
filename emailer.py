"""Daily briefing via the local Outlook desktop app — no password stored.

Sends through your already-signed-in Outlook using COM automation, so the
script never needs your email password. Windows + Outlook desktop only, and it
must run in your logged-in desktop session (see README scheduling notes).
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict

from datetime import date

from config import (EMAIL_TO, EMAIL_ALERT_TO, EMAIL_ATTACH_REPORT,
                   LIVE_WORKBOOK_LINK)

log = logging.getLogger(__name__)

_OL_MAIL_ITEM = 0  # olMailItem


def _send(to_addr: str, subject: str, body: str, attachment: Path | None = None) -> None:
    # Imported lazily so the module still imports on non-Windows machines.
    import win32com.client

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(_OL_MAIL_ITEM)
    mail.To = to_addr
    mail.Subject = subject
    mail.Body = body
    if attachment is not None and attachment.exists():
        mail.Attachments.Add(str(attachment.resolve()))
    mail.Send()
    log.info("Sent email via Outlook to %s: %s", to_addr, subject)


def _pretty_date(iso: str) -> str:
    """'2026-06-16' -> 'Tuesday, June 16, 2026'; pass anything else through."""
    try:
        return date.fromisoformat(iso).strftime("%A, %B %d, %Y")
    except (ValueError, TypeError):
        return iso


def send_daily_briefing(
    briefing: Dict[str, Any],
    diff: Dict[str, Any],
    excel_path: Path,
    today_str: str,
) -> None:
    today_pretty = _pretty_date(today_str)
    prev_pretty = _pretty_date(diff.get("prev_date", "")) if diff.get("prev_date") else ""

    lines = [f"Queue briefing for {today_pretty}.", ""]
    # The live sheet is the headline: an always-current, co-authored workbook.
    if LIVE_WORKBOOK_LINK:
        lines.append(f"LIVE queue (updates all day): {LIVE_WORKBOOK_LINK}")
        lines.append("")
    lines.append(briefing.get("briefing", "(no briefing available)"))
    lines.append("")
    vs = f" (vs {prev_pretty})" if prev_pretty else ""
    lines.append(
        f"Changes{vs}:  new={len(diff['new'])}  returning={len(diff.get('returning', []))}  "
        f"removed={len(diff['removed'])}  changed={len(diff['changed'])}  "
        f"lingering(3+ days)={len(diff['persistent'])}"
    )
    lines.append("")

    items = briefing.get("action_items", []) or []
    if items:
        lines.append("Top action items:")
        for item in items:
            rank = item.get("rank", "?")
            job = item.get("job", "?")
            reason = item.get("reason", "")
            lines.append(f"  {rank}. Job {job} — {reason}")
        lines.append("")

    anomalies = briefing.get("anomalies", []) or []
    if anomalies:
        lines.append("Anomalies flagged:")
        for a in anomalies:
            lines.append(f"  - {a}")
        lines.append("")

    # The dated .xlsx is always saved to disk for the archive; attach it only
    # when configured (default: off once a live link is set — the link is the
    # point), but always note where the archived copy lives.
    attachment = excel_path if EMAIL_ATTACH_REPORT else None
    label = "attached; archived at" if EMAIL_ATTACH_REPORT else "archived at"
    lines.append(f"Dated report ({label}): {excel_path}")

    _send(EMAIL_TO, f"Daily Queue Briefing — {today_pretty}", "\n".join(lines),
          attachment=attachment)


def send_alert(subject_suffix: str, error_text: str) -> None:
    """Best-effort failure alert. Swallows its own errors so we don't loop."""
    try:
        _send(
            EMAIL_ALERT_TO,
            f"ALERT: Daily Queue script failed — {subject_suffix}",
            f"The daily queue script failed:\n\n{error_text}\n",
        )
    except Exception as e:
        log.exception("Failed to send alert email: %s", e)
