"""Plain-text email notifications via SMTP."""
from __future__ import annotations

import logging
import smtplib
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Dict

from config import (
    SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD,
    EMAIL_FROM, EMAIL_TO, EMAIL_ALERT_TO,
)

log = logging.getLogger(__name__)


def _send(to_addr: str, subject: str, body: str) -> None:
    msg = EmailMessage()
    msg["From"] = EMAIL_FROM
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.set_content(body)

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.starttls()
        s.login(SMTP_USER, SMTP_PASSWORD)
        s.send_message(msg)
    log.info("Sent email to %s: %s", to_addr, subject)


def send_daily_briefing(
    briefing: Dict[str, Any],
    diff: Dict[str, Any],
    excel_path: Path,
    today_str: str,
) -> None:
    lines = []
    lines.append(briefing.get("briefing", "(no briefing available)"))
    lines.append("")
    lines.append(
        f"Counts:  new={len(diff['new'])}  removed={len(diff['removed'])}  "
        f"changed={len(diff['changed'])}  lingering(3+ days)={len(diff['persistent'])}"
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

    lines.append(f"Excel report saved to: {excel_path}")

    _send(EMAIL_TO, f"Daily Queue Briefing — {today_str}", "\n".join(lines))


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
