"""Claude API analysis of queue changes.

Sends today's structured diff to Claude and asks for:
  1) A natural-language briefing paragraph
  2) Anomaly flags (high $, no assignee, unrealistic dates, lingering jobs)
  3) Ranked top 3-5 action items
"""
from __future__ import annotations

import json
import logging
from datetime import date
from typing import Any, Dict

import anthropic

from config import ANTHROPIC_API_KEY, CLAUDE_MODEL

log = logging.getLogger(__name__)

SYSTEM_PROMPT = """You are an operations analyst preparing a daily briefing for an engineering team lead.
You will receive a structured JSON snapshot of yesterday vs today's work queue, with:
  - new orders (in today, not yesterday)
  - removed orders (in yesterday, not today — likely completed)
  - changed orders (field-level diffs)
  - persistent orders (3+ consecutive days in the queue)
  - aggregate counts

Output STRICT JSON only, no prose outside the JSON, matching this schema:
{
  "briefing": "3-5 sentence natural-language summary of the day's queue. Mention notable patterns, customers with multiple jobs, dollar totals, schedule slips. Conversational but specific.",
  "anomalies": ["Short bullet strings flagging things worth investigating: unusually high $ values, jobs with no Assigned To, End Dates that seem unrealistic relative to Start, jobs lingering 5+ days, etc."],
  "action_items": [
    {"rank": 1, "job": "######", "reason": "Why this needs attention today"},
    ...
  ]
}
Give 3-5 action items, ranked by urgency. Be concrete. If the queue is quiet, say so honestly rather than padding."""


def analyze(diff: Dict[str, Any], today: date) -> Dict[str, Any]:
    """Call Claude on the diff. Returns the parsed analysis dict.

    Raises on API or parsing errors so the caller can send an alert email.
    """
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # Trim the payload so we don't blow tokens on huge fields
    payload = {
        "date": today.isoformat(),
        "summary": {
            "today_job_count": diff["today_count"],
            "yesterday_job_count": diff["yesterday_count"],
            "new_count": len(diff["new"]),
            "removed_count": len(diff["removed"]),
            "changed_count": len(diff["changed"]),
            "persistent_count": len(diff["persistent"]),
        },
        "new_orders": diff["new"],
        "removed_orders": diff["removed"],
        "changed_orders": diff["changed"],
        "persistent_orders": [
            {"job": p["job"], "customer": p["customer"], "days_in_queue": p["days"],
             "end_date": p["snapshot"].get("end_date"), "assigned_to": p["snapshot"].get("assigned_to"),
             "total_price": p["snapshot"].get("total_price")}
            for p in diff["persistent"]
        ],
    }

    log.info("Calling Claude (%s) with %d new, %d removed, %d changed, %d persistent jobs",
             CLAUDE_MODEL, len(diff["new"]), len(diff["removed"]), len(diff["changed"]), len(diff["persistent"]))

    response = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=2048,
        thinking={"type": "adaptive"},
        system=SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": f"Here is today's queue diff. Produce the briefing JSON.\n\n{json.dumps(payload, indent=2)}",
        }],
    )

    # Pull the text out — adaptive thinking puts ThinkingBlocks before the TextBlock
    text = next((b.text for b in response.content if b.type == "text"), "")
    if not text:
        raise RuntimeError("Claude returned no text content")

    # Strip ```json fences if Claude added them
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = cleaned.strip("`")
        if cleaned.startswith("json"):
            cleaned = cleaned[4:]
        cleaned = cleaned.strip()
        if cleaned.endswith("```"):
            cleaned = cleaned[:-3].strip()

    try:
        return json.loads(cleaned)
    except json.JSONDecodeError as e:
        log.error("Claude response was not valid JSON. Raw text:\n%s", text)
        raise RuntimeError(f"Failed to parse Claude JSON output: {e}")
