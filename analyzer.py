"""Claude API analysis of queue changes.

The briefing is intentionally scoped: it talks ONLY about newly-appeared
orders (new + returning), and reasons only over a whitelist of fields. Price,
End Date, Assigned To, Checker, status notes, and approval/credit flags are
deliberately withheld from the model — FanNet date is the only timing signal.
"""
from __future__ import annotations

import json
import logging
from datetime import date
from typing import Any, Dict

import anthropic

from config import ANTHROPIC_API_KEY, CLAUDE_MODEL

log = logging.getLogger(__name__)

# The ONLY fields the AI is allowed to see. Everything else (total_price,
# end_date, assigned_to, checker, status_note, unapproved, credit_hold,
# has_notes) is withheld on purpose so the briefing can't reference it.
AI_FIELDS = [
    "job", "status", "customer", "primary_rep", "item", "design",
    "oper", "start_date", "fannet_date", "plan_hrs", "ship_with",
]


def _trim(job: Dict[str, Any]) -> Dict[str, Any]:
    """Keep only the whitelisted fields the AI is allowed to reason about."""
    out = {k: job.get(k, "") for k in AI_FIELDS}
    if job.get("_last_seen"):
        out["last_seen"] = job["_last_seen"]
    return out


SYSTEM_PROMPT = """You are an operations analyst preparing a daily briefing for an engineering team lead.

Each day you receive the orders that have just appeared on the engineering work queue:
  - new orders (never seen before)
  - returning orders (back after previously dropping off; "last_seen" = when they were last here)
plus aggregate counts for context.

Your briefing is ONLY about what is newly on the board today (new + returning). Do not editorialize about the rest of the queue.

For each order you may use ONLY these fields: job, status, customer, primary_rep (rep), item, design, oper (operation), start_date, fannet_date, plan_hrs (planned hours), ship_with.

You must IGNORE and never mention: total price / dollar values, End Date (use FanNet date as the only date/deadline), Assigned To, Checker, status notes, and any approval / credit-hold / notes flags. These fields are intentionally not provided.

Use FanNet date as the timing signal (e.g. which new orders have the soonest FanNet dates).

Output STRICT JSON only, no prose outside the JSON, matching this schema:
{
  "briefing": "3-5 sentence summary of what is NEW on the board today: how many new/returning orders, which customers and reps, which designs, notable FanNet timing, and any ship-with groupings. Conversational but specific. If nothing is new, say so plainly rather than padding.",
  "anomalies": ["Short bullets about the NEW/returning orders worth a look: soonest FanNet dates, the same customer or design showing up on multiple new orders, ship-with links, or possible duplicate new orders (same customer + design + FanNet). Use only the allowed fields."],
  "action_items": [
    {"rank": 1, "job": "######", "reason": "Why this new order needs attention, framed by FanNet timing / customer / design / ship-with"},
    ...
  ]
}
Give up to 5 action items drawn only from the new/returning orders, ranked by FanNet urgency. If there are no new orders, return an empty action_items list and say the board is quiet."""


def analyze(diff: Dict[str, Any], today: date) -> Dict[str, Any]:
    """Call Claude on the diff. Returns the parsed analysis dict.

    Raises on API or parsing errors so the caller can send an alert email.
    """
    if not ANTHROPIC_API_KEY:
        raise RuntimeError("ANTHROPIC_API_KEY is not set (check your .env).")
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # Only the newly-appeared orders are sent, trimmed to the allowed fields.
    payload = {
        "date": today.isoformat(),
        "summary": {
            "today_job_count": diff["today_count"],
            "new_count": len(diff["new"]),
            "returning_count": len(diff.get("returning", [])),
        },
        "new_orders": [_trim(j) for j in diff["new"]],
        "returning_orders": [_trim(j) for j in diff.get("returning", [])],
    }

    log.info("Calling Claude (%s) with %d new, %d returning orders",
             CLAUDE_MODEL, len(diff["new"]), len(diff.get("returning", [])))

    # claude-opus-4-7 only supports adaptive thinking. Thinking tokens count
    # against max_tokens, and at the default "high" effort the model almost
    # always thinks — so max_tokens must be generous or the JSON answer gets
    # truncated mid-stream. "medium" effort keeps reasoning proportionate to
    # this small task while leaving plenty of headroom for the output.
    response = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=8000,
        thinking={"type": "adaptive"},
        output_config={"effort": "medium"},
        system=SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": f"Here is today's queue diff. Produce the briefing JSON.\n\n{json.dumps(payload, indent=2)}",
        }],
    )

    if getattr(response, "stop_reason", None) == "max_tokens":
        raise RuntimeError(
            "Claude hit the max_tokens limit before finishing — the JSON is "
            "likely truncated. Increase max_tokens in analyzer.py."
        )

    # Pull the text out — adaptive mode returns thinking block(s) before the text
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
