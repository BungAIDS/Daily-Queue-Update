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
from operations import operations_glossary, route_owner, routing_glossary

log = logging.getLogger(__name__)

# The ONLY fields the AI is allowed to see. Everything else (total_price,
# assigned_to, checker, status_note, unapproved, credit_hold, has_notes) is
# withheld on purpose so the briefing can't reference it.
AI_FIELDS = [
    "job", "status", "customer", "primary_rep", "item", "design",
    "oper", "start_date", "end_date", "fannet_date", "plan_hrs", "ship_with",
]


def _trim(job: Dict[str, Any]) -> Dict[str, Any]:
    """Keep only the whitelisted fields the AI is allowed to reason about."""
    out = {k: job.get(k, "") for k in AI_FIELDS}
    if job.get("_last_seen"):
        out["last_seen"] = job["_last_seen"]
    handler = route_owner(job.get("design", ""), job.get("oper", ""))
    if handler:
        out["handler"] = handler
    return out


SYSTEM_PROMPT = """You are an operations analyst preparing a daily briefing for an engineering team lead.

DOMAIN CONTEXT
__OPERATIONS_GLOSSARY__

__ROUTING_GLOSSARY__

When you reference an operation in the briefing, translate the number into its workflow step (e.g. "Op 200 (straight-to-shop drafting)"). Operation is a first-class grouping signal: spotting which operations are loaded heaviest, where bottlenecks may form on a workflow path, and which new orders enter which step are all valuable.

Each day you receive:
  - new orders (never seen before)
  - returning orders (back after previously dropping off; "last_seen" = when they were last here)
  - full_queue_for_context: EVERY order currently on the board (same fields), provided so you can spot groupings — do not summarize the full queue itself.
plus aggregate counts.

Your briefing is ONLY about what is newly on the board today (new + returning). Do not editorialize about the rest of the queue. BUT use full_queue_for_context to judge whether each new order is part of something bigger: a new order may join an existing cluster of the same design, operation, or customer, or its ship-with partner may already be on the board even if that partner is not new. Call those connections out.

For each order you may use ONLY these fields: job, status, customer, primary_rep (rep), item, design, oper (operation — this is a first-class grouping signal), start_date, end_date, fannet_date, plan_hrs (planned hours), ship_with.

You must IGNORE and never mention: total price / dollar values, Assigned To, Checker, status notes, and any approval / credit-hold / notes flags. These fields are intentionally not provided.

TIMING — two dates matter, weigh them together:
  - End Date is the engineering commitment / target date. Respect it: an order at or past its End Date is slipping and worth flagging.
  - FanNet date is when the product is actually needed downstream.
  - Relative urgency: if an order's FanNet date is much further out than its End Date, a late or approaching End Date is LESS urgent — there is downstream slack. If the End Date and FanNet date are both near, the order is genuinely time-critical. Rank by this combined judgment, and when you flag a late End Date, mention how far out its FanNet date is so the reader can gauge real urgency.

Output STRICT JSON only, no prose outside the JSON, matching this schema:
{
  "briefing": "3-5 sentence summary of what is NEW on the board today: how many new/returning orders, which customers and reps, which designs and operations, notable End Date / FanNet timing (and any orders already past their End Date), and any ship-with groupings. When a new order joins an existing design / operation / customer cluster on the board, note how many total there are now. Conversational but specific. If nothing is new, say so plainly rather than padding.",
  "anomalies": ["Short bullets about the NEW/returning orders worth a look: orders at or past their End Date (note how far out their FanNet date is); soonest deadlines; a new order that joins an existing cluster of the same design, operation, or customer (say how many total are now on the board); a new order whose ship_with partner is already on the board; or possible duplicate new orders (same customer + design + oper + FanNet). Use only the allowed fields."],
  "action_items": [
    {"rank": 1, "job": "######", "reason": "Why this new order needs attention, framed by End Date vs FanNet urgency / customer / design / operation / ship-with (existing partners on the board are fair game as context)"},
    ...
  ]
}
Give up to 5 action items drawn only from the new/returning orders, ranked by combined End Date + FanNet urgency. If there are no new orders, return an empty action_items list and say the board is quiet."""

SYSTEM_PROMPT = SYSTEM_PROMPT.replace("__OPERATIONS_GLOSSARY__", operations_glossary())
SYSTEM_PROMPT = SYSTEM_PROMPT.replace("__ROUTING_GLOSSARY__", routing_glossary())


def analyze(diff: Dict[str, Any], today: date, all_jobs: list | None = None) -> Dict[str, Any]:
    """Call Claude on the diff. Returns the parsed analysis dict.

    all_jobs is the full current queue, sent (trimmed) as grouping context so
    the briefing can see when a new order joins an existing cluster or has a
    ship-with partner already on the board. Raises on API/parse errors.
    """
    if not ANTHROPIC_API_KEY:
        raise RuntimeError("ANTHROPIC_API_KEY is not set (check your .env).")
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # Only the newly-appeared orders are the subject, but the full board is sent
    # as context. All of it is trimmed to the allowed fields.
    payload = {
        "date": today.isoformat(),
        "summary": {
            "today_job_count": diff["today_count"],
            "new_count": len(diff["new"]),
            "returning_count": len(diff.get("returning", [])),
        },
        "new_orders": [_trim(j) for j in diff["new"]],
        "returning_orders": [_trim(j) for j in diff.get("returning", [])],
        "full_queue_for_context": [_trim(j) for j in (all_jobs or [])],
    }

    log.info("Calling Claude (%s) with %d new, %d returning orders (%d on board for context)",
             CLAUDE_MODEL, len(diff["new"]), len(diff.get("returning", [])), len(all_jobs or []))

    # Pick a thinking mode by model. Opus 4.x only supports adaptive thinking
    # (manual/disabled get rejected), and on this small structured-JSON task
    # the extra reasoning earns its keep. Sonnet doesn't need thinking here —
    # adaptive on Sonnet inflated cost to ~$0.15/run when disabled would have
    # been ~$0.02. Pick automatically based on CLAUDE_MODEL.
    if CLAUDE_MODEL.startswith("claude-opus-4"):
        kwargs = dict(
            max_tokens=8000,
            thinking={"type": "adaptive"},
            output_config={"effort": "medium"},
        )
    else:
        kwargs = dict(
            max_tokens=4000,
            thinking={"type": "disabled"},
        )

    response = client.messages.create(
        model=CLAUDE_MODEL,
        system=SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": f"Here is today's queue diff. Produce the briefing JSON.\n\n{json.dumps(payload, indent=2)}",
        }],
        **kwargs,
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
