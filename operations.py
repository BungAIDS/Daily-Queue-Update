"""Operation glossary — what each Oper number on the cbcinsider dispatch queue means.

THIS IS THE HELPER DOCUMENT for operations. The AI briefing loads it at every
run and uses it to translate op numbers (19, 20, 200, etc.) into the real
workflow step they represent. Edit this file any time you want to refine what
the AI knows about your operations — no other code changes required.

Workflow overview:

  Custom path
    Op 19  (mgr prep)        ->  Op 20  (custom drafting)
                                       -> Op 22  (double check, extra-custom)
                                       -> Op 23  (customer print sent, fab prints deferred)

  Straight-to-shop path
    Op 21  (mgr prep)        ->  Op 200 (macro drafting, send to shop)

  Unapproved orders
    Op 201 (customer print only, no fabrication prints)
"""
from __future__ import annotations

from typing import Dict

# Keys are the op number as a STRING (matches the scraped field).
OPERATIONS: Dict[str, Dict[str, str]] = {
    "19": {
        "name": "Engineering Manager Prep (custom path)",
        "description": (
            "Engineering manager prepares the order so it's ready for "
            "custom drafting on Op 20. The manager's prep step for the "
            "more-custom workflow."
        ),
    },
    "20": {
        "name": "Custom Drafting",
        "description": (
            "Drafting for the more-custom orders. Comes after Op 19 "
            "(manager prep). May feed into Op 22 (double check) or "
            "Op 23 (customer print sent, fab prints deferred)."
        ),
    },
    "21": {
        "name": "Engineering Manager Prep (straight-to-shop path)",
        "description": (
            "Engineering manager's prep step for Op 200 — the straight-"
            "to-shop equivalent of Op 19. Often absent from the board."
        ),
    },
    "22": {
        "name": "Double Check (extra-custom orders)",
        "description": (
            "Double-check step for the extra-custom orders that came "
            "through Op 20."
        ),
    },
    "23": {
        "name": "Customer Print Sent, Fab Prints Deferred",
        "description": (
            "Used when an Op 20 order needs the customer drawing sent "
            "now, but fabrication (shop) prints are deferred. The 20 -> "
            "23 move means: get the customer their print, postpone fab "
            "prints."
        ),
    },
    "200": {
        "name": "Straight-to-Shop Drafting (macro)",
        "description": (
            "Less custom: runs a macro for a customer drawing and sends "
            "it out. The straight-to-shop drafting path. Comes after Op 21."
        ),
    },
    "201": {
        "name": "Customer Print Only (order unapproved)",
        "description": (
            "Used when an order is unapproved — only the customer print "
            "is produced; no fabrication prints."
        ),
    },
}


# ---------------------------------------------------------------------------
# ROUTING RULES — who should handle an order, by design / operation.
#
# These are BUSINESS rules (who the work belongs to), NOT the scraped
# "Assigned To" field (which the AI ignores). Edit / add rules freely. The
# first matching rule wins. route_owner() is evaluated in code so the answer
# is deterministic — the AI just reports the result.
# ---------------------------------------------------------------------------

def _design_num(design: str):
    """Designs are strings like '95' or 'EMSI'. Return the int, or None."""
    try:
        return int(str(design).strip())
    except (TypeError, ValueError):
        return None


def route_owner(design: str, oper: str) -> str | None:
    """Return the person who should handle this order, or None if no rule matches."""
    d = _design_num(design)
    o = str(oper).strip()

    # Design 95 on Op 19 (custom manager prep) -> Michael.
    if d == 95 and o == "19":
        return "Michael"
    # Design 55 or Design 58 (any operation) -> Michael.
    if d in (55, 58):
        return "Michael"

    return None


# Human-readable version of the same rules, for the AI prompt.
ROUTING_RULES_TEXT = [
    "Design 95 on Op 19 (custom manager prep) -> handled by Michael.",
    "Design 55 or Design 58 (any operation) -> handled by Michael.",
]


def routing_glossary() -> str:
    """Render the routing rules as plaintext for inclusion in the AI prompt."""
    lines = [
        "ROUTING RULES (who should handle an order, by design/operation — these are",
        "business rules, separate from the scraped 'Assigned To' field which you ignore):",
    ]
    for r in ROUTING_RULES_TEXT:
        lines.append(f"  - {r}")
    lines.append(
        "A new/returning order that matches a rule carries a \"handler\" field naming "
        "the owner. Call out the handler when briefing those orders (e.g. \"Michael's\")."
    )
    return "\n".join(lines)


def operations_glossary() -> str:
    """Render the glossary as plaintext for inclusion in the AI prompt."""
    lines = [
        "OPERATION GLOSSARY (translate op numbers into workflow steps when briefing):",
    ]
    for code in sorted(OPERATIONS, key=lambda c: int(c)):
        info = OPERATIONS[code]
        lines.append(f"  Op {code} — {info['name']}: {info['description']}")
    lines.append("")
    lines.append("Workflow paths:")
    lines.append("  Custom path:           Op 19 (mgr prep) -> Op 20 (custom drafting) -> Op 22 (double check) or Op 23 (cust print, fab deferred)")
    lines.append("  Straight-to-shop path: Op 21 (mgr prep) -> Op 200 (macro drafting)")
    lines.append("  Unapproved orders:     Op 201 (customer print only)")
    return "\n".join(lines)
