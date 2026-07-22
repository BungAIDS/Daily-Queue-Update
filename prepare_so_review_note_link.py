"""Handle a GL Queue Explorer ``glqnote:`` review-note link.

The Explorer is a static file page, so this per-user protocol handler bridges
an explicit browser click to the existing ``so_review_notes.json`` queue.  The
payload is URL query data only; it is validated and passed to ``so_review`` as
data, never executed as a command.
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import so_review

PUSH_ON_SAVE = (
    os.environ.get("SO_REVIEW_NOTE_PUSH_ON_SAVE", "1").strip().lower()
    not in {"0", "false", "no", "off"}
)


ROOT = Path(__file__).resolve().parent
SCHEME = "glqnote"


def console_python_path(root: Path = ROOT, executable: str | Path | None = None) -> Path:
    """Prefer this checkout's console Python, never a windowless pythonw.exe."""
    for candidate in (
        root / "venv" / "Scripts" / "python.exe",
        root / ".venv" / "Scripts" / "python.exe",
    ):
        if candidate.exists():
            return candidate
    current = Path(executable or sys.executable)
    if current.name.lower() == "pythonw.exe":
        console = current.with_name("python.exe")
        if console.exists():
            return console
    return current


def protocol_command(
    root: Path = ROOT,
    executable: str | Path | None = None,
) -> str:
    """Windows registry command line for this handler; ``%1`` is the URI."""
    return subprocess.list2cmdline([
        str(console_python_path(root, executable)),
        str(root / Path(__file__).name),
        "%1",
    ])


def parse_note_uri(uri: str) -> dict[str, str]:
    """Return a validated note payload from ``glqnote:?order=...``."""
    value = str(uri).strip()
    if value[:len(SCHEME) + 1].lower() != f"{SCHEME}:":
        raise ValueError(f"expected a {SCHEME}: link")
    parsed = urlparse(value)
    params = parse_qs(parsed.query, keep_blank_values=True)
    one = lambda key: (params.get(key) or [""])[0].strip()
    order = one("order")
    action = (one("action") or "add").casefold()
    note = one("note")
    note_id = one("note_id")
    if not order.isdigit():
        raise ValueError("order must contain digits only")
    if action not in {"add", "delete"}:
        raise ValueError("unsupported note action")
    if action == "add":
        if not note:
            raise ValueError("note text is required")
        if len(note) > 4000:
            raise ValueError("note text is too long")
    elif not note_id.isdigit():
        raise ValueError("delete requires a numeric note_id")
    return {
        "action": action,
        "order": order,
        "note_id": note_id,
        "item_no": one("item_no"),
        "item_text": one("item_text")[:1000],
        "note": note,
        "row_key": one("row_key")[:500],
    }


def publish_review_notes() -> bool:
    """Best-effort immediate publish of the updated SO review note store."""
    if not PUSH_ON_SAVE:
        return False
    try:
        import data_push
        return data_push.push_data()
    except Exception:  # noqa: BLE001 - note saving must not fail because git did
        return False


def delete_note(store: dict, order: str, note_id: str) -> bool:
    """Remove one open/pending review note by id and order."""
    before = len(store.get("notes") or [])
    store["notes"] = [
        n for n in (store.get("notes") or [])
        if not (str(n.get("id", "")) == str(note_id)
                and str(n.get("order", "")) == str(order))
    ]
    return len(store["notes"]) != before


def record_from_uri(uri: str) -> tuple[bool, bool]:
    payload = parse_note_uri(uri)
    store = so_review.load_store()
    if payload["action"] == "delete":
        changed = delete_note(store, payload["order"], payload["note_id"])
    else:
        changed = bool(so_review.record_note(
            store,
            payload["order"],
            payload["item_no"],
            payload["item_text"],
            payload["note"],
            row_key=payload["row_key"],
        ))
    pushed = False
    if changed:
        so_review.save_store(store)
        pushed = publish_review_notes()
    return changed, pushed


def main(argv: list[str] | None = None) -> int:
    args = list(argv if argv is not None else sys.argv[1:])
    if not args:
        print(f"Usage: python {Path(__file__).name} '{SCHEME}:?order=...&note=...'")
        return 2
    try:
        added, pushed = record_from_uri(args[0])
    except Exception as exc:
        print(f"Could not record note: {exc}")
        return 1
    if added:
        print("Updated note queue and pushed order-data branch." if pushed
              else "Updated note queue. Git data push was skipped or failed.")
    else:
        print("Note already exists.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
