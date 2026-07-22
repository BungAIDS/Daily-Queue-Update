"""Handle a GL Queue Explorer ``glqnote:`` review-note link.

The Explorer is a static file page, so this per-user protocol handler bridges
an explicit browser click to the existing ``so_review_notes.json`` queue.  The
payload is URL query data only; it is validated and passed to ``so_review`` as
data, never executed as a command.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import so_review


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
    note = one("note")
    if not order.isdigit():
        raise ValueError("order must contain digits only")
    if not note:
        raise ValueError("note text is required")
    if len(note) > 4000:
        raise ValueError("note text is too long")
    return {
        "order": order,
        "item_no": one("item_no"),
        "item_text": one("item_text")[:1000],
        "note": note,
        "row_key": one("row_key")[:500],
    }


def record_from_uri(uri: str) -> bool:
    payload = parse_note_uri(uri)
    store = so_review.load_store()
    added = so_review.record_note(
        store,
        payload["order"],
        payload["item_no"],
        payload["item_text"],
        payload["note"],
        row_key=payload["row_key"],
    )
    if added:
        so_review.save_store(store)
    return bool(added)


def main(argv: list[str] | None = None) -> int:
    args = list(argv if argv is not None else sys.argv[1:])
    if not args:
        print(f"Usage: python {Path(__file__).name} '{SCHEME}:?order=...&note=...'")
        return 2
    try:
        added = record_from_uri(args[0])
    except Exception as exc:
        print(f"Could not record note: {exc}")
        return 1
    print("Recorded note." if added else "Note already exists.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
