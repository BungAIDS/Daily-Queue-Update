"""Fill the Drawing Transmittal Word document from a TransmittalData.

This is the writing half of step 2 of *Completing Transmittals*. It takes the
data gathered by `transmittal_data.py` and stamps a copy of the master template
(`DWG TRANSMITTAL MASTER.doc`): the TO emails, date, order #, subject (customer),
P.O. #, the one approval/record checkbox, the drawing-table rows, and the
engineer's signed initials ("BY: __").

The template's checkboxes are real Word FORMCHECKBOX form fields, so filling is
done through the **desktop Word app over COM** — the same no-password automation
this project already uses for Outlook (emailer.py) and Excel (live_excel.py).
That makes it Windows + Word only, like those.

The planning step (`plan_fill`) is pure and unit-tested; the COM step
(`apply_fill_word`) is a thin applier so the mapping is verifiable without Word.

The output is numbered from what's already gone out: the suffix on
`<order> DWG TRANSMITTAL-NN.doc` is one past the highest **sent** transmittal,
read from the saved Outlook `.msg` emails in the job's TRANSMITTAL folder (a
saved `.msg` proves it was actually mailed; a `.doc` alone is just a draft).
After writing, the finished doc is opened in Word so it can be eyeballed.

    python transmittal_doc.py 421693                 # gather + fill, then open it
    python transmittal_doc.py 421693 --initials DAG  # override the signature
    python transmittal_doc.py 421693 --no-open       # write it but don't open Word
"""
from __future__ import annotations

import argparse
import logging
import os
import re
import sys
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import List, Optional, Tuple

import engineers
import transmittal_data as td

log = logging.getLogger(__name__)

# The master template lives in the repo root (uploaded alongside the code).
TEMPLATE_PATH = Path(__file__).resolve().parent / "DWG TRANSMITTAL MASTER.doc"

# The three FORMCHECKBOX fields in document order map to these boxes.
BOX_ORDER = ["sales", "approval", "record"]

# Each box's identifying wording in its template paragraph ("For sales purposes
# only…", "For approval only…", "For record only…"). Used to tick the right
# FORMCHECKBOX by its label rather than trusting position alone.
BOX_LABEL_RES = {
    "sales": re.compile(r"sales\s+purposes", re.IGNORECASE),
    "approval": re.compile(r"approval\s+only", re.IGNORECASE),
    "record": re.compile(r"record\s+only", re.IGNORECASE),
}

# Word COM constant: a checkbox form field (wdFieldFormCheckBox).
_WD_FORM_CHECKBOX = 71


@dataclass
class FillPlan:
    """Everything the Word applier writes — semantic, not Word-specific, so it
    can be asserted in tests."""
    order: str
    subject: str                       # the customer, goes on the SUBJECT line
    po: str
    date: str                          # MM/DD/YYYY
    initials: str
    box_index: int                     # index into BOX_ORDER (0..2)
    to_emails: List[str] = field(default_factory=list)
    rows: List[Tuple[str, str, str, str, str, str]] = field(default_factory=list)
    # rows: (EMAIL, PRINT, MANUAL, DRAWING NO., REV., DESCRIPTION); the first
    # three are "X" or "".
    warnings: List[str] = field(default_factory=list)


def _x(flag: bool) -> str:
    return "X" if flag else ""


def plan_fill(
    data: td.TransmittalData,
    initials: str,
    today: Optional[date] = None,
) -> FillPlan:
    """Turn a TransmittalData + signature into a FillPlan (pure)."""
    today = today or date.today()
    box_index = BOX_ORDER.index(data.box) if data.box in BOX_ORDER else BOX_ORDER.index("approval")
    rows = [
        (_x(r.email), _x(r.print), _x(r.manual), r.drawing_no, r.rev, r.description)
        for r in data.drawing_rows
    ]
    plan = FillPlan(
        order=str(data.order),
        subject=data.customer,
        po=data.po,
        date=today.strftime("%m/%d/%Y"),
        initials=initials,
        box_index=box_index,
        to_emails=list(data.emails),
        rows=rows,
        warnings=list(data.warnings),
    )
    if not initials:
        plan.warnings.append(
            "No signature initials — pass --initials or add the user to engineers.SIGNATURES."
        )
    if not plan.to_emails:
        plan.warnings.append("No recipient emails to place in the TO block.")
    return plan


# --------------------------------------------------------------------------- #
# Word COM applier (Windows + Word only)                                       #
# --------------------------------------------------------------------------- #
def _set_paragraph_value(doc, label_substr: str, value: str) -> bool:
    """Rewrite the first paragraph that contains `label_substr` to
    '<label> <value>', preserving everything up to and including the label.
    Returns True if a paragraph matched."""
    target = label_substr.lower()
    for para in doc.Paragraphs:
        text = para.Range.Text
        low = text.lower()
        idx = low.find(target)
        if idx == -1:
            continue
        # Keep through the label, then write the value, dropping the old
        # placeholder (XXXXXX / xxxxxx) that trailed it.
        head = text[: idx + len(label_substr)]
        rng = para.Range
        # Strip the trailing paragraph mark Word includes in Range.Text.
        rng.End = rng.End - 1
        rng.Text = f"{head} {value}".rstrip()
        return True
    return False


def _fill_to_block(doc, emails: List[str]) -> None:
    """Put every recipient on the TO line (one per line, as line breaks within the
    one paragraph), then DELETE the leftover 'xxx' placeholder lines so no unused
    placeholders remain."""
    linebreak = chr(11)  # vertical tab = a line break inside the paragraph
    to_text = ("TO: " + linebreak.join(emails)) if emails else "TO:"
    for para in doc.Paragraphs:
        if para.Range.Text.strip().lower().startswith("to:"):
            r = para.Range
            r.End = r.End - 1            # exclude the trailing paragraph mark
            r.Text = to_text
            break
    # Remove the placeholder 'xxx' lines (the TO block's only standalone 'xxx'
    # paragraphs — "BY: xxx" is "by: xxx", not a match). Bottom-up so the
    # collection indices stay valid as we delete.
    paras = doc.Paragraphs
    for k in range(paras.Count, 0, -1):
        p = paras(k)
        if p.Range.Text.strip().lower() == "xxx":
            p.Range.Delete()


def pick_checkbox_index(paragraph_texts: List[str], box: str, fallback: int) -> int:
    """The index of the checkbox whose paragraph text names `box` (see
    BOX_LABEL_RES). Falls back to the positional index when the label matches
    no (or more than one) paragraph — pure, so the mapping is unit-testable."""
    rx = BOX_LABEL_RES.get(box)
    if rx is not None:
        hits = [i for i, t in enumerate(paragraph_texts) if rx.search(t or "")]
        if len(hits) == 1:
            return hits[0]
    return fallback


def _check_box(doc, box_index: int) -> None:
    """Tick the FORMCHECKBOX for BOX_ORDER[box_index] and clear the others.

    The box is located by the wording of its own paragraph ('sales purposes' /
    'approval only' / 'record only'), with the document-order index only as the
    fallback — so a template whose form fields enumerate unexpectedly can't
    tick the wrong box silently. What was ticked is logged for the run log."""
    box = BOX_ORDER[box_index] if 0 <= box_index < len(BOX_ORDER) else "approval"
    checkboxes = [ff for ff in doc.FormFields if ff.Type == _WD_FORM_CHECKBOX]
    if len(checkboxes) != len(BOX_ORDER):
        log.warning("Template has %d checkbox form fields (expected %d) — "
                    "relying on the paragraph labels to find the %r box.",
                    len(checkboxes), len(BOX_ORDER), box)
    texts: List[str] = []
    for ff in checkboxes:
        try:
            texts.append(ff.Range.Paragraphs(1).Range.Text or "")
        except Exception:  # noqa: BLE001 - label read is best-effort
            texts.append("")
    target = pick_checkbox_index(texts, box, box_index)
    for i, ff in enumerate(checkboxes):
        try:
            ff.CheckBox.Value = (i == target)
        except Exception as e:  # noqa: BLE001
            log.warning("Could not set checkbox %d: %s", i, e)
    if 0 <= target < len(texts):
        log.info("Ticked the %r box: %s", box, texts[target].strip()[:90] or "(no label read)")
    else:
        log.warning("No checkbox ticked — index %d out of range of %d found.",
                    target, len(checkboxes))


def _fill_table(doc, rows: List[Tuple[str, str, str, str, str, str]]) -> None:
    """Write the drawing rows into the EMAIL/PRINT/MANUAL/DRAWING NO./REV./
    DESCRIPTION table (the first table whose header row carries those columns)."""
    table = None
    for t in doc.Tables:
        try:
            header = " ".join(t.Cell(1, c).Range.Text for c in range(1, 7)).upper()
        except Exception:  # noqa: BLE001 - a table with too few columns
            continue
        if "DRAWING NO" in header and "DESCRIPTION" in header:
            table = t
            break
    if table is None:
        log.warning("Could not locate the drawing table in the template.")
        return
    for i, row in enumerate(rows):
        r = i + 2  # row 1 is the header
        while table.Rows.Count < r:
            table.Rows.Add()
        for c, value in enumerate(row, start=1):
            try:
                cell = table.Cell(r, c)
                cell.Range.Text = value
            except Exception as e:  # noqa: BLE001
                log.warning("Could not write table cell (%d,%d): %s", r, c, e)


def _scrub_placeholders(doc) -> None:
    """Final safety net: clear any template x-placeholders a field-specific fill
    didn't catch (``xxx`` / ``xxxxxx`` / the ``xx/xx/xxxx`` date mask), so a
    generated transmittal never ships with leftover x's.

    Only a run of THREE OR MORE x's (and the date mask) is removed — the single
    ``X`` marks written into the EMAIL/PRINT/MANUAL table columns are left
    untouched. Pure x-runs never occur in real transmittal data (emails, customer
    names, P.O. and drawing numbers), so this is safe to sweep document-wide."""
    # (Word wildcard pattern, …). Date mask first so its 4-x year isn't half-eaten.
    patterns = [
        r"[Xx]{2}/[Xx]{2}/[Xx]{4}",  # the xx/xx/xxxx date mask
        r"<[Xx]{3,}>",               # a whole "word" that is only x's (3+)
    ]
    for pattern in patterns:
        find = doc.Content.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()
        find.Text = pattern
        find.Replacement.Text = ""
        find.Forward = True
        find.Wrap = 1            # wdFindContinue
        find.MatchWildcards = True
        try:
            find.Execute(Replace=2)  # wdReplaceAll
        except Exception as e:  # noqa: BLE001 - a scrub miss must not fail the fill
            log.warning("Placeholder scrub for %r did not run: %s", pattern, e)


def _open_for_review(path: Path) -> None:
    """Open the finished transmittal so a human can eyeball it before it goes out.
    On Windows this launches it in Word; elsewhere it's just logged."""
    try:
        if sys.platform == "win32":
            os.startfile(str(Path(path).resolve()))  # noqa: S606 - opening our own output
        else:
            log.info("Transmittal ready for review: %s", path)
    except Exception as e:  # noqa: BLE001 - never let opening-for-review fail the run
        log.warning("Could not open the transmittal for review (%s): %s", path, e)


def apply_fill_word(plan: FillPlan, template: Path, out_path: Path,
                    open_after: bool = False) -> Path:
    """Open the template in desktop Word, apply the plan, and save to out_path
    (kept as .doc). When ``open_after`` is set, the saved doc is reopened for the
    engineer to review. Windows + Word only."""
    import win32com.client  # lazy: only needed on the Windows box that fills docs

    template = Path(template).resolve()
    if not template.exists():
        raise FileNotFoundError(
            f"Transmittal template not found: {template}. It ships with the code "
            "(DWG TRANSMITTAL MASTER.doc) — make sure you pulled the latest branch."
        )
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = None
    try:
        doc = word.Documents.Open(str(template))
        _fill_to_block(doc, plan.to_emails)
        _set_paragraph_value(doc, "DATE:", plan.date)
        _set_paragraph_value(doc, "CBC ORDER #:", plan.order)
        _set_paragraph_value(doc, "SUBJECT:", plan.subject)
        # The master spells it "P.0. #" (zero); a re-saved copy may use "P.O. #".
        if not _set_paragraph_value(doc, "P.0. #:", plan.po):
            _set_paragraph_value(doc, "P.O. #:", plan.po)
        _set_paragraph_value(doc, "BY:", plan.initials)
        _check_box(doc, plan.box_index)
        _fill_table(doc, plan.rows)
        # Belt-and-suspenders: wipe any x-placeholder the field fills above missed.
        _scrub_placeholders(doc)
        # SaveAs2 with the Word 97-2003 format (0) to keep the .doc extension.
        doc.SaveAs(str(out_path.resolve()), FileFormat=0)
        log.info("Wrote transmittal: %s", out_path)
    finally:
        if doc is not None:
            doc.Close(SaveChanges=False)
        word.Quit()
    # Reopen the saved file (outside the automation instance we just quit) so it
    # comes up in a normal, visible Word window for review.
    if open_after:
        _open_for_review(out_path)
    return out_path


def transmittal_dir(data: td.TransmittalData) -> Optional[Path]:
    """The job's TRANSMITTAL folder (a subfolder of its AutoCAD folder) — where
    transmittals live and usually already exists. None if the job folder is
    unknown."""
    return (Path(data.folder) / "TRANSMITTAL") if data.folder else None


def existing_transmittals(tdir: Optional[Path], order: str) -> list[Path]:
    """Any transmittal Word docs already in the TRANSMITTAL folder for this order
    (often a semi-filled one started by hand), newest revision last."""
    if not tdir or not tdir.exists():
        return []
    hits = [p for p in tdir.glob("*.doc*")
            if "transmittal" in p.name.lower() and str(order) in p.name]
    return sorted(hits)


# A transmittal file is named "<order> DWG TRANSMITTAL-NN[.ext]"; this pulls the
# NN suffix out of the name (leading zeros tolerated).
_SUFFIX_RE = re.compile(r"transmittal\s*-\s*0*(\d+)", re.IGNORECASE)


def sent_transmittal_numbers(tdir: Optional[Path], order: str) -> list[int]:
    """The suffix numbers of transmittals already *sent* for this order, read from
    the saved Outlook emails (``*.msg``) in the job's TRANSMITTAL folder.

    A saved ``.msg`` is the proof a transmittal actually went out — the Word
    ``.doc`` alone is only a draft that may never have been emailed. So numbering
    keys off the ``.msg`` files, not the ``.doc`` files. (Outlook hides the
    ``.msg`` extension in Explorer via NeverShowExt, so these look extension-less
    there, but on disk they really are ``*.msg``.)"""
    if not tdir or not tdir.exists():
        return []
    nums: list[int] = []
    for p in tdir.rglob("*.msg"):
        stem = p.stem
        if "transmittal" not in stem.lower() or str(order) not in stem:
            continue
        m = _SUFFIX_RE.search(stem)
        if m:
            nums.append(int(m.group(1)))
    return sorted(set(nums))


def next_transmittal_suffix(tdir: Optional[Path], order: str) -> str:
    """The two-digit suffix for the NEXT transmittal: one more than the highest
    already-sent (``.msg``) suffix, or ``01`` when none has been sent yet
    (e.g. sent -01 and -02 -> the next is ``03``)."""
    sent = sent_transmittal_numbers(tdir, order)
    nxt = (max(sent) + 1) if sent else 1
    return f"{nxt:02d}"


# Our generated transmittals go in their own subfolder of the job's TRANSMITTAL
# folder, so they never collide with the hand-made / semi-filled docs sitting
# alongside. ("for now" — a clean separation while this is being validated.)
GENERATED_SUBDIR = "AUTO-GENERATED"


def default_out_path(order: str, data: Optional[td.TransmittalData] = None,
                     base_dir: Optional[Path] = None) -> Path:
    """Where the filled transmittal is written: the job's
    <AutoCAD folder>/TRANSMITTAL/AUTO-GENERATED/<order> DWG TRANSMITTAL-NN.doc —
    a dedicated subfolder so our output never clashes with the hand-made docs in
    the TRANSMITTAL folder. The NN suffix is one past the highest already-sent
    transmittal (see ``next_transmittal_suffix``: it counts the saved ``.msg``
    emails in the TRANSMITTAL folder). Falls back to ./transmittal_out/<order>/
    when the job folder isn't known."""
    tdir = transmittal_dir(data) if data is not None else None
    suffix = next_transmittal_suffix(tdir, order)
    if tdir is not None:
        out_dir = tdir / GENERATED_SUBDIR
    elif base_dir:
        out_dir = Path(base_dir) / str(order)
    else:
        out_dir = Path.cwd() / "transmittal_out" / str(order)
    return out_dir / f"{order} DWG TRANSMITTAL-{suffix}.doc"


def fill_transmittal(
    data: td.TransmittalData,
    initials: str = "",
    *,
    template: Path = TEMPLATE_PATH,
    out_path: Optional[Path] = None,
    today: Optional[date] = None,
    open_doc: bool = True,
) -> Tuple[Path, FillPlan]:
    """Gather-to-doc convenience: build the plan and (on Windows) write the doc
    into the job's TRANSMITTAL folder, then open it for review unless ``open_doc``
    is False. Returns (out_path, plan)."""
    initials = initials or engineers.signature_for_user()
    plan = plan_fill(data, initials, today=today)
    # Surface (but never clobber) a semi-filled transmittal already in the folder.
    for p in existing_transmittals(transmittal_dir(data), data.order):
        plan.warnings.append(f"Existing transmittal in the folder (left as-is): {p.name}")
    out = out_path or default_out_path(data.order, data=data)
    apply_fill_word(plan, template, out, open_after=open_doc)
    return out, plan


# --------------------------------------------------------------------------- #
# CLI                                                                          #
# --------------------------------------------------------------------------- #
def main(argv: Optional[List[str]] = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    ap = argparse.ArgumentParser(description="Fill the Drawing Transmittal doc for an order.")
    ap.add_argument("order")
    ap.add_argument("--initials", default="", help="signature (default: from the Windows login)")
    ap.add_argument("--customer", default="", help="override customer (else read from SO)")
    ap.add_argument("--out", default="", help="output .doc path")
    ap.add_argument("--plan-only", action="store_true",
                    help="print the fill plan without opening Word (works off-Windows)")
    ap.add_argument("--no-open", action="store_true",
                    help="don't open the finished .doc for review after writing it")
    args = ap.parse_args(argv)

    data = td.gather(args.order, customer=args.customer)
    initials = args.initials or engineers.signature_for_user()
    plan = plan_fill(data, initials)

    print(f"\n=== Fill plan for {args.order} ===")
    print(f"  Date    : {plan.date}")
    print(f"  Order   : {plan.order}")
    print(f"  Subject : {plan.subject}")
    print(f"  P.O. #  : {plan.po}")
    print(f"  Box     : {BOX_ORDER[plan.box_index]}")
    print(f"  Because : {data.box_evidence or '(no status evidence recorded)'}")
    print(f"  BY      : {plan.initials or '(none)'}")
    print(f"  TO:")
    for e in plan.to_emails:
        print(f"     {e}")
    print(f"  Table:")
    for (em, pr, mn, no, rev, desc) in plan.rows:
        marks = "".join(c if v else "·" for c, v in (("E", em), ("P", pr), ("M", mn)))
        print(f"     [{marks}] {no or '(blank)':<14} {desc}")
    for w in plan.warnings:
        print(f"  ! {w}")

    if args.plan_only:
        return 0

    out = Path(args.out) if args.out else default_out_path(args.order, data=data)
    try:
        apply_fill_word(plan, TEMPLATE_PATH, out, open_after=not args.no_open)
        print(f"\nWrote: {out}")
    except ImportError:
        print("\n(win32com not available here — run on the Windows box, or use --plan-only.)")
        return 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
