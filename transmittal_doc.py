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

    python transmittal_doc.py 421693                 # gather + fill (preview)
    python transmittal_doc.py 421693 --initials DAG  # override the signature
"""
from __future__ import annotations

import argparse
import logging
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
    """Place recipient emails one-per-line into the TO block: the 'TO:' paragraph
    plus the placeholder 'xxx' lines beneath it."""
    paras = list(doc.Paragraphs)
    for i, para in enumerate(paras):
        if para.Range.Text.strip().lower().startswith("to:"):
            first = emails[0] if emails else ""
            r = para.Range
            r.End = r.End - 1
            r.Text = f"TO: {first}".rstrip()
            # The following placeholder lines ('xxx') take the remaining emails;
            # any left over are cleared to blank.
            rest = emails[1:]
            j = i + 1
            while j < len(paras) and paras[j].Range.Text.strip().lower() in ("xxx", ""):
                r2 = paras[j].Range
                r2.End = r2.End - 1
                r2.Text = rest.pop(0) if rest else ""
                if not rest and paras[j].Range.Text.strip() == "":
                    break
                j += 1
            return


def _check_box(doc, box_index: int) -> None:
    """Tick the FORMCHECKBOX at `box_index` (sales/approval/record) and clear the
    others."""
    checkboxes = [ff for ff in doc.FormFields if ff.Type == _WD_FORM_CHECKBOX]
    for i, ff in enumerate(checkboxes):
        try:
            ff.CheckBox.Value = (i == box_index)
        except Exception as e:  # noqa: BLE001
            log.warning("Could not set checkbox %d: %s", i, e)


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


def apply_fill_word(plan: FillPlan, template: Path, out_path: Path) -> Path:
    """Open the template in desktop Word, apply the plan, and save to out_path
    (kept as .doc). Windows + Word only."""
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
        # SaveAs2 with the Word 97-2003 format (0) to keep the .doc extension.
        doc.SaveAs(str(out_path.resolve()), FileFormat=0)
        log.info("Wrote transmittal: %s", out_path)
    finally:
        if doc is not None:
            doc.Close(SaveChanges=False)
        word.Quit()
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


# Our generated transmittals go in their own subfolder of the job's TRANSMITTAL
# folder, so they never collide with the hand-made / semi-filled docs sitting
# alongside. ("for now" — a clean separation while this is being validated.)
GENERATED_SUBDIR = "AUTO-GENERATED"


def default_out_path(order: str, data: Optional[td.TransmittalData] = None,
                     base_dir: Optional[Path] = None) -> Path:
    """Where the filled transmittal is written: the job's
    <AutoCAD folder>/TRANSMITTAL/AUTO-GENERATED/<order> DWG TRANSMITTAL-01.doc —
    a dedicated subfolder so our output never clashes with the hand-made docs in
    the TRANSMITTAL folder. Falls back to ./transmittal_out/<order>/ when the job
    folder isn't known. Always the same name (overwrites our own prior run)."""
    tdir = transmittal_dir(data) if data is not None else None
    if tdir is not None:
        tdir = tdir / GENERATED_SUBDIR
    elif base_dir:
        tdir = Path(base_dir) / str(order)
    else:
        tdir = Path.cwd() / "transmittal_out" / str(order)
    return tdir / f"{order} DWG TRANSMITTAL-01.doc"


def fill_transmittal(
    data: td.TransmittalData,
    initials: str = "",
    *,
    template: Path = TEMPLATE_PATH,
    out_path: Optional[Path] = None,
    today: Optional[date] = None,
) -> Tuple[Path, FillPlan]:
    """Gather-to-doc convenience: build the plan and (on Windows) write the doc
    into the job's TRANSMITTAL folder. Returns (out_path, plan)."""
    initials = initials or engineers.signature_for_user()
    plan = plan_fill(data, initials, today=today)
    # Surface (but never clobber) a semi-filled transmittal already in the folder.
    for p in existing_transmittals(transmittal_dir(data), data.order):
        plan.warnings.append(f"Existing transmittal in the folder (left as-is): {p.name}")
    out = out_path or default_out_path(data.order, data=data)
    apply_fill_word(plan, template, out)
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
        apply_fill_word(plan, TEMPLATE_PATH, out)
        print(f"\nWrote: {out}")
    except ImportError:
        print("\n(win32com not available here — run on the Windows box, or use --plan-only.)")
        return 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
