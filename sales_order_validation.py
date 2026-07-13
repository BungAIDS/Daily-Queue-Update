"""Validate a Sales Order PDF before it can enter the active archive or stores."""
from __future__ import annotations

import json
import re
import uuid
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path


ORDER_HEADER_RE = re.compile(
    r"\bOrder\s*#\s*(?:Rep\s*Ref\.?\s*#|RepRef\.?\s*#).*?(?:\r?\n)+\s*([0-9]{5,7}[A-Z]?)\b",
    re.IGNORECASE | re.DOTALL,
)
ORDER_JOB_TABLE_RE = re.compile(
    r"\bOrder\s*#\s+Job\s*#[^\r\n]*(?:\r?\n)+\s*([0-9]{5,7}[A-Z]?)\b",
    re.IGNORECASE,
)
ORDER_INLINE_RE = re.compile(r"\bOrder\s*#\s*[:\-]?\s*([0-9]{5,7}[A-Z]?)\b", re.IGNORECASE)
SALES_ORDER_RE = re.compile(
    r"\bSales\s+Order\s+(?:No\.?\s*)?([0-9]{5,7}[A-Z]?)\b", re.IGNORECASE
)
ORDER_VERIFICATION_RE = re.compile(
    r"\bOrder\s+Cust\s+PO\s+Ship\s+Via\s+Package\s+Prepaid\s+Date\s+Order\s+Terms\s+Verification\s+Date"
    r"\s*(?:\r?\n)+\s*([0-9]{5,7}[A-Z]?)\b",
    re.IGNORECASE,
)
ORDER_VERIFICATION_TITLE_RE = re.compile(r"\bOrder\s+Verification\s+Report\b", re.IGNORECASE)
SALES_ORDER_TITLE_RE = re.compile(
    r"\b(?:Chicago\s+Blower\s+Corporation\s+)?Sales\s+Order\b", re.IGNORECASE
)

DOCUMENT_KIND_SALES_ORDER = "SALES_ORDER"
DOCUMENT_KIND_ORDER_VERIFICATION = "ORDER_VERIFICATION"
DOCUMENT_KIND_UNKNOWN = "UNKNOWN"

# Everything populated from the contents or provenance of a Sales Order PDF.
# ``so_imi`` is intentionally absent: it comes from the AutoCAD job folder, not
# from the PDF.  These fields are cleared together when a stored document is
# proven to be an Order Verification Report.
SALES_ORDER_DERIVED_FIELDS = frozenset({
    "co_number", "co_history",
    "so_pdf", "so_verified_at", "so_validation", "so_internal_order",
    "so_validation_method", "so_document_kind", "so_source_type", "so_quarantine",
    "so_design_desc", "so_size", "so_arrangement", "so_motor_pos", "so_class",
    "so_rotation", "so_discharge", "so_pct_width", "so_wheel_type",
    "so_design_temp", "so_max_temp", "so_special_temp",
    "so_emails", "so_po", "so_released",
    "line_items", "line_item_tags", "line_item_count",
    "dwg_reuse", "dwg_reuse_label", "dwg_reuse_note",
})

SO_INVALIDATED_AT = "so_invalidated_at"


@dataclass(frozen=True)
class SalesOrderValidation:
    expected_order: str
    internal_order: str
    status: str
    method: str
    error: str = ""
    document_kind: str = DOCUMENT_KIND_UNKNOWN

    @property
    def matched(self) -> bool:
        return self.status == "MATCH"


@dataclass(frozen=True)
class SalesOrderAcceptance:
    path: str | None
    validation: SalesOrderValidation
    quarantine_path: str = ""


def normalize_order(value: str) -> str:
    return re.sub(r"[^0-9A-Z]", "", str(value or "").upper())


def clear_sales_order_data(record: dict, invalidated_at: str = "") -> bool:
    """Remove all PDF-derived Sales Order data from one stored job record."""
    changed = False
    for field in SALES_ORDER_DERIVED_FIELDS:
        if field in record:
            record.pop(field, None)
            changed = True
    if invalidated_at and record.get(SO_INVALIDATED_AT) != invalidated_at:
        record[SO_INVALIDATED_AT] = invalidated_at
        changed = True
    return changed


def is_order_verification_record(record: dict) -> bool:
    """True when stored provenance identifies an Order Verification Report."""
    return (
        str(record.get("so_document_kind") or "").upper()
        == DOCUMENT_KIND_ORDER_VERIFICATION
        or str(record.get("so_source_type") or "").casefold() == "cs_salesorder"
        or str(record.get("so_validation_method") or "").casefold()
        == "order-verification-table"
        or "orderverificationreportviewer" in str(record.get("so_pdf") or "").casefold()
    )


def is_true_sales_order_record(record: dict) -> bool:
    """True when provenance and a stored path identify a genuine Sales Order."""
    kind = str(record.get("so_document_kind") or "").upper()
    source = str(record.get("so_source_type") or "").casefold()
    return (
        kind == DOCUMENT_KIND_SALES_ORDER
        and bool(str(record.get("so_pdf") or "").strip())
        and source in {"", "cbc_salesorder"}
    )


def effective_sales_order_invalidation(*records: dict) -> str:
    """Newest invalidation that has not been superseded by a later true SO."""
    invalidated = max(
        (str(record.get(SO_INVALIDATED_AT) or "") for record in records if record),
        default="",
    )
    verified = max(
        (str(record.get("so_verified_at") or "") for record in records if record),
        default="",
    )
    return invalidated if invalidated and invalidated >= verified else ""


def extract_internal_order(text: str) -> tuple[str, str]:
    for regex, method in (
        (ORDER_JOB_TABLE_RE, "order-job-table"),
        (ORDER_HEADER_RE, "header"),
        (ORDER_INLINE_RE, "inline"),
        (ORDER_VERIFICATION_RE, "order-verification-table"),
        (SALES_ORDER_RE, "sales-order-label"),
    ):
        match = regex.search(text or "")
        if match:
            return normalize_order(match.group(1)), method
    return "", "none"


def classify_sales_order_document(text: str) -> str:
    """Identify the two same-order PDF layouts CBC exposes as SalesOrder pids."""
    text = text or ""
    if ORDER_VERIFICATION_TITLE_RE.search(text) or ORDER_VERIFICATION_RE.search(text):
        return DOCUMENT_KIND_ORDER_VERIFICATION
    if SALES_ORDER_TITLE_RE.search(text):
        return DOCUMENT_KIND_SALES_ORDER
    return DOCUMENT_KIND_UNKNOWN


def _extract_first_page(path: Path) -> str:
    try:
        from pypdf import PdfReader

        reader = PdfReader(str(path), strict=False)
        if reader.pages:
            text = reader.pages[0].extract_text() or ""
            if text.strip():
                return text
    except Exception:
        pass

    import pdfplumber

    with pdfplumber.open(str(path)) as pdf:
        return (pdf.pages[0].extract_text() or "") if pdf.pages else ""


def validate_sales_order_pdf(
    path: str | Path,
    expected_order: str,
    required_document_kind: str | None = None,
) -> SalesOrderValidation:
    expected = normalize_order(expected_order)
    try:
        text = _extract_first_page(Path(path))
        internal, method = extract_internal_order(text)
        document_kind = classify_sales_order_document(text)
    except Exception as exc:  # noqa: BLE001 - validation must fail closed
        return SalesOrderValidation(
            expected,
            "",
            "ERROR",
            "none",
            f"{type(exc).__name__}: {exc}",
            DOCUMENT_KIND_UNKNOWN,
        )
    if not internal:
        return SalesOrderValidation(
            expected, "", "UNREADABLE", method, document_kind=document_kind
        )
    if internal != expected:
        status = "MISMATCH"
    elif required_document_kind and document_kind != required_document_kind:
        status = "WRONG_DOCUMENT"
    else:
        status = "MATCH"
    return SalesOrderValidation(
        expected, internal, status, method, document_kind=document_kind
    )


def modal_text_matches_job(text: str, expected_order: str) -> bool:
    expected = normalize_order(expected_order)
    if not expected:
        return False
    return bool(re.search(rf"(?<![A-Z0-9]){re.escape(expected)}(?:-\d+)?(?![A-Z0-9])", text.upper()))


def _quarantine_root(destination: Path, expected_order: str) -> Path:
    expected = normalize_order(expected_order)
    active_root = destination.parent.parent
    for parent in destination.parents:
        if normalize_order(parent.name) == expected:
            active_root = parent.parent
            break
    return active_root.parent / f"{active_root.name} QUARANTINE"


def staging_path(destination: str | Path, expected_order: str) -> Path:
    destination = Path(destination)
    return (
        _quarantine_root(destination, expected_order)
        / "pending-downloads"
        / datetime.now().strftime("%Y%m%d")
        / normalize_order(expected_order)
        / uuid.uuid4().hex
        / destination.name
    )


def _unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stamp = datetime.now().strftime("%H%M%S%f")
    return path.with_name(f"{path.stem}-{stamp}-{uuid.uuid4().hex[:8]}{path.suffix}")


def quarantine_candidate(
    candidate: str | Path,
    destination: str | Path,
    validation: SalesOrderValidation,
    bucket: str = "rejected-downloads",
) -> Path:
    candidate = Path(candidate)
    destination = Path(destination)
    actual = validation.internal_order or "UNKNOWN"
    target = _unique_path(
        _quarantine_root(destination, validation.expected_order)
        / bucket
        / datetime.now().strftime("%Y%m%d")
        / validation.expected_order
        / f"{validation.status}-{actual}"
        / candidate.name
    )
    target.parent.mkdir(parents=True, exist_ok=True)
    original_path = str(candidate)
    candidate.replace(target)
    metadata = {
        **asdict(validation),
        "original_path": original_path,
        "quarantine_path": str(target),
        "quarantined_at": datetime.now().isoformat(timespec="seconds"),
    }
    try:
        target.with_suffix(target.suffix + ".validation.json").write_text(
            json.dumps(metadata, indent=2), encoding="utf-8"
        )
    except OSError:
        pass
    return target


def quarantine_superseded_verification_reports(
    destination: str | Path, expected_order: str
) -> list[str]:
    """Move sibling verification reports aside once a true SO is confirmed."""
    destination = Path(destination)
    moved: list[str] = []
    try:
        siblings = list(destination.parent.glob("*.pdf"))
    except OSError:
        return moved
    for path in siblings:
        try:
            if path.resolve() == destination.resolve():
                continue
            validation = validate_sales_order_pdf(path, expected_order)
            if validation.document_kind != DOCUMENT_KIND_ORDER_VERIFICATION:
                continue
            target = quarantine_candidate(
                path,
                destination,
                validation,
                bucket="superseded-order-verification-reports",
            )
            moved.append(str(target))
        except OSError:
            continue
    return moved


def finalize_candidate(
    candidate: str | Path,
    destination: str | Path,
    expected_order: str,
    required_document_kind: str | None = DOCUMENT_KIND_SALES_ORDER,
) -> SalesOrderAcceptance:
    candidate = Path(candidate)
    destination = Path(destination)
    required_document_kind = required_document_kind or DOCUMENT_KIND_SALES_ORDER
    validation = validate_sales_order_pdf(
        candidate, expected_order, required_document_kind
    )
    if not validation.matched:
        quarantined = quarantine_candidate(candidate, destination, validation)
        return SalesOrderAcceptance(None, validation, str(quarantined))

    if candidate.resolve() == destination.resolve():
        if required_document_kind == DOCUMENT_KIND_SALES_ORDER:
            quarantine_superseded_verification_reports(destination, expected_order)
        return SalesOrderAcceptance(str(destination), validation)

    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        existing = validate_sales_order_pdf(
            destination, expected_order, required_document_kind
        )
        if existing.matched:
            duplicate = quarantine_candidate(
                candidate, destination, validation, bucket="duplicate-valid-downloads"
            )
            if required_document_kind == DOCUMENT_KIND_SALES_ORDER:
                quarantine_superseded_verification_reports(destination, expected_order)
            return SalesOrderAcceptance(str(destination), existing, str(duplicate))
        quarantine_candidate(destination, destination, existing, bucket="rejected-existing-files")
    candidate.replace(destination)
    if required_document_kind == DOCUMENT_KIND_SALES_ORDER:
        quarantine_superseded_verification_reports(destination, expected_order)
    return SalesOrderAcceptance(str(destination), validation)


def accept_existing(
    destination: str | Path,
    expected_order: str,
    required_document_kind: str | None = DOCUMENT_KIND_SALES_ORDER,
) -> SalesOrderAcceptance | None:
    destination = Path(destination)
    if not destination.exists():
        return None
    return finalize_candidate(
        destination, destination, expected_order, required_document_kind
    )


def failed_acceptance(expected_order: str, error: str) -> SalesOrderAcceptance:
    validation = SalesOrderValidation(
        normalize_order(expected_order), "", "DOWNLOAD_FAILED", "none", error
    )
    return SalesOrderAcceptance(None, validation)
