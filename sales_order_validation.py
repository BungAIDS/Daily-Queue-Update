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


@dataclass(frozen=True)
class SalesOrderValidation:
    expected_order: str
    internal_order: str
    status: str
    method: str
    error: str = ""

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


def validate_sales_order_pdf(path: str | Path, expected_order: str) -> SalesOrderValidation:
    expected = normalize_order(expected_order)
    try:
        internal, method = extract_internal_order(_extract_first_page(Path(path)))
    except Exception as exc:  # noqa: BLE001 - validation must fail closed
        return SalesOrderValidation(expected, "", "ERROR", "none", f"{type(exc).__name__}: {exc}")
    if not internal:
        return SalesOrderValidation(expected, "", "UNREADABLE", method)
    status = "MATCH" if internal == expected else "MISMATCH"
    return SalesOrderValidation(expected, internal, status, method)


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


def finalize_candidate(
    candidate: str | Path, destination: str | Path, expected_order: str
) -> SalesOrderAcceptance:
    candidate = Path(candidate)
    destination = Path(destination)
    validation = validate_sales_order_pdf(candidate, expected_order)
    if not validation.matched:
        quarantined = quarantine_candidate(candidate, destination, validation)
        return SalesOrderAcceptance(None, validation, str(quarantined))

    if candidate.resolve() == destination.resolve():
        return SalesOrderAcceptance(str(destination), validation)

    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        existing = validate_sales_order_pdf(destination, expected_order)
        if existing.matched:
            duplicate = quarantine_candidate(
                candidate, destination, validation, bucket="duplicate-valid-downloads"
            )
            return SalesOrderAcceptance(str(destination), existing, str(duplicate))
        quarantine_candidate(destination, destination, existing, bucket="rejected-existing-files")
    candidate.replace(destination)
    return SalesOrderAcceptance(str(destination), validation)


def accept_existing(destination: str | Path, expected_order: str) -> SalesOrderAcceptance | None:
    destination = Path(destination)
    if not destination.exists():
        return None
    return finalize_candidate(destination, destination, expected_order)


def failed_acceptance(expected_order: str, error: str) -> SalesOrderAcceptance:
    validation = SalesOrderValidation(
        normalize_order(expected_order), "", "DOWNLOAD_FAILED", "none", error
    )
    return SalesOrderAcceptance(None, validation)
