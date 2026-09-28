"""Domain models and enumerations for Beacon purchasing documents."""

from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from enum import Enum
from typing import Any

from pydantic import BaseModel, Field, field_validator


class DocumentType(str, Enum):
    ACKNOWLEDGEMENT = "ACKNOWLEDGEMENT"
    REVISION = "REVISION"
    INVOICE = "INVOICE"
    SHIPPING_NOTICE = "SHIPPING_NOTICE"
    RECEIPT_ONLY = "RECEIPT_ONLY"
    OTHER = "OTHER"
    UNKNOWN = "UNKNOWN"


class CommitmentStatus(str, Enum):
    FULL = "FULL"
    PARTIAL = "PARTIAL"
    RECEIPT_ACKNOWLEDGED_NO_SCHEDULE = "RECEIPT_ACKNOWLEDGED_NO_SCHEDULE"
    UNKNOWN = "UNKNOWN"


class DateType(str, Enum):
    PROMISE_DATE = "PROMISE_DATE"
    DELIVERY_DATE = "DELIVERY_DATE"
    SHIP_DATE = "SHIP_DATE"
    DELIVERY_WINDOW = "DELIVERY_WINDOW"
    UNKNOWN = "UNKNOWN"


class DatePrecision(str, Enum):
    EXACT_DATE = "EXACT_DATE"
    WEEK = "WEEK"
    WEEK_RANGE = "WEEK_RANGE"
    TEXT_ONLY = "TEXT_ONLY"
    UNKNOWN = "UNKNOWN"


class Severity(str, Enum):
    RED = "RED"
    YELLOW = "YELLOW"
    INFO = "INFO"
    GREEN = "GREEN"


class MatchMethod(str, Enum):
    EXACT_CUSTOMER_PN = "EXACT_CUSTOMER_PN"
    APPROVED_VENDOR_CROSSWALK = "APPROVED_VENDOR_CROSSWALK"
    EXACT_VENDOR_PN_TO_BEACON_PN = "EXACT_VENDOR_PN_TO_BEACON_PN"
    RULE_BASED = "RULE_BASED"
    SINGLE_REMAINING_LINE = "SINGLE_REMAINING_LINE"
    MANUAL_REVIEW = "MANUAL_REVIEW"
    UNMATCHED = "UNMATCHED"


class CrosswalkStatus(str, Enum):
    APPROVED = "APPROVED"
    PROPOSED = "PROPOSED"
    REJECTED = "REJECTED"


class DocumentLifecycle(str, Enum):
    ACTIVE = "ACTIVE"
    SUPERSEDED = "SUPERSEDED"
    INFORMATIONAL = "INFORMATIONAL"
    REVIEW_REQUIRED = "REVIEW_REQUIRED"


class ExtractionMode(str, Enum):
    NATIVE_TEXT = "NATIVE_TEXT"
    VISION = "VISION"
    OCR_LOCAL = "OCR_LOCAL"
    LOCAL_TEXT = "LOCAL_TEXT"
    CACHED = "CACHED"
    FAILED = "FAILED"


class PromiseWindow(BaseModel):
    raw_text: str | None = None
    start_date: date | None = None
    end_date: date | None = None
    date_type: DateType = DateType.UNKNOWN
    precision: DatePrecision = DatePrecision.UNKNOWN

    @property
    def earliest(self) -> date | None:
        return self.start_date

    @property
    def latest(self) -> date | None:
        return self.end_date or self.start_date

    @property
    def is_usable(self) -> bool:
        return self.start_date is not None or self.end_date is not None


class ConfirmationLine(BaseModel):
    source_line_number: int | None = None
    customer_part_number: str | None = None
    vendor_part_number: str | None = None
    description: str | None = None
    quantity: float | None = None
    uom: str | None = None
    unit_price: Decimal | None = None
    currency: str | None = None
    promise: PromiseWindow = Field(default_factory=PromiseWindow)
    page_number: int | None = None

    @field_validator("unit_price", mode="before")
    @classmethod
    def _coerce_price(cls, v: Any) -> Any:
        if v is None or v == "":
            return None
        if isinstance(v, Decimal):
            return v
        return Decimal(str(v))


class ExtractedDocument(BaseModel):
    source_file: str
    source_sha256: str
    vendor_name_raw: str | None = None
    po_number: str | None = None
    document_number: str | None = None
    document_date: date | None = None
    document_type: DocumentType = DocumentType.UNKNOWN
    commitment_status: CommitmentStatus = CommitmentStatus.UNKNOWN

    is_revision: bool = False
    supersedes_all_prior_for_po: bool = False

    partial_sequence: int | None = None
    partial_total: int | None = None

    document_currency: str | None = None
    lines: list[ConfirmationLine] = Field(default_factory=list)

    free_text_notes: list[str] = Field(default_factory=list)
    extraction_warnings: list[str] = Field(default_factory=list)

    extraction_mode: ExtractionMode = ExtractionMode.NATIVE_TEXT
    model_name: str | None = None
    prompt_version: str | None = None
    schema_version: str | None = None
    extracted_at: datetime | None = None
    page_count: int | None = None
    native_text_chars: int | None = None
    cached: bool = False

    # Lifecycle filled later
    lifecycle: DocumentLifecycle = DocumentLifecycle.ACTIVE
    vendor_id: str | None = None
    vendor_name_resolved: str | None = None


class OpenPOLine(BaseModel):
    po_number: str
    po_date: date | None = None
    vendor_id: str | None = None
    vendor_name: str | None = None
    line_number: int
    our_pn: str
    our_description: str | None = None
    qty_ordered: Decimal
    unit_price: Decimal
    required_date: date | None = None
    currency_assumed: str = "USD"


class VendorMasterRow(BaseModel):
    vendor_id: str
    vendor_name: str
    country: str | None = None
    ap_email: str | None = None
    known_pn_mapping_note: str | None = None
    currency: str | None = None


class CrosswalkEntry(BaseModel):
    vendor_id: str
    vendor_name: str | None = None
    vendor_pn: str
    beacon_pn: str
    evidence_count: int = 0
    vendor_pn_total_count: int = 0
    purity: float = 0.0
    first_seen: date | None = None
    last_seen: date | None = None
    status: CrosswalkStatus = CrosswalkStatus.PROPOSED
    source: str = "historical_confirmation"


class MatchResult(BaseModel):
    po_line: OpenPOLine | None = None
    confirmation_line: ConfirmationLine | None = None
    source_document: ExtractedDocument | None = None
    method: MatchMethod = MatchMethod.UNMATCHED
    confidence: float = 0.0
    explanation: str = ""


class ReconciledLine(BaseModel):
    po_number: str
    po_line_number: int | None = None
    vendor_id: str | None = None
    vendor_name: str | None = None
    vendor_contact: str | None = None
    beacon_pn: str | None = None
    vendor_pn: str | None = None
    description: str | None = None

    qty_ordered: Decimal | None = None
    qty_confirmed: Decimal | None = None
    qty_variance: Decimal | None = None

    po_unit_price: Decimal | None = None
    po_currency_assumed: str | None = None
    confirmed_unit_price_raw: Decimal | None = None
    confirmed_currency: str | None = None
    fx_rate_used: Decimal | None = None
    confirmed_price_usd: Decimal | None = None
    unit_price_difference: Decimal | None = None
    unit_price_difference_pct: Decimal | None = None
    price_impact_on_order: Decimal | None = None
    price_stated: bool = False

    required_date: date | None = None
    commitment_date_type: DateType | None = None
    promise_start: date | None = None
    promise_end: date | None = None
    promise_raw: str | None = None
    projected_full_qty_promise_date: date | None = None
    days_late: int | None = None
    minimum_late_days: int | None = None
    maximum_late_days: int | None = None

    match_method: MatchMethod | None = None
    match_confidence: float | None = None
    match_explanation: str | None = None

    source_file: str | None = None
    source_document_number: str | None = None
    document_date: date | None = None
    document_status: str | None = None
    document_type: DocumentType | None = None
    source_sha256: str | None = None

    issue_codes: list[str] = Field(default_factory=list)
    severity: Severity = Severity.GREEN
    suggested_action: str | None = None
    sort_key: tuple = Field(default_factory=tuple)

    extra_source_files: list[str] = Field(default_factory=list)
    schedule_details: list[str] = Field(default_factory=list)
    is_actionable: bool = False
    unknown_po: bool = False
    extraction_warnings: list[str] = Field(default_factory=list)


class RunStats(BaseModel):
    started_at: datetime | None = None
    finished_at: datetime | None = None
    pdfs_processed: int = 0
    acknowledgments: int = 0
    revisions: int = 0
    invoices: int = 0
    scanned_pdfs: int = 0
    cached_extractions: int = 0
    ai_extractions: int = 0
    local_extractions: int = 0
    vision_extractions: int = 0
    native_text_openai: int = 0
    vision_openai: int = 0
    local_fallback: int = 0
    review_required: int = 0
    extraction_failures: int = 0
    open_po_lines: int = 0
    matched_lines: int = 0
    missing_lines: int = 0
    red_issues: int = 0
    yellow_issues: int = 0
    unknown_pos: int = 0
    warnings: int = 0
    offline: bool = False
    refresh_cache: bool = False
    llm_provider: str | None = None
    llm_model: str | None = None
    inputs: dict[str, str] = Field(default_factory=dict)
    generated_files: list[str] = Field(default_factory=list)
