"""Extract structured confirmation documents from PDFs.

Pipeline:
    PDF -> native text quality check
         -> good text: LLM (or local parser fallback)
         -> scan: vision LLM, else OCR + local parser
    Then Pydantic validation. Failures become REVIEW_REQUIRED records.
The open PO list is never provided to the extractor.
"""

from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
from typing import Any

from pydantic import ValidationError

from src import config
from src.config import PROMPT_VERSION, SCHEMA_VERSION
from src.llm_client import (
    AuthenticationFailed,
    ExtractionError,
    LLMClient,
    build_llm_client,
    is_auth_error,
    sanitize_error,
)
from src.local_extractor import classify_document, parse_document_text
from src.models import (
    CommitmentStatus,
    ConfirmationLine,
    DateType,
    DocumentType,
    ExtractedDocument,
    ExtractionMode,
)
from src.normalize import (
    detect_currency,
    extract_po_number,
    normalize_part_number,
    normalize_po_number,
    parse_date,
    parse_promise,
)
from src.pdf_reader import PdfDocument, ocr_pdf, read_pdf, render_pages_png

logger = logging.getLogger("beacon")


def _now() -> datetime:
    return datetime.now(timezone.utc)


def cache_path(cache_dir: Path, sha256: str) -> Path:
    return cache_dir / f"{sha256}.json"


def load_cache(
    cache_dir: Path,
    sha256: str,
    *,
    prompt_version: str,
    schema_version: str,
    model_name: str,
) -> ExtractedDocument | None:
    path = cache_path(cache_dir, sha256)
    if not path.exists():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if payload.get("prompt_version") != prompt_version:
        return None
    if payload.get("schema_version") != schema_version:
        return None
    if payload.get("model_name") != model_name:
        return None
    if payload.get("file_hash") != sha256:
        return None
    doc_data = payload.get("document")
    if not doc_data:
        return None
    try:
        doc = ExtractedDocument.model_validate(doc_data)
    except ValidationError:
        return None
    doc.cached = True
    return doc


def save_cache(
    cache_dir: Path,
    doc: ExtractedDocument,
    *,
    model_name: str,
) -> None:
    cache_dir.mkdir(parents=True, exist_ok=True)
    payload = {
        "file_hash": doc.source_sha256,
        "prompt_version": doc.prompt_version,
        "schema_version": doc.schema_version,
        "model": model_name,
        "model_name": model_name,
        "timestamp": (doc.extracted_at or _now()).isoformat(),
        "extraction_mode": doc.extraction_mode.value,
        "document": json.loads(doc.model_dump_json()),
    }
    cache_path(cache_dir, doc.source_sha256).write_text(
        json.dumps(payload, indent=2, default=str),
        encoding="utf-8",
    )


def _payload_to_document(
    payload: dict[str, Any],
    pdf: PdfDocument,
    *,
    mode: ExtractionMode,
    model_name: str,
    warnings: list[str] | None = None,
) -> ExtractedDocument:
    prefer_dmy = False
    raw_vendor = payload.get("vendor_name_raw") or ""
    if re_search_de(raw_vendor) or (payload.get("document_currency") or "").upper() == "EUR":
        prefer_dmy = True

    doc_date = parse_date(payload.get("document_date"), prefer_dmy=prefer_dmy)
    po = normalize_po_number(payload.get("po_number")) or extract_po_number(str(payload.get("po_number") or ""))
    currency = payload.get("document_currency")
    if currency:
        currency = str(currency).upper()

    lines: list[ConfirmationLine] = []
    for raw in payload.get("lines") or []:
        promise_raw = raw.get("promise_raw_text") or raw.get("promise") or None
        if isinstance(promise_raw, dict):
            promise_raw = promise_raw.get("raw_text")
        hint = raw.get("date_type") or ""
        promise = parse_promise(promise_raw, header_hint=hint, prefer_dmy=prefer_dmy)
        # Python owns week-range parsing; never collapse KW windows.
        if raw.get("date_type"):
            try:
                dt = DateType(raw["date_type"])
                if promise.date_type in {DateType.UNKNOWN, DateType.PROMISE_DATE}:
                    promise.date_type = dt
            except ValueError:
                pass
        price = raw.get("unit_price")
        if price is not None and price != "":
            try:
                price = Decimal(str(price))
            except Exception:
                price = None
        else:
            price = None
        qty = raw.get("quantity")
        try:
            qty_f = float(qty) if qty is not None and qty != "" else None
        except (TypeError, ValueError):
            qty_f = None
        line_currency = raw.get("currency")
        if line_currency:
            line_currency = str(line_currency).upper()
        lines.append(
            ConfirmationLine(
                source_line_number=raw.get("source_line_number"),
                customer_part_number=raw.get("customer_part_number"),
                vendor_part_number=raw.get("vendor_part_number"),
                description=raw.get("description"),
                quantity=qty_f,
                uom=raw.get("uom"),
                unit_price=price,
                currency=line_currency or currency,
                promise=promise,
                page_number=raw.get("page_number"),
            )
        )

    try:
        doc_type = DocumentType(payload.get("document_type") or "UNKNOWN")
    except ValueError:
        doc_type = DocumentType.UNKNOWN
    try:
        commitment = CommitmentStatus(payload.get("commitment_status") or "UNKNOWN")
    except ValueError:
        commitment = CommitmentStatus.UNKNOWN

    warn = list(payload.get("extraction_warnings") or [])
    if warnings:
        warn.extend(warnings)

    return ExtractedDocument(
        source_file=str(pdf.path),
        source_sha256=pdf.sha256,
        vendor_name_raw=payload.get("vendor_name_raw"),
        po_number=po,
        document_number=payload.get("document_number"),
        document_date=doc_date,
        document_type=doc_type,
        commitment_status=commitment,
        is_revision=bool(payload.get("is_revision")),
        supersedes_all_prior_for_po=bool(payload.get("supersedes_all_prior_for_po")),
        partial_sequence=payload.get("partial_sequence"),
        partial_total=payload.get("partial_total"),
        document_currency=currency,
        lines=lines,
        free_text_notes=list(payload.get("free_text_notes") or []),
        extraction_warnings=warn,
        extraction_mode=mode,
        model_name=model_name,
        prompt_version=PROMPT_VERSION,
        schema_version=SCHEMA_VERSION,
        extracted_at=_now(),
        page_count=pdf.page_count,
        native_text_chars=len(pdf.native_text.strip()),
    )


def re_search_de(text: str) -> bool:
    return bool(text) and any(
        tok in text.lower()
        for tok in ("gmbh", "ostmark", "deutschland", "n\u00fcrnberg", "nuernberg")
    )


def failed_document(pdf: PdfDocument, error: str, mode: ExtractionMode, model_name: str) -> ExtractedDocument:
    return ExtractedDocument(
        source_file=str(pdf.path),
        source_sha256=pdf.sha256,
        document_type=DocumentType.UNKNOWN,
        extraction_mode=ExtractionMode.FAILED,
        model_name=model_name,
        prompt_version=PROMPT_VERSION,
        schema_version=SCHEMA_VERSION,
        extracted_at=_now(),
        page_count=pdf.page_count,
        native_text_chars=len(pdf.native_text.strip()),
        extraction_warnings=[error],
    )


def _safe_exc(exc: BaseException) -> str:
    return sanitize_error(exc)


def _line_match_key(customer_pn, vendor_pn, quantity, source_line_number) -> tuple:
    pn = normalize_part_number(customer_pn or vendor_pn or "") or ""
    qty = None
    if quantity is not None and quantity != "":
        try:
            qty = float(quantity)
        except (TypeError, ValueError):
            qty = None
    return (pn, qty, source_line_number)


def refine_extracted_document(doc: ExtractedDocument, native_text: str) -> ExtractedDocument:
    """Python-owned corrections from document text. Never invents omitted lines."""
    text = native_text or ""
    if not text.strip() or doc.extraction_mode == ExtractionMode.FAILED:
        return doc

    doc_type, is_rev, supersedes, commitment = classify_document(text)
    if doc_type == DocumentType.RECEIPT_ONLY:
        doc.document_type = DocumentType.RECEIPT_ONLY
        doc.commitment_status = CommitmentStatus.RECEIPT_ACKNOWLEDGED_NO_SCHEDULE
        doc.is_revision = False
        doc.supersedes_all_prior_for_po = False
    elif doc_type == DocumentType.INVOICE:
        doc.document_type = DocumentType.INVOICE
    elif doc_type == DocumentType.REVISION:
        doc.document_type = DocumentType.REVISION
        doc.is_revision = True
        if supersedes:
            doc.supersedes_all_prior_for_po = True
        if doc.commitment_status == CommitmentStatus.UNKNOWN:
            doc.commitment_status = commitment

    currency = detect_currency(text, doc.document_currency)
    if currency:
        doc.document_currency = currency
        for ln in doc.lines:
            if not ln.currency:
                ln.currency = currency

    local = parse_document_text(text)
    local_by_pn_qty: dict[tuple, dict] = {}
    local_by_src: dict[int, dict] = {}
    for raw in local.get("lines") or []:
        key = _line_match_key(
            raw.get("customer_part_number"),
            raw.get("vendor_part_number"),
            raw.get("quantity"),
            raw.get("source_line_number"),
        )
        if key[0] or key[1] is not None:
            local_by_pn_qty[(key[0], key[1])] = raw
        if raw.get("source_line_number") is not None:
            local_by_src[int(raw["source_line_number"])] = raw

    for ln in doc.lines:
        raw = None
        key = _line_match_key(ln.customer_part_number, ln.vendor_part_number, ln.quantity, ln.source_line_number)
        if (key[0], key[1]) in local_by_pn_qty:
            raw = local_by_pn_qty[(key[0], key[1])]
        elif ln.source_line_number is not None:
            raw = local_by_src.get(int(ln.source_line_number))
        if not raw:
            continue
        if ln.unit_price is None and raw.get("unit_price") not in (None, ""):
            try:
                ln.unit_price = Decimal(str(raw["unit_price"]))
            except Exception:
                pass
        if not ln.currency and raw.get("currency"):
            ln.currency = str(raw["currency"]).upper()
        if (not ln.promise.raw_text) and raw.get("promise_raw_text"):
            ln.promise = parse_promise(raw.get("promise_raw_text"), header_hint=raw.get("date_type") or "")
    return doc


def _try_llm(client: LLMClient, pdf: PdfDocument, use_vision: bool) -> dict:
    last_err = None
    for attempt in range(config.LLM_MAX_RETRIES):
        hint = str(last_err) if last_err else None
        try:
            if use_vision:
                images = render_pages_png(pdf.path)
                payload = client.extract_from_images(images, retry_hint=hint)
            else:
                payload = client.extract_from_text(pdf.native_text, retry_hint=hint)
            # Light structural check
            if not isinstance(payload, dict):
                raise ExtractionError("Model did not return a JSON object")
            payload.setdefault("lines", [])
            return payload
        except Exception as exc:  # noqa: BLE001 — retry malformed responses only
            if is_auth_error(exc):
                raise AuthenticationFailed(_safe_exc(exc)) from exc
            last_err = exc
            logger.info("  LLM extraction attempt %s failed: %s", attempt + 1, _safe_exc(exc))
    raise ExtractionError(f"LLM extraction failed after retries: {last_err}")


def extract_one(
    path: Path,
    *,
    cache_dir: Path,
    refresh_cache: bool = False,
    offline: bool = False,
    llm: LLMClient | None = None,
) -> ExtractedDocument:
    pdf = read_pdf(path)
    model_name = (llm.model_name if llm else "local-parser-v1")

    if not refresh_cache:
        cached = load_cache(
            cache_dir,
            pdf.sha256,
            prompt_version=PROMPT_VERSION,
            schema_version=SCHEMA_VERSION,
            model_name=model_name,
        )
        if cached:
            refine_extracted_document(cached, pdf.native_text)
            return cached

    if offline and llm is None:
        # Offline: local parser still allowed for uncached docs (no API calls).
        pass
    elif offline and llm is not None:
        # Offline forbids API. If cache missed, fall back to local parser.
        llm = None

    use_llm = llm is not None and not offline

    try:
        if pdf.text_usable:
            if use_llm:
                try:
                    payload = _try_llm(llm, pdf, use_vision=False)
                    mode = ExtractionMode.NATIVE_TEXT
                except AuthenticationFailed:
                    raise
                except Exception as exc:  # noqa: BLE001
                    logger.warning(
                        "OpenAI text extraction failed for %s; falling back to local parser (%s)",
                        pdf.filename,
                        _safe_exc(exc),
                    )
                    payload = parse_document_text(pdf.native_text)
                    mode = ExtractionMode.LOCAL_TEXT
                    payload.setdefault("extraction_warnings", []).append(
                        "Fell back to local text parser after LLM error"
                    )
            else:
                payload = parse_document_text(pdf.native_text)
                mode = ExtractionMode.LOCAL_TEXT
        else:
            if use_llm:
                try:
                    payload = _try_llm(llm, pdf, use_vision=True)
                    mode = ExtractionMode.VISION
                except AuthenticationFailed:
                    raise
                except Exception as exc:  # noqa: BLE001
                    logger.warning(
                        "OpenAI vision extraction failed for %s; falling back to OCR (%s)",
                        pdf.filename,
                        _safe_exc(exc),
                    )
                    ocr_text = ocr_pdf(pdf.path)
                    if not ocr_text.strip():
                        raise ExtractionError(
                            "Vision LLM failed and OCR produced empty output"
                        ) from exc
                    payload = parse_document_text(ocr_text)
                    mode = ExtractionMode.OCR_LOCAL
                    payload.setdefault("extraction_warnings", []).append(
                        "Fell back to OCR after LLM vision error; native PDF text was empty"
                    )
            else:
                ocr_text = ocr_pdf(pdf.path)
                if not ocr_text.strip():
                    raise ExtractionError("Scan had no native text and OCR produced empty output")
                payload = parse_document_text(ocr_text)
                mode = ExtractionMode.OCR_LOCAL
                payload.setdefault("extraction_warnings", []).append("Extracted via OCR; native PDF text was empty")
    except AuthenticationFailed:
        raise
    except Exception as exc:  # noqa: BLE001
        doc = failed_document(pdf, f"Extraction failed: {_safe_exc(exc)}", ExtractionMode.FAILED, model_name)
        save_cache(cache_dir, doc, model_name=model_name)
        return doc

    try:
        doc = _payload_to_document(payload, pdf, mode=mode, model_name=model_name)
    except ValidationError as exc:
        doc = failed_document(pdf, f"Validation failed: {exc}", mode, model_name)
        save_cache(cache_dir, doc, model_name=model_name)
        return doc

    if not doc.document_currency:
        doc.document_currency = detect_currency(pdf.native_text, doc.document_currency)
    refine_extracted_document(doc, pdf.native_text)

    save_cache(cache_dir, doc, model_name=model_name)
    return doc


def extract_all(
    paths: list[Path],
    *,
    cache_dir: Path,
    refresh_cache: bool = False,
    offline: bool = False,
    logger_=None,
) -> list[ExtractedDocument]:
    log = logger_ or logger
    client = None if offline else build_llm_client()
    docs: list[ExtractedDocument] = []
    total = len(paths)
    for i, path in enumerate(paths, start=1):
        doc = extract_one(
            path,
            cache_dir=cache_dir,
            refresh_cache=refresh_cache,
            offline=offline,
            llm=client,
        )
        mode = "cached" if doc.cached else doc.extraction_mode.value.lower()
        if doc.cached:
            tag = "cached"
        elif doc.extraction_mode == ExtractionMode.VISION:
            tag = "scan - vision extraction"
        elif doc.extraction_mode == ExtractionMode.OCR_LOCAL:
            tag = "scan - ocr extraction"
        elif doc.extraction_mode == ExtractionMode.NATIVE_TEXT:
            tag = "native text - extracted"
        elif doc.extraction_mode == ExtractionMode.LOCAL_TEXT:
            tag = "native text - local parser"
        elif doc.extraction_mode == ExtractionMode.FAILED:
            tag = "FAILED"
        else:
            tag = mode
        log.info("[%s/%s] %s - %s - %s", i, total, path.name, tag, doc.document_type.value)
        docs.append(doc)
    return docs
