"""Deterministic PO-line reconciliation.

AI has already extracted document facts. This module never calls an LLM.
"""

from __future__ import annotations

from collections import defaultdict
from datetime import date
from decimal import Decimal

from src.config import (
    BASE_CURRENCY,
    FOREIGN_CURRENCY_PRICE_PCT_TOL,
    LATE_YELLOW_MAX_DAYS,
    MATERIAL_PRICE_IMPACT,
    MATERIAL_PRICE_INCREASE_PCT,
    PRICE_ABS_TOL,
    QTY_OVER_DEFAULT_RED,
    SAME_CURRENCY_PRICE_PCT_TOL,
)
from src.currency import FxTable
from src.document_versions import ResolvedDocuments
from src.matcher import match_document_to_po
from src.models import (
    CommitmentStatus,
    ConfirmationLine,
    CrosswalkEntry,
    DateType,
    DocumentType,
    ExtractedDocument,
    MatchMethod,
    MatchResult,
    OpenPOLine,
    ReconciledLine,
    Severity,
    VendorMasterRow,
)
from src.normalize import compare_promise_to_required, round_money, vendor_key
from src.part_crosswalk import propose_from_matches


SEVERITY_ORDER = {Severity.RED: 0, Severity.YELLOW: 1, Severity.INFO: 2, Severity.GREEN: 3}


def _dec(value) -> Decimal | None:
    if value is None:
        return None
    return Decimal(str(value))


def _action_for(row: ReconciledLine) -> str:
    codes = set(row.issue_codes)
    po = row.po_number
    vendor = row.vendor_name or "the vendor"
    pn = row.beacon_pn or row.vendor_pn or "the part"
    line = row.po_line_number
    req = row.required_date.isoformat() if row.required_date else "the required date"

    if "UNKNOWN_PO" in codes:
        return (
            f"Verify whether this acknowledgment belongs to Beacon and whether PO number "
            f"{po} is correct."
        )
    if "VENDOR_MISMATCH" in codes:
        return f"Confirm the document vendor matches the vendor on {po}."
    if "NO_FORMAL_ACK_FOUND" in codes:
        return (
            f"No formal acknowledgment found for {po}"
            + (f" line {line}" if line else "")
            + ". Request an order confirmation (the invoice/shipping notice is not sufficient)."
        )
    if "MISSING_PO_LINE" in codes:
        qty = f"{row.qty_ordered:,.0f}" if row.qty_ordered is not None else "ordered"
        return f"Ask {vendor} to confirm {po} line {line}, {pn}, qty {qty}, required {req}."
    if "QTY_SHORT" in codes:
        short = None
        if row.qty_ordered is not None and row.qty_confirmed is not None:
            short = row.qty_ordered - row.qty_confirmed
        extra = f" {short:,.0f}-unit" if short is not None else ""
        return f"Resolve{extra} short confirmation on {po} line {line} ({pn}) before release."
    if "QTY_OVER" in codes:
        over = None
        if row.qty_ordered is not None and row.qty_confirmed is not None:
            over = row.qty_confirmed - row.qty_ordered
        extra = f" {over:,.0f}-unit" if over is not None else ""
        return f"Resolve{extra} over-commitment on {po} line {line} ({pn}); supplier committed unauthorized excess quantity."
    if "PROMISE_MISSING" in codes or row.document_type == DocumentType.RECEIPT_ONLY:
        return f"Request committed quantity and promise date for {po} line {line} ({pn})."
    if "PROMISE_LATE" in codes:
        days = row.days_late if row.days_late is not None else row.minimum_late_days
        day_txt = f"{days} days late" if days is not None else "later than required"
        return f"Ask {vendor} whether required date {req} can be recovered; current promise is {day_txt}."
    if "PROMISE_WINDOW_STRADDLES_REQUIRED" in codes:
        return (
            f"Delivery window for {po} line {line} straddles required date {req}. "
            f"Ask {vendor} to commit a date on or before {req}."
        )
    if "PRICE_HIGH" in codes:
        diff = row.unit_price_difference
        pct = row.unit_price_difference_pct
        if diff is not None and pct is not None:
            return f"Confirm acceptance of ${diff:.4f}/unit increase (~{pct:.2%}) on {po} line {line}."
        return f"Confirm the confirmed unit price increase on {po} line {line}."
    if "PRICE_LOW" in codes:
        return f"Confirm whether the lower confirmed price on {po} line {line} is intentional."
    if "PRICE_NOT_STATED" in codes:
        return f"Confirmation for {po} line {line} omits price; do not treat this as price acceptance."
    if "PART_MATCH_UNCERTAIN" in codes:
        return f"Review part-number match on {po} line {line} before releasing."
    if "EXTRA_CONFIRMATION_LINE" in codes:
        return f"Acknowledgment contains an extra line not on {po}; confirm whether it belongs on another PO."
    if "CURRENCY_UNCERTAIN" in codes:
        return f"FX conversion was unavailable for {po} line {line}; review raw confirmed price before accepting."
    if "DATE_SEMANTICS_WARNING" in codes:
        return (
            f"{vendor} stated a ship date (not a Beacon receipt/promise date) on {po}. "
            "Confirm whether the ship date will still meet the required date at Beacon."
        )
    if "INVOICE_NOT_ACK" in codes:
        return f"Document for {po} is an invoice, not a formal acknowledgment. Request a confirmation if still open."
    if "EXTRACTION_WARNING" in codes:
        return f"Review extraction warnings for {po} before relying on the parsed values."
    return f"Review {po} line {line} with {vendor}."


def _severity_from_codes(codes: list[str], *, price_pct: Decimal | None, price_impact: Decimal | None, days_late: int | None) -> Severity:
    red_codes = {
        "UNKNOWN_PO",
        "VENDOR_MISMATCH",
        "NO_FORMAL_ACK_FOUND",
        "MISSING_PO_LINE",
        "QTY_SHORT",
        "PROMISE_MISSING",
        "EXTRACTION_WARNING",  # only if it blocked reconciliation — handled separately
    }
    if QTY_OVER_DEFAULT_RED:
        red_codes.add("QTY_OVER")
    if "PRICE_HIGH" in codes:
        material = False
        if price_pct is not None and price_pct >= MATERIAL_PRICE_INCREASE_PCT:
            material = True
        if price_impact is not None and price_impact >= MATERIAL_PRICE_IMPACT:
            material = True
        if material:
            red_codes.add("PRICE_HIGH")
    if "PROMISE_LATE" in codes:
        if days_late is None or days_late >= 3:
            red_codes.add("PROMISE_LATE")
    if "PART_MATCH_UNCERTAIN" in codes and any(
        c in codes for c in ("MISSING_PO_LINE", "QTY_SHORT", "UNKNOWN_PO")
    ):
        red_codes.add("PART_MATCH_UNCERTAIN")

    blocking_extract = "EXTRACTION_WARNING" in codes and (
        "UNKNOWN_PO" in codes or "MISSING_PO_LINE" in codes or not codes
    )

    if any(c in red_codes for c in codes) or blocking_extract:
        # PRICE_HIGH may be YELLOW if not material
        if set(codes) <= {"PRICE_HIGH", "DATE_SEMANTICS_WARNING", "PRICE_NOT_STATED"} and "PRICE_HIGH" not in red_codes:
            pass
        else:
            if any(c in red_codes for c in codes):
                return Severity.RED

    yellow_codes = {
        "PROMISE_WINDOW_STRADDLES_REQUIRED",
        "PART_MATCH_UNCERTAIN",
        "PRICE_NOT_STATED",
        "PRICE_HIGH",
        "PRICE_LOW",
        "DATE_SEMANTICS_WARNING",
        "CURRENCY_UNCERTAIN",
        "EXTRA_CONFIRMATION_LINE",
        "PROMISE_LATE",
    }
    if any(c in yellow_codes for c in codes):
        return Severity.YELLOW

    info_codes = {
        "DOCUMENT_REVISION",
        "DOCUMENT_SUPERSEDED",
        "INVOICE_NOT_ACK",
    }
    if any(c in info_codes for c in codes):
        return Severity.INFO
    if codes:
        return Severity.INFO
    return Severity.GREEN


def _sort_tuple(row: ReconciledLine) -> tuple:
    missing = 1 if any(c in row.issue_codes for c in ("MISSING_PO_LINE", "QTY_SHORT", "NO_FORMAL_ACK_FOUND", "UNKNOWN_PO")) else 0
    impact = abs(row.price_impact_on_order or Decimal("0"))
    late = row.days_late if row.days_late is not None else row.minimum_late_days or 0
    return (
        SEVERITY_ORDER.get(row.severity, 9),
        0 if missing else 1,
        -float(impact),
        -int(late),
        row.po_number,
        row.po_line_number or 0,
    )


def projected_full_qty_promise(
    commitments: list[tuple[Decimal, date | None]],
    qty_ordered: Decimal,
) -> date | None:
    """Sort by commitment end date, cumulative qty, first date covering ordered qty."""
    usable = [(q, d) for q, d in commitments if q is not None and d is not None]
    if not usable:
        return None
    usable.sort(key=lambda t: t[1])
    running = Decimal("0")
    for q, d in usable:
        running += q
        if running >= qty_ordered:
            return d
    return None


def _vendor_contact(vendors: list[VendorMasterRow], vendor_id: str | None, vendor_name: str | None) -> str | None:
    if vendor_id:
        for v in vendors:
            if v.vendor_id == vendor_id:
                return v.ap_email
    if vendor_name:
        key = vendor_key(vendor_name)
        for v in vendors:
            if vendor_key(v.vendor_name) == key:
                return v.ap_email
    return None


def reconcile(
    open_pos: list[OpenPOLine],
    vendors: list[VendorMasterRow],
    resolved: ResolvedDocuments,
    *,
    crosswalk: list[CrosswalkEntry],
    fx: FxTable,
) -> tuple[list[ReconciledLine], list[MatchResult], list[CrosswalkEntry]]:
    by_po: dict[str, list[OpenPOLine]] = defaultdict(list)
    for ln in open_pos:
        by_po[ln.po_number].append(ln)

    all_matches: list[MatchResult] = []
    proposed_pairs: list[tuple[str, str, str, str]] = []
    rows: list[ReconciledLine] = []

    # Index active/info docs
    docs_by_po: dict[str, list[ExtractedDocument]] = defaultdict(list)
    unknown_docs: list[ExtractedDocument] = []
    for d in resolved.all_documents:
        if d.po_number and d.po_number in by_po:
            docs_by_po[d.po_number].append(d)
        elif d.po_number and d.po_number not in by_po:
            unknown_docs.append(d)
        elif not d.po_number:
            unknown_docs.append(d)

    # Unknown POs
    for doc in unknown_docs:
        codes = ["UNKNOWN_PO"]
        if doc.extraction_warnings:
            codes.append("EXTRACTION_WARNING")
        if doc.document_type == DocumentType.INVOICE:
            codes.append("INVOICE_NOT_ACK")
        row = ReconciledLine(
            po_number=doc.po_number or "(missing PO number)",
            vendor_name=doc.vendor_name_raw,
            vendor_contact=_vendor_contact(vendors, None, doc.vendor_name_raw),
            description="; ".join(filter(None, [ln.description for ln in doc.lines])) or None,
            qty_confirmed=_dec(sum((ln.quantity or 0) for ln in doc.lines)) if doc.lines else None,
            source_file=doc.source_file,
            source_document_number=doc.document_number,
            document_date=doc.document_date,
            document_status=doc.lifecycle.value,
            document_type=doc.document_type,
            source_sha256=doc.source_sha256,
            issue_codes=codes,
            unknown_po=True,
            extraction_warnings=list(doc.extraction_warnings),
        )
        row.severity = _severity_from_codes(codes, price_pct=None, price_impact=None, days_late=None)
        row.suggested_action = _action_for(row)
        row.is_actionable = row.severity in {Severity.RED, Severity.YELLOW}
        row.sort_key = _sort_tuple(row)
        rows.append(row)

    # Per PO line
    for po_number, po_lines in by_po.items():
        vendor_id = po_lines[0].vendor_id
        vendor_name = po_lines[0].vendor_name
        contact = _vendor_contact(vendors, vendor_id, vendor_name)
        po_docs = docs_by_po.get(po_number, [])
        active = [d for d in po_docs if d.lifecycle.value == "ACTIVE"]
        informational = [d for d in po_docs if d.lifecycle.value == "INFORMATIONAL"]
        superseded = [d for d in po_docs if d.lifecycle.value == "SUPERSEDED"]

        formal_active = [
            d
            for d in active
            if d.document_type in {DocumentType.ACKNOWLEDGEMENT, DocumentType.REVISION, DocumentType.RECEIPT_ONLY}
        ]
        receipt_only = [d for d in formal_active if d.document_type == DocumentType.RECEIPT_ONLY or d.commitment_status == CommitmentStatus.RECEIPT_ACKNOWLEDGED_NO_SCHEDULE]

        # Match each active formal doc independently (split acks must aggregate).
        matches_by_line: dict[int, list[MatchResult]] = defaultdict(list)
        extra_unmatched: list[MatchResult] = []
        doc_issues: list[str] = []
        for doc in [d for d in active if d.document_type != DocumentType.INVOICE]:
            matches, issues = match_document_to_po(
                doc,
                po_lines,
                crosswalk=crosswalk,
                fx=fx,
                expected_vendor_id=vendor_id,
                expected_vendor_name=vendor_name,
            )
            doc_issues.extend(issues)
            all_matches.extend(matches)
            for m in matches:
                if m.po_line is not None:
                    matches_by_line[m.po_line.line_number].append(m)
                    conf = m.confirmation_line
                    if (
                        conf
                        and conf.vendor_part_number
                        and m.po_line.our_pn
                        and normalize_ne(conf.vendor_part_number, m.po_line.our_pn)
                        and m.method
                        in {
                            MatchMethod.RULE_BASED,
                            MatchMethod.SINGLE_REMAINING_LINE,
                            MatchMethod.APPROVED_VENDOR_CROSSWALK,
                        }
                    ):
                        proposed_pairs.append(
                            (vendor_id or "", vendor_name or "", conf.vendor_part_number, m.po_line.our_pn)
                        )
                else:
                    extra_unmatched.append(m)

        for po in po_lines:
            line_matches = matches_by_line.get(po.line_number, [])
            codes: list[str] = []
            if "VENDOR_MISMATCH" in doc_issues:
                codes.append("VENDOR_MISMATCH")

            if not formal_active:
                if informational:
                    codes.append("NO_FORMAL_ACK_FOUND")
                    codes.append("INVOICE_NOT_ACK")
                else:
                    codes.append("NO_FORMAL_ACK_FOUND")

            if formal_active and not line_matches:
                if receipt_only and all(not d.lines for d in receipt_only) and not any(d.lines for d in formal_active):
                    codes.append("PROMISE_MISSING")
                elif any(d.commitment_status == CommitmentStatus.RECEIPT_ACKNOWLEDGED_NO_SCHEDULE for d in formal_active) and not any(
                    d.lines for d in formal_active
                ):
                    codes.append("PROMISE_MISSING")
                else:
                    codes.append("MISSING_PO_LINE")

            qty_confirmed = Decimal("0")
            commitments: list[tuple[Decimal, date | None]] = []
            prices: list[Decimal] = []
            price_raws: list[tuple[Decimal, str | None]] = []
            schedules: list[str] = []
            source_files: list[str] = []
            warnings: list[str] = []
            match_method = None
            match_conf = None
            match_expl = None
            vendor_pn = None
            commitment_type = None
            promise_start = None
            promise_end = None
            promise_raw = None
            conf_currency = None
            fx_rate = None
            confirmed_usd = None
            fx_warning = None
            doc_number = None
            doc_date = None
            doc_status = None
            doc_type = None
            sha = None

            for m in line_matches:
                conf = m.confirmation_line
                doc = m.source_document
                if not conf or not doc:
                    continue
                source_files.append(doc.source_file)
                warnings.extend(doc.extraction_warnings)
                if doc.is_revision:
                    codes.append("DOCUMENT_REVISION")
                q = _dec(conf.quantity) or Decimal("0")
                qty_confirmed += q
                end_d = conf.promise.latest
                commitments.append((q, end_d))
                if conf.promise.raw_text or end_d:
                    schedules.append(
                        f"{q:g} on {conf.promise.raw_text or (end_d.isoformat() if end_d else '?')} "
                        f"[{doc.source_file.split('/')[-1]}]"
                    )
                if conf.unit_price is not None:
                    prices.append(conf.unit_price)
                    price_raws.append((conf.unit_price, conf.currency or doc.document_currency))
                if vendor_pn is None:
                    vendor_pn = conf.vendor_part_number or conf.customer_part_number
                if match_method is None or (m.confidence or 0) > (match_conf or 0):
                    match_method = m.method
                    match_conf = m.confidence
                    match_expl = m.explanation
                if commitment_type is None:
                    commitment_type = conf.promise.date_type
                if conf.promise.start_date:
                    promise_start = min(filter(None, [promise_start, conf.promise.start_date])) if promise_start else conf.promise.start_date
                if conf.promise.end_date:
                    promise_end = max(filter(None, [promise_end, conf.promise.end_date])) if promise_end else conf.promise.end_date
                if promise_raw is None:
                    promise_raw = conf.promise.raw_text
                doc_number = doc.document_number
                doc_date = doc.document_date
                doc_status = doc.lifecycle.value
                doc_type = doc.document_type
                sha = doc.source_sha256
                if m.method == MatchMethod.MANUAL_REVIEW:
                    codes.append("PART_MATCH_UNCERTAIN")

            if superseded:
                codes.append("DOCUMENT_SUPERSEDED")

            qty_confirmed_out = qty_confirmed if line_matches else None
            qty_var = (qty_confirmed - po.qty_ordered) if qty_confirmed_out is not None else None
            if qty_var is not None:
                if qty_var < 0:
                    codes.append("QTY_SHORT")
                elif qty_var > 0:
                    codes.append("QTY_OVER")

            price_stated = bool(price_raws)
            confirmed_price_raw = price_raws[0][0] if price_raws else None
            conf_currency = price_raws[0][1] if price_raws else None
            if line_matches and not price_stated:
                codes.append("PRICE_NOT_STATED")

            if price_stated:
                usd, rate, _month, fx_warning = fx.to_usd(
                    confirmed_price_raw,
                    conf_currency,
                    doc_date,
                )
                fx_rate = rate
                confirmed_usd = usd
                if fx_warning:
                    codes.append("CURRENCY_UNCERTAIN")
                if confirmed_usd is not None:
                    diff = confirmed_usd - po.unit_price
                    pct = (diff / po.unit_price) if po.unit_price else None
                    foreign = bool(conf_currency and conf_currency.upper() != BASE_CURRENCY)
                    tol = FOREIGN_CURRENCY_PRICE_PCT_TOL if foreign else SAME_CURRENCY_PRICE_PCT_TOL
                    if pct is not None and abs(diff) > PRICE_ABS_TOL and abs(pct) > tol:
                        codes.append("PRICE_HIGH" if diff > 0 else "PRICE_LOW")
                else:
                    codes.append("CURRENCY_UNCERTAIN")
            else:
                diff = None
                pct = None

            # Dates
            days_late = None
            min_late = None
            max_late = None
            proj = None
            if line_matches:
                # If any commitment lacks a usable date while qty is confirmed
                usable_dates = [d for _, d in commitments if d is not None]
                if not usable_dates and "PROMISE_MISSING" not in codes and "MISSING_PO_LINE" not in codes:
                    codes.append("PROMISE_MISSING")
                else:
                    # Split shipments: judge lateness by the date the FULL ordered
                    # quantity is promised, not the earliest partial.
                    proj = projected_full_qty_promise(commitments, po.qty_ordered)
                    from src.models import PromiseWindow

                    dated = [d for _, d in commitments if d is not None]
                    is_split = len(dated) > 1 and len({d.isoformat() for d in dated}) > 1
                    if is_split and proj is not None:
                        window = PromiseWindow(
                            raw_text=promise_raw,
                            start_date=proj,
                            end_date=proj,
                            date_type=commitment_type or DateType.UNKNOWN,
                        )
                    else:
                        window = PromiseWindow(
                            raw_text=promise_raw,
                            start_date=promise_start,
                            end_date=promise_end or promise_start,
                            date_type=commitment_type or DateType.UNKNOWN,
                        )
                    cmp = compare_promise_to_required(window, po.required_date)
                    if cmp["status"] == "LATE":
                        codes.append("PROMISE_LATE")
                        days_late = cmp["days_late"]
                        min_late = cmp["minimum_late_days"]
                        max_late = cmp["maximum_late_days"]
                    elif cmp["status"] == "STRADDLES":
                        codes.append("PROMISE_WINDOW_STRADDLES_REQUIRED")
                        min_late = cmp["minimum_late_days"]
                        max_late = cmp["maximum_late_days"]
                    elif cmp["status"] == "ON_TIME":
                        days_late = 0
                        min_late = 0
                        max_late = 0
                    elif cmp["status"] == "MISSING" and "PROMISE_MISSING" not in codes:
                        codes.append("PROMISE_MISSING")

                if commitment_type == DateType.SHIP_DATE:
                    codes.append("DATE_SEMANTICS_WARNING")

            if any(d.extraction_warnings for d in po_docs if d.lifecycle.value == "ACTIVE"):
                if any("failed" in (w.lower()) for d in po_docs for w in d.extraction_warnings):
                    codes.append("EXTRACTION_WARNING")

            # Unique codes, stable order
            seen = set()
            uniq = []
            for c in codes:
                if c not in seen:
                    seen.add(c)
                    uniq.append(c)
            codes = uniq

            impact = None
            if price_stated and confirmed_usd is not None:
                impact = round_money((confirmed_usd - po.unit_price) * po.qty_ordered, 2)

            row = ReconciledLine(
                po_number=po.po_number,
                po_line_number=po.line_number,
                vendor_id=vendor_id,
                vendor_name=vendor_name,
                vendor_contact=contact,
                beacon_pn=po.our_pn,
                vendor_pn=vendor_pn,
                description=po.our_description,
                qty_ordered=po.qty_ordered,
                qty_confirmed=qty_confirmed_out,
                qty_variance=qty_var,
                po_unit_price=po.unit_price,
                po_currency_assumed=po.currency_assumed,
                confirmed_unit_price_raw=confirmed_price_raw,
                confirmed_currency=conf_currency,
                fx_rate_used=fx_rate,
                confirmed_price_usd=confirmed_usd,
                unit_price_difference=(confirmed_usd - po.unit_price) if confirmed_usd is not None else None,
                unit_price_difference_pct=pct,
                price_impact_on_order=impact,
                price_stated=price_stated,
                required_date=po.required_date,
                commitment_date_type=commitment_type,
                promise_start=promise_start,
                promise_end=promise_end,
                promise_raw=promise_raw,
                projected_full_qty_promise_date=proj,
                days_late=days_late,
                minimum_late_days=min_late,
                maximum_late_days=max_late,
                match_method=match_method,
                match_confidence=match_conf,
                match_explanation=match_expl,
                source_file=source_files[0] if source_files else (formal_active[0].source_file if formal_active else (informational[0].source_file if informational else None)),
                source_document_number=doc_number or (formal_active[0].document_number if formal_active else None),
                document_date=doc_date or (formal_active[0].document_date if formal_active else None),
                document_status=doc_status or (formal_active[0].lifecycle.value if formal_active else ("INFORMATIONAL" if informational else None)),
                document_type=doc_type or (formal_active[0].document_type if formal_active else (informational[0].document_type if informational else None)),
                source_sha256=sha,
                issue_codes=codes,
                extra_source_files=source_files,
                schedule_details=schedules,
                extraction_warnings=warnings,
            )
            row.severity = _severity_from_codes(
                codes,
                price_pct=abs(pct) if pct is not None else None,
                price_impact=abs(impact) if impact is not None else None,
                days_late=days_late if days_late is not None else min_late,
            )
            # PRICE_NOT_STATED alone on a fully matched line -> INFO if no other issues
            leftover = set(codes) - {"PRICE_NOT_STATED", "DOCUMENT_REVISION", "DOCUMENT_SUPERSEDED"}
            if codes == ["PRICE_NOT_STATED"] or leftover == set() and "PRICE_NOT_STATED" in codes and row.severity != Severity.RED:
                if row.severity != Severity.RED and "QTY_SHORT" not in codes:
                    if set(codes) <= {"PRICE_NOT_STATED", "DOCUMENT_REVISION", "DOCUMENT_SUPERSEDED"}:
                        row.severity = Severity.INFO
            row.suggested_action = _action_for(row)
            row.is_actionable = row.severity in {Severity.RED, Severity.YELLOW}
            row.sort_key = _sort_tuple(row)
            rows.append(row)

        for m in extra_unmatched:
            doc = m.source_document
            conf = m.confirmation_line
            extra_row = ReconciledLine(
                po_number=po_number,
                vendor_id=vendor_id,
                vendor_name=vendor_name,
                vendor_contact=contact,
                vendor_pn=conf.vendor_part_number if conf else None,
                beacon_pn=conf.customer_part_number if conf else None,
                description=conf.description if conf else None,
                qty_confirmed=_dec(conf.quantity) if conf else None,
                confirmed_unit_price_raw=conf.unit_price if conf else None,
                source_file=doc.source_file if doc else None,
                source_document_number=doc.document_number if doc else None,
                document_date=doc.document_date if doc else None,
                document_status=doc.lifecycle.value if doc else None,
                document_type=doc.document_type if doc else None,
                match_method=m.method,
                match_confidence=m.confidence,
                match_explanation=m.explanation,
                issue_codes=["EXTRA_CONFIRMATION_LINE"]
                + (["PART_MATCH_UNCERTAIN"] if m.method == MatchMethod.MANUAL_REVIEW else []),
            )
            extra_row.severity = Severity.YELLOW
            extra_row.suggested_action = _action_for(extra_row)
            extra_row.is_actionable = True
            extra_row.sort_key = _sort_tuple(extra_row)
            rows.append(extra_row)

    proposed = propose_from_matches(resolved.all_documents, proposed_pairs, crosswalk)
    rows.sort(key=lambda r: r.sort_key)
    return rows, all_matches, proposed


def normalize_ne(a: str, b: str) -> bool:
    from src.normalize import normalize_part_number

    return normalize_part_number(a) != normalize_part_number(b)
