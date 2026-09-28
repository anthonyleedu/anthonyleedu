"""PO-line matching hierarchy tests."""

from datetime import date
from decimal import Decimal

from src.currency import FxTable
from src.matcher import match_document_to_po
from src.models import (
    ConfirmationLine,
    CrosswalkEntry,
    CrosswalkStatus,
    ExtractedDocument,
    MatchMethod,
    OpenPOLine,
)


def _po(**kwargs) -> OpenPOLine:
    base = dict(
        po_number="PO-1",
        vendor_id="V002",
        vendor_name="Heritage Cold Heading",
        line_number=1,
        our_pn="CHB-9472-3",
        our_description="Cold-headed blank for HX-9472-3",
        qty_ordered=Decimal("50000"),
        unit_price=Decimal("0.84"),
        required_date=date(2026, 6, 8),
    )
    base.update(kwargs)
    return OpenPOLine(**base)


def _doc(lines, **kwargs) -> ExtractedDocument:
    base = dict(
        source_file="h.pdf",
        source_sha256="h",
        vendor_name_raw="Heritage Cold Heading",
        po_number="PO-1",
        lines=lines,
    )
    base.update(kwargs)
    return ExtractedDocument(**base)


FX = FxTable([{"month": "2026-05", "currency": "USD", "rate_to_usd": 1}])


def test_exact_customer_pn():
    po = _po(our_pn="BAR-A286-375", vendor_id="V001", vendor_name="Apex Bar & Tube Co.")
    doc = _doc(
        [ConfirmationLine(customer_part_number="BAR-A286-375", quantity=2500, unit_price=Decimal("7.84"))],
        vendor_name_raw="Apex Bar & Tube Co.",
    )
    matches, _ = match_document_to_po(doc, [po], crosswalk=[], fx=FX, expected_vendor_id="V001", expected_vendor_name="Apex")
    assert matches[0].method == MatchMethod.EXACT_CUSTOMER_PN
    assert matches[0].po_line.our_pn == "BAR-A286-375"


def test_approved_crosswalk():
    po = _po()
    xw = [
        CrosswalkEntry(
            vendor_id="V002",
            vendor_name="Heritage Cold Heading",
            vendor_pn="APH-441",
            beacon_pn="CHB-9472-3",
            evidence_count=30,
            vendor_pn_total_count=30,
            purity=1.0,
            status=CrosswalkStatus.APPROVED,
        )
    ]
    doc = _doc([ConfirmationLine(vendor_part_number="APH-441", quantity=50000, unit_price=Decimal("0.84"))])
    matches, _ = match_document_to_po(doc, [po], crosswalk=xw, fx=FX, expected_vendor_id="V002", expected_vendor_name="Heritage")
    assert matches[0].method == MatchMethod.APPROVED_VENDOR_CROSSWALK
    assert matches[0].explanation == "Approved crosswalk APH-441 maps to CHB-9472-3"


def test_vendor_pn_equals_beacon_pn():
    po = _po(our_pn="CHB-7715", line_number=2)
    doc = _doc([ConfirmationLine(vendor_part_number="CHB-7715", quantity=25000, unit_price=Decimal("0.62"))])
    matches, _ = match_document_to_po(doc, [po], crosswalk=[], fx=FX, expected_vendor_id="V002", expected_vendor_name="Heritage")
    assert matches[0].method == MatchMethod.EXACT_VENDOR_PN_TO_BEACON_PN


def test_contextual_price_qty_matching():
    po = _po(our_pn="CHB-9472-4", qty_ordered=Decimal("15000"), unit_price=Decimal("0.86"))
    doc = _doc([ConfirmationLine(vendor_part_number="APH-441-OS", quantity=15000, unit_price=Decimal("0.86"))])
    matches, _ = match_document_to_po(doc, [po], crosswalk=[], fx=FX, expected_vendor_id="V002", expected_vendor_name="Heritage")
    assert matches[0].method in {MatchMethod.RULE_BASED, MatchMethod.SINGLE_REMAINING_LINE}
    assert matches[0].po_line.our_pn == "CHB-9472-4"


def test_ambiguous_match_returns_review():
    po_a = _po(our_pn="AAA", line_number=1, qty_ordered=Decimal("100"), unit_price=Decimal("1.00"), our_description="alpha")
    po_b = _po(our_pn="BBB", line_number=2, qty_ordered=Decimal("100"), unit_price=Decimal("1.00"), our_description="beta")
    doc = _doc([ConfirmationLine(vendor_part_number="ZZZ", quantity=100, unit_price=Decimal("1.00"), description="(see PO for description)")])
    matches, issues = match_document_to_po(
        doc, [po_a, po_b], crosswalk=[], fx=FX, expected_vendor_id="V002", expected_vendor_name="Heritage"
    )
    assert any(m.method in {MatchMethod.MANUAL_REVIEW, MatchMethod.UNMATCHED} for m in matches)
    assert "PART_MATCH_UNCERTAIN" in issues or any(m.method == MatchMethod.UNMATCHED for m in matches)
