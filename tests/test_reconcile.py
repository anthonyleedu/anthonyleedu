"""Reconciliation business-rule tests."""

from datetime import date
from decimal import Decimal

from src.currency import FxTable
from src.document_versions import resolve_active_documents
from src.models import (
    CommitmentStatus,
    ConfirmationLine,
    DocumentType,
    ExtractedDocument,
    OpenPOLine,
    VendorMasterRow,
)
from src.normalize import parse_promise
from src.reconcile import projected_full_qty_promise, reconcile


FX = FxTable(
    [
        {"month": "2026-05", "currency": "USD", "rate_to_usd": 1},
        {"month": "2026-05", "currency": "EUR", "rate_to_usd": 1.0902},
    ]
)
VENDORS = [VendorMasterRow(vendor_id="V001", vendor_name="Apex Bar & Tube Co.", ap_email="orders@apexbar.com")]


def _po(**kwargs) -> OpenPOLine:
    base = dict(
        po_number="PO-4500050001",
        vendor_id="V001",
        vendor_name="Apex Bar & Tube Co.",
        line_number=1,
        our_pn="BAR-A286-375",
        our_description="A286 bar stock, 0.375 dia, mill-cert",
        qty_ordered=Decimal("2500"),
        unit_price=Decimal("7.84"),
        required_date=date(2026, 6, 10),
    )
    base.update(kwargs)
    return OpenPOLine(**base)


def _doc(**kwargs) -> ExtractedDocument:
    base = dict(
        source_file="a.pdf",
        source_sha256="a",
        vendor_name_raw="Apex Bar & Tube Co.",
        po_number="PO-4500050001",
        document_type=DocumentType.ACKNOWLEDGEMENT,
        document_date=date(2026, 5, 17),
        document_currency="USD",
        lines=[],
    )
    base.update(kwargs)
    return ExtractedDocument(**base)


def _run(pos, docs):
    resolved = resolve_active_documents(docs)
    rows, _, _ = reconcile(pos, VENDORS, resolved, crosswalk=[], fx=FX)
    return rows


def test_missing_line():
    pos = [
        _po(line_number=1, our_pn="BAR-A286-375"),
        _po(line_number=2, our_pn="BAR-A286-250", qty_ordered=Decimal("1500"), unit_price=Decimal("4.71")),
    ]
    docs = [
        _doc(
            lines=[
                ConfirmationLine(
                    customer_part_number="BAR-A286-375",
                    quantity=2500,
                    unit_price=Decimal("7.84"),
                    promise=parse_promise("06/10/2026"),
                )
            ]
        )
    ]
    rows = _run(pos, docs)
    missing = [r for r in rows if r.beacon_pn == "BAR-A286-250"][0]
    assert "MISSING_PO_LINE" in missing.issue_codes
    assert missing.severity.value == "RED"


def test_quantity_short():
    pos = [_po(po_number="PO-4500050007", our_pn="BAR-A286-250", qty_ordered=Decimal("1500"), unit_price=Decimal("4.71"))]
    docs = [
        _doc(
            po_number="PO-4500050007",
            lines=[
                ConfirmationLine(
                    customer_part_number="BAR-A286-250",
                    quantity=1425,
                    unit_price=Decimal("4.71"),
                    promise=parse_promise("06/15/2026"),
                )
            ],
        )
    ]
    row = [r for r in _run(pos, docs) if r.po_number == "PO-4500050007"][0]
    assert row.qty_variance == Decimal("-75")
    assert "QTY_SHORT" in row.issue_codes


def test_quantity_over():
    pos = [_po(qty_ordered=Decimal("100"))]
    docs = [
        _doc(
            lines=[
                ConfirmationLine(
                    customer_part_number="BAR-A286-375",
                    quantity=150,
                    unit_price=Decimal("7.84"),
                    promise=parse_promise("06/10/2026"),
                )
            ]
        )
    ]
    row = _run(pos, docs)[0]
    assert "QTY_OVER" in row.issue_codes
    assert row.severity.value == "RED"


def test_price_increase():
    pos = [_po(po_number="PO-4500050002", our_pn="BAR-CRES-250", qty_ordered=Decimal("1000"), unit_price=Decimal("3.92"), required_date=date(2026, 5, 21))]
    docs = [
        _doc(
            po_number="PO-4500050002",
            lines=[
                ConfirmationLine(
                    customer_part_number="BAR-CRES-250",
                    quantity=1000,
                    unit_price=Decimal("4.0102"),
                    currency="USD",
                    promise=parse_promise("05/21/2026"),
                )
            ],
        )
    ]
    row = _run(pos, docs)[0]
    assert abs(row.unit_price_difference - Decimal("0.0902")) < Decimal("0.00001")
    assert abs(row.unit_price_difference_pct - Decimal("0.023010204081632653")) < Decimal("0.0002")
    assert "PRICE_HIGH" in row.issue_codes


def test_missing_price_not_treated_as_match():
    pos = [_po()]
    docs = [
        _doc(
            lines=[
                ConfirmationLine(
                    customer_part_number="BAR-A286-375",
                    quantity=2500,
                    promise=parse_promise("06/10/2026"),
                )
            ]
        )
    ]
    row = _run(pos, docs)[0]
    assert "PRICE_NOT_STATED" in row.issue_codes
    assert row.price_stated is False


def test_late_date():
    pos = [_po(required_date=date(2026, 5, 21), our_pn="HT-PRECIP", qty_ordered=Decimal("500"), unit_price=Decimal("2.4"), po_number="PO-X")]
    docs = [
        _doc(
            po_number="PO-X",
            vendor_name_raw="Continental Quality Heat Treat",
            lines=[
                ConfirmationLine(
                    description="Precipitation heat treat A286, per lb",
                    quantity=500,
                    promise=parse_promise("06/11/2026"),
                )
            ],
        )
    ]
    row = _run(pos, docs)[0]
    assert "PROMISE_LATE" in row.issue_codes
    assert row.days_late == 21


def test_missing_schedule():
    pos = [_po(po_number="PO-4500050019", our_pn="PLAT-PASV", qty_ordered=Decimal("25000"), unit_price=Decimal("0.08"))]
    docs = [
        _doc(
            po_number="PO-4500050019",
            document_type=DocumentType.RECEIPT_ONLY,
            commitment_status=CommitmentStatus.RECEIPT_ACKNOWLEDGED_NO_SCHEDULE,
            lines=[],
        )
    ]
    row = _run(pos, docs)[0]
    assert "PROMISE_MISSING" in row.issue_codes
    assert "MISSING_PO_LINE" not in row.issue_codes


def test_split_quantities_projected_date():
    ordered = Decimal("2500")
    proj = projected_full_qty_promise(
        [
            (Decimal("1500"), date(2026, 5, 17)),
            (Decimal("1000"), date(2026, 5, 24)),
        ],
        ordered,
    )
    assert proj == date(2026, 5, 24)

    pos = [
        _po(
            po_number="PO-4500050016",
            vendor_id="V003",
            vendor_name="Liberty Surface Finishing",
            our_pn="PLAT-PASV",
            qty_ordered=Decimal("2500"),
            unit_price=Decimal("0.08"),
            required_date=date(2026, 5, 17),
        )
    ]
    docs = [
        _doc(
            source_file="p1.pdf",
            source_sha256="p1",
            po_number="PO-4500050016",
            vendor_name_raw="Liberty Surface Finishing",
            partial_sequence=1,
            partial_total=2,
            lines=[ConfirmationLine(description="Passivation per AMS 2700, per piece", quantity=1500, promise=parse_promise("05/17/2026"))],
        ),
        _doc(
            source_file="p2.pdf",
            source_sha256="p2",
            po_number="PO-4500050016",
            vendor_name_raw="Liberty Surface Finishing",
            document_date=date(2026, 5, 16),
            partial_sequence=2,
            partial_total=2,
            lines=[ConfirmationLine(description="Passivation per AMS 2700, per piece", quantity=1000, promise=parse_promise("05/24/2026"))],
        ),
    ]
    row = _run(pos, docs)[0]
    assert row.qty_confirmed == Decimal("2500")
    assert "QTY_SHORT" not in row.issue_codes
    assert row.projected_full_qty_promise_date == date(2026, 5, 24)
