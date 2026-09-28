"""Document lifecycle: revision vs split vs invoice."""

from datetime import date

from src.document_versions import resolve_active_documents
from src.models import ConfirmationLine, DocumentLifecycle, DocumentType, ExtractedDocument, PromiseWindow
from src.normalize import parse_promise


def _doc(**kwargs) -> ExtractedDocument:
    defaults = dict(
        source_file="x.pdf",
        source_sha256=kwargs.get("source_sha256", "abc"),
        po_number="PO-1",
        document_type=DocumentType.ACKNOWLEDGEMENT,
        lines=[ConfirmationLine(quantity=1, promise=parse_promise("05/21/2026"))],
    )
    defaults.update(kwargs)
    return ExtractedDocument(**defaults)


def test_quickship_revision_supersedes_original():
    original = _doc(
        source_file="quickship_01.pdf",
        source_sha256="orig",
        po_number="PO-4500050030",
        document_date=date(2026, 5, 6),
        document_type=DocumentType.ACKNOWLEDGEMENT,
        lines=[ConfirmationLine(quantity=450, customer_part_number="MISC-SPR-001")],
    )
    revision = _doc(
        source_file="quickship_02.pdf",
        source_sha256="rev",
        po_number="PO-4500050030",
        document_date=date(2026, 5, 11),
        document_type=DocumentType.REVISION,
        is_revision=True,
        supersedes_all_prior_for_po=True,
        lines=[ConfirmationLine(quantity=500, customer_part_number="MISC-SPR-001")],
    )
    resolved = resolve_active_documents([original, revision])
    assert revision in resolved.active
    assert original in resolved.superseded
    assert original.lifecycle == DocumentLifecycle.SUPERSEDED
    active_qty = [ln.quantity for d in resolved.active for ln in d.lines]
    assert active_qty == [500]


def test_liberty_split_both_active():
    p1 = _doc(
        source_file="liberty_01.pdf",
        source_sha256="p1",
        po_number="PO-4500050016",
        document_date=date(2026, 5, 13),
        partial_sequence=1,
        partial_total=2,
        lines=[ConfirmationLine(quantity=1500, promise=parse_promise("05/17/2026"))],
    )
    p2 = _doc(
        source_file="liberty_02.pdf",
        source_sha256="p2",
        po_number="PO-4500050016",
        document_date=date(2026, 5, 16),
        partial_sequence=2,
        partial_total=2,
        lines=[ConfirmationLine(quantity=1000, promise=parse_promise("05/24/2026"))],
    )
    resolved = resolve_active_documents([p1, p2])
    assert p1 in resolved.active
    assert p2 in resolved.active
    assert resolved.superseded == []
    total = sum(ln.quantity or 0 for d in resolved.active for ln in d.lines)
    assert total == 2500


def test_invoice_is_informational():
    inv = _doc(
        source_file="quickship_03.pdf",
        source_sha256="inv",
        po_number="PO-4500050032",
        document_type=DocumentType.INVOICE,
        lines=[ConfirmationLine(quantity=500)],
    )
    ack = _doc(
        source_file="other.pdf",
        source_sha256="ack",
        po_number="PO-OTHER",
        document_type=DocumentType.ACKNOWLEDGEMENT,
    )
    resolved = resolve_active_documents([inv, ack])
    assert inv in resolved.informational
    assert inv.lifecycle == DocumentLifecycle.INFORMATIONAL
    assert ack in resolved.active
