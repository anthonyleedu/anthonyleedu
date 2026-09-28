from decimal import Decimal

from src.document_extractor import refine_extracted_document
from src.llm_client import AuthenticationFailed, is_auth_error, sanitize_error
from src.models import (
    CommitmentStatus,
    ConfirmationLine,
    DocumentType,
    ExtractedDocument,
    ExtractionMode,
)
from src.normalize import parse_promise


def test_sanitize_error_strips_openai_key_material():
    # Concatenated so the file does not contain a scanner-looking sk-proj token.
    fake_token = "sk-" + "proj-EXAMPLEONLY-NOTAREALSECRET-TESTFIXTURE"
    raw = f"Error code: 401 - Incorrect API key provided: {fake_token}"
    cleaned = sanitize_error(Exception(raw))
    assert fake_token not in cleaned
    assert "NOTAREALSECRET" not in cleaned
    assert "sk-***" in cleaned


def test_is_auth_error_detects_invalidated_and_incorrect_keys():
    assert is_auth_error(Exception("Error code: 401 - token_invalidated"))
    assert is_auth_error(Exception("Incorrect API key provided: sk-***azEA"))
    assert is_auth_error(AuthenticationFailed("rejected"))
    assert not is_auth_error(Exception("rate limit exceeded"))


def test_refine_marks_receipt_only_without_inventing_lines():
    text = (
        "Liberty Surface Finishing\n"
        "Process Order Acknowledgment\n"
        "PO Reference: PO-4500050019\n"
        "We acknowledge receipt of your purchase order and incoming parts shipment. "
        "Process schedule to be confirmed upon completion of incoming inspection.\n"
    )
    doc = ExtractedDocument(
        source_file="liberty_03.pdf",
        source_sha256="x",
        po_number="PO-4500050019",
        document_type=DocumentType.ACKNOWLEDGEMENT,
        commitment_status=CommitmentStatus.UNKNOWN,
        extraction_mode=ExtractionMode.NATIVE_TEXT,
        lines=[],
    )
    refine_extracted_document(doc, text)
    assert doc.document_type == DocumentType.RECEIPT_ONLY
    assert doc.commitment_status == CommitmentStatus.RECEIPT_ACKNOWLEDGED_NO_SCHEDULE
    assert doc.lines == []


def test_refine_fills_missing_price_from_visible_text_not_new_lines():
    text = (
        "Ostmark Werkzeug GmbH\n"
        "Auftragsbestätigung / Order Confirmation\n"
        "Ihre Bestellung / Your PO: PO-4500050027\n"
        "Pos.\nArt.-Nr. (Ostmark)\nBezeichnung\nMenge\nEUR / Stück\nLiefertermin\n"
        "1\nTOOL-DIE-9472\n(siehe Bestellung)\n25\n€1140.8000\nKW 20-22 / 2026\n"
        "Preise in EUR.\n"
    )
    doc = ExtractedDocument(
        source_file="ostmark_01.pdf",
        source_sha256="y",
        po_number="PO-4500050027",
        document_type=DocumentType.ACKNOWLEDGEMENT,
        document_currency="EUR",
        extraction_mode=ExtractionMode.NATIVE_TEXT,
        lines=[
            ConfirmationLine(
                source_line_number=1,
                vendor_part_number="TOOL-DIE-9472",
                quantity=25.0,
                unit_price=None,
                currency=None,
                promise=parse_promise("KW 20-22 / 2026"),
            )
        ],
    )
    refine_extracted_document(doc, text)
    assert doc.lines[0].unit_price == Decimal("1140.8000")
    assert doc.lines[0].currency == "EUR"
    assert len(doc.lines) == 1

