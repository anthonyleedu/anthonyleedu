"""Deterministic fallback parser for native PDF / OCR text.

Used when no LLM API key is configured, or as a repair path.
Parses facts from document text - never from filenames or the open-PO list.
"""

from __future__ import annotations

import re
from src.models import (
    CommitmentStatus,
    ConfirmationLine,
    DateType,
    DocumentType,
)
from src.normalize import (
    detect_currency,
    extract_po_number,
    normalize_part_number,
    parse_number,
    parse_promise,
)

FOOTER_START = re.compile(
    r"^(terms|note:|notes:|zahlungsbedingungen|all work|customer-supplied|"
    r"regards|thank you|fob |material cert|preise in|lemme know|"
    r"please contact|hi lisa|lisa -)",
    re.IGNORECASE,
)

PART_TOKEN = re.compile(r"^[A-Z]{2,}[A-Z0-9]*(?:-[A-Z0-9]+)+$", re.IGNORECASE)
LINE_NO = re.compile(r"^\d{1,3}$")
QTY_TOKEN = re.compile(r"^\d{1,3}(?:,\d{3})+(?:\.\d+)?$|^\d+(?:\.\d+)?$")
PRICE_TOKEN = re.compile("^[\\$\u20ac]\\s*\\d")
UOM_TOKEN = re.compile(r"^(ea|lbs?|lb|kg|pcs?|pc|each|stuck|stk|ib|ibs)$", re.IGNORECASE)

PARTIAL_RE = re.compile(r"part\s+(\d+)\s+of\s+(\d+)", re.IGNORECASE)
DOCNO_PATTERNS = [
    re.compile(r"(?:Apex Order\s*#|Order\s*#)\s*[:\s]*([A-Z0-9\-]+)", re.IGNORECASE),
    re.compile(r"SO\s*#\s*([A-Z0-9\-]+)", re.IGNORECASE),
    re.compile(r"Process Order\s*#\s*[:\s]*([A-Z0-9\-]+)", re.IGNORECASE),
    re.compile(r"Auftrag-Nr\.\s*[:\s]*([A-Z0-9\-]+)", re.IGNORECASE),
    re.compile(r"Continental WO:\s*([A-Z0-9\-]+)", re.IGNORECASE),
    re.compile(r"WO:\s*([A-Z0-9\-]+)", re.IGNORECASE),
]

QUICKSHIP_LINE = re.compile(
    r"(?P<pn>[A-Z0-9][A-Z0-9\-]+)\s+qty\s+(?P<qty>[\d,]+)\s+@\s+"
    r"(?P<price>[$]?[\d,\.]+)\s*(?:ea|each)?\s+ship\s+(?P<ship>[\d/\.\-]+)",
    re.IGNORECASE,
)

LABELED = {
    "description": re.compile(r"^Description:\s*(.+)$", re.IGNORECASE),
    "quantity": re.compile(r"^Quantity:\s*([\d,]+)\s*([A-Za-z]*)", re.IGNORECASE),
    "promise": re.compile(r"^(?:Promise|Liefertermin|Delivery):\s*(.+)$", re.IGNORECASE),
    "customer_po": re.compile(r"(?:Customer PO|Your PO|Your reference|PO Reference|Ihre Bestellung)[:\s]+(.+)$", re.IGNORECASE),
    "date": re.compile(
        r"^(?:Date|Acknowledgment Date|Issued|Datum\s*/\s*Date|Datum):\s*(.+)$",
        re.IGNORECASE,
    ),
    "vendor_doc": re.compile(
        r"^(?:Apex Order\s*#|Process Order\s*#|Auftrag-Nr\.|Continental WO|SO #)\s*[:\s]*(.+)$",
        re.IGNORECASE,
    ),
}


HEADER_ALIASES = {
    "line": "line",
    "item": "line",
    "pos.": "line",
    "pos": "line",
    "customer pn": "customer_pn",
    "customer p/n": "customer_pn",
    "our p/n": "vendor_pn",
    "art.-nr. (ostmark)": "vendor_pn",
    "art.-nr.": "vendor_pn",
    "vendor pn": "vendor_pn",
    "description": "description",
    "bezeichnung": "description",
    "qty": "qty",
    "quantity": "qty",
    "menge": "qty",
    "unit price": "price",
    "unit": "price",
    "eur / stuck": "price",
    "eur/stuck": "price",
    "eur / stueck": "price",
    "eur/stueck": "price",
    "price": "price",
    "promise date": "promise",
    "promise": "promise",
    "liefertermin": "promise",
    "process": "process",
}


def _lines(text: str) -> list[str]:
    return [ln.strip() for ln in text.replace("\r", "").split("\n") if ln.strip()]


def classify_document(text: str) -> tuple[DocumentType, bool, bool, CommitmentStatus]:
    low = text.lower()
    is_revision = bool(
        re.search(r"\brevised\b|\brevision\b|disregard our earlier|supersede|please disregard", low)
    )
    supersedes = bool(re.search(r"disregard our earlier|supersede", low))
    looks_like_ack = bool(
        re.search(
            r"order acknowledgment|order acknowledgement|order confirmation|auftragsbestätigung|"
            r"sales order confirmation|process order acknowledgment",
            low,
        )
    )
    # "Net 30 from invoice date" is not an invoice document.
    is_invoice = (not looks_like_ack) and bool(
        re.search(r"subject:\s*.*\binvoice\b|\bplease find invoice\b|^invoice\b|\binvoice\b\s*[-:]", low, re.M)
        or (re.search(r"\binvoice\b", low) and not re.search(r"invoice date", low))
    )
    receipt_only = bool(
        re.search(
            r"schedule to be confirmed|to be confirmed upon|acknowledge receipt of your purchase order",
            low,
        )
    ) and not re.search(r"promise date|liefertermin|ship \d", low)

    if is_invoice:
        return DocumentType.INVOICE, False, False, CommitmentStatus.UNKNOWN
    if is_revision:
        return DocumentType.REVISION, True, supersedes, CommitmentStatus.FULL
    if receipt_only:
        return (
            DocumentType.RECEIPT_ONLY,
            False,
            False,
            CommitmentStatus.RECEIPT_ACKNOWLEDGED_NO_SCHEDULE,
        )
    if re.search(
        r"acknowledgment|acknowledgement|order confirmation|auftragsbest.*tigung|"
        r"sales order confirmation|process order acknowledgment",
        low,
    ):
        partial = PARTIAL_RE.search(text)
        status = CommitmentStatus.PARTIAL if partial else CommitmentStatus.FULL
        return DocumentType.ACKNOWLEDGEMENT, False, False, status
    return DocumentType.UNKNOWN, is_revision, supersedes, CommitmentStatus.UNKNOWN


def infer_vendor_name(text: str) -> str | None:
    patterns = [
        r"APEX BAR & TUBE CO\.?",
        r"Heritage Cold Heading",
        r"Liberty Surface Finishing",
        r"CONTINENTAL QUALITY HEAT TREAT",
        r"Ostmark Werkzeug GmbH",
        r"QuickShip Industrial",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            raw = m.group(0)
            return raw.title() if raw.isupper() else raw
    if re.search(r"quickship", text, re.IGNORECASE):
        return "QuickShip Industrial"
    for ln in _lines(text)[:5]:
        if len(ln) > 4 and not re.search(r"beacon fasteners|from:|to:|subject:", ln, re.IGNORECASE):
            return ln
    return None


def _find_partial(text: str) -> tuple[int | None, int | None]:
    m = PARTIAL_RE.search(text)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None


def _document_number(text: str) -> str | None:
    for rx in DOCNO_PATTERNS:
        m = rx.search(text)
        if m:
            return m.group(1).strip()
    return None


def _document_date_raw(text: str) -> str | None:
    for ln in _lines(text)[:25]:
        m = LABELED["date"].search(ln)
        if m:
            return m.group(1).strip()
        m = re.search(r"^Date:\s*(.+)$", ln, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    # QuickShip email Date header
    m = re.search(r"^Date:\s*([0-9./\-]+)\s*$", text, re.IGNORECASE | re.MULTILINE)
    if m:
        return m.group(1)
    return None


def _map_header(token: str) -> str | None:
    key = token.strip().lower()
    key = re.sub(r"\s+", " ", key)
    key = (
        key.replace("\u00fc", "u")
        .replace("\u00f6", "o")
        .replace("\u00e4", "a")
        .replace("\u00df", "ss")
    )
    return HEADER_ALIASES.get(key)


def _looks_header_block(window: list[str]) -> list[str] | None:
    mapped = []
    for w in window:
        col = _map_header(w)
        if col is None:
            return None
        mapped.append(col)
    if "qty" in mapped and ("promise" in mapped or "price" in mapped or "description" in mapped):
        return mapped
    return None


def parse_table_lines(lines: list[str]) -> list[ConfirmationLine]:
    """Detect a header block of consecutive mapped labels, then row-major cells."""
    best: list[ConfirmationLine] = []
    for i in range(len(lines)):
        for width in range(7, 3, -1):
            window = lines[i : i + width]
            if len(window) < width:
                continue
            cols = _looks_header_block(window)
            if not cols:
                continue
            body: list[str] = []
            for ln in lines[i + width :]:
                if FOOTER_START.search(ln):
                    break
                if re.match(r"^(Terms:|Note:|Zahlungsbedingungen|All work|Regards,|Thank you)", ln, re.I):
                    break
                body.append(ln)
            rows = _cells_to_rows(body, cols)
            if rows:
                return rows
            if not best and rows:
                best = rows
    return best


def _cells_to_rows(body: list[str], cols: list[str]) -> list[ConfirmationLine]:
    width = len(cols)
    # Stop when remaining cells aren't a multiple? Still try.
    rows: list[ConfirmationLine] = []
    # If the first body cell isn't a line number or part/description, abort.
    i = 0
    while i + width - 1 < len(body):
        chunk = body[i : i + width]
        rec = dict(zip(cols, chunk))
        line = _row_from_dict(rec)
        if line is None:
            # maybe footer started
            break
        rows.append(line)
        i += width
        # If next leftover looks like footer, stop
        if i < len(body) and FOOTER_START.search(body[i]):
            break
    return rows


def _row_from_dict(rec: dict[str, str]) -> ConfirmationLine | None:
    qty = parse_number(rec.get("qty"))
    price = parse_number(rec.get("price"))
    desc = rec.get("description")
    cust = normalize_part_number(rec.get("customer_pn"))
    vend = normalize_part_number(rec.get("vendor_pn"))
    promise_raw = rec.get("promise")
    line_no = None
    raw_line = rec.get("line")
    if raw_line and LINE_NO.match(raw_line.strip()):
        line_no = int(raw_line.strip())
    # Heritage "Unit" column is actually unit price; already mapped to price.
    process = rec.get("process")
    if process and not vend:
        # keep process in description, not as a fake PN
        if desc:
            desc = f"{desc} [{process}]"
        else:
            desc = process

    if qty is None and not cust and not vend and not promise_raw:
        return None
    currency = None
    raw_price = rec.get("price") or ""
    if "\u20ac" in raw_price or "EUR" in (raw_price.upper() if raw_price else ""):
        currency = "EUR"
    elif "$" in raw_price:
        currency = "USD"

    header_hint = "promise"
    if promise_raw and re.search(r"\bship\b", promise_raw, re.I):
        header_hint = "ship"
    prefer_dmy = bool(promise_raw and "." in (promise_raw or "") and "/" not in (promise_raw or ""))
    promise = parse_promise(promise_raw, header_hint=header_hint, prefer_dmy=prefer_dmy)

    uom = None
    return ConfirmationLine(
        source_line_number=line_no,
        customer_part_number=cust,
        vendor_part_number=vend,
        description=desc,
        quantity=float(qty) if qty is not None else None,
        uom=uom,
        unit_price=price,
        currency=currency,
        promise=promise,
    )


def parse_labeled_block(text: str) -> list[ConfirmationLine]:
    """Continental-style PROCESS DETAILS labeled fields."""
    desc = qty = uom = promise_raw = None
    for ln in _lines(text):
        m = LABELED["description"].match(ln)
        if m:
            desc = m.group(1).strip()
            continue
        m = LABELED["quantity"].match(ln)
        if m:
            qty = parse_number(m.group(1))
            uom_raw = (m.group(2) or "").strip()
            if uom_raw.lower() in {"ib", "ibs"}:
                uom_raw = "lb" if uom_raw.lower() == "ib" else "lbs"
            uom = uom_raw or None
            continue
        m = LABELED["promise"].match(ln)
        if m:
            promise_raw = m.group(1).strip()
    if qty is None and not promise_raw and not desc:
        return []
    promise = parse_promise(promise_raw, header_hint="promise")
    return [
        ConfirmationLine(
            source_line_number=1,
            description=desc,
            quantity=float(qty) if qty is not None else None,
            uom=uom,
            promise=promise,
        )
    ]


def parse_quickship(text: str) -> list[ConfirmationLine]:
    lines = []
    for m in QUICKSHIP_LINE.finditer(text):
        qty = parse_number(m.group("qty"))
        price = parse_number(m.group("price"))
        pn = normalize_part_number(m.group("pn"))
        promise = parse_promise(f"ship {m.group('ship')}", header_hint="ship")
        cur = "EUR" if ("\u20ac" in m.group("price") or "EUR" in m.group("price").upper()) else "USD"
        lines.append(
            ConfirmationLine(
                customer_part_number=pn,
                vendor_part_number=None,
                quantity=float(qty) if qty is not None else None,
                unit_price=price,
                currency=cur,
                uom="EA",
                promise=promise,
            )
        )
    return lines


def parse_document_text(text: str) -> dict:
    doc_type, is_revision, supersedes, commitment = classify_document(text)
    vendor = infer_vendor_name(text)
    po = extract_po_number(text)
    seq, tot = _find_partial(text)
    currency = detect_currency(text)
    notes = []
    if seq and tot:
        notes.append(f"Split acknowledgment part {seq} of {tot}")
        commitment = CommitmentStatus.PARTIAL

    lines = parse_quickship(text)
    if not lines:
        lines = parse_table_lines(_lines(text))
    if not lines:
        lines = parse_labeled_block(text)

    # Do not invent lines for receipt-only documents.
    if commitment == CommitmentStatus.RECEIPT_ACKNOWLEDGED_NO_SCHEDULE:
        lines = []

    for ln in lines:
        if ln.currency is None:
            ln.currency = currency
        if ln.promise.date_type == DateType.UNKNOWN and ln.promise.is_usable:
            ln.promise.date_type = DateType.PROMISE_DATE

    if doc_type == DocumentType.INVOICE:
        # Keep shipment evidence but do not treat as ack.
        pass

    payload = {
        "vendor_name_raw": vendor,
        "po_number": po,
        "document_number": _document_number(text),
        "document_date": _document_date_raw(text),
        "document_type": doc_type.value,
        "commitment_status": commitment.value,
        "is_revision": is_revision,
        "supersedes_all_prior_for_po": supersedes,
        "partial_sequence": seq,
        "partial_total": tot,
        "document_currency": currency,
        "free_text_notes": notes,
        "extraction_warnings": [],
        "lines": [
            {
                "source_line_number": ln.source_line_number,
                "customer_part_number": ln.customer_part_number,
                "vendor_part_number": ln.vendor_part_number,
                "description": ln.description,
                "quantity": ln.quantity,
                "uom": ln.uom,
                "unit_price": str(ln.unit_price) if ln.unit_price is not None else None,
                "currency": ln.currency,
                "promise_raw_text": ln.promise.raw_text,
                "date_type": ln.promise.date_type.value,
                "page_number": ln.page_number,
            }
            for ln in lines
        ],
    }
    return payload
