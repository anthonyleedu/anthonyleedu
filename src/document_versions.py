"""Document lifecycle: revisions vs split/partial acks vs invoices.

A. REVISION — later document REPLACES earlier confirmation for that PO.
B. PARTIAL / SPLIT — multiple documents are simultaneously valid and aggregated.
C. INVOICE / shipping notice — evidence only; not a formal acknowledgment.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

from src.models import (
    DocumentLifecycle,
    DocumentType,
    ExtractedDocument,
)


FORMAL_ACK_TYPES = {DocumentType.ACKNOWLEDGEMENT, DocumentType.REVISION, DocumentType.RECEIPT_ONLY}
INFORMATIONAL_TYPES = {DocumentType.INVOICE, DocumentType.SHIPPING_NOTICE, DocumentType.OTHER}


@dataclass
class ResolvedDocuments:
    active: list[ExtractedDocument] = field(default_factory=list)
    superseded: list[ExtractedDocument] = field(default_factory=list)
    informational: list[ExtractedDocument] = field(default_factory=list)
    review: list[ExtractedDocument] = field(default_factory=list)
    all_documents: list[ExtractedDocument] = field(default_factory=list)

    def active_for_po(self, po_number: str | None) -> list[ExtractedDocument]:
        if not po_number:
            return []
        return [d for d in self.active if d.po_number == po_number]

    def all_for_po(self, po_number: str | None) -> list[ExtractedDocument]:
        if not po_number:
            return []
        return [d for d in self.all_documents if d.po_number == po_number]


def _sort_key(doc: ExtractedDocument) -> tuple:
    d = doc.document_date or date.min
    return (d, doc.source_file)


def resolve_active_documents(documents: list[ExtractedDocument]) -> ResolvedDocuments:
    """Assign lifecycle status. Does not look at open PO lines."""
    resolved = ResolvedDocuments(all_documents=list(documents))

    by_po: dict[str, list[ExtractedDocument]] = {}
    no_po: list[ExtractedDocument] = []
    for doc in documents:
        if doc.extraction_mode.value == "FAILED" or doc.extraction_warnings and not doc.po_number and not doc.lines:
            doc.lifecycle = DocumentLifecycle.REVIEW_REQUIRED
            resolved.review.append(doc)
            continue
        if not doc.po_number:
            no_po.append(doc)
            continue
        by_po.setdefault(doc.po_number, []).append(doc)

    for doc in no_po:
        if doc.document_type in INFORMATIONAL_TYPES:
            doc.lifecycle = DocumentLifecycle.INFORMATIONAL
            resolved.informational.append(doc)
        else:
            doc.lifecycle = DocumentLifecycle.REVIEW_REQUIRED
            resolved.review.append(doc)

    for _po, docs in by_po.items():
        docs_sorted = sorted(docs, key=_sort_key)
        # Identify explicit revisions that supersede all prior for the PO.
        superseding = [
            d
            for d in docs_sorted
            if d.supersedes_all_prior_for_po or (d.is_revision and d.document_type == DocumentType.REVISION)
        ]

        superseded_ids: set[str] = set()
        if superseding:
            latest = superseding[-1]
            for d in docs_sorted:
                if d is latest:
                    continue
                # Only supersede earlier formal acks, not later invoices.
                if d.document_type in INFORMATIONAL_TYPES:
                    continue
                if (d.document_date or date.min) <= (latest.document_date or date.max) and d.source_sha256 != latest.source_sha256:
                    superseded_ids.add(d.source_sha256)

        for d in docs_sorted:
            if d.document_type in INFORMATIONAL_TYPES:
                d.lifecycle = DocumentLifecycle.INFORMATIONAL
                resolved.informational.append(d)
                continue
            if d.source_sha256 in superseded_ids:
                d.lifecycle = DocumentLifecycle.SUPERSEDED
                resolved.superseded.append(d)
                continue
            if d.document_type in FORMAL_ACK_TYPES or d.lines:
                d.lifecycle = DocumentLifecycle.ACTIVE
                resolved.active.append(d)
            else:
                d.lifecycle = DocumentLifecycle.REVIEW_REQUIRED
                resolved.review.append(d)

    return resolved
