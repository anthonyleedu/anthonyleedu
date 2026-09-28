"""Match confirmation lines to Beacon PO lines.

Matching is confined to the PO stated on the document.
The extractor never sees the open PO list; this module does.
"""

from __future__ import annotations

from decimal import Decimal

from rapidfuzz import fuzz

from src.config import (
    BASE_CURRENCY,
    PLACEHOLDER_DESCRIPTIONS,
    RULE_MATCH_AUTO_THRESHOLD,
    RULE_MATCH_MARGIN,
)
from src.currency import FxTable
from src.models import (
    ConfirmationLine,
    CrosswalkEntry,
    ExtractedDocument,
    MatchMethod,
    MatchResult,
    OpenPOLine,
)
from src.normalize import normalize_part_number, vendor_key
from src.part_crosswalk import approved_map


def _placeholder_desc(text: str | None) -> bool:
    if not text:
        return True
    return text.strip().lower() in PLACEHOLDER_DESCRIPTIONS or "see po" in text.lower() or "siehe bestellung" in text.lower()


def _price_close(a: Decimal | None, b: Decimal | None, *, foreign: bool) -> bool:
    if a is None or b is None:
        return False
    if b == 0:
        return a == 0
    pct = abs(a - b) / abs(b)
    tol = Decimal("0.005") if foreign else Decimal("0.001")
    return pct <= tol or abs(a - b) <= Decimal("0.01")


def _qty_close(a: float | Decimal | None, b: Decimal | None) -> bool:
    if a is None or b is None:
        return False
    return abs(Decimal(str(a)) - b) <= Decimal("0.0001")


def score_rule(
    conf: ConfirmationLine,
    po: OpenPOLine,
    *,
    confirmed_usd: Decimal | None,
    foreign: bool,
) -> tuple[float, list[str]]:
    score = 0.0
    why: list[str] = []
    if confirmed_usd is not None and _price_close(confirmed_usd, po.unit_price, foreign=foreign):
        score += 0.40
        why.append("price matches")
    if _qty_close(conf.quantity, po.qty_ordered):
        score += 0.30
        why.append("quantity matches")
    if not _placeholder_desc(conf.description) and po.our_description:
        sim = fuzz.token_set_ratio(conf.description, po.our_description) / 100.0
        if sim >= 0.72:
            score += 0.20
            why.append(f"description similarity {sim:.0%}")
    if conf.source_line_number is not None and conf.source_line_number == po.line_number:
        score += 0.10
        why.append("line number matches")
    return score, why


def match_document_to_po(
    doc: ExtractedDocument,
    po_lines: list[OpenPOLine],
    *,
    crosswalk: list[CrosswalkEntry],
    fx: FxTable,
    expected_vendor_id: str | None,
    expected_vendor_name: str | None,
) -> tuple[list[MatchResult], list[str]]:
    """Return matches (and leftover unmatched confirmation lines as UNMATCHED results)."""
    issues: list[str] = []
    if not po_lines:
        issues.append("UNKNOWN_PO")
        return [], issues

    po_vendor_id = po_lines[0].vendor_id
    po_vendor_name = po_lines[0].vendor_name
    if doc.vendor_name_raw and po_vendor_name:
        if vendor_key(doc.vendor_name_raw) not in vendor_key(po_vendor_name) and vendor_key(po_vendor_name) not in vendor_key(doc.vendor_name_raw):
            # soft check — names can be abbreviated
            ratio = fuzz.token_set_ratio(doc.vendor_name_raw, po_vendor_name)
            if ratio < 70:
                issues.append("VENDOR_MISMATCH")

    approved = approved_map(crosswalk)
    remaining_po = list(po_lines)
    remaining_conf = list(doc.lines)
    results: list[MatchResult] = []

    def take_po(po: OpenPOLine) -> None:
        remaining_po[:] = [p for p in remaining_po if not (p.po_number == po.po_number and p.line_number == po.line_number)]

    def take_conf(conf: ConfirmationLine) -> None:
        remaining_conf[:] = [c for c in remaining_conf if c is not conf]

    # 1. Exact customer PN
    for conf in list(remaining_conf):
        cpn = normalize_part_number(conf.customer_part_number)
        if not cpn:
            continue
        hits = [p for p in remaining_po if p.our_pn == cpn]
        if len(hits) == 1:
            po = hits[0]
            results.append(
                MatchResult(
                    po_line=po,
                    confirmation_line=conf,
                    source_document=doc,
                    method=MatchMethod.EXACT_CUSTOMER_PN,
                    confidence=0.99,
                    explanation=f"Customer PN {cpn} matches Beacon {po.our_pn}",
                )
            )
            take_po(po)
            take_conf(conf)

    # 2. Approved historical vendor crosswalk
    vid = po_vendor_id or expected_vendor_id
    for conf in list(remaining_conf):
        vpn = normalize_part_number(conf.vendor_part_number)
        if not vpn or not vid:
            continue
        beacon = approved.get((vid, vpn))
        if not beacon:
            continue
        hits = [p for p in remaining_po if p.our_pn == beacon]
        if len(hits) == 1:
            po = hits[0]
            results.append(
                MatchResult(
                    po_line=po,
                    confirmation_line=conf,
                    source_document=doc,
                    method=MatchMethod.APPROVED_VENDOR_CROSSWALK,
                    confidence=0.92,
                    explanation=f"Approved crosswalk {vpn} ? {beacon}",
                )
            )
            take_po(po)
            take_conf(conf)

    # 3. vendor PN exactly equals a Beacon our_pn on that PO
    for conf in list(remaining_conf):
        vpn = normalize_part_number(conf.vendor_part_number)
        if not vpn:
            continue
        hits = [p for p in remaining_po if p.our_pn == vpn]
        if len(hits) == 1:
            po = hits[0]
            results.append(
                MatchResult(
                    po_line=po,
                    confirmation_line=conf,
                    source_document=doc,
                    method=MatchMethod.EXACT_VENDOR_PN_TO_BEACON_PN,
                    confidence=0.90,
                    explanation=f"Vendor PN {vpn} equals Beacon PN on this PO",
                )
            )
            take_po(po)
            take_conf(conf)

    # 4. Rule-based on remaining
    while remaining_conf and remaining_po:
        scored: list[tuple[float, ConfirmationLine, OpenPOLine, list[str]]] = []
        for conf in remaining_conf:
            foreign = bool(conf.currency and conf.currency.upper() != BASE_CURRENCY)
            usd, _, _, _ = fx.to_usd(
                conf.unit_price,
                conf.currency or doc.document_currency,
                doc.document_date,
            )
            for po in remaining_po:
                s, why = score_rule(conf, po, confirmed_usd=usd, foreign=foreign)
                scored.append((s, conf, po, why))
        if not scored:
            break
        scored.sort(key=lambda t: t[0], reverse=True)
        best = scored[0]
        # second best for same confirmation line
        same_conf = [t for t in scored if t[1] is best[1]]
        same_conf.sort(key=lambda t: t[0], reverse=True)
        second = same_conf[1][0] if len(same_conf) > 1 else 0.0
        best_score, conf, po, why = best

        if best_score >= RULE_MATCH_AUTO_THRESHOLD and (best_score - second) >= RULE_MATCH_MARGIN:
            results.append(
                MatchResult(
                    po_line=po,
                    confirmation_line=conf,
                    source_document=doc,
                    method=MatchMethod.RULE_BASED,
                    confidence=round(best_score, 3),
                    explanation="Rule-based: " + ", ".join(why) if why else f"score {best_score:.2f}",
                )
            )
            take_po(po)
            take_conf(conf)
            continue

        # 5. Single remaining line — supporting evidence only
        if len(remaining_conf) == 1 and len(remaining_po) == 1:
            conf = remaining_conf[0]
            po = remaining_po[0]
            results.append(
                MatchResult(
                    po_line=po,
                    confirmation_line=conf,
                    source_document=doc,
                    method=MatchMethod.SINGLE_REMAINING_LINE,
                    confidence=max(0.55, best_score),
                    explanation="Only one unmatched confirmation line and one unmatched PO line remain",
                )
            )
            take_po(po)
            take_conf(conf)
            break

        # Ambiguous: attach MANUAL_REVIEW to best pair without consuming others blindly
        results.append(
            MatchResult(
                po_line=None,
                confirmation_line=conf,
                source_document=doc,
                method=MatchMethod.MANUAL_REVIEW,
                confidence=round(best_score, 3),
                explanation=(
                    f"Ambiguous match (best {best_score:.2f}, margin {best_score - second:.2f}). "
                    + (", ".join(why) if why else "Insufficient unique evidence")
                ),
            )
        )
        take_conf(conf)
        issues.append("PART_MATCH_UNCERTAIN")
        break

    for conf in remaining_conf:
        results.append(
            MatchResult(
                po_line=None,
                confirmation_line=conf,
                source_document=doc,
                method=MatchMethod.UNMATCHED,
                confidence=0.0,
                explanation="No unique PO line match on this purchase order",
            )
        )
        issues.append("EXTRA_CONFIRMATION_LINE")

    return results, issues
