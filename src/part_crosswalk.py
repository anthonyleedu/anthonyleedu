"""Historical vendor part-number crosswalk from ERP confirmation rows."""

from __future__ import annotations

from collections import defaultdict
from datetime import date

from src.config import HISTORICAL_MAPPING_MIN_PURITY, HISTORICAL_MAPPING_MIN_SUPPORT
from src.models import CrosswalkEntry, CrosswalkStatus, ExtractedDocument
from src.normalize import normalize_part_number, parse_date


def build_historical_crosswalk(
    confirmation_rows: list[dict],
    po_line_rows: list[dict],
    po_header_rows: list[dict],
    vendor_rows: list[dict],
) -> list[CrosswalkEntry]:
    """Join confirmation to po_line on po_number+line_no, then to vendor via po_header."""
    po_lines = {(r["po_number"], int(r["line_no"])): r for r in po_line_rows}
    headers = {r["po_number"]: r for r in po_header_rows}
    vendors = {r["vendor_id"]: r for r in vendor_rows}

    # vendor_id + vendor_pn -> list of (beacon_pn, date)
    pairs: dict[tuple[str, str], list[tuple[str, date | None]]] = defaultdict(list)

    for c in confirmation_rows:
        vpn = normalize_part_number(c.get("vendor_pn"))
        if not vpn:
            continue
        key = (c.get("po_number"), int(c["line_no"]) if c.get("line_no") is not None else None)
        if key[1] is None or key not in po_lines:
            continue
        pl = po_lines[key]
        beacon = normalize_part_number(pl.get("part_id"))
        if not beacon:
            continue
        hdr = headers.get(c.get("po_number"))
        if not hdr:
            continue
        vid = hdr.get("vendor_id")
        if not vid:
            continue
        seen = parse_date(c.get("doc_date")) or parse_date(c.get("promised_date"))
        pairs[(vid, vpn)].append((beacon, seen))

    entries: list[CrosswalkEntry] = []
    for (vid, vpn), items in pairs.items():
        total = len(items)
        counts: dict[str, int] = defaultdict(int)
        first: dict[str, date | None] = {}
        last: dict[str, date | None] = {}
        for beacon, seen in items:
            counts[beacon] += 1
            if seen:
                if first.get(beacon) is None or seen < first[beacon]:
                    first[beacon] = seen
                if last.get(beacon) is None or seen > last[beacon]:
                    last[beacon] = seen
        vname = (vendors.get(vid) or {}).get("vendor_name")
        for beacon, n in counts.items():
            purity = n / total if total else 0.0
            status = (
                CrosswalkStatus.APPROVED
                if n >= HISTORICAL_MAPPING_MIN_SUPPORT and purity >= HISTORICAL_MAPPING_MIN_PURITY
                else CrosswalkStatus.PROPOSED
            )
            entries.append(
                CrosswalkEntry(
                    vendor_id=vid,
                    vendor_name=vname,
                    vendor_pn=vpn,
                    beacon_pn=beacon,
                    evidence_count=n,
                    vendor_pn_total_count=total,
                    purity=round(purity, 4),
                    first_seen=first.get(beacon),
                    last_seen=last.get(beacon),
                    status=status,
                    source="historical_confirmation",
                )
            )
    entries.sort(key=lambda e: (e.vendor_name or "", e.vendor_pn, -e.evidence_count))
    return entries


def approved_map(entries: list[CrosswalkEntry]) -> dict[tuple[str, str], str]:
    """(vendor_id, vendor_pn) -> beacon_pn for APPROVED mappings only.

    If multiple approved targets exist (shouldn't, given purity), skip that vendor_pn.
    """
    grouped: dict[tuple[str, str], list[CrosswalkEntry]] = defaultdict(list)
    for e in entries:
        if e.status == CrosswalkStatus.APPROVED:
            grouped[(e.vendor_id, e.vendor_pn)].append(e)
    out = {}
    for k, items in grouped.items():
        items = sorted(items, key=lambda e: -e.purity)
        if len(items) == 1 or items[0].purity >= 0.99:
            out[k] = items[0].beacon_pn
    return out


def propose_from_matches(
    documents: list[ExtractedDocument],
    matched_pairs: list[tuple[str, str, str, str]],
    existing: list[CrosswalkEntry],
) -> list[CrosswalkEntry]:
    """Create PROPOSED mappings from confident current-week matches.

    matched_pairs: (vendor_id, vendor_name, vendor_pn, beacon_pn)
    """
    existing_keys = {(e.vendor_id, e.vendor_pn, e.beacon_pn) for e in existing}
    extra: list[CrosswalkEntry] = []
    seen = set()
    for vendor_id, vendor_name, vendor_pn, beacon_pn in matched_pairs:
        vpn = normalize_part_number(vendor_pn)
        bpn = normalize_part_number(beacon_pn)
        if not vpn or not bpn or vpn == bpn:
            continue
        key = (vendor_id, vpn, bpn)
        if key in existing_keys or key in seen:
            continue
        seen.add(key)
        extra.append(
            CrosswalkEntry(
                vendor_id=vendor_id,
                vendor_name=vendor_name,
                vendor_pn=vpn,
                beacon_pn=bpn,
                evidence_count=1,
                vendor_pn_total_count=1,
                purity=1.0,
                status=CrosswalkStatus.PROPOSED,
                source="current_document_context",
            )
        )
    return extra
