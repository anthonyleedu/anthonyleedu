"""Historical vendor-PN crosswalk tests."""

from src.models import CrosswalkStatus
from src.part_crosswalk import build_historical_crosswalk


def test_evidence_count_purity_and_approval():
    confirmations = []
    po_lines = []
    headers = []
    for i in range(5):
        po = f"PO-{i}"
        headers.append({"po_number": po, "vendor_id": "V002"})
        po_lines.append({"po_number": po, "line_no": 1, "part_id": "CHB-9472-3"})
        confirmations.append(
            {
                "po_number": po,
                "line_no": 1,
                "vendor_pn": "APH-441",
                "doc_date": "2026-01-01",
            }
        )
    # One weaker mapping for a different vendor PN
    headers.append({"po_number": "PO-X", "vendor_id": "V002"})
    po_lines.append({"po_number": "PO-X", "line_no": 1, "part_id": "CHB-9472-4"})
    confirmations.append({"po_number": "PO-X", "line_no": 1, "vendor_pn": "APH-441-OS", "doc_date": "2026-02-01"})

    vendors = [{"vendor_id": "V002", "vendor_name": "Heritage Cold Heading"}]
    entries = build_historical_crosswalk(confirmations, po_lines, headers, vendors)
    approved = [e for e in entries if e.vendor_pn == "APH-441"][0]
    assert approved.evidence_count == 5
    assert approved.purity == 1.0
    assert approved.status == CrosswalkStatus.APPROVED

    proposed = [e for e in entries if e.vendor_pn == "APH-441-OS"][0]
    assert proposed.evidence_count == 1
    assert proposed.status == CrosswalkStatus.PROPOSED


def test_low_purity_not_auto_approved():
    headers = [{"po_number": f"PO-{i}", "vendor_id": "V005"} for i in range(4)]
    po_lines = [
        {"po_number": "PO-0", "line_no": 1, "part_id": "TOOL-INS-CAR"},
        {"po_number": "PO-1", "line_no": 1, "part_id": "TOOL-INS-CAR"},
        {"po_number": "PO-2", "line_no": 1, "part_id": "TOOL-DIE-9472"},
        {"po_number": "PO-3", "line_no": 1, "part_id": "TOOL-INS-CAR"},
    ]
    confs = [{"po_number": f"PO-{i}", "line_no": 1, "vendor_pn": "OST-X", "doc_date": "2026-01-01"} for i in range(4)]
    vendors = [{"vendor_id": "V005", "vendor_name": "Ostmark"}]
    entries = build_historical_crosswalk(confs, po_lines, headers, vendors)
    car = [e for e in entries if e.beacon_pn == "TOOL-INS-CAR"][0]
    assert car.evidence_count == 3
    assert car.purity == 0.75
    assert car.status == CrosswalkStatus.PROPOSED
