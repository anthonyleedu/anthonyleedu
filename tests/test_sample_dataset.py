"""Integration assertions against the supplied Beacon sample dataset.

Skipped automatically if the sample files are not present, so a generic
install without the case-study PDFs still has a green unit-test suite.
"""

from __future__ import annotations

from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

from src.currency import FxTable
from src.document_extractor import extract_all
from src.document_versions import resolve_active_documents
from src.erp import load_erp
from src.file_discovery import discover_from_data_dir
from src.inputs import load_open_pos, load_vendor_master
from src.part_crosswalk import build_historical_crosswalk
from src.reconcile import reconcile
from src.vendor_performance import analyze_vendor_performance

DATA = Path("data")
ERP = DATA / "beacon_erp.db"
OPEN = DATA / "open_pos.csv"


def sample_available() -> bool:
    return ERP.exists() and OPEN.exists() and (DATA / "confirmations").exists()


pytestmark = pytest.mark.skipif(not sample_available(), reason="Beacon sample dataset not present")


@pytest.fixture(scope="module")
def pipeline(tmp_path_factory):
    discovered = discover_from_data_dir(DATA)
    open_pos = load_open_pos(discovered.open_pos)
    vendors = load_vendor_master(discovered.vendor_master)
    conn, schema, erp = load_erp(discovered.erp)
    try:
        fx = FxTable(erp["fx_rate"])
        xw = build_historical_crosswalk(erp["confirmation"], erp["po_line"], erp["po_header"], erp["vendor_master"])
        cache = tmp_path_factory.mktemp("cache")
        docs = extract_all(discovered.confirmations, cache_dir=cache, offline=True)
        resolved = resolve_active_documents(docs)
        rows, _, _ = reconcile(open_pos, vendors, resolved, crosswalk=xw, fx=fx)
        perf = analyze_vendor_performance(erp)
        return {
            "docs": docs,
            "resolved": resolved,
            "rows": rows,
            "perf": perf,
            "xw": xw,
        }
    finally:
        conn.close()


def _row(rows, po, line=None, pn=None):
    hits = [r for r in rows if r.po_number == po]
    if line is not None:
        hits = [r for r in hits if r.po_line_number == line]
    if pn is not None:
        hits = [r for r in hits if r.beacon_pn == pn]
    assert hits, f"No reconciled row for {po} line={line} pn={pn}"
    return hits[0]


def test_pdf_count_is_about_thirty_five(pipeline):
    assert 30 <= len(pipeline["docs"]) <= 50


def test_apex_missing_line(pipeline):
    missing = _row(pipeline["rows"], "PO-4500050001", pn="BAR-A286-250")
    assert "MISSING_PO_LINE" in missing.issue_codes
    present = _row(pipeline["rows"], "PO-4500050001", pn="BAR-A286-375")
    assert "MISSING_PO_LINE" not in present.issue_codes


def test_apex_price_variance(pipeline):
    row = _row(pipeline["rows"], "PO-4500050002")
    assert row.po_unit_price == Decimal("3.92")
    assert row.confirmed_unit_price_raw == Decimal("4.0102")
    assert abs(row.unit_price_difference - Decimal("0.0902")) < Decimal("0.00001")
    assert abs(float(row.unit_price_difference_pct) - 0.0230) < 0.0005
    assert "PRICE_HIGH" in row.issue_codes


def test_apex_qty_short(pipeline):
    row = _row(pipeline["rows"], "PO-4500050007")
    assert row.qty_ordered == Decimal("1500")
    assert row.qty_confirmed == Decimal("1425")
    assert row.qty_variance == Decimal("-75")
    assert "QTY_SHORT" in row.issue_codes


def test_apex_rogue_unknown_po(pipeline):
    row = _row(pipeline["rows"], "PO-4500060619")
    assert row.unknown_po or "UNKNOWN_PO" in row.issue_codes


def test_continental_scan_and_late(pipeline):
    scanned = [d for d in pipeline["docs"] if "continental" in Path(d.source_file).name.lower()]
    assert scanned
    assert all((d.native_text_chars or 0) < 80 or d.extraction_mode.value in {"OCR_LOCAL", "VISION"} for d in scanned)
    row = _row(pipeline["rows"], "PO-4500050022")
    assert row.qty_confirmed == Decimal("500")
    assert row.promise_start == date(2026, 6, 11)
    assert row.required_date == date(2026, 5, 21)
    assert row.days_late == 21
    assert "PROMISE_LATE" in row.issue_codes


def test_continental_other_pos(pipeline):
    r23 = _row(pipeline["rows"], "PO-4500050023")
    assert r23.qty_confirmed == Decimal("750")
    assert r23.promise_start == date(2026, 5, 23)
    r24 = _row(pipeline["rows"], "PO-4500050024")
    assert r24.qty_confirmed == Decimal("350")
    assert r24.promise_start == date(2026, 5, 9)
    r25 = _row(pipeline["rows"], "PO-4500050025")
    assert r25.qty_confirmed == Decimal("350")
    assert r25.promise_start == date(2026, 5, 7)


def test_heritage_crosswalk_and_proposed(pipeline):
    r9 = _row(pipeline["rows"], "PO-4500050009")
    assert r9.vendor_pn == "APH-441"
    assert r9.beacon_pn == "CHB-9472-3"
    r12 = _row(pipeline["rows"], "PO-4500050012")
    assert r12.beacon_pn == "CHB-9472-4"
    assert r12.qty_confirmed == Decimal("15000")
    assert r12.days_late == 14
    r15_os = _row(pipeline["rows"], "PO-4500050015", pn="CHB-9472-4")
    r15_7715 = _row(pipeline["rows"], "PO-4500050015", pn="CHB-7715")
    r15_441 = _row(pipeline["rows"], "PO-4500050015", pn="CHB-9472-3")
    assert r15_os.qty_confirmed == Decimal("25000")
    assert r15_7715.qty_confirmed == Decimal("25000")
    assert r15_441.qty_confirmed == Decimal("5000")


def test_liberty_split_not_shortage(pipeline):
    row = _row(pipeline["rows"], "PO-4500050016")
    assert row.qty_ordered == Decimal("2500")
    assert row.qty_confirmed == Decimal("2500")
    assert "QTY_SHORT" not in row.issue_codes
    assert row.projected_full_qty_promise_date == date(2026, 5, 24)
    active = [d for d in pipeline["resolved"].active if d.po_number == "PO-4500050016"]
    assert len(active) == 2


def test_liberty_receipt_only_no_schedule(pipeline):
    row = _row(pipeline["rows"], "PO-4500050019")
    assert "PROMISE_MISSING" in row.issue_codes or row.document_type.value == "RECEIPT_ONLY"


def test_ostmark_eur_week_window(pipeline):
    rows = [r for r in pipeline["rows"] if r.po_number == "PO-4500050027"]
    assert len(rows) == 2
    for r in rows:
        assert r.confirmed_currency == "EUR"
        assert r.promise_raw and "KW" in r.promise_raw.upper()
        assert r.promise_start != r.promise_end
        assert "PROMISE_WINDOW_STRADDLES_REQUIRED" in r.issue_codes or r.promise_start <= r.required_date <= r.promise_end
    car = [r for r in rows if r.beacon_pn == "TOOL-INS-CAR"][0]
    assert car.vendor_pn == "OST-CAR-A-100"


def test_quickship_revision_qty_500(pipeline):
    row = _row(pipeline["rows"], "PO-4500050030")
    assert row.qty_confirmed == Decimal("500")
    superseded = [d for d in pipeline["resolved"].superseded if d.po_number == "PO-4500050030"]
    assert superseded
    assert any(ln.quantity == 450 for d in superseded for ln in d.lines)


def test_quickship_invoice_not_ack(pipeline):
    docs = [d for d in pipeline["docs"] if d.po_number == "PO-4500050032"]
    assert any(d.document_type.value == "INVOICE" for d in docs)
    for r in [x for x in pipeline["rows"] if x.po_number == "PO-4500050032" and x.po_line_number]:
        assert "NO_FORMAL_ACK_FOUND" in r.issue_codes or "INVOICE_NOT_ACK" in r.issue_codes


def test_historical_metrics(pipeline):
    perf = pipeline["perf"]
    assert perf.as_of == date(2026, 5, 17)
    monthly = {m.month: float(m.received_value) for m in perf.monthly}
    expected = {
        "2025-09": 114453.14,
        "2025-10": 3533227.93,
        "2025-11": 4785787.08,
        "2025-12": 3668417.52,
        "2026-01": 5091552.47,
        "2026-02": 3126206.31,
        "2026-03": 3722784.93,
        "2026-04": 3223456.87,
        "2026-05": 930853.62,
    }
    for month, val in expected.items():
        assert month in monthly
        assert abs(monthly[month] - val) < 0.05
    assert any(m.is_partial for m in perf.monthly if m.month == "2026-05")

    by_name = {v.vendor_name: v for v in perf.vendors}
    assert abs(float(by_name["Apex Bar & Tube Co."].received_value) - 23978261.68) < 0.05
    assert abs(by_name["Apex Bar & Tube Co."].otd_required - 0.827) < 0.005
    assert abs(by_name["Continental Quality Heat Treat"].otd_required - 0.391) < 0.005
    assert abs(by_name["Continental Quality Heat Treat"].otd_promise - 0.566) < 0.01
    assert perf.call_first is not None
    assert perf.call_first.vendor_name == "Continental Quality Heat Treat"


def test_no_hallucinated_missing_line(pipeline):
    # Apex 0001 ack does not contain BAR-A286-250; extractor must not invent it.
    docs = [d for d in pipeline["docs"] if d.po_number == "PO-4500050001"]
    pns = []
    for d in docs:
        for ln in d.lines:
            pns.append(ln.customer_part_number or ln.vendor_part_number)
    assert "BAR-A286-250" not in pns
