"""Stable completion-date tests using signed receipt quantities."""

from datetime import date
from decimal import Decimal

from src.erp import stable_completion_date


def _r(txn_id, qty, day, action="R"):
    return {"txn_id": txn_id, "qty": qty, "txn_date": day, "action_type": action}


def test_one_full_receipt():
    recs = [_r(1, 100, "2026-01-10")]
    assert stable_completion_date(recs, 100) == date(2026, 1, 10)


def test_multiple_partials():
    recs = [_r(1, 40, "2026-01-05"), _r(2, 60, "2026-01-12")]
    assert stable_completion_date(recs, 100) == date(2026, 1, 12)


def test_partial_then_completion():
    recs = [_r(1, 25, "01/02/2026"), _r(2, 25, "01/08/2026"), _r(3, 50, "01/20/2026")]
    assert stable_completion_date(recs, 100) == date(2026, 1, 20)


def test_receipt_then_reversal_below_ordered():
    recs = [_r(1, 100, "2026-01-10", "R"), _r(2, -40, "2026-01-15", "RV")]
    assert stable_completion_date(recs, 100) is None


def test_receipt_reversal_rereceipt_stable_date():
    recs = [
        _r(1, 100, "2026-01-10", "R"),
        _r(2, -100, "2026-01-12", "RV"),
        _r(3, 100, "2026-01-20", "R"),
    ]
    # First time cumulative >= ordered is Jan 10, but it later falls below; stable date is Jan 20.
    assert stable_completion_date(recs, 100) == date(2026, 1, 20)


def test_incomplete_line():
    recs = [_r(1, 40, "2026-01-10")]
    assert stable_completion_date(recs, 100) is None


def test_reversal_uses_signed_qty_not_abs():
    # If someone abs()'d RV, this would look complete; signed logic keeps it incomplete.
    recs = [_r(1, 80, "2026-01-10"), _r(2, -20, "2026-01-11", "RV")]
    assert stable_completion_date(recs, 100) is None
    net = sum(r["qty"] for r in recs)
    assert net == 60
