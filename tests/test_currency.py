"""FX conversion tests."""

from datetime import date
from decimal import Decimal

from src.currency import FxTable
from src.normalize import parse_number


def _fx():
    return FxTable(
        [
            {"month": "2026-05", "currency": "USD", "rate_to_usd": 1.0},
            {"month": "2026-05", "currency": "EUR", "rate_to_usd": 1.0902},
            {"month": "2026-04", "currency": "EUR", "rate_to_usd": 1.0874},
        ]
    )


def test_usd_identity():
    fx = _fx()
    usd, rate, month, warn = fx.to_usd(Decimal("3.92"), "USD", date(2026, 5, 17))
    assert usd == Decimal("3.920000") or usd == Decimal("3.92")
    assert rate == Decimal("1")
    assert warn is None


def test_eur_conversion():
    fx = _fx()
    usd, rate, month, warn = fx.to_usd(Decimal("1140.80"), "EUR", date(2026, 5, 1))
    assert warn is None
    assert rate == Decimal("1.0902")
    assert month == "2026-05"
    expected = Decimal("1140.80") * Decimal("1.0902")
    assert abs(usd - expected) < Decimal("0.00001")


def test_missing_fx_rate():
    fx = FxTable([{"month": "2026-05", "currency": "USD", "rate_to_usd": 1.0}])
    usd, rate, month, warn = fx.to_usd(Decimal("10"), "JPY", date(2026, 5, 1))
    assert usd is None
    assert rate is None
    assert warn is not None
    assert "JPY" in warn


def test_parse_european_and_us_numbers():
    assert parse_number("1,425") == Decimal("1425")
    assert parse_number("$4.0102") == Decimal("4.0102")
    assert parse_number("1140.8000") == Decimal("1140.8000")
