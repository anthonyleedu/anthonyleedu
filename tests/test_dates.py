"""Date and ISO-week parsing tests."""

from datetime import date

from src.normalize import compare_promise_to_required, parse_date, parse_iso_week_range, parse_promise
from src.models import DatePrecision, DateType


def test_us_slash_date():
    assert parse_date("05/21/2026") == date(2026, 5, 21)
    assert parse_date("Acknowledgment Date: 05/06/2026") == date(2026, 5, 6)


def test_iso_date():
    assert parse_date("2026-06-10") == date(2026, 6, 10)


def test_german_dotted_date_is_dmy():
    assert parse_date("01.05.2026") == date(2026, 5, 1)
    assert parse_date("30.04.2026") == date(2026, 4, 30)
    assert parse_date("08.06.2026") == date(2026, 6, 8)


def test_iso_week_exact():
    start, end, prec, raw = parse_iso_week_range("KW 20 / 2026")
    assert prec == DatePrecision.WEEK
    assert start == date.fromisocalendar(2026, 20, 1)
    assert end == date.fromisocalendar(2026, 20, 7)
    assert start.weekday() == 0
    assert end.weekday() == 6


def test_iso_week_range():
    start, end, prec, raw = parse_iso_week_range("KW 20-22 / 2026")
    assert prec == DatePrecision.WEEK_RANGE
    assert start == date.fromisocalendar(2026, 20, 1)  # Monday
    assert end == date.fromisocalendar(2026, 22, 7)  # Sunday
    assert raw.lower().startswith("kw")
    # Must not collapse to a single fake exact date
    assert start != end


def test_parse_promise_preserves_week_range():
    w = parse_promise("KW 20-22 / 2026")
    assert w.date_type == DateType.DELIVERY_WINDOW
    assert w.precision == DatePrecision.WEEK_RANGE
    assert w.start_date == date.fromisocalendar(2026, 20, 1)
    assert w.end_date == date.fromisocalendar(2026, 22, 7)
    assert "KW" in (w.raw_text or "").upper()


def test_range_straddles_required_date():
    w = parse_promise("KW 20-22 / 2026")
    required = date(2026, 5, 15)  # inside week 20-22
    cmp = compare_promise_to_required(w, required)
    assert cmp["straddles"] is True
    assert cmp["late"] is False
    assert cmp["on_time"] is False
    assert cmp["minimum_late_days"] == 0
    assert cmp["maximum_late_days"] == (w.end_date - required).days


def test_range_entirely_after_required_is_late():
    w = parse_promise("KW 20-22 / 2026")
    required = date(2026, 5, 1)
    cmp = compare_promise_to_required(w, required)
    assert cmp["late"] is True
    assert cmp["minimum_late_days"] == (w.start_date - required).days


def test_range_entirely_before_required_is_on_time():
    w = parse_promise("KW 20-22 / 2026")
    required = date(2026, 6, 15)
    cmp = compare_promise_to_required(w, required)
    assert cmp["on_time"] is True


def test_ship_date_semantics():
    w = parse_promise("ship 05/21/2026", header_hint="ship")
    assert w.date_type == DateType.SHIP_DATE
    assert w.start_date == date(2026, 5, 21)
    assert w.end_date == date(2026, 5, 21)
