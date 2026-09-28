"""Deterministic normalization utilities: POs, parts, vendors, numbers, dates."""

from __future__ import annotations

import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import Iterable

from src.models import DatePrecision, DateType, PromiseWindow

_PO_RE = re.compile(r"\bPO[-\s]?(\d{7,12})\b", re.IGNORECASE)
_ISO_WEEK_RANGE_RE = re.compile(
    r"\bKW\s*(\d{1,2})\s*[\-\u2013\u2014]+\s*(\d{1,2})\s*/\s*(\d{4})\b",
    re.IGNORECASE,
)
_ISO_WEEK_RE = re.compile(r"\bKW\s*(\d{1,2})\s*/\s*(\d{4})\b", re.IGNORECASE)
_ISO_WEEK_ALT_RE = re.compile(
    r"\b(?:ISO\s*)?WEEK(?:S)?\s*(\d{1,2})\s*(?:[\-\u2013\u2014]+\s*(\d{1,2}))?\s*(?:/|,)?\s*(\d{4})\b",
    re.IGNORECASE,
)

_US_DATE_RE = re.compile(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b")
_ISO_DATE_RE = re.compile(r"\b(\d{4})-(\d{2})-(\d{2})\b")
_DE_DATE_RE = re.compile(r"\b(\d{1,2})\.(\d{1,2})\.(\d{4})\b")
_US_DASH_DATE_RE = re.compile(r"\b(\d{1,2})-(\d{1,2})-(\d{4})\b")

_CURRENCY_RE = re.compile(r"\b(USD|EUR|GBP|CAD|JPY|CHF)\b", re.IGNORECASE)


def normalize_po_number(value: str | None) -> str | None:
    if value is None:
        return None
    text = str(value).strip().upper()
    if not text:
        return None
    text = text.replace(" ", "")
    text = text.strip(".:;#")
    m = _PO_RE.search(text) or _PO_RE.search("PO-" + text if text.isdigit() else text)
    if m:
        return f"PO-{m.group(1)}"
    # Preserve identifier if it already looks like a PO.
    if text.startswith("PO"):
        return text
    return text


def extract_po_number(text: str | None) -> str | None:
    if not text:
        return None
    m = _PO_RE.search(text)
    if m:
        return f"PO-{m.group(1)}"
    return None


def normalize_part_number(value: str | None) -> str | None:
    if value is None:
        return None
    text = str(value).strip().upper()
    if not text:
        return None
    text = re.sub(r"\s+", " ", text)
    text = text.replace(" ", "")
    text = text.strip(".,;:")
    # Do not strip meaningful hyphens.
    if text in {"N/A", "NA", "NONE", "-", "â€”"}:
        return None
    return text or None


def normalize_vendor_name(value: str | None) -> str | None:
    if value is None:
        return None
    text = re.sub(r"\s+", " ", str(value).strip())
    return text or None


def vendor_key(value: str | None) -> str:
    name = normalize_vendor_name(value) or ""
    return name.casefold()


def parse_number(value: str | int | float | Decimal | None) -> Decimal | None:
    if value is None or value == "":
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    text = (
        str(value)
        .strip()
        .replace("\u20ac", "")
        .replace("$", "")
        .replace("\u00a3", "")
        .replace("\u00a0", "")
        .replace(" ", "")
    )
    neg = False
    if text.startswith("(") and text.endswith(")"):
        neg = True
        text = text[1:-1]
    # European thousands: 1.234,56 vs US 1,234.56
    if re.search(r"^\d{1,3}(\.\d{3})+,\d+$", text):
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", "")
    text = re.sub(r"[^0-9.\-]", "", text)
    if text in {"", "-", ".", "-."}:
        return None
    try:
        n = Decimal(text)
    except InvalidOperation:
        return None
    return -n if neg else n


def parse_qty(value: str | int | float | Decimal | None) -> Decimal | None:
    return parse_number(value)


def detect_currency(text: str | None, default: str | None = None) -> str | None:
    if not text:
        return default
    if "\u20ac" in text or re.search(r"\bEUR\b|preise in\s*eur|euro", text, re.IGNORECASE):
        return "EUR"
    if "$" in text or re.search(r"\bUSD\b", text, re.IGNORECASE):
        return "USD"
    # Do not treat process codes such as CAD (cadmium) as Canadian dollars.
    m = re.search(r"\b(GBP|JPY|CHF)\b", text, re.IGNORECASE)
    if m:
        return m.group(1).upper()
    return default


def _safe_date(y: int, m: int, d: int) -> date | None:
    try:
        return date(y, m, d)
    except ValueError:
        return None


def parse_date(value: str | date | datetime | None, *, prefer_dmy: bool = False) -> date | None:
    """Parse mixed ERP/PDF date formats.

    Dotted dates (01.05.2026) are treated as DMY (German).
    Slashed dates (05/06/2026) are treated as MDY (US) unless prefer_dmy.
    ISO dates are unambiguous.
    """
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if not text:
        return None

    m = _ISO_DATE_RE.search(text)
    if m:
        return _safe_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))

    m = _DE_DATE_RE.search(text)
    if m:
        return _safe_date(int(m.group(3)), int(m.group(2)), int(m.group(1)))

    m = _US_DATE_RE.search(text)
    if m:
        a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if prefer_dmy:
            return _safe_date(y, b, a)
        return _safe_date(y, a, b)

    m = _US_DASH_DATE_RE.search(text)
    if m:
        a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if a > 12:
            return _safe_date(y, b, a)
        return _safe_date(y, a, b)

    for fmt in ("%Y/%m/%d", "%d-%m-%Y", "%b %d, %Y", "%B %d, %Y", "%d %b %Y", "%d %B %Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def iso_week_bounds(year: int, week: int) -> tuple[date, date]:
    start = date.fromisocalendar(year, week, 1)
    end = date.fromisocalendar(year, week, 7)
    return start, end


def parse_iso_week_range(text: str | None) -> tuple[date, date, DatePrecision, str] | None:
    if not text:
        return None
    m = _ISO_WEEK_RANGE_RE.search(text)
    if m:
        w1, w2, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        start, _ = iso_week_bounds(year, w1)
        _, end = iso_week_bounds(year, w2)
        return start, end, DatePrecision.WEEK_RANGE, m.group(0)
    m = _ISO_WEEK_RE.search(text)
    if m:
        week, year = int(m.group(1)), int(m.group(2))
        start, end = iso_week_bounds(year, week)
        return start, end, DatePrecision.WEEK, m.group(0)
    m = _ISO_WEEK_ALT_RE.search(text)
    if m:
        w1 = int(m.group(1))
        w2 = int(m.group(2)) if m.group(2) else w1
        year = int(m.group(3))
        start, _ = iso_week_bounds(year, w1)
        _, end = iso_week_bounds(year, w2)
        prec = DatePrecision.WEEK_RANGE if w2 != w1 else DatePrecision.WEEK
        return start, end, prec, m.group(0)
    return None


def infer_date_type(raw_text: str | None, header_hint: str | None = None) -> DateType:
    blob = f"{raw_text or ''} {header_hint or ''}".lower()
    if "kw" in blob or "week" in blob or "window" in blob or "-" in (raw_text or "") and "kw" in blob:
        if "kw" in blob or "week" in blob:
            return DateType.DELIVERY_WINDOW
    if re.search(r"\bship(?:ment|ping)?\b", blob):
        return DateType.SHIP_DATE
    if "liefertermin" in blob or "delivery" in blob:
        return DateType.DELIVERY_DATE
    if "promise" in blob:
        return DateType.PROMISE_DATE
    return DateType.UNKNOWN


def parse_promise(
    raw_text: str | None,
    *,
    header_hint: str | None = None,
    prefer_dmy: bool = False,
) -> PromiseWindow:
    raw = (raw_text or "").strip() or None
    window = PromiseWindow(raw_text=raw)
    if not raw:
        return window

    iso = parse_iso_week_range(raw)
    if iso:
        start, end, prec, matched = iso
        window.start_date = start
        window.end_date = end
        window.precision = prec
        window.date_type = DateType.DELIVERY_WINDOW
        window.raw_text = matched
        return window

    # Range of two calendar dates
    dates = []
    for rx, kind in (
        (_DE_DATE_RE, "dmy"),
        (_ISO_DATE_RE, "iso"),
        (_US_DATE_RE, "mdy"),
    ):
        for m in rx.finditer(raw):
            if kind == "dmy":
                d = _safe_date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            elif kind == "iso":
                d = _safe_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            else:
                d = _safe_date(int(m.group(3)), int(m.group(1)), int(m.group(2)))
            if d:
                dates.append(d)
        if dates:
            break

    if not dates:
        d = parse_date(raw, prefer_dmy=prefer_dmy)
        if d:
            dates = [d]

    if len(dates) >= 2:
        window.start_date = min(dates[0], dates[1])
        window.end_date = max(dates[0], dates[1])
        window.precision = DatePrecision.EXACT_DATE
        window.date_type = DateType.DELIVERY_WINDOW
    elif len(dates) == 1:
        window.start_date = dates[0]
        window.end_date = dates[0]
        window.precision = DatePrecision.EXACT_DATE
        window.date_type = infer_date_type(raw, header_hint)
        if window.date_type == DateType.UNKNOWN:
            window.date_type = infer_date_type(header_hint, None) or DateType.PROMISE_DATE
            if window.date_type == DateType.UNKNOWN:
                window.date_type = DateType.PROMISE_DATE
    else:
        window.precision = DatePrecision.TEXT_ONLY
        window.date_type = infer_date_type(raw, header_hint)

    if window.date_type == DateType.UNKNOWN:
        window.date_type = infer_date_type(raw, header_hint)
    return window


def compare_promise_to_required(
    promise: PromiseWindow,
    required: date | None,
) -> dict:
    """Evaluate a promise window against Beacon's required date.

    Returns on_time / late / straddles plus min/max late days.
    """
    result = {
        "status": "MISSING",
        "on_time": False,
        "late": False,
        "straddles": False,
        "days_late": None,
        "minimum_late_days": None,
        "maximum_late_days": None,
    }
    if required is None or not promise.is_usable:
        return result

    start = promise.start_date
    end = promise.latest
    if start is None and end is None:
        return result

    if start is None:
        start = end
    if end is None:
        end = start
    assert start is not None and end is not None

    min_late = (start - required).days
    max_late = (end - required).days

    if end <= required:
        result.update(
            status="ON_TIME",
            on_time=True,
            late=False,
            straddles=False,
            days_late=0,
            minimum_late_days=0,
            maximum_late_days=0,
        )
    elif start > required:
        result.update(
            status="LATE",
            on_time=False,
            late=True,
            straddles=False,
            days_late=min_late,
            minimum_late_days=min_late,
            maximum_late_days=max_late,
        )
    else:
        # start <= required < end
        result.update(
            status="STRADDLES",
            on_time=False,
            late=False,
            straddles=True,
            days_late=None,
            minimum_late_days=0,
            maximum_late_days=max_late,
        )
    return result


def first_present(values: Iterable[str | None]) -> str | None:
    for v in values:
        if v and str(v).strip():
            return str(v).strip()
    return None


def month_key(d: date | None) -> str | None:
    if d is None:
        return None
    return d.strftime("%Y-%m")


def add_months(d: date, months: int) -> date:
    y = d.year + (d.month - 1 + months) // 12
    m = (d.month - 1 + months) % 12 + 1
    return date(y, m, 1)


def round_money(value: Decimal | float | int | None, places: int = 4) -> Decimal | None:
    if value is None:
        return None
    q = Decimal("1").scaleb(-places)
    return Decimal(value).quantize(q)
