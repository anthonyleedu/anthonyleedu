"""FX conversion using ERP monthly rates.

fx_rate.rate_to_usd means: amount_in_currency * rate_to_usd = USD amount.
"""

from __future__ import annotations

from datetime import date
from decimal import Decimal

from src.config import BASE_CURRENCY, FX_LOOKBACK_MONTHS
from src.normalize import add_months, month_key, round_money


class FxTable:
    def __init__(self, rows: list[dict] | None = None):
        self._rates: dict[tuple[str, str], Decimal] = {}
        for row in rows or []:
            month = str(row.get("month") or "").strip()
            currency = str(row.get("currency") or "").strip().upper()
            rate = row.get("rate_to_usd")
            if not month or not currency or rate is None:
                continue
            self._rates[(month, currency)] = Decimal(str(rate))
            # Identity for USD if missing
        if not any(c == "USD" for _, c in self._rates):
            pass

    def rate(
        self,
        currency: str | None,
        as_of: date | None,
        *,
        lookback_months: int = FX_LOOKBACK_MONTHS,
    ) -> tuple[Decimal | None, str | None, str | None]:
        """Return (rate, month_used, warning)."""
        cur = (currency or BASE_CURRENCY).upper()
        if cur == "USD" or cur == BASE_CURRENCY:
            # Still try table; identity otherwise.
            if as_of:
                mk = month_key(as_of)
                r = self._rates.get((mk, cur))
                if r is not None:
                    return r, mk, None
            return Decimal("1"), month_key(as_of) if as_of else None, None

        if as_of is None:
            return None, None, "No document date available for FX month selection"

        mk = month_key(as_of)
        r = self._rates.get((mk, cur))
        if r is not None:
            return r, mk, None

        # nearest earlier month
        cursor = date(as_of.year, as_of.month, 1)
        for _ in range(lookback_months):
            cursor = add_months(cursor, -1)
            mk2 = month_key(cursor)
            r = self._rates.get((mk2, cur))
            if r is not None:
                return r, mk2, f"Used earlier FX month {mk2} (document month {mk} missing)"
        return None, None, f"No FX rate for {cur} near {mk}"

    def to_usd(
        self,
        amount: Decimal | None,
        currency: str | None,
        as_of: date | None,
    ) -> tuple[Decimal | None, Decimal | None, str | None, str | None]:
        """Return (usd_amount, rate, month_used, warning)."""
        if amount is None:
            return None, None, None, None
        rate, month_used, warning = self.rate(currency, as_of)
        if rate is None:
            return None, None, month_used, warning
        return round_money(amount * rate, 6), rate, month_used, warning
