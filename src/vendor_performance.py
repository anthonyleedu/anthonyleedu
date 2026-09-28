"""Historical vendor performance from ERP receipts and confirmations.

All metrics are derived from the database. Nothing is hardcoded.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from statistics import mean

from src.config import MIN_DUE_LINES_FOR_CALL
from src.erp import analysis_as_of, completion_map
from src.models import CrosswalkEntry
from src.normalize import month_key, parse_date, round_money


@dataclass
class MonthlyValue:
    month: str
    received_value: Decimal
    is_partial: bool = False


@dataclass
class LateLine:
    vendor_id: str
    vendor_name: str
    po_number: str
    line_no: int
    part_id: str | None
    qty_ordered: Decimal
    required_date: date | None
    completion_date: date | None
    promised_date: date | None
    days_late_required: int | None
    days_late_promise: int | None
    incomplete: bool
    received_value: Decimal


@dataclass
class VendorMetrics:
    vendor_id: str
    vendor_name: str
    received_value: Decimal = Decimal("0")
    due_lines: int = 0
    completed_due_lines: int = 0
    on_time_required: int = 0
    promise_due_lines: int = 0
    on_time_promise: int = 0
    late_days_completed: list[int] = field(default_factory=list)
    incomplete_overdue: int = 0
    promise_slippage: list[int] = field(default_factory=list)
    positive_slippage: list[int] = field(default_factory=list)
    confirmation_lag: list[int] = field(default_factory=list)
    qc_holds: int = 0
    open_qc_holds: int = 0

    @property
    def otd_required(self) -> float | None:
        if not self.due_lines:
            return None
        return self.on_time_required / self.due_lines

    @property
    def otd_promise(self) -> float | None:
        if not self.promise_due_lines:
            return None
        return self.on_time_promise / self.promise_due_lines

    @property
    def avg_days_late(self) -> float | None:
        if not self.late_days_completed:
            return None
        return mean(self.late_days_completed)

    @property
    def avg_promise_slippage(self) -> float | None:
        if not self.promise_slippage:
            return None
        return mean(self.promise_slippage)

    @property
    def pct_promises_later_than_required(self) -> float | None:
        if not self.promise_slippage:
            return None
        return sum(1 for s in self.promise_slippage if s > 0) / len(self.promise_slippage)

    @property
    def avg_positive_pushout(self) -> float | None:
        if not self.positive_slippage:
            return None
        return mean(self.positive_slippage)

    @property
    def avg_confirmation_lag(self) -> float | None:
        if not self.confirmation_lag:
            return None
        return mean(self.confirmation_lag)


@dataclass
class CallFirst:
    vendor_id: str
    vendor_name: str
    explanation: str
    otd_required: float | None
    otd_promise: float | None
    received_value: Decimal
    avg_days_late: float | None
    pct_promises_later: float | None
    avg_positive_pushout: float | None


@dataclass
class PerformanceResult:
    as_of: date
    period_start: date | None
    monthly: list[MonthlyValue]
    vendors: list[VendorMetrics]
    late_lines: list[LateLine]
    call_first: CallFirst | None
    total_received: Decimal
    overall_otd_required: float | None
    total_due_lines: int
    notes: list[str]


def _pick_active_confirmation(rows: list[dict]) -> dict | None:
    """Active confirmation = superseded_by IS NULL; if duplicates, latest doc_date then conf_id."""
    active = [r for r in rows if r.get("superseded_by") in (None, "", 0)]
    if not active:
        return None
    active.sort(key=lambda r: (parse_date(r.get("doc_date")) or date.min, r.get("conf_id") or 0))
    return active[-1]


def analyze_vendor_performance(
    data: dict[str, list[dict]],
    *,
    crosswalk: list[CrosswalkEntry] | None = None,
) -> PerformanceResult:
    receipts = data["receipt_txn"]
    po_lines = data["po_line"]
    po_headers = {r["po_number"]: r for r in data["po_header"]}
    vendors = {r["vendor_id"]: r for r in data["vendor_master"]}
    qc_holds = data.get("qc_hold") or []
    confirmations = data.get("confirmation") or []

    as_of = analysis_as_of(receipts)
    if as_of is None:
        raise RuntimeError("receipt_txn has no parseable txn_date values")

    completions = completion_map(po_lines, receipts)
    pl_map = {(r["po_number"], int(r["line_no"])): r for r in po_lines}

    # Monthly received value from dated receipt transactions (signed qty * PO unit price)
    monthly_acc: dict[str, Decimal] = defaultdict(lambda: Decimal("0"))
    vendor_val: dict[str, Decimal] = defaultdict(lambda: Decimal("0"))
    line_receipt_value: dict[tuple[str, int], Decimal] = defaultdict(lambda: Decimal("0"))
    txn_dates: list[date] = []
    for r in receipts:
        d = parse_date(r.get("txn_date"))
        if d is None:
            continue
        txn_dates.append(d)
        pl = pl_map.get((r["po_number"], int(r["line_no"])))
        if not pl:
            continue
        val = Decimal(str(r["qty"])) * Decimal(str(pl["unit_price"]))
        monthly_acc[d.strftime("%Y-%m")] += val
        hdr = po_headers.get(r["po_number"])
        if hdr:
            vendor_val[hdr["vendor_id"]] += val
        line_receipt_value[(r["po_number"], int(r["line_no"]))] += val

    period_start = min(txn_dates) if txn_dates else None
    start_month = period_start.strftime("%Y-%m") if period_start else None
    as_of_month = as_of.strftime("%Y-%m")
    monthly = [
        MonthlyValue(
            month=m,
            received_value=round_money(v, 2) or Decimal("0"),
            is_partial=(m == as_of_month or (start_month is not None and m == start_month)),
        )
        for m, v in sorted(monthly_acc.items())
    ]

    conf_by_line: dict[tuple[str, int], list[dict]] = defaultdict(list)
    for c in confirmations:
        conf_by_line[(c["po_number"], int(c["line_no"]))].append(c)

    metrics: dict[str, VendorMetrics] = {}
    for vid, v in vendors.items():
        metrics[vid] = VendorMetrics(vendor_id=vid, vendor_name=v["vendor_name"], received_value=round_money(vendor_val.get(vid, 0), 2) or Decimal("0"))

    late_lines: list[LateLine] = []

    for pl in po_lines:
        hdr = po_headers.get(pl["po_number"])
        if not hdr:
            continue
        vid = hdr["vendor_id"]
        m = metrics.get(vid)
        if m is None:
            continue
        key = (pl["po_number"], int(pl["line_no"]))
        req = parse_date(pl.get("required_date"))
        cd = completions.get(key)
        conf = _pick_active_confirmation(conf_by_line.get(key, []))
        promised = parse_date(conf.get("promised_date")) if conf else None
        po_date = parse_date(hdr.get("po_date"))
        doc_date = parse_date(conf.get("doc_date")) if conf else None

        if req is not None and req <= as_of:
            m.due_lines += 1
            if cd:
                m.completed_due_lines += 1
                if cd <= req:
                    m.on_time_required += 1
                else:
                    delay = (cd - req).days
                    m.late_days_completed.append(delay)
                    late_lines.append(
                        LateLine(
                            vendor_id=vid,
                            vendor_name=m.vendor_name,
                            po_number=pl["po_number"],
                            line_no=int(pl["line_no"]),
                            part_id=pl.get("part_id"),
                            qty_ordered=Decimal(str(pl["qty_ordered"])),
                            required_date=req,
                            completion_date=cd,
                            promised_date=promised,
                            days_late_required=delay,
                            days_late_promise=(cd - promised).days if promised else None,
                            incomplete=False,
                            received_value=round_money(line_receipt_value.get(key, 0), 2) or Decimal("0"),
                        )
                    )
            else:
                m.incomplete_overdue += 1
                late_lines.append(
                    LateLine(
                        vendor_id=vid,
                        vendor_name=m.vendor_name,
                        po_number=pl["po_number"],
                        line_no=int(pl["line_no"]),
                        part_id=pl.get("part_id"),
                        qty_ordered=Decimal(str(pl["qty_ordered"])),
                        required_date=req,
                        completion_date=None,
                        promised_date=promised,
                        days_late_required=(as_of - req).days,
                        days_late_promise=None,
                        incomplete=True,
                        received_value=round_money(line_receipt_value.get(key, 0), 2) or Decimal("0"),
                    )
                )

        if promised is not None and promised <= as_of:
            m.promise_due_lines += 1
            if cd and cd <= promised:
                m.on_time_promise += 1

        if promised is not None and req is not None:
            slip = (promised - req).days
            m.promise_slippage.append(slip)
            if slip > 0:
                m.positive_slippage.append(slip)

        if doc_date is not None and po_date is not None:
            m.confirmation_lag.append((doc_date - po_date).days)

    for q in qc_holds:
        hdr = po_headers.get(q.get("po_number"))
        if not hdr:
            continue
        m = metrics.get(hdr["vendor_id"])
        if not m:
            continue
        m.qc_holds += 1
        released = str(q.get("released") or "").strip().upper()
        if released in {"N", "0", "FALSE", "NO"}:
            m.open_qc_holds += 1

    vendor_list = sorted(metrics.values(), key=lambda v: (-(v.received_value or 0), v.vendor_name))
    total_received = sum((v.received_value or 0) for v in vendor_list)
    total_due = sum(v.due_lines for v in vendor_list)
    total_otd = sum(v.on_time_required for v in vendor_list)
    overall = (total_otd / total_due) if total_due else None

    eligible = [v for v in vendor_list if v.due_lines >= MIN_DUE_LINES_FOR_CALL and v.otd_required is not None]
    call = None
    if eligible:
        # Lowest required-date OTD; tie-break larger received value.
        eligible.sort(key=lambda v: (v.otd_required, -(v.received_value or 0)))
        winner = eligible[0]
        otd_r = winner.otd_required
        otd_p = winner.otd_promise
        pct_push = winner.pct_promises_later_than_required
        avg_late = winner.avg_days_late
        avg_push = winner.avg_positive_pushout
        explanation = (
            f"{winner.vendor_name} is the strongest candidate for an immediate supplier discussion. "
            f"Only {otd_r:.1%} of due lines were completed by Beacon's required date"
            + (f", and {otd_p:.1%} met {winner.vendor_name.split()[0]}'s own promise date" if otd_p is not None else "")
            + ". "
        )
        if pct_push is not None:
            explanation += (
                f"The issue appears upstream as well: {pct_push:.1%} of {winner.vendor_name.split()[0]} "
                f"confirmations promise later than Beacon requested. "
            )
        if avg_late is not None:
            explanation += (
                f"Completed late lines averaged {avg_late:.1f} calendar days late to the required date. "
            )
        if avg_push is not None:
            explanation += (
                f"When the vendor pushes dates out, the average positive promise slippage is {avg_push:.1f} days. "
            )
        explanation += (
            f"Received value in the extract is ${winner.received_value:,.2f}. "
            "This discussion should focus on capacity, quoted lead times, and realistic commitment dates "
            "rather than only expediting individual late orders. "
            "QC holds on this supplier's POs are shown for context and are not assumed to be supplier-caused."
        )
        call = CallFirst(
            vendor_id=winner.vendor_id,
            vendor_name=winner.vendor_name,
            explanation=explanation,
            otd_required=otd_r,
            otd_promise=otd_p,
            received_value=winner.received_value,
            avg_days_late=avg_late,
            pct_promises_later=pct_push,
            avg_positive_pushout=avg_push,
        )

    notes = [
        f"Analysis as-of date is MAX(receipt_txn.txn_date) = {as_of.isoformat()}, not today's date.",
        (
            f"Boundary months {start_month} and {as_of_month} are PARTIAL "
            f"because the extract starts {period_start.isoformat() if period_start else 'n/a'} "
            f"and ends {as_of.isoformat()}."
        ),
        "Received value = signed receipt_txn.qty * po_line.unit_price (Beacon base valuation).",
        "Reversals (action_type=RV) are stored as negative qty and are summed as signed values.",
        "On-time to required date uses the stable full-quantity completion date, not first receipt.",
        "Incomplete overdue lines count as NOT on time.",
        "Active vendor promises use confirmation.superseded_by IS NULL.",
        "MRP messages are not used as a vendor-performance metric (part-level, vendor-ambiguous).",
        "QC holds are associated to supplier POs but not proven supplier-caused.",
    ]

    late_lines.sort(key=lambda x: (-(x.days_late_required or 0), x.vendor_name, x.po_number))
    return PerformanceResult(
        as_of=as_of,
        period_start=period_start,
        monthly=monthly,
        vendors=vendor_list,
        late_lines=late_lines,
        call_first=call,
        total_received=round_money(total_received, 2) or Decimal("0"),
        overall_otd_required=overall,
        total_due_lines=total_due,
        notes=notes,
    )
