"""Task 2 Excel: plant-manager vendor performance workbook."""

from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import xlsxwriter

from src.models import CrosswalkEntry
from src.vendor_performance import PerformanceResult

NAVY = "#1B365D"
WHITE = "#FFFFFF"
RED = "#C0392B"
AMBER = "#D4A017"
GREEN = "#1E8449"
BORDER = "#C5CDD6"
PARTIAL = "#F5CBA7"


def _fmt(wb, **kwargs):
    return wb.add_format(kwargs)


def _formats(wb):
    base = dict(font_name="Calibri", font_size=10, border=1, border_color=BORDER, valign="vcenter")
    return {
        "header": _fmt(wb, **base, bold=True, font_color=WHITE, bg_color=NAVY, text_wrap=True, align="center"),
        "title": _fmt(wb, font_name="Calibri", font_size=18, bold=True, font_color=NAVY),
        "subtitle": _fmt(wb, font_name="Calibri", font_size=11, italic=True, font_color="#4A5568"),
        "section": _fmt(wb, font_name="Calibri", font_size=13, bold=True, font_color=NAVY),
        "label": _fmt(wb, font_name="Calibri", font_size=10, bold=True, font_color="#34495E"),
        "kpi": _fmt(wb, font_name="Calibri", font_size=16, bold=True, font_color=NAVY, num_format="$#,##0"),
        "kpi_pct": _fmt(wb, font_name="Calibri", font_size=16, bold=True, font_color=NAVY, num_format="0.0%"),
        "kpi_int": _fmt(wb, font_name="Calibri", font_size=16, bold=True, font_color=NAVY, num_format="#,##0"),
        "callout": _fmt(
            wb,
            font_name="Calibri",
            font_size=11,
            text_wrap=True,
            valign="top",
            bg_color="#FDEDEC",
            font_color="#7B241C",
            border=1,
            border_color=RED,
        ),
        "text": _fmt(wb, **base, text_wrap=True),
        "text_wrap": _fmt(wb, font_name="Calibri", font_size=10, text_wrap=True, valign="top"),
        "money": _fmt(wb, **base, num_format="$#,##0.00"),
        "pct": _fmt(wb, **base, num_format="0.0%"),
        "num": _fmt(wb, **base, num_format="#,##0.0"),
        "int": _fmt(wb, **base, num_format="#,##0"),
        "date": _fmt(wb, **base, num_format="yyyy-mm-dd"),
        "partial": _fmt(wb, **base, bg_color=PARTIAL, num_format="$#,##0.00"),
        "note": _fmt(wb, font_name="Calibri", font_size=10, text_wrap=True, valign="top"),
    }


def _write_header(ws, headers, fmts, widths, row=0):
    ws.set_row(row, 32)
    for i, h in enumerate(headers):
        ws.write(row, i, h, fmts["header"])
        ws.set_column(i, i, widths[i] if i < len(widths) else 14)
    ws.freeze_panes(row + 1, 0)
    ws.autofilter(row, 0, row, len(headers) - 1)


def write_task2_workbook(
    path: Path,
    result: PerformanceResult,
    crosswalk: list[CrosswalkEntry],
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(path))
    fmts = _formats(wb)

    # ----- EXECUTIVE_SUMMARY -----
    ws = wb.add_worksheet("EXECUTIVE_SUMMARY")
    ws.set_column(0, 0, 28)
    ws.set_column(1, 1, 22)
    ws.set_column(2, 8, 16)
    ws.hide_gridlines(2)

    ws.write(0, 0, "Beacon Fasteners — Vendor Performance", fmts["title"])
    period = f"{result.period_start.isoformat() if result.period_start else '—'} through {result.as_of.isoformat()}"
    ws.write(1, 0, f"Historical purchasing extract  ·  analysis as-of {result.as_of.isoformat()}  ·  {period}", fmts["subtitle"])

    ws.write(3, 0, "Analysis as-of", fmts["label"])
    ws.write_datetime(3, 1, datetime(result.as_of.year, result.as_of.month, result.as_of.day), fmts["date"])
    ws.write(4, 0, "Total received value", fmts["label"])
    ws.write_number(4, 1, float(result.total_received), fmts["kpi"])
    ws.write(5, 0, "Vendors analyzed", fmts["label"])
    ws.write_number(5, 1, len(result.vendors), fmts["kpi_int"])
    ws.write(6, 0, "Due PO lines", fmts["label"])
    ws.write_number(6, 1, result.total_due_lines, fmts["kpi_int"])
    ws.write(7, 0, "Overall required-date OTD", fmts["label"])
    if result.overall_otd_required is not None:
        ws.write_number(7, 1, result.overall_otd_required, fmts["kpi_pct"])
    start_month = result.period_start.strftime("%Y-%m") if result.period_start else None
    as_of_month = result.as_of.strftime("%Y-%m")
    ws.write(8, 0, "Partial months", fmts["label"])
    if result.period_start:
        ws.write(
            8,
            1,
            (
                f"{start_month} PARTIAL (extract starts {result.period_start.isoformat()}); "
                f"{as_of_month} PARTIAL (receipts through {result.as_of.isoformat()})"
            ),
            fmts["subtitle"],
        )
    else:
        ws.write(8, 1, f"{as_of_month} PARTIAL (receipts through {result.as_of.isoformat()})", fmts["subtitle"])

    # Monthly table for the chart
    ws.write(10, 0, "Received value by month", fmts["section"])
    ws.write(11, 0, "Month", fmts["header"])
    ws.write(11, 1, "Received value", fmts["header"])
    ws.write(11, 2, "Partial?", fmts["header"])
    for i, m in enumerate(result.monthly):
        ws.write(12 + i, 0, m.month, fmts["text"])
        money_fmt = fmts["partial"] if m.is_partial else fmts["money"]
        ws.write_number(12 + i, 1, float(m.received_value), money_fmt)
        ws.write(12 + i, 2, "PARTIAL" if m.is_partial else "", fmts["text"])

    n_months = len(result.monthly)
    chart1 = wb.add_chart({"type": "column"})
    chart1.add_series(
        {
            "name": "Received value",
            "categories": ["EXECUTIVE_SUMMARY", 12, 0, 11 + n_months, 0],
            "values": ["EXECUTIVE_SUMMARY", 12, 1, 11 + n_months, 1],
            "fill": {"color": NAVY},
        }
    )
    chart1.set_title({"name": "Monthly received value (USD)"})
    chart1.set_y_axis({"num_format": "$#,##0", "major_gridlines": {"visible": False}})
    chart1.set_legend({"none": True})
    chart1.set_size({"width": 620, "height": 280})
    chart1.set_style(10)
    ws.insert_chart(10, 4, chart1)

    # Vendor OTD table + chart
    start = 12 + n_months + 2
    ws.write(start, 0, "Required-date OTD by vendor", fmts["section"])
    ws.write(start + 1, 0, "Vendor", fmts["header"])
    ws.write(start + 1, 1, "OTD required", fmts["header"])
    ws.write(start + 1, 2, "OTD promise", fmts["header"])
    for i, v in enumerate(result.vendors):
        ws.write(start + 2 + i, 0, v.vendor_name, fmts["text"])
        if v.otd_required is not None:
            ws.write_number(start + 2 + i, 1, v.otd_required, fmts["pct"])
        if v.otd_promise is not None:
            ws.write_number(start + 2 + i, 2, v.otd_promise, fmts["pct"])

    n_v = len(result.vendors)
    chart2 = wb.add_chart({"type": "bar"})
    chart2.add_series(
        {
            "name": "Required-date OTD",
            "categories": ["EXECUTIVE_SUMMARY", start + 2, 0, start + 1 + n_v, 0],
            "values": ["EXECUTIVE_SUMMARY", start + 2, 1, start + 1 + n_v, 1],
            "fill": {"color": NAVY},
        }
    )
    chart2.add_series(
        {
            "name": "Promise-date OTD",
            "categories": ["EXECUTIVE_SUMMARY", start + 2, 0, start + 1 + n_v, 0],
            "values": ["EXECUTIVE_SUMMARY", start + 2, 2, start + 1 + n_v, 2],
            "fill": {"color": "#5D6D7E"},
        }
    )
    chart2.set_title({"name": "On-time delivery by vendor"})
    chart2.set_x_axis({"num_format": "0%", "min": 0, "max": 1})
    chart2.set_size({"width": 620, "height": 280})
    chart2.set_style(10)
    ws.insert_chart(start, 4, chart2)

    call_row = start + n_v + 4
    ws.write(call_row, 0, "Who should we call first?", fmts["section"])
    if result.call_first:
        ws.merge_range(call_row + 1, 0, call_row + 6, 3, result.call_first.explanation, fmts["callout"])
        ws.set_row(call_row + 1, 28)
    else:
        ws.write(call_row + 1, 0, "Insufficient due-line volume to recommend a call.", fmts["subtitle"])

    # ----- VENDOR_COMPARISON -----
    ws = wb.add_worksheet("VENDOR_COMPARISON")
    headers = [
        "Vendor ID", "Vendor Name", "Received Value", "Due Lines", "Completed Due Lines",
        "Required-Date OTD %", "Promise-Due Lines", "Promise-Date OTD %",
        "Average Days Late - Completed Late Lines", "Incomplete Overdue Lines",
        "Average Promise Slippage", "% Promises Later Than Required",
        "Avg Positive Promise Pushout", "Average Confirmation Lag",
        "QC Holds", "Open QC Holds",
    ]
    widths = [12, 34, 16, 12, 18, 18, 16, 18, 22, 18, 18, 22, 22, 18, 12, 14]
    _write_header(ws, headers, fmts, widths)
    # Sort: lowest required OTD first (operational pain), then spend
    vendors = sorted(
        result.vendors,
        key=lambda v: (v.otd_required if v.otd_required is not None else 9, -(v.received_value or 0)),
    )
    for i, v in enumerate(vendors, start=1):
        vals = [
            v.vendor_id,
            v.vendor_name,
            float(v.received_value or 0),
            v.due_lines,
            v.completed_due_lines,
            v.otd_required,
            v.promise_due_lines,
            v.otd_promise,
            v.avg_days_late,
            v.incomplete_overdue,
            v.avg_promise_slippage,
            v.pct_promises_later_than_required,
            v.avg_positive_pushout,
            v.avg_confirmation_lag,
            v.qc_holds,
            v.open_qc_holds,
        ]
        kinds = [
            "t", "t", "m", "i", "i", "p", "i", "p", "n", "i", "n", "p", "n", "n", "i", "i",
        ]
        for c, (val, kind) in enumerate(zip(vals, kinds)):
            if val is None:
                ws.write_blank(i, c, None, fmts["text"])
            elif kind == "m":
                ws.write_number(i, c, float(val), fmts["money"])
            elif kind == "p":
                ws.write_number(i, c, float(val), fmts["pct"])
            elif kind == "i":
                ws.write_number(i, c, int(val), fmts["int"])
            elif kind == "n":
                ws.write_number(i, c, float(val), fmts["num"])
            else:
                ws.write(i, c, val, fmts["text"])
    if vendors:
        ws.autofilter(0, 0, len(vendors), len(headers) - 1)
        ws.conditional_format(
            1, 5, len(vendors), 5,
            {"type": "3_color_scale", "min_color": "#C0392B", "mid_color": "#F4D03F", "max_color": "#1E8449"},
        )

    # ----- MONTHLY_RECEIVED_VALUE -----
    ws = wb.add_worksheet("MONTHLY_RECEIVED_VALUE")
    _write_header(ws, ["Month", "Received Value (USD)", "Partial month?"], fmts, [14, 22, 16])
    for i, m in enumerate(result.monthly, start=1):
        ws.write(i, 0, m.month, fmts["text"])
        ws.write_number(i, 1, float(m.received_value), fmts["partial"] if m.is_partial else fmts["money"])
        ws.write(i, 2, "PARTIAL" if m.is_partial else "", fmts["text"])
    chart = wb.add_chart({"type": "column"})
    chart.add_series(
        {
            "name": "Received value",
            "categories": ["MONTHLY_RECEIVED_VALUE", 1, 0, len(result.monthly), 0],
            "values": ["MONTHLY_RECEIVED_VALUE", 1, 1, len(result.monthly), 1],
            "fill": {"color": NAVY},
        }
    )
    chart.set_title({"name": "Received value by month"})
    chart.set_y_axis({"num_format": "$#,##0"})
    chart.set_legend({"none": True})
    chart.set_size({"width": 720, "height": 320})
    ws.insert_chart(1, 4, chart)

    # ----- LATE_LINES -----
    ws = wb.add_worksheet("LATE_LINES")
    headers = [
        "Vendor", "PO Number", "Line", "Part", "Qty Ordered", "Required Date",
        "Completion Date", "Promised Date", "Days Late vs Required",
        "Days Late vs Promise", "Incomplete?", "Line received value",
    ]
    widths = [32, 16, 8, 16, 12, 14, 16, 14, 18, 16, 12, 18]
    _write_header(ws, headers, fmts, widths)
    for i, ln in enumerate(result.late_lines, start=1):
        ws.write(i, 0, ln.vendor_name, fmts["text"])
        ws.write(i, 1, ln.po_number, fmts["text"])
        ws.write_number(i, 2, ln.line_no, fmts["int"])
        ws.write(i, 3, ln.part_id or "", fmts["text"])
        ws.write_number(i, 4, float(ln.qty_ordered), fmts["int"])
        if ln.required_date:
            ws.write_datetime(i, 5, datetime(ln.required_date.year, ln.required_date.month, ln.required_date.day), fmts["date"])
        if ln.completion_date:
            ws.write_datetime(i, 6, datetime(ln.completion_date.year, ln.completion_date.month, ln.completion_date.day), fmts["date"])
        else:
            ws.write(i, 6, "", fmts["text"])
        if ln.promised_date:
            ws.write_datetime(i, 7, datetime(ln.promised_date.year, ln.promised_date.month, ln.promised_date.day), fmts["date"])
        else:
            ws.write(i, 7, "", fmts["text"])
        if ln.days_late_required is not None:
            ws.write_number(i, 8, ln.days_late_required, fmts["int"])
        if ln.days_late_promise is not None:
            ws.write_number(i, 9, ln.days_late_promise, fmts["int"])
        else:
            ws.write(i, 9, "", fmts["text"])
        ws.write(i, 10, "Y" if ln.incomplete else "N", fmts["text"])
        ws.write_number(i, 11, float(ln.received_value), fmts["money"])
    if result.late_lines:
        ws.autofilter(0, 0, len(result.late_lines), len(headers) - 1)

    # ----- PART_CROSSWALK -----
    ws = wb.add_worksheet("PART_CROSSWALK")
    headers = [
        "Vendor ID", "Vendor Name", "Vendor PN", "Beacon PN", "Evidence Count",
        "Vendor PN Total", "Purity", "First Seen", "Last Seen", "Status", "Source",
    ]
    widths = [12, 32, 16, 16, 14, 14, 10, 13, 13, 12, 24]
    _write_header(ws, headers, fmts, widths)
    for i, e in enumerate(crosswalk, start=1):
        ws.write(i, 0, e.vendor_id, fmts["text"])
        ws.write(i, 1, e.vendor_name or "", fmts["text"])
        ws.write(i, 2, e.vendor_pn, fmts["text"])
        ws.write(i, 3, e.beacon_pn, fmts["text"])
        ws.write_number(i, 4, e.evidence_count, fmts["int"])
        ws.write_number(i, 5, e.vendor_pn_total_count, fmts["int"])
        ws.write_number(i, 6, e.purity, fmts["pct"])
        if e.first_seen:
            ws.write_datetime(i, 7, datetime(e.first_seen.year, e.first_seen.month, e.first_seen.day), fmts["date"])
        if e.last_seen:
            ws.write_datetime(i, 8, datetime(e.last_seen.year, e.last_seen.month, e.last_seen.day), fmts["date"])
        ws.write(i, 9, e.status.value, fmts["text"])
        ws.write(i, 10, e.source, fmts["text"])
    if crosswalk:
        ws.autofilter(0, 0, len(crosswalk), len(headers) - 1)

    # ----- DATA_NOTES -----
    ws = wb.add_worksheet("DATA_NOTES")
    ws.set_column(0, 0, 100)
    ws.write(0, 0, "Data notes, definitions, and assumptions", fmts["title"])
    notes = [
        "Definitions",
        "• Received value: signed receipt quantity × PO line unit price (Beacon base-currency valuation).",
        "• Due line: po_line.required_date <= analysis as-of date.",
        "• Stable completion date: earliest receipt date at which cumulative signed qty >= ordered AND never falls below ordered afterward.",
        "• OTD to required date: completion_date exists AND completion_date <= required_date. Incomplete overdue lines are not on time.",
        "• OTD to vendor promise: evaluated only when promised_date <= as-of. Same completion-date rule.",
        "• Promise slippage: promised_date - required_date (positive means vendor already committed later than Beacon asked).",
        "• Confirmation lag: confirmation.doc_date - po_header.po_date.",
        "",
        "Assumptions",
        "1. Open-PO / ERP unit prices are Beacon base-currency USD values. The open PO CSV has no currency column.",
        "2. Ostmark document prices are EUR and are normalized with ERP fx_rate (amount × rate_to_usd) when compared in Task 1.",
        "3. QuickShip states a ship date, which may not equal Beacon's required receipt date.",
        "4. required_date is treated as Beacon's needed/completion date for OTD.",
        "5. A missing confirmation price is NOT interpreted as price acceptance.",
        "6. Invoices are not automatically treated as acknowledgments.",
        "7. QC holds are associated to supplier POs but not necessarily proven supplier-caused defects.",
        "8. MRP messages are part-level and are not used as a core vendor-performance metric (vendor attribution is ambiguous).",
        (
            f"9. Boundary months "
            f"{result.period_start.strftime('%Y-%m') if result.period_start else 'start'} and "
            f"{result.as_of.strftime('%Y-%m')} are PARTIAL: the extract starts "
            f"{result.period_start.isoformat() if result.period_start else 'n/a'} and ends "
            f"{result.as_of.isoformat()}. Dollar values are the receipts that exist in those months; "
            "they are not annualized."
        ),
        "10. New vendor part mappings are PROPOSED unless historical evidence_count >= 3 and purity >= 0.95.",
        "11. Receipt reversals use the signed qty stored in receipt_txn. RV rows are negative; they are not abs()'d or sign-flipped again.",
        "12. Active confirmations are rows where superseded_by IS NULL. If duplicates remain, the latest doc_date / conf_id is used.",
        "",
        "Source tables",
        "receipt_txn, po_line, po_header, vendor_master, confirmation, qc_hold, fx_rate, part_master.",
        "The ERP database is opened read-only (sqlite URI mode=ro) and is never modified.",
        "",
        "Limitations",
        "• Calendar days are used for lateness, not plant-calendar workdays.",
        "• FX rates are monthly; small rounding differences vs an intra-month rate are expected.",
        "• This is a decision-support prototype, not a production ERP integration.",
        "",
        "Questions I Would Ask Lisa / Plant Manager",
        "• Are open-PO prices always USD?",
        "• Does required_date mean arrival at Beacon, ship date, or another milestone?",
        "• Should a supplier's lower price also be flagged, or only increases?",
        "• What price variance tolerance does purchasing actually use?",
        "• Are split/partial acknowledgments normal and acceptable?",
        "• When a supplier gives a delivery window (e.g. KW 20–22), how should purchasing judge it?",
        "• Does on-time mean first receipt or full ordered quantity received? (This tool uses full quantity.)",
        "• Should plant-calendar workdays rather than calendar days be used for lateness?",
        "• Are all QC holds supplier-responsible?",
        "• Can vendor part mappings ever be auto-approved into a master file?",
        "• How should an invoice/shipping notice be handled when no formal acknowledgment was received?",
        "• Which discrepancies cause the most business pain today?",
    ]
    for i, line in enumerate(notes, start=2):
        style = fmts["section"] if line and not line.startswith("•") and not line.startswith("1") and line[0].isupper() and "•" not in line and len(line) < 60 else fmts["note"]
        if line in {"Definitions", "Assumptions", "Source tables", "Limitations", "Questions I Would Ask Lisa / Plant Manager"}:
            style = fmts["section"]
        ws.write(i, 0, line, style)
        if "call" in line.lower():
            pass

    wb.close()
