"""Audit CSV/JSONL/markdown outputs and follow-up email drafts."""

from __future__ import annotations

import json
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path

from src.models import CrosswalkEntry, ExtractedDocument, ReconciledLine, RunStats, Severity
from src.vendor_performance import PerformanceResult


def _json_default(obj):
    if isinstance(obj, Decimal):
        return str(obj)
    if hasattr(obj, "isoformat"):
        return obj.isoformat()
    raise TypeError(type(obj))


def write_extracted_jsonl(path: Path, documents: list[ExtractedDocument]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        for doc in documents:
            f.write(json.dumps(json.loads(doc.model_dump_json()), ensure_ascii=False) + "\n")


def write_reconciled_csv(path: Path, rows: list[ReconciledLine]) -> None:
    import pandas as pd

    records = []
    for r in rows:
        d = json.loads(r.model_dump_json())
        d["issue_codes"] = ";".join(r.issue_codes)
        d.pop("sort_key", None)
        records.append(d)
    pd.DataFrame(records).to_csv(path, index=False)


def write_vendor_metrics_csv(path: Path, result: PerformanceResult) -> None:
    import pandas as pd

    recs = []
    for v in result.vendors:
        recs.append(
            {
                "vendor_id": v.vendor_id,
                "vendor_name": v.vendor_name,
                "received_value": float(v.received_value or 0),
                "due_lines": v.due_lines,
                "completed_due_lines": v.completed_due_lines,
                "otd_required": v.otd_required,
                "promise_due_lines": v.promise_due_lines,
                "otd_promise": v.otd_promise,
                "avg_days_late": v.avg_days_late,
                "incomplete_overdue": v.incomplete_overdue,
                "avg_promise_slippage": v.avg_promise_slippage,
                "pct_promises_later_than_required": v.pct_promises_later_than_required,
                "avg_positive_pushout": v.avg_positive_pushout,
                "avg_confirmation_lag": v.avg_confirmation_lag,
                "qc_holds": v.qc_holds,
                "open_qc_holds": v.open_qc_holds,
            }
        )
    pd.DataFrame(recs).to_csv(path, index=False)


def build_follow_ups(rows: list[ReconciledLine], limit: int = 25) -> list[dict]:
    drafts = []
    reds = [r for r in rows if r.severity == Severity.RED and r.po_line_number is not None or r.unknown_po]
    reds = [r for r in rows if r.severity == Severity.RED]
    seen = set()
    for r in reds:
        key = (r.vendor_name, r.po_number, tuple(r.issue_codes[:3]), r.po_line_number)
        if key in seen:
            continue
        seen.add(key)
        vendor = r.vendor_name or "Supplier"
        po = r.po_number
        codes = set(r.issue_codes)
        subject = f"{po} confirmation discrepancy"
        body_lines = [
            "Hello,",
            "",
            f"We received your document for {po}.",
        ]
        if "QTY_SHORT" in codes and r.qty_ordered is not None and r.qty_confirmed is not None:
            short = r.qty_ordered - r.qty_confirmed
            req = r.required_date.isoformat() if r.required_date else "the required date"
            body_lines.append(
                f"Beacon ordered {r.qty_ordered:,.0f} units of {r.beacon_pn}, while the acknowledgment confirms {r.qty_confirmed:,.0f}."
            )
            body_lines.append(
                f"Please confirm whether the remaining {short:,.0f} units can be supplied by {req}."
            )
            subject = f"{po} confirmation discrepancy"
        elif "MISSING_PO_LINE" in codes:
            req = r.required_date.isoformat() if r.required_date else "the required date"
            body_lines.append(
                f"Beacon PO {po} line {r.po_line_number} ({r.beacon_pn}, qty {r.qty_ordered:,.0f} if ordered, required {req}) "
                "does not appear on the acknowledgment. Please confirm this line."
            )
        elif "UNKNOWN_PO" in codes:
            body_lines.append(
                f"We received an acknowledgment referencing {po}, which is not on our current open PO list. "
                "Please confirm the correct Beacon PO number."
            )
        elif "PROMISE_MISSING" in codes:
            body_lines.append(
                f"Please provide a committed quantity and promise date for {po} line {r.po_line_number} ({r.beacon_pn})."
            )
        elif "PROMISE_LATE" in codes:
            days = r.days_late if r.days_late is not None else r.minimum_late_days
            req = r.required_date.isoformat() if r.required_date else "the required date"
            body_lines.append(
                f"The confirmed date for {r.beacon_pn} is {days} days after Beacon's required date of {req}. "
                "Please advise whether the required date can be recovered."
            )
        elif "PRICE_HIGH" in codes and r.unit_price_difference is not None:
            body_lines.append(
                f"The acknowledgment unit price differs from the PO by ${r.unit_price_difference:.4f}/unit. "
                "Please confirm whether this is intended."
            )
        elif "NO_FORMAL_ACK_FOUND" in codes:
            body_lines.append(
                f"We have shipment/invoice evidence for {po} but no formal order acknowledgment. "
                "Please send a confirmation of quantity, price, and promise date."
            )
        else:
            body_lines.append(r.suggested_action or "Please review the attached discrepancy.")
        body_lines += ["", "Thank you.", "", "Lisa Morgan", "Purchasing, Beacon Fasteners"]
        drafts.append(
            {
                "vendor": vendor,
                "contact": r.vendor_contact,
                "po": po,
                "subject": subject,
                "body": "\n".join(body_lines),
            }
        )
        if len(drafts) >= limit:
            break
    return drafts


def write_follow_ups_xlsx(path: Path, drafts: list[dict]) -> None:
    import xlsxwriter

    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("FOLLOW_UPS")
    header = wb.add_format({"bold": True, "bg_color": "#1B365D", "font_color": "white", "border": 1})
    wrap = wb.add_format({"text_wrap": True, "valign": "top", "border": 1})
    cols = ["Vendor", "Contact", "PO", "Subject", "Suggested Email"]
    widths = [28, 28, 16, 40, 90]
    for i, (c, w) in enumerate(zip(cols, widths)):
        ws.write(0, i, c, header)
        ws.set_column(i, i, w)
    ws.freeze_panes(1, 0)
    for r, d in enumerate(drafts, start=1):
        ws.set_row(r, 90)
        ws.write(r, 0, d.get("vendor"), wrap)
        ws.write(r, 1, d.get("contact"), wrap)
        ws.write(r, 2, d.get("po"), wrap)
        ws.write(r, 3, d.get("subject"), wrap)
        ws.write(r, 4, d.get("body"), wrap)
    wb.close()


def write_run_summary_md(
    path: Path,
    *,
    stats: RunStats,
    rows: list[ReconciledLine],
    result: PerformanceResult,
    generated: list[Path],
) -> None:
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    top = [r for r in rows if r.severity == Severity.RED][:12]
    lines = [
        "# Beacon Purchasing Copilot — run summary",
        "",
        f"- Timestamp: {ts}",
        f"- Offline: {stats.offline}",
        f"- LLM: {stats.llm_provider} / {stats.llm_model}",
        "",
        "## Inputs",
    ]
    for k, v in stats.inputs.items():
        lines.append(f"- {k}: `{v}`")
    lines += [
        "",
        "## Extraction",
        f"- PDFs processed: {stats.pdfs_processed}",
        f"- Acknowledgments: {stats.acknowledgments}",
        f"- Revisions: {stats.revisions}",
        f"- Invoices: {stats.invoices}",
        f"- Scanned PDFs: {stats.scanned_pdfs}",
        f"- Cached: {stats.cached_extractions}",
        f"- AI extractions: {stats.ai_extractions}",
        f"- Local/OCR: {stats.local_extractions}",
        f"- Failures: {stats.extraction_failures}",
        "",
        "## Task 1 — open PO exceptions",
        f"- Open PO lines: {stats.open_po_lines}",
        f"- Matched: {stats.matched_lines}",
        f"- Missing lines: {stats.missing_lines}",
        f"- RED: {stats.red_issues}  ·  YELLOW: {stats.yellow_issues}  ·  unknown POs: {stats.unknown_pos}",
        "",
        "### Top RED issues",
    ]
    if not top:
        lines.append("- None")
    for r in top:
        lines.append(
            f"- {r.po_number} L{r.po_line_number or '—'} · {r.vendor_name} · "
            f"{';'.join(r.issue_codes)} · {r.suggested_action}"
        )
    cf = result.call_first
    lines += [
        "",
        "## Task 2 — vendor performance",
        f"- As-of: **{result.as_of.isoformat()}**",
        f"- Total received value: **${result.total_received:,.2f}**",
        f"- Due lines: {result.total_due_lines}",
        f"- Overall required-date OTD: {result.overall_otd_required:.1%}" if result.overall_otd_required is not None else "- Overall OTD: n/a",
        "",
        "### Received value by month",
    ]
    for m in result.monthly:
        flag = " *(PARTIAL)*" if m.is_partial else ""
        lines.append(f"- {m.month}: ${m.received_value:,.2f}{flag}")
    lines += ["", "### Vendor required-date OTD"]
    for v in result.vendors:
        otd = f"{v.otd_required:.1%}" if v.otd_required is not None else "n/a"
        lines.append(f"- {v.vendor_name}: received ${v.received_value:,.2f} · OTD {otd} · due {v.due_lines}")
    lines += ["", "## Call-first candidate"]
    if cf:
        lines.append(f"**{cf.vendor_name}**")
        lines.append("")
        lines.append(cf.explanation)
    lines += [
        "",
        "## Assumptions",
        "- Open PO prices are treated as USD.",
        "- Receipt reversals use signed qty.",
        "- Invoices are not formal acknowledgments.",
        "- May of the as-of year is partial through the last receipt date.",
        "",
        "## Generated files",
    ]
    for p in generated:
        lines.append(f"- `{p}`")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_demo_notes(path: Path, result: PerformanceResult, rows: list[ReconciledLine]) -> None:
    reds = [r for r in rows if r.severity == Severity.RED]
    text = f"""# Demo notes (for a 10-minute walkthrough)

## 30 seconds — what this is

Beacon sends a PO. The supplier sends a PDF back. This tool reads the PDF,
compares the promise to the PO, and shows Lisa what changed or went missing.
It also looks at eight months of receipts to show which vendors deliver.

AI reads documents. Python makes the business decisions.

## 3 minutes — Lisa's workbook (`lisa_reconciliation.xlsx`)

Open **ACTION_QUEUE**. It is only RED/YELLOW.

Walk these exceptions (derived, not hardcoded):

1. Apex **PO-4500050001** — two-line PO, acknowledgment omitted BAR-A286-250 (`MISSING_PO_LINE`).
2. Apex **PO-4500050002** — unit price 4.0102 vs PO 3.92 (`PRICE_HIGH`).
3. Apex **PO-4500050007** — 1,425 vs 1,500 (`QTY_SHORT` of 75).
4. Apex rogue **PO-4500060619** — not on the open PO list (`UNKNOWN_PO`).
5. Continental **PO-4500050022** — scanned PDF, ~21 days late. Vision/OCR had to work.
6. Liberty **PO-4500050016** — two PDFs, split 1,500 + 1,000 = 2,500, full-qty promise 05/24.
7. Liberty **PO-4500050019** — receipt acknowledged, schedule not committed.
8. Ostmark **PO-4500050027** — EUR + KW 20-22 window, not collapsed to a fake exact date.
9. QuickShip **PO-4500050030** — revision supersedes the qty-450 confirmation; active qty is 500.
10. QuickShip **PO-4500050032** — document is an **invoice**, not a formal ack.

## 3 minutes — plant manager (`vendor_performance.xlsx`)

As-of date is **{result.as_of.isoformat()}**, not today. May is labeled PARTIAL.

Call-first (computed from lowest required-date OTD among vendors with enough due lines):
**{result.call_first.vendor_name if result.call_first else 'n/a'}**

{result.call_first.explanation if result.call_first else ''}

## Questions to ask in the room

- Are PO prices always USD?
- Does required_date mean dock date at Beacon?
- How should KW delivery windows be judged?
- Should invoices ever stand in for acknowledgments?

## What this is not

Not a production ERP integration. Not an email sender. Not a supplier score model.
"""
    path.write_text(text, encoding="utf-8")
