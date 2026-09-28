"""Task 1 Excel: Lisa's morning reconciliation workbook."""

from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import xlsxwriter

from src.file_discovery import display_path

from src.models import (
    CrosswalkEntry,
    ExtractedDocument,
    ReconciledLine,
    RunStats,
    Severity,
)

NAVY = "#1B365D"
NAVY_DARK = "#12243F"
WHITE = "#FFFFFF"
RED_BG = "#F8D7DA"
RED_FONT = "#721C24"
YELLOW_BG = "#FFF3CD"
YELLOW_FONT = "#856404"
GREEN_BG = "#D4EDDA"
GREEN_FONT = "#155724"
INFO_BG = "#E8EEF4"
GRAY = "#F4F6F8"
BORDER = "#C5CDD6"
HEADER_H = 36


def _fmt(wb, **kwargs):
    return wb.add_format(kwargs)


def _formats(wb):
    base = dict(font_name="Calibri", font_size=10, border=1, border_color=BORDER, valign="vcenter")
    return {
        "header": _fmt(wb, **base, bold=True, font_color=WHITE, bg_color=NAVY, text_wrap=True, align="center"),
        "title": _fmt(wb, font_name="Calibri", font_size=16, bold=True, font_color=NAVY),
        "subtitle": _fmt(wb, font_name="Calibri", font_size=10, italic=True, font_color="#4A5568"),
        "text": _fmt(wb, **base, text_wrap=True),
        "text_red": _fmt(wb, **base, text_wrap=True, bg_color=RED_BG, font_color=RED_FONT),
        "text_yellow": _fmt(wb, **base, text_wrap=True, bg_color=YELLOW_BG, font_color=YELLOW_FONT),
        "text_green": _fmt(wb, **base, text_wrap=True, bg_color=GREEN_BG, font_color=GREEN_FONT),
        "text_info": _fmt(wb, **base, text_wrap=True, bg_color=INFO_BG),
        "num": _fmt(wb, **base, num_format="#,##0.00"),
        "int": _fmt(wb, **base, num_format="#,##0"),
        "money": _fmt(wb, **base, num_format="$#,##0.0000"),
        "money2": _fmt(wb, **base, num_format="$#,##0.00"),
        "pct": _fmt(wb, **base, num_format="0.00%"),
        "date": _fmt(wb, **base, num_format="yyyy-mm-dd"),
        "sev_red": _fmt(wb, **base, bold=True, bg_color="#C0392B", font_color=WHITE, align="center"),
        "sev_yellow": _fmt(wb, **base, bold=True, bg_color="#D4A017", font_color=WHITE, align="center"),
        "sev_green": _fmt(wb, **base, bold=True, bg_color="#1E8449", font_color=WHITE, align="center"),
        "sev_info": _fmt(wb, **base, bold=True, bg_color="#5D6D7E", font_color=WHITE, align="center"),
        "label": _fmt(wb, font_name="Calibri", font_size=10, bold=True, font_color=NAVY),
        "kpi": _fmt(wb, font_name="Calibri", font_size=14, bold=True, font_color=NAVY),
    }


def _sev_fmt(fmts, sev: Severity):
    return {
        Severity.RED: fmts["sev_red"],
        Severity.YELLOW: fmts["sev_yellow"],
        Severity.GREEN: fmts["sev_green"],
        Severity.INFO: fmts["sev_info"],
    }[sev]


def _row_fmt(fmts, sev: Severity):
    return {
        Severity.RED: fmts["text_red"],
        Severity.YELLOW: fmts["text_yellow"],
        Severity.GREEN: fmts["text_green"],
        Severity.INFO: fmts["text_info"],
    }[sev]


def _write_header(ws, headers, fmts, widths):
    ws.set_row(0, HEADER_H)
    for i, h in enumerate(headers):
        ws.write(0, i, h, fmts["header"])
        ws.set_column(i, i, widths[i] if i < len(widths) else 14)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, 0, len(headers) - 1)
    ws.set_default_row(18)


def _cell(ws, r, c, value, fmt, kind="text"):
    if value is None or value == "":
        ws.write_blank(r, c, None, fmt)
        return
    if kind == "date" and isinstance(value, (date, datetime)):
        ws.write_datetime(r, c, datetime(value.year, value.month, value.day) if isinstance(value, date) and not isinstance(value, datetime) else value, fmt)
        return
    if isinstance(value, Decimal):
        ws.write_number(r, c, float(value), fmt)
        return
    if kind in {"num", "money", "pct", "int"} and isinstance(value, (int, float, Decimal)):
        ws.write_number(r, c, float(value), fmt)
        return
    ws.write(r, c, value, fmt)


def _basename(path: str | None) -> str | None:
    if not path:
        return None
    return Path(path).name


def _hyperlink(ws, r, c, path: str | None, fmt):
    if not path:
        ws.write_blank(r, c, None, fmt)
        return
    # Write a cwd-relative or basename path. Do not embed machine-absolute
    # file:// URIs in the workbook.
    shown = display_path(path) or Path(path).name
    ws.write(r, c, shown, fmt)


ACTION_COLUMNS = [
    ("Priority", 10),
    ("Severity", 10),
    ("PO Number", 16),
    ("PO Line", 10),
    ("Vendor", 28),
    ("Vendor Contact", 28),
    ("Beacon Part Number", 18),
    ("Vendor Part Number", 18),
    ("Description", 32),
    ("Issue Codes", 36),
    ("Suggested Action", 48),
    ("Qty Ordered", 12),
    ("Qty Confirmed", 14),
    ("Qty Variance", 12),
    ("PO Unit Price", 13),
    ("PO Currency", 12),
    ("Confirmation Unit Price", 18),
    ("Confirmation Currency", 16),
    ("FX Rate", 10),
    ("Normalized Confirmation Price", 20),
    ("Price Difference", 14),
    ("Price Difference %", 14),
    ("Estimated Price Impact", 16),
    ("Required Date", 13),
    ("Commitment Type", 16),
    ("Promise Start", 13),
    ("Promise End", 13),
    ("Projected Full Qty Promise", 18),
    ("Days Late", 10),
    ("Match Method", 22),
    ("Match Confidence", 14),
    ("Source Document", 22),
    ("Source Document Number", 18),
    ("Document Date", 13),
    ("Document Status", 16),
]


def _write_line_row(ws, r, row: ReconciledLine, fmts, *, include_priority: bool, priority: int | None):
    base = _row_fmt(fmts, row.severity)
    money = fmts["money"]
    money2 = fmts["money2"]
    pct = fmts["pct"]
    dt = fmts["date"]
    num = fmts["num"]
    codes = ";".join(row.issue_codes)
    vals = []
    if include_priority:
        vals.append(("int", priority))
    vals.extend(
        [
            ("sev", row.severity.value),
            ("text", row.po_number),
            ("int", row.po_line_number),
            ("text", row.vendor_name),
            ("text", row.vendor_contact),
            ("text", row.beacon_pn),
            ("text", row.vendor_pn),
            ("text", row.description),
            ("text", codes),
            ("text", row.suggested_action),
            ("num", row.qty_ordered),
            ("num", row.qty_confirmed),
            ("num", row.qty_variance),
            ("money", row.po_unit_price),
            ("text", row.po_currency_assumed),
            ("money", row.confirmed_unit_price_raw),
            ("text", row.confirmed_currency),
            ("num", row.fx_rate_used),
            ("money", row.confirmed_price_usd),
            ("money", row.unit_price_difference),
            ("pct", row.unit_price_difference_pct),
            ("money2", row.price_impact_on_order),
            ("date", row.required_date),
            ("text", row.commitment_date_type.value if row.commitment_date_type else None),
            ("date", row.promise_start),
            ("date", row.promise_end),
            ("date", row.projected_full_qty_promise_date),
            ("int", row.days_late),
            ("text", row.match_method.value if row.match_method else None),
            ("num", row.match_confidence),
            ("file", row.source_file),
            ("text", row.source_document_number),
            ("date", row.document_date),
            ("text", row.document_status),
        ]
    )
    for c, (kind, value) in enumerate(vals):
        if kind == "sev":
            ws.write(r, c, value, _sev_fmt(fmts, row.severity))
        elif kind == "file":
            _hyperlink(ws, r, c, value, base)
        elif kind == "date":
            _cell(ws, r, c, value, dt, "date")
        elif kind == "money":
            _cell(ws, r, c, value, money, "money")
        elif kind == "money2":
            _cell(ws, r, c, value, money2, "money")
        elif kind == "pct":
            _cell(ws, r, c, value, pct, "pct")
        elif kind == "num":
            _cell(ws, r, c, value, num, "num")
        elif kind == "int":
            _cell(ws, r, c, value, fmts["int"] if isinstance(value, (int, float, Decimal)) else base, "int")
        else:
            _cell(ws, r, c, value, base)


def write_task1_workbook(
    path: Path,
    *,
    rows: list[ReconciledLine],
    documents: list[ExtractedDocument],
    crosswalk: list[CrosswalkEntry],
    stats: RunStats,
    follow_ups: list[dict] | None = None,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(path))
    fmts = _formats(wb)

    actionable = [r for r in rows if r.is_actionable]
    actionable.sort(key=lambda r: r.sort_key)

    # ACTION_QUEUE
    ws = wb.add_worksheet("ACTION_QUEUE")
    headers = [h for h, _ in ACTION_COLUMNS]
    widths = [w for _, w in ACTION_COLUMNS]
    _write_header(ws, headers, fmts, widths)
    for i, row in enumerate(actionable, start=1):
        ws.set_row(i, 32)
        _write_line_row(ws, i, row, fmts, include_priority=True, priority=i)
    if actionable:
        ws.autofilter(0, 0, len(actionable), len(headers) - 1)
    ws.freeze_panes(1, 3)

    # ALL_PO_LINES
    ws = wb.add_worksheet("ALL_PO_LINES")
    all_headers = headers[1:]  # no Priority
    all_widths = widths[1:]
    _write_header(ws, all_headers, fmts, all_widths)
    po_rows = [r for r in rows if not (r.issue_codes == ["EXTRA_CONFIRMATION_LINE"] and r.po_line_number is None)]
    # Keep extras too actually — ALL_PO_LINES should be one row per Beacon PO line plus unknown POs
    po_rows = [r for r in rows if r.po_line_number is not None or r.unknown_po]
    for i, row in enumerate(po_rows, start=1):
        ws.set_row(i, 28)
        _write_line_row(ws, i, row, fmts, include_priority=False, priority=None)
    if po_rows:
        ws.autofilter(0, 0, len(po_rows), len(all_headers) - 1)

    # DOCUMENTS
    ws = wb.add_worksheet("DOCUMENTS")
    dheaders = [
        "Source File", "SHA256", "Vendor (raw)", "PO Number", "Document Number", "Document Date",
        "Document Type", "Commitment Status", "Lifecycle", "Revision?", "Supersedes prior?",
        "Part N", "Part M", "Currency", "Line Count", "Extraction Mode", "Model",
        "Cached?", "Warnings", "Notes",
    ]
    dwidths = [22, 16, 28, 16, 18, 13, 18, 28, 16, 12, 16, 10, 10, 10, 12, 16, 18, 10, 40, 30]
    _write_header(ws, dheaders, fmts, dwidths)
    for i, doc in enumerate(documents, start=1):
        vals = [
            display_path(doc.source_file) or Path(doc.source_file).name,
            doc.source_sha256[:12],
            doc.vendor_name_raw,
            doc.po_number,
            doc.document_number,
            doc.document_date,
            doc.document_type.value,
            doc.commitment_status.value,
            doc.lifecycle.value,
            "Y" if doc.is_revision else "N",
            "Y" if doc.supersedes_all_prior_for_po else "N",
            doc.partial_sequence,
            doc.partial_total,
            doc.document_currency,
            len(doc.lines),
            doc.extraction_mode.value,
            doc.model_name,
            "Y" if doc.cached else "N",
            "; ".join(doc.extraction_warnings),
            "; ".join(doc.free_text_notes),
        ]
        for c, v in enumerate(vals):
            if isinstance(v, date):
                _cell(ws, i, c, v, fmts["date"], "date")
            else:
                _hyperlink(ws, i, c, doc.source_file, fmts["text"]) if c == 0 else _cell(ws, i, c, v, fmts["text"])
    if documents:
        ws.autofilter(0, 0, len(documents), len(dheaders) - 1)

    # PART_MAPPING
    ws = wb.add_worksheet("PART_MAPPING")
    pheaders = [
        "Vendor ID", "Vendor Name", "Vendor PN", "Beacon PN", "Evidence Count",
        "Vendor PN Total", "Purity", "First Seen", "Last Seen", "Status", "Source",
    ]
    pwidths = [12, 32, 16, 16, 14, 14, 10, 13, 13, 12, 24]
    _write_header(ws, pheaders, fmts, pwidths)
    for i, e in enumerate(crosswalk, start=1):
        vals = [
            e.vendor_id, e.vendor_name, e.vendor_pn, e.beacon_pn, e.evidence_count,
            e.vendor_pn_total_count, e.purity, e.first_seen, e.last_seen, e.status.value, e.source,
        ]
        for c, v in enumerate(vals):
            if isinstance(v, date):
                _cell(ws, i, c, v, fmts["date"], "date")
            elif isinstance(v, float) and c == 6:
                _cell(ws, i, c, v, fmts["pct"], "pct")
            else:
                _cell(ws, i, c, v, fmts["text"])
    if crosswalk:
        ws.autofilter(0, 0, len(crosswalk), len(pheaders) - 1)

    # UNMATCHED_LINES
    ws = wb.add_worksheet("UNMATCHED_LINES")
    uheaders = ["PO Number", "Vendor", "Vendor PN", "Customer PN", "Description", "Qty", "Price",
                "Match Method", "Explanation", "Source File", "Issue Codes"]
    uwidths = [16, 28, 16, 16, 36, 12, 12, 22, 40, 22, 28]
    _write_header(ws, uheaders, fmts, uwidths)
    unmatched = [r for r in rows if r.match_method and r.match_method.value in {"UNMATCHED", "MANUAL_REVIEW"} and r.po_line_number is None]
    unmatched += [r for r in rows if "MISSING_PO_LINE" in r.issue_codes or "UNKNOWN_PO" in r.issue_codes]
    seen = set()
    uniq = []
    for r in unmatched:
        k = (r.po_number, r.po_line_number, r.source_file, tuple(r.issue_codes), r.vendor_pn)
        if k in seen:
            continue
        seen.add(k)
        uniq.append(r)
    for i, r in enumerate(uniq, start=1):
        vals = [
            r.po_number, r.vendor_name, r.vendor_pn, r.beacon_pn, r.description,
            r.qty_confirmed if r.qty_confirmed is not None else r.qty_ordered,
            r.confirmed_unit_price_raw, r.match_method.value if r.match_method else None,
            r.match_explanation, _basename(r.source_file), ";".join(r.issue_codes),
        ]
        for c, v in enumerate(vals):
            _cell(ws, i, c, v, _row_fmt(fmts, r.severity))
    if uniq:
        ws.autofilter(0, 0, len(uniq), len(uheaders) - 1)

    # RUN_SUMMARY
    ws = wb.add_worksheet("RUN_SUMMARY")
    ws.set_column(0, 0, 36)
    ws.set_column(1, 1, 48)
    ws.write(0, 0, "Beacon Purchasing Copilot — Run Summary", fmts["title"])
    ws.write(1, 0, "Task 1 · Open PO vs supplier acknowledgments", fmts["subtitle"])
    items = [
        ("PDFs processed", stats.pdfs_processed),
        ("Acknowledgments", stats.acknowledgments),
        ("Revisions", stats.revisions),
        ("Invoices", stats.invoices),
        ("Scanned PDFs", stats.scanned_pdfs),
        ("OpenAI-backed documents", stats.openai_backed),
        ("Native text + OpenAI (incl. cache)", stats.native_text_openai),
        ("Vision + OpenAI (incl. cache)", stats.vision_openai),
        ("Fresh AI extractions this run", stats.ai_extractions),
        ("Cached extractions", stats.cached_extractions),
        ("Local/OCR fallbacks", stats.local_fallback),
        ("Open PO lines", stats.open_po_lines),
        ("Matched lines", stats.matched_lines),
        ("Missing lines", stats.missing_lines),
        ("RED issues", stats.red_issues),
        ("YELLOW issues", stats.yellow_issues),
        ("Unknown POs", stats.unknown_pos),
        ("Extraction failures", stats.extraction_failures),
        ("LLM provider", stats.llm_provider),
        ("LLM model", stats.llm_model),
        ("Offline", "Y" if stats.offline else "N"),
    ]
    ws.write(3, 0, "Metric", fmts["header"])
    ws.write(3, 1, "Value", fmts["header"])
    for i, (k, v) in enumerate(items, start=4):
        ws.write(i, 0, k, fmts["label"])
        ws.write(i, 1, v if v is not None else "", fmts["text"])
    note_row = 4 + len(items) + 2
    ws.write(note_row, 0, "Assumptions", fmts["label"])
    ws.write(note_row + 1, 0, "Open PO unit prices are treated as USD (CSV has no currency column).", fmts["subtitle"])
    ws.write(note_row + 2, 0, "AI extracts document facts; Python computes qty/price/date discrepancies.", fmts["subtitle"])
    ws.write(note_row + 3, 0, "Invoices are not treated as formal acknowledgments.", fmts["subtitle"])
    ws.write(note_row + 4, 0, "QuickShip 'ship' dates are compared for visibility and flagged DATE_SEMANTICS_WARNING.", fmts["subtitle"])
    ws.write(
        note_row + 5,
        0,
        "Cached OpenAI results still count as OpenAI-backed; Fresh AI extractions this run is the API-call count.",
        fmts["subtitle"],
    )

    if follow_ups:
        ws = wb.add_worksheet("FOLLOW_UPS")
        fheaders = ["Vendor", "Contact", "PO", "Subject", "Suggested Email"]
        fwidths = [28, 28, 16, 40, 80]
        _write_header(ws, fheaders, fmts, fwidths)
        for i, fu in enumerate(follow_ups, start=1):
            ws.set_row(i, 80)
            for c, key in enumerate(["vendor", "contact", "po", "subject", "body"]):
                _cell(ws, i, c, fu.get(key), fmts["text"])
        ws.autofilter(0, 0, len(follow_ups), 4)

    wb.close()
