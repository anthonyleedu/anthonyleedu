# Beacon Purchasing Copilot — run summary

- Timestamp: 2026-09-28 19:27 UTC
- Offline: True
- LLM: local / local-parser-v1

## Inputs
- open_pos: `/workspace/data/open_pos.csv`
- vendor_master: `/workspace/data/vendor_master.csv`
- erp: `/workspace/data/beacon_erp.db`
- confirmations_dir: `/workspace/data/confirmations`
- pdf_count: `35`
- skipped_pdfs: `(none)`

## Extraction
- PDFs processed: 35
- Acknowledgments: 32
- Revisions: 1
- Invoices: 1
- Scanned PDFs: 4
- Native text + OpenAI: 0
- Vision + OpenAI: 0
- Local parser/OCR fallback: 35
- Required review: 0
- Cached: 35
- Fresh AI extractions: 0
- Fresh local/OCR: 0
- Failures: 0

## Task 1 — open PO exceptions
- Open PO lines: 44
- Matched: 40
- Missing lines: 1
- RED: 10  ·  YELLOW: 4  ·  unknown POs: 1

### Top RED issues
- PO-4500050001 L2 · Apex Bar & Tube Co. · MISSING_PO_LINE · Ask Apex Bar & Tube Co. to confirm PO-4500050001 line 2, BAR-A286-250, qty 1,500, required 2026-06-10.
- PO-4500050007 L1 · Apex Bar & Tube Co. · QTY_SHORT · Resolve 75-unit short confirmation on PO-4500050007 line 1 (BAR-A286-250) before release.
- PO-4500050032 L1 · QuickShip Industrial · NO_FORMAL_ACK_FOUND;INVOICE_NOT_ACK · No formal acknowledgment found for PO-4500050032 line 1. Request an order confirmation (the invoice/shipping notice is not sufficient).
- PO-4500050032 L2 · QuickShip Industrial · NO_FORMAL_ACK_FOUND;INVOICE_NOT_ACK · No formal acknowledgment found for PO-4500050032 line 2. Request an order confirmation (the invoice/shipping notice is not sufficient).
- PO-4500060619 L— · Apex Bar & Tube Co. · UNKNOWN_PO · Verify whether this acknowledgment belongs to Beacon and whether PO number PO-4500060619 is correct.
- PO-4500050002 L1 · Apex Bar & Tube Co. · PRICE_HIGH · Confirm acceptance of $0.0902/unit increase (~2.30%) on PO-4500050002 line 1.
- PO-4500050022 L1 · Continental Quality Heat Treat · PRICE_NOT_STATED;PROMISE_LATE · Ask Continental Quality Heat Treat whether required date 2026-05-21 can be recovered; current promise is 21 days late.
- PO-4500050012 L1 · Heritage Cold Heading · PROMISE_LATE · Ask Heritage Cold Heading whether required date 2026-06-04 can be recovered; current promise is 14 days late.
- PO-4500050016 L1 · Liberty Surface Finishing · PRICE_NOT_STATED;PROMISE_LATE · Ask Liberty Surface Finishing whether required date 2026-05-17 can be recovered; current promise is 7 days late.
- PO-4500050019 L1 · Liberty Surface Finishing · PROMISE_MISSING · Request committed quantity and promise date for PO-4500050019 line 1 (PLAT-PASV).

## Task 2 — vendor performance
- As-of: **2026-05-17**
- Total received value: **$28,196,739.87**
- Due lines: 930
- Overall required-date OTD: 74.3%

### Received value by month
- 2025-09: $114,453.14
- 2025-10: $3,533,227.93
- 2025-11: $4,785,787.08
- 2025-12: $3,668,417.52
- 2026-01: $5,091,552.47
- 2026-02: $3,126,206.31
- 2026-03: $3,722,784.93
- 2026-04: $3,223,456.87
- 2026-05: $930,853.62 *(PARTIAL)*

### Vendor required-date OTD
- Apex Bar & Tube Co.: received $23,978,261.68 · OTD 82.7% · due 225
- Heritage Cold Heading: received $1,721,393.87 · OTD 69.3% · due 205
- Continental Quality Heat Treat: received $1,508,567.40 · OTD 39.1% · due 138
- Ostmark Werkzeug GmbH: received $867,316.80 · OTD 66.7% · due 57
- Liberty Surface Finishing: received $107,952.00 · OTD 85.5% · due 193
- QuickShip Industrial: received $13,248.12 · OTD 94.6% · due 112

## Call-first candidate
**Continental Quality Heat Treat**

Continental Quality Heat Treat is the strongest candidate for an immediate supplier discussion. Only 39.1% of due lines were completed by Beacon's required date, and 56.6% met Continental's own promise date. The issue appears upstream as well: 91.2% of Continental confirmations promise later than Beacon requested. Completed late lines averaged 15.6 calendar days late to the required date. When the vendor pushes dates out, the average positive promise slippage is 9.3 days. Received value in the extract is $1,508,567.40. This discussion should focus on capacity, quoted lead times, and realistic commitment dates rather than only expediting individual late orders. QC holds on this supplier's POs are shown for context and are not assumed to be supplier-caused.

## Assumptions
- Open PO prices are treated as USD.
- Receipt reversals use signed qty.
- Invoices are not formal acknowledgments.
- May of the as-of year is partial through the last receipt date.

## Generated files
- `output/lisa_reconciliation.xlsx`
- `output/vendor_performance.xlsx`
- `output/extracted_documents.jsonl`
- `output/reconciled_po_lines.csv`
- `output/vendor_metrics.csv`
- `output/run_summary.md`
- `output/follow_up_drafts.xlsx`
- `output/demo_notes.md`
