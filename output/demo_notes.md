# Demo notes (for a 30-minute presentation / live walkthrough)

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

As-of date is **2026-05-17**, not today. Boundary months 2025-09 and 2026-05 are labeled PARTIAL.

Call-first (computed from lowest required-date OTD among vendors with enough due lines):
**Continental Quality Heat Treat**

Continental Quality Heat Treat is the strongest candidate for an immediate supplier discussion. Only 39.1% of due lines were completed by Beacon's required date, and 56.6% met Continental's own promise date. The issue appears upstream as well: 91.2% of Continental confirmations promise later than Beacon requested. Completed late lines averaged 15.6 calendar days late to the required date. When the vendor pushes dates out, the average positive promise slippage is 9.3 days. Received value in the extract is $1,508,567.40. This discussion should focus on capacity, quoted lead times, and realistic commitment dates rather than only expediting individual late orders. QC holds on this supplier's POs are shown for context and are not assumed to be supplier-caused.

## Questions to ask in the room

- Are PO prices always USD?
- Does required_date mean dock date at Beacon?
- How should KW delivery windows be judged?
- Should invoices ever stand in for acknowledgments?

## What this is not

Not a production ERP integration. Not an email sender. Not a supplier score model.
