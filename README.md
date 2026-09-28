# Beacon Purchasing Copilot

Beacon sends a PO telling a supplier what it wants. The supplier sends a PDF back saying what it promises. This tool reads the PDF, compares the promise to Beacon's PO, and highlights anything that changed or went missing. It also looks at historical receipts to show which vendors reliably deliver.

Two workbooks are the product:

- `output/lisa_reconciliation.xlsx` — what Lisa opens every morning
- `output/vendor_performance.xlsx` — what the plant manager opens for the supplier discussion

There is no web app. Run one command, open the Excel files.

## Why AI

Supplier acknowledgments are messy: scans, German layouts, emails-saved-as-PDF, revisions, split confirmations, invoices pretending to be confirmations.

**AI reads documents. Deterministic Python makes business decisions.**

The language model is asked only to transcribe facts that are actually printed on the PDF. It never sees Beacon's open PO list, so it cannot hallucinate a silently omitted line. Quantity variance, price variance, on-time logic, joins, FX, and "who do we call first?" are all ordinary Python.

## Architecture

```
PDF
  |
  v
native PDF text (PyMuPDF)
  |
  +---- good text ----> structured LLM extraction (or local parser if no API key)
  |
  +---- scan/poor text -> render pages -> vision LLM
                                      or tesseract OCR + local parser
                                      |
                                      v
                              validated structured JSON
                                      |
                                      v
                            deterministic Python logic
                           /                          \
                          v                            v
                PO reconciliation                ERP analysis
                          |                            |
                          v                            v
              lisa_reconciliation.xlsx        vendor_performance.xlsx
```

## Installation

Python 3.11+. Tesseract is optional but required for Continental-style scans when no vision LLM key is configured.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# optional, for scanned PDFs without an LLM key
sudo apt-get install -y tesseract-ocr
```

## Configuration

```bash
cp .env.example .env
```

| Variable | Purpose |
|---|---|
| `LLM_PROVIDER` | `openai`, `anthropic`, or `local` |
| `LLM_MODEL` | e.g. `gpt-4o-mini` |
| `OPENAI_API_KEY` / `ANTHROPIC_API_KEY` | never committed |
| `LLM_TIMEOUT_SECONDS` | default 60 |
| `BASE_CURRENCY` | assumed currency of the open-PO CSV (default `USD`) |

If no API key is present, the tool falls back to a local text/OCR parser. That keeps the sample dataset reproducible for an evaluator who does not want to spend tokens. Extraction results are cached under `cache/extractions/` keyed by PDF SHA-256, prompt version, schema version, and model.

## Run

With the repository layout (`data/` contains the CSVs, SQLite extract, and confirmation PDFs):

```bash
python main.py --data-dir ./data --output ./output
```

Explicit paths:

```bash
python main.py \
    --confirmations ./data/confirmations \
    --open-pos ./data/open_pos.csv \
    --vendor-master ./data/vendor_master.csv \
    --erp ./data/beacon_erp.db \
    --output ./output
```

Useful flags:

```bash
python main.py --ping-llm
python main.py --data-dir ./data --output ./output --refresh-cache
python main.py --data-dir ./data --output ./output --offline
```

`--ping-llm` loads `.env` and makes one tiny authenticated OpenAI call. It logs key presence and length, never the key itself.

If an API key is present, extraction authenticates first. A rejected key aborts the run so local-parser fallback cannot be cached as if it were an OpenAI result. Use `--offline` when you want the local/OCR path on purpose.

`--offline` makes no API calls. Valid cache entries are reused; anything else is parsed locally / via OCR.

`--data-dir` discovers `*open_pos*.csv`, `*vendor_master*.csv`, `*beacon_erp*.db`, and confirmation PDFs recursively. Candidate-prompt PDFs and ERP exports are skipped. New vendor PDFs dropped into `data/confirmations/` are picked up on the next run with no code changes.

## Outputs

| File | Who it's for |
|---|---|
| `output/lisa_reconciliation.xlsx` | Buyer (Lisa) |
| `output/vendor_performance.xlsx` | Plant manager |
| `output/extracted_documents.jsonl` | Audit of every extraction |
| `output/reconciled_po_lines.csv` | Machine-readable Task 1 |
| `output/vendor_metrics.csv` | Machine-readable Task 2 |
| `output/run_summary.md` | Presentation notes for this run |
| `output/follow_up_drafts.xlsx` | Unsent email drafts for RED issues |
| `output/demo_notes.md` | 10-minute walkthrough script |

### Lisa's workbook

1. **ACTION_QUEUE** (first sheet) — only RED/YELLOW rows, sorted by severity, shortage, dollar impact, days late, PO.
2. **ALL_PO_LINES** — every open PO line plus unknown-PO documents.
3. **DOCUMENTS** — every PDF, including superseded revisions and invoices.
4. **PART_MAPPING** — historical + proposed vendor PN crosswalk.
5. **UNMATCHED_LINES** — extra or uncertain confirmation lines.
6. **RUN_SUMMARY** — counts for the run.
7. **FOLLOW_UPS** — draft emails; nothing is sent.

### Manager's workbook

1. **EXECUTIVE_SUMMARY** — as-of date, spend, OTD, monthly chart, vendor OTD chart, call-first write-up.
2. **VENDOR_COMPARISON**
3. **MONTHLY_RECEIVED_VALUE**
4. **LATE_LINES** — the rows behind the OTD math.
5. **PART_CROSSWALK**
6. **DATA_NOTES** — definitions, assumptions, limitations, questions.

## Tests

```bash
pytest
```

Unit tests cover dates (including `KW 20-22 / 2026`), FX, revisions vs splits vs invoices, matching, reconciliation, receipt reversals, and crosswalk approval. An integration module asserts the known sample cases and is skipped if the dataset is absent.

## Assumptions

1. Open-PO unit prices are Beacon base-currency **USD** because the CSV has no currency column.
2. Ostmark prices are EUR and are normalized with ERP `fx_rate.rate_to_usd` (`amount * rate = USD`).
3. QuickShip says **ship date**, which may not equal Beacon's required receipt date. Compared for visibility and flagged `DATE_SEMANTICS_WARNING`.
4. `required_date` is treated as Beacon's needed / completion date for OTD.
5. A missing confirmation price is **not** price acceptance.
6. Invoices are not formal acknowledgments.
7. QC holds are associated to a supplier's POs, not proven supplier-caused defects.
8. MRP messages are part-level and are **not** used as a vendor metric.
9. The last month of the extract is **PARTIAL** through `MAX(receipt_txn.txn_date)` (not today's date).
10. New vendor part mappings are **PROPOSED** unless historical `evidence_count >= 3` and `purity >= 0.95`.
11. Receipt reversals use the **signed** `qty` in the database. `RV` rows are already negative; they are not abs()'d.
12. On-time means the **stable full-quantity completion date**, not first receipt.

The ERP file is opened read-only:

```python
sqlite3.connect(f"file:{path}?mode=ro", uri=True)
```

## Known limitations

- This is a decision-support prototype, not a production ERP integration.
- Local parser coverage is tuned to native-text layouts plus Continental OCR. Unusual new templates should go through the LLM path.
- Lateness is calendar days, not plant-calendar workdays.
- Monthly FX can introduce small differences versus an intra-month rate.
- Follow-up emails are drafts only.
- Cache files must not be treated as a source of secrets (keys are never written there).

## Questions I Would Ask Lisa / Plant Manager

- Are open-PO prices always USD?
- Does `required_date` mean arrival at Beacon, ship date, or another milestone?
- Should a supplier's lower price also be flagged, or only increases?
- What price variance tolerance does purchasing actually use?
- Are split/partial acknowledgments normal and acceptable?
- When a supplier gives a delivery window (KW 20–22), how should purchasing judge it?
- Does on-time mean first receipt or full ordered quantity received? (This tool uses full quantity.)
- Should plant-calendar workdays rather than calendar days be used for lateness?
- Are all QC holds supplier-responsible?
- Can vendor part mappings ever be auto-approved into a master file?
- How should an invoice/shipping notice be handled when no formal acknowledgment was received?
- Which discrepancies cause the most business pain today?

## Future improvements

- Email inbox ingestion instead of a folder of PDFs
- Human approval workflow for proposed vendor PN mappings
- Persistent vendor-PN master written back (with audit) to a controlled store — never to the source ERP extract
- Direct ERP integration for live open POs
- Automated (but reviewed) supplier follow-ups
- A lightweight web UI only after the Excel workflow is validated with Lisa

Do not treat this as production-ready purchasing software. It is a complete, honest take-home that an operations user can open tomorrow morning.
