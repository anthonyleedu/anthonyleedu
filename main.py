#!/usr/bin/env python3
"""Beacon Purchasing Copilot — confirmation reconciliation + vendor performance."""

from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path

from src import config
from src.document_extractor import extract_all
from src.document_versions import resolve_active_documents
from src.erp import load_erp
from src.excel_task1 import write_task1_workbook
from src.excel_task2 import write_task2_workbook
from src.file_discovery import display_path, resolve_inputs
from src.currency import FxTable
from src.inputs import load_open_pos, load_vendor_master
from src.logging_utils import setup_logging
from src.llm_client import AuthenticationFailed, llm_available, ping_openai
from src.models import DocumentType, ExtractionMode, ReconciledLine, RunStats, Severity
from src.part_crosswalk import build_historical_crosswalk
from src.reconcile import reconcile
from src.reporting import (
    build_follow_ups,
    write_demo_notes,
    write_extracted_jsonl,
    write_follow_ups_xlsx,
    write_reconciled_csv,
    write_run_summary_md,
    write_vendor_metrics_csv,
)
from src.vendor_performance import analyze_vendor_performance


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Reconcile supplier confirmation PDFs to Beacon open POs and analyze vendor performance."
    )
    p.add_argument("--data-dir", type=Path, help="Discover CSVs, ERP DB, and confirmation PDFs under this directory.")
    p.add_argument("--confirmations", type=Path, help="Directory of vendor confirmation PDFs (recursive).")
    p.add_argument("--open-pos", type=Path, help="Open PO CSV path.")
    p.add_argument("--vendor-master", type=Path, help="Vendor master CSV path.")
    p.add_argument("--erp", type=Path, help="Read-only Beacon ERP SQLite database.")
    p.add_argument("--output", type=Path, default=Path("output"), help="Output directory.")
    p.add_argument("--cache-dir", type=Path, default=config.DEFAULT_CACHE_DIR, help="Extraction cache directory.")
    p.add_argument("--refresh-cache", action="store_true", help="Ignore extraction cache and re-extract.")
    p.add_argument("--offline", action="store_true", help="Do not call LLM APIs. Use cache and local/OCR parsers.")
    p.add_argument(
        "--ping-llm",
        action="store_true",
        help="Load .env, verify OPENAI_API_KEY is present, and make one tiny authenticated API call.",
    )
    return p.parse_args(argv)


def tally_stats(docs, rows: list[ReconciledLine], stats: RunStats, review_count: int = 0) -> None:
    stats.pdfs_processed = len(docs)
    stats.acknowledgments = sum(1 for d in docs if d.document_type == DocumentType.ACKNOWLEDGEMENT)
    stats.revisions = sum(1 for d in docs if d.document_type == DocumentType.REVISION)
    stats.invoices = sum(1 for d in docs if d.document_type == DocumentType.INVOICE)
    stats.scanned_pdfs = sum(
        1
        for d in docs
        if d.extraction_mode in {ExtractionMode.VISION, ExtractionMode.OCR_LOCAL} or (d.native_text_chars or 0) < 80
    )
    stats.cached_extractions = sum(1 for d in docs if d.cached)
    stats.native_text_openai = sum(1 for d in docs if d.extraction_mode == ExtractionMode.NATIVE_TEXT)
    stats.vision_openai = sum(1 for d in docs if d.extraction_mode == ExtractionMode.VISION)
    stats.local_fallback = sum(
        1 for d in docs if d.extraction_mode in {ExtractionMode.LOCAL_TEXT, ExtractionMode.OCR_LOCAL}
    )
    stats.openai_backed = stats.native_text_openai + stats.vision_openai
    stats.review_required = review_count
    stats.ai_extractions = sum(
        1 for d in docs if d.extraction_mode in {ExtractionMode.NATIVE_TEXT, ExtractionMode.VISION} and not d.cached
    )
    stats.local_extractions = sum(
        1 for d in docs if d.extraction_mode in {ExtractionMode.LOCAL_TEXT, ExtractionMode.OCR_LOCAL} and not d.cached
    )
    stats.vision_extractions = stats.vision_openai
    stats.extraction_failures = sum(1 for d in docs if d.extraction_mode == ExtractionMode.FAILED)
    stats.open_po_lines = sum(1 for r in rows if r.po_line_number is not None and not r.unknown_po)
    stats.matched_lines = sum(
        1 for r in rows if r.po_line_number is not None and r.qty_confirmed is not None and "MISSING_PO_LINE" not in r.issue_codes
    )
    stats.missing_lines = sum(1 for r in rows if "MISSING_PO_LINE" in r.issue_codes)
    stats.red_issues = sum(1 for r in rows if r.severity == Severity.RED)
    stats.yellow_issues = sum(1 for r in rows if r.severity == Severity.YELLOW)
    stats.unknown_pos = sum(1 for r in rows if r.unknown_po or "UNKNOWN_PO" in r.issue_codes)
    stats.warnings = sum(len(d.extraction_warnings) for d in docs)


def run(args: argparse.Namespace, logger: logging.Logger) -> int:
    stats = RunStats(
        started_at=datetime.now(timezone.utc),
        offline=args.offline,
        refresh_cache=args.refresh_cache,
        llm_provider=config.LLM_PROVIDER if llm_available() and not args.offline else "local",
        llm_model=config.LLM_MODEL if llm_available() and not args.offline else "local-parser-v1",
    )
    data_dir = args.data_dir
    if data_dir is None and args.open_pos is None:
        # Sensible default for this repository layout
        candidate = Path("data")
        if candidate.exists():
            data_dir = candidate

    discovered = resolve_inputs(
        data_dir=data_dir,
        confirmations=args.confirmations,
        open_pos=args.open_pos,
        vendor_master=args.vendor_master,
        erp=args.erp,
    )
    output = Path(args.output)
    output.mkdir(parents=True, exist_ok=True)
    cache_dir = Path(args.cache_dir)
    cache_dir.mkdir(parents=True, exist_ok=True)

    stats.inputs = {
        "open_pos": display_path(discovered.open_pos) or str(discovered.open_pos),
        "vendor_master": display_path(discovered.vendor_master) or str(discovered.vendor_master),
        "erp": display_path(discovered.erp) or str(discovered.erp),
        "confirmations_dir": display_path(Path(discovered.confirmations[0]).parent)
        or str(Path(discovered.confirmations[0]).parent),
        "pdf_count": str(len(discovered.confirmations)),
        "skipped_pdfs": ", ".join(p.name for p in discovered.skipped_pdfs) or "(none)",
    }
    logger.info("Open POs:        %s", discovered.open_pos)
    logger.info("Vendor master:   %s", discovered.vendor_master)
    logger.info("ERP database:    %s", discovered.erp)
    logger.info("PDFs discovered: %s (skipped %s)", len(discovered.confirmations), len(discovered.skipped_pdfs))
    env = config._ENV_STATUS
    logger.info(
        "LLM config: provider=%s model=%s .env_exists=%s key_present=%s key_length=%s",
        stats.llm_provider,
        stats.llm_model,
        env.get("env_file_exists"),
        env.get("openai_key_present"),
        env.get("openai_key_length"),
    )
    if not args.offline and llm_available():
        ping = ping_openai()
        if not ping.get("ok"):
            logger.error(
                "OpenAI authentication failed before extraction (%s). "
                "Refusing to extract so local-parser fallback cannot poison the API cache. "
                "Fix OPENAI_API_KEY in .env or rerun with --offline.",
                ping.get("error") or "ping unsuccessful",
            )
            return 1

    open_pos = load_open_pos(discovered.open_pos)
    vendors = load_vendor_master(discovered.vendor_master)
    logger.info("Loaded %s open PO lines across %s POs", len(open_pos), len({r.po_number for r in open_pos}))

    conn, schema, erp_data = load_erp(discovered.erp)
    try:
        fx = FxTable(erp_data.get("fx_rate") or [])
        historical = build_historical_crosswalk(
            erp_data["confirmation"],
            erp_data["po_line"],
            erp_data["po_header"],
            erp_data["vendor_master"],
        )
        logger.info("Historical crosswalk entries: %s (approved %s)", len(historical), sum(1 for e in historical if e.status.value == "APPROVED"))

        logger.info("Extracting %s PDFs (offline=%s refresh=%s) ...", len(discovered.confirmations), args.offline, args.refresh_cache)
        try:
            documents = extract_all(
                discovered.confirmations,
                cache_dir=cache_dir,
                refresh_cache=args.refresh_cache,
                offline=args.offline,
                logger_=logger,
            )
        except AuthenticationFailed as exc:
            logger.error(
                "OpenAI authentication failed during extraction (%s). Cache was not updated.",
                exc,
            )
            return 1
        resolved = resolve_active_documents(documents)
        logger.info(
            "Documents active=%s superseded=%s informational=%s review=%s",
            len(resolved.active),
            len(resolved.superseded),
            len(resolved.informational),
            len(resolved.review),
        )

        rows, _matches, proposed = reconcile(
            open_pos,
            vendors,
            resolved,
            crosswalk=historical,
            fx=fx,
        )
        crosswalk = historical + proposed

        perf = analyze_vendor_performance(erp_data, crosswalk=crosswalk)
        tally_stats(documents, rows, stats, review_count=len(resolved.review))
        for doc in documents:
            doc.source_file = display_path(doc.source_file) or doc.source_file
        for row in rows:
            if row.source_file:
                row.source_file = display_path(row.source_file) or row.source_file
            row.extra_source_files = [display_path(p) or p for p in row.extra_source_files]
        follow_ups = build_follow_ups(rows)

        out_task1 = output / "lisa_reconciliation.xlsx"
        out_task2 = output / "vendor_performance.xlsx"
        out_jsonl = output / "extracted_documents.jsonl"
        out_csv = output / "reconciled_po_lines.csv"
        out_vm = output / "vendor_metrics.csv"
        out_md = output / "run_summary.md"
        out_fu = output / "follow_up_drafts.xlsx"
        out_demo = output / "demo_notes.md"

        write_task1_workbook(
            out_task1,
            rows=rows,
            documents=documents,
            crosswalk=crosswalk,
            stats=stats,
            follow_ups=follow_ups,
        )
        write_task2_workbook(out_task2, perf, crosswalk)
        write_extracted_jsonl(out_jsonl, documents)
        write_reconciled_csv(out_csv, rows)
        write_vendor_metrics_csv(out_vm, perf)
        write_follow_ups_xlsx(out_fu, follow_ups)
        generated = [out_task1, out_task2, out_jsonl, out_csv, out_vm, out_md, out_fu, out_demo]
        stats.generated_files = [display_path(p) or str(p) for p in generated]
        write_run_summary_md(out_md, stats=stats, rows=rows, result=perf, generated=stats.generated_files)
        write_demo_notes(out_demo, perf, rows)

        stats.finished_at = datetime.now(timezone.utc)
        logger.info("")
        logger.info("Processed: %s", stats.pdfs_processed)
        logger.info("Successful: %s", stats.pdfs_processed - stats.extraction_failures)
        logger.info("OpenAI-backed: %s (fresh this run %s, cached %s)", stats.openai_backed, stats.ai_extractions, stats.cached_extractions)
        logger.info("Local parser/OCR fallback: %s", stats.local_fallback)
        logger.info("Review required: %s", stats.review_required)
        logger.info("Vision/OCR scans: %s", stats.scanned_pdfs)
        logger.info("Warnings: %s", stats.warnings)
        logger.info("RED issues: %s  YELLOW: %s  missing lines: %s  unknown POs: %s", stats.red_issues, stats.yellow_issues, stats.missing_lines, stats.unknown_pos)
        if perf.call_first:
            logger.info("Call first: %s (required-date OTD %.1f%%)", perf.call_first.vendor_name, 100 * (perf.call_first.otd_required or 0))
        logger.info("Output: %s", display_path(output) or str(output))
        return extraction_exit_code(stats.extraction_failures)
    finally:
        conn.close()


def extraction_exit_code(failures: int) -> int:
    return 0 if failures == 0 else 1


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    logger = setup_logging()
    if args.ping_llm:
        logger.info(
            "dotenv: env_file_exists=%s openai_key_present=%s openai_key_length=%s",
            config._ENV_STATUS.get("env_file_exists"),
            config._ENV_STATUS.get("openai_key_present"),
            config._ENV_STATUS.get("openai_key_length"),
        )
        try:
            result = ping_openai()
        except Exception as exc:  # noqa: BLE001
            logger.error("OpenAI ping failed: %s", exc)
            return 1
        if not result.get("ok"):
            logger.error("OpenAI ping unsuccessful: %s", result.get("error"))
            return 1
        logger.info(
            "OpenAI authentication OK (model=%s finish=%s reply_chars=%s)",
            result.get("resolved_model") or result.get("model"),
            result.get("finish_reason"),
            result.get("reply_chars"),
        )
        return 0
    try:
        return run(args, logger)
    except FileNotFoundError as exc:
        logger.error("Fatal: %s", exc)
        return 2
    except Exception as exc:  # noqa: BLE001
        logger.exception("Fatal: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
