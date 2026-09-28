"""Run-level extraction metrics and process exit code."""

from src.file_discovery import display_path
from src.models import ExtractedDocument, ExtractionMode, ReconciledLine, RunStats, Severity
from main import extraction_exit_code, tally_stats


def test_extraction_exit_code_is_nonzero_on_failures():
    assert extraction_exit_code(0) == 0
    assert extraction_exit_code(1) == 1
    assert extraction_exit_code(12) == 1


def test_cached_openai_documents_count_as_openai_backed():
    docs = [
        ExtractedDocument(
            source_file="data/confirmations/apex_01.pdf",
            source_sha256="a",
            extraction_mode=ExtractionMode.NATIVE_TEXT,
            cached=True,
        ),
        ExtractedDocument(
            source_file="data/confirmations/continental_01.pdf",
            source_sha256="b",
            extraction_mode=ExtractionMode.VISION,
            cached=True,
        ),
        ExtractedDocument(
            source_file="data/confirmations/local.pdf",
            source_sha256="c",
            extraction_mode=ExtractionMode.LOCAL_TEXT,
            cached=False,
        ),
    ]
    stats = RunStats()
    tally_stats(docs, [], stats)
    assert stats.pdfs_processed == 3
    assert stats.openai_backed == 2
    assert stats.cached_extractions == 2
    assert stats.ai_extractions == 0
    assert stats.local_fallback == 1
    assert stats.native_text_openai == 1
    assert stats.vision_openai == 1


def test_fresh_openai_extractions_are_counted_separately_from_cache():
    docs = [
        ExtractedDocument(
            source_file="a.pdf",
            source_sha256="a",
            extraction_mode=ExtractionMode.NATIVE_TEXT,
            cached=False,
        ),
        ExtractedDocument(
            source_file="b.pdf",
            source_sha256="b",
            extraction_mode=ExtractionMode.VISION,
            cached=True,
        ),
    ]
    stats = RunStats()
    tally_stats(docs, [], stats)
    assert stats.openai_backed == 2
    assert stats.ai_extractions == 1
    assert stats.cached_extractions == 1
    assert stats.local_fallback == 0


def test_display_path_prefers_relative_source_names(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    nested = tmp_path / "data" / "confirmations"
    nested.mkdir(parents=True)
    pdf = nested / "apex_01.pdf"
    pdf.write_bytes(b"%PDF")
    assert display_path(pdf) == "data/confirmations/apex_01.pdf"
    assert display_path("/tmp/not-in-this-workspace/other.pdf") == "other.pdf"


def test_tally_red_yellow_counts():
    rows = [
        ReconciledLine(po_number="PO-1", severity=Severity.RED, issue_codes=["MISSING_PO_LINE"]),
        ReconciledLine(po_number="PO-2", severity=Severity.YELLOW, issue_codes=["DATE_SEMANTICS_WARNING"]),
        ReconciledLine(po_number="PO-3", severity=Severity.GREEN),
    ]
    stats = RunStats()
    tally_stats([], rows, stats)
    assert stats.red_issues == 1
    assert stats.yellow_issues == 1
    assert stats.missing_lines == 1
