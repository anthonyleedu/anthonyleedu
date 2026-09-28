"""PDF reading, SHA-256, native text quality, rendering, and OCR."""

from __future__ import annotations

import hashlib
import io
import logging
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

import pymupdf

from src.config import OCR_DPI, TEXT_MIN_ALNUM_RATIO, TEXT_MIN_CHARS, VISION_DPI

logger = logging.getLogger("beacon")

PURCHASING_HINTS = (
    "po",
    "order",
    "customer",
    "qty",
    "quantity",
    "bestellung",
    "auftrag",
    "confirmation",
    "invoice",
    "acknowledgment",
    "acknowledgement",
    "liefer",
    "promise",
)


@dataclass
class PdfPage:
    page_number: int
    text: str


@dataclass
class PdfDocument:
    path: Path
    sha256: str
    page_count: int
    pages: list[PdfPage] = field(default_factory=list)
    native_text: str = ""
    text_usable: bool = False
    is_scan: bool = False
    alphanumeric_ratio: float = 0.0

    @property
    def filename(self) -> str:
        return self.path.name


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def alphanumeric_ratio(text: str) -> float:
    stripped = text.strip()
    if not stripped:
        return 0.0
    alnum = sum(ch.isalnum() for ch in stripped)
    return alnum / len(stripped)


def text_quality_ok(text: str) -> bool:
    stripped = text.strip()
    if len(stripped) < TEXT_MIN_CHARS:
        return False
    if alphanumeric_ratio(stripped) < TEXT_MIN_ALNUM_RATIO:
        return False
    return True


def read_pdf(path: Path) -> PdfDocument:
    path = Path(path)
    digest = sha256_file(path)
    doc = pymupdf.open(path)
    try:
        pages = []
        texts = []
        for i, page in enumerate(doc, start=1):
            t = page.get_text() or ""
            pages.append(PdfPage(page_number=i, text=t))
            texts.append(t)
        native = "\n".join(texts)
        usable = text_quality_ok(native)
        return PdfDocument(
            path=path,
            sha256=digest,
            page_count=doc.page_count,
            pages=pages,
            native_text=native,
            text_usable=usable,
            is_scan=not usable,
            alphanumeric_ratio=alphanumeric_ratio(native),
        )
    finally:
        doc.close()


def render_pages_png(path: Path, dpi: int = VISION_DPI) -> list[bytes]:
    doc = pymupdf.open(path)
    images: list[bytes] = []
    try:
        zoom = dpi / 72.0
        matrix = pymupdf.Matrix(zoom, zoom)
        for page in doc:
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            images.append(pix.tobytes("png"))
    finally:
        doc.close()
    return images


def ocr_pdf(path: Path, dpi: int = OCR_DPI) -> str:
    """OCR a PDF via tesseract if available. Returns concatenated page text."""
    tess = shutil.which("tesseract")
    if not tess:
        raise RuntimeError(
            "Native PDF text is unusable and tesseract-ocr is not installed. "
            "Install tesseract or configure an LLM vision provider."
        )
    doc = pymupdf.open(path)
    parts: list[str] = []
    try:
        zoom = dpi / 72.0
        matrix = pymupdf.Matrix(zoom, zoom)
        with tempfile.TemporaryDirectory() as td:
            td_path = Path(td)
            for i, page in enumerate(doc, start=1):
                pix = page.get_pixmap(matrix=matrix, alpha=False)
                img_path = td_path / f"page-{i}.png"
                pix.save(img_path)
                result = subprocess.run(
                    [tess, str(img_path), "stdout", "--psm", "6", "-l", "eng"],
                    capture_output=True,
                    text=True,
                    check=False,
                )
                if result.returncode != 0:
                    logger.warning("tesseract failed on %s page %s: %s", path.name, i, result.stderr.strip())
                parts.append(result.stdout or "")
    finally:
        doc.close()
    return "\n".join(parts).strip()
