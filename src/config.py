"""Central configuration and tunable thresholds.

All magic numbers live here so business rules stay explainable.
"""

from __future__ import annotations

import os
from decimal import Decimal
from pathlib import Path

from dotenv import load_dotenv

REPO_ROOT = Path(__file__).resolve().parents[1]


def load_environment() -> dict:
    """Load repo-root .env via python-dotenv. Never returns secret values."""
    env_file = REPO_ROOT / ".env"
    # Repo-root first, then cwd, without overriding a key already in the process env.
    load_dotenv(dotenv_path=env_file, override=False)
    load_dotenv(override=False)
    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    return {
        "env_file": str(env_file),
        "env_file_exists": env_file.exists(),
        "openai_key_present": bool(key),
        "openai_key_length": len(key),
    }


_ENV_STATUS = load_environment()

# --- Extraction ---
TEXT_MIN_CHARS = 80
TEXT_MIN_ALNUM_RATIO = 0.45
PROMPT_VERSION = "procurement-extract-v1"
SCHEMA_VERSION = "extracted-document-v1"
OCR_DPI = 200
VISION_DPI = 160

# --- Currency ---
BASE_CURRENCY = os.getenv("BASE_CURRENCY", "USD").upper()
SAME_CURRENCY_PRICE_PCT_TOL = Decimal("0.001")  # 0.1%
FOREIGN_CURRENCY_PRICE_PCT_TOL = Decimal("0.005")  # 0.5%
PRICE_ABS_TOL = Decimal("0.005")
FX_LOOKBACK_MONTHS = 2
MATERIAL_PRICE_INCREASE_PCT = Decimal("0.015")  # 1.5% -> RED
MATERIAL_PRICE_IMPACT = Decimal("25")  # $ impact that upgrades to RED

# --- Matching ---
RULE_MATCH_AUTO_THRESHOLD = 0.70
RULE_MATCH_MARGIN = 0.20
DESCRIPTION_SCORE_MIN = 0.72
PLACEHOLDER_DESCRIPTIONS = (
    "(see po for description)",
    "(siehe bestellung)",
    "see po for description",
    "siehe bestellung",
)

# --- Historical crosswalk ---
HISTORICAL_MAPPING_MIN_SUPPORT = 3
HISTORICAL_MAPPING_MIN_PURITY = 0.95

# --- Dates / lateness ---
LATE_YELLOW_MAX_DAYS = 2
QTY_OVER_YELLOW_PCT = Decimal("0.02")  # unused: overage defaults RED
QTY_OVER_DEFAULT_RED = True

# --- Vendor call-first ---
MIN_DUE_LINES_FOR_CALL = 20

# --- LLM ---
LLM_PROVIDER = os.getenv("LLM_PROVIDER", "openai").strip().lower()
LLM_MODEL = os.getenv("LLM_MODEL", "gpt-4o-mini").strip()
OPENAI_API_KEY = (os.getenv("OPENAI_API_KEY") or "").strip()
ANTHROPIC_API_KEY = (os.getenv("ANTHROPIC_API_KEY") or "").strip()
LLM_TIMEOUT_SECONDS = int(os.getenv("LLM_TIMEOUT_SECONDS", "60"))
LLM_MAX_RETRIES = 3

# --- Paths ---
DEFAULT_CACHE_DIR = Path("cache/extractions")
DEFAULT_OUTPUT_DIR = Path("output")

EXTRACTION_SYSTEM_PROMPT = """You are a procurement document extraction engine.

Your only task is to transcribe and normalize facts that are actually present in the supplied document.

Do not reconcile the document to a purchase order.
Do not guess what the customer probably ordered.
Do not invent omitted lines.
Do not repair the vendor's document.
Do not infer a price from another source.

If a value is not stated, return null.

Classify what kind of document this is:
- acknowledgement/order confirmation
- revision
- invoice
- shipping notice
- receipt acknowledgement with no schedule
- other/unknown

Extract:
- vendor name
- Beacon PO number
- vendor document/order number
- document date
- document currency
- whether the document explicitly revises/supersedes an earlier document
- whether it is part N of M
- every actual line visibile in the document

For each visible line extract:
- source/vendor line number
- customer part number if explicitly shown
- vendor part number if explicitly shown
- description
- quantity
- UOM
- unit price
- currency
- promise/ship/delivery date

Preserve the meaning of dates.

If the document says "ship 05/21/2026":
    date_type must be SHIP_DATE.

If it says "Promise Date":
    date_type must be PROMISE_DATE.

If it gives a delivery window:
    preserve the full range.

If it uses an ISO calendar week such as:
    KW 20-22 / 2026
return:
    raw_text = "KW 20-22 / 2026"
and identify it as a week range.

If a supplier says schedule is still to be confirmed:
    do NOT invent a promise date.

If the document is an invoice:
    classify it as invoice even if it contains PO lines.

If the document says to disregard an earlier confirmation:
    mark it as a revision and supersedes_all_prior_for_po=true.

Return only data matching the provided JSON schema.
"""
