"""LLM provider adapters for structured document extraction.

AI reads documents. Python makes business decisions.
The complete open-PO list is NEVER sent to the model.
"""

from __future__ import annotations

import json
import logging
from typing import Any, Protocol

from tenacity import retry, retry_if_exception_type, stop_after_attempt, wait_exponential

from src import config
from src.config import EXTRACTION_SYSTEM_PROMPT

logger = logging.getLogger("beacon")

EXTRACTION_JSON_SCHEMA: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "vendor_name_raw": {"type": ["string", "null"]},
        "po_number": {"type": ["string", "null"]},
        "document_number": {"type": ["string", "null"]},
        "document_date": {"type": ["string", "null"], "description": "As printed, any format"},
        "document_type": {
            "type": "string",
            "enum": [
                "ACKNOWLEDGEMENT",
                "REVISION",
                "INVOICE",
                "SHIPPING_NOTICE",
                "RECEIPT_ONLY",
                "OTHER",
                "UNKNOWN",
            ],
        },
        "commitment_status": {
            "type": "string",
            "enum": ["FULL", "PARTIAL", "RECEIPT_ACKNOWLEDGED_NO_SCHEDULE", "UNKNOWN"],
        },
        "is_revision": {"type": "boolean"},
        "supersedes_all_prior_for_po": {"type": "boolean"},
        "partial_sequence": {"type": ["integer", "null"]},
        "partial_total": {"type": ["integer", "null"]},
        "document_currency": {"type": ["string", "null"]},
        "free_text_notes": {"type": "array", "items": {"type": "string"}},
        "extraction_warnings": {"type": "array", "items": {"type": "string"}},
        "lines": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "source_line_number": {"type": ["integer", "null"]},
                    "customer_part_number": {"type": ["string", "null"]},
                    "vendor_part_number": {"type": ["string", "null"]},
                    "description": {"type": ["string", "null"]},
                    "quantity": {"type": ["number", "null"]},
                    "uom": {"type": ["string", "null"]},
                    "unit_price": {"type": ["number", "string", "null"]},
                    "currency": {"type": ["string", "null"]},
                    "promise_raw_text": {"type": ["string", "null"]},
                    "date_type": {
                        "type": "string",
                        "enum": [
                            "PROMISE_DATE",
                            "DELIVERY_DATE",
                            "SHIP_DATE",
                            "DELIVERY_WINDOW",
                            "UNKNOWN",
                        ],
                    },
                    "page_number": {"type": ["integer", "null"]},
                },
                "required": [
                    "source_line_number",
                    "customer_part_number",
                    "vendor_part_number",
                    "description",
                    "quantity",
                    "uom",
                    "unit_price",
                    "currency",
                    "promise_raw_text",
                    "date_type",
                    "page_number",
                ],
            },
        },
    },
    "required": [
        "vendor_name_raw",
        "po_number",
        "document_number",
        "document_date",
        "document_type",
        "commitment_status",
        "is_revision",
        "supersedes_all_prior_for_po",
        "partial_sequence",
        "partial_total",
        "document_currency",
        "free_text_notes",
        "extraction_warnings",
        "lines",
    ],
}


class ExtractionError(Exception):
    """Raised when the model response cannot be repaired."""


class LLMClient(Protocol):
    provider_name: str
    model_name: str

    def extract_from_text(self, text: str, *, retry_hint: str | None = None) -> dict:
        ...

    def extract_from_images(self, images_png: list[bytes], *, retry_hint: str | None = None) -> dict:
        ...


def llm_available() -> bool:
    provider = config.LLM_PROVIDER
    if provider == "local":
        return False
    if provider == "anthropic":
        return bool(config.ANTHROPIC_API_KEY)
    return bool(config.OPENAI_API_KEY)


def ping_openai() -> dict:
    """Tiny authenticated request. Does not echo the API key."""
    from openai import OpenAI

    status = {
        "env_file_exists": bool(config._ENV_STATUS.get("env_file_exists")),
        "key_present": bool(config.OPENAI_API_KEY),
        "key_length": len(config.OPENAI_API_KEY or ""),
        "model": config.LLM_MODEL or "gpt-4o-mini",
        "ok": False,
    }
    if not config.OPENAI_API_KEY:
        status["error"] = "OPENAI_API_KEY is empty after python-dotenv load"
        return status
    client = OpenAI(api_key=config.OPENAI_API_KEY, timeout=min(config.LLM_TIMEOUT_SECONDS, 30))
    response = client.chat.completions.create(
        model=status["model"],
        messages=[{"role": "user", "content": "Reply with the single word pong."}],
        max_tokens=8,
        temperature=0,
    )
    text = (response.choices[0].message.content or "").strip()
    status.update(
        {
            "ok": True,
            "response_id_present": bool(getattr(response, "id", None)),
            "finish_reason": response.choices[0].finish_reason,
            "reply_chars": len(text),
            "resolved_model": getattr(response, "model", None),
        }
    )
    return status


def build_llm_client() -> LLMClient | None:
    if not llm_available():
        return None
    if config.LLM_PROVIDER == "anthropic":
        return AnthropicClient(model=config.LLM_MODEL or "claude-sonnet-4-5")
    return OpenAIClient(model=config.LLM_MODEL or "gpt-4o-mini")


def _user_payload(kind: str, body: str, retry_hint: str | None) -> str:
    extra = ""
    if retry_hint:
        extra = (
            "\n\nThe previous response failed validation. "
            f"Fix these issues and return valid JSON only:\n{retry_hint}"
        )
    return (
        f"Extract procurement facts from this {kind}. "
        "Do not invent omitted lines or prices.\n\n"
        f"{body}{extra}"
    )


class OpenAIClient:
    provider_name = "openai"

    def __init__(self, model: str):
        from openai import OpenAI

        self.model_name = model
        self._client = OpenAI(api_key=config.OPENAI_API_KEY, timeout=config.LLM_TIMEOUT_SECONDS)

    def extract_from_text(self, text: str, *, retry_hint: str | None = None) -> dict:
        return self._complete(
            [
                {"role": "system", "content": EXTRACTION_SYSTEM_PROMPT},
                {
                    "role": "user",
                    "content": _user_payload("document text", text[:24000], retry_hint),
                },
            ]
        )

    def extract_from_images(self, images_png: list[bytes], *, retry_hint: str | None = None) -> dict:
        import base64

        content: list[dict] = [
            {
                "type": "text",
                "text": _user_payload("scanned document image(s)", "", retry_hint),
            }
        ]
        for img in images_png:
            b64 = base64.b64encode(img).decode("ascii")
            content.append(
                {
                    "type": "image_url",
                    "image_url": {"url": f"data:image/png;base64,{b64}"},
                }
            )
        return self._complete(
            [
                {"role": "system", "content": EXTRACTION_SYSTEM_PROMPT},
                {"role": "user", "content": content},
            ]
        )

    @retry(
        reraise=True,
        stop=stop_after_attempt(2),
        wait=wait_exponential(multiplier=1, min=1, max=8),
        retry=retry_if_exception_type(Exception),
    )
    def _complete(self, messages: list[dict]) -> dict:
        response = self._client.chat.completions.create(
            model=self.model_name,
            messages=messages,
            temperature=0,
            response_format={
                "type": "json_schema",
                "json_schema": {
                    "name": "extracted_confirmation",
                    "strict": True,
                    "schema": EXTRACTION_JSON_SCHEMA,
                },
            },
        )
        content = response.choices[0].message.content
        if not content:
            raise ExtractionError("Empty model response")
        return json.loads(content)


class AnthropicClient:
    provider_name = "anthropic"

    def __init__(self, model: str):
        import anthropic

        self.model_name = model
        self._client = anthropic.Anthropic(
            api_key=config.ANTHROPIC_API_KEY,
            timeout=config.LLM_TIMEOUT_SECONDS,
        )

    def extract_from_text(self, text: str, *, retry_hint: str | None = None) -> dict:
        return self._complete(
            [{"role": "user", "content": _user_payload("document text", text[:24000], retry_hint)}]
        )

    def extract_from_images(self, images_png: list[bytes], *, retry_hint: str | None = None) -> dict:
        import base64

        blocks: list[dict] = [
            {
                "type": "text",
                "text": _user_payload("scanned document image(s)", "", retry_hint),
            }
        ]
        for img in images_png:
            blocks.append(
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/png",
                        "data": base64.b64encode(img).decode("ascii"),
                    },
                }
            )
        return self._complete([{"role": "user", "content": blocks}])

    @retry(
        reraise=True,
        stop=stop_after_attempt(2),
        wait=wait_exponential(multiplier=1, min=1, max=8),
        retry=retry_if_exception_type(Exception),
    )
    def _complete(self, messages: list) -> dict:
        response = self._client.messages.create(
            model=self.model_name,
            max_tokens=4096,
            system=EXTRACTION_SYSTEM_PROMPT,
            messages=messages,
            temperature=0,
        )
        text = "".join(block.text for block in response.content if getattr(block, "type", "") == "text")
        text = text.strip()
        if text.startswith("```"):
            text = text.strip("`")
            if text.startswith("json"):
                text = text[4:]
            text = text.strip()
        if not text:
            raise ExtractionError("Empty Anthropic response")
        return json.loads(text)
