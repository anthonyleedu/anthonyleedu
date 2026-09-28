"""Structured logging helpers for extraction progress."""

from __future__ import annotations

import logging
import sys


def setup_logging(level: int = logging.INFO) -> logging.Logger:
    logger = logging.getLogger("beacon")
    if logger.handlers:
        logger.setLevel(level)
        return logger
    logger.setLevel(level)
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter("%(message)s"))
    logger.addHandler(handler)
    logger.propagate = False
    return logger


def progress(logger: logging.Logger, index: int, total: int, filename: str, *parts: str) -> None:
    suffix = " - ".join(p for p in parts if p)
    logger.info("[%s/%s] %s - %s", index, total, filename, suffix)
