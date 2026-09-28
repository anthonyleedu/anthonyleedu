"""Read-only ERP access, schema introspection, and receipt completion dates."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Any

from src.normalize import parse_date


REQUIRED_TABLES = {
    "confirmation": {
        "conf_id",
        "po_number",
        "line_no",
        "vendor_pn",
        "part_id",
        "confirmed_qty",
        "confirmed_price",
        "currency",
        "promised_date",
        "doc_date",
        "source_file",
        "superseded_by",
    },
    "fx_rate": {"month", "currency", "rate_to_usd"},
    "po_header": {"po_number", "vendor_id", "po_date", "buyer", "status"},
    "po_line": {
        "po_number",
        "line_no",
        "part_id",
        "qty_ordered",
        "unit_price",
        "required_date",
        "qty_received",
    },
    "receipt_txn": {"txn_id", "po_number", "line_no", "action_type", "qty", "txn_date"},
    "vendor_master": {"vendor_id", "vendor_name", "country", "currency"},
    "part_master": {"part_id", "description", "part_class", "uom", "std_cost"},
    "qc_hold": {"hold_id", "po_number", "reason_code", "hold_date", "released"},
}


class ErpSchemaError(RuntimeError):
    pass


@dataclass
class TableSchema:
    name: str
    columns: list[dict]


@dataclass
class ErpSchema:
    tables: list[str]
    table_info: dict[str, TableSchema] = field(default_factory=dict)


def open_readonly(db_path: Path) -> sqlite3.Connection:
    path = Path(db_path).resolve()
    if not path.exists():
        raise FileNotFoundError(f"ERP database not found: {path}")
    conn = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    return conn


def introspect_schema(conn: sqlite3.Connection) -> ErpSchema:
    tables = [
        r[0]
        for r in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY 1"
        )
    ]
    info: dict[str, TableSchema] = {}
    for t in tables:
        cols = [dict(r) for r in conn.execute(f"PRAGMA table_info({t})")]
        info[t] = TableSchema(name=t, columns=cols)
    return ErpSchema(tables=tables, table_info=info)


def validate_schema(schema: ErpSchema) -> None:
    problems: list[str] = []
    for table, required_cols in REQUIRED_TABLES.items():
        if table not in schema.table_info:
            problems.append(f"missing required table '{table}'")
            continue
        present = {c["name"] for c in schema.table_info[table].columns}
        missing = sorted(required_cols - present)
        if missing:
            problems.append(f"table '{table}' missing columns: {', '.join(missing)}")
    if problems:
        raise ErpSchemaError(
            "ERP schema validation failed:\n  - "
            + "\n  - ".join(problems)
            + "\nOpen the SQLite file and confirm it is the Beacon extract."
        )


def fetch_table(conn: sqlite3.Connection, name: str) -> list[dict[str, Any]]:
    return [dict(r) for r in conn.execute(f"SELECT * FROM {name}")]


def load_erp(db_path: Path) -> tuple[sqlite3.Connection, ErpSchema, dict[str, list[dict]]]:
    conn = open_readonly(db_path)
    schema = introspect_schema(conn)
    validate_schema(schema)
    data = {t: fetch_table(conn, t) for t in schema.tables}
    return conn, schema, data


def stable_completion_date(
    receipts: list[dict],
    qty_ordered: Decimal | float | int,
) -> date | None:
    """Earliest txn_date where cumulative signed qty >= ordered AND never falls below afterward.

    Receipts must already be sorted by (txn_date, txn_id).
    Uses the SIGNED qty stored in the database. Do not abs() reversals.
    """
    ordered = Decimal(str(qty_ordered))
    if not receipts:
        return None
    history: list[tuple[date, Decimal, Any]] = []
    running = Decimal("0")
    for r in receipts:
        d = parse_date(r.get("txn_date"))
        if d is None:
            continue
        qty = Decimal(str(r.get("qty") or 0))
        running += qty
        history.append((d, running, r.get("txn_id")))
    if not history or history[-1][1] < ordered:
        return None
    for i, (dt, cum, _tid) in enumerate(history):
        if cum >= ordered and all(h[1] >= ordered for h in history[i:]):
            return dt
    return None


def completion_map(
    po_lines: list[dict],
    receipts: list[dict],
) -> dict[tuple[str, int], date | None]:
    by_line: dict[tuple[str, int], list[dict]] = {}
    for r in receipts:
        key = (r["po_number"], int(r["line_no"]))
        by_line.setdefault(key, []).append(r)
    for key, rows in by_line.items():
        rows.sort(key=lambda r: (parse_date(r["txn_date"]) or date.min, r.get("txn_id") or 0))

    out: dict[tuple[str, int], date | None] = {}
    for pl in po_lines:
        key = (pl["po_number"], int(pl["line_no"]))
        out[key] = stable_completion_date(by_line.get(key, []), pl["qty_ordered"])
    return out


def analysis_as_of(receipts: list[dict]) -> date | None:
    dates = [parse_date(r.get("txn_date")) for r in receipts]
    dates = [d for d in dates if d]
    return max(dates) if dates else None
