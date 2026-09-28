"""Load open PO CSV and vendor master CSV with schema tolerance."""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path

import pandas as pd

from src.config import BASE_CURRENCY
from src.models import OpenPOLine, VendorMasterRow
from src.normalize import normalize_part_number, normalize_po_number, parse_date, parse_number


def _col(df: pd.DataFrame, *names: str) -> str | None:
    lower = {c.lower().strip(): c for c in df.columns}
    for n in names:
        if n.lower() in lower:
            return lower[n.lower()]
    return None


def load_open_pos(path: Path) -> list[OpenPOLine]:
    df = pd.read_csv(path)
    po_c = _col(df, "po_number", "po", "purchase_order")
    date_c = _col(df, "po_date")
    vid_c = _col(df, "vendor_id")
    vname_c = _col(df, "vendor_name")
    line_c = _col(df, "line_number", "line_no", "po_line")
    pn_c = _col(df, "our_pn", "part_id", "beacon_pn")
    desc_c = _col(df, "our_description", "description")
    qty_c = _col(df, "qty_ordered", "quantity")
    price_c = _col(df, "unit_price")
    req_c = _col(df, "required_date")
    missing = [n for n, c in {
        "po_number": po_c, "line_number": line_c, "our_pn": pn_c,
        "qty_ordered": qty_c, "unit_price": price_c,
    }.items() if c is None]
    if missing:
        raise ValueError(f"open PO CSV missing required columns: {missing}. Found: {list(df.columns)}")

    rows: list[OpenPOLine] = []
    for rec in df.to_dict(orient="records"):
        pn = normalize_part_number(str(rec[pn_c]))
        po = normalize_po_number(str(rec[po_c]))
        if not po or not pn:
            continue
        rows.append(
            OpenPOLine(
                po_number=po,
                po_date=parse_date(rec.get(date_c) if date_c else None),
                vendor_id=str(rec[vid_c]).strip() if vid_c and pd.notna(rec.get(vid_c)) else None,
                vendor_name=str(rec[vname_c]).strip() if vname_c and pd.notna(rec.get(vname_c)) else None,
                line_number=int(rec[line_c]),
                our_pn=pn,
                our_description=None if desc_c is None or pd.isna(rec.get(desc_c)) else str(rec[desc_c]),
                qty_ordered=parse_number(rec[qty_c]) or Decimal("0"),
                unit_price=parse_number(rec[price_c]) or Decimal("0"),
                required_date=parse_date(rec.get(req_c) if req_c else None),
                currency_assumed=BASE_CURRENCY,
            )
        )
    return rows


def load_vendor_master(path: Path) -> list[VendorMasterRow]:
    df = pd.read_csv(path)
    vid = _col(df, "vendor_id")
    name = _col(df, "vendor_name")
    if vid is None or name is None:
        raise ValueError(f"vendor master missing vendor_id/vendor_name. Found: {list(df.columns)}")
    country = _col(df, "country")
    email = _col(df, "ap_email", "email")
    note = _col(df, "known_pn_mapping_note", "note")
    currency = _col(df, "currency")
    rows = []
    for rec in df.to_dict(orient="records"):
        rows.append(
            VendorMasterRow(
                vendor_id=str(rec[vid]).strip(),
                vendor_name=str(rec[name]).strip(),
                country=None if not country or pd.isna(rec.get(country)) else str(rec[country]),
                ap_email=None if not email or pd.isna(rec.get(email)) else str(rec[email]),
                known_pn_mapping_note=None if not note or pd.isna(rec.get(note)) else str(rec[note]),
                currency=None if not currency or pd.isna(rec.get(currency)) else str(rec[currency]).upper(),
            )
        )
    return rows
