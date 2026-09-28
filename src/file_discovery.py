"""Discover open-PO CSV, vendor master, ERP database, and confirmation PDFs."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

SKIP_PDF_NAME_TOKENS = (
    "candidate",
    "prompt",
    "instruction",
    "case-study",
    "casestudy",
    "beacon_erp",
    "erp_export",
    "database",
)


@dataclass
class DiscoveredInputs:
    open_pos: Path | None = None
    vendor_master: Path | None = None
    erp: Path | None = None
    confirmations: list[Path] = field(default_factory=list)
    skipped_pdfs: list[Path] = field(default_factory=list)

    def require(self) -> None:
        missing = []
        if self.open_pos is None or not self.open_pos.exists():
            missing.append("open PO CSV (*open_pos*.csv)")
        if self.vendor_master is None or not self.vendor_master.exists():
            missing.append("vendor master CSV (*vendor_master*.csv)")
        if self.erp is None or not self.erp.exists():
            missing.append("ERP SQLite database (*beacon_erp*.db)")
        if missing:
            raise FileNotFoundError(
                "Required input(s) not found: "
                + ", ".join(missing)
                + ". Pass --data-dir or explicit --open-pos/--vendor-master/--erp paths."
            )
        if not self.confirmations:
            raise FileNotFoundError(
                "No vendor confirmation PDFs were discovered. "
                "Pass --confirmations DIR or place PDFs under --data-dir."
            )


def _pick_best(matches: list[Path], preferred_names: tuple[str, ...]) -> Path | None:
    if not matches:
        return None
    exact = [p for p in matches if p.name.lower() in preferred_names]
    if exact:
        return sorted(exact)[0]
    return sorted(matches, key=lambda p: (len(p.name), p.name.lower()))[0]


def display_path(path: str | Path | None) -> str | None:
    """Prefer a cwd-relative path so audit files do not embed machine-absolute Cursor paths."""
    if path is None or str(path).strip() == "":
        return None
    p = Path(path)
    try:
        resolved = p.resolve()
        rel = resolved.relative_to(Path.cwd().resolve())
        return rel.as_posix()
    except (ValueError, OSError):
        return p.name


def is_vendor_confirmation_pdf(path: Path) -> bool:
    name = path.name.lower()
    if path.suffix.lower() != ".pdf":
        return False
    return not any(tok in name for tok in SKIP_PDF_NAME_TOKENS)


def discover_pdfs(root: Path) -> tuple[list[Path], list[Path]]:
    pdfs = sorted(p for p in root.rglob("*.pdf") if p.is_file())
    keep, skip = [], []
    for p in pdfs:
        if is_vendor_confirmation_pdf(p):
            keep.append(p)
        else:
            skip.append(p)
    return keep, skip


def discover_from_data_dir(data_dir: Path) -> DiscoveredInputs:
    data_dir = data_dir.resolve()
    if not data_dir.exists():
        raise FileNotFoundError(f"data dir does not exist: {data_dir}")

    csvs = list(data_dir.rglob("*.csv"))
    dbs = list(data_dir.rglob("*.db")) + list(data_dir.rglob("*.sqlite"))

    open_pos = _pick_best(
        [p for p in csvs if "open_pos" in p.name.lower() or "open-po" in p.name.lower()],
        ("open_pos.csv",),
    )
    vendor_master = _pick_best(
        [p for p in csvs if "vendor_master" in p.name.lower() or "vendormaster" in p.name.lower()],
        ("vendor_master.csv",),
    )
    erp = _pick_best(
        [p for p in dbs if "beacon_erp" in p.name.lower() or "erp" in p.name.lower()],
        ("beacon_erp.db",),
    )
    pdfs, skipped = discover_pdfs(data_dir)
    return DiscoveredInputs(
        open_pos=open_pos,
        vendor_master=vendor_master,
        erp=erp,
        confirmations=pdfs,
        skipped_pdfs=skipped,
    )


def resolve_inputs(
    *,
    data_dir: Path | None,
    confirmations: Path | None,
    open_pos: Path | None,
    vendor_master: Path | None,
    erp: Path | None,
) -> DiscoveredInputs:
    discovered = DiscoveredInputs()
    if data_dir is not None:
        discovered = discover_from_data_dir(data_dir)

    if open_pos is not None:
        discovered.open_pos = Path(open_pos)
    if vendor_master is not None:
        discovered.vendor_master = Path(vendor_master)
    if erp is not None:
        discovered.erp = Path(erp)

    if confirmations is not None:
        cpath = Path(confirmations)
        if cpath.is_file() and cpath.suffix.lower() == ".pdf":
            discovered.confirmations = [cpath]
        elif cpath.is_dir():
            pdfs, skipped = discover_pdfs(cpath)
            discovered.confirmations = pdfs
            discovered.skipped_pdfs = skipped
        else:
            raise FileNotFoundError(f"confirmations path not found: {cpath}")

    discovered.require()
    return discovered
