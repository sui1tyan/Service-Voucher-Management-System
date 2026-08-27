"""CSV exports designed for Excel and audit-friendly handoff."""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Any, Iterable, Mapping


VOUCHER_EXPORT_FIELDS = (
    "voucher_id",
    "created_at",
    "customer_name",
    "contact_number",
    "units",
    "particulars",
    "problem",
    "recipient",
    "technician_id",
    "technician_name",
    "status",
    "solution",
    "ref_bill",
    "ref_bill_date",
    "amount_rm",
    "tech_commission",
    "created_by",
    "updated_at",
    "updated_by",
)

COMMISSION_EXPORT_FIELDS = (
    "id",
    "created_at",
    "staff_name",
    "bill_type",
    "bill_no",
    "voucher_id",
    "total_amount",
    "commission_amount",
    "note",
)


def _write_csv(
    destination: str | Path,
    rows: Iterable[Mapping[str, Any]],
    fieldnames: tuple[str, ...],
) -> str:
    path = Path(destination).expanduser()
    if path.suffix.lower() != ".csv":
        path = path.with_suffix(".csv")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in fieldnames})
    return str(path.resolve())


def export_vouchers_csv(
    destination: str | Path, rows: Iterable[Mapping[str, Any]]
) -> str:
    return _write_csv(destination, rows, VOUCHER_EXPORT_FIELDS)


def export_commissions_csv(
    destination: str | Path, rows: Iterable[Mapping[str, Any]]
) -> str:
    return _write_csv(destination, rows, COMMISSION_EXPORT_FIELDS)
