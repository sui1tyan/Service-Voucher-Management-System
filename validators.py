"""Validation and normalization for business records."""

from __future__ import annotations

import re
from datetime import datetime
from typing import Any, Mapping

from config import BILL_TYPES, VOUCHER_STATUSES


PHONE_RE = re.compile(r"^\+?[0-9][0-9\s().-]{3,23}[0-9]$")


class ValidationError(ValueError):
    """Raised when a business record cannot be safely persisted."""


def clean_text(value: Any, *, limit: int = 2_000) -> str:
    text = str(value or "").strip()
    if len(text) > limit:
        raise ValidationError(f"Text must not exceed {limit} characters.")
    return text


def normalize_phone(value: Any) -> str:
    phone = clean_text(value, limit=25)
    if not PHONE_RE.fullmatch(phone):
        raise ValidationError("Enter a valid contact number using 5-20 digits.")
    digits = re.sub(r"\D", "", phone)
    if not 5 <= len(digits) <= 20:
        raise ValidationError("Enter a valid contact number using 5-20 digits.")
    return ("+" if phone.startswith("+") else "") + digits


def normalize_date(value: Any, *, optional: bool = True) -> str:
    text = clean_text(value, limit=20)
    if not text and optional:
        return ""
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    raise ValidationError("Use date format YYYY-MM-DD or DD-MM-YYYY.")


def _non_negative_number(value: Any, label: str) -> float:
    if value in (None, ""):
        return 0.0
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise ValidationError(f"{label} must be a number.") from exc
    if number < 0:
        raise ValidationError(f"{label} cannot be negative.")
    return round(number, 2)


def validate_voucher(data: Mapping[str, Any]) -> dict[str, Any]:
    customer_name = clean_text(data.get("customer_name"), limit=120)
    if not customer_name:
        raise ValidationError("Customer name is required.")

    particulars = clean_text(data.get("particulars"), limit=2_000)
    problem = clean_text(data.get("problem"), limit=2_000)
    if not particulars:
        raise ValidationError("Particulars are required.")
    if not problem:
        raise ValidationError("Problem description is required.")

    try:
        units = int(data.get("units", 1))
    except (TypeError, ValueError) as exc:
        raise ValidationError("Units must be a whole number.") from exc
    if not 1 <= units <= 10_000:
        raise ValidationError("Units must be between 1 and 10,000.")

    status = clean_text(data.get("status") or "Pending", limit=30)
    if status not in VOUCHER_STATUSES:
        raise ValidationError("Select a valid voucher status.")

    return {
        "customer_name": customer_name,
        "contact_number": normalize_phone(data.get("contact_number")),
        "units": units,
        "particulars": particulars,
        "problem": problem,
        "staff_name": clean_text(data.get("staff_name"), limit=120),
        "status": status,
        "recipient": clean_text(data.get("recipient"), limit=120),
        "solution": clean_text(data.get("solution"), limit=2_000),
        "technician_id": clean_text(data.get("technician_id"), limit=60),
        "technician_name": clean_text(data.get("technician_name"), limit=120),
        "ref_bill": clean_text(data.get("ref_bill"), limit=60),
        "ref_bill_date": normalize_date(data.get("ref_bill_date"), optional=True),
        "amount_rm": _non_negative_number(data.get("amount_rm"), "Amount"),
        "tech_commission": _non_negative_number(
            data.get("tech_commission"), "Technician commission"
        ),
    }


def validate_staff(data: Mapping[str, Any]) -> dict[str, str]:
    name = clean_text(data.get("name"), limit=120)
    if not name:
        raise ValidationError("Staff name is required.")
    phone = clean_text(data.get("phone"), limit=25)
    if phone:
        phone = normalize_phone(phone)
    return {
        "position": clean_text(data.get("position"), limit=80),
        "staff_id_opt": clean_text(data.get("staff_id_opt"), limit=60),
        "name": name,
        "phone": phone,
        "email": clean_text(data.get("email"), limit=160),
        "note": clean_text(data.get("note"), limit=1_000),
        "photo_path": clean_text(data.get("photo_path"), limit=500),
    }


def validate_commission(data: Mapping[str, Any]) -> dict[str, Any]:
    bill_type = clean_text(data.get("bill_type"), limit=10).upper()
    if bill_type not in BILL_TYPES:
        raise ValidationError("Bill type must be CS or INV.")
    bill_no = clean_text(data.get("bill_no"), limit=60)
    if not bill_no:
        raise ValidationError("Bill number is required.")
    total = _non_negative_number(data.get("total_amount"), "Total amount")
    commission = _non_negative_number(data.get("commission_amount"), "Commission")
    if commission > total:
        raise ValidationError("Commission cannot exceed the total amount.")
    try:
        staff_id = int(data.get("staff_id"))
    except (TypeError, ValueError) as exc:
        raise ValidationError("Select a staff member.") from exc
    return {
        "staff_id": staff_id,
        "bill_type": bill_type,
        "bill_no": bill_no,
        "total_amount": total,
        "commission_amount": commission,
        "bill_image_path": clean_text(data.get("bill_image_path"), limit=500),
        "voucher_id": clean_text(data.get("voucher_id"), limit=60),
        "note": clean_text(data.get("note"), limit=1_000),
    }
