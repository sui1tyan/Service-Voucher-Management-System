from __future__ import annotations

import pytest

from validators import (
    ValidationError,
    normalize_date,
    normalize_phone,
    validate_commission,
    validate_voucher,
)


def test_phone_and_date_normalization() -> None:
    assert normalize_phone("+60 16-123 4567") == "+60161234567"
    assert normalize_date("27-08-2026") == "2026-08-27"
    assert normalize_date("") == ""
    with pytest.raises(ValidationError):
        normalize_phone("abc")


def test_voucher_requires_business_fields(voucher_payload: dict) -> None:
    invalid = dict(voucher_payload)
    invalid["customer_name"] = ""
    with pytest.raises(ValidationError, match="Customer name"):
        validate_voucher(invalid)


def test_commission_cannot_exceed_total() -> None:
    with pytest.raises(ValidationError, match="cannot exceed"):
        validate_commission(
            {
                "staff_id": 1,
                "bill_type": "CS",
                "bill_no": "1",
                "total_amount": 10,
                "commission_amount": 11,
            }
        )
