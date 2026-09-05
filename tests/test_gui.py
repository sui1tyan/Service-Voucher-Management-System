from __future__ import annotations

from typing import Any

from gui import VoucherDialog


class _Field:
    def __init__(self, value: Any) -> None:
        self.value = str(value)

    def get(self, *_args: Any) -> str:
        return self.value


class _VoucherDialogStub:
    def __init__(self, payload: dict[str, Any], voucher_id: str) -> None:
        self.voucher: dict[str, Any] = {}
        self.entries = {
            field: _Field(payload.get(field, ""))
            for field, _label in VoucherDialog.ENTRY_FIELDS
        }
        self.textboxes = {
            field: _Field(payload.get(field, ""))
            for field in ("particulars", "problem", "solution")
        }
        self.recipient = _Field(payload.get("recipient", ""))
        self.status = _Field(payload.get("status", "Pending"))
        self.voucher_id_entry = _Field(voucher_id)
        self.result: dict[str, Any] | None = None
        self.destroyed = False

    def destroy(self) -> None:
        self.destroyed = True


def test_voucher_dialog_preserves_edited_id(voucher_payload: dict[str, Any]) -> None:
    dialog = _VoucherDialogStub(voucher_payload, "0042000")

    VoucherDialog._save(dialog)  # type: ignore[arg-type]

    assert dialog.result is not None
    assert dialog.result["voucher_id"] == "42000"
    assert dialog.destroyed


def test_voucher_dialog_allows_unchanged_legacy_id(
    voucher_payload: dict[str, Any],
) -> None:
    dialog = _VoucherDialogStub(voucher_payload, "SV-LEGACY")
    dialog.voucher = {"voucher_id": "SV-LEGACY"}

    VoucherDialog._save(dialog)  # type: ignore[arg-type]

    assert dialog.result is not None
    assert dialog.result["voucher_id"] == "SV-LEGACY"
