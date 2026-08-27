from __future__ import annotations

from auth import (
    hash_pwd,
    validate_password_policy,
    validate_username,
    verify_pwd,
)


def test_password_hash_and_verification() -> None:
    password_hash = hash_pwd("StrongPassword!2026")
    assert password_hash != b"StrongPassword!2026"
    assert verify_pwd("StrongPassword!2026", password_hash)
    assert not verify_pwd("wrong-password", password_hash)
    assert not verify_pwd("anything", b"invalid-hash")


def test_password_policy() -> None:
    assert validate_password_policy("Short1!") is not None
    assert validate_password_policy("alllowercase!2026") is not None
    assert validate_password_policy("ALLUPPERCASE!2026") is not None
    assert validate_password_policy("NoDigitsHere!!") is not None
    assert validate_password_policy("NoSymbolsHere2026") is not None
    assert validate_password_policy("ValidPassword!2026") is None
    assert validate_password_policy("A1!" + ("界" * 30)) is not None


def test_username_policy() -> None:
    assert validate_username("owner") is None
    assert validate_username("service.admin-01") is None
    assert validate_username("x") is not None
    assert validate_username("invalid username") is not None
