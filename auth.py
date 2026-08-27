"""Password hashing and account input validation."""

from __future__ import annotations

import re

import bcrypt


USERNAME_RE = re.compile(r"^[A-Za-z0-9._-]{3,40}$")


def hash_pwd(password: str) -> bytes:
    """Hash a password with a per-password bcrypt salt."""

    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt(rounds=12))


def verify_pwd(password: str, password_hash: bytes | str | None) -> bool:
    """Return False for malformed hashes instead of crashing the login screen."""

    if not password or not password_hash:
        return False
    try:
        if isinstance(password_hash, str):
            password_hash = password_hash.encode("utf-8")
        return bcrypt.checkpw(password.encode("utf-8"), password_hash)
    except (TypeError, ValueError):
        return False


def validate_password_policy(password: str) -> str | None:
    """Return a user-facing policy error, or None when the password is valid."""

    if not password:
        return "Password cannot be empty."
    if len(password) < 12:
        return "Password must be at least 12 characters."
    if len(password.encode("utf-8")) > 72:
        return "Password must not exceed 72 UTF-8 bytes."
    if not re.search(r"[A-Z]", password):
        return "Include at least one uppercase letter."
    if not re.search(r"[a-z]", password):
        return "Include at least one lowercase letter."
    if not re.search(r"\d", password):
        return "Include at least one digit."
    if not re.search(r"[^\w\s]", password):
        return "Include at least one symbol."
    return None


def validate_username(username: str) -> str | None:
    """Validate a username used as a stable login identifier."""

    if not username:
        return "Username cannot be empty."
    if not USERNAME_RE.fullmatch(username):
        return "Use 3-40 letters, numbers, dots, underscores, or hyphens."
    return None
