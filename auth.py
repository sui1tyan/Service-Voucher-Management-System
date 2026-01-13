import bcrypt
import re

def hash_pwd(pwd: str) -> bytes:
    return bcrypt.hashpw(pwd.encode("utf-8"), bcrypt.gensalt())

def verify_pwd(pwd: str, hp: bytes) -> bool:
    try:
        return bcrypt.checkpw(pwd.encode("utf-8"), hp)
    except Exception:
        return False

def validate_password_policy(pw: str) -> str | None:
    if not pw: return "Password cannot be empty."
    s = str(pw)
    if len(s) < 10: return "Password must be at least 10 characters."
    if not re.search(r"[A-Z]", s): return "Include at least one uppercase letter."
    if not re.search(r"[a-z]", s): return "Include at least one lowercase letter."
    if not re.search(r"\d", s): return "Include at least one digit."
    if not re.search(r"[^\w\s]", s): return "Include at least one symbol."
    return None
