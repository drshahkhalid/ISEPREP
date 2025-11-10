"""
Authentication & password hashing utilities.

Hash format stored in users.password_hash:
    pbkdf2_sha256$<iterations>$<base64(salt)>$<base64(hash)>

Functions:
    hash_password(plain) -> hash string
    verify_password(plain, stored_hash) -> bool
    authenticate(username, password) -> (success, user_dict|None, migrated:bool, message)

If a legacy (plaintext) password is detected (i.e., the stored value is NOT in the pbkdf2 format)
and the provided password matches exactly, the password is re-hashed and updated (migration).
"""

import sqlite3
import hashlib
import base64
import secrets
from db import connect_db

DEFAULT_ITERATIONS = 260000


def hash_password(password: str, iterations: int = DEFAULT_ITERATIONS) -> str:
    if not isinstance(password, str):
        raise TypeError("Password must be a string")
    salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations)
    return f"pbkdf2_sha256${iterations}${base64.b64encode(salt).decode()}${base64.b64encode(dk).decode()}"


def verify_password(password: str, stored_hash: str) -> bool:
    try:
        algo, iterations, salt_b64, hash_b64 = stored_hash.split("$", 3)
        if algo != "pbkdf2_sha256":
            return False
        iterations = int(iterations)
        salt = base64.b64decode(salt_b64)
        original_hash = base64.b64decode(hash_b64)
        test_hash = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations)
        return secrets.compare_digest(original_hash, test_hash)
    except Exception:
        return False


def is_pdkdf2_hash(value: str) -> bool:
    if not value or "$" not in value:
        return False
    parts = value.split("$")
    return len(parts) == 4 and parts[0] == "pbkdf2_sha256" and parts[1].isdigit()


def authenticate(username: str, password: str):
    """
    Returns (success: bool, user_row: dict|None, migrated: bool, message: str)

    Migration logic:
      - If stored password is plaintext (legacy) and matches exactly, re-hash and update row.
      - If hashing update fails, still allow login (just no migration message).
    """
    conn = connect_db()
    if conn is None:
        return False, None, False, "Database connection failed"
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    try:
        cur.execute('SELECT user_id, username, password_hash, role, preferred_language FROM users WHERE username = ?', (username,))
        row = cur.fetchone()
        if not row:
            return False, None, False, "Invalid username or password"

        stored = row["password_hash"] or ""
        migrated = False

        if is_pdkdf2_hash(stored):
            # Normal path
            if not verify_password(password, stored):
                return False, None, False, "Invalid username or password"
        else:
            # Legacy plaintext
            if stored != password:
                return False, None, False, "Invalid username or password"
            # Migrate
            new_hash = hash_password(password)
            try:
                cur.execute('UPDATE users SET password_hash=? WHERE user_id=?', (new_hash, row["user_id"]))
                conn.commit()
                migrated = True
            except sqlite3.Error:
                conn.rollback()
                # Migration failure does not block login

        user_obj = {
            "user_id": row["user_id"],
            "username": row["username"],
            "role": row["role"],
            "preferred_language": (row["preferred_language"] or "EN").upper()
        }
        return True, user_obj, migrated, "Login successful"

    finally:
        cur.close()
        conn.close()
