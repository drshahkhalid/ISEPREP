"""
role_map.py

Obfuscation mapping between canonical role names and short symbols.

NOTE:
  - This provides ONLY obfuscation (it does NOT increase security).
  - Always use canonical role names ("admin", "hq", "coordinator", "manager", "supervisor")
    for authorization logic. Symbols are for display or lightweight concealment.
  - decode_role() accepts either a symbol or a canonical role name and returns
    the canonical role. If value is unrecognized it is returned unchanged.

Mapping:
    admin        -> "@"
    hq           -> "&"
    coordinator  -> "("
    manager      -> "~"
    supervisor   -> "$"
"""

ROLE_TO_SYMBOL = {
    "admin": "@",
    "hq": "&",
    "coordinator": "(",
    "manager": "~",
    "supervisor": "$",
}

SYMBOL_TO_ROLE = {sym: role for role, sym in ROLE_TO_SYMBOL.items()}


def encode_role(role: str) -> str:
    """
    Return the symbol for a canonical role name.
    If role is already a symbol or unknown, it is returned unchanged.
    """
    if not role:
        return role
    r = role.lower()
    return ROLE_TO_SYMBOL.get(r, role)


def decode_role(value: str) -> str:
    """
    Accept either a canonical role name or a symbol and return the canonical role name.
    If value is unknown it is returned unchanged.
    """
    if not value:
        return value
    low = value.lower()
    # Already canonical?
    if low in ROLE_TO_SYMBOL:
        return low
    # Symbol?
    if value in SYMBOL_TO_ROLE:
        return SYMBOL_TO_ROLE[value]
    return value


def is_known_symbol(value: str) -> bool:
    return value in SYMBOL_TO_ROLE


def is_canonical_role(value: str) -> bool:
    return (value or "").lower() in ROLE_TO_SYMBOL


def all_canonical_roles():
    return list(ROLE_TO_SYMBOL.keys())


def all_symbols():
    return list(SYMBOL_TO_ROLE.keys())


if __name__ == "__main__":
    # Simple self‑test
    for r in ROLE_TO_SYMBOL:
        sym = encode_role(r)
        back = decode_role(sym)
        assert back == r, f"Round trip failed for {r}"
    for sym in SYMBOL_TO_ROLE:
        assert decode_role(sym) in ROLE_TO_SYMBOL, f"Symbol decode failed for {sym}"
    print("role_map self-test passed.")