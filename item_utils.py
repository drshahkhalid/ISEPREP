# item_utils.py
# Shared helpers for item designation columns and type filtering logic

def get_designation_col(lang_code: str) -> str:
    """
    Return the correct designation column based on current language.
    Defaults to English if unknown.
    """
    return {
        "fr": "designation_fr",
        "es": "designation_sp"
    }.get(lang_code, "designation_en")


def build_type_filter(level: str, designation_col: str) -> str:
    """
    Build SQL fragment to handle type fallback logic:
    - Match exact type (KIT/MODULE/ITEM)
    - If type is NULL, fall back to prefix heuristics (code/designation contains KIT/MOD/ITE)
    """
    return f"""
        (
            UPPER(type) = %s OR
            (
                type IS NULL AND (
                    (UPPER(code) LIKE %s) OR
                    (UPPER({designation_col}) LIKE %s)
                )
            )
        )
    """
