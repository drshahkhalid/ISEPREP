import json
import os
from typing import Any, Dict, List, Optional

class LanguageManager:
    """
    Centralized language manager with support for:
      - Nested translation keys (dot notation)
      - Enum mapping helpers:
          * enum_map(section_key): canonical -> display mapping
          * enum_reverse_map(section_key): display -> canonical mapping
          * enum_to_display(section_key, canonical_value)
          * enum_to_display_list(section_key, canonical_list)
          * enum_to_canonical(section_key, display_value)
      - Graceful fallbacks if keys are missing or files invalid
    """

    def __init__(self, default_lang: str = "en"):
        self.lang_code: str = default_lang
        self.lang_lang: str = default_lang  # compatibility alias
        self.translations: Dict[str, Any] = {}
        self._enum_cache: Dict[str, Dict[str, str]] = {}           # section_key -> canonical->display
        self._enum_rev_cache: Dict[str, Dict[str, str]] = {}        # section_key -> normalized display->canonical
        self.load_language(default_lang)

    # ---------------- Core loading ----------------
    def load_language(self, lang_code: str) -> None:
        """
        Load translation JSON for given language code from data/translations/{lang_code}.json.
        Falls back to empty dict on error or missing file.
        """
        file_path = os.path.join("data", "translations", f"{lang_code}.json")
        try:
            if os.path.exists(file_path):
                with open(file_path, "r", encoding="utf-8") as f:
                    self.translations = json.load(f)
            else:
                print(f"[Warning] Translation file not found: {file_path}")
                self.translations = {}
        except json.JSONDecodeError:
            print(f"[Error] Failed to parse JSON: {file_path}")
            self.translations = {}
        except Exception as e:
            print(f"[Error] Unexpected when loading translations: {e}")
            self.translations = {}

        # Update language code + compatibility alias
        self.lang_code = lang_code
        self.lang_lang = lang_code

        # Clear enum caches when language changes
        self._enum_cache.clear()
        self._enum_rev_cache.clear()

    def set_language(self, lang_code: str) -> None:
        """
        Switch active language and reload translations.
        """
        self.load_language(lang_code)

    # ---------------- Dot-notation translation ----------------
    def _get_nested(self, keys: List[str], data: Any) -> Any:
        """
        Navigate nested dict using a list of keys.
        Returns None if any segment is missing.
        """
        cur = data
        for k in keys:
            if isinstance(cur, dict) and k in cur:
                cur = cur[k]
            else:
                return None
        return cur

    def t(self, key: str, fallback: Optional[str] = None, **kwargs) -> str:
        """
        Translate a key using dot notation.
        Fallbacks gracefully to provided fallback or the key itself.
        Supports basic str.format(**kwargs) for placeholders.
        """
        parts = key.split(".")
        text = self._get_nested(parts, self.translations)

        if text is None:
            text = fallback if fallback is not None else key

        if kwargs:
            try:
                text = text.format(**kwargs)
            except Exception:
                # Ignore formatting errors to avoid crashing UI
                pass
        return text

    # ---------------- Section helpers ----------------
    def get_section(self, section_key: str) -> Dict[str, Any]:
        """
        Return a nested section dict for a key like 'stock_in.in_types_map'.
        If not found or not a dict, returns {}.
        """
        parts = section_key.split(".")
        sect = self._get_nested(parts, self.translations)
        return sect if isinstance(sect, dict) else {}

    # ---------------- Enum mapping (canonical EN <-> display) ----------------
    def _build_enum_maps(self, section_key: str) -> None:
        """
        Build and cache mapping dicts for a section holding a canonical->display map.
        Example JSON section:
            "stock_in": {
              "in_types_map": {
                "In MSF": "Entrée MSF",
                "In Local Purchase": "Achat local",
                ...
              }
            }
        Cache:
            self._enum_cache[section_key]     = {canonical: display}
            self._enum_rev_cache[section_key] = {normalized_display: canonical}
        """
        if section_key in self._enum_cache:
            return

        section = self.get_section(section_key)
        enum_map: Dict[str, str] = {}
        enum_rev: Dict[str, str] = {}

        # Only include str->str entries
        for k, v in section.items():
            if isinstance(k, str) and isinstance(v, str):
                enum_map[k] = v
                # Normalize display for reverse lookup (trim + casefold for robustness)
                norm_display = v.strip().casefold()
                enum_rev[norm_display] = k

        self._enum_cache[section_key] = enum_map
        self._enum_rev_cache[section_key] = enum_rev

    def enum_map(self, section_key: str) -> Dict[str, str]:
        """
        Return canonical->display mapping for the section key.
        If section missing, returns empty dict.
        """
        self._build_enum_maps(section_key)
        return self._enum_cache.get(section_key, {})

    def enum_reverse_map(self, section_key: str) -> Dict[str, str]:
        """
        Return normalized display->canonical reverse mapping for the section key.
        If section missing, returns empty dict.
        """
        self._build_enum_maps(section_key)
        return self._enum_rev_cache.get(section_key, {})

    def enum_to_display(self, section_key: str, canonical_value: str, fallback: Optional[str] = None) -> str:
        """
        Map a canonical English enum value to display (localized) text.
        Falls back gracefully to canonical_value or provided fallback.
        """
        if not canonical_value:
            return fallback if fallback is not None else ""
        m = self.enum_map(section_key)
        return m.get(canonical_value, fallback if fallback is not None else canonical_value)

    def enum_to_display_list(self, section_key: str, canonical_list: List[str]) -> List[str]:
        """
        Map a list of canonical English enum values to a list of localized display labels.
        Each element falls back to its canonical if translation missing.
        """
        m = self.enum_map(section_key)
        out: List[str] = []
        for val in canonical_list:
            out.append(m.get(val, val))
        return out

    def enum_to_canonical(self, section_key: str, display_value: str, fallback: Optional[str] = None) -> str:
        """
        Map a localized display label back to canonical English.
        Matching is case-insensitive and ignores surrounding spaces.
        Falls back gracefully to provided fallback or the original display_value.
        """
        if not display_value:
            return fallback if fallback is not None else ""
        rev = self.enum_reverse_map(section_key)
        norm = display_value.strip().casefold()
        return rev.get(norm, fallback if fallback is not None else display_value)

# Global instance
lang = LanguageManager()

# Backward compatibility alias
lang_lang = lang