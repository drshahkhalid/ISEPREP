import json
import os

class LanguageManager:
    """
    Centralized language manager.
    - Loads translations from JSON files under `data/translations/`.
    - Supports nested keys using dot notation (e.g., "kits.add_item").
    - Falls back gracefully to the key or a provided fallback text.
    """

    def __init__(self, default_lang="en"):
        # Default language code
        self.lang_code = default_lang
        # Compatibility alias (used in some older modules)
        self.lang_lang = default_lang
        self.translations = {}

        # Load initial language
        self.load_language(default_lang)

    def load_language(self, lang_code):
        """
        Load translation JSON for given language code.
        Falls back to empty dict if file not found.
        """
        file_path = os.path.join("data", "translations", f"{lang_code}.json")

        if os.path.exists(file_path):
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    self.translations = json.load(f)
            except json.JSONDecodeError:
                print(f"[Error] Failed to parse JSON: {file_path}")
                self.translations = {}
        else:
            print(f"[Warning] Translation file not found: {file_path}")
            self.translations = {}

        # Update language code
        self.lang_code = lang_code
        self.lang_lang = lang_code

    def _get_nested(self, keys, data):
        """
        Helper: Navigate nested dict using key list.
        """
        for key in keys:
            if isinstance(data, dict) and key in data:
                data = data[key]
            else:
                return None
        return data

    def t(self, key, fallback=None, **kwargs):
        """
        Translate a key.
        - Supports dot notation for nested JSON keys.
        - Allows placeholder replacement via kwargs.
        """
        keys = key.split(".")
        text = self._get_nested(keys, self.translations)

        # Fallbacks
        if text is None:
            text = fallback if fallback is not None else key

        # Placeholder formatting
        if kwargs:
            try:
                text = text.format(**kwargs)
            except KeyError:
                # Ignore missing placeholders
                pass

        return text

    def set_language(self, lang_code):
        """
        Switch language and reload translations.
        """
        self.load_language(lang_code)


# Global instance
lang = LanguageManager()

# Backward compatibility alias
lang_lang = lang
