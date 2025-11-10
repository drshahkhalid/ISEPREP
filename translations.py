import json
import os

# Folder containing translation JSON files
TRANSLATION_FOLDER = os.path.join(os.path.dirname(__file__), "data", "translations")

# Global variables
_translations = {}        # Holds all language dictionaries
_current_language = "en"  # Default language

def load_translations():
    """
    Load all JSON translation files into _translations dictionary
    """
    global _translations
    _translations.clear()
    print(f"Loading translations from: {TRANSLATION_FOLDER}")

    if not os.path.exists(TRANSLATION_FOLDER):
        print(f"Error: Translation folder {TRANSLATION_FOLDER} does not exist")
        return

    # Loop through all JSON files in translations folder
    for filename in os.listdir(TRANSLATION_FOLDER):
        if filename.endswith(".json"):
            lang_code = filename.split(".")[0]  # e.g., en.json → "en"
            file_path = os.path.join(TRANSLATION_FOLDER, filename)
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    _translations[lang_code] = json.load(f)
                print(f"Loaded translation file: {filename}")
            except Exception as e:
                print(f"Error loading {filename}: {e}")

def set_language(language_code: str):
    """
    Set active language. Defaults to English if code not found.
    """
    global _current_language
    if language_code in _translations:
        _current_language = language_code
        print(f"Set language to: {language_code}")
    else:
        print(f"[Warning] Language '{language_code}' not found. Falling back to English.")
        _current_language = "en"

def t(key: str, **kwargs) -> str:
    """
    Get translated text using nested key (e.g., 'login.username').

    - Splits key by '.' to traverse nested dictionaries
    - Falls back to key itself if not found
    """
    parts = key.split(".")
    value = _translations.get(_current_language, {})
    for part in parts:
        if isinstance(value, dict) and part in value:
            value = value[part]
        else:
            print(f"Translation key not found: {key}")
            return key.format(**kwargs) if kwargs else key
    return value.format(**kwargs) if kwargs else value

# Load translations immediately
load_translations()