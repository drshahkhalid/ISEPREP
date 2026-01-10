# Multilingual Support Documentation

## Overview
IsEPREP now supports full multilingual operation in three languages:
- **English (en)** - Default
- **French (fr)** - Français
- **Spanish (es)** - Español

## How It Works

### Language Manager
The `language_manager.py` module handles all translations:
- Loads translation files (en.json, fr.json, es.json) from the repository root
- Provides `lang.t(key, fallback)` method for retrieving translated text
- Supports nested keys using dot notation (e.g., "stock_in.all_scenarios")

### Translation Files
Each language has its own JSON file with identical structure:
```json
{
  "stock_in": {
    "title": "Stock In",
    "all_scenarios": "All Scenarios",
    "in_msf": "In MSF",
    ...
  }
}
```

### Usage in Code
```python
from language_manager import lang

# Simple translation
title = lang.t("stock_in.title", "Stock In")

# Translation with placeholder
message = lang.t("stock_in.loaded_records", "Loaded {count} records").format(count=10)

# Change language
lang.set_language('fr')  # Switch to French
```

## Stock In Module (in_.py)

### Dropdown Translations
All dropdown options are translated dynamically:
- **Scenarios**: "All Scenarios" / "Tous les scénarios" / "Todos los Escenarios"
- **IN Types**: 
  - "In MSF" / "En MSF" / "En MSF"
  - "In Local Purchase" / "En achat local" / "En Compra Local"
  - "In from Quarantine" / "Depuis la quarantaine" / "Desde Cuarentena"
  - And all other IN type options

### Document Number Generation
The system maintains consistent document numbers across all languages:

**Format**: `YYYY/MM/<PROJECT_CODE>/<ABBR>/<SERIAL>`

**Example**:
- English: "In MSF" → Document: 2026/01/PRJ/IMSF/0001
- French: "En MSF" → Document: 2026/01/PRJ/IMSF/0001
- Spanish: "En MSF" → Document: 2026/01/PRJ/IMSF/0001

All three languages produce the same abbreviation (IMSF), ensuring consistent tracking.

### Excel Export
Excel exports use the active language for:
- Sheet title
- Column headers (Code, Description, Scenario, Kit, Module, etc.)
- IN Type label
- Date and document number formatting

**Example French Export**:
```
Date: 2026-01-10 14:30:00        Numéro de Document: 2026/01/PRJ/IMSF/0001
                                                Entrée de stock
                                                Project - CODE
                                                Type ENTRÉE: En MSF

Code | Description | Scénario | Kit | Module | Qté Std | ...
```

### Database Consistency
Backend operations remain in English to ensure data integrity:
- **Document abbreviations**: IMSF, ILP, IFQ, IDN, IREU, ISNM, IBR, IRL, ICOR
- **unique_id format**: `scenario/kit/module/code/std_qty/expiry`
- **Transaction types**: Stored in English for consistent reporting

## Adding New Translations

### Step 1: Add English Key
Edit `en.json`:
```json
{
  "stock_in": {
    ...
    "new_feature": "New Feature Text"
  }
}
```

### Step 2: Add French Translation
Edit `fr.json`:
```json
{
  "stock_in": {
    ...
    "new_feature": "Texte de nouvelle fonctionnalité"
  }
}
```

### Step 3: Add Spanish Translation
Edit `es.json`:
```json
{
  "stock_in": {
    ...
    "new_feature": "Texto de nueva funcionalidad"
  }
}
```

### Step 4: Use in Code
```python
label_text = lang.t("stock_in.new_feature", "New Feature Text")
```

## Best Practices

### 1. Always Use Translation Keys
❌ **Bad**: `label = "Stock In"`
✅ **Good**: `label = lang.t("stock_in.title", "Stock In")`

### 2. Provide Fallback Text
Always provide a fallback that matches the English text:
```python
lang.t("stock_in.title", "Stock In")  # "Stock In" is fallback
```

### 3. Use Placeholders for Dynamic Content
```python
# In JSON
"loaded_records": "Loaded {count} records"

# In code
msg = lang.t("stock_in.loaded_records", "Loaded {count} records").format(count=10)
```

### 4. Keep Keys Organized
Use hierarchical structure:
- `stock_in.*` - Stock In module
- `dialog_titles.*` - Dialog titles
- `receive_kit.*` - Receive Kit module

### 5. Maintain Consistency
Ensure all three JSON files have the same keys:
```bash
# Check consistency
python3 -c "
import json
with open('en.json') as f: en = set(json.load(f)['stock_in'].keys())
with open('fr.json') as f: fr = set(json.load(f)['stock_in'].keys())
with open('es.json') as f: es = set(json.load(f)['stock_in'].keys())
print('EN:', len(en), 'FR:', len(fr), 'ES:', len(es))
print('Missing in FR:', en - fr)
print('Missing in ES:', en - es)
"
```

## Testing Translations

### Manual Testing
1. Set language in login screen
2. Navigate to Stock In
3. Verify all UI elements are translated:
   - Labels
   - Dropdowns
   - Buttons
   - Error messages
4. Test Excel export
5. Check document numbers

### Automated Testing
```python
from language_manager import lang

# Test all languages load
for code in ['en', 'fr', 'es']:
    lang.set_language(code)
    assert lang.t('stock_in.title', 'FAIL') != 'FAIL'
    print(f"✓ {code} loaded")
```

## Troubleshooting

### Translation Not Showing
1. Check the key exists in JSON file
2. Verify JSON syntax is valid
3. Ensure language file is in correct location
4. Restart application to reload translations

### Inconsistent Abbreviations
The document number generation includes comprehensive mapping for all languages. If new IN types are added:
1. Add English translation to `en.json`
2. Add French translation to `fr.json`
3. Add Spanish translation to `es.json`
4. Update `generate_document_number()` in `in_.py` to include all three language variations

### JSON Validation
```bash
# Validate JSON syntax
python3 -c "import json; json.load(open('en.json'))" && echo "✓ en.json valid"
python3 -c "import json; json.load(open('fr.json'))" && echo "✓ fr.json valid"
python3 -c "import json; json.load(open('es.json'))" && echo "✓ es.json valid"
```

## Future Enhancements
- Add more languages (German, Arabic, etc.)
- Create translation management tool
- Implement language-specific date formatting
- Add currency formatting per language
- Support right-to-left languages

## Support
For translation issues or to add new languages, contact:
- Shah Khalid (e-pool Pharmacy Coordinator, OCG)
- GitHub: drshahkhalid/ISEPREP
