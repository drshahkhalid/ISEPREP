def get_translations():
    # Get translations from the database
    translations = {
        'IN Type': _localize_text('IN Type'),
        'comments': _localize_comments('comments'),
        'popups': _localize_popups('popups')
    }
    return translations


def save_to_database(data):
    # Insert or update record in the database
    data['document_number'] = data.get('document_number', '').strip()  # Keep document number in English
    data['translations'] = get_translations()  # Ensure translations are saved correctly
    db.save(data)  

