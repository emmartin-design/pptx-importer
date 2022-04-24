

def get_error_dict():
    """
    This provides a stable dictionary to show where errors appear in a report, preventing PPT's "repair" message
    """
    return {
        'categories': ['data mismatch', 'error', 'data mismatch', 'error'],
        'error': [0, 0, 0, 0]
    }
