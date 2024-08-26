def preprocess_name(name):
    """Preprocess names by converting to uppercase and stripping whitespace."""
    if isinstance(name, str):
        return name.upper().strip()
    return ""


def extract_surname(name):
    """Extract the surname (assumed to be the last part of the name)."""
    parts = name.split()
    if len(parts) > 1:
        return parts[-1]  # Return the last word as the surname
    return ""
