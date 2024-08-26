# preprocessing.py

def preprocess_name(name):
    """Preprocess names by converting to uppercase, removing prefixes, and stripping whitespace."""
    if isinstance(name, str):
        # Remove common prefixes like "Mr."
        name = name.upper().strip()
        prefixes = ["MR.", "MRS.", "MS.", "DR."]  # Add more prefixes if needed
        for prefix in prefixes:
            if name.startswith(prefix):
                name = name[len(prefix):].strip()
        return name
    return ""


def extract_surname(name):
    """Extract the surname (assumed to be the last part of the name)."""
    parts = name.split()
    if len(parts) > 1:
        return parts[-1]  # Return the last word as the surname
    return ""


def normalize_name(name, surname):
    """Normalize the name by ensuring the surname is consistent."""
    name_parts = name.split()
    # If the surname is in the name, ensure it is consistently used
    if surname in name_parts:
        surname_index = name_parts.index(surname)
        # Keep the surname and any initials or first names before it
        normalized_name = " ".join(name_parts[:surname_index + 1])
        return normalized_name
    return name  # Return the original name if surname not found
