# correction_logic.py

import pandas as pd
from thefuzz import fuzz

from src.preprocessing import preprocess_name, extract_surname, normalize_name


def correct_names_in_excel(file_path, similarity_threshold=90):
    """
    Correct names in Excel sheets based on strong similarity to names in column 1,
    focusing on surname-based corrections and normalizing initials.

    Args:
        file_path (str): Path to the Excel file.
        similarity_threshold (int): Threshold for similarity score (0-100).
    """
    # Load the Excel file
    xls = pd.ExcelFile(file_path)

    # Create a dictionary to hold corrected DataFrames
    corrected_sheets = {}

    # Store corrections made across all sheets
    all_corrections = {}

    # Iterate over each sheet in the Excel file (skipping the first one)
    for sheet_name in xls.sheet_names[1:]:
        df = pd.read_excel(xls, sheet_name=sheet_name)

        # Skip processing if the sheet is empty
        if df.empty:
            continue

        # Ensure there are at least 1 column and up to 9 columns
        max_cols = min(9, df.shape[1])
        if max_cols < 2:
            corrected_sheets[sheet_name] = df
            continue

        # Preprocess correct names from column 1
        correct_names = df.iloc[:, 0].dropna().apply(preprocess_name).unique()

        # Create a dictionary mapping surnames to the most common full name
        surname_to_name = {}
        for name in correct_names:
            surname = extract_surname(name)
            if surname:
                surname_to_name[surname] = name

        # Initialize corrections dictionary for the current sheet
        sheet_corrections = {}

        # Iterate through each correct name and its surname
        for surname, correct_name in surname_to_name.items():
            # Iterate through columns 2 to max_cols
            for col in range(1, max_cols):
                # Get the column data
                column_data = df.iloc[:, col]

                # Iterate through each cell in the column
                for idx, cell_value in column_data.items():
                    original_value = cell_value
                    preprocessed_value = preprocess_name(cell_value)

                    # Skip if cell is empty
                    if not preprocessed_value:
                        continue

                    # Normalize the name with respect to the surname
                    normalized_value = normalize_name(preprocessed_value, surname)

                    # Extract the surname from the preprocessed value
                    value_surname = extract_surname(normalized_value)

                    # Calculate similarity score based on the surname
                    if value_surname and fuzz.ratio(value_surname, surname) >= similarity_threshold:
                        # If surnames match and values are different, correct it
                        if normalized_value != correct_name:
                            # Make the correction in the DataFrame
                            df.iat[idx, col] = correct_name

                            # Record the correction
                            sheet_corrections[original_value] = correct_name

        # Add sheet corrections to all corrections
        if sheet_corrections:
            all_corrections[sheet_name] = sheet_corrections

        # Store the corrected DataFrame
        corrected_sheets[sheet_name] = df

    return corrected_sheets, all_corrections
