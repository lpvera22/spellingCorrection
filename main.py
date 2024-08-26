import json

import pandas as pd
from thefuzz import fuzz


def preprocess_name(name):
    """Preprocess names by converting to uppercase and stripping whitespace."""
    if isinstance(name, str):
        return name.upper().strip()
    return ""


def correct_names_in_excel(file_path, similarity_threshold=90):
    """
    Correct names in Excel sheets based on strong similarity to names in column 1.

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

        # Initialize corrections dictionary for current sheet
        sheet_corrections = {}

        # Iterate through each correct name
        for correct_name in correct_names:
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

                    # Calculate similarity score
                    similarity_score = fuzz.ratio(preprocessed_value, correct_name)

                    # If similarity is above threshold and values are different, correct it
                    if similarity_score >= similarity_threshold and preprocessed_value != correct_name:
                        # Make the correction in DataFrame
                        df.iat[idx, col] = correct_name

                        # Record the correction
                        sheet_corrections[original_value] = correct_name

        # Add sheet corrections to all corrections
        if sheet_corrections:
            all_corrections[sheet_name] = sheet_corrections

        # Store the corrected DataFrame
        corrected_sheets[sheet_name] = df

    # Save the corrected sheets back to a new Excel file
    corrected_file_path = file_path.replace('.xlsx', '_corrected.xlsx')
    with pd.ExcelWriter(corrected_file_path, engine='openpyxl') as writer:
        # Write the first (original) sheet unmodified
        first_sheet = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
        first_sheet.to_excel(writer, sheet_name=xls.sheet_names[0], index=False)

        # Write corrected sheets
        for sheet_name, df in corrected_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Save corrections as JSON
    json_file_path = file_path.replace('.xlsx', '_corrections.json')
    with open(json_file_path, 'w', encoding='utf-8') as json_file:
        json.dump(all_corrections, json_file, ensure_ascii=False, indent=4)

    print(f"Corrections have been made and saved to '{corrected_file_path}'.")
    print(f"Corrections dictionary has been saved to '{json_file_path}'.")


correct_names_in_excel(file_path="data/REP - Consolidated DLN 2024.xlsx")
