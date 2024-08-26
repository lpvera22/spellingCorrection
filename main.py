# main.py

import json

import pandas as pd

from src.correction_logic import correct_names_in_excel


def save_corrected_file(file_path, corrected_sheets):
    """Save the corrected sheets back to a new Excel file."""
    corrected_file_path = file_path.replace('.xlsx', '_corrected.xlsx')
    with pd.ExcelWriter(corrected_file_path, engine='openpyxl') as writer:
        # Write the first (original) sheet unmodified
        xls = pd.ExcelFile(file_path)
        first_sheet = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
        first_sheet.to_excel(writer, sheet_name=xls.sheet_names[0], index=False)

        # Write corrected sheets
        for sheet_name, df in corrected_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Corrections have been made and saved to '{corrected_file_path}'.")


def save_corrections_as_json(file_path, all_corrections):
    """Save corrections as JSON."""
    json_file_path = file_path.replace('.xlsx', '_corrections.json')
    with open(json_file_path, 'w', encoding='utf-8') as json_file:
        json.dump(all_corrections, json_file, ensure_ascii=False, indent=4)

    print(f"Corrections dictionary has been saved to '{json_file_path}'.")


if __name__ == "__main__":
    file_path = "data/REP - Consolidated DLN 2024.xlsx"
    similarity_threshold = 90

    # Correct names in Excel
    corrected_sheets, all_corrections = correct_names_in_excel(file_path, similarity_threshold)

    # Save the results
    save_corrected_file(file_path, corrected_sheets)
    save_corrections_as_json(file_path, all_corrections)
