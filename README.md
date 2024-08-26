# Name Correction in Excel Sheets

This project provides a Python-based solution to standardize and correct names across multiple Excel sheets by focusing on surname consistency, removing prefixes like "Mr.", and normalizing name formats. It utilizes fuzzy matching to identify similar names and applies corrections based on a predefined similarity threshold.

## Project Structure

- `preprocessing.py`: Contains helper functions for preprocessing names, extracting surnames, removing prefixes, and normalizing names.
- `correction_logic.py`: Implements the core logic for correcting names across Excel sheets based on surname-based matching and normalization.
- `main.py`: The main script that ties everything together, processes the Excel file, and saves the corrected sheets and a log of corrections made.
- `data/`: This is the directory where your Excel files are stored. Replace the example file path with your actual file paths when using the script.

## How It Works

### 1. Preprocessing
Names are preprocessed to:
- Convert to uppercase.
- Strip leading and trailing whitespace.
- Remove common prefixes like "Mr.", "Mrs.", "Ms.", and "Dr."
- Extract surnames for consistency checks.

### 2. Correction Logic
- The script loads an Excel file and processes each sheet (excluding the first one) to ensure name consistency.
- The first column of each sheet is assumed to contain the correct names. The script identifies the most common full name for each surname.
- It then checks the remaining columns for similar names, using fuzzy matching to determine if the names should be corrected.
- Names are normalized by ensuring that all instances of a surname are corrected to the most common form.

### 3. Output
- **Corrected Excel File**: A new Excel file is saved with the corrected names.
- **Corrections Log**: A JSON file is created to log all the corrections made during the process.

## Installation and Setup

### Prerequisites
- Python 3.7+
- Required Python packages:
  - `pandas`
  - `thefuzz`
  - `openpyxl`

You can install the required packages using pip:

```bash
pip install pandas thefuzz openpyxl
```
## Running the Project
- Place your Excel file in the data/ directory.
- Open a terminal and navigate to the project directory.
- Run the main script:
    ```bash
    python main.py
    ```
- This will process the Excel file, correct the names, and save the corrected file and corrections log in the same directory as the original file.
### Configuration
You can adjust the similarity threshold for name corrections by modifying the similarity_threshold variable in main.py. This threshold determines how closely two names must match to be considered the same.
## Example
Assume you have an Excel file with the following names:
- Sheet 1 (Correct Names): Devanand Shukla, Sumesh Appunni, P.N.Kurup
- Sheet 2 (Names to be Corrected): DN Shukla, Mr. A Sumesh, P.N.Kurup, Shukla Devanand, Sumesh Appuni, Kurup Narayanan

After running the script:
- DN Shukla, Mr. A Sumesh, P.N.Kurup, Shukla Devanand, Sumesh Appuni, and Kurup Narayanan will be corrected to Devanand Shukla, Sumesh Appunni, and P.N.Kurup, respectively.

## Limitations
- The script assumes the first sheet contains the correct names.
- It only processes up to the first 9 columns in each sheet.
- Some edge cases (e.g., very similar names with different meanings) might require manual review after processing.
## Contributing
If you'd like to contribute to this project, feel free to fork the repository and submit a pull request with your improvements or bug fixes.