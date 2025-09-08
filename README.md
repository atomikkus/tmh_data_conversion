# SAV to Excel Converter

## Description

This script converts SPSS (.sav) files to Excel (.xlsx) files. It performs the following operations:
- Maps variable labels from the SAV file to column headers in the Excel file.
- Maps value labels for coded variables (e.g., converting `1` to `"Male"`).
- Handles SPSS date formats, converting them to date strings in `YYYY-MM-DD` format.

## Installation

1.  **Clone the repository or download the files.**

2.  **Create a Python virtual environment.** This is recommended to avoid conflicts with other projects.
    ```bash
    python -m venv venv
    ```

3.  **Activate the virtual environment.**
    -   On Windows:
        ```bash
        venv\Scripts\activate
        ```
    -   On macOS and Linux:
        ```bash
        source venv/bin/activate
        ```

4.  **Install the required dependencies.**
    ```bash
    pip install -r requirements.txt
    ```

## Usage

1.  **Place your .sav files** in the `SAV_RAW` directory.

2.  **Run the script** from the command line, providing the input and output file paths. The converted file will be saved in the `Converted` directory.

    ```bash
    python sav_to_excel.py SAV_RAW\<your_input_file>.sav Converted\<your_output_file>.xlsx
    ```

    **Example:**
    ```bash
    python sav_to_excel.py "SAV_RAW\MTB Analysis 2024_FINAL_Ranadheer.sav" "Converted\MTB Analysis 2024_FINAL_Ranadheer.xlsx"
    ```
