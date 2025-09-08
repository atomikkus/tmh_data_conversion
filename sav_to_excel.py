import pandas as pd
import pyreadstat
import argparse
import os
from datetime import datetime, timedelta

def spss_date_to_string(spss_date):
    if pd.isna(spss_date):
        return None
    try:
        # SPSS dates are seconds since October 14, 1582
        start_date = datetime(1582, 10, 14)
        return (start_date + timedelta(seconds=spss_date)).strftime('%Y-%m-%d')
    except (ValueError, TypeError, OverflowError):
        return None

def convert_sav_to_excel(sav_path, excel_path):
    """
    Converts an SPSS .sav file to an Excel .xlsx file,
    mapping variable labels to column headers and handling dates.

    Args:
        sav_path (str): The path to the input .sav file.
        excel_path (str): The path to the output .xlsx file.
    """
    try:
        # Read the .sav file
        df, meta = pyreadstat.read_sav(sav_path, disable_datetime_conversion=True, apply_value_formats=True)

        # Create a mapping from variable names to variable labels, only if the label exists
        column_map = {var: label for var, label in meta.column_names_to_labels.items() if label}

        # Rename the DataFrame columns if there are any labels
        if column_map:
            df.rename(columns=column_map, inplace=True)

        # Handle specific date columns
        date_columns = ['Date_Discussed_MTB', 'Date_NGS_Perfomed']
        for col in date_columns:
            # The column might have been renamed
            new_col_name = column_map.get(col, col)
            if new_col_name in df.columns:
                df[new_col_name] = df[new_col_name].apply(spss_date_to_string)

        # Write the DataFrame to an Excel file
        df.to_excel(excel_path, index=False)

        print(f"Successfully converted '{sav_path}' to '{excel_path}'")

    except FileNotFoundError:
        print(f"Error: The file '{sav_path}' was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert a .sav file to an Excel file.")
    parser.add_argument("sav_file", help="The path to the input .sav file.")
    parser.add_argument("output_dir", nargs='?', default="Converted", 
                       help="The output directory for the Excel file (default: 'Converted').")
    args = parser.parse_args()

    # Create the output directory if it doesn't exist
    if not os.path.exists(args.output_dir):
        os.makedirs(args.output_dir)

    # Generate output filename based on the input .sav filename
    sav_basename = os.path.basename(args.sav_file)
    excel_filename = os.path.splitext(sav_basename)[0] + ".xlsx"
    excel_path = os.path.join(args.output_dir, excel_filename)

    convert_sav_to_excel(args.sav_file, excel_path)