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

def batch_convert_sav_to_excel(sav_files, output_dir="Converted"):
    """
    Convert multiple .sav files to Excel files in batch mode.
    
    Args:
        sav_files (list): List of paths to .sav files
        output_dir (str): Output directory for Excel files
        
    Returns:
        dict: Results with success/failure status for each file
    """
    results = {
        'successful': [],
        'failed': [],
        'total': len(sav_files)
    }
    
    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    print(f"Starting batch conversion of {len(sav_files)} files...")
    
    for i, sav_file in enumerate(sav_files, 1):
        try:
            print(f"Processing file {i}/{len(sav_files)}: {os.path.basename(sav_file)}")
            
            # Generate output filename
            sav_basename = os.path.basename(sav_file)
            excel_filename = os.path.splitext(sav_basename)[0] + ".xlsx"
            excel_path = os.path.join(output_dir, excel_filename)
            
            # Convert the file
            convert_sav_to_excel(sav_file, excel_path)
            results['successful'].append({
                'input': sav_file,
                'output': excel_path,
                'filename': excel_filename
            })
            print(f"✓ Successfully converted: {excel_filename}")
            
        except Exception as e:
            error_msg = f"Failed to convert {sav_file}: {e}"
            results['failed'].append({
                'input': sav_file,
                'error': str(e)
            })
            print(f"✗ {error_msg}")
    
    # Print summary
    print(f"\n=== Batch Conversion Summary ===")
    print(f"Total files: {results['total']}")
    print(f"Successful: {len(results['successful'])}")
    print(f"Failed: {len(results['failed'])}")
    
    if results['failed']:
        print(f"\nFailed files:")
        for failed in results['failed']:
            print(f"  - {os.path.basename(failed['input'])}: {failed['error']}")
    
    return results

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert .sav file(s) to Excel file(s).")
    parser.add_argument("sav_files", nargs='+', help="Path(s) to input .sav file(s). Can be multiple files or a directory.")
    parser.add_argument("output_dir", nargs='?', default="Converted", 
                       help="The output directory for the Excel file(s) (default: 'Converted').")
    parser.add_argument("--batch", action="store_true", 
                       help="Process multiple files in batch mode (default when multiple files provided).")
    args = parser.parse_args()

    # Check if any of the arguments are directories
    sav_files = []
    for path in args.sav_files:
        if os.path.isdir(path):
            # Find all .sav files in the directory
            for root, dirs, files in os.walk(path):
                for file in files:
                    if file.lower().endswith('.sav'):
                        sav_files.append(os.path.join(root, file))
        else:
            sav_files.append(path)
    
    # Remove duplicates and sort
    sav_files = sorted(list(set(sav_files)))
    
    if not sav_files:
        print("No .sav files found!")
        exit(1)
    
    # Determine if we should use batch mode
    use_batch = args.batch or len(sav_files) > 1
    
    if use_batch:
        batch_convert_sav_to_excel(sav_files, args.output_dir)
    else:
        # Single file mode
        sav_file = sav_files[0]
        excel_filename = os.path.splitext(os.path.basename(sav_file))[0] + ".xlsx"
        excel_path = os.path.join(args.output_dir, excel_filename)
        convert_sav_to_excel(sav_file, excel_path)