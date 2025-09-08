import gradio as gr
import os
import tempfile
import shutil
import zipfile
from sav_to_excel import convert_sav_to_excel, batch_convert_sav_to_excel

def process_single_sav_file(sav_file, output_dir="Converted"):
    """
    Process a single uploaded SAV file using the sav_to_excel script.
    
    Args:
        sav_file: The uploaded .sav file from Gradio
        output_dir: The output directory for the Excel file
        
    Returns:
        tuple: (excel_file_path, status_message)
    """
    if sav_file is None:
        return None, "Please upload a .sav file first."
    
    # Check if the file has a .sav extension
    if not sav_file.name.lower().endswith('.sav'):
        return None, "Please upload a valid .sav file."
    
    try:
        # Create output directory if it doesn't exist
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Generate output filename based on the input .sav filename
        sav_basename = os.path.basename(sav_file.name)
        excel_filename = os.path.splitext(sav_basename)[0] + ".xlsx"
        excel_path = os.path.join(output_dir, excel_filename)

        # Use the convert_sav_to_excel function from sav_to_excel.py
        convert_sav_to_excel(sav_file.name, excel_path)
        
        return excel_path, f"Successfully converted '{sav_basename}' to '{excel_filename}'"
        
    except Exception as e:
        return None, f"An error occurred: {e}"

def process_batch_sav_files(sav_files, output_dir="Converted"):
    """
    Process multiple uploaded SAV files using batch conversion.
    
    Args:
        sav_files: List of uploaded .sav files from Gradio
        output_dir: The output directory for the Excel files
        
    Returns:
        tuple: (zip_file_path, status_message)
    """
    if not sav_files or len(sav_files) == 0:
        return None, "Please upload .sav files first."
    
    # Filter valid .sav files
    valid_files = []
    for file in sav_files:
        if file and file.name.lower().endswith('.sav'):
            valid_files.append(file.name)
    
    if not valid_files:
        return None, "No valid .sav files found. Please upload files with .sav extension."
    
    try:
        # Create output directory if it doesn't exist
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Use batch conversion
        results = batch_convert_sav_to_excel(valid_files, output_dir)
        
        # Create a zip file with all converted Excel files
        zip_filename = f"converted_files_{len(results['successful'])}_files.zip"
        zip_path = os.path.join(output_dir, zip_filename)
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for success in results['successful']:
                if os.path.exists(success['output']):
                    zipf.write(success['output'], success['filename'])
        
        # Create status message
        status_msg = f"Batch conversion completed!\n"
        status_msg += f"Total files: {results['total']}\n"
        status_msg += f"Successful: {len(results['successful'])}\n"
        status_msg += f"Failed: {len(results['failed'])}\n"
        
        if results['failed']:
            status_msg += f"\nFailed files:\n"
            for failed in results['failed']:
                status_msg += f"• {os.path.basename(failed['input'])}: {failed['error']}\n"
        
        return zip_path, status_msg
        
    except Exception as e:
        return None, f"An error occurred during batch processing: {e}"

# Create the Gradio interface
def create_interface():
    with gr.Blocks(title="SAV to Excel Converter", theme=gr.themes.Soft()) as interface:
        gr.Markdown("""
        # 📊 SAV to Excel Converter
        
        Convert SPSS .sav files to Excel .xlsx files with proper formatting.
        The converter will:
        - Map variable labels to column headers
        - Handle date columns properly
        - Preserve data integrity
        - Support both single file and batch processing
        """)
        
        with gr.Tabs():
            # Single File Tab
            with gr.Tab("Single File"):
                with gr.Row():
                    with gr.Column(scale=1):
                        single_file_input = gr.File(
                            label="Upload SAV File",
                            file_types=[".sav"],
                            file_count="single"
                        )
                        
                        single_output_dir = gr.Textbox(
                            label="Output Directory",
                            value="Converted",
                            placeholder="Enter output directory (default: Converted)"
                        )
                        
                        single_convert_btn = gr.Button(
                            "Convert to Excel",
                            variant="primary",
                            size="lg"
                        )
                        
                    with gr.Column(scale=1):
                        single_status_output = gr.Textbox(
                            label="Status",
                            interactive=False,
                            lines=3
                        )
                        
                        single_file_output = gr.File(
                            label="Download Excel File",
                            interactive=False
                        )
            
            # Batch Processing Tab
            with gr.Tab("Batch Processing"):
                with gr.Row():
                    with gr.Column(scale=1):
                        batch_files_input = gr.File(
                            label="Upload Multiple SAV Files",
                            file_types=[".sav"],
                            file_count="multiple"
                        )
                        
                        batch_output_dir = gr.Textbox(
                            label="Output Directory",
                            value="Converted",
                            placeholder="Enter output directory (default: Converted)"
                        )
                        
                        batch_convert_btn = gr.Button(
                            "Convert All Files",
                            variant="primary",
                            size="lg"
                        )
                        
                    with gr.Column(scale=1):
                        batch_status_output = gr.Textbox(
                            label="Status",
                            interactive=False,
                            lines=8
                        )
                        
                        batch_file_output = gr.File(
                            label="Download ZIP File",
                            interactive=False
                        )
        
        # Set up the conversion processes
        single_convert_btn.click(
            fn=process_single_sav_file,
            inputs=[single_file_input, single_output_dir],
            outputs=[single_file_output, single_status_output]
        )
        
        batch_convert_btn.click(
            fn=process_batch_sav_files,
            inputs=[batch_files_input, batch_output_dir],
            outputs=[batch_file_output, batch_status_output]
        )
        
        # Example section
        gr.Markdown("""
        ## 📝 How to use:
        
        ### Single File Mode:
        1. Go to "Single File" tab
        2. Upload one .sav file
        3. Click "Convert to Excel"
        4. Download the converted Excel file
        
        ### Batch Processing Mode:
        1. Go to "Batch Processing" tab
        2. Upload multiple .sav files (or select multiple at once)
        3. Click "Convert All Files"
        4. Download the ZIP file containing all converted Excel files
        
        ## ℹ️ Features:
        - **Single & Batch Processing**: Handle one file or multiple files at once
        - **Automatic variable label mapping**: Column headers use descriptive labels
        - **Date column handling**: Proper date formatting for SPSS dates
        - **Progress tracking**: See conversion status and results
        - **Error handling**: Detailed error messages for failed conversions
        - **ZIP download**: Batch results packaged in a convenient ZIP file
        - **Clean, formatted output**: Preserves original data structure
        """)
    
    return interface

if __name__ == "__main__":
    # Create and launch the interface
    app = create_interface()
    app.launch(
        server_name="0.0.0.0",
        server_port=7860,
        show_error=True
    )
