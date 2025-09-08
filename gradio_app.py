import gradio as gr
import os
import tempfile
import shutil
from sav_to_excel import convert_sav_to_excel

def process_sav_file(sav_file, output_dir="Converted"):
    """
    Process the uploaded SAV file using the sav_to_excel script.
    
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

# Create the Gradio interface
def create_interface():
    with gr.Blocks(title="SAV to Excel Converter", theme=gr.themes.Soft()) as interface:
        gr.Markdown("""
        # 📊 SAV to Excel Converter
        
        Upload an SPSS .sav file and convert it to an Excel .xlsx file with proper formatting.
        The converter will:
        - Map variable labels to column headers
        - Handle date columns properly
        - Preserve data integrity
        """)
        
        with gr.Row():
            with gr.Column(scale=1):
                file_input = gr.File(
                    label="Upload SAV File",
                    file_types=[".sav"],
                    file_count="single"
                )
                
                output_dir_input = gr.Textbox(
                    label="Output Directory",
                    value="Converted",
                    placeholder="Enter output directory (default: Converted)"
                )
                
                convert_btn = gr.Button(
                    "Convert to Excel",
                    variant="primary",
                    size="lg"
                )
                
            with gr.Column(scale=1):
                status_output = gr.Textbox(
                    label="Status",
                    interactive=False,
                    lines=3
                )
                
                file_output = gr.File(
                    label="Download Excel File",
                    interactive=False
                )
        
        # Set up the conversion process
        convert_btn.click(
            fn=process_sav_file,
            inputs=[file_input, output_dir_input],
            outputs=[file_output, status_output]
        )
        
        # Example section
        gr.Markdown("""
        ## 📝 How to use:
        1. Click "Upload SAV File" and select your .sav file
        2. Click "Convert to Excel" to process the file
        3. Download the converted Excel file when ready
        
        ## ℹ️ Features:
        - Automatic variable label mapping
        - Date column handling
        - Preserves original data structure
        - Clean, formatted output
        """)
    
    return interface

if __name__ == "__main__":
    # Create and launch the interface
    app = create_interface()
    app.launch(
        server_name="0.0.0.0",
        server_port=7860,
        share=True,
        show_error=True
    )
