#Windows: Download and install pdftk from pdflabs.com.
#Mac:
#!brew install pdftk-java
#Linux:
#sudo apt-get install pdftk
import os
import subprocess
import shutil

# Define the directories
#edit it when you try with your files
pdf_directory = r"C:\Users\Nat\pdf2excel\Final_code_oit Automation"
output_directory = r"C:\Users\Nat\pdf2excel\Final_code_oit Automation\Extracts"


def extract_attachments(pdf_path, temp_dir):
    try:
        # Run the pdftk command to extract attachments
        output = subprocess.run(['pdftk', pdf_path, 'unpack_files', 'output', temp_dir],
                                capture_output=True, text=True)
        
        if output.returncode == 0:
            print(f"Extracted attachments from {pdf_path}")
            return True
        else:
            print(f"Failed to extract attachments from {pdf_path}: {output.stderr}")
            return False
    except Exception as e:
        print(f"Error processing {pdf_path}: {e}")
        return False
        
def process_pdfs(pdf_directory, output_directory):
    # Create the output directory if it doesn't exist
    os.makedirs(output_directory, exist_ok=True)
    temp_dir = os.path.join(output_directory, 'temp')
    
    # Iterate through all the PDF files in the given directory
    for root, dirs, files in os.walk(pdf_directory):
        for file in files:
            if file.lower().endswith(".pdf"):
                pdf_path = os.path.join(root, file)
                file_prefix = os.path.splitext(file)[0]
                
                # Create a temporary directory for extraction
                os.makedirs(temp_dir, exist_ok=True)
                
                if extract_attachments(pdf_path, temp_dir):
                    # Move and rename extracted CSV files
                    for extracted_file in os.listdir(temp_dir):
                        if extracted_file.lower().endswith(".csv"):
                            src_path = os.path.join(temp_dir, extracted_file)
                            dest_file_name = f"{file_prefix}_{extracted_file}"
                            dest_path = os.path.join(output_directory, dest_file_name)
                            shutil.move(src_path, dest_path)
                            print(f"Saved {dest_file_name}")
                
                # Clean up the temporary directory
                shutil.rmtree(temp_dir)
                
# Process all PDFs in the directory and extract CSV attachments
process_pdfs(pdf_directory, output_directory)
                        
