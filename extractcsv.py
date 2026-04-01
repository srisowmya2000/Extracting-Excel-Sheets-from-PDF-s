#Windows: Download and install pdftk from pdflabs.com.
#Mac:
#!brew install pdftk-java
#Linux:
#sudo apt-get install pdftk
import os
import shutil
import subprocess
import argparse


def extract_attachments(pdf_path, temp_dir):
    try:
        output = subprocess.run(
            ["pdftk", pdf_path, "unpack_files", "output", temp_dir],
            capture_output=True,
            text=True,
        )
        if output.returncode == 0:
            print(f"Extracted attachments from {pdf_path}")
            return True
        print(f"Failed to extract attachments from {pdf_path}: {output.stderr}")
        return False
    except Exception as e:
        print(f"Error processing {pdf_path}: {e}")
        return False


def process_pdfs(pdf_directory, output_directory):
    os.makedirs(output_directory, exist_ok=True)
    temp_dir = os.path.join(output_directory, "temp")

    for root, _, files in os.walk(pdf_directory):
        for file in files:
            if file.lower().endswith(".pdf"):
                pdf_path = os.path.join(root, file)
                file_prefix = os.path.splitext(file)[0]

                os.makedirs(temp_dir, exist_ok=True)

                if extract_attachments(pdf_path, temp_dir):
                    for extracted_file in os.listdir(temp_dir):
                        if extracted_file.lower().endswith(".csv"):
                            src_path = os.path.join(temp_dir, extracted_file)
                            dest_file_name = f"{file_prefix}_{extracted_file}"
                            dest_path = os.path.join(output_directory, dest_file_name)
                            shutil.move(src_path, dest_path)
                            print(f"Saved {dest_file_name}")

                    shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Extract embedded CSV files from PDF reports."
    )
    parser.add_argument("pdf_directory", help="Directory containing PDF files")
    parser.add_argument("output_directory", help="Directory to save extracted CSV files")
    args = parser.parse_args()

    process_pdfs(args.pdf_directory, args.output_directory)
