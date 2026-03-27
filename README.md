# Extract CSV Files from PDF Reports

A Python utility to extract embedded CSV files from PDF reports in bulk.

This project scans a folder of PDF reports, unpacks file attachments from each PDF using `pdftk`, and saves all extracted `.csv` files into an output directory with cleaned, prefixed filenames.

## Why this project is useful

Many enterprise and reporting workflows generate PDF files that contain embedded CSV attachments. Manually opening each report and exporting attachments is slow and repetitive.

This script automates that process by:

- scanning a directory for PDF files
- extracting embedded attachments from each PDF
- filtering for CSV files only
- renaming the extracted CSVs using the source PDF filename
- saving everything into one output folder

## Features

- Bulk processing of PDF files
- Extracts embedded file attachments from PDFs
- Saves only CSV attachments
- Automatically renames output files to keep them organized
- Works well for report-processing and automation workflows

## How it works

The script:

1. Walks through a directory of PDF files
2. Uses `pdftk` to unpack embedded files from each PDF
3. Checks the extracted files for `.csv` attachments
4. Renames each CSV using the original PDF filename as a prefix
5. Moves the final files into the output directory

## Example

If the input folder contains:

- `report_january.pdf`
- `report_february.pdf`

and those PDFs contain embedded CSV files such as:

- `data.csv`
- `summary.csv`

the output may look like:

- `report_january_data.csv`
- `report_february_summary.csv`

## Project structure

```bash
.
├── extractcsv.py
├── requirements.txt
└── README.md
