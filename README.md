# PDF Table Extraction Tool

This tool is designed to detect and extract tables from PDF documents and export them to Excel files. It uses the PyPDF2 library to read PDF files and pandas to handle and export the data.

## Features

- Load PDF files and extract text.
- Detect tables in the extracted text.
- Combine and clean the detected tables.
- Export the combined table data to an Excel file.

## Requirements

- Python 3.x
- PyPDF2
- pandas
- openpyxl

## Installation

1. Clone the repository or download the script.

2. Install the required Python libraries using pip:

```bash
pip install PyPDF2 pandas openpyxl
```

## Usage

### Script Overview
The main functions in the script are:

- load_pdf(file_path): Loads the PDF file and returns a PDF reader object.
- detect_tables(page_text): Detects tables in the text extracted from a PDF page.
- extract_combined_table(reader): Extracts and combines tables from all pages of the PDF.
- export_to_excel(table_data, output_file): Exports the combined table data to an Excel file.
- process_pdf(file_path, output_file): Processes a PDF file and saves the extracted table to an Excel file.

### Running the Script
- Save the script to a file, e.g., extract_tables.py.
- Place the PDF files you want to process in the same directory as the script or specify their paths.
- Update the pdf_path and excel_output_path variables with the paths to your PDF files and desired output Excel files.
- Run the script to extract tables from a PDF:

    ```python
    python extract_tables.py
    ```

## Note

For this assignment I have added the 3 pdf test files name in the script itself along with the name of the extracted excel sheet name. So, If you run this script it will run for the test PDFs themselves. Also, While testing the file test6 wasn't getting read, so I believe it's corrupted. Kindly take a look into it.
