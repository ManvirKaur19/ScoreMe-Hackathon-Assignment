import PyPDF2
import re
import pandas as pd


def load_pdf(file_path):
    try:
        pdf = open(file_path, 'rb')
        reader = PyPDF2.PdfReader(
            pdf)
        return reader
    except Exception as e:
        print(f"Error loading PDF: {e}")
        return None


def detect_tables(page_text):
    table_data = []
    lines = page_text.split('\n')
    table = []

    for line in lines:
        if re.match(r'^\s*[\d\w\s\.,-]+$',
                    line):
            table.append(line.split())
        else:
            if table:
                table_data.append(table)
                table = []

    if table:
        table_data.append(table)

    return table_data


def extract_combined_table(reader):
    all_tables = []

    for page in reader.pages:
        page_text = page.extract_text()
        tables = detect_tables(page_text)
        all_tables.extend(tables)

    combined_table = []
    for table in all_tables:
        combined_table.extend(table)

    return combined_table


def export_to_excel(table_data, output_file):
    try:
        df = pd.DataFrame(table_data)
        df.to_excel(output_file, index=False,
                    header=False)
        print(f"Table extracted and saved to {output_file}")
    except Exception as e:
        print(f"Error exporting to Excel: {e}")


def process_pdf(file_path, output_file):
    try:
        pdf_reader = load_pdf(file_path)
        if pdf_reader:
            combined_table = extract_combined_table(pdf_reader)
            export_to_excel(combined_table, output_file)
    except Exception as e:
        print(f"Error processing PDF: {e}")

pdf_path = 'test3.pdf'
excel_output_path = 'extracted_table3.xlsx'
process_pdf(pdf_path, excel_output_path)
pdf_path = 'test5.pdf'
excel_output_path = 'extracted_table5.xlsx'
process_pdf(pdf_path, excel_output_path)
pdf_path = 'test6.pdf'
excel_output_path = 'extracted_table6.xlsx'
process_pdf(pdf_path, excel_output_path)
