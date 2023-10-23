import openpyxl
from PyPDF2 import PdfReader
import os

# Paths declarations
OUTPUT_XLSX_PATH = "output.xlsx"
GETPDFID_DIR = "getpdfid"

def get_unique_id_from_pdf(pdf_path):
    """Extract the Unique ID from the PDF metadata"""
    reader = PdfReader(open(pdf_path, "rb"))
    return reader.metadata.get('/UniqueID', None)

def get_name_mapping(unique_id, mapping_data):
    """Retrieve the name corresponding to a Unique ID from the mapping data"""
    return mapping_data.get(unique_id, None)

# Load the mapping data from the output.xlsx
workbook = openpyxl.load_workbook(OUTPUT_XLSX_PATH)
sheet = workbook.active
mapping_data = {row[2]: row[0] for row in sheet.iter_rows(values_only=True) if row[2]}

# Go through all PDFs in the getpdfid folder
results = {}
for pdf_file in os.listdir(GETPDFID_DIR):
    if pdf_file.endswith(".pdf"):
        pdf_path = os.path.join(GETPDFID_DIR, pdf_file)
        unique_id = get_unique_id_from_pdf(pdf_path)
        name = get_name_mapping(unique_id, mapping_data)
        if name:
            results[pdf_file] = name
        else:
            print(f"No mapping found for {pdf_file} with Unique ID: {unique_id}")

# Display the results
for pdf_file, name in results.items():
    print(f"PDF {pdf_file} เป็นของคุณ {name}")
