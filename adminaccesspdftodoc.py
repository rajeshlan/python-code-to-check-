import zipfile
import os
import re
import PyPDF2
import win32com.client
from pdf2docx import Converter

# Define a function to extract PDFs from a zipped file
def extract_pdfs_from_zip(zip_file_path, output_dir):
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        zip_ref.extractall(output_dir)

# Define a function to convert PDF to DOC using win32com
def convert_pdf_to_doc(pdf_file_path, doc_file_path):
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Add()
    doc.SaveAs(doc_file_path, 0)
    doc.Close()
    word.Quit()

# Define a function to extract names, phone numbers, and emails from a DOCX file
def extract_data_from_docx(docx_file_path):
    # Add your code here to extract data from the DOCX file
    pass  # Placeholder, replace with actual code

# Main function to extract data from a zipped file
def main(zip_file_path, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)  # Create the output directory if it doesn't exist

    extract_pdfs_from_zip(zip_file_path, output_dir)
    pdf_files = [f for f in os.listdir(output_dir) if f.endswith('.pdf')]

    for pdf_file in pdf_files:
        pdf_file_path = os.path.join(output_dir, pdf_file)
        doc_file_path = os.path.splitext(pdf_file_path)[0] + '.doc'

        convert_pdf_to_doc(pdf_file_path, doc_file_path)

        # Add code here to extract data from the DOC file
        # Call extract_data_from_docx function with doc_file_path as an argument

if __name__ == "__main__":
    zip_file_path = r"C:\Users\rajes\OneDrive\Desktop\CV SEPT112023-20230913T044719Z-001\CV SEPT112023"
    output_dir = r"C:\Users\rajes\OneDrive\Desktop\CV (1)"
    main(zip_file_path, output_dir)
