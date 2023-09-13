import zipfile
import os
import re
import PyPDF2
from pdf2docx import Converter
from docx import Document

# Define a function to extract PDFs from a zipped file
def extract_pdfs_from_zip(zip_file_path, output_dir):
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        zip_ref.extractall(output_dir)

# Define a function to convert PDF to DOCX using pdf2docx
def convert_pdf_to_docx(pdf_file_path, docx_file_path):
    cv = Converter(pdf_file_path)
    cv.convert(docx_file_path, start=0, end=None)
    cv.close()

# Define a function to extract names, phone numbers, and emails from a DOCX file
def extract_data_from_docx(docx_file_path):
    doc = Document(docx_file_path)
    text = "\n".join([para.text for para in doc.paragraphs])

    # Regular expressions to find names, phone numbers, and emails
    name_pattern = r'\b[A-Z][a-z]+ [A-Z][a-z]+\b'
    phone_pattern = r'\b\d{1,4}[-\.\s]?\d{1,3}[-\.\s]?\d{1,4}[-\.\s]?\d{1,4}\b'
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b'

    names = re.findall(name_pattern, text)
    phones = re.findall(phone_pattern, text)
    emails = re.findall(email_pattern, text)

    return names, phones, emails

# Main function to extract data from a zipped file
def main(zip_file_path, output_dir):
    extract_pdfs_from_zip(zip_file_path, output_dir)
    pdf_files = [f for f in os.listdir(output_dir) if f.endswith('.pdf')]

    for pdf_file in pdf_files:
        pdf_file_path = os.path.join(output_dir, pdf_file)
        docx_file_path = os.path.splitext(pdf_file_path)[0] + '.docx'

        convert_pdf_to_docx(pdf_file_path, docx_file_path)
        names, phones, emails = extract_data_from_docx(docx_file_path)

        print(f"Data from {pdf_file}:")
        print("Names:", names)
        print("Phone Numbers:", phones)
        print("Emails:", emails)
        print("\n")

if __name__ == "__main__":
    zip_file_path = r"C:\Users\rajes\OneDrive\Desktop\CV SEPT112023.zip"  # Replace with the path to your zip file
    output_dir = r"C:\Users\rajes\OneDrive\Desktop\CV (1)"  # Replace with the output directory path
    main(zip_file_path, output_dir)
