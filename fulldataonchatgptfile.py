import re
import os
import PyPDF2
import zipfile
import pandas as pd

# Function to extract names, phone numbers, and email addresses from a text
def extract_info(text):
    # Regular expressions for extracting names, phone numbers, and email addresses
    name_pattern = r'\b[A-Z][a-z]+\s[A-Z][a-z]+\b'
    phone_pattern = r'\b\d{10}\b'
    email_pattern = r'\b\w+@\w+\.(?:com|co|in)\b'

    # Extracting names, phone numbers, and email addresses
    names = re.findall(name_pattern, text)
    phone_numbers = re.findall(phone_pattern, text)
    email_addresses = re.findall(email_pattern, text)

    return names, phone_numbers, email_addresses

# Function to extract text from a PDF file
def extract_text_from_pdf(pdf_path):
    pdf_text = ""
    try:
        pdf_file = open(pdf_path, 'rb')
        pdf_reader = PyPDF2.PdfFileReader(pdf_file)
        for page_num in range(pdf_reader.numPages):
            page = pdf_reader.getPage(page_num)
            pdf_text += page.extractText()
        pdf_file.close()
    except Exception as e:
        print(f"Error extracting text from PDF: {str(e)}")

    return pdf_text

# Extract text from PDF files
pdf_folder = r'C:\Users\rajes\OneDrive\Desktop\CV for assignment'
output_folder = 'output'

for filename in os.listdir(pdf_folder):
    if filename.endswith('.pdf'):
        pdf_path = os.path.join(pdf_folder, filename)
        pdf_text = extract_text_from_pdf(pdf_path)

        # Extract names, phone numbers, and email addresses from the extracted text
        names, phone_numbers, email_addresses = extract_info(pdf_text)

        # Print the results
        print(f"File: {filename}")
        print(f"Names: {', '.join(names)}")
        print(f"Phone Numbers: {', '.join(phone_numbers)}")
        print(f"Email Addresses: {', '.join(email_addresses)}")
        print("\n")

# You can modify the 'pdf_folder' variable to point to the folder containing your PDF files.
