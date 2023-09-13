import os
import re
import PyPDF2

# Function to extract text from a PDF file
def extract_text_from_pdf(pdf_file):
    text = ''
    try:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        for page in pdf_reader.pages:
            text += page.extract_text()
    except Exception as e:
        print(f"Error extracting text from {pdf_file}: {str(e)}")
    return text

# Function to extract names and phone/contact numbers using regex
def extract_names_and_numbers(text):
    # Define regex patterns for names and phone numbers
    name_pattern = re.compile(r'(?i)Name: ([^\n]*)')
    phone_pattern = re.compile(r'(\(?\d{3}\D{0,3}\d{3}\D{0,3}\d{4})')

    # Find all matches of names and phone numbers in the text
    names = re.findall(name_pattern, text)
    phone_numbers = re.findall(phone_pattern, text)

    return names, phone_numbers

# Specify the directory containing PDF files
pdf_directory = "path/to/pdf/files_directory"

# Process the PDF files and extract names and phone numbers
for filename in os.listdir(pdf_directory):
    if filename.endswith(".pdf"):
        pdf_file = os.path.join(pdf_directory, filename)
        text = extract_text_from_pdf(pdf_file)
        names, phone_numbers = extract_names_and_numbers(text)

        # Display extracted names and phone numbers
        print(f"File: {pdf_file}")
        print("Extracted Names:")
        for name in names:
            print(name)

        print("Extracted Phone Numbers:")
        for phone_number in phone_numbers:
            print(phone_number)

        print("=" * 40)
