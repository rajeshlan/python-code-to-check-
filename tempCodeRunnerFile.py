import re
import os
import PyPDF2
import zipfile

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

# Function to extract names, emails, and phone numbers from text
def extract_info(text):
    # Extract email addresses using regex
    emails = re.findall(r'\S+@\S+', text)

    # Extract phone numbers using regex
    phone_numbers = re.findall(r'\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b', text)

    # Assuming names are in the format "First Last"
    names = re.findall(r'\b[A-Z][a-z]+\s[A-Z][a-z]+\b', text)

    return names, emails, phone_numbers

# Function to process a file (PDF or ZIP containing PDFs)
def process_file(file_path):
    extracted_data = {
        'Names': [],
        'Emails': [],
        'Phone Numbers': []
    }

    if file_path.lower().endswith('.pdf'):
        # Extract text from the PDF
        pdf_text = extract_text_from_pdf(file_path)
        # Extract information from the text
        names, emails, phone_numbers = extract_info(pdf_text)
        extracted_data['Names'] = names
        extracted_data['Emails'] = emails
        extracted_data['Phone Numbers'] = phone_numbers

    elif file_path.lower().endswith('.zip'):
        # Extract text from all PDFs in the ZIP archive
        try:
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                for file_info in zip_ref.infolist():
                    if file_info.filename.lower().endswith('.pdf'):
                        pdf_text = extract_text_from_pdf(zip_ref.open(file_info.filename))
                        # Extract information from the text
                        names, emails, phone_numbers = extract_info(pdf_text)
                        extracted_data['Names'].extend(names)
                        extracted_data['Emails'].extend(emails)
                        extracted_data['Phone Numbers'].extend(phone_numbers)
        except Exception as e:
            print(f"Error extracting PDFs from ZIP: {str(e)}")

    return extracted_data

# Main program
if __name__ == '__main__':
    file_path = input("C:\Users\rajes\OneDrive\Desktop\CV's for assignment: ")
    if os.path.exists(file_path):
        extracted_data = process_file(file_path)

        print("Extracted Names:")
        for name in extracted_data['Names']:
            print(name)

        print("\nExtracted Emails:")
        for email in extracted_data['Emails']:
            print(email)

        print("\nExtracted Phone Numbers:")
        for phone_number in extracted_data['Phone Numbers']:
            print(phone_number)
    else:
        print("File not found.")
