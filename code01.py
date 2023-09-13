import re
import os
import PyPDF2
import zipfile
import pandas as pd

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

    # Extract names (assuming "First Last" format)
    names = re.findall(r'\b[A-Z][a-z]+\s[A-Z][a-z]+\b', text)

    return names, emails, phone_numbers

# Function to process a file (PDF or ZIP containing PDFs)
def process_file(file_path):
    extracted_data = {
        'Name': [],
        'Email': [],
        'Phone Number': []
    }

    if file_path.lower().endswith('.pdf'):
        # Extract text from the PDF
        pdf_text = extract_text_from_pdf(file_path)
        # Extract information from the text
        names, emails, phone_numbers = extract_info(pdf_text)
        extracted_data['Name'] = names
        extracted_data['Email'] = emails
        extracted_data['Phone Number'] = phone_numbers

    elif file_path.lower().endswith('.zip'):
        # Extract text from all PDFs in the ZIP archive
        try:
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                for file_info in zip_ref.infolist():
                    if file_info.filename.lower().endswith('.pdf'):
                        pdf_text = extract_text_from_pdf(zip_ref.open(file_info.filename))
                        # Extract information from the text
                        names, emails, phone_numbers = extract_info(pdf_text)
                        extracted_data['Name'].extend(names)
                        extracted_data['Email'].extend(emails)
                        extracted_data['Phone Number'].extend(phone_numbers)
        except Exception as e:
            print(f"Error extracting PDFs from ZIP: {str(e)}")

    return extracted_data

# Main program
if __name__ == '__main__':
    file_path = input("Enter the path to the PDF file or ZIP file containing PDFs: ")


    if os.path.exists(file_path):
        extracted_data = process_file(file_path)

        # Create a DataFrame from the extracted data
        df = pd.DataFrame(extracted_data)

        # Save the DataFrame to a spreadsheet (e.g., CSV)
        output_file = "extracted_data.csv"
        df.to_csv(output_file, index=False)

        print(f"Data extracted and saved to '{output_file}'")
    else:
        print("File not found.")
