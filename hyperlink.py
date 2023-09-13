import re
import os
import pandas as pd
import pdfplumber

# Function to extract names, phone numbers, email addresses, city names, and states from a text
def extract_info(text):
    # Regular expressions for extracting names, phone numbers, email addresses, city names, and states
    name_pattern = r'\b[A-Z][a-z]+\s[A-Z][a-z]+\b'
    phone_pattern = r'(\+\d{1,3}\s?)?\(?\d{3}\)?[-.\s]?\d{3,4}[-.\s]?\d{4}'  # Matches phone numbers with or without country code
    email_pattern = r'\b\w+@\w+\.\w+\b'  # Matches email addresses
    city_state_pattern = r'\b([A-Z][a-z]+), ([A-Z]{2})\b'  # Matches "City, State" format

    # Extracting names, phone numbers, email addresses, city names, and states
    names = re.findall(name_pattern, text)
    phone_numbers = re.findall(phone_pattern, text)
    email_addresses = re.findall(email_pattern, text)

    # Extract city names and states
    city_state_matches = re.findall(city_state_pattern, text)
    city_names = [match[0] for match in city_state_matches]
    states = [match[1] for match in city_state_matches]

    return names, phone_numbers, email_addresses, city_names, states

# Function to extract text from a PDF file using pdfplumber
def extract_text_from_pdf(pdf_path):
    pdf_text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                pdf_text += page.extract_text()
    except Exception as e:
        print(f"Error extracting text from PDF '{pdf_path}': {str(e)}")
        pdf_text = ""  # Set an empty string if extraction fails

    return pdf_text

# Define the PDF folder and output folder
pdf_folder = r'C:\Users\rajes\OneDrive\Desktop\CV SEPT112023-20230913T044719Z-001\CV SEPT112023'
output_folder = r'C:\Users\rajes\OneDrive\Desktop\CV for assignment\output'  # Change this to your desired output folder

# Create the output folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Create a list to store the extracted information
extracted_data = []

# Extract text from PDF files and store the extracted information
for filename in os.listdir(pdf_folder):
    if filename.endswith('.pdf'):
        pdf_path = os.path.join(pdf_folder, filename)
        
        # Try to extract text from the PDF
        pdf_text = extract_text_from_pdf(pdf_path)

        # Extract names, phone numbers, email addresses, city names, and states from the extracted text
        names, phone_numbers, email_addresses, city_names, states = extract_info(pdf_text)

        # Append the extracted information to the list
        extracted_data.append({
            'File': filename,
            'Names': ', '.join(names),
            'Phone Numbers': ', '.join(phone_numbers),
            'Email Addresses': ', '.join(email_addresses),
            'City Names': ', '.join(city_names),
            'States': ', '.join(states)
        })

# Create a pandas DataFrame from the extracted data
df = pd.DataFrame(extracted_data)

# Save the DataFrame to a spreadsheet (Excel format) in the output folder
output_excel = os.path.join(output_folder, 'extracted_data.xlsx')
df.to_excel(output_excel, index=False)

# Print the DataFrame (optional)
print(df)
