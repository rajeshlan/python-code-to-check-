import os
from pdf2docx import Converter

# Define a function to convert PDF to DOCX
def convert_pdf_to_docx(pdf_file_path, output_dir, font_dir=None):
    # Create a Converter instance
    cv = Converter(pdf_file_path)

    # Specify the font directory if provided
    if font_dir:
        cv.fallback_fonts.append(font_dir)

    # Convert the PDF to DOCX and save in the output directory
    cv.convert(output_dir, start=0, end=None)
    cv.close()

# Main function to convert PDFs to DOCX
def main(input_folder, output_folder, font_dir=None):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)  # Create the output directory if it doesn't exist

    pdf_files = [f for f in os.listdir(input_folder) if f.endswith('.pdf')]

    for pdf_file in pdf_files:
        pdf_file_path = os.path.join(input_folder, pdf_file)
        docx_file_path = os.path.join(output_folder, os.path.splitext(pdf_file)[0] + '.docx')

        convert_pdf_to_docx(pdf_file_path, docx_file_path, font_dir)

if __name__ == "__main__":
    input_folder = r"C:\Users\rajes\OneDrive\Desktop\CV (1)"
    output_folder = r"C:\Users\rajes\OneDrive\Desktop\CV (1)\converted"
    font_dir = r"C:\Windows\Fonts"  # Specify the path to the directory containing fonts

    main(input_folder, output_folder, font_dir)
