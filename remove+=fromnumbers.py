import openpyxl

# Open the Excel file
file_path = r'C:\Users\rajes\OneDrive\Desktop\errornumbers.xlsx'
workbook = openpyxl.load_workbook(file_path)

# Specify the sheet name where you want to remove the formulas
sheet_name = 'errornumbers'  # Replace with your sheet name
worksheet = workbook[sheet_name]

# Loop through the cells in the specified sheet
for row in worksheet.iter_rows(values_only=True):
    for cell in row:
        if cell is not None and isinstance(cell, str) and cell.startswith('=+'):
            # Remove the "=+" prefix and assign the numeric value
            cell_value = cell[2:]  # Remove the first two characters (=+)
            cell.value = cell_value

# Save the modified Excel file
output_file_path = 'modified_excel_file.xlsx'
workbook.save(output_file_path)

# Close the workbook
workbook.close()

print(f"Modified Excel file saved as '{output_file_path}'")
