import pandas as pd
from docx import Document
import os

# Folder path where your Excel files are stored
folder_path = r"C:/Users/Shreyash Bhardwaj/Desktop/Excel_Extraction"

# List of Excel files
excel_files = [
    # "asansol.xlt.xls",
    "Bajaj AMC Training Enrollment Form - NISM 13 Common Derivative  (Responses).xlsx",
    # "Bankura.xlsx",
    # "Bardhaman.xlt.xls",
    # "Durgapur.xlsx",
    # "Howrah.xlt.xls",
    # "Kharagpur.xls",
    # "Siliguri.xls"
]

# Process each Excel file
for file in excel_files:
    try:
        ext = os.path.splitext(file)[1].lower()
        engine = 'openpyxl' if ext == '.xlsx' else 'xlrd'

        # Construct full file path
        file_path = os.path.join(folder_path, file)

        # Read the Excel file
        df = pd.read_excel(file_path, engine=engine)

        # Extract Column F (index 5)
        column_data = df.iloc[:, 5].dropna().astype(str).tolist()
        comma_separated = ', '.join(column_data)

        # Create and save the Word document
        doc = Document()
        doc.add_heading(f'Data from Column F - {file}', level=1)
        doc.add_paragraph(comma_separated)

        output_name = os.path.splitext(file)[0] + ".docx"
        output_path = os.path.join(folder_path, output_name)

        doc.save(output_path)
        print(f"✅ Saved: {output_name}")

    except Exception as e:
        print(f"⚠️ Failed to process {file}: {e}")
