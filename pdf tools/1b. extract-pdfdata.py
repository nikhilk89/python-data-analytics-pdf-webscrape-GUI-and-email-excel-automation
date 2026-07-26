#one time - extract data from form for all required fields
import os
import pandas as pd
from pypdf import PdfReader


def extract_form_fields(pdf_path):
    # Initialize the reader
    reader = PdfReader(pdf_path)

    # Extract interactive AcroForm fields
    fields = reader.get_fields()

    if not fields:
        print("No fillable interactive fields found in this PDF.")
        return {}

    extracted_data = {}

    for field_name, field_info in fields.items():
        # Retrieve value if available (/V key)
        field_value = field_info.get('/V', None)

        # Clean up string wrapper if present
        if hasattr(field_value, 'get_object'):
            field_value = field_value.get_object()

        # Handle None or empty values cleanly
        extracted_data[field_name] = field_value if field_value is not None else ""

    return extracted_data


def save_to_excel(pdf_path, excel_path):
    # 1. Extract data using your function
    form_data = extract_form_fields(pdf_path)

    if not form_data:
        print("Skipping Excel export as no fields were extracted.")
        return

    # 2. Convert dictionary to a pandas DataFrame
    # Wrapping form_data in a list [form_data] creates 1 row where keys are column headers
    df = pd.DataFrame([form_data])

    # 3. Export to Excel
    df.to_excel(excel_path, index=False, engine="openpyxl")
    print(f"Successfully extracted fields and saved to '{excel_path}'")


# --- Run Extraction & Export ---
pdf_file = r"fillformpdf/chase-dd.pdf"
excel_output = r"fillformpdf/ddfields.xlsx"

save_to_excel(pdf_file, excel_output)