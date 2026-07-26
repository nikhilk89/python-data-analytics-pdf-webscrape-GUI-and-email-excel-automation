import os
import pandas as pd
from pypdf import PdfReader, PdfWriter

#Load the Excel data
excel_file = r"ddfields.xlsx"
df = pd.read_excel(excel_file)
df = df.fillna("") # Fill missing/NaN values with empty blanks, pandas sees blank in excel ans Na
df['Date'] = df['Date'].dt.strftime('%m/%d/%Y')  # Convert the column to standard date format (MM/DD/YYYY) right in pandas

#fillable pdf forms
pdf_template_path = "chase-dd.pdf"
output_folder = "chase-dd filled forms8"

os.makedirs(output_folder, exist_ok=True)

#enter info
for index, row in df.iterrows():
    client_name = str(row['Customer name']).strip()


    if not client_name or client_name.lower() == '':
        continue

    print(f"Processing form for: {client_name}...")

    # Load the template PDF for writing
    reader = PdfReader(pdf_template_path)
    writer = PdfWriter()
    writer.append(reader)

    field_mapping = {
        "Customer name": str(row['Customer name']),
        "Address": str(row['Address']),
        "City": str(row['City']),
        "State": str(row['State']),
        "ZIP code": str(row['ZIP code']),
        "Name of business": str(row['Name of business']),
        "Bank's Routing Number": str(row["Bank's Routing Number"]),
        "Date": str(row['Date']),
        "Checking Account Number": str(row['Checking Account Number']),
        "Savings Account Number": str(row['Savings Account Number']),
    }

#radio button tick with checkings/savings check
    account_type = str(row['Checking or Savings']).lower()
    if 'checking' in account_type:
        field_mapping["Checking or Savings"] = "/Checking Account"
    elif 'savings' in account_type:
        field_mapping["Checking or Savings"] = "/Savings Account"


    writer.update_page_form_field_values(writer.pages[0], field_mapping)

    #new name
    safe_client_name = "".join(c for c in client_name if c.isalnum() or c in (' ', '_', '-')).rstrip()
    output_filename = os.path.join(output_folder, f"Direct_Deposit_{safe_client_name}.pdf")

    with open(output_filename, "wb") as output_file:
        writer.write(output_file)

print(f"\n All client forms generated successfully in the '{output_folder}' folder!")