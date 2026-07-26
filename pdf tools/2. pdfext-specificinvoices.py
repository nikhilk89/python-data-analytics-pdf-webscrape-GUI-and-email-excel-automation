import os
import re
import pandas as pd
import pdfplumber


#functions code for vendors
def v1_bioplex(text, file_name):
    inv_num = re.search(
        r"INVOICE\s*#?\s*([A-Za-z0-9-]+)", text, re.IGNORECASE
    )
    inv_date = re.search(r"DATE:\s*([\d{1,2}\.-]+)", text, re.IGNORECASE)
    total_due = re.search(
        r"TOTAL DUE\s*[\n|]*\s*([\d,]+\.\d{2})", text, re.IGNORECASE
    )

    invoice_id = inv_num.group(1) if inv_num else file_name

    summary = {
        "File Name": file_name,
        "Vendor": "Bioplex",
        "Invoice No": invoice_id,
        "Date": inv_date.group(1) if inv_date else "",
        "Total Amount": total_due.group(1) if total_due else "",
    }

    line_items = []
    # Pattern to match line items (qquantity,description,Unit price,total)
    pattern = r"(\d+)\s+([A-Za-z0-9\s\(\)\/,-]+?)\s+(\d+\.\d{2})\s+(\d+\.\d{2})"
    matches = re.findall(pattern, text)

    for qty, desc, price, total in matches:
        if not any(
            x in desc.upper() for x in ["SUBTOTAL", "TOTAL", "SALES TAX"]
        ):
            line_items.append(
                {
                    "File Name": file_name,
                    "Invoice No": invoice_id,
                    "Quantity": qty.strip(),
                    "Description": desc.strip().replace("\n", " "),
                    "Unit Price": price.strip(),
                    "Total": total.strip(),
                }
            )

    return summary, line_items


def v2_zencorp(text, file_name):
    inv_num = re.search(r"Invoice\s+No\.\s*(\d+)", text, re.IGNORECASE)
    inv_date = re.search(r"(\d{2}\.\d{2}\.\d{4})", text)
    total_due = re.search(
        r"Total Due\s*\|\s*([\d\.]+)", text, re.IGNORECASE
    )

    invoice_id = inv_num.group(1) if inv_num else file_name

    summary = {
        "File Name": file_name,
        "Vendor": "Zencorporations",
        "Invoice No": invoice_id,
        "Date": inv_date.group(1) if inv_date else "",
        "Total Amount": total_due.group(1) if total_due else "",
    }

    line_items = []
    pattern = r"([A-Za-z0-9\s'’-]+?)\s*Qty\.\s*(\d+)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})"
    matches = re.findall(pattern, text)

    for desc, qty, price, total in matches:
        line_items.append(
            {
                "File Name": file_name,
                "Invoice No": invoice_id,
                "Quantity": qty.strip(),
                "Description": desc.strip().replace("\n", " "),
                "Unit Price": price.strip(),
                "Total": total.strip(),
            }
        )

    return summary, line_items


def v3_conincorp(text, file_name):
    inv_num = re.search(r"INVOICE\s+(\d+)", text, re.IGNORECASE)
    inv_date = re.search(r"Date:\s*([\d{2}\.\d{2}\.\d{4}]+)", text)

    invoice_id = inv_num.group(1) if inv_num else file_name

    line_items = []
    total_amount = 0.0

    pattern = r"\d+\s+(\d+)\s+([0-9-]+\s+-\s+[A-Za-z0-9\s\(\)]+?)\s+(\d+\.\d{2})\s+(\d+\.\d{2})"
    matches = re.findall(pattern, text)

    for qty, desc, price, total in matches:
        total_amount += float(total)
        line_items.append(
            {
                "File Name": file_name,
                "Invoice No": invoice_id,
                "Quantity": qty.strip(),
                "Description": desc.strip().replace("\n", " "),
                "Unit Price": price.strip(),
                "Total": total.strip(),
            }
        )

    summary = {
        "File Name": file_name,
        "Vendor": "ConIncorporated",
        "Invoice No": invoice_id,
        "Date": inv_date.group(1) if inv_date else "",
        "Total Amount": f"{total_amount:.2f}",
    }

    return summary, line_items



#main proagram
folder_path = r"invoices\\dvendor"

all_summaries = []
all_line_items = []

if os.path.exists(folder_path):
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, file_name)
            print(f"Reading: {file_name}")

            full_text = ""
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        full_text += text + "\n"

            #select based on vendor
            if "bioplex" in full_text.lower():
                summary, items = v1_bioplex(full_text, file_name)
            elif "zencorporations" in full_text.lower():
                summary, items = v2_zencorp(full_text, file_name)
            elif "conincorporated" in full_text.lower():
                summary, items = v3_conincorp(full_text, file_name)
            else:
                print(f"Unknown Vendor Layout in {file_name}. Skipping.")
                continue

            all_summaries.append(summary)
            all_line_items.extend(items)

    #save to excel
    output_excel = "processed_vendor_invoices.xlsx"
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        pd.DataFrame(all_summaries).to_excel(
            writer, sheet_name="Invoice Summaries", index=False
        )
        pd.DataFrame(all_line_items).to_excel(
            writer, sheet_name="Line Items", index=False
        )

    print(f"\nExtraction Complete! saved '{output_excel}' successfully.")
else:
    print(
        f"Directory '{folder_path}' not found. Please create the folder to add your PDFs."
    )