import json
import os
import pymupdf
from PIL import Image
import pytesseract

pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Program Files\Tesseract-OCR\tesseract.exe"
)

def extract_text(pdf_path: str) -> dict:
    # Open PDF document
    pdf_doc = pymupdf.open(pdf_path)
    doc_data = {
        "filename": pdf_path,
        "total_pages": len(pdf_doc),
        "pages": [],
    }

    for page_num in range(len(pdf_doc)):
        page = pdf_doc[page_num]

        # Render page into image with 300 DPI
        pix = page.get_pixmap(dpi=300)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Run OCR on image
        page_txt = pytesseract.image_to_string(img, config="--psm 6")


        doc_data["pages"].append(
            {"page_number": page_num + 1, "text": page_txt.strip()}
        )

    return doc_data


if __name__ == "__main__":
    file_path = "scansmpl4ocr.pdf"

    # Extract OCR data
    data = extract_text(file_path)

    # Ensure output directory exists
    output_dir = "static"
    os.makedirs(output_dir, exist_ok=True)

    # Save to static/ocr.json
    output_file = os.path.join(output_dir, "ocr.json")
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

    print(f"Extraction complete! Saved to {output_file}")