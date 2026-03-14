import pytesseract
from pdf2image import convert_from_path
import cv2
import numpy as np
import re
import json
import pandas as pd
import time

# -------------------------------------------------
#  TESSERACT PATH
# -------------------------------------------------
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


# -------------------------------------------------
#  PDF → IMAGES
# -------------------------------------------------
def pdf_to_images(pdf_path, dpi=300):
    return convert_from_path(pdf_path, dpi=dpi)


# -------------------------------------------------
#  IMAGE PREPROCESSING
# -------------------------------------------------
def preprocess_image(pil_image):
    img = np.array(pil_image.convert("RGB"))
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    gray = cv2.medianBlur(gray, 3)
    thresh = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY, 11, 2
    )
    return thresh


# -------------------------------------------------
#  CROP TOTAL REGION (BOTTOM-RIGHT)
# -------------------------------------------------
def crop_total_region(image):
    h, w = image.shape[:2]
    return image[int(h * 0.70):int(h * 0.98), int(w * 0.55):int(w * 0.98)]


# -------------------------------------------------
#  OCR CLEANER (Fix OCR Mistakes)
# -------------------------------------------------
def clean_ocr_amount(val):
    return (
        val.upper()
           .replace("S", "5")
           .replace("O", "0")
           .replace(",", "")
           .replace(" ", "")
    )


# -------------------------------------------------
#  TEXT OCR
# -------------------------------------------------
def extract_text(image):
    return pytesseract.image_to_string(image, lang='eng')


# -------------------------------------------------
#  EXTRACT BASIC FIELDS
# -------------------------------------------------
def extract_fields(text):
    data = {}

    def find(patterns):
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
            if match:
                val = match.group(1).strip()
                return re.sub(r'\s+', ' ', val)
        return "Not found"

    data["invoice_number"] = find([
        r'Invoice\s*No\.?\s*([A-Z0-9\-\/]+)',
        r'BILL\s*NO\.?[\s:\-]*\n?([A-Z0-9\-\/]+)'
    ])

    data["date"] = find([
        r'Dated[\s:\n]*([0-9]{1,2}[-/\.][A-Za-z]{3,9}[-/\.]?[0-9]{2,4})',
        r'\b([0-9]{1,2}[-/\.][A-Za-z]{3,9}[-/\.]?[0-9]{2,4})\b'
    ])

    data["seller"] = find([
        r'Tax Invoice\s*\n*([A-Z][A-Z\s&\.]+)'
    ])

    data["buyer"] = find([
        r'Buyer\s*\(Bill to\)\s*\n*([A-Z][A-Z\s&\.]+)'
    ])

    return data


# -------------------------------------------------
#  TOTAL + TAX EXTRACTION (Uses Cropped Region)
# -------------------------------------------------
def extract_amounts(full_text, total_region_text):

    # ---------- TAX ----------
    tax_match = re.search(
        r'(IGST|CGST|SGST)[^\d]+([0-9,]+\.\d{2})',
        full_text,
        re.IGNORECASE
    )
    tax_amount = clean_ocr_amount(tax_match.group(2)) if tax_match else "Not found"

    # ---------- TOTAL (bottom-right region only) ----------
    total_match = re.search(
        r'([0-9][0-9,]*\.\d{2})',
        total_region_text
    )

    total_amount = clean_ocr_amount(total_match.group(1)) if total_match else "Not found"

    return {
        "tax_amount": tax_amount,
        "total_amount": total_amount
    }


# -------------------------------------------------
#  PRODUCT EXTRACTION (WORKING)
# -------------------------------------------------
def extract_products(text):
    lines = text.split("\n")
    products = []

    product_pattern = re.compile(
        r'^\s*(\d+)\s*\|\s*'                   # SL No.
        r'([A-Za-z0-9\"\'\s\-/\(\)\.]+?)\s+'   # Description
        r'([0-9]{8})\s+'                       # HSN
        r'([0-9,]+)\s*nos\s+'                  # Quantity
        r'([0-9,]+\.\d{2})\s*nos\s+'           # Rate
        r'([0-9,]+\.\d{2})',                   # Amount
        re.IGNORECASE
    )

    for line in lines:
        match = product_pattern.match(line.strip())
        if match:
            products.append({
                "sl_no": match.group(1),
                "description": match.group(2).strip(),
                "hsn_sac": match.group(3),
                "quantity": match.group(4),
                "rate": match.group(5).replace(",", ""),
                "amount": match.group(6).replace(",", "")
            })

    return products


# -------------------------------------------------
#  MAIN PIPELINE
# -------------------------------------------------
def process_invoice(pdf_path, save_text=False):
    images = pdf_to_images(pdf_path)
    full_text = ""
    total_region_text = ""

    for img in images:
        pre_img = preprocess_image(img)

        # FULL PAGE OCR
        text = extract_text(pre_img)
        full_text += text + "\n"

        # TOTAL REGION OCR
        total_region = crop_total_region(pre_img)
        total_text = extract_text(total_region)
        total_region_text += total_text + "\n"

    # Save OCR text (optional)
    if save_text:
        with open("ocr_full.txt", "w", encoding="utf-8") as f:
            f.write(full_text)
        with open("ocr_total_region.txt", "w", encoding="utf-8") as f:
            f.write(total_region_text)

    print("---- TOTAL REGION OCR ----")
    print(total_region_text)
    print("--------------------------")

    # Extract fields
    fields = extract_fields(full_text)

    # Extract TOTAL + TAX
    amounts = extract_amounts(full_text, total_region_text)
    fields.update(amounts)

    # Product details
    fields["products"] = extract_products(full_text)

    return fields


# -------------------------------------------------
#  SAVE JSON
# -------------------------------------------------
def save_combined_json(data, output_file="invoice_output.json"):
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)
    print(f"✅ JSON saved: {output_file}")


# -------------------------------------------------
#  SAVE EXCEL
# -------------------------------------------------
def save_to_excel(data, output_file=None):
    if output_file is None:
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        output_file = f"invoice_output_{timestamp}.xlsx"

    fields_df = pd.DataFrame([{k: v for k, v in data.items() if k != 'products'}])
    products_df = pd.DataFrame(data.get("products", []))

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        fields_df.to_excel(writer, sheet_name="Invoice Info", index=False)
        products_df.to_excel(writer, sheet_name="Products", index=False)

    print(f"📊 Excel saved: {output_file}")


# -------------------------------------------------
#  MAIN EXECUTION
# -------------------------------------------------
if __name__ == "__main__":
    pdf_path = "116.pdf"
    result = process_invoice(pdf_path, save_text=True)

    print(json.dumps(result, indent=4))
    save_combined_json(result)
    save_to_excel(result)
