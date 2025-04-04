from PIL import Image
import pytesseract
from openpyxl import Workbook

# Optional: Set path to tesseract executable if not in PATH
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def extract_text_from_image(image_path):
    image = Image.open(image_path)
    extracted_text = pytesseract.image_to_string(image)
    return extracted_text

def write_text_to_excel(text, output_excel_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "OCR Output"

    # Split text by lines and write to Excel
    for row_idx, line in enumerate(text.splitlines(), start=1):
        ws.cell(row=row_idx, column=1, value=line)

    wb.save(output_excel_path)

# Example usage
image_path = "3EA9C631-1420-4132-84B6-3D1E0E02433F.png"
excel_path = "output_text.xlsx"

text = extract_text_from_image(image_path)
write_text_to_excel(text, excel_path)

print("Text extracted and saved to Excel.")
