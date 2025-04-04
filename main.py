import os
import sys

from PIL import Image
import pytesseract
from openpyxl import Workbook
from pathlib import Path

def extract_text_from_image(image_path):
    image = Image.open(image_path)
    extracted_text = pytesseract.image_to_string(image)
    return extracted_text

def write_text_to_excel(extracted_text, output_excel_path):
    wb = Workbook()
    ws = wb.active

    # Split text by lines and write to Excel
    for row_idx, line in enumerate(extracted_text.splitlines(), start=1):
        ws.cell(row=row_idx, column=1, value=line)

    wb.save(output_excel_path)

# Define input and output directories
input_dir = Path("in")
output_dir = Path("out")
output_dir.mkdir(exist_ok=True)

# Counter for processed files
processed_files = 0

# Iterate over all image files in the input directory
if not input_dir.exists():
    print(f"Input directory '{input_dir}' does not exist.")
    sys.exit(1)
for image_path in input_dir.glob("*"):
    # Process only files with valid image extensions
    if image_path.suffix.lower() not in [".png", ".jpg", ".jpeg", ".tiff", ".bmp", ".gif"]:
        continue

    text = extract_text_from_image(image_path)
    # Create an Excel file name based on the image file stem
    excel_path = output_dir / f"{image_path.stem}.xlsx"
    write_text_to_excel(text, excel_path)
    processed_files += 1
    print(f'Processed "{image_path.name}" and saved to "{excel_path}".')

print(f"Total files processed: {processed_files}")
os.system("pause")
