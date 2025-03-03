#!/usr/bin/env python
import os
import subprocess

# Path to SumatraPDF executable
SUMATRA_PATH = r"C:\portableapps\sumatrapdf\sumatrapdf.exe"

# Folder where PDFs are stored
PDF_FOLDER = r"C:\Users\benoi\Downloads\ebay_manuals"

# Default printer name
PRINTER_NAME = "Brother HL-L8360CDW [Wireless]"

# Predefined print settings for specific PDFs
PRINT_SETTINGS = {
    "Singer 3337": [
        "color,1,simplex,fit,paper=letter",
        "monochrome,2,duplex,fit,paper=letter",
        "monochrome,3-34,duplex,fit,paper=letter",
        "color,102,simplex,fit,paper=letter",
    ],
    "Bernette b05": [
        "color,1,82,duplex,fit,paper=letter",
    ],
    "OtherManual": [
        "monochrome,1-5,simplex,fit,paper=letter"
    ]
}

def print_pdf(file_name):
    """Prints a PDF with predefined settings or default settings if not found."""
    pdf_path = os.path.join(PDF_FOLDER, f"{file_name}.pdf")

    if not os.path.exists(pdf_path):
        print(f"❌ File not found: {pdf_path}")
        return

    # Get print settings for this file or use default
    print_settings = PRINT_SETTINGS.get(file_name, ["color,fit,paper=letter"])

    print(f"🖨 Printing: {file_name}.pdf ...")
    
    for setting in print_settings:
        print(f"   ➜ Applying print settings: {setting}")
        subprocess.run([SUMATRA_PATH, "-print-to", PRINTER_NAME, "-print-settings", setting, pdf_path], check=True)

    print("✅ Printing completed!")

if __name__ == "__main__":
    file_name = input("Enter the PDF filename (without .pdf): ").strip()
    print_pdf(file_name)

