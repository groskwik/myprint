#!/usr/bin/env python
#!/usr/bin/env python
import os
import subprocess

# Path to SumatraPDF executable
SUMATRA_PATH = r"C:\portableapps\sumatrapdf\sumatrapdf.exe"

# Folder where PDFs are stored
PDF_FOLDER = r"C:\Users\benoi\Downloads\ebay_manuals"

# Default printer name
PRINTER_NAME = "Brother HL-L8360CDW [Wireless]"

# Predefined print settings for specific PDFs (keys stored in lowercase)
PRINT_SETTINGS = {
    "singer 3337": [
        "color,1,simplex,fit,paper=letter",
        "monochrome,2,duplex,fit,paper=letter",
        "monochrome,3-34,duplex,fit,paper=letter",
        "color,102,simplex,fit,paper=letter",
    ],
    "bernette b05": [
        "color,1,82,duplex,fit,paper=letter",
    ],
    "canon sx530 hs": [
        "color,1,169,duplexshort,fit,paper=letter,landscape",
    ],
    "Brother Innov-is xp3 embrodery": [
        "color,1,212,duplex,fit,paper=letter",
    ],
    "Humminbird HELIX 5": [
        "color,1-190,duplex,fit,paper=letter",
        "monochrome,191-214,duplex,fit,paper=letter",
        "color,215,simplex,fit,paper=letter",
    ],
    "Leica Q3 43": [
        "color,1,264,duplex,fit,paper=letter",
    ],
    "Canon EOS M50 Mark II": [
        "color,1,709,duplex,fit,paper=letter",
    ],
    "othermanual": [
        "monochrome,1-5,simplex,fit,paper=letter"
    ]
}
def find_pdf(partial_name):
    """Finds a PDF file in the folder that contains the given string (case insensitive)."""
    partial_name_lower = partial_name.lower()

    matching_files = [
        f for f in os.listdir(PDF_FOLDER)
        if f.lower().endswith(".pdf") and partial_name_lower in f.lower()
    ]

    if not matching_files:
        print(f"No PDF found containing: {partial_name}")
        return None
    
    if len(matching_files) > 1:
        print("Multiple matches found:")
        for idx, file in enumerate(matching_files, start=1):
            print(f"{idx}. {file}")
        choice = input("Enter the number of the file you want to print: ").strip()
        if not choice.isdigit() or int(choice) < 1 or int(choice) > len(matching_files):
            print("Invalid choice.")
            return None
        return matching_files[int(choice) - 1]

    return matching_files[0]  # Return the only match

def print_pdf(partial_name):
    """Prints a PDF with predefined settings or default settings if not found."""
    pdf_file = find_pdf(partial_name)
    if not pdf_file:
        return

    pdf_path = os.path.join(PDF_FOLDER, pdf_file)
    file_name_without_ext = os.path.splitext(pdf_file)[0].lower()

    # Get print settings for this file or use default
    print_settings = PRINT_SETTINGS.get(file_name_without_ext, ["color,fit,paper=letter"])

    print(f"Printing: {pdf_file} ...")
    
    for setting in print_settings:
        print(f"Applying print settings: {setting}")
        subprocess.run([SUMATRA_PATH, "-print-to", PRINTER_NAME, "-print-settings", setting, pdf_path], check=True)

    print("Printing completed!")

if __name__ == "__main__":
    file_name = input("Enter part of the PDF filename: ").strip()
    print_pdf(file_name)

