#!/usr/bin/env python
import os
import subprocess
import time

# Path to SumatraPDF executable
SUMATRA_PATH = r"C:\portableapps\sumatrapdf\sumatrapdf.exe"

# Folder where PDFs are stored
PDF_FOLDER = r"C:\Users\benoi\Downloads\ebay_manuals"
PDF_FOLDER = r"C:\Users\benoi\Downloads\manuals"

# Available printers
PRINTERS = {
    "1": "Brother HL-L8360CDW [Wireless]",
    "2": "Brother HL-L3290CDW [Wireless]",
    "3": "Brother HL-L8360CDW [Landscape]"
}

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
    "brother innov-is xp3 embrodery": [
        "color,1,212,duplex,fit,paper=letter",
    ],
    "humminbird helix 5": [
        "color,1-190,duplex,fit,paper=letter",
        "monochrome,191-214,duplex,fit,paper=letter",
        "color,215,simplex,fit,paper=letter",
    ],
    "leica q3 43": [
        "color,1-264,duplex,fit,paper=letter",
    ],
    "canon eos m50 mark ii": [
        "color,1-709,duplex,fit,paper=letter",
    ],
    "hp16c-oh-en_mod": [
        "color,1-140,duplex,fit,paper=letter",
    ],
    "othermanual": [
        "monochrome,1-5,simplex,fit,paper=letter"
    ]
}

def select_printer():
    """Prompts the user to select a printer."""
    print("Select a printer:")
    for key, name in PRINTERS.items():
        print(f"{key}. {name}")
    
    choice = input("Enter the number of the printer: ").strip()
    if choice not in PRINTERS:
        print("Invalid choice. Defaulting to Brother HL-L8360CDW [Wireless].")
        return PRINTERS["1"]
    return PRINTERS[choice]

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

def print_pdf(printer_name, partial_name):
    """Prints a PDF with predefined settings or default settings if not found."""
    pdf_file = find_pdf(partial_name)
    if not pdf_file:
        return

    pdf_path = os.path.join(PDF_FOLDER, pdf_file)
    file_name_without_ext = os.path.splitext(pdf_file)[0].lower()

    # Get print settings for this file or use default
    print_settings = PRINT_SETTINGS.get(file_name_without_ext, ["color,fit,paper=letter"])

    print(f"Printing: {pdf_file} on {printer_name}...")

    for setting in print_settings:
        # Split settings to check for a page range
        setting_parts = setting.split(",")
        pages = None
        for part in setting_parts:
            if "-" in part and all(p.isdigit() for p in part.split("-")):
                pages = part
                break

        if pages:
            # Extract start and end page
            start_page, end_page = map(int, pages.split("-"))
            batch_size = 70
            current_page = start_page

            while current_page <= end_page:
                batch_end = min(current_page + batch_size - 1, end_page)
                batch_setting = setting.replace(pages, f"{current_page}-{batch_end}")

                print(f"Printing batch: {current_page}-{batch_end}")
                subprocess.run([SUMATRA_PATH, "-print-to", printer_name, "-print-settings", batch_setting, pdf_path], check=True)

                current_page += batch_size
                if current_page <= end_page:
                    print("Waiting for 3 minutes before next batch...")
                    time.sleep(180)  # Small delay to prevent printer overload
        else:
            # Print normally if no page range is specified
            print(f"Applying print settings: {setting}")
            subprocess.run([SUMATRA_PATH, "-print-to", printer_name, "-print-settings", setting, pdf_path], check=True)

    print("Printing completed!")

if __name__ == "__main__":
    selected_printer = select_printer()
    file_name = input("Enter part of the PDF filename: ").strip()
    print_pdf(selected_printer, file_name)
