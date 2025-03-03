#!/usr/bin/env python
#!/usr/bin/env python
import os
import subprocess
import time

# Path to SumatraPDF executable
SUMATRA_PATH = r"C:\portableapps\sumatrapdf\sumatrapdf.exe"

# Folder where PDFs are stored
PDF_FOLDER = r"C:\Users\benoi\Downloads\ebay_manuals"

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
    "brother innov-is xp3 sewing": [
        "color,1,216,duplex,fit,paper=letter",
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
        "color,609-709,duplex,fit,paper=letter",
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
    return PRINTERS.get(choice, PRINTERS["1"])

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
    """Prints a PDF with predefined settings or user-defined page ranges."""
    pdf_file = find_pdf(partial_name)
    if not pdf_file:
        return

    pdf_path = os.path.join(PDF_FOLDER, pdf_file)
    file_name_without_ext = os.path.splitext(pdf_file)[0].lower()
    
    # Get custom page range input
    custom_range = input("Enter the page range to print (e.g., 40-45), or press Enter for default: ").strip()
    
    # Get predefined print settings or use default
    print_settings = PRINT_SETTINGS.get(file_name_without_ext, ["color,fit,paper=letter"])
    
    if custom_range:
        new_print_settings = []
        for setting in print_settings:
            setting_parts = setting.split(",")
            new_setting = []
            replaced = False
            
            for part in setting_parts:
                if "-" in part and all(p.isdigit() for p in part.split("-")):
                    new_setting.append(custom_range)
                    replaced = True
                else:
                    new_setting.append(part)
            
            if not replaced:
                new_setting.insert(1, custom_range)  # Insert page range after color mode
            
            new_print_settings.append(",".join(new_setting))
        print_settings = new_print_settings
    
    print(f"Printing: {pdf_file} on {printer_name}...")

    for setting in print_settings:
        print(f"Applying print settings: {setting}")
        subprocess.run([SUMATRA_PATH, "-print-to", printer_name, "-print-settings", setting, pdf_path], check=True)
    
    print("Printing completed!")

if __name__ == "__main__":
    selected_printer = select_printer()
    file_name = input("Enter part of the PDF filename: ").strip()
    print_pdf(selected_printer, file_name)

