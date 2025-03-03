#!/usr/bin/env python
import os
import subprocess
import time
from PyPDF2 import PdfReader

# Path to SumatraPDF executable
SUMATRA_PATH = r"C:\portableapps\sumatrapdf\sumatrapdf.exe"

# Folders where PDFs are stored
PDF_FOLDERS = [
    r"C:\Users\benoi\Downloads\ebay_manuals",
    r"C:\Users\benoi\Downloads\manuals"
]

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
    "otari mtr-10 series": [
        "monochrome,1-180,duplex,fit,paper=letter",
        "color,181,188,duplex,fit,paper=letter",
        "monochrome,189-290,duplex,fit,paper=letter",
    ],
    "singer 2263": [
        "color,1,simplex,fit,paper=letter",
        "monochrome,2,duplex,fit,paper=letter",
        "monochrome,3-62,duplex,fit,paper=letter",
    ],
    "bernette b05": [
        "color,1-82,duplex,fit,paper=letter",
    ],
    "canon sx530 hs": [
        "color,1-169,duplexshort,fit,paper=letter,landscape",
    ],
    "brother innov-is xp3 embroidery": [
        "color,1-212,duplex,fit,paper=letter",
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
    "bernina 570qe": [
        "color,1-214,duplex,fit,paper=letter",
    ],
    "hp16c-oh-en_mod": [
        "color,1-140,duplex,fit,paper=letter",
    ],
    "hp15c-oh-en": [
        "color,1-300,duplex,fit,paper=letter",
    ],
    "hp42s-elec-en": [
        "monochrome,3-164,duplex,fit,paper=letter",
    ],
    "hp42s-elec-en_mod": [
        "color,1,simplex,fit,paper=letter",
        "monochrome,3-164,duplex,fit,paper=letter",
    ],
    "gopro hero 9 black camera_mod": [
        "color,3-145,duplex,fit,paper=letter",
    ],
    "hp42s-easycourse": [
        "monochrome,3-388,duplex,fit,paper=letter",
    ],
    "hp42s-easycourse_mod": [
        "color,1,simplex,fit,paper=letter",
        "monochrome,3-388,duplex,fit,paper=letter",
    ],
    "nikon d3500": [
        "monochrome,1-36,duplex,fit,paper=letter",
    ],
    "nikon d3300": [
        "monochrome,1-144,duplex,fit,paper=letter",
    ],
    "nikon d6": [
        "monochrome,1-316,duplex,fit,paper=letter",
    ],
    "hp71-rpn": [
        "monochrome,1-12,duplex,fit,paper=letter",
    ],
    "canon eos 6d": [
        "color,1-404,duplex,fit,paper=letter",
    ],
    "dm42": [
        "color,2-29,duplex,fit,paper=letter",
    ],
    "dm41x_user_manual": [
        "color,3-44,duplex,fit,paper=letter",
    ],
    "brother ce8080 ce8080prw_mod": [
        "color,1-72,duplex,fit,paper=letter",
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
    """Finds a PDF file in the specified folders that contains the given string (case insensitive)."""
    partial_name_lower = partial_name.lower()

    matching_files = []
    for folder in PDF_FOLDERS:
        for f in os.listdir(folder):
            if f.lower().endswith(".pdf") and partial_name_lower in f.lower():
                matching_files.append(os.path.join(folder, f))

    if not matching_files:
        print(f"No PDF found containing: {partial_name}")
        return None
    
    if len(matching_files) > 1:
        print("Multiple matches found:")
        for idx, file in enumerate(matching_files, start=1):
            print(f"{idx}. {os.path.basename(file)}")
        choice = input("Enter the number of the file you want to print: ").strip()
        if not choice.isdigit() or int(choice) < 1 or int(choice) > len(matching_files):
            print("Invalid choice.")
            return None
        return matching_files[int(choice) - 1]

    return matching_files[0]  # Return the only match

def get_pdf_page_count(pdf_path):
    """Returns the number of pages in the given PDF file."""
    try:
        with open(pdf_path, "rb") as f:
            reader = PdfReader(f)
            return len(reader.pages)
    except Exception as e:
        print(f"Error reading PDF: {e}")
        return None

def print_pdf(printer_name, partial_name):
    """Prints a PDF with predefined settings or user-defined page ranges."""
    pdf_path = find_pdf(partial_name)
    if not pdf_path:
        return

    # Display the number of pages in the selected PDF
    page_count = get_pdf_page_count(pdf_path)
    if page_count:
        print(f"The document '{os.path.basename(pdf_path)}' has {page_count} pages.")
    else:
        print("Unable to determine the number of pages.")

    file_name_without_ext = os.path.splitext(os.path.basename(pdf_path))[0].lower()
    
    # Get custom page range input
    custom_range = input("Enter the page range to print (e.g., 40-45), or press Enter for default: ").strip()
    
    # Get predefined print settings or use default
    print_settings = PRINT_SETTINGS.get(file_name_without_ext, ["color,fit,paper=letter"])
    
    if custom_range:
        new_print_settings = []
        for setting in print_settings:
            setting_parts = setting.split(",")
            modified_setting = []
            replaced = False
            
            for part in setting_parts:
                if "-" in part and part.replace("-", "").isdigit():
                    modified_setting.append(custom_range)
                    replaced = True
                elif part.isdigit():
                    continue  # Remove existing standalone numbers
                else:
                    modified_setting.append(part)
            
            if not replaced:
                modified_setting.insert(1, custom_range)  # Ensure correct placement of page range
            
            new_print_settings.append(",".join(modified_setting))
        print_settings = new_print_settings

    print(f"Printing: {os.path.basename(pdf_path)} on {printer_name}...")

    batch_size = 70  # Number of pages per batch
    delay_between_batches = 180  # Delay in seconds between batches

    for setting in print_settings:
        # Extract page range from the setting
        setting_parts = setting.split(",")
        page_range = None
        for part in setting_parts:
            if "-" in part and part.replace("-", "").isdigit():
                page_range = part
                break

        if page_range:
            start_page, end_page = map(int, page_range.split("-"))
            current_page = start_page

            while current_page <= end_page:
                batch_end = min(current_page + batch_size - 1, end_page)
                batch_range = f"{current_page}-{batch_end}"
                batch_setting = setting.replace(page_range, batch_range)

                print(f"Printing pages {batch_range} with settings: {batch_setting}")
                subprocess.run([SUMATRA_PATH, "-print-to", printer_name, "-print-settings", batch_setting, pdf_path], check=True)

                current_page += batch_size
                if current_page <= end_page:
                    print(f"Waiting for {delay_between_batches // 60} minutes before next batch...")
                    time.sleep(delay_between_batches)
        else:
            # Print normally if no page range is specified
            print(f"Applying print settings: {setting}")
            subprocess.run([SUMATRA_PATH, "-print-to", printer_name, "-print-settings", setting, pdf_path], check=True)

    print("Printing completed!")

if __name__ == "__main__":
    selected_printer = select_printer()
    file_name = input("Enter part of the PDF filename: ").strip()
    print_pdf(selected_printer, file_name)
