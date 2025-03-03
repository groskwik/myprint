    ],
    "canon eos m50 mark ii": [
        "color,1-709,duplex,fit,paper=letter",
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

def get_closest_matching_setting(file_name, custom_range):
    """Finds the most relevant print setting for a given custom page range."""
    settings = PRINT_SETTINGS.get(file_name, [])
    custom_start, custom_end = map(int, custom_range.split("-"))

    best_match = None
    for setting in settings:
        setting_parts = setting.split(",")
        for part in setting_parts:
            if "-" in part and all(p.isdigit() for p in part.split("-")):
                start, end = map(int, part.split("-"))
                if start <= custom_start <= end or start <= custom_end <= end:
                    best_match = setting
                    break

    return best_match or settings[0] if settings else "color,fit,paper=letter"

def print_pdf(printer_name, partial_name):
    """Prints a PDF with predefined settings or a custom range if specified."""
    pdf_file = find_pdf(partial_name)
    if not pdf_file:
        return

    pdf_path = os.path.join(PDF_FOLDER, pdf_file)
    file_name_without_ext = os.path.splitext(pdf_file)[0].lower()

    # Ask for custom page range
    custom_range = input("Enter the page range to print (or press Enter to use defaults): ").strip()

    if custom_range:
        print_setting = get_closest_matching_setting(file_name_without_ext, custom_range)
    else:
        print_setting = PRINT_SETTINGS.get(file_name_without_ext, ["color,fit,paper=letter"])

    print(f"Printing: {pdf_file} on {printer_name}...")

    if isinstance(print_setting, list):
        print_settings = print_setting  # Use predefined settings
    else:
        print_settings = [print_setting]  # Use user-defined setting

    for setting in print_settings:
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

