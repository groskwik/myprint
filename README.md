# Automated PDF Printing Script

This Python script automates the process of searching for PDF files within specified directories and printing them using predefined or custom settings. It leverages the [SumatraPDF](https://www.sumatrapdfreader.org/docs/Command-line-arguments) application for command-line printing and the [PyPDF2](https://pypdf2.readthedocs.io/en/latest/) library to retrieve PDF metadata.

## Features

- **Printer Selection**: Choose from a list of predefined printers or set a default printer.
- **PDF Search**: Locate PDF files in specified directories based on partial filenames.
- **Predefined Print Settings**: Apply custom print settings for specific PDF files.
- **Custom Page Range Printing**: Specify custom page ranges for printing.
- **Batch Printing**: Print large documents in batches with configurable delays between batches.

## Requirements

- **Python 3.x**
- **[SumatraPDF](https://www.sumatrapdfreader.org/docs/Command-line-arguments)**: Ensure the path to the SumatraPDF executable is correctly set in the script.
- **[PyPDF2](https://pypdf2.readthedocs.io/en/latest/)**: Install using `pip install PyPDF2`.

## Setup

1. **Configure SumatraPDF Path**: Set the `SUMATRA_PATH` variable to the location of your SumatraPDF executable. For example:

   ```python
   SUMATRA_PATH = r"C:\path\to\sumatrapdf.exe"
   ```

2. **Define PDF Directories**: List the directories where your PDF files are stored in the `PDF_FOLDERS` variable. For example:

   ```python
   PDF_FOLDERS = [
       r"C:\Users\YourUsername\Documents\PDFs",
       r"D:\Manuals"
   ]
   ```

3. **Specify Printers**: Add your available printers to the `PRINTERS` dictionary. For example:

   ```python
   PRINTERS = {
       "1": "Printer_Name_1",
       "2": "Printer_Name_2"
   }
   ```

4. **Set Print Settings**: Define any predefined print settings for specific PDFs in the `PRINT_SETTINGS` dictionary. Use the lowercase filename (without extension) as the key. For example:

   ```python
   PRINT_SETTINGS = {
       "example_manual": [
           "color,1-10,duplex,fit,paper=letter"
       ]
   }
   ```

## Usage

1. **Run the Script**: Execute the script in a Python environment.

2. **Select a Printer**: When prompted, choose a printer by entering the corresponding number.

3. **Enter PDF Filename**: Provide a part of the PDF filename you wish to print. The script will search for matching files in the specified directories.

4. **Custom Page Range**: If desired, input a custom page range (e.g., `1-5`). Press Enter to use the default settings.

5. **Printing**: The script will print the selected PDF using the specified settings. For large documents, it will print in batches with pauses between each batch to manage printer load.

## Notes

- **SumatraPDF Command-Line Options**: The script utilizes SumatraPDF's command-line options for printing. Detailed information about these options can be found in the [SumatraPDF Command-Line Arguments documentation](https://www.sumatrapdfreader.org/docs/Command-line-arguments).

- **Error Handling**: Ensure that the specified printer names and PDF paths are correct to avoid errors during execution.

- **Batch Printing**: The script is configured to print documents in batches of 70 pages with a 3-minute delay between batches. Adjust the `batch_size` and `delay_between_batches` variables as needed.

## Disclaimer

This script is provided as-is without warranty of any kind. Ensure you have the necessary permissions to access the specified directories and printers. Use at your own risk.

## License

This project is licensed under the MIT License. 
