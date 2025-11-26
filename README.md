# MyPrint — Intelligent PDF Printing Utility

MyPrint is a Python-based command-line tool designed to streamline the printing of manuals, sewing machine guides, camera documentation, and other large PDF files.  
It features fuzzy searching, automated print presets, JSON-based settings management, and robust SumatraPDF integration.

A companion tool, **`manage_print_settings.py`**, allows easy editing of the print settings database with fuzzy search support.

---

## ✨ Features

### 🔍 Smart Fuzzy PDF Search
Enter part of a filename — MyPrint uses fuzzy matching to locate the intended PDF across multiple folders:

- Case-insensitive  
- Handles typos and partial names  
- Shows a ranked match list with numeric selection  

### 🖨 Printer Selection
Choose a printer at runtime from a predefined list or automatically fall back to a default printer.

### 🎛 JSON-Based Print Settings
Each manual can have custom print profiles stored in a shared JSON file:

```
print_settings.json
```

Profiles include color mode, page ranges, duplex modes, orientation, paper size, and scaling.

### 🔄 Interactive Print Settings Editor
Use:

```
python manage_print_settings.py
```

to:

- Search entries fuzzily  
- Add new manuals  
- Modify existing print settings  
- Delete entries  
- Validate the JSON automatically  

### 📚 Batch Printing for Large PDFs
MyPrint can automatically split big PDFs into print-safe chunks (default: 70 pages per batch), adding a delay between batches.

### 🎚 Custom Page Range Override
If no print profile exists—or you want a special range—you can enter a custom page range manually.

### 📄 PDF Metadata Extraction
MyPrint uses PyPDF2 to read:

- Page count  
- Document title  
- Basic metadata  

---

## 📦 Requirements

- **Python 3.10 or newer**
- **SumatraPDF** (for command-line printing)  
  https://www.sumatrapdfreader.org/docs/Command-line-arguments  
- **Python packages:**

```
pip install PyPDF2 rapidfuzz
```

---

## 🧩 Setup

### 1. Set the path to SumatraPDF
Edit in `myprint.py`:

```python
SUMATRA_PATH = r"C:\path\to\sumatrapdf.exe"
```

### 2. Configure your PDF folders
MyPrint recursively scans these directories:

```python
PDF_FOLDERS = [
    r"C:\Users\YourName\Manuals",
    r"D:\PDFs"
]
```

### 3. Configure your available printers
Example:

```python
PRINTERS = {
    "1": "Brother_HL2350DW",
    "2": "HP_OfficeJet_Pro_9015",
    "3": "Kyocera_Color"
}
```

### 4. Create or edit your print settings database
Print settings live in:

```
print_settings.json
```

Example entry:

```json
"nikon d850": [
  "monochrome,1-400,duplex,fit,paper=letter"
]
```

Keys must be lowercase manual names without `.pdf`.

---

## 🚀 Usage

Run:

```
python myprint.py
```

### Step-by-step workflow

1. **Choose a printer**  
   MyPrint lists your configured printers with their IDs.

2. **Search for a PDF**  
   Enter part of the name (e.g., `"nikon 85"`).  
   MyPrint shows a fuzzy-matched list:

   ```
   1) nikon d850.pdf
   2) nikon d810.pdf
   3) nikon d800.pdf
   ```

   Then select the correct number.

3. **Check for predefined print settings**  
   If a match exists in `print_settings.json`, MyPrint applies it automatically.

4. **Optional: Custom page range**  
   If no preset exists, or if you want a one-off override, enter:

   ```
   1-50
   ```

   or press Enter for full-document printing.

5. **Batch printing**  
   For documents larger than the batch size (default = 70 pages), MyPrint prints them in waves:

   ```
   Printing pages 1–70...
   Waiting 180 seconds...
   Printing pages 71–140...
   ```

---

## 🛠 Managing Print Settings

Launch:

```
python manage_print_settings.py
```

This interactive tool allows you to:

- Fuzzy-find a manual name  
- Add a new entry  
- Edit existing settings  
- Remove outdated entries  
- Save the updated JSON database  

---

## ⚠ Notes

- Ensure printer names match exactly what Windows reports.  
- SumatraPDF may fail silently if a printer name is invalid.  
- PDF metadata is extracted using PyPDF2.  
- All print setting keys must be lowercase for consistent matching.

---

## 📄 License

MIT License.
