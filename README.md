
# PDF to Excel to Access Uploader

This application extracts structured data from PDF reports, lets you review and edit the results in Excel, then uploads approved records into a Microsoft Access database. It is packaged for non-technical users and comes with a simple desktop interface.

## Quick Start for end users

### What you need

- Windows PC with the Microsoft Access Database Engine installed. The bitness must match the executable you run.
- An Access database file `.accdb` or `.mdb` configured in `src/config/mapping.yml`.
- The provided `src` folder structure.

### Folder layout

```bash
src/
  pdf_to_access_app.exe
  config/
    mapping.yml
  PDFs/
    inbox/
    archive/
    rejected/
  output/
    review.xlsx
  logs/
```

### Steps

1. Place PDFs into `input/inbox`.
2. Run `pdf_to_access_app.exe`.
3. Click **Extract to Excel** to create `output/review.xlsx`.
4. Click **Open review workbook**, fix any values, set `_review_status` to `APPROVED` or `REJECTED`.
5. Save and close Excel.
6. Click **Upload to Access**. The app inserts approved rows, moves PDFs to `archive` or `rejected`, and writes a log in `logs/`.
7. Use **Test Access connection** if needed.

## Developer guide

### Prerequisites

- Python 3.11+
- Install dependencies

  ```bash
  pip install -r requirements.txt
  ```

- Ensure the Microsoft Access ODBC driver is installed on the build machine.

### Configuration: `src/config/mapping.yml`

Defines fields to extract, optional de-duplication keys, and Access connection details.

Supported strategies:

- `keyword_right` captures text to the right of a keyword in the same block.
- `keyword_line` captures the remainder of the line after a keyword.
- `regex` extracts with a regular expression.

### Run locally

```bash
python src/pdf_to_access_app.py
```

### Build a single-file executable

```bash
pyinstaller --onefile --noconsole src/pdf_to_access_app.py
# Output at dist/pdf_to_access_app.exe
```

### Code structure

- `src/pdf_to_access_app.py` is the Tkinter desktop app with two phases
  - Extract to Excel using PyMuPDF
  - Upload to Access using pyodbc
- `src/extract_to_excel.py` and `src/upload_to_access.py` are optional CLI scripts for headless use.

### Troubleshooting

- If Access driver errors occur, match 32-bit or 64-bit driver to the executable.
- If extraction returns empty fields, refine `mapping.yml` keywords or regex patterns.
- For scanned PDFs, add an OCR pre-pass to make text searchable.
