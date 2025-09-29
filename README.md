# PDF to Excel to Access Uploader

This application extracts structured data from PDF reports, lets you review and edit the results in Excel, then uploads approved records into a Microsoft Access database. It ships as a single desktop app.

## Quick Start for end users (Windows)

### What you need

* A Windows PC with the **Microsoft Access Database Engine** installed. The bitness must match the executable.
* An Access database file (`.accdb` or `.mdb`) whose path is set in `src/config/mapping.yml`.
* The provided `src` folder structure.

### Folder layout

```bash
src/
  pdf_to_access_app.exe
  config/
    mapping.yml
  input/
    inbox/
    archive/
    rejected/
  output/
    review.xlsx
  logs/
```

### Steps

1. Place PDFs into `src/input/inbox`.
2. Run `src/pdf_to_access_app.exe`.
3. Select **Extract to Excel**. This creates `src/output/review.xlsx`.
4. Select **Open review workbook**, correct values as needed, then set `_review_status` to `APPROVED` or `REJECTED`.
5. Save and close Excel.
6. Select **Upload to Access**. The app inserts approved rows, moves PDFs to `input/archive` or `input/rejected`, and writes a log in `logs/`.
7. Use **Test Access connection** to verify connectivity if required.

## Developer guide

### Prerequisites

* Python 3.11 or later.
* Install dependencies:

  ```bash
  pip install -r requirements.txt
  ```

* For Windows upload and Windows builds, install the **Microsoft Access ODBC driver** (Access Database Engine) that matches your Python and build bitness.

### Configuration: `src/config/mapping.yml`

Defines fields to extract, optional de-duplication keys, and Access connection details.

Supported strategies:

* `keyword_right` captures text to the right of a keyword in the same block.
* `keyword_line` captures the remainder of the line after a keyword.
* `regex` extracts using a regular expression.

Minimum keys:

```yaml
fields:
  - name: "case_id"
    find:
      type: "keyword_line"
      keywords: ["Case ID"]

dedupe_key: ["case_id"]   # optional

access:
  db_path: "C:/Data/Access/MyDatabase.accdb"
  table: "Cases"
  column_map:
    case_id: "CaseID"
```

### Run locally

```bash
python src/pdf_to_access_app.py
```

The app expects:

* PDFs in `src/input/inbox`
* Configuration in `src/config/mapping.yml`
* It writes `src/output/review.xlsx` and logs to `src/logs`

### Build a single-file executable (Windows)

Run on a Windows machine or a Windows CI runner:

```bash
pyinstaller --onefile --noconsole src/pdf_to_access_app.py
# Output: dist/pdf_to_access_app.exe
```

You can also use a Windows GitHub Actions runner to produce the `.exe`.

### Troubleshooting

* **Access driver errors**. Ensure the Access Database Engine is installed and that its bitness matches the executable.
* **Nothing extracted**. Adjust `mapping.yml` keywords or regex. Confirm your PDFs are text-searchable. For scanned PDFs, add an OCR pre-pass.
* **Upload issues**. Confirm `access.table` and all `access.column_map` destinations exist in the Access database.
* **File in use**. Close `review.xlsx` before running the upload.
