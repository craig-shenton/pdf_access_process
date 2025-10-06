# PDF to CSV Review Helper

This application extracts structured data from PDF reports, lets you review and edit the results in a spreadsheet, and produces a CSV ready for manual import into Microsoft Access (or any other downstream system).

## Quick Start for end users (Windows)

### What you need

* A Windows PC. Install the **Microsoft Access Database Engine** if you plan to import the CSV into Access.
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
    review.csv
  logs/
```

### Steps

1. Place PDFs into `src/input/inbox`.
2. Run `src/pdf_to_access_app.exe`.
3. Select **Extract for review (CSV)**. This creates `src/output/review.csv`.
4. Select **Open review CSV**, correct values as needed, then set `_review_status` to `APPROVED` or `REJECTED` for each row.
5. Save and close the CSV.
6. Use your own Microsoft Access process (for example, a saved macro or TransferText action) to load the approved rows from `review.csv`.

## Developer guide

### Prerequisites

* Python 3.11 or later.
* Install dependencies:

  ```bash
  pip install -r requirements.txt
  ```

* Install the Microsoft Access Database Engine only if you plan to run Access imports manually on the same machine.

### Configuration: `src/config/mapping.yml`

Defines fields to extract, optional de-duplication keys, and output file names.

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

output:
  review_csv: "review.csv"
```

### Run locally

```bash
python src/pdf_to_access_app.py
```

The app expects:

* PDFs in `src/input/inbox`
* Configuration in `src/config/mapping.yml`
* It writes `src/output/review.csv`

### Build a single-file executable (Windows)

Run on a Windows machine or a Windows CI runner:

```bash
pyinstaller --onefile --noconsole src/pdf_to_access_app.py
# Output: dist/pdf_to_access_app.exe
```

You can also use a Windows GitHub Actions runner to produce the `.exe`.

### Troubleshooting

* **Nothing extracted**. Adjust `mapping.yml` keywords or regex. Confirm your PDFs are text-searchable. For scanned PDFs, add an OCR pre-pass.
* **Spreadsheet shows strange characters**. Ensure your editor opens `review.csv` as UTF-8 with BOM.
* **Access import issues**. Confirm your manual Access process references the correct CSV path and column names.
