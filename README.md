# PDF to CSV to Access Uploader

This application extracts structured data from PDF reports, lets you review and edit the results in a CSV grid, then uploads approved records into a Microsoft Access database via a bulk-import macro. It ships as a single desktop app.

## Quick Start for end users (Windows)

### What you need

* A Windows PC with the **Microsoft Access Database Engine** installed. The bitness must match the executable.
* An Access database file (`.accdb` or `.mdb`) whose path is set in `src/config/mapping.yml`.
* The provided `src` folder structure.
* A macro inside your Access database that can consume the Access-ready CSV (see [Configuring the Access macro](#configuring-the-access-macro)).

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
    access_ready.csv
  logs/
```

### Steps

1. Place PDFs into `src/input/inbox`.
2. Run `src/pdf_to_access_app.exe`.
3. Select **Extract to CSV**. This creates `src/output/review.csv` and `src/output/access_ready.csv`.
4. Select **Open review CSV**, correct values as needed, then set `_review_status` to `APPROVED` or `REJECTED`.
5. Save and close the CSV editor.
6. Select **Upload to Access**. The app runs your Access macro against the Access-ready CSV, moves PDFs to `input/archive` or `input/rejected`, and writes a log in `logs/`.
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
* It writes `src/output/review.csv`, `src/output/access_ready.csv`, and logs to `src/logs`

### Configuring the Access macro

The bulk upload button launches `msaccess.exe` with the `/x <macro>` switch so your database can run a macro that imports the approved rows from the Access-ready CSV. To set this up:

1. Open your Access database (`mapping.yml > access.db_path`).
2. Create a module (e.g., `modBulkImport`) with a VBA helper that reads the command-line argument and runs `TransferText`:

   ```vb
   Public Function ImportApprovedFromCsv()
       Dim csvPath As String
       csvPath = Command()

       If Len(csvPath) = 0 Then
           MsgBox "CSV path not provided by uploader."
           Exit Function
       End If

       DoCmd.SetWarnings False
       DoCmd.TransferText acImportDelim, "ApprovedImportSpec", "ApprovedTable", csvPath, True
       DoCmd.SetWarnings True
   End Function
   ```

   Replace `ApprovedImportSpec` with the name of your saved import specification (or `Null` to use defaults) and `ApprovedTable` with the Access table that should receive the approved rows.

3. Create a macro named to match `mapping.yml > access.bulk_import.macro` (default: `ImportApproved`). The macro should run the VBA function above via the **RunCode** action and can include any additional logic you need (e.g., clearing staging tables first).
4. Save the macro. When the application executes **Upload to Access**, it will call:

   ```
   msaccess.exe "<db_path>" /x ImportApproved /cmd "<full path to access_ready.csv>"
   ```

   Inside Access, the `Command()` function returns the CSV filename supplied via `/cmd`.

If you rename the macro or need a different CSV filename, update `config/mapping.yml` accordingly.

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
* **File in use**. Close `review.csv` before running the upload.
