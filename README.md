# PDF to CSV Review Helper

This desktop application extracts structured information from PDF lab reports, lets you verify the
results in a spreadsheet, and exports a CSV that is ready for a downstream Microsoft Access import
(or any other workflow that consumes CSV files).

## Features

* Point the app at a root folder and it will ensure the expected sub-folders exist.
* Extract values from PDFs using keyword, keyword-right, or regular expression strategies defined in
  `config/mapping.yml`.
* Produce a review-ready CSV that highlights extraction success, failures, and notes.
* Launch the review CSV directly from the UI for quick editing before you import approved rows into
  Access.

## Project layout

```
config/
  mapping.yml        # field definitions and output file name
input/
  inbox/             # drop PDFs here
  archive/
  rejected/
output/
  review.csv         # generated CSV for review and manual import
logs/
src/
  pdf_to_access_app.py
```

The packaged Windows executable mirrors this layout inside the folder where it is extracted.

## Running the app from source

1. Install Python 3.11 or newer.
2. Install dependencies:
   ```bash
   python -m pip install -r requirements.txt
   ```
3. Launch the application from the project root:
   ```bash
   python src/pdf_to_access_app.py
   ```
4. Use the **Project root** picker to point the UI at your working folder (the default is the
   repository root). The application will create the expected sub-folders if they do not exist.
5. Drop PDFs into `input/inbox`, choose **Extract for review (CSV)**, and then review the generated
   CSV before importing the approved records into Access using your own process.

## Running the packaged EXE

A pre-built `pdf_to_access_app.exe` can be distributed to Windows users. Place it alongside the
`config`, `input`, `output`, and `logs` folders. Double-click the executable to launch the same UI as
running from source.

## Building a Windows executable (64-bit)

You can generate a standalone executable either locally on Windows or via GitHub Actions.

### Local build

```bash
python -m pip install -r requirements.txt
pyinstaller --onefile --noconsole src/pdf_to_access_app.py
# Result: dist/pdf_to_access_app.exe
```

### GitHub Actions build

Trigger the **Build Windows EXE** workflow. It uses a 64-bit Python runtime to install dependencies
and run PyInstaller, then publishes the resulting `pdf_to_access_app.exe` artifact as
`pdf_to_access_app-win64`.

## Configuration tips

* `fields`: list the data points you want to extract. Each field requires a `find` strategy with a
  `type` (`keyword_line`, `keyword_right`, or `regex`) and supporting parameters.
* `dedupe_key`: optional list of columns used to flag potential duplicates in the review CSV.
* `output.review_csv`: file name for the generated CSV. The file is written to the `output`
  directory under the selected project root.

## Troubleshooting

* **No PDFs found** – Confirm that your PDFs are in the `input/inbox` directory under the selected
  root.
* **Empty or incorrect extraction results** – Refine the strategies in `config/mapping.yml` or
  ensure the PDFs contain searchable text (run OCR if necessary).
* **Cannot open the review CSV from the app** – Open the CSV manually from the `output` folder and
  verify that Microsoft Excel or your preferred editor is installed.
