# OCR Local Processing Tool

This tool provides a minimal, self-contained way to extract text from PDFs and images **fully offline**, with no reliance on internet services or LLMs. It uses:

* **pypdf** to detect and extract embedded text from PDFs
* **pdf2image + Poppler** to rasterise scanned pages
* **OpenCV** for pre-processing
* **Tesseract via pytesseract** for OCR
* **Pandas** for structured review outputs

It is designed for sensitive data such as PII, where all processing must happen locally.

## Features

* Detects whether a PDF page already has embedded text, to avoid unnecessary OCR.
* Rasterises and OCRs pages with images only.
* Pre-processing (denoise, adaptive threshold) for better OCR accuracy.
* Exports:

  * `document.txt` with concatenated text from all pages
  * `page_XXX.csv` files with word-level bounding boxes and confidence scores

## Requirements

### System dependencies

* **Tesseract OCR**

  * Ubuntu/Debian: `sudo apt-get install tesseract-ocr`
  * macOS: `brew install tesseract`
  * Windows: install from [UB Mannheim build](https://github.com/UB-Mannheim/tesseract/wiki)

* **Poppler** (for PDF rasterisation)

  * Ubuntu/Debian: `sudo apt-get install poppler-utils`
  * macOS: `brew install poppler`
  * Windows: install Poppler for Windows and add the `bin` folder to PATH.

### Python packages

Install into a virtual environment:

```bash
pip install -r requirements.txt
```

Contents of `requirements.txt`:

```bash
pytesseract
pdf2image
opencv-python
pypdf
pandas
openpyxl
```

## Usage

```bash
# OCR a scanned PDF
python3 src/ocr_local.py path/to/document.pdf --out results

# OCR a single image (PNG, JPG, TIFF, etc.)
python3 src/ocr_local.py path/to/image.jpg --out results_img
```

Outputs:

* `document.txt` – concatenated text from all pages
* `page_001.csv`, `page_002.csv`, etc. – word-level tokens with positions and confidence

## Example

```bash
python3 src/ocr_local.py test-ocr.pdf --out output
```

Console output:

```bash
Processing PDF: test-ocr.pdf
  Page 1/1: used embedded text
Done. Outputs in: /.../output
  - document.txt for human reading
  - page_001.csv for review
```

`document.txt` now contains the recognised text.
`page_001.csv` contains structured data for each recognised token.

## Developer Notes

* Pre-processing happens in `preprocess_for_ocr`: greyscale → denoise → adaptive threshold.
  This can be customised for specific document types (e.g. forms, tables).
* `ocr_image_to_tsv` calls Tesseract with `--psm 6` (assumes uniform text blocks).
  Adjust for different layouts:

  * `--psm 4` for columnar text
  * `--psm 11` for sparse text (forms, scattered fields)
* Confidence scores in the CSV can be used to flag low-quality outputs for human review.
