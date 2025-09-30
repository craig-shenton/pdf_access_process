import argparse
import os
from pathlib import Path
import sys
import tempfile

import cv2
import numpy as np
import pandas as pd
import pytesseract as tess
from pdf2image import convert_from_path
from pypdf import PdfReader


def has_embedded_text(pdf_path: Path, page_index: int) -> bool:
    """Return True if the given PDF page contains extractable text."""
    reader = PdfReader(str(pdf_path))
    if page_index < 0 or page_index >= len(reader.pages):
        return False
    text = reader.pages[page_index].extract_text() or ""
    # Heuristic to ignore spurious whitespace
    return len(text.strip()) >= 5


def rasterise_pdf_pages(pdf_path: Path, dpi: int = 300, pages=None):
    """Return list of PIL images for requested pages."""
    return convert_from_path(str(pdf_path), dpi=dpi, first_page=None, last_page=None, fmt="png", thread_count=2)


def preprocess_for_ocr(img_bgr: np.ndarray) -> np.ndarray:
    """Minimal but robust preprocessing for OCR."""
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    # Light denoise
    den = cv2.fastNlMeansDenoising(gray, h=15)
    # Adaptive threshold for uneven lighting
    return cv2.adaptiveThreshold(
        den, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 10
    )


def ocr_image_to_tsv(img_bgr: np.ndarray, lang: str = "eng"):
    """Run Tesseract and return TSV dict."""
    proc = preprocess_for_ocr(img_bgr)
    data = tess.image_to_data(
        proc, lang=lang, config="--psm 6", output_type=tess.Output.DATAFRAME
    )
    # Drop rows that are NaN or empty
    data = data.dropna(subset=["text"])
    data["text"] = data["text"].astype(str)
    data = data[data["text"].str.strip() != ""]
    return data


def join_text(lines_df: pd.DataFrame) -> str:
    """Concatenate recognised text roughly in reading order."""
    if lines_df.empty:
        return ""
    # Keep only word level entries
    words = lines_df[lines_df["level"] == 5].copy()
    if words.empty:
        return ""
    # Group by line for crude line reconstruction
    group_cols = ["page_num", "block_num", "par_num", "line_num"]
    text_lines = []
    for _, group in words.groupby(group_cols):
        # Sort by x
        group = group.sort_values(by="left")
        line_text = " ".join(group["text"].tolist())
        text_lines.append(line_text)
    return "\n".join(text_lines)


def save_outputs(base_out: Path, page_idx: int, tsv_df: pd.DataFrame, page_text: str):
    page_tag = f"page_{page_idx+1:03d}"
    # CSV with all tokens, boxes and confidence
    csv_path = base_out / f"{page_tag}.csv"
    tsv_df.to_csv(csv_path, index=False)
    # Append to a single TXT for convenience
    with open(base_out / "document.txt", "a", encoding="utf-8") as f:
        f.write(f"\n===== {page_tag} =====\n")
        f.write(page_text.strip() + "\n")


def ensure_tools_available():
    """Basic checks for local binaries."""
    # Tesseract
    try:
        _ = tess.get_tesseract_version()
    except Exception as e:
        print(
            "Tesseract is not available. Install it and ensure the binary is on PATH.\n"
            "Windows: https://github.com/UB-Mannheim/tesseract/wiki\n"
            "macOS: brew install tesseract\n"
            "Linux: sudo apt-get install tesseract-ocr",
            file=sys.stderr,
        )
        raise
    # pdf2image requires Poppler for PDFs


def extract_text_from_pdf_page(pdf_path: Path, page_index: int) -> str:
    reader = PdfReader(str(pdf_path))
    return (reader.pages[page_index].extract_text() or "").strip()


def process_pdf(pdf_path: Path, out_dir: Path, dpi: int, lang: str):
    print(f"Processing PDF: {pdf_path}")
    out_dir.mkdir(parents=True, exist_ok=True)
    # Clear or create combined text file
    (out_dir / "document.txt").write_text("", encoding="utf-8")

    reader = PdfReader(str(pdf_path))
    n_pages = len(reader.pages)
    # Render all pages once to keep logic simple
    pil_pages = convert_from_path(str(pdf_path), dpi=dpi, fmt="png", thread_count=2)

    for i, pil in enumerate(pil_pages):
        if page_has_text := has_embedded_text(pdf_path, i):
            # Use native text to avoid OCR errors
            native_text = extract_text_from_pdf_page(pdf_path, i)
            # Build a minimal DataFrame to keep output shape consistent
            df = pd.DataFrame(
                {
                    "level": [],
                    "page_num": [],
                    "block_num": [],
                    "par_num": [],
                    "line_num": [],
                    "word_num": [],
                    "left": [],
                    "top": [],
                    "width": [],
                    "height": [],
                    "conf": [],
                    "text": [],
                }
            )
            save_outputs(out_dir, i, df, native_text)
            print(f"  Page {i+1}/{n_pages}: used embedded text")
        else:
            img = cv2.cvtColor(np.array(pil), cv2.COLOR_RGB2BGR)
            df = ocr_image_to_tsv(img, lang=lang)
            page_text = join_text(df)
            save_outputs(out_dir, i, df, page_text)
            print(f"  Page {i+1}/{n_pages}: OCR complete")


def process_image(image_path: Path, out_dir: Path, lang: str):
    print(f"Processing image: {image_path}")
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "document.txt").write_text("", encoding="utf-8")

    img = cv2.imread(str(image_path))
    if img is None:
        raise ValueError(f"Cannot read image at {image_path}")
    df = ocr_image_to_tsv(img, lang=lang)
    page_text = join_text(df)
    save_outputs(out_dir, 0, df, page_text)
    print("  Image OCR complete")


def main():
    parser = argparse.ArgumentParser(description="Local OCR for PDFs and images. No network usage.")
    parser.add_argument("input_path", type=str, help="Path to a PDF or an image file")
    parser.add_argument("--out", type=str, default="ocr_output", help="Output folder")
    parser.add_argument("--dpi", type=int, default=300, help="DPI for PDF rasterisation")
    parser.add_argument("--lang", type=str, default="eng", help="Tesseract language, e.g. eng")
    args = parser.parse_args()

    ensure_tools_available()

    input_path = Path(args.input_path)
    out_dir = Path(args.out)

    if not input_path.exists():
        print(f"Input not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    if input_path.suffix.lower() == ".pdf":
        process_pdf(input_path, out_dir, dpi=args.dpi, lang=args.lang)
    else:
        process_image(input_path, out_dir, lang=args.lang)

    print(f"\nDone. Outputs in: {out_dir.resolve()}")
    print("  - document.txt for human reading")
    print("  - page_XXX.csv for review with boxes and confidences")


if __name__ == "__main__":
    main()
