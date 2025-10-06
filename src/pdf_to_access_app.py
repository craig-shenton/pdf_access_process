import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
import re
import yaml

from ui.app import PdfToAccessApp

from output.review_writer import build_review_dataframe, write_review_csv

APP_TITLE = "PDF to CSV Review Helper"
DEFAULT_ROOT = Path.cwd()

# Centralise folder names to avoid drift

IN_DIR = "input"
INBOX_DIR = f"{IN_DIR}/inbox"
ARCHIVE_DIR = f"{IN_DIR}/archive"
REJECTED_DIR = f"{IN_DIR}/rejected"
CONFIG_DIR = "config"
OUTPUT_DIR = "output"
LOGS_DIR = "logs"
PDF_EXTS = (".pdf", ".PDF")

def ensure_dirs(root: Path):
    (root / CONFIG_DIR).mkdir(parents=True, exist_ok=True)
    (root / INBOX_DIR).mkdir(parents=True, exist_ok=True)
    (root / ARCHIVE_DIR).mkdir(parents=True, exist_ok=True)
    (root / REJECTED_DIR).mkdir(parents=True, exist_ok=True)
    (root / OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
    (root / LOGS_DIR).mkdir(parents=True, exist_ok=True)

def load_cfg(root: Path):
    cfg_path = root / CONFIG_DIR / "mapping.yml"
    if not cfg_path.exists():
        raise FileNotFoundError(f"Cannot find {cfg_path}")
    with open(cfg_path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    # Light validation
    if "fields" not in cfg:
        raise ValueError("Missing 'fields' in mapping.yml")
    if "output" not in cfg:
        raise ValueError("Missing 'output' section in mapping.yml")
    if "review_csv" not in cfg["output"]:
        raise ValueError("Missing 'output.review_csv' in mapping.yml")
    return cfg

def log_line(textbox: tk.Text, msg: str):
    textbox.configure(state="normal")
    textbox.insert("end", msg + "\n")
    textbox.see("end")
    textbox.configure(state="disabled")
    textbox.update()

def read_pdf_text_blocks(pdf_path: Path):
    import fitz  # PyMuPDF
    pages = []
    # Ensure the document is closed promptly
    with fitz.open(pdf_path) as doc:
        for p in doc:
            blocks = p.get_text("blocks")
            pages.append(blocks)
    return pages

def find_value_in_blocks(blocks, strategy):
    t = strategy.get("type")
    if t == "regex":
        pattern = strategy["pattern"]
        full_text = "\n".join(b[4] for b in blocks if isinstance(b, (list, tuple)) and len(b) >= 5)
        m = re.search(pattern, full_text, flags=re.IGNORECASE)
        if m:
            return m.group(1) if m.groups() else m.group(0)
        return None

    keywords = [k.lower() for k in strategy.get("keywords", [])]

    if t == "keyword_line":
        for b in blocks:
            text = b[4] if len(b) >= 5 else ""
            low = text.lower()
            for kw in keywords:
                if kw in low:
                    idx = low.find(kw)
                    tail = text[idx + len(kw):].strip(" :\t\r\n")
                    if tail:
                        return tail
        return None

    if t == "keyword_right":
        for b in blocks:
            text = b[4] if len(b) >= 5 else ""
            low = text.lower()
            for kw in keywords:
                if kw in low:
                    parts = re.split(re.escape(kw), text, flags=re.IGNORECASE, maxsplit=1)
                    if len(parts) == 2:
                        right = parts[1].strip(" :\t\r\n")
                        right = right.splitlines()[0].strip()
                        if right:
                            return right
        return None

    return None

def extract_from_pdf(pdf_path: Path, cfg: dict):
    record = {"_source_pdf": pdf_path.name}
    try:
        pages = read_pdf_text_blocks(pdf_path)
        for field in cfg["fields"]:
            name = field["name"]
            strat = field["find"]
            value = None
            for blocks in pages:
                value = find_value_in_blocks(blocks, strat)
                if value:
                    break
            record[name] = value
        record["_extraction_ok"] = True
        record["_notes"] = ""
    except Exception as e:
        record["_extraction_ok"] = False
        record["_notes"] = f"Extraction error: {e}"
    return record

def extract_to_review(root: Path, logbox: tk.Text):
    try:
        cfg = load_cfg(root)
    except Exception as e:
        messagebox.showerror(APP_TITLE, f"Config error. {e}")
        return

    inbox = root / INBOX_DIR
    pdfs = sorted(p for p in inbox.glob("*") if p.suffix in PDF_EXTS)
    if not pdfs:
        messagebox.showinfo(APP_TITLE, f"No PDFs found in {inbox}")
        return

    log_line(logbox, f"Found {len(pdfs)} PDFs")
    rows = []
    for i, pdf in enumerate(pdfs, 1):
        log_line(logbox, f"[{i}/{len(pdfs)}] Extracting {pdf.name}")
        rec = extract_from_pdf(pdf, cfg)
        rows.append(rec)

    df = build_review_dataframe(rows, cfg)
    outdir = root / OUTPUT_DIR
    outdir.mkdir(parents=True, exist_ok=True)
    review_path = write_review_csv(outdir, cfg["output"]["review_csv"], df)
    log_line(logbox, f"Review CSV written to: {review_path}")
    log_line(
        logbox,
        "Use the review CSV to make corrections and import approved rows into Access manually.",
    )
    messagebox.showinfo(
        APP_TITLE,
        (
            "Review CSV created.\n"
            f"Path: {review_path}\n"
            "After reviewing, import approved rows into Access using your own process."
        ),
    )


def open_review(root: Path):
    try:
        cfg = load_cfg(root)
    except Exception as e:
        messagebox.showerror(APP_TITLE, f"Config error. {e}")
        return
    csv_path = root / OUTPUT_DIR / cfg["output"]["review_csv"]
    if not csv_path.exists():
        messagebox.showinfo(APP_TITLE, "No review CSV found yet.")
        return
    try:
        import os
        os.startfile(str(csv_path))
    except Exception as e:
        messagebox.showerror(APP_TITLE, f"Could not open review CSV. {e}")

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("780x520")
        self.resizable(True, True)

        self.root_dir = tk.StringVar(value=str(DEFAULT_ROOT))
        self.create_widgets()

    def browse_root(self):
        chosen = filedialog.askdirectory(title="Select project root")
        if chosen:
            self.root_dir.set(chosen)

    def create_widgets(self):
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        row1 = ttk.Frame(frm)
        row1.pack(fill="x", pady=(0, 8))
        ttk.Label(row1, text="Project root folder:").pack(side="left")
        ttk.Entry(row1, textvariable=self.root_dir).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row1, text="Browse…", command=self.browse_root).pack(side="left")

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=6)
        ttk.Button(btns, text="Extract for review (CSV)", command=self.on_extract).pack(side="left", padx=4)
        ttk.Button(btns, text="Open review CSV", command=self.on_open_review).pack(side="left", padx=4)

        ttk.Label(frm, text="Activity log").pack(anchor="w")
        self.logbox = tk.Text(frm, height=20, wrap="word", state="disabled", borderwidth=1, relief="solid")
        self.logbox.pack(fill="both", expand=True, pady=(2, 6))

        ttk.Label(
            frm,
            text="Place PDFs in input/inbox. Review output/review CSV, then import approved rows into Access manually.",
        ).pack(anchor="w")

    def on_extract(self):
        root = Path(self.root_dir.get())
        ensure_dirs(root)
        try:
            extract_to_review(root, self.logbox)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Extraction failed. {e}")

    def on_open_review(self):
        root = Path(self.root_dir.get())
        ensure_dirs(root)
        open_review(root)

if __name__ == "__main__":
    main()
