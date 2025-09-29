import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
import re
import shutil
import yaml
import pandas as pd
from datetime import datetime

# Heavy dependencies imported lazily inside functions: fitz (PyMuPDF), pyodbc, openpyxl

APP_TITLE = "PDF to Excel to Access"
DEFAULT_ROOT = Path.cwd()

# Centralise folder names to avoid drift

IN_DIR = "input"
INBOX_DIR = f"{IN_DIR}/inbox"
ARCHIVE_DIR = f"{IN_DIR}/archive"
REJECTED_DIR = f"{IN_DIR}/rejected"
CONFIG_DIR = "config"
OUTPUT_DIR = "output"
LOGS_DIR = "logs"
REVIEW_FILE = "review.xlsx"
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
    for k in ["fields", "access"]:
        if k not in cfg:
            raise ValueError(f"Missing '{k}' in mapping.yml")
    for k in ["db_path", "table", "column_map"]:
        if k not in cfg["access"]:
            raise ValueError(f"Missing 'access.{k}' in mapping.yml")
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

def write_review_workbook(root: Path, rows: list, cfg: dict):
    outdir = root / OUTPUT_DIR
    review_xlsx = outdir / REVIEW_FILE
    df = pd.DataFrame(rows)

    # Build dedupe key if configured and columns exist
    dk = cfg.get("dedupe_key", [])
    if dk and all(col in df.columns for col in dk):
        df["_dedupe_key"] = df[dk].astype(str).agg("|".join, axis=1)
    else:
        df["_dedupe_key"] = ""

    df["_review_status"] = "PENDING"
    df["_review_comment"] = ""

    ordered = ["_source_pdf", "_extraction_ok", "_dedupe_key", "_review_status", "_review_comment", "_notes"]
    field_names = [f["name"] for f in cfg["fields"]]
    cols = [c for c in ordered if c in df.columns] + [c for c in field_names if c in df.columns]
    df = df.reindex(columns=cols)

    df.to_excel(review_xlsx, index=False)
    return review_xlsx

def extract_to_excel(root: Path, logbox: tk.Text):
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

    review_path = write_review_workbook(root, rows, cfg)
    log_line(logbox, f"Review workbook written to: {review_path}")
    messagebox.showinfo(APP_TITLE, f"Review workbook created:\n{review_path}")

def get_conn(db_path: str):
    import pyodbc
    conn_str = (
        r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
        rf"DBQ={db_path};"
    )
    return pyodbc.connect(conn_str)

def row_exists(cursor, table, key_cols, row, colmap):
    if not key_cols:
        return False
    good_keys = [k for k in key_cols if k in colmap]
    if not good_keys:
        return False
    where = " AND ".join([f"[{colmap[k]}]=?" for k in good_keys])
    sql = f"SELECT COUNT(1) FROM [{table}] WHERE {where}"
    params = [row.get(k) for k in good_keys]
    cursor.execute(sql, params)
    cnt = cursor.fetchone()[0]
    return cnt > 0

def insert_row(cursor, table, row, colmap):
    cols = []
    vals = []
    qmarks = []
    for src_col, dest_col in colmap.items():
        # Guard against missing columns in the review sheet
        val = row.get(src_col) if src_col in row.index else None
        cols.append(f"[{dest_col}]")
        vals.append(val)
        qmarks.append("?")
    sql = f"INSERT INTO [{table}] ({', '.join(cols)}) VALUES ({', '.join(qmarks)})"
    cursor.execute(sql, vals)

def upload_to_access(root: Path, logbox: tk.Text):
    try:
        cfg = load_cfg(root)
    except Exception as e:
        messagebox.showerror(APP_TITLE, f"Config error. {e}")
        return

    review_xlsx = root / OUTPUT_DIR / REVIEW_FILE
    if not review_xlsx.exists():
        messagebox.showerror(APP_TITLE, f"Cannot find {review_xlsx}. Run extraction first.")
        return

    df = pd.read_excel(review_xlsx).fillna("")
    if "_review_status" not in df.columns:
        messagebox.showerror(APP_TITLE, "review.xlsx is missing the _review_status column.")
        return
    approved = df[df["_review_status"].astype(str).str.upper() == "APPROVED"].copy()
    if approved.empty:
        messagebox.showinfo(APP_TITLE, "No APPROVED rows found.")
        return

    access_db = cfg["access"]["db_path"]
    table = cfg["access"]["table"]
    colmap = cfg["access"]["column_map"]
    key_cols = cfg.get("dedupe_key", [])

    log_path = root / LOGS_DIR / f"upload_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    ok, skipped, failed = 0, 0, 0

    try:
        conn = get_conn(access_db)
        cur = conn.cursor()
    except Exception as e:
        messagebox.showerror(APP_TITLE, f"Access connection failed. {e}")
        return

    inbox = root / INBOX_DIR
    archive = root / ARCHIVE_DIR
    rejected = root / REJECTED_DIR

    with open(log_path, "w", encoding="utf-8") as log:
        for _, row in approved.iterrows():
            src = row.get("_source_pdf", "")
            pdf_path = inbox / src
            try:
                if row_exists(cur, table, key_cols, row, colmap):
                    log.write(f"SKIP duplicate key for {src}\n")
                    log_line(logbox, f"SKIP duplicate for {src}")
                    skipped += 1
                    if pdf_path.exists():
                        shutil.move(str(pdf_path), str(archive / src))
                    continue

                insert_row(cur, table, row, colmap)
                ok += 1
                log.write(f"OK inserted {src}\n")
                log_line(logbox, f"OK inserted {src}")
                if pdf_path.exists():
                    shutil.move(str(pdf_path), str(archive / src))
            except Exception as e:
                failed += 1
                log.write(f"FAIL {src}: {e}\n")
                log_line(logbox, f"FAIL {src}: {e}")
                if pdf_path.exists():
                    shutil.move(str(pdf_path), str(rejected / src))

    try:
        conn.commit()
        cur.close()
        conn.close()
    except Exception:
        pass

    summary = f"Upload complete. OK={ok}, SKIPPED={skipped}, FAILED={failed}\nLog: {log_path}"
    log_line(logbox, summary)
    messagebox.showinfo(APP_TITLE, summary)

def test_access(root: Path):
    try:
        cfg = load_cfg(root)
        db = cfg["access"]["db_path"]
        conn = get_conn(db)
        conn.close()
        messagebox.showinfo(APP_TITLE, f"Access connection OK:\n{db}")
    except Exception as e:
        messagebox.showerror(APP_TITLE, f"Access connection failed. {e}")

def open_review(root: Path):
    xlsx = root / OUTPUT_DIR / REVIEW_FILE
    if not xlsx.exists():
        messagebox.showinfo(APP_TITLE, "No review.xlsx found yet.")
        return
    try:
        import os
        os.startfile(str(xlsx))
    except Exception as e:
        messagebox.showerror(APP_TITLE, f"Could not open workbook. {e}")

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
        ttk.Button(btns, text="Extract to Excel", command=self.on_extract).pack(side="left", padx=4)
        ttk.Button(btns, text="Open review workbook", command=self.on_open_review).pack(side="left", padx=4)
        ttk.Button(btns, text="Upload to Access", command=self.on_upload).pack(side="left", padx=4)
        ttk.Button(btns, text="Test Access connection", command=self.on_test_access).pack(side="left", padx=4)

        ttk.Label(frm, text="Activity log").pack(anchor="w")
        self.logbox = tk.Text(frm, height=20, wrap="word", state="disabled", borderwidth=1, relief="solid")
        self.logbox.pack(fill="both", expand=True, pady=(2, 6))

        ttk.Label(frm, text="Place PDFs in input/inbox. Configure config/mapping.yml.").pack(anchor="w")

    def on_extract(self):
        root = Path(self.root_dir.get())
        ensure_dirs(root)
        try:
            extract_to_excel(root, self.logbox)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Extraction failed. {e}")

    def on_upload(self):
        root = Path(self.root_dir.get())
        ensure_dirs(root)
        try:
            upload_to_access(root, self.logbox)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Upload failed. {e}")

    def on_test_access(self):
        root = Path(self.root_dir.get())
        ensure_dirs(root)
        test_access(root)

    def on_open_review(self):
        root = Path(self.root_dir.get())
        ensure_dirs(root)
        open_review(root)

if __name__ == "__main__":
    app = App()
    app.mainloop()
