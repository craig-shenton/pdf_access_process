import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
import re
import shutil
import yaml
from datetime import datetime

# Heavy dependencies imported lazily inside functions: fitz (PyMuPDF)

from output.review_writer import (
    REVIEW_STATUS_COLUMN,
    build_review_dataframe,
    load_review_dataframe,
    write_access_ready_csv,
    write_review_csv,
)
from integrations.access_bulk import import_csv_to_access
from ui.review_panel import ReviewPanel

APP_TITLE = "PDF to CSV to Access"
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
    for k in ["fields", "access"]:
        if k not in cfg:
            raise ValueError(f"Missing '{k}' in mapping.yml")
    for k in ["db_path", "column_map", "bulk_import"]:
        if k not in cfg["access"]:
            raise ValueError(f"Missing 'access.{k}' in mapping.yml")
    if "output" not in cfg:
        raise ValueError("Missing 'output' section in mapping.yml")
    for k in ["review_csv", "access_ready_csv"]:
        if k not in cfg["output"]:
            raise ValueError(f"Missing 'output.{k}' in mapping.yml")
    for k in ["msaccess_path", "macro"]:
        if k not in cfg["access"]["bulk_import"]:
            raise ValueError(f"Missing 'access.bulk_import.{k}' in mapping.yml")
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
        return None

    inbox = root / INBOX_DIR
    pdfs = sorted(p for p in inbox.glob("*") if p.suffix in PDF_EXTS)
    if not pdfs:
        messagebox.showinfo(APP_TITLE, f"No PDFs found in {inbox}")
        return None

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
    access_ready_path, _ = write_access_ready_csv(
        df,
        outdir,
        cfg["output"]["access_ready_csv"],
        cfg,
    )
    log_line(logbox, f"Review CSV updated: {review_path}")
    log_line(logbox, f"Access-ready CSV written to: {access_ready_path}")
    log_line(logbox, "Review table refreshed in the application UI.")
    messagebox.showinfo(
        APP_TITLE,
        "Extraction complete. Review data has been loaded into the in-app table.",
    )
    return df


def upload_to_access(root: Path, logbox: tk.Text):
    try:
        cfg = load_cfg(root)
    except Exception as e:
        messagebox.showerror(APP_TITLE, f"Config error. {e}")
        return

    review_csv = root / OUTPUT_DIR / cfg["output"]["review_csv"]
    if not review_csv.exists():
        messagebox.showerror(APP_TITLE, f"Cannot find {review_csv}. Run extraction first.")
        return

    try:
        df = load_review_dataframe(review_csv)
    except Exception as e:
        messagebox.showerror(APP_TITLE, f"Could not read {review_csv}: {e}")
        return
    if REVIEW_STATUS_COLUMN not in df.columns:
        messagebox.showerror(APP_TITLE, f"{review_csv.name} is missing the {REVIEW_STATUS_COLUMN} column.")
        return

    outdir = root / OUTPUT_DIR
    outdir.mkdir(parents=True, exist_ok=True)
    access_ready_path, approved_rows = write_access_ready_csv(
        df,
        outdir,
        cfg["output"]["access_ready_csv"],
        cfg,
    )

    if approved_rows.empty:
        messagebox.showinfo(APP_TITLE, "No APPROVED rows found.")
        return

    bulk_cfg = cfg["access"]["bulk_import"]
    inbox = root / INBOX_DIR
    archive = root / ARCHIVE_DIR
    rejected = root / REJECTED_DIR
    archive.mkdir(parents=True, exist_ok=True)
    rejected.mkdir(parents=True, exist_ok=True)

    log_path = root / LOGS_DIR / f"bulk_import_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    log_path.parent.mkdir(parents=True, exist_ok=True)

    log_messages = []

    def record(msg: str):
        log_messages.append(msg)
        log_line(logbox, msg)

    record(f"Preparing bulk import with {len(approved_rows)} approved rows")

    try:
        result = import_csv_to_access(
            msaccess_path=bulk_cfg["msaccess_path"],
            database_path=cfg["access"]["db_path"],
            csv_path=access_ready_path,
            macro=bulk_cfg.get("macro"),
            timeout=bulk_cfg.get("timeout_sec", 600),
            use_cmd_argument=bulk_cfg.get("use_cmd_argument", True),
            extra_args=bulk_cfg.get("extra_args", []),
            workdir=bulk_cfg.get("workdir"),
            log_path=log_path,
        )
        record(f"Access import completed with return code {result.returncode}")
        success = True
    except Exception as e:
        record(f"Access import failed: {e}")
        success = False

    target_dir = archive if success else rejected
    action = "archived" if success else "rejected"
    for _, row in approved_rows.iterrows():
        src = row.get("_source_pdf", "")
        if not src:
            continue
        pdf_path = inbox / src
        if not pdf_path.exists():
            continue
        dest = target_dir / src
        try:
            shutil.move(str(pdf_path), str(dest))
            record(f"{action.upper()} {src}")
        except Exception as move_err:
            record(f"Failed to move {src}: {move_err}")

    if log_messages:
        with open(log_path, "a", encoding="utf-8") as log:
            for msg in log_messages:
                log.write(msg + "\n")

    if success:
        summary = (
            f"Bulk upload complete. Imported {len(approved_rows)} rows.\n"
            f"Access-ready CSV: {access_ready_path}\nLog: {log_path}"
        )
        messagebox.showinfo(APP_TITLE, summary)
    else:
        summary = (
            f"Bulk upload failed. Approved PDFs moved to rejected.\n"
            f"Access-ready CSV: {access_ready_path}\nLog: {log_path}"
        )
        messagebox.showerror(APP_TITLE, summary)


def test_access(root: Path):
    try:
        cfg = load_cfg(root)
        db = Path(cfg["access"]["db_path"])
        bulk_cfg = cfg["access"]["bulk_import"]
        msaccess = Path(bulk_cfg["msaccess_path"])
        issues = []
        if not db.exists():
            issues.append(f"Database not found: {db}")
        if not msaccess.exists():
            issues.append(f"msaccess.exe not found: {msaccess}")
        if issues:
            raise FileNotFoundError("; ".join(issues))
        messagebox.showinfo(APP_TITLE, f"Access configuration looks OK.\nDB: {db}\nmsaccess: {msaccess}")
    except Exception as e:
        messagebox.showerror(APP_TITLE, f"Access connection failed. {e}")

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("780x520")
        self.resizable(True, True)

        self.root_dir = tk.StringVar(value=str(DEFAULT_ROOT))
        self.create_widgets()
        self.try_load_existing_review()

    def browse_root(self):
        chosen = filedialog.askdirectory(title="Select project root")
        if chosen:
            self.root_dir.set(chosen)
            self.try_load_existing_review()

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
        ttk.Button(btns, text="Bulk upload to Access", command=self.on_upload).pack(side="left", padx=4)
        ttk.Button(btns, text="Test Access setup", command=self.on_test_access).pack(side="left", padx=4)

        ttk.Label(frm, text="Review extracted data").pack(anchor="w", pady=(4, 2))
        self.review_panel = ReviewPanel(frm)
        self.review_panel.pack(fill="both", expand=True, pady=(0, 8))
        self.review_panel.set_save_callback(lambda: self.save_review_changes(show_message=True))

        ttk.Label(frm, text="Activity log").pack(anchor="w")
        self.logbox = tk.Text(frm, height=12, wrap="word", state="disabled", borderwidth=1, relief="solid")
        self.logbox.pack(fill="both", expand=True, pady=(2, 6))

        ttk.Label(frm, text="Place PDFs in input/inbox. Use the in-app review table above before bulk upload.").pack(anchor="w")

    def on_extract(self):
        root = Path(self.root_dir.get())
        ensure_dirs(root)
        try:
            df = extract_to_review(root, self.logbox)
            if df is not None:
                self.review_panel.load_dataframe(df)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Extraction failed. {e}")

    def on_upload(self):
        root = Path(self.root_dir.get())
        ensure_dirs(root)
        # Persist any staged review edits before uploading
        self.save_review_changes(show_message=False)
        try:
            upload_to_access(root, self.logbox)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Upload failed. {e}")

    def on_test_access(self):
        root = Path(self.root_dir.get())
        ensure_dirs(root)
        test_access(root)

    def save_review_changes(self, show_message: bool = True) -> bool:
        if not self.review_panel.has_data():
            return False

        if hasattr(self.review_panel, "is_dirty") and not self.review_panel.is_dirty():
            if show_message:
                messagebox.showinfo(APP_TITLE, "No review changes to save.")
            return False

        root = Path(self.root_dir.get())
        ensure_dirs(root)
        try:
            cfg = load_cfg(root)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Config error. {e}")
            return False

        try:
            df = self.review_panel.get_dataframe()
        except ValueError:
            return False

        outdir = root / OUTPUT_DIR
        outdir.mkdir(parents=True, exist_ok=True)
        try:
            path = write_review_csv(outdir, cfg["output"]["review_csv"], df)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Failed to save review CSV. {e}")
            return False

        log_line(self.logbox, f"Review CSV saved: {path}")
        if show_message:
            messagebox.showinfo(APP_TITLE, f"Review changes saved to:\n{path}")
        if hasattr(self.review_panel, "mark_clean"):
            self.review_panel.mark_clean()
        return True

    def try_load_existing_review(self) -> None:
        root = Path(self.root_dir.get())
        try:
            cfg = load_cfg(root)
        except Exception:
            # Config not ready yet; clear any stale data
            self.review_panel.clear()
            return

        review_csv = root / OUTPUT_DIR / cfg["output"]["review_csv"]
        if not review_csv.exists():
            self.review_panel.clear()
            return

        try:
            df = load_review_dataframe(review_csv)
        except Exception as e:
            log_line(self.logbox, f"Failed to load existing review CSV: {e}")
            self.review_panel.clear()
            return

        self.review_panel.load_dataframe(df)
        log_line(self.logbox, f"Loaded review data from: {review_csv}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
