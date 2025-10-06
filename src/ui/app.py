"""Tkinter user interface for the PDF to Access workflow."""
from __future__ import annotations

import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core.config import ConfigError
from core.workflow import (
    NoApprovedRowsError,
    NoInputFilesError,
    ReviewFileMissingError,
    WorkflowError,
    WorkflowService,
)

APP_TITLE = "PDF to CSV to Access"
DEFAULT_ROOT = Path.cwd()


class PdfToAccessApp(tk.Tk):
    """Tkinter application orchestrating the PDF to Access workflow."""

    def __init__(self, *, default_root: Path | None = None) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("780x520")
        self.resizable(True, True)

        root_dir = default_root or DEFAULT_ROOT
        self.root_dir = tk.StringVar(value=str(root_dir))

        self._build_layout()

    # ------------------------------------------------------------------
    # UI helpers
    # ------------------------------------------------------------------
    def _build_layout(self) -> None:
        container = ttk.Frame(self, padding=12)
        container.pack(fill="both", expand=True)

        row = ttk.Frame(container)
        row.pack(fill="x", pady=(0, 8))
        ttk.Label(row, text="Project root folder:").pack(side="left")
        ttk.Entry(row, textvariable=self.root_dir).pack(
            side="left", fill="x", expand=True, padx=6
        )
        ttk.Button(row, text="Browse…", command=self._on_browse).pack(side="left")

        buttons = ttk.Frame(container)
        buttons.pack(fill="x", pady=6)
        ttk.Button(
            buttons, text="Extract for review (CSV)", command=self._on_extract
        ).pack(side="left", padx=4)
        ttk.Button(
            buttons, text="Open review CSV", command=self._on_open_review
        ).pack(side="left", padx=4)
        ttk.Button(
            buttons, text="Bulk upload to Access", command=self._on_upload
        ).pack(side="left", padx=4)
        ttk.Button(
            buttons, text="Test Access setup", command=self._on_test_access
        ).pack(side="left", padx=4)

        ttk.Label(container, text="Activity log").pack(anchor="w")
        self.logbox = tk.Text(
            container,
            height=20,
            wrap="word",
            state="disabled",
            borderwidth=1,
            relief="solid",
        )
        self.logbox.pack(fill="both", expand=True, pady=(2, 6))

        ttk.Label(
            container,
            text="Place PDFs in input/inbox. Review output/review CSV before bulk upload.",
        ).pack(anchor="w")

    def _append_log(self, message: str) -> None:
        self.logbox.configure(state="normal")
        self.logbox.insert("end", message + "\n")
        self.logbox.see("end")
        self.logbox.configure(state="disabled")
        self.logbox.update()

    def _create_service(self) -> WorkflowService:
        root = Path(self.root_dir.get()).expanduser()
        service = WorkflowService(root)
        service.ensure_directories()
        return service

    # ------------------------------------------------------------------
    # Event handlers
    # ------------------------------------------------------------------
    def _on_browse(self) -> None:
        chosen = filedialog.askdirectory(title="Select project root")
        if chosen:
            self.root_dir.set(chosen)

    def _on_extract(self) -> None:
        service = self._create_service()
        try:
            summary = service.extract_for_review(progress=self._append_log)
        except NoInputFilesError as err:
            messagebox.showinfo(APP_TITLE, str(err))
        except ConfigError as err:
            messagebox.showerror(APP_TITLE, f"Config error. {err}")
        except WorkflowError as err:
            messagebox.showerror(APP_TITLE, f"Extraction failed. {err}")
        else:
            messagebox.showinfo(
                APP_TITLE, f"Review CSV created:\n{summary.review_csv}"
            )

    def _on_upload(self) -> None:
        service = self._create_service()
        try:
            summary = service.upload_to_access(progress=self._append_log)
        except ReviewFileMissingError as err:
            messagebox.showerror(APP_TITLE, str(err))
        except NoApprovedRowsError as err:
            messagebox.showinfo(APP_TITLE, str(err))
        except ConfigError as err:
            messagebox.showerror(APP_TITLE, f"Config error. {err}")
        except WorkflowError as err:
            messagebox.showerror(APP_TITLE, f"Upload failed. {err}")
        else:
            if summary.success:
                messagebox.showinfo(
                    APP_TITLE,
                    (
                        f"Bulk upload complete. Imported {summary.approved_count} rows.\n"
                        f"Access-ready CSV: {summary.access_ready_csv}\n"
                        f"Log: {summary.log_path}"
                    ),
                )
            else:
                error_text = f"Error: {summary.error}" if summary.error else ""
                messagebox.showerror(
                    APP_TITLE,
                    (
                        "Bulk upload failed. Approved PDFs moved to rejected.\n"
                        f"Access-ready CSV: {summary.access_ready_csv}\n"
                        f"Log: {summary.log_path}"
                        + (f"\n{error_text}" if error_text else "")
                    ),
                )

    def _on_test_access(self) -> None:
        service = self._create_service()
        try:
            db_path, msaccess_path = service.test_access()
        except ConfigError as err:
            messagebox.showerror(APP_TITLE, f"Config error. {err}")
        except WorkflowError as err:
            messagebox.showerror(APP_TITLE, f"Access connection failed. {err}")
        else:
            messagebox.showinfo(
                APP_TITLE,
                f"Access configuration looks OK.\nDB: {db_path}\nmsaccess: {msaccess_path}",
            )

    def _on_open_review(self) -> None:
        service = self._create_service()
        try:
            review_path = service.get_review_csv_path()
        except ConfigError as err:
            messagebox.showerror(APP_TITLE, f"Config error. {err}")
        except FileNotFoundError:
            messagebox.showinfo(APP_TITLE, "No review CSV found yet.")
        else:
            try:
                os.startfile(str(review_path))  # type: ignore[attr-defined]
            except Exception as err:  # pragma: no cover - platform specific
                messagebox.showerror(APP_TITLE, f"Could not open review CSV. {err}")


__all__ = ["PdfToAccessApp", "APP_TITLE"]
