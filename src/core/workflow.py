"""Workflow orchestration helpers used by the GUI."""
from __future__ import annotations

import shutil
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from integrations.access_bulk import import_csv_to_access
from output.review_writer import (
    REVIEW_STATUS_COLUMN,
    build_review_dataframe,
    load_review_dataframe,
    write_access_ready_csv,
    write_review_csv,
)

from .config import load_config
from .extraction import extract_from_pdf

IN_DIR = "input"
INBOX_DIR = f"{IN_DIR}/inbox"
ARCHIVE_DIR = f"{IN_DIR}/archive"
REJECTED_DIR = f"{IN_DIR}/rejected"
OUTPUT_DIR = "output"
LOGS_DIR = "logs"
PDF_EXTENSIONS = (".pdf",)

ProgressCallback = Callable[[str], None]


class WorkflowError(RuntimeError):
    """Base class for workflow errors."""


class NoInputFilesError(WorkflowError):
    """Raised when no PDFs are available for extraction."""

    def __init__(self, inbox: Path):
        super().__init__(f"No PDFs found in {inbox}")
        self.inbox = inbox


class ReviewFileMissingError(WorkflowError):
    """Raised when the review CSV file cannot be located."""

    def __init__(self, path: Path):
        super().__init__(f"Cannot find review CSV at {path}")
        self.path = path


class NoApprovedRowsError(WorkflowError):
    """Raised when no APPROVED rows exist in the review CSV."""

    def __init__(self) -> None:
        super().__init__("No APPROVED rows found. Review CSV must contain APPROVED rows before upload.")


@dataclass
class ExtractionSummary:
    """Information about a completed extraction run."""

    review_csv: Path
    access_ready_csv: Path
    rows: list[dict[str, Any]] = field(default_factory=list)


@dataclass
class UploadSummary:
    """Information about a bulk upload attempt."""

    access_ready_csv: Path
    log_path: Path
    approved_count: int
    moved_to_archive: list[Path] = field(default_factory=list)
    moved_to_rejected: list[Path] = field(default_factory=list)
    success: bool = False
    return_code: int | None = None
    error: Exception | None = None


class WorkflowService:
    """Coordinate extraction and Access upload tasks for the UI."""

    def __init__(self, root: Path, *, config_loader=load_config):
        self.root = Path(root)
        self._config_loader = config_loader

    # ------------------------------------------------------------------
    # Directory management
    # ------------------------------------------------------------------
    def ensure_directories(self) -> None:
        """Ensure that required project directories exist."""

        for relative in (
            IN_DIR,
            INBOX_DIR,
            ARCHIVE_DIR,
            REJECTED_DIR,
            OUTPUT_DIR,
            LOGS_DIR,
        ):
            (self.root / relative).mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------------
    # Configuration access
    # ------------------------------------------------------------------
    def _load_config(self) -> dict[str, Any]:
        return self._config_loader(self.root)

    # ------------------------------------------------------------------
    # Extraction
    # ------------------------------------------------------------------
    def extract_for_review(self, *, progress: ProgressCallback | None = None) -> ExtractionSummary:
        """Extract text from PDFs into the review and Access-ready CSVs."""

        callback = progress or (lambda _msg: None)
        config = self._load_config()
        inbox = self.root / INBOX_DIR

        pdfs = sorted(
            pdf
            for pdf in inbox.glob("**/*")
            if pdf.is_file() and pdf.suffix.lower() in PDF_EXTENSIONS
        )
        if not pdfs:
            raise NoInputFilesError(inbox)

        callback(f"Found {len(pdfs)} PDFs")
        rows: list[dict[str, Any]] = []
        for index, pdf in enumerate(pdfs, start=1):
            callback(f"[{index}/{len(pdfs)}] Extracting {pdf.name}")
            record = extract_from_pdf(pdf, config["fields"])
            rows.append(record)

        dataframe = build_review_dataframe(rows, config)
        outdir = self.root / OUTPUT_DIR
        review_csv = write_review_csv(outdir, config["output"]["review_csv"], dataframe)
        access_ready_csv, _ = write_access_ready_csv(
            dataframe, outdir, config["output"]["access_ready_csv"], config
        )

        callback(f"Review CSV written to: {review_csv}")
        callback(f"Access-ready CSV written to: {access_ready_csv}")

        return ExtractionSummary(review_csv=review_csv, access_ready_csv=access_ready_csv, rows=rows)

    # ------------------------------------------------------------------
    # Access upload
    # ------------------------------------------------------------------
    def upload_to_access(self, *, progress: ProgressCallback | None = None) -> UploadSummary:
        """Generate the Access-ready CSV and trigger the Access bulk upload."""

        callback = progress or (lambda _msg: None)
        config = self._load_config()
        outdir = self.root / OUTPUT_DIR

        review_csv = outdir / config["output"]["review_csv"]
        if not review_csv.exists():
            raise ReviewFileMissingError(review_csv)

        review_df = load_review_dataframe(review_csv)
        if REVIEW_STATUS_COLUMN not in review_df.columns:
            raise WorkflowError(
                f"{review_csv.name} is missing the {REVIEW_STATUS_COLUMN} column."
            )

        access_ready_csv, approved_rows = write_access_ready_csv(
            review_df,
            outdir,
            config["output"]["access_ready_csv"],
            config,
        )

        if approved_rows.empty:
            raise NoApprovedRowsError()

        bulk_cfg = config["access"]["bulk_import"]
        log_path = self.root / LOGS_DIR / f"bulk_import_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        log_path.parent.mkdir(parents=True, exist_ok=True)

        messages: list[str] = []

        def record(message: str) -> None:
            messages.append(message)
            callback(message)

        summary = UploadSummary(
            access_ready_csv=access_ready_csv,
            log_path=log_path,
            approved_count=len(approved_rows),
        )

        record(f"Preparing bulk import with {len(approved_rows)} approved rows")

        inbox = self.root / INBOX_DIR
        archive = self.root / ARCHIVE_DIR
        rejected = self.root / REJECTED_DIR
        archive.mkdir(parents=True, exist_ok=True)
        rejected.mkdir(parents=True, exist_ok=True)

        try:
            result = import_csv_to_access(
                msaccess_path=bulk_cfg["msaccess_path"],
                database_path=config["access"]["db_path"],
                csv_path=access_ready_csv,
                macro=bulk_cfg.get("macro"),
                timeout=bulk_cfg.get("timeout_sec", 600),
                use_cmd_argument=bulk_cfg.get("use_cmd_argument", True),
                extra_args=bulk_cfg.get("extra_args"),
                workdir=bulk_cfg.get("workdir"),
                log_path=log_path,
            )
            summary.success = True
            summary.return_code = result.returncode
            record(f"Access import completed with return code {result.returncode}")
        except Exception as exc:  # pragma: no cover - exercised via tests
            summary.success = False
            summary.error = exc
            record(f"Access import failed: {exc}")

        target_dir = archive if summary.success else rejected
        moved: list[Path] = []
        for _, row in approved_rows.iterrows():
            source_pdf = row.get("_source_pdf", "")
            if not source_pdf:
                continue
            src_path = inbox / str(source_pdf)
            if not src_path.exists():
                continue
            dest = target_dir / src_path.name
            try:
                dest.parent.mkdir(parents=True, exist_ok=True)
                shutil.move(str(src_path), str(dest))
                moved.append(dest)
                record(
                    f"{('ARCHIVED' if summary.success else 'REJECTED')} {src_path.name}"
                )
            except Exception as move_err:  # pragma: no cover - best effort
                record(f"Failed to move {src_path.name}: {move_err}")

        if summary.success:
            summary.moved_to_archive = moved
            summary.moved_to_rejected = []
        else:
            summary.moved_to_archive = []
            summary.moved_to_rejected = moved

        if messages:
            with open(log_path, "a", encoding="utf-8") as log_file:
                for message in messages:
                    log_file.write(message + "\n")

        return summary

    # ------------------------------------------------------------------
    # Utilities
    # ------------------------------------------------------------------
    def test_access(self) -> tuple[Path, Path]:
        """Validate that the configured Access dependencies exist."""

        config = self._load_config()
        db_path = Path(config["access"]["db_path"])
        msaccess_path = Path(config["access"]["bulk_import"]["msaccess_path"])
        issues = []
        if not db_path.exists():
            issues.append(f"Database not found: {db_path}")
        if not msaccess_path.exists():
            issues.append(f"msaccess.exe not found: {msaccess_path}")
        if issues:
            raise WorkflowError("; ".join(issues))
        return db_path, msaccess_path

    def get_review_csv_path(self) -> Path:
        """Return the path to the review CSV, ensuring it exists."""

        config = self._load_config()
        review_path = self.root / OUTPUT_DIR / config["output"]["review_csv"]
        if not review_path.exists():
            raise FileNotFoundError(review_path)
        return review_path
