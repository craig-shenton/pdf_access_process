"""Core helpers for the PDF to Access application."""
from __future__ import annotations

from .config import ConfigError, load_config
from .extraction import extract_from_pdf, find_value_in_blocks, read_pdf_text_blocks

__all__ = [
    "ConfigError",
    "load_config",
    "extract_from_pdf",
    "find_value_in_blocks",
    "read_pdf_text_blocks",
    "NoApprovedRowsError",
    "NoInputFilesError",
    "ReviewFileMissingError",
    "WorkflowError",
    "ExtractionSummary",
    "UploadSummary",
    "WorkflowService",
]


# Lazy imports -----------------------------------------------------------------

def __getattr__(name: str):
    if name in {
        "NoApprovedRowsError",
        "NoInputFilesError",
        "ReviewFileMissingError",
        "WorkflowError",
        "ExtractionSummary",
        "UploadSummary",
        "WorkflowService",
    }:
        from .workflow import (  # noqa: WPS433 - imported lazily to avoid heavy deps
            ExtractionSummary,
            NoApprovedRowsError,
            NoInputFilesError,
            ReviewFileMissingError,
            UploadSummary,
            WorkflowError,
            WorkflowService,
        )

        return {
            "NoApprovedRowsError": NoApprovedRowsError,
            "NoInputFilesError": NoInputFilesError,
            "ReviewFileMissingError": ReviewFileMissingError,
            "WorkflowError": WorkflowError,
            "ExtractionSummary": ExtractionSummary,
            "UploadSummary": UploadSummary,
            "WorkflowService": WorkflowService,
        }[name]
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


def __dir__() -> list[str]:  # pragma: no cover - used for tooling
    return sorted(__all__)
