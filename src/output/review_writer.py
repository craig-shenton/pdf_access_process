"""Helpers for preparing the CSV outputs used during review."""

from pathlib import Path
from typing import Iterable

import pandas as pd

REVIEW_STATUS_COLUMN = "_review_status"
REVIEW_COMMENT_COLUMN = "_review_comment"
DEDUPE_COLUMN = "_dedupe_key"
NOTES_COLUMN = "_notes"
SOURCE_PDF_COLUMN = "_source_pdf"
REVIEW_CSV_ENCODING = "utf-8-sig"
APPROVED_VALUE = "APPROVED"
PENDING_VALUE = "PENDING"


def build_review_dataframe(rows: Iterable[dict], cfg: dict) -> pd.DataFrame:
    """Return a DataFrame ready for review CSV export."""
    df = pd.DataFrame(rows)
    if df.empty:
        field_names = [field["name"] for field in cfg.get("fields", [])]
        columns = [SOURCE_PDF_COLUMN, "_extraction_ok", NOTES_COLUMN]
        df = pd.DataFrame(columns=columns + field_names)

    df = df.fillna("")

    dedupe_key = cfg.get("dedupe_key", [])
    if dedupe_key and all(col in df.columns for col in dedupe_key):
        df[DEDUPE_COLUMN] = df[dedupe_key].astype(str).agg("|".join, axis=1)
    else:
        df[DEDUPE_COLUMN] = df.get(DEDUPE_COLUMN, "")

    df[REVIEW_STATUS_COLUMN] = df.get(REVIEW_STATUS_COLUMN, PENDING_VALUE)
    df[REVIEW_COMMENT_COLUMN] = df.get(REVIEW_COMMENT_COLUMN, "")
    df[NOTES_COLUMN] = df.get(NOTES_COLUMN, "")

    ordered = [
        SOURCE_PDF_COLUMN,
        "_extraction_ok",
        DEDUPE_COLUMN,
        REVIEW_STATUS_COLUMN,
        REVIEW_COMMENT_COLUMN,
        NOTES_COLUMN,
    ]
    field_names = [field["name"] for field in cfg.get("fields", [])]
    columns = [c for c in ordered if c in df.columns] + [c for c in field_names if c in df.columns]
    df = df.reindex(columns=columns)

    return df


def write_review_csv(outdir: Path, filename: str, df: pd.DataFrame) -> Path:
    outdir.mkdir(parents=True, exist_ok=True)
    path = outdir / filename
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False, encoding=REVIEW_CSV_ENCODING)
    return path


def load_review_dataframe(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    return df
