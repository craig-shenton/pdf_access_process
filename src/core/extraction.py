"""PDF text extraction helpers."""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Callable, Mapping, Sequence


Block = Sequence[Any]
PageBlocks = Sequence[Block]
BlockReader = Callable[[Path], Sequence[PageBlocks]]


def read_pdf_text_blocks(pdf_path: Path) -> list[list[Any]]:
    """Return the text blocks for each page of a PDF document."""
    import fitz  # PyMuPDF

    pages: list[list[Any]] = []
    with fitz.open(pdf_path) as document:  # type: ignore[attr-defined]
        for page in document:
            blocks = page.get_text("blocks")
            pages.append(blocks)
    return pages


def find_value_in_blocks(blocks: PageBlocks, strategy: Mapping[str, Any]) -> str | None:
    """Extract a value from a set of blocks using the provided strategy."""

    strategy_type = strategy.get("type")
    if strategy_type == "regex":
        pattern = strategy.get("pattern")
        if not pattern:
            return None
        flags = re.IGNORECASE if strategy.get("ignore_case", True) else 0
        full_text = "\n".join(
            str(block[4])
            for block in blocks
            if isinstance(block, (list, tuple)) and len(block) >= 5
        )
        match = re.search(pattern, full_text, flags=flags)
        if match:
            return match.group(1) if match.groups() else match.group(0)
        return None

    keywords = [str(keyword).lower() for keyword in strategy.get("keywords", [])]

    if strategy_type == "keyword_line":
        for block in blocks:
            text = str(block[4]) if len(block) >= 5 else ""
            lower = text.lower()
            for keyword in keywords:
                if keyword in lower:
                    index = lower.find(keyword)
                    tail = text[index + len(keyword):].strip(" :\t\r\n")
                    if tail:
                        return tail
        return None

    if strategy_type == "keyword_right":
        for block in blocks:
            text = str(block[4]) if len(block) >= 5 else ""
            lower = text.lower()
            for keyword in keywords:
                if keyword in lower:
                    parts = re.split(re.escape(keyword), text, flags=re.IGNORECASE, maxsplit=1)
                    if len(parts) == 2:
                        right = parts[1].strip(" :\t\r\n")
                        first_line = right.splitlines()[0].strip()
                        if first_line:
                            return first_line
        return None

    return None


def extract_from_pdf(
    pdf_path: Path,
    fields: Sequence[Mapping[str, Any]],
    *,
    block_reader: Callable[[Path], Sequence[PageBlocks]] | None = None,
) -> dict[str, Any]:
    """Extract configured fields from ``pdf_path``.

    Parameters
    ----------
    pdf_path:
        The PDF document to parse.
    fields:
        Sequence of field configuration dictionaries containing ``name`` and
        ``find`` entries.
    block_reader:
        Optional callable overriding :func:`read_pdf_text_blocks`, useful for
        unit testing.
    """

    reader = block_reader or read_pdf_text_blocks
    record: dict[str, Any] = {"_source_pdf": Path(pdf_path).name}

    try:
        pages = reader(Path(pdf_path))
        for field in fields:
            name = field.get("name")
            strategy = field.get("find", {})
            value = None
            for blocks in pages:
                value = find_value_in_blocks(blocks, strategy)
                if value:
                    break
            if name:
                record[name] = value
        record["_extraction_ok"] = True
        record["_notes"] = record.get("_notes", "")
    except Exception as exc:  # pragma: no cover - exercised in tests via stub
        record["_extraction_ok"] = False
        record["_notes"] = f"Extraction error: {exc}"

    return record
