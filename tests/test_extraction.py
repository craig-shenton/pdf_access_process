from __future__ import annotations

from pathlib import Path

from core.extraction import extract_from_pdf, find_value_in_blocks


def test_find_value_in_blocks_regex() -> None:
    blocks = [
        (0, 0, 0, 0, "Case ID: 12345"),
        (0, 0, 0, 0, "Other text"),
    ]
    strategy = {"type": "regex", "pattern": r"Case ID:\s*(\d+)"}
    assert find_value_in_blocks(blocks, strategy) == "12345"


def test_find_value_in_blocks_keyword_line() -> None:
    blocks = [
        (0, 0, 0, 0, "Case Number: XYZ"),
    ]
    strategy = {"type": "keyword_line", "keywords": ["case number"]}
    assert find_value_in_blocks(blocks, strategy) == "XYZ"


def test_find_value_in_blocks_keyword_right() -> None:
    blocks = [
        (0, 0, 0, 0, "Name: John Doe"),
    ]
    strategy = {"type": "keyword_right", "keywords": ["name"]}
    assert find_value_in_blocks(blocks, strategy) == "John Doe"


def test_extract_from_pdf_uses_block_reader(tmp_path: Path) -> None:
    pdf_path = tmp_path / "dummy.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n...")

    def fake_reader(_path: Path):
        return [
            [
                (0, 0, 0, 0, "Case ID: 12345"),
                (0, 0, 0, 0, "Name: Jane Smith"),
            ]
        ]

    fields = [
        {"name": "case_id", "find": {"type": "regex", "pattern": r"Case ID:\s*(\d+)"}},
        {"name": "name", "find": {"type": "keyword_right", "keywords": ["name"]}},
    ]

    record = extract_from_pdf(pdf_path, fields, block_reader=fake_reader)

    assert record["_source_pdf"] == "dummy.pdf"
    assert record["case_id"] == "12345"
    assert record["name"] == "Jane Smith"
    assert record["_extraction_ok"] is True


def test_extract_from_pdf_handles_errors(tmp_path: Path) -> None:
    pdf_path = tmp_path / "dummy.pdf"
    pdf_path.write_bytes(b"%PDF-1.4\n...")

    def failing_reader(_path: Path):
        raise RuntimeError("boom")

    record = extract_from_pdf(pdf_path, [], block_reader=failing_reader)

    assert record["_extraction_ok"] is False
    assert "boom" in record["_notes"]
