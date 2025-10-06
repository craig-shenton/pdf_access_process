from __future__ import annotations

import textwrap
from pathlib import Path

import pytest

from core.config import ConfigError, load_config


def write_config(root: Path, contents: str) -> Path:
    config_dir = root / "config"
    config_dir.mkdir(parents=True, exist_ok=True)
    path = config_dir / "mapping.yml"
    path.write_text(textwrap.dedent(contents), encoding="utf-8")
    return path


def base_config() -> str:
    return """
    fields:
      - name: case_id
        find:
          type: keyword_line
          keywords: ["Case ID"]
    output:
      review_csv: review.csv
      access_ready_csv: access.csv
    access:
      db_path: /tmp/access.accdb
      column_map:
        case_id: CaseID
      bulk_import:
        msaccess_path: /tmp/MSACCESS.EXE
    """


def test_load_config_success(tmp_path: Path) -> None:
    write_config(tmp_path, base_config())
    config = load_config(tmp_path)

    assert config["output"]["review_csv"] == "review.csv"
    assert config["access"]["bulk_import"]["timeout_sec"] == 600
    assert config["fields"][0]["find"]["type"] == "keyword_line"


def test_load_config_missing_fields_section(tmp_path: Path) -> None:
    write_config(
        tmp_path,
        """
        output:
          review_csv: review.csv
          access_ready_csv: access.csv
        access:
          db_path: /tmp/access.accdb
          column_map: {case_id: CaseID}
          bulk_import:
            msaccess_path: /tmp/MSACCESS.EXE
        """,
    )

    with pytest.raises(ConfigError):
        load_config(tmp_path)


def test_load_config_requires_field_name(tmp_path: Path) -> None:
    write_config(
        tmp_path,
        """
        fields:
          - find:
              type: keyword_line
              keywords: ["Case ID"]
        output:
          review_csv: review.csv
          access_ready_csv: access.csv
        access:
          db_path: /tmp/access.accdb
          column_map: {case_id: CaseID}
          bulk_import:
            msaccess_path: /tmp/MSACCESS.EXE
        """,
    )

    with pytest.raises(ConfigError):
        load_config(tmp_path)
