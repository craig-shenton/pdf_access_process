"""Configuration loading and validation helpers for the PDF access workflow."""
from __future__ import annotations

from pathlib import Path
from typing import Any, Iterable, Mapping

import yaml


class ConfigError(RuntimeError):
    """Raised when the YAML configuration cannot be loaded or validated."""


CONFIG_DIR = "config"
CONFIG_FILENAME = "mapping.yml"

_REQUIRED_TOP_LEVEL_KEYS = ("fields", "output", "access")
_REQUIRED_OUTPUT_KEYS = ("review_csv", "access_ready_csv")
_REQUIRED_ACCESS_KEYS = ("db_path", "column_map", "bulk_import")
_REQUIRED_BULK_IMPORT_KEYS = ("msaccess_path",)


def _require_keys(mapping: Mapping[str, Any], keys: Iterable[str], context: str) -> None:
    missing = [key for key in keys if key not in mapping]
    if missing:
        raise ConfigError(
            f"Missing {', '.join(missing)} in {context}. Check your mapping.yml file."
        )


def load_config(root: Path) -> dict[str, Any]:
    """Load and validate the application configuration from ``mapping.yml``.

    Parameters
    ----------
    root:
        Root directory of the project. The configuration is expected to live in
        ``<root>/config/mapping.yml``.

    Returns
    -------
    dict[str, Any]
        A dictionary containing the validated configuration data.

    Raises
    ------
    ConfigError
        If the configuration file is missing or does not satisfy the expected
        structure.
    """

    cfg_path = Path(root) / CONFIG_DIR / CONFIG_FILENAME
    if not cfg_path.exists():
        raise ConfigError(f"Cannot find configuration file at {cfg_path}")

    with open(cfg_path, "r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle) or {}

    if not isinstance(data, dict):
        raise ConfigError("mapping.yml must define a YAML mapping at the root level")

    _require_keys(data, _REQUIRED_TOP_LEVEL_KEYS, "mapping.yml")

    fields = data["fields"]
    if not isinstance(fields, list) or not fields:
        raise ConfigError("'fields' must be a non-empty list in mapping.yml")

    for index, field in enumerate(fields):
        if not isinstance(field, dict):
            raise ConfigError(f"Field entry #{index + 1} must be a mapping")
        _require_keys(field, ("name", "find"), f"fields[{index}]")
        find = field["find"]
        if not isinstance(find, dict):
            raise ConfigError(f"fields[{index}].find must be a mapping")
        if "type" not in find:
            raise ConfigError(f"fields[{index}].find is missing a 'type' entry")

    output = data["output"]
    if not isinstance(output, dict):
        raise ConfigError("'output' must be a mapping in mapping.yml")
    _require_keys(output, _REQUIRED_OUTPUT_KEYS, "output")

    access = data["access"]
    if not isinstance(access, dict):
        raise ConfigError("'access' must be a mapping in mapping.yml")
    _require_keys(access, _REQUIRED_ACCESS_KEYS, "access")

    column_map = access["column_map"]
    if not isinstance(column_map, dict) or not column_map:
        raise ConfigError("'access.column_map' must be a non-empty mapping")

    bulk_import = access["bulk_import"]
    if not isinstance(bulk_import, dict):
        raise ConfigError("'access.bulk_import' must be a mapping")
    _require_keys(bulk_import, _REQUIRED_BULK_IMPORT_KEYS, "access.bulk_import")

    bulk_import.setdefault("timeout_sec", 600)
    bulk_import.setdefault("use_cmd_argument", True)
    bulk_import.setdefault("extra_args", [])

    dedupe_key = data.get("dedupe_key", [])
    if dedupe_key and not isinstance(dedupe_key, list):
        raise ConfigError("'dedupe_key' must be a list when provided")
    data["dedupe_key"] = dedupe_key

    return data
