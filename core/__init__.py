"""Expose the ``src.core`` package as a top-level ``core`` package."""
from importlib import import_module
import sys

_pkg = import_module("src.core")

# Replace this module with the implementation from ``src.core``.
sys.modules[__name__] = _pkg
