"""Expose the ``src.ui`` package at the top level."""
from importlib import import_module
import sys

_pkg = import_module("src.ui")
sys.modules[__name__] = _pkg
