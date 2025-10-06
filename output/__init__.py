"""Expose the ``src.output`` package at the top level."""
from importlib import import_module
import sys

_pkg = import_module("src.output")
sys.modules[__name__] = _pkg
