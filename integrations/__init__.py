"""Expose the ``src.integrations`` package at the top level."""
from importlib import import_module
import sys

_pkg = import_module("src.integrations")
sys.modules[__name__] = _pkg
