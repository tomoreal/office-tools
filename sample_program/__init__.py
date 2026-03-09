"""XBRLP: Lightweight Python XBRL Parser for Japanese financial data."""

from .file_loader import FileLoader
from .parser import Arc, Fact, Label, Parser, QName

__version__ = "0.1.0"
__all__ = ["Parser", "Fact", "QName", "Arc", "Label", "FileLoader"]