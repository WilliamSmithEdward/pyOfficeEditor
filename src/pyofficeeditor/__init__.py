"""pyOfficeEditor: the document surface of Office files, in pure Python.

Where pyOpenVBA edits the VBA project inside an Office file, this library
edits the document itself: the cell, the paragraph, the slide, the table.
No Office installation, no COM, and no third-party packages.

What is here today is the foundation every host sits on:

- :mod:`pyofficeeditor.opc` -- Open Packaging Conventions: parts, content
  types, relationships.
- :mod:`pyofficeeditor._xml` -- XML that reproduces the bytes it was parsed
  from, so editing one cell leaves the rest of a worksheet untouched.
- :mod:`pyofficeeditor._zip` -- the ZIP container, read and written field
  for field.

The Excel surface is being built on top of these; Word, PowerPoint and
Access follow in that order.
"""

from __future__ import annotations

from pyofficeeditor._xml import Element, XmlDocument
from pyofficeeditor.exceptions import (
    PackageError,
    PyOfficeEditorError,
    UnsupportedFormatError,
    XmlError,
    ZipError,
)
from pyofficeeditor.opc import OpcPackage, Relationship, Relationships

__version__ = "0.1.0"

__all__ = [
    "Element",
    "OpcPackage",
    "PackageError",
    "PyOfficeEditorError",
    "Relationship",
    "Relationships",
    "UnsupportedFormatError",
    "XmlDocument",
    "XmlError",
    "ZipError",
    "__version__",
]
