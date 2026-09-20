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

from importlib.metadata import PackageNotFoundError, version

from pyofficeeditor._xml import Element, XmlDocument
from pyofficeeditor.exceptions import (
    PackageError,
    PyOfficeEditorError,
    UnsupportedFormatError,
    XmlError,
    ZipError,
)
from pyofficeeditor.opc import OpcPackage, Relationship, Relationships

try:
    #: Read from the installed distribution rather than written here.
    #:
    #: A literal has to be remembered at every release and was not: 0.1.1,
    #: 0.2.0 and 0.2.1 all shipped saying 0.1.0, because nothing links a
    #: string in this file to the version in pyproject.toml. Asking the
    #: metadata cannot drift, and ``tests/test_version.py`` fails if the
    #: two ever disagree.
    __version__ = version("pyOfficeEditor")
except PackageNotFoundError:  # pragma: no cover - a source tree, uninstalled
    # Importable straight from a checkout, where there is no distribution to
    # ask. Saying so beats guessing a number.
    __version__ = "0.0.0+source"

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
