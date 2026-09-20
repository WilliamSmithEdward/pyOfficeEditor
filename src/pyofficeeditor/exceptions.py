"""The exception hierarchy shared by every layer.

Every parser and writer in this package raises from here, or raises
``ValueError`` for a caller-input bug.  No bare ``Exception`` and no
``KeyError`` reaches user code.
"""

from __future__ import annotations


class PyOfficeEditorError(Exception):
    """Root of every error this library raises."""


class UnsupportedFormatError(PyOfficeEditorError):
    """A file, extension or container feature this library does not handle."""


class ZipError(PyOfficeEditorError):
    """A malformed or unsupported ZIP container."""


class XmlError(PyOfficeEditorError):
    """Malformed XML, or an XML construct this library refuses to process."""


class PackageError(PyOfficeEditorError):
    """A malformed OPC package: a missing part, a broken relationship,
    a content type that cannot be resolved."""


__all__ = [
    "PackageError",
    "PyOfficeEditorError",
    "UnsupportedFormatError",
    "XmlError",
    "ZipError",
]
