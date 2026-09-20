"""The version the package reports, and where it comes from.

``__version__`` was a literal in ``__init__.py`` for three releases and was
never updated, so 0.1.1, 0.2.0 and 0.2.1 all reported 0.1.0. Nothing linked
the string to ``pyproject.toml``, and nothing failed.

The link is closed in two places rather than one, and neither reads TOML:

- ``publish.yml`` refuses to release unless the git tag matches the version
  in ``pyproject.toml``, and the build writes that version into the
  distribution's metadata.
- these tests refuse a ``__version__`` that is not the metadata's.

So a wrong number cannot get past the tag check, and a right number cannot
be reported wrongly. The first draft of this file read ``pyproject.toml``
directly and broke the 3.10 job, because ``tomllib`` is stdlib only from
3.11 and this library supports 3.10.
"""

from __future__ import annotations

from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

import pytest

import pyofficeeditor


def installed() -> str:
    """What the distribution's own metadata says, or a skip.

    Running these from a checkout with nothing installed is legitimate, and
    there is no metadata to compare against then.
    """
    try:
        return version("pyOfficeEditor")
    except PackageNotFoundError:  # pragma: no cover - depends on the environment
        pytest.skip("pyOfficeEditor is not installed, so there is no metadata to read")


class TestTheReportedVersion:
    def test_it_is_the_installed_version(self) -> None:
        """The check that was missing. A literal in ``__init__.py`` drifted
        silently across three releases because nothing compared the two."""
        assert pyofficeeditor.__version__ == installed(), (
            "the attribute and the distribution metadata disagree. After a version "
            "bump, reinstall with `pip install -e .`; the package otherwise reports "
            "whatever it was last installed at."
        )

    def test_it_is_read_from_the_metadata(self) -> None:
        """A guard against going back to a hardcoded string.

        The source has to ask the distribution. There is one literal, the
        fallback for an uninstalled checkout, and it is deliberately not a
        version number.
        """
        source = Path(pyofficeeditor.__file__).read_text(encoding="utf-8")
        assert 'version("pyOfficeEditor")' in source, (
            "__version__ is no longer read from the installed metadata; a literal "
            "here goes stale at the next release, which is what happened to 0.1.1, "
            "0.2.0 and 0.2.1."
        )
        literals = [
            line.strip()
            for line in source.splitlines()
            if line.strip().startswith('__version__ = "')
        ]
        assert literals == ['__version__ = "0.0.0+source"'], literals

    def test_it_looks_like_a_version(self) -> None:
        assert pyofficeeditor.__version__
        assert pyofficeeditor.__version__[0].isdigit()

    def test_it_is_exported(self) -> None:
        assert "__version__" in pyofficeeditor.__all__
