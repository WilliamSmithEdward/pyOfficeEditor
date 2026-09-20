"""The version the package reports, and where it comes from.

``__version__`` was a literal in ``__init__.py`` for three releases and was
never updated, so 0.1.1, 0.2.0 and 0.2.1 all reported 0.1.0. Nothing linked
the string to ``pyproject.toml``, and nothing failed. These tests are that
link.
"""

from __future__ import annotations

from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

import pytest

# tomllib is stdlib and ships no stubs, which strict pyright objects to.
import tomllib  # pyright: ignore[reportMissingTypeStubs]

import pyofficeeditor

PYPROJECT = Path(__file__).resolve().parent.parent / "pyproject.toml"


def declared() -> str:
    """The version ``pyproject.toml`` declares, which the release workflow
    also checks the git tag against."""
    with PYPROJECT.open("rb") as handle:
        return tomllib.load(handle)["project"]["version"]


class TestTheReportedVersion:
    def test_it_matches_what_pyproject_declares(self) -> None:
        """The check that was missing. A literal in ``__init__.py`` drifted
        silently across three releases."""
        if not PYPROJECT.is_file():
            pytest.skip("not a source checkout, so there is nothing to compare against")
        try:
            installed = version("pyOfficeEditor")
        except PackageNotFoundError:
            pytest.skip("pyOfficeEditor is not installed, so there is no metadata to read")
        assert installed == declared(), (
            "the installed distribution and pyproject.toml disagree. Reinstall with "
            "`pip install -e .` after a version bump, or the package will report the "
            "version it was last installed at."
        )

    def test_the_package_reports_the_installed_version(self) -> None:
        try:
            installed = version("pyOfficeEditor")
        except PackageNotFoundError:
            pytest.skip("pyOfficeEditor is not installed")
        assert pyofficeeditor.__version__ == installed

    def test_it_is_read_from_the_metadata(self) -> None:
        """A guard against someone going back to a hardcoded string.

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
            line for line in source.splitlines() if line.strip().startswith('__version__ = "')
        ]
        assert literals == ['    __version__ = "0.0.0+source"'], literals

    def test_it_looks_like_a_version(self) -> None:
        assert pyofficeeditor.__version__
        assert pyofficeeditor.__version__[0].isdigit()

    def test_it_is_exported(self) -> None:
        assert "__version__" in pyofficeeditor.__all__
