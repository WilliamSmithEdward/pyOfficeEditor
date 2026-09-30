"""The fuzz corpus, replayed through the fuzz targets on every run.

tests/fuzz_corpus/<target> seeds fuzz/fuzz_office.py; a fuzz finding joins
it as a regression seed. Each seed must be read, or refused with the
library's own error, and hold the target's round-trip where one applies.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pyofficeeditor import PyOfficeEditorError
from pyofficeeditor._xml import XmlDocument
from pyofficeeditor._zip import ZipArchive
from pyofficeeditor.excel import Workbook
from pyofficeeditor.excel._calc import parse
from pyofficeeditor.excel._calc.lexer import FormulaSyntaxError

CORPUS = Path(__file__).parent / "fuzz_corpus"
TARGETS = ("zip", "xml", "workbook", "formula")
SEEDS = [(target, path) for target in TARGETS for path in sorted((CORPUS / target).iterdir())]


def test_every_target_has_seeds() -> None:
    assert {target for target, _ in SEEDS} == set(TARGETS)


@pytest.mark.parametrize(("target", "path"), SEEDS, ids=[f"{t}/{p.name}" for t, p in SEEDS])
def test_a_seed_is_read_or_refused_cleanly(target: str, path: Path) -> None:
    data = path.read_bytes()
    if target == "zip":
        try:
            archive = ZipArchive.from_bytes(data)
        except PyOfficeEditorError:
            return
        written = archive.to_bytes()
        end = written.rfind(b"PK\x05\x06")
        if written[end - 20 : end - 16] == b"PK\x06\x07":
            return  # the format's zip64 ambiguity, as fuzz/fuzz_office.py explains
        again = ZipArchive.from_bytes(written)
        assert [(m.name, m.stored) for m in again.members()] == [
            (m.name, m.stored) for m in archive.members()
        ]
        assert again.to_bytes() == written
    elif target == "xml":
        try:
            document = XmlDocument.parse(data)
        except PyOfficeEditorError:
            return
        assert document.to_bytes() == data
    elif target == "workbook":
        try:
            book = Workbook.from_bytes(data)
            for sheet in book.sheets:
                list(sheet.cell_elements())
                _ = sheet.merged_ranges
        except PyOfficeEditorError:
            pass
    else:
        try:
            parse(data.decode("utf-8", errors="replace"))
        except FormulaSyntaxError:
            pass
