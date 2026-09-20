"""Shared fixtures.

The committed packages were authored by Excel; the generated ones are
authored by openpyxl, which writes the same parts a different legal way.
Testing against both is the point: a reader that only ever sees one
producer's output encodes that producer's habits as rules.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

FIXTURES = Path(__file__).parent / "fixtures"
EXCEL_FIXTURES = FIXTURES / "excel"

#: Authored by Excel: growth hints, mc:Ignorable, a CRLF after the
#: declaration.  The byte-fidelity gate.
MINIMAL_XLSM = EXCEL_FIXTURES / "excel_authored_minimal.xlsm"
#: Authored by Excel: 32 parts and a wider relationship graph.
POWERQUERY_XLSX = EXCEL_FIXTURES / "excel_authored_powerquery.xlsx"
#: Authored by Excel: the binary workbook format, whose worksheets are
#: .bin parts rather than XML.  Proves the container layer does not assume
#: its members are text.
BINARY_XLSB = EXCEL_FIXTURES / "excel_authored_binary.xlsb"
#: Authored by real Excel through pyvbaharness, on demand.  Tests that need
#: these skip when scripts/build_excel_fixtures.py has not been run.
LIVE_EMPTY_XLSX = EXCEL_FIXTURES / "empty.xlsx"
LIVE_SAMPLE_XLSX = EXCEL_FIXTURES / "sample.xlsx"
#: Authored by Excel: two ListObjects, one with a totals row and a calculated
#: column, plus defined names at both scopes, column widths and a hyperlink.
LIVE_STRUCTURES_XLSX = EXCEL_FIXTURES / "structures.xlsx"
#: Authored by Excel: every element that used to make an insertion refuse.
#: Data validation with formulas, a protected range, a saved sort,
#: scenarios, a shape in a drawing part, a form control and a comment,
#: the last two anchored through VML.
LIVE_REFUSED_XLSX = EXCEL_FIXTURES / "refused.xlsx"


@pytest.fixture(scope="session")
def minimal_xlsm_bytes() -> bytes:
    return MINIMAL_XLSM.read_bytes()


@pytest.fixture(scope="session")
def powerquery_xlsx_bytes() -> bytes:
    return POWERQUERY_XLSX.read_bytes()


@pytest.fixture(scope="session")
def binary_xlsb_bytes() -> bytes:
    return BINARY_XLSB.read_bytes()


@pytest.fixture(scope="session")
def excel_authored_packages() -> list[tuple[str, bytes]]:
    """Every committed Excel-authored package, for the fidelity gates."""
    return [(p.name, p.read_bytes()) for p in (MINIMAL_XLSM, POWERQUERY_XLSX, BINARY_XLSB)]


@pytest.fixture(scope="session")
def openpyxl_xlsx_bytes() -> bytes:
    """A workbook openpyxl wrote: a sharedStrings part, formula cells, and
    none of Excel's growth hints."""
    openpyxl = pytest.importorskip("openpyxl", reason="the dev extra provides openpyxl")
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Data"
    sheet["A1"] = "Region"
    sheet["B1"] = "Units"
    sheet["A2"] = "North"
    sheet["B2"] = 120
    sheet["A3"] = "South"
    sheet["B3"] = 340
    sheet["C2"] = "=B2*2"
    sheet["A5"] = "  padded  "
    sheet["A6"] = "amp & lt < gt >"
    workbook.create_sheet("Notes")["A1"] = "Second sheet"
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


@pytest.fixture(scope="session")
def data_descriptor_zip_bytes() -> bytes:
    """An archive whose members carry data descriptors.

    CPython's zipfile writes them when the output stream cannot seek, which
    is the shape a streaming producer emits: the local header zeroes the
    CRC and both sizes and only the central directory states the truth.
    """

    class Unseekable(io.RawIOBase):
        def __init__(self) -> None:
            self.buffer = bytearray()

        def writable(self) -> bool:
            return True

        def write(self, data: object) -> int:  # pyright: ignore[reportIncompatibleMethodOverride]
            payload = bytes(data)  # type: ignore[arg-type]
            self.buffer += payload
            return len(payload)

        def seekable(self) -> bool:
            return False

    sink = Unseekable()
    with zipfile.ZipFile(sink, "w", zipfile.ZIP_DEFLATED) as archive:  # type: ignore[arg-type]
        archive.writestr("[Content_Types].xml", b'<?xml version="1.0"?><Types/>')
        archive.writestr("a.xml", b"<a/>" * 100)
        archive.writestr("b.xml", b"<b/>" * 100)
    return bytes(sink.buffer)


@pytest.fixture(scope="session")
def live_sample_xlsx() -> Path:
    """The path to the Excel-authored sample, or a skip if it is absent.

    Session-scoped because it is a constant, and because the live gate's
    module-scoped fixtures request it: a narrower scope cannot be consumed by
    a wider one.
    """
    if not LIVE_SAMPLE_XLSX.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author sample.xlsx with real Excel")
    return LIVE_SAMPLE_XLSX


@pytest.fixture(scope="session")
def live_structures_xlsx() -> Path:
    """The Excel-authored workbook carrying tables and defined names."""
    if not LIVE_STRUCTURES_XLSX.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author structures.xlsx with real Excel")
    return LIVE_STRUCTURES_XLSX


@pytest.fixture(scope="session")
def live_refused_xlsx() -> Path:
    """The Excel-authored workbook carrying every once-refused element."""
    if not LIVE_REFUSED_XLSX.is_file():
        pytest.skip("run scripts/build_excel_fixtures.py to author refused.xlsx with real Excel")
    return LIVE_REFUSED_XLSX
