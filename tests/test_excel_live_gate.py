# pyvbaharness ships no type stubs and is an optional extra, so strict
# inference cannot see through it here. The suppressions are scoped to this
# file, which is the only one that imports it.
# pyright: reportMissingTypeStubs=false, reportUnknownMemberType=false
# pyright: reportAttributeAccessIssue=false, reportUnknownArgumentType=false
"""The live gate: does Excel accept what this library writes?

Every other test compares bytes against bytes. This one asks the only
authority that matters. It opens a workbook this library wrote in real
Excel, checks that Excel did not have to repair it, and reads the cells back
through Excel's own object model.

It is the strongest available check on three things nothing else can prove:

- the package is valid, since Excel repairs or refuses one that is not
- ``fullCalcOnLoad`` works, since a changed input must produce a
  recalculated formula result rather than the stale cache still in the file
- a date written as a serial plus a number format really shows as a date

Opt in with ``RUN_LIVE_EXCEL=1``. It needs Windows, Excel, and the ``live``
extra, and it drives Excel through pyvbaharness, which holds a machine-wide
lock and kills the process it owns by recorded pid if anything wedges.
"""

from __future__ import annotations

import datetime as dt
import os
from pathlib import Path

import pytest

from pyofficeeditor.excel import CellError, Workbook

pytestmark = [
    pytest.mark.live,
    pytest.mark.skipif(
        os.environ.get("RUN_LIVE_EXCEL") != "1",
        reason="set RUN_LIVE_EXCEL=1 to drive real Excel",
    ),
]

#: Read back through Excel rather than through this library.
_READ_BACK = r"""
Public Function Probe(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim parts As String

    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Target)
    Set ws = wb.Worksheets("Data")

    ' Excel recalculates on load when fullCalcOnLoad is set. D2 is =B2*C2
    ' and B2 was changed from 120 to 200, so a correct file gives 850 and a
    ' stale cache gives 510.
    parts = "D2=" & CStr(ws.Range("D2").Value)
    parts = parts & "|D6=" & CStr(ws.Range("D6").Value)
    parts = parts & "|A2=" & ws.Range("A2").Value
    parts = parts & "|H1=" & ws.Range("H1").Value
    parts = parts & "|H2=" & CStr(ws.Range("H2").Value)
    parts = parts & "|H3=" & CStr(ws.Range("H3").Value)
    parts = parts & "|H4=" & Format(ws.Range("H4").Value, "yyyy-mm-dd")
    parts = parts & "|H4isdate=" & CStr(IsDate(ws.Range("H4").Value))
    parts = parts & "|H5=" & Format(ws.Range("H5").Value, "yyyy-mm-dd hh:nn")
    parts = parts & "|H7=" & CStr(ws.Range("H7").Text)
    parts = parts & "|A9=[" & ws.Range("A9").Value & "]"
    parts = parts & "|Z50=" & ws.Range("Z50").Value
    parts = parts & "|C2empty=" & CStr(IsEmpty(ws.Range("C2").Value))
    parts = parts & "|Notes=" & wb.Worksheets("Notes").Range("B5").Value
    parts = parts & "|sheets=" & CStr(wb.Worksheets.Count)

    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Probe = parts
End Function
"""


@pytest.fixture()
def written(live_sample_xlsx: Path, tmp_path: Path) -> Path:
    """A workbook this library edited, on disk, ready for Excel."""
    book = Workbook.open(live_sample_xlsx)
    sheet = book["Data"]
    sheet["B2"].value = 200
    sheet["H1"].value = "brand new string"
    sheet["H2"].value = 3.5
    sheet["H3"].value = True
    sheet["H4"].value = dt.date(2026, 7, 4)
    sheet["H5"].value = dt.datetime(2026, 7, 4, 13, 45)
    sheet["H7"].value = CellError("#N/A")
    sheet["A9"].value = "  still padded  "
    sheet["Z50"].value = "far away"
    sheet["C2"].clear()
    book["Notes"]["B5"].value = "on the second sheet"
    target = tmp_path / "edited.xlsx"
    book.save(target)
    return target


def _probe(path: Path) -> dict[str, str]:
    from pyvbaharness import ExcelSession

    with ExcelSession() as excel:
        excel.new_workbook()
        result = excel.run_vba(_READ_BACK, proc="Probe", args=(str(path),), timeout=180)
        assert result.outcome == "passed", f"Excel refused the workbook: {result!r}"
        raw = str(result.value)
    return dict(part.split("=", 1) for part in raw.split("|"))


def test_excel_opens_the_edited_workbook_and_recalculates(written: Path) -> None:
    """The gate. If Excel had to repair the file, opening it raises here."""
    seen = _probe(written)

    # C2 was cleared, so =B2*C2 is 200 * empty = 0. The point is that Excel
    # recalculated at all: the file still carries the cached 510.
    assert seen["D2"] == "0", f"D2 was not recalculated; Excel reported {seen['D2']}"
    assert seen["D6"] != "3265", "D6's cached total survived, so nothing recalculated"

    assert seen["A2"] == "North", "an existing shared string"
    assert seen["H1"] == "brand new string", "a shared string this library added"
    assert seen["H2"] == "3.5"
    assert seen["H3"] == "True"
    assert seen["H4"] == "2026-07-04"
    assert seen["H4isdate"] == "True", "the number format has to make it a date"
    assert seen["H5"] == "2026-07-04 13:45"
    assert seen["H7"] == "#N/A", "an error value, not the text of one"
    assert seen["A9"] == "[  still padded  ]", "xml:space held the spaces"
    assert seen["Z50"] == "far away"
    assert seen["C2empty"] == "True"
    assert seen["Notes"] == "on the second sheet"
    assert seen["sheets"] == "2"


def test_excel_opens_an_untouched_workbook(live_sample_xlsx: Path, tmp_path: Path) -> None:
    """A no-op save must produce a file Excel is equally happy with."""
    book = Workbook.open(live_sample_xlsx)
    target = tmp_path / "untouched.xlsx"
    target.write_bytes(book.to_bytes())
    assert target.read_bytes() == live_sample_xlsx.read_bytes(), "a no-op save changed bytes"
    seen = _probe(target)
    assert seen["D2"] == "510", "nothing changed, so the cached value stands"
    assert seen["A2"] == "North"
