"""Author the Excel test fixtures with real Excel.

The library's fidelity gates only mean something against bytes Excel
actually wrote.  openpyxl writes the same parts a different legal way: no
growth-hint extra fields, different ZIP version words, a different styles
part.  Both are worth testing against, and only one of them can be
generated without Office.

Run this on a Windows machine with Excel installed:

    python -m pip install -e ".[dev]" --group live
    python scripts/build_excel_fixtures.py          # only what is missing
    python scripts/build_excel_fixtures.py --force  # rebuild everything

A fixture that already exists is left alone. Excel does not produce the same
bytes twice, since it stamps every part with a fresh revision GUID, so
rebuilding a committed fixture churns it for no benefit and buries the real
change in a diff. ``--force`` is there for when a fixture's *content* needs to
change, which is a deliberate act.

Excel runs hidden and the harness terminates the instance it owns by recorded
process id even if a step wedges.
"""

from __future__ import annotations

import sys
from pathlib import Path

FIXTURES = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "excel"

#: XlFileFormat values.  Excel will not save .xlsx through the harness's
#: save_as, which only knows the macro-enabled formats, so the format is
#: named explicitly in VBA.
XL_OPEN_XML_WORKBOOK = 51

_BUILD_SAMPLE = r"""
Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Data"

    ' Text, so the workbook gets a real sharedStrings part.
    ws.Range("A1").Value = "Region"
    ws.Range("B1").Value = "Units"
    ws.Range("C1").Value = "Price"
    ws.Range("D1").Value = "Total"
    ws.Range("E1").Value = "Shipped"
    ws.Range("F1").Value = "Ordered"
    ws.Range("A1:F1").Font.Bold = True

    ws.Range("A2").Value = "North"
    ws.Range("A3").Value = "South"
    ws.Range("A4").Value = "East"
    ws.Range("A5").Value = "West"

    ' Numbers, including one that is deliberately not an integer.
    ws.Range("B2").Value = 120
    ws.Range("B3").Value = 340
    ws.Range("B4").Value = 95
    ws.Range("B5").Value = 210
    ws.Range("C2").Value = 4.25
    ws.Range("C3").Value = 4.25
    ws.Range("C4").Value = 5.5
    ws.Range("C5").Value = 3.75

    ' Formulas, so there is something with a cached value to read.
    ws.Range("D2:D5").Formula = "=B2*C2"
    ws.Range("D6").Formula = "=SUM(D2:D5)"

    ' Booleans and dates, the two types that do not look like numbers.
    ws.Range("E2").Value = True
    ws.Range("E3").Value = False
    ws.Range("E4").Value = True
    ws.Range("E5").Value = False
    ws.Range("F2").Value = DateSerial(2026, 1, 15)
    ws.Range("F3").Value = DateSerial(2026, 2, 1)
    ws.Range("F4").Value = DateSerial(2026, 2, 28)
    ws.Range("F5").Value = DateSerial(2026, 3, 9)
    ws.Range("F2:F5").NumberFormat = "yyyy-mm-dd"

    ' A string with characters that have to be escaped, and one with
    ' leading and trailing spaces that needs xml:space="preserve".
    ws.Range("A8").Value = "ampersand & angle < bracket > quote "" done"
    ws.Range("A9").Value = "  padded  "

    ' An inline formula error, so the error cell type appears.
    ws.Range("B8").Formula = "=1/0"

    ' A merged range and a second sheet, so structure is exercised too.
    ws.Range("A11:C11").Merge
    ws.Range("A11").Value = "Merged heading"

    wb.Worksheets.Add After:=ws
    wb.Worksheets(2).Name = "Notes"
    wb.Worksheets(2).Range("A1").Value = "Second sheet"

    ws.Activate
    ws.Range("A1").Select

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True

    Build = wb.FullName
End Function
"""

_BUILD_EMPTY = r"""
Public Function Build(ByVal Target As String) As String
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = ActiveWorkbook.FullName
End Function
"""

_BUILD_STRUCTURES = r"""
Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Tabled"

    ws.Range("A1").Value = "Region"
    ws.Range("B1").Value = "Units"
    ws.Range("C1").Value = "Revenue"
    ws.Range("A2").Value = "North"
    ws.Range("B2").Value = 120
    ws.Range("C2").Formula = "=B2*10"
    ws.Range("A3").Value = "South"
    ws.Range("B3").Value = 340
    ws.Range("C3").Formula = "=B3*10"
    ws.Range("A4").Value = "East"
    ws.Range("B4").Value = 95
    ws.Range("C4").Formula = "=B4*10"

    ' A real ListObject, with a totals row so the shape is complete.
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C4"), , xlYes)
    lo.Name = "SalesTable"
    lo.TableStyle = "TableStyleMedium2"
    lo.ShowTotals = True
    lo.ListColumns("Units").TotalsCalculation = xlTotalsCalculationSum

    ' A second, plainer table on another sheet, so two table parts exist.
    wb.Worksheets.Add After:=ws
    wb.Worksheets(2).Name = "Second"
    wb.Worksheets(2).Range("A1").Value = "Key"
    wb.Worksheets(2).Range("A2").Value = "k1"
    wb.Worksheets(2).ListObjects.Add(xlSrcRange, _
        wb.Worksheets(2).Range("A1:A2"), , xlYes).Name = "KeyTable"

    ' A workbook-scoped defined name and a sheet-scoped one, since renaming a
    ' sheet has to repoint both.
    wb.Names.Add Name:="TotalUnits", RefersTo:="=Tabled!$B$2:$B$4"
    ws.Names.Add Name:="LocalRegion", RefersTo:="=Tabled!$A$2"

    ' Column widths and a row height, which live outside sheetData.
    ws.Columns("A").ColumnWidth = 18
    ws.Rows(1).RowHeight = 24

    ' A hyperlink, which needs its own relationship from the sheet part.
    ws.Hyperlinks.Add Anchor:=ws.Range("E1"), Address:="https://example.invalid/", _
        TextToDisplay:="a link"

    ws.Activate
    ws.Range("A1").Select

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = wb.FullName
End Function
"""


def build(session: object, source: str, target: Path, label: str, *, force: bool) -> None:
    from pyvbaharness import ExcelSession

    assert isinstance(session, ExcelSession)
    if target.exists():
        if not force:
            print(f"  kept  {target.name} (already there; pass --force to rebuild)")
            return
        target.unlink()
    session.reset_sheets()
    result = session.run_vba(source, proc="Build", args=(str(target),), timeout=180)
    if result.outcome != "passed":
        raise SystemExit(f"{label}: Excel refused the build ({result.outcome}): {result!r}")
    if not target.is_file():
        raise SystemExit(f"{label}: Excel reported success but {target} is not there.")
    print(f"  wrote {target.relative_to(FIXTURES.parent.parent.parent)} ({target.stat().st_size} bytes)")


def main() -> int:
    try:
        from pyvbaharness import ExcelSession
    except ImportError:
        print(
            'pyvbaharness is not installed. Install the live extra:\n'
            '    python -m pip install -e ".[dev]" --group live',
            file=sys.stderr,
        )
        return 2

    force = "--force" in sys.argv
    wanted = [
        ("empty.xlsx", _BUILD_EMPTY),
        ("sample.xlsx", _BUILD_SAMPLE),
        ("structures.xlsx", _BUILD_STRUCTURES),
    ]
    if not force and all((FIXTURES / name).exists() for name, _ in wanted):
        print("every fixture is already there; nothing to do (pass --force to rebuild)")
        return 0

    FIXTURES.mkdir(parents=True, exist_ok=True)
    print(f"authoring fixtures in {FIXTURES}")
    with ExcelSession() as excel:
        for name, source in wanted:
            excel.new_workbook()
            build(excel, source, FIXTURES / name, name, force=force)
    print("done")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
