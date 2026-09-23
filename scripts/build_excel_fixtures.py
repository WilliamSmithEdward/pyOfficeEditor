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

import html
import json
import os
import posixpath
import re
import struct
import sys
import zipfile
import zlib
from collections.abc import Callable
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
    print(f"  wrote {target.name} ({target.stat().st_size} bytes)")



#: Every element that used to make an insertion refuse, on one sheet.
#: Between them they cover all five address notations, including the two
#: zero-based ones and the two cases where an address is stored twice and
#: Excel refuses the workbook if the copies disagree.
_BUILD_REFUSED = r"""
Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "R"

    For i = 1 To 12
        ws.Cells(i, 1).Value = i
        ws.Cells(i, 2).Value = i * 3
        ws.Cells(i, 3).Value = "n" & i
    Next i
    ws.Range("E1").Value = "alpha"
    ws.Range("E2").Value = "beta"

    ' Data validation: a list from a range, a bound that is a reference,
    ' and a custom formula with a relative reference.
    With ws.Range("G2:G9").Validation
        .Delete
        .Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=$E$1:$E$2"
    End With
    With ws.Range("H2:H9").Validation
        .Delete
        .Add Type:=1, AlertStyle:=1, Operator:=1, Formula1:="1", Formula2:="=$B$5"
    End With
    With ws.Range("I2:I9,K2:K4").Validation
        .Delete
        .Add Type:=7, AlertStyle:=1, Operator:=1, Formula1:="=ISNUMBER($A2)"
    End With

    ws.Protection.AllowEditRanges.Add Title:="Editable", Range:=ws.Range("B2:B9")

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("B2:B9"), Order:=2
        .SetRange ws.Range("A1:C9")
        .Header = 1
        .Apply
    End With

    ws.Scenarios.Add Name:="High", ChangingCells:=ws.Range("A2:A3"), Values:=Array(99, 98)

    ' A shape, whose anchor lives in a drawing part.
    ws.Shapes.AddShape 1, ws.Range("E6").Left, ws.Range("E6").Top, 80, 40
    ' A form control, anchored inline and again in VML.
    ws.Buttons.Add ws.Range("E10").Left, ws.Range("E10").Top, 70, 24
    ' A comment, whose cell is in the comments part and again in VML.
    ws.Range("C3").AddComment "note here"

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = wb.FullName
End Function
"""


#: Sheet-level settings: protection with and without a password, a tab
#: colour, view flags, outline grouping, page setup and three visibility
#: states across four sheets.
_BUILD_SETTINGS = r"""
Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim other As Worksheet

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "S"
    ws.Range("A1:D20").Value = 1

    wb.Windows(1).Zoom = 85
    ws.Tab.Color = RGB(255, 0, 0)
    wb.Windows(1).DisplayGridlines = False
    wb.Windows(1).DisplayHeadings = False

    With ws.PageSetup
        .Orientation = 2
        .PaperSize = 9
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(1)
        .BottomMargin = Application.InchesToPoints(1)
        .HeaderMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.5)
        .CenterHorizontally = True
        .PrintArea = "$A$1:$D$20"
        .PrintTitleRows = "$1:$1"
        .PrintTitleColumns = "$A:$A"
        .LeftHeader = "left head"
        .CenterHeader = "&Bbold centre&B"
        .RightFooter = "Page &P of &N"
        .PrintGridlines = True
        .PrintHeadings = True
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 2
        .FirstPageNumber = 3
    End With

    ws.Rows("5:8").Group
    ws.Columns("B:C").Group

    Set other = wb.Worksheets.Add
    other.Name = "Hidden"
    other.Visible = 0
    Set other = wb.Worksheets.Add
    other.Name = "VeryHidden"
    other.Visible = 2

    ws.Protect Password:="secret", DrawingObjects:=True, Contents:=True, Scenarios:=True

    Set other = wb.Worksheets.Add
    other.Name = "Plain"
    other.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowFormattingCells:=True, AllowSorting:=True

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = wb.FullName
End Function
"""


#: Hyperlinks of every kind, an autofilter, and outline grouping.
_BUILD_LINKS = r"""
Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "L"
    ws.Range("A1").Value = "Region"
    ws.Range("B1").Value = "Units"
    ws.Range("C1").Value = "When"
    For i = 2 To 12
        ws.Cells(i, 1).Value = "r" & i
        ws.Cells(i, 2).Value = i * 2
        ws.Cells(i, 3).Value = DateSerial(2026, 1, i)
    Next i

    ws.Hyperlinks.Add Anchor:=ws.Range("E2"), Address:="https://example.com/a", _
        ScreenTip:="go there", TextToDisplay:="Example"
    ws.Hyperlinks.Add Anchor:=ws.Range("E3"), Address:="", SubAddress:="L!A1", _
        TextToDisplay:="top"
    ws.Hyperlinks.Add Anchor:=ws.Range("E4"), Address:="mailto:a@b.c", TextToDisplay:="mail"
    ws.Hyperlinks.Add Anchor:=ws.Range("E5"), Address:="https://example.com/b", _
        SubAddress:="frag", TextToDisplay:="both"
    ws.Hyperlinks.Add Anchor:=ws.Range("G2:G4"), Address:="https://example.com/c"

    ws.Range("A1:C12").AutoFilter Field:=1, Criteria1:=Array("r2", "r3"), Operator:=7
    ws.Range("A1:C12").AutoFilter Field:=2, Criteria1:=">=10"

    ws.Rows("5:8").Group
    ws.Rows("9:11").Group
    ws.Rows("5:8").EntireRow.Hidden = True
    ws.Columns("B:C").Group
    ws.Outline.SummaryRow = 0
    ws.Outline.SummaryColumn = 0

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = wb.FullName
End Function
"""


#: One shape of each MsoAutoShapeType the table names, so the preset
#: geometry mapping is held to what Excel writes rather than transcribed.
_BUILD_GEOMETRY = r"""
Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim s As Shape
    Dim kinds As Variant
    Dim i As Long

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "G"

    kinds = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, _
                  19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 92, 93, 94, 95, 96)
    For i = LBound(kinds) To UBound(kinds)
        On Error Resume Next
        Set s = ws.Shapes.AddShape(kinds(i), 10 + i * 5, 10, 40, 30)
        If Err.Number = 0 Then s.Name = "T" & kinds(i)
        Err.Clear
        On Error GoTo 0
    Next i

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = wb.FullName
End Function
"""

#: One of every Forms control that has something to say, wired up, so the
#: control reader is held to what Excel writes rather than to the schema.
#:
#: The committed shapes fixture carries a Button and nothing else, whose
#: whole part is ``objectType="Button" lockText="1"``. It cannot show that
#: a linked cell is read correctly because it has no linked cell, and it
#: cannot show a value at all.
#:
#: The measuring matters because the current value is spelled three
#: different ways and none of them is a single ``val`` attribute:
#:
#: - a check box and an option button write ``checked="Checked"`` or
#:   ``checked="Mixed"``, and write nothing when they are off
#: - a drop down and a list box write ``sel``, the 1-based selection,
#:   and leave ``val`` at 0
#: - a spinner and a scroll bar write ``val``
#:
#: So a reader that reports ``val`` answers 0 for a ticked box and 0 for a
#: drop down with the third item chosen. Both look plausible and both are
#: wrong. Excel's own object model answers -4146 for an unticked box,
#: which is ``xlOff`` and not 0, and that is the number recorded here.
_BUILD_CONTROLS = r"""
Private Function SafeLinked(ByVal s As Shape) As String
    On Error Resume Next
    SafeLinked = s.ControlFormat.LinkedCell
    If Err.Number <> 0 Then SafeLinked = ""
    Err.Clear
End Function

Private Function SafeList(ByVal s As Shape) As String
    On Error Resume Next
    SafeList = s.ControlFormat.ListFillRange
    If Err.Number <> 0 Then SafeList = ""
    Err.Clear
End Function

Private Function SafeValue(ByVal s As Shape) As String
    On Error Resume Next
    SafeValue = CStr(s.ControlFormat.Value)
    If Err.Number <> 0 Then SafeValue = ""
    Err.Clear
End Function

Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim s As Shape
    Dim out As String

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Controls"

    ws.Range("A1").Value = "Alpha"
    ws.Range("A2").Value = "Beta"
    ws.Range("A3").Value = "Gamma"

    ' A button: a macro, and nothing wired to a cell.
    Set s = ws.Shapes.AddFormControl(xlButtonControl, 200, 20, 90, 30)
    s.Name = "Go"
    s.TextFrame.Characters.Text = "Press"
    s.OnAction = "Clicked"

    ' A ticked check box with a linked cell: the case the consumer pins.
    Set s = ws.Shapes.AddFormControl(xlCheckBox, 200, 60, 110, 20)
    s.Name = "Tick"
    s.TextFrame.Characters.Text = "Enabled"
    s.ControlFormat.LinkedCell = "$D$6"
    s.ControlFormat.Value = xlOn

    ' An unticked one, which writes no checked attribute at all.
    Set s = ws.Shapes.AddFormControl(xlCheckBox, 200, 90, 110, 20)
    s.Name = "Untick"
    s.ControlFormat.LinkedCell = "$D$7"
    s.ControlFormat.Value = xlOff

    ' And the third state, which is neither.
    Set s = ws.Shapes.AddFormControl(xlCheckBox, 200, 120, 110, 20)
    s.Name = "Part"
    s.ControlFormat.Value = xlMixed

    ' A drop down with the second item chosen: sel moves, val stays 0.
    Set s = ws.Shapes.AddFormControl(xlDropDown, 200, 150, 110, 20)
    s.Name = "Pick"
    s.ControlFormat.ListFillRange = "$A$1:$A$3"
    s.ControlFormat.LinkedCell = "$D$8"
    s.ControlFormat.Value = 2

    ' A list box, wired the same way.
    Set s = ws.Shapes.AddFormControl(xlListBox, 200, 180, 110, 50)
    s.Name = "Many"
    s.ControlFormat.ListFillRange = "$A$1:$A$3"
    s.ControlFormat.LinkedCell = "$D$9"

    ' Two option buttons, so firstButton shows up on the first of them.
    Set s = ws.Shapes.AddFormControl(xlOptionButton, 200, 240, 110, 20)
    s.Name = "First"
    s.TextFrame.Characters.Text = "Choose me"
    s.ControlFormat.LinkedCell = "$D$10"
    s.ControlFormat.Value = xlOn

    Set s = ws.Shapes.AddFormControl(xlOptionButton, 200, 270, 110, 20)
    s.Name = "Second"
    s.TextFrame.Characters.Text = "Or me"

    ' A spinner and a scroll bar, which carry val and a range.
    Set s = ws.Shapes.AddFormControl(xlSpinner, 200, 300, 20, 30)
    s.Name = "Step"
    s.ControlFormat.LinkedCell = "$D$11"
    s.ControlFormat.Min = 0
    s.ControlFormat.Max = 50
    s.ControlFormat.Value = 7

    Set s = ws.Shapes.AddFormControl(xlScrollBar, 200, 340, 110, 20)
    s.Name = "Slide"
    s.ControlFormat.LinkedCell = "$D$12"
    s.ControlFormat.Min = 0
    s.ControlFormat.Max = 100
    s.ControlFormat.Value = 25

    ' A group box and a label, wired to nothing, to prove absence reads
    ' as absence rather than as zero.
    Set s = ws.Shapes.AddFormControl(xlGroupBox, 340, 240, 130, 50)
    s.Name = "Set"
    s.TextFrame.Characters.Text = "A group"

    Set s = ws.Shapes.AddFormControl(xlLabel, 340, 300, 110, 20)
    s.Name = "Caption"
    s.TextFrame.Characters.Text = "A label"

    ' An ordinary AutoShape too, so the reader is made to tell a drawing
    ' shape from a control on one sheet.
    Set s = ws.Shapes.AddShape(1, 20, 20, 80, 40)
    s.Name = "Plain"

    ' Saved before the measuring, deliberately. OnAction reports the macro
    ' qualified by the workbook holding it, so reading it from an unsaved
    ' book records "Book8!Clicked" and the number changes every rebuild.
    Application.DisplayAlerts = False
    wb.SaveAs Target, 52
    Application.DisplayAlerts = True

    For Each s In ws.Shapes
        out = out & s.Name & "|" & s.Type & "|" & s.OnAction & "|" _
            & SafeLinked(s) & "|" & SafeList(s) & "|" & SafeValue(s) & vbLf
    Next s

    Build = out
End Function
"""

#: One of every autofilter criterion, each on its own sheet.
#:
#: Each kind is a different child element rather than a variation on one,
#: so a model built from the two kinds a typical file carries would be
#: wrong about the other four. ``links.xlsx`` has a value list and one
#: comparison; this adds the rest.
#:
#: One kind per sheet, because stacking them fails: after three criteria
#: almost no rows are left showing and Excel refuses a top-ten call that
#: has nothing to rank. That is the probe's problem rather than a limit
#: worth recording, and separate sheets avoid it.
#:
#: What each sheet is for is the row count it hides as much as the markup.
#: Excel stores a filter twice, the criteria and a ``hidden`` flag on each
#: row, and does not recompute the first when the workbook opens.
_BUILD_FILTERS = r"""
Private Function Safe(ByVal f As Object, ByVal which As Long) As String
    On Error Resume Next
    If which = 1 Then Safe = CStr(f.Criteria1)
    If which = 2 Then Safe = CStr(f.Criteria2)
    If which = 3 Then Safe = CStr(f.Operator)
    If Err.Number <> 0 Then Safe = "n/a"
    Err.Clear
End Function

Private Sub Fill(ByVal ws As Worksheet)
    Dim i As Long
    ws.Range("A1").Value = "Name"
    ws.Range("B1").Value = "Units"
    ws.Range("C1").Value = "When"
    For i = 2 To 21
        ws.Cells(i, 1).Value = "r" & i
        ws.Cells(i, 2).Value = (i * 7) Mod 50
        ws.Cells(i, 3).Value = DateSerial(2026, ((i - 2) Mod 12) + 1, 5)
    Next i
    ' One blank, so a blanks filter has something to keep.
    ws.Range("A15").ClearContents
End Sub

Private Function Report(ByVal ws As Worksheet) As String
    Dim i As Long
    Dim f As Object
    Dim hidden As Long
    Dim out As String

    out = ws.Name & ";" & CStr(ws.AutoFilterMode)
    If ws.AutoFilterMode Then
        For i = 1 To ws.AutoFilter.Filters.Count
            Set f = ws.AutoFilter.Filters(i)
            If f.On Then
                out = out & ";" & CStr(i) & ":" & Safe(f, 1) & _
                      ":" & Safe(f, 2) & ":" & Safe(f, 3)
            End If
        Next i
    End If
    For i = 2 To 21
        If ws.Rows(i).Hidden Then hidden = hidden + 1
    Next i
    Report = out & ";hidden=" & CStr(hidden)
End Function

Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim out As String
    Dim names As Variant
    Dim i As Long

    Set wb = ActiveWorkbook
    names = Array("Values", "Compare", "Between", "TopTen", "Blanks", "Dynamic", _
                  "Numbers", "Dates")

    For i = LBound(names) To UBound(names)
        If wb.Worksheets.Count <= i Then wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count)
        Set ws = wb.Worksheets(i + 1)
        ws.Name = names(i)
        Fill ws
    Next i

    ' A value list.
    wb.Worksheets("Values").Range("A1:C21").AutoFilter _
        Field:=1, Criteria1:=Array("r2", "r3", "r4"), Operator:=7

    ' One comparison.
    wb.Worksheets("Compare").Range("A1:C21").AutoFilter Field:=2, Criteria1:=">=10"

    ' Two comparisons joined by And.
    wb.Worksheets("Between").Range("A1:C21").AutoFilter _
        Field:=2, Criteria1:=">=5", Operator:=1, Criteria2:="<=40"

    ' Top ten items, on a sheet nothing else has narrowed.
    wb.Worksheets("TopTen").Range("A1:C21").AutoFilter _
        Field:=2, Criteria1:="3", Operator:=3

    ' Blanks, which is its own spelling rather than a value.
    wb.Worksheets("Blanks").Range("A1:C21").AutoFilter Field:=1, Criteria1:="="

    ' Above average, which stores the average Excel worked out.
    wb.Worksheets("Dynamic").Range("A1:C21").AutoFilter Field:=2, Operator:=11, Criteria1:=33

    ' A value list over numbers, which writes the plain number as text.
    wb.Worksheets("Numbers").Range("A1:C21").AutoFilter _
        Field:=2, Criteria1:=Array("14", "21", "28"), Operator:=7

    ' And over dates, which through the object model keeps the text it was
    ' given rather than writing a dateGroupItem.
    wb.Worksheets("Dates").Range("A1:C21").AutoFilter _
        Field:=3, Criteria1:=Array("1/5/2026", "3/5/2026"), Operator:=7

    For i = 1 To wb.Worksheets.Count
        out = out & Report(wb.Worksheets(i)) & "|"
    Next i

    Application.DisplayAlerts = False
    wb.SaveAs Target, 51
    Application.DisplayAlerts = True

    Build = out
End Function
"""

#: Notes, which Excel's object model still calls comments: plain, on two
#: lines, showing, resized, on the first row and column, with a word in
#: bold, and one beside a form control, since the two share a VML part.
#:
#: The reply is one line per note, with what Excel's object model says of
#: it. The text comes back with its line breaks spelled ``\n``, because the
#: harness turns a control character into a space.
_BUILD_COMMENTS = r"""
Private Function Report(ByVal c As Comment) As String
    Report = c.Parent.Parent.Name & "!" & c.Parent.Address(False, False) & "|" & _
             Replace(c.Text, vbLf, "\n") & "|" & c.Author & "|" & CStr(c.Visible) & "|" & _
             CStr(c.Shape.Left) & "|" & CStr(c.Shape.Top) & "|" & _
             CStr(c.Shape.Width) & "|" & CStr(c.Shape.Height)
End Function

Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim both As Worksheet
    Dim c As Comment
    Dim out As String

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Notes"
    ws.Range("A1:D6").Value = 1

    ws.Range("C3").AddComment "plain note"
    Set c = ws.Range("E5").AddComment("two" & vbLf & "lines")
    c.Visible = True
    Set c = ws.Range("B8").AddComment("sized")
    c.Shape.Width = 200
    c.Shape.Height = 80
    ws.Range("A1").AddComment "corner"
    Set c = ws.Range("D10").AddComment("bold start")
    c.Shape.TextFrame.Characters(1, 4).Font.Bold = True

    Set both = wb.Worksheets.Add(After:=ws)
    both.Name = "Both"
    both.Buttons.Add 100, 30, 60, 20
    both.Range("B2").AddComment "beside a button"

    For Each ws In wb.Worksheets
        For Each c In ws.Comments
            out = out & Report(c) & vbLf
        Next c
    Next ws

    wb.Worksheets(1).Activate
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = out
End Function
"""

def _png(width: int, height: int, dpi: int | None) -> bytes:
    """A plain red PNG, with a ``pHYs`` chunk when ``dpi`` is given."""

    def chunk(kind: bytes, body: bytes) -> bytes:
        crc = zlib.crc32(kind + body) & 0xFFFFFFFF
        return struct.pack(">I", len(body)) + kind + body + struct.pack(">I", crc)

    rows = b"".join(b"\x00" + b"\xff\x00\x00" * width for _ in range(height))
    body = chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
    if dpi is not None:
        per_metre = round(dpi / 0.0254)
        body += chunk(b"pHYs", struct.pack(">IIB", per_metre, per_metre, 1))
    body += chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b"")
    return b"\x89PNG\r\n\x1a\n" + body


#: The one-pixel GIF every web page used to carry.
_GIF = (
    b"GIF89a\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff!\xf9\x04\x01\x00\x00\x00\x00"
    b",\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;"
)

#: Pictures as ``Shapes.AddPicture`` puts them in: at their own size, one
#: at 144 dots to the inch, one stretched, a GIF, and one with alternative
#: text, the same image twice among them. VBA writes the images itself from
#: hex, so the recipe needs nothing on disk.
_BUILD_PICTURES = r"""
Private Function Written(ByVal Name As String, ByVal Hex As String) As String
    Dim f As Integer
    Dim i As Long
    Dim data() As Byte
    ReDim data(Len(Hex) \ 2 - 1)
    For i = 0 To UBound(data)
        data(i) = CByte("&H" & Mid(Hex, i * 2 + 1, 2))
    Next i
    Written = Environ("TEMP") & "\" & Name
    f = FreeFile
    Open Written For Binary Access Write As #f
    Put #f, , data
    Close #f
End Function

Private Function Report(ByVal s As Shape) As String
    Report = s.Name & "|" & CStr(s.Type) & "|" & CStr(s.Left) & "|" & CStr(s.Top) & "|" & _
             CStr(s.Width) & "|" & CStr(s.Height) & "|" & s.AlternativeText & "|" & CStr(s.LockAspectRatio)
End Function

Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim s As Shape
    Dim plain As String
    Dim dense As String
    Dim dot As String
    Dim out As String

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Pictures"
    plain = Written("pyofficeeditor_plain.png", "{PLAIN}")
    dense = Written("pyofficeeditor_dense.png", "{DENSE}")
    dot = Written("pyofficeeditor_dot.gif", "{DOT}")

    Set s = ws.Shapes.AddPicture(plain, msoFalse, msoTrue, 100, 50, -1, -1)
    s.Name = "Natural"
    Set s = ws.Shapes.AddPicture(dense, msoFalse, msoTrue, 300, 50, -1, -1)
    s.Name = "Dense"
    Set s = ws.Shapes.AddPicture(plain, msoFalse, msoTrue, 300, 200, 60, 30)
    s.Name = "Stretched"
    Set s = ws.Shapes.AddPicture(dot, msoFalse, msoTrue, 100, 300, -1, -1)
    s.Name = "Dot"
    s.AlternativeText = "a single pixel"

    For Each s In ws.Shapes
        out = out & Report(s) & vbLf
    Next s
    Kill plain
    Kill dense
    Kill dot

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = out
End Function
""".replace("{PLAIN}", _png(120, 80, 96).hex().upper()).replace(
    "{DENSE}", _png(120, 80, 144).hex().upper()
).replace("{DOT}", _GIF.hex().upper())

#: Text in more than one font in one cell, made through ``Characters`` the
#: way a person makes it: a bold word, a red one, a bigger one, underline
#: and superscript, another typeface, a line break, and a cell whose own
#: font is bold. The reply is one line per cell with each font change Excel
#: reports: where it starts, then the typeface, size, bold, italic, colour,
#: underline and superscript.
_BUILD_RICHTEXT = r"""
Private Function Describe(ByVal c As Range) As String
    Dim i As Long
    Dim f As Font
    Dim out As String
    Dim last As String
    Dim now As String
    For i = 1 To Len(c.Value)
        Set f = c.Characters(i, 1).Font
        now = f.Name & "," & CStr(f.Size) & "," & CStr(f.Bold) & "," & CStr(f.Italic) & "," & _
              CStr(f.Color) & "," & CStr(f.Underline) & "," & CStr(f.Superscript)
        If now <> last Then out = out & CStr(i) & ":" & now & ";"
        last = now
    Next i
    Describe = out
End Function

Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long
    Dim out As String

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Rich"

    ws.Range("A1").Value = "bold plain"
    ws.Range("A1").Characters(1, 4).Font.Bold = True
    ws.Range("A2").Value = "plain red end"
    ws.Range("A2").Characters(7, 3).Font.Color = RGB(255, 0, 0)
    ws.Range("A3").Value = "big and italic"
    ws.Range("A3").Characters(1, 3).Font.Size = 16
    ws.Range("A3").Characters(9, 6).Font.Italic = True
    ws.Range("A4").Value = "under over"
    ws.Range("A4").Characters(1, 5).Font.Underline = xlUnderlineStyleSingle
    ws.Range("A4").Characters(7, 4).Font.Superscript = True
    ws.Range("A5").Value = "font change"
    ws.Range("A5").Characters(6, 6).Font.Name = "Courier New"
    ws.Range("A6").Value = "two" & vbLf & "lines bold"
    ws.Range("A6").Characters(11, 4).Font.Bold = True
    ws.Range("A6").WrapText = True
    ws.Range("A7").Value = "all bold"
    ws.Range("A7").Font.Bold = True
    ws.Range("A7").Characters(1, 3).Font.Italic = True

    For i = 1 To 7
        out = out & "A" & CStr(i) & "|" & Describe(ws.Cells(i, 1)) & vbLf
    Next i

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = out
End Function
"""

#: Text XML cannot carry as it is, everywhere a workbook keeps text: cells
#: with control characters, a lone carriage return, a CRLF and a literal
#: "_x0041_", formula results and a formula naming a sheet called
#: "a_x0041_b", a note, a validation's messages, a hyperlink's tip, a
#: table's headers, a page header and a defined name's comment. The reply
#: is one line per piece of text: what it is, then the UTF-16 code of each
#: character Excel reports for it, in hex, so no character is lost on the
#: way. Triple single quotes, because the formulas hold three double ones.
_BUILD_ESCAPES = r'''
Private Function Codes(ByVal s As String) As String
    Dim i As Long
    Dim out As String
    For i = 1 To Len(s)
        out = out & Hex(AscW(Mid(s, i, 1))) & " "
    Next i
    Codes = Trim(out)
End Function

Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim other As Worksheet
    Dim lo As ListObject
    Dim nm As Name
    Dim texts As Variant
    Dim headers As Variant
    Dim i As Long
    Dim out As String

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Text"
    Set other = wb.Worksheets.Add(After:=ws)
    other.Name = "a_x0041_b"
    other.Range("A1").Value = 7

    texts = Array("a" & Chr(1) & "b", "tab" & vbTab & "x", "cr" & vbCr & "x", "crlf" & vbCrLf & "x", _
                  "lf" & vbLf & "x", "_x0041_", "_X0041_", "_x004_", "a_x005F_b", _
                  "bell" & Chr(7) & "end" & Chr(31), "high" & ChrW(&HFFFE) & "x", "_x0001_" & Chr(1), _
                  "__x0041_", "_x00e9_")
    For i = 0 To UBound(texts)
        ws.Cells(i + 1, 1).Value = texts(i)
    Next i

    ws.Range("B1").Formula = "=""a""&CHAR(1)&""b""&CHAR(10)&""c""&CHAR(13)&""d"""
    ws.Range("B2").Formula = "=""_x0041_"""
    ws.Range("B3").Formula = "='a_x0041_b'!A1"

    ws.Range("C1").AddComment "note" & Chr(1) & "x" & vbLf & "y" & vbCr & "z _x0041_"

    With ws.Range("D1").Validation
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="1", Formula2:="9"
        .InputTitle = "ti" & Chr(1) & "t"
        .InputMessage = "in" & vbLf & "put" & Chr(1) & " _x0041_"
        .ErrorMessage = "er" & vbCr & "ror"
    End With

    ws.Hyperlinks.Add Anchor:=ws.Range("E1"), Address:="https://example.com/", _
                      ScreenTip:="tip" & Chr(1) & vbLf & "x", TextToDisplay:="link"

    headers = Array("head" & vbLf & "two", "_x0041_", "plain", "ctl" & Chr(1) & "x", _
                    "tab" & vbTab & "x", "cr" & vbCr & "x")
    For i = 0 To UBound(headers)
        ws.Cells(1, 7 + i).Value = headers(i)
        ws.Cells(2, 7 + i).Value = i
    Next i
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 7), ws.Cells(2, 7 + UBound(headers))), , xlYes)

    ws.PageSetup.CenterHeader = "head" & vbLf & "er _x0041_"

    Set nm = wb.Names.Add(Name:="nm", RefersTo:="=1")
    nm.Comment = "c" & vbLf & "_x0041_"

    For i = 0 To UBound(texts)
        out = out & "Text!A" & CStr(i + 1) & "|" & Codes(CStr(ws.Cells(i + 1, 1).Value)) & vbLf
    Next i
    For i = 1 To 3
        out = out & "Text!B" & CStr(i) & "|" & Codes(CStr(ws.Cells(i, 2).Value)) & vbLf
        out = out & "Text!B" & CStr(i) & " formula|" & Codes(ws.Cells(i, 2).Formula) & vbLf
    Next i
    out = out & "note|" & Codes(ws.Range("C1").Comment.Text) & vbLf
    out = out & "validation title|" & Codes(ws.Range("D1").Validation.InputTitle) & vbLf
    out = out & "validation message|" & Codes(ws.Range("D1").Validation.InputMessage) & vbLf
    out = out & "validation error|" & Codes(ws.Range("D1").Validation.ErrorMessage) & vbLf
    out = out & "hyperlink tip|" & Codes(ws.Hyperlinks(1).ScreenTip) & vbLf
    For i = 1 To lo.ListColumns.Count
        out = out & "table column " & CStr(i) & "|" & Codes(lo.ListColumns(i).Name) & vbLf
        out = out & "table header " & CStr(i) & "|" & Codes(CStr(lo.HeaderRowRange.Cells(1, i).Value)) & vbLf
    Next i
    out = out & "page header|" & Codes(ws.PageSetup.CenterHeader) & vbLf
    out = out & "sheet name|" & Codes(other.Name) & vbLf
    out = out & "name comment|" & Codes(nm.Comment) & vbLf

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = out
End Function
'''

#: Every built-in cell style Excel lists, each applied to a cell of its own
#: in column B beside its name, a style of the workbook's own in D1, and
#: two styles put over formatting a cell already had: "Good" over a bold,
#: centred cell showing two decimals, and "Currency" then italic. The reply
#: is a line per cell describing what Excel shows, and a line per built-in
#: style saying which aspects it sets.
_BUILD_STYLES = r'''
Private Function Describe(ByVal c As Range) As String
    Dim s As String
    Dim edges As Variant
    Dim e As Variant
    s = c.Style.Name & "|" & c.NumberFormat & "|" & c.Font.Name & "|" & CStr(c.Font.Size) & "|" & _
        CStr(c.Font.Bold) & "|" & CStr(c.Font.Italic) & "|" & CStr(c.Font.Color) & "|" & _
        CStr(c.Interior.Pattern) & "|" & CStr(c.Interior.Color) & "|" & CStr(c.HorizontalAlignment)
    edges = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
    For Each e In edges
        s = s & "|" & CStr(c.Borders(e).LineStyle) & "," & CStr(c.Borders(e).Weight) & "," & CStr(c.Borders(e).Color)
    Next e
    Describe = s
End Function

Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim st As Style
    Dim mine As Style
    Dim i As Long
    Dim r As Long
    Dim out As String

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Styles"
    For Each st In wb.Styles
        If st.BuiltIn And st.Name <> "Normal" Then
            i = i + 1
            ws.Cells(i, 1).Value = st.Name
            ws.Cells(i, 2).Style = st.Name
            ws.Cells(i, 2).Value = 1234.5
            out = out & "includes|" & st.Name & "|" & CStr(st.IncludeNumber) & "," & CStr(st.IncludeFont) & "," & _
                  CStr(st.IncludePatterns) & "," & CStr(st.IncludeBorder) & "," & CStr(st.IncludeAlignment) & "," & _
                  CStr(st.IncludeProtection) & vbLf
        End If
    Next st

    Set mine = wb.Styles.Add("Mine")
    mine.IncludeNumber = False
    mine.IncludeAlignment = False
    mine.IncludeBorder = False
    mine.IncludeProtection = False
    mine.Font.Bold = True
    mine.Font.Color = RGB(0, 0, 255)
    mine.Interior.Color = RGB(255, 255, 0)
    ws.Range("D1").Style = "Mine"
    ws.Range("D1").Value = "mine"

    ws.Range("D2").Value = 1.5
    ws.Range("D2").NumberFormat = "0.00"
    ws.Range("D2").Font.Bold = True
    ws.Range("D2").HorizontalAlignment = xlCenter
    ws.Range("D2").Style = "Good"
    ws.Range("D3").Value = 2.5
    ws.Range("D3").Style = "Currency"
    ws.Range("D3").Font.Italic = True

    For r = 1 To i
        out = out & "B" & CStr(r) & "|" & Describe(ws.Cells(r, 2)) & vbLf
    Next r
    For r = 1 To 3
        out = out & "D" & CStr(r) & "|" & Describe(ws.Cells(r, 4)) & vbLf
    Next r

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    Application.DisplayAlerts = True
    Build = out
End Function
'''

#: Charts on a data sheet, on another sheet and on a chart sheet: columns
#: with a typed title, a line over two areas, a pie whose title is linked
#: to a cell, a scatter, and bars. After the workbook is saved, each edit a
#: chart's references have to follow is made on a fresh copy by Excel and
#: saved to the temporary folder. The references are read from those files
#: rather than from Excel's object model, which reports a deleted reference
#: by its old address until the file is reopened, and then refuses to
#: report the series at all. The reply is each chart's series formulas and
#: title before any edit, and where each edited file is.
_BUILD_CHARTS = r'''
Private Function Report(ByVal wb As Workbook, ByVal label As String) As String
    Dim out As String
    Dim co As ChartObject
    Dim ws As Worksheet
    Dim ch As Chart
    For Each ws In wb.Worksheets
        For Each co In ws.ChartObjects
            out = out & label & "|" & ws.Name & "/" & co.Name & "|" & Series(co.Chart) & vbLf
        Next co
    Next ws
    For Each ch In wb.Charts
        out = out & label & "|" & ch.Name & "|" & Series(ch) & vbLf
    Next ch
    Report = out
End Function

Private Function Series(ByVal ch As Chart) As String
    Dim s As Series
    Dim out As String
    For Each s In ch.SeriesCollection
        out = out & s.Formula & ";"
    Next s
    out = out & "|"
    If ch.HasTitle Then out = out & ch.ChartTitle.Formula
    Series = out
End Function

Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rep As Worksheet
    Dim co As ChartObject
    Dim cs As Chart
    Dim r As Long
    Dim out As String
    Dim ops As Variant
    Dim i As Long
    Dim again As Workbook
    Dim edited As String
    Dim base As String

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Data"
    ws.Range("A1:D1").Value = Array("Month", "Sales", "Costs", "Profit")
    For r = 2 To 6
        ws.Cells(r, 1).Value = "M" & CStr(r - 1)
        ws.Cells(r, 2).Value = r * 10
        ws.Cells(r, 3).Value = r * 4
        ws.Cells(r, 4).Formula = "=B" & CStr(r) & "-C" & CStr(r)
    Next r

    Set co = ws.ChartObjects.Add(Left:=300, Top:=10, Width:=300, Height:=200)
    co.Name = "Columns"
    co.Chart.ChartType = xlColumnClustered
    co.Chart.SetSourceData Source:=ws.Range("A1:C6")
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = "Plain title"

    Set co = ws.ChartObjects.Add(Left:=300, Top:=220, Width:=300, Height:=200)
    co.Name = "Line"
    co.Chart.ChartType = xlLine
    co.Chart.SetSourceData Source:=ws.Range("A1:A6,D1:D6")

    Set rep = wb.Worksheets.Add(After:=ws)
    rep.Name = "Report"
    Set co = rep.ChartObjects.Add(Left:=10, Top:=10, Width:=300, Height:=200)
    co.Name = "Pie"
    co.Chart.ChartType = xlPie
    co.Chart.SetSourceData Source:=ws.Range("A1:B6")
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Formula = "=Data!$B$1"

    Set co = rep.ChartObjects.Add(Left:=10, Top:=220, Width:=300, Height:=200)
    co.Name = "Scatter"
    co.Chart.ChartType = xlXYScatter
    co.Chart.SetSourceData Source:=ws.Range("B1:C6")

    Set cs = wb.Charts.Add(After:=wb.Sheets(wb.Sheets.Count))
    cs.Name = "Bars"
    cs.ChartType = xlBarClustered
    cs.SetSourceData Source:=ws.Range("A1:B6")

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51

    ' The fixture stays open as the active workbook, so everything is read
    ' from a copy of it opened afresh, as each edit's result is. The copies
    ' go in the temporary folder, so a build that fails leaves none beside
    ' the fixtures.
    base = Environ("TEMP") & "\pyofficeeditor_charts_base.xlsx"
    wb.SaveCopyAs base
    Set again = Workbooks.Open(base)
    out = Report(again, "base")
    again.Close SaveChanges:=False
    out = out & "file|base|" & base & vbLf
    ops = Array("insert rows 3:4", "insert row 1", "delete row 4", "delete rows 2:6", "delete row 1", _
                "insert column B", "delete column C", "delete column A", "rename Data")
    For i = 0 To UBound(ops)
        Set again = Workbooks.Open(base)
        Select Case ops(i)
            Case "insert rows 3:4": again.Worksheets("Data").Rows("3:4").Insert
            Case "insert row 1": again.Worksheets("Data").Rows(1).Insert
            Case "delete row 4": again.Worksheets("Data").Rows(4).Delete
            Case "delete rows 2:6": again.Worksheets("Data").Rows("2:6").Delete
            Case "delete row 1": again.Worksheets("Data").Rows(1).Delete
            Case "insert column B": again.Worksheets("Data").Columns("B").Insert
            Case "delete column C": again.Worksheets("Data").Columns("C").Delete
            Case "delete column A": again.Worksheets("Data").Columns("A").Delete
            Case "rename Data": again.Worksheets("Data").Name = "Q1 Data"
        End Select
        edited = Environ("TEMP") & "\pyofficeeditor_charts_" & CStr(i) & ".xlsx"
        again.SaveAs Filename:=edited, FileFormat:=51
        again.Close SaveChanges:=False
        out = out & "file|" & ops(i) & "|" & edited & vbLf
    Next i
    Application.DisplayAlerts = True
    Build = out
End Function
'''

#: Data!A1:C11 feeding two pivot tables, each from a cache of its own:
#: Summary on a sheet of its own at A3, and Beside on the data sheet at F3,
#: where the edits reach it. Each edit is made by Excel on a fresh copy and
#: saved to the temporary folder, or reported refused; the pivot tables are
#: read from those files. The reply is where each file is.
_BUILD_PIVOTS = r'''
Public Function Build(ByVal Target As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pv As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim r As Long
    Dim ops As Variant
    Dim i As Long
    Dim again As Workbook
    Dim out As String
    Dim base As String
    Dim edited As String

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "Data"
    ws.Range("A1:C1").Value = Array("Region", "Product", "Sales")
    For r = 2 To 11
        ws.Cells(r, 1).Value = Choose((r Mod 3) + 1, "North", "South", "West")
        ws.Cells(r, 2).Value = Choose((r Mod 2) + 1, "Tea", "Coffee")
        ws.Cells(r, 3).Value = r * 10
    Next r
    Set pv = wb.Worksheets.Add(After:=ws)
    pv.Name = "Pivot"
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data!R1C1:R11C3")
    Set pt = pc.CreatePivotTable(TableDestination:=pv.Range("A3"), TableName:="Summary")
    pt.PivotFields("Region").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Sales"), "Total", xlSum
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data!R1C1:R11C3")
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("F3"), TableName:="Beside")
    pt.PivotFields("Product").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Sales"), "Sum", xlSum

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=Target, FileFormat:=51
    base = Environ("TEMP") & "\pyofficeeditor_pivots_base.xlsx"
    wb.SaveCopyAs base
    out = "file|base|" & base & vbLf

    ops = Array("insert row 1", "insert row 3", "insert row 7", "insert rows 5:6", "delete row 3", _
                "delete row 4", "delete rows 3:6", "delete rows 2:7", "delete rows 1:11", "delete rows 1:2", _
                "delete row 7", "insert column A", "insert column F", "insert column H", "delete column B", _
                "delete column F", "delete columns E:H", "insert row 1 on Pivot", "rename Data")
    For i = 0 To UBound(ops)
        Set again = Workbooks.Open(base)
        Set ws = again.Worksheets("Data")
        On Error Resume Next
        Select Case ops(i)
            Case "insert row 1": ws.Rows(1).Insert
            Case "insert row 3": ws.Rows(3).Insert
            Case "insert row 7": ws.Rows(7).Insert
            Case "insert rows 5:6": ws.Rows("5:6").Insert
            Case "delete row 3": ws.Rows(3).Delete
            Case "delete row 4": ws.Rows(4).Delete
            Case "delete rows 3:6": ws.Rows("3:6").Delete
            Case "delete rows 2:7": ws.Rows("2:7").Delete
            Case "delete rows 1:11": ws.Rows("1:11").Delete
            Case "delete rows 1:2": ws.Rows("1:2").Delete
            Case "delete row 7": ws.Rows(7).Delete
            Case "insert column A": ws.Columns("A").Insert
            Case "insert column F": ws.Columns("F").Insert
            Case "insert column H": ws.Columns("H").Insert
            Case "delete column B": ws.Columns("B").Delete
            Case "delete column F": ws.Columns("F").Delete
            Case "delete columns E:H": ws.Columns("E:H").Delete
            Case "insert row 1 on Pivot": again.Worksheets("Pivot").Rows(1).Insert
            Case "rename Data": ws.Name = "Q1 Data"
        End Select
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            again.Close SaveChanges:=False
            out = out & "refused|" & ops(i) & vbLf
        Else
            On Error GoTo 0
            edited = Environ("TEMP") & "\pyofficeeditor_pivots_" & CStr(i) & ".xlsx"
            again.SaveAs Filename:=edited, FileFormat:=51
            again.Close SaveChanges:=False
            out = out & "file|" & ops(i) & "|" & edited & vbLf
        End If
    Next i
    Application.DisplayAlerts = True
    Build = out
End Function
'''

#: What each field of a described cell is.
_STYLE_FIELDS = (
    "style", "number_format", "font_name", "font_size", "bold", "italic", "font_color",
    "pattern", "fill_color", "horizontal", "left", "top", "bottom", "right",
)

#: Which fields of the reported line mean what.
_CONTROL_FIELDS = ("type", "macro", "linked_cell", "list_range", "value")

#: The same for a picture, after its name.
_PICTURE_FIELDS = ("type", "left", "top", "width", "height", "description", "locks_aspect")

#: The same for a note, after its sheet-qualified address.
_COMMENT_FIELDS = ("text", "author", "visible", "left", "top", "width", "height")


#: What a measured fixture's reply turns into: an entry per thing measured.
Answers = dict[str, dict[str, object]]


def build_answers(
    session: object,
    source: str,
    target: Path,
    label: str,
    *,
    force: bool,
    parse: Callable[[str], Answers],
) -> None:
    """Author a fixture and record what Excel's object model said about it.

    The measurements are the point of this one, so they are written beside
    the workbook instead of being read once and remembered. A reader that
    disagrees with them is wrong about Excel, which is the only authority
    that settles it. ``parse`` reads the fixture's own reply.
    """
    from pyvbaharness import ExcelSession

    assert isinstance(session, ExcelSession)
    answers_path = target.with_name(f"{target.stem}_answers.json")
    if target.exists() and answers_path.exists():
        if not force:
            print(f"  kept  {target.name} (already there; pass --force to rebuild)")
            return
        target.unlink()
        answers_path.unlink(missing_ok=True)

    session.reset_sheets()
    result = session.run_vba(source, proc="Build", args=(str(target),), timeout=180)
    if result.outcome != "passed":
        raise SystemExit(f"{label}: Excel refused the build ({result.outcome}): {result!r}")
    if not target.is_file():
        raise SystemExit(f"{label}: Excel reported success but {target} is not there.")

    answers = parse(str(result.value or ""))
    answers_path.write_text(
        json.dumps(answers, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    print(f"  wrote {target.name} ({target.stat().st_size} bytes) and {answers_path.name}")


def control_answers(reply: str) -> Answers:
    """One line per control: its name, then the fields of ``_CONTROL_FIELDS``."""
    answers: Answers = {}
    for line in reply.splitlines():
        if not line.strip():
            continue
        parts = line.split("|")
        name = parts[0]
        fields = (parts[1:] + [""] * len(_CONTROL_FIELDS))[: len(_CONTROL_FIELDS)]
        entry: dict[str, object] = {}
        for key, raw in zip(_CONTROL_FIELDS, fields, strict=True):
            if key in ("type", "value"):
                entry[key] = int(raw) if raw.strip("-").isdigit() else raw
            else:
                entry[key] = raw
        answers[name] = entry
    return answers


def richtext_answers(reply: str) -> Answers:
    """One line per cell: its address, then each font change as
    ``start:name,size,bold,italic,colour,underline,superscript;``."""
    answers: Answers = {}
    for line in reply.splitlines():
        if not line.strip():
            continue
        address, _, body = line.partition("|")
        runs: list[dict[str, object]] = []
        for piece in body.split(";"):
            if not piece:
                continue
            start, _, fields = piece.partition(":")
            name, size, bold, italic, color, underline, superscript = fields.split(",")
            runs.append({
                "start": int(start),
                "name": name,
                "size": float(size),
                "bold": bold == "True",
                "italic": italic == "True",
                "color": int(color),
                "underline": int(underline),
                "superscript": superscript == "True",
            })
        answers[address] = {"runs": runs}
    return answers


def pivot_answers(reply: str) -> Answers:
    """For the file as built and for each edit Excel made: whether Excel
    refused it, and otherwise each pivot table's location and the source of
    the cache it reads, with how many caches the workbook kept. Read from
    the saved files, which are removed once read."""
    answers: Answers = {}
    for line in reply.splitlines():
        if not line.strip():
            continue
        parts = line.split("|")
        if parts[0] == "refused":
            answers[parts[1]] = {"refused": True}
            continue
        _, state, where = parts
        path = Path(where)
        answers[state] = {"refused": False, **_pivot_state(path)}
        path.unlink()
    return answers


def _pivot_state(path: Path) -> dict[str, object]:
    """Each pivot table's location and its cache's source, and how many
    caches there are, read with nothing but the archive and a few
    patterns."""
    tables: dict[str, dict[str, str]] = {}
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        for name in names:
            if not re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name):
                continue
            text = package.read(name).decode("utf-8")
            table = re.search(r'<pivotTableDefinition\b[^>]*\bname="([^"]*)"', text)
            location = re.search(r'<location\b[^>]*\bref="([^"]*)"', text)
            rels = package.read(f"xl/pivotTables/_rels/{name.rpartition('/')[2]}.rels").decode("utf-8")
            target = re.search(r'Target="([^"]*pivotCacheDefinition[^"]*)"', rels)
            source = ""
            if target is not None:
                cache = posixpath.normpath(posixpath.join("xl/pivotTables", target.group(1)))
                found = re.search(r"<worksheetSource\b[^>]*>", package.read(cache).decode("utf-8"))
                if found is not None:
                    sheet = re.search(r'\bsheet="([^"]*)"', found.group(0))
                    ref = re.search(r'\bref="([^"]*)"', found.group(0))
                    source = f"{html.unescape(sheet.group(1)) if sheet else ''}!{ref.group(1) if ref else ''}"
            if table is not None and location is not None:
                tables[html.unescape(table.group(1))] = {"location": location.group(1), "source": source}
        caches = sum(1 for name in names if re.fullmatch(r"xl/pivotCache/pivotCacheDefinition\d+\.xml", name))
    return {"tables": tables, "caches": caches}


def chart_answers(reply: str) -> Answers:
    """Each chart's series formulas and title before any edit, as Excel's
    object model gave them, and each chart's references, as Excel wrote
    them, in the file saved before any edit and in each saved after one.
    The files are in the temporary folder and are removed once read."""
    answers: Answers = {}
    for line in reply.splitlines():
        if not line.strip():
            continue
        parts = line.split("|")
        if parts[0] == "file":
            _, state, where = parts
            path = Path(where)
            for chart, references in _chart_references(path).items():
                answers[f"{state}|{chart}"] = {"references": references}
            path.unlink()
            continue
        state, chart, series, title = parts
        answers[f"reported {state}|{chart}"] = {
            "series": [formula for formula in series.split(";") if formula],
            "title": title,
        }
    return answers


def _chart_references(path: Path) -> dict[str, list[str]]:
    """Every chart in a workbook, as ``Sheet/Name`` or a chart sheet's
    name, and the references in its part in the order the part holds them.

    Read with nothing but the archive and a few patterns, following each
    sheet's drawing to the chart parts it names, so the answers owe nothing
    to the library they are there to check.
    """

    def text(package: zipfile.ZipFile, part: str) -> str:
        return package.read(part).decode("utf-8")

    def targets(package: zipfile.ZipFile, part: str) -> dict[str, str]:
        folder, _, name = part.rpartition("/")
        rels = f"{folder}/_rels/{name}.rels"
        if rels not in package.namelist():
            return {}
        found: dict[str, str] = {}
        for entry in re.findall(r"<Relationship\b[^>]*>", text(package, rels)):
            identifier = re.search(r'\bId="([^"]*)"', entry)
            target = re.search(r'\bTarget="([^"]*)"', entry)
            if identifier and target:
                found[identifier.group(1)] = posixpath.normpath(posixpath.join(folder, target.group(1)))
        return found

    charts: dict[str, list[str]] = {}
    with zipfile.ZipFile(path) as package:
        books = targets(package, "xl/workbook.xml")
        for entry in re.findall(r"<sheet\b[^>]*>", text(package, "xl/workbook.xml")):
            name = html.unescape(re.search(r'\bname="([^"]*)"', entry).group(1))  # type: ignore[union-attr]
            sheet_part = books[re.search(r'\br:id="([^"]*)"', entry).group(1)]  # type: ignore[union-attr]
            for drawing in targets(package, sheet_part).values():
                if "/drawings/" not in drawing or not drawing.endswith(".xml"):
                    continue
                drawn = targets(package, drawing)
                for frame in re.findall(r"<xdr:graphicFrame\b.*?</xdr:graphicFrame>", text(package, drawing), re.S):
                    chart = re.search(r'<c:chart\b[^>]*\br:id="([^"]*)"', frame)
                    frame_name = re.search(r'<xdr:cNvPr\b[^>]*\bname="([^"]*)"', frame)
                    if chart is None or frame_name is None:
                        continue
                    references = re.findall(r"<(?:\w+:)?f>(.*?)</(?:\w+:)?f>", text(package, drawn[chart.group(1)]))
                    key = name if "/chartsheets/" in sheet_part else f"{name}/{html.unescape(frame_name.group(1))}"
                    charts[key] = [html.unescape(reference) for reference in references]
    return charts


def style_answers(reply: str) -> Answers:
    """A line per built-in style, ``includes|name|`` and six flags, and a
    line per cell, its address and then the fields of ``_STYLE_FIELDS``,
    kept as the text Excel gave them."""
    answers: Answers = {}
    for line in reply.splitlines():
        if not line.strip():
            continue
        parts = line.split("|")
        if parts[0] == "includes":
            names = ("number_format", "font", "fill", "border", "alignment", "protection")
            flags = [flag == "True" for flag in parts[2].split(",")]
            answers[f"includes {parts[1]}"] = dict(zip(names, flags, strict=True))
            continue
        answers[parts[0]] = dict(zip(_STYLE_FIELDS, parts[1:], strict=True))
    return answers


def escape_answers(reply: str) -> Answers:
    """One line per piece of text: what it is, then its UTF-16 codes in
    hex. A surrogate pair makes one character, as it does in Excel."""
    answers: Answers = {}
    for line in reply.splitlines():
        if not line.strip():
            continue
        key, _, codes = line.partition("|")
        units = "".join(chr(int(code, 16)) for code in codes.split())
        text = units.encode("utf-16-le", "surrogatepass").decode("utf-16-le", "surrogatepass")
        answers[key] = {"text": text}
    return answers


def picture_answers(reply: str) -> Answers:
    """One line per picture: its name, then the fields of ``_PICTURE_FIELDS``."""
    answers: Answers = {}
    for line in reply.splitlines():
        if not line.strip():
            continue
        parts = line.split("|")
        fields = (parts[1:] + [""] * len(_PICTURE_FIELDS))[: len(_PICTURE_FIELDS)]
        entry: dict[str, object] = {}
        for key, raw in zip(_PICTURE_FIELDS, fields, strict=True):
            if key == "type":
                entry[key] = int(raw)
            elif key in ("left", "top", "width", "height"):
                entry[key] = float(raw)
            elif key == "locks_aspect":
                entry[key] = raw == "-1"
            else:
                entry[key] = raw
        answers[parts[0]] = entry
    return answers


def comment_answers(reply: str) -> Answers:
    """One line per note: ``Sheet!A1``, then the fields of ``_COMMENT_FIELDS``."""
    answers: Answers = {}
    for line in reply.splitlines():
        if not line.strip():
            continue
        parts = line.split("|")
        fields = (parts[1:] + [""] * len(_COMMENT_FIELDS))[: len(_COMMENT_FIELDS)]
        entry: dict[str, object] = {}
        for key, raw in zip(_COMMENT_FIELDS, fields, strict=True):
            if key == "text":
                entry[key] = raw.replace("\\n", "\n")
            elif key == "visible":
                entry[key] = raw == "True"
            elif key in ("left", "top", "width", "height"):
                entry[key] = float(raw)
            else:
                entry[key] = raw
        answers[parts[0]] = entry
    return answers


def filter_answers(reply: str) -> Answers:
    """One entry per sheet: the name, whether the autofilter is on, each
    live filter's criteria and operator, and how many rows Excel hid. The
    last is the measurement that matters, because Excel does not recompute
    a filter when the workbook opens."""
    answers: Answers = {}
    for row in reply.split("|"):
        if not row.strip():
            continue
        parts = row.split(";")
        name, mode = parts[0], parts[1] == "True"
        hidden = 0
        filters: dict[str, dict[str, str]] = {}
        for piece in parts[2:]:
            if piece.startswith("hidden="):
                hidden = int(piece.partition("=")[2] or 0)
                continue
            index, _, rest = piece.partition(":")
            first, second, operator = [*rest.split(":"), "", "", ""][:3]
            filters[index] = {
                "criteria1": first,
                "criteria2": second,
                "operator": operator,
            }
        answers[name] = {"on": mode, "hidden": hidden, "filters": filters}
    return answers


def main() -> int:
    try:
        from pyvbaharness import ExcelSession
        from pyvbaharness.session import HarnessConfig
    except ImportError:
        print(
            'pyvbaharness is not installed. Install the live extra:\n'
            '    python -m pip install -e ".[dev]" --group live',
            file=sys.stderr,
        )
        return 2
    # Seconds to wait for another session to finish with Excel, as the live
    # gate waits; by default the build does not wait.
    wait = float(os.environ.get("LIVE_EXCEL_LOCK_WAIT", "0"))

    force = "--force" in sys.argv
    wanted = [
        ("empty.xlsx", _BUILD_EMPTY),
        ("sample.xlsx", _BUILD_SAMPLE),
        ("structures.xlsx", _BUILD_STRUCTURES),
        ("refused.xlsx", _BUILD_REFUSED),
        ("settings.xlsx", _BUILD_SETTINGS),
        ("links.xlsx", _BUILD_LINKS),
        ("geometry.xlsx", _BUILD_GEOMETRY),
    ]
    #: Fixtures whose measurements are recorded beside them, each with what
    #: reads its reply back.
    measured = [
        ("controls.xlsm", _BUILD_CONTROLS, control_answers),
        ("filters.xlsx", _BUILD_FILTERS, filter_answers),
        ("comments.xlsx", _BUILD_COMMENTS, comment_answers),
        ("pictures.xlsx", _BUILD_PICTURES, picture_answers),
        ("richtext.xlsx", _BUILD_RICHTEXT, richtext_answers),
        ("escapes.xlsx", _BUILD_ESCAPES, escape_answers),
        ("styles.xlsx", _BUILD_STYLES, style_answers),
        ("charts.xlsx", _BUILD_CHARTS, chart_answers),
        ("pivots.xlsx", _BUILD_PIVOTS, pivot_answers),
    ]

    everything = [name for name, _ in wanted] + [name for name, _, _ in measured]
    everything += [f"{Path(name).stem}_answers.json" for name, _, _ in measured]
    if not force and all((FIXTURES / name).exists() for name in everything):
        print("every fixture is already there; nothing to do (pass --force to rebuild)")
        return 0

    FIXTURES.mkdir(parents=True, exist_ok=True)
    print(f"authoring fixtures in {FIXTURES}")
    with ExcelSession(HarnessConfig(lock_wait_s=wait)) as excel:
        for name, source in wanted:
            excel.new_workbook()
            build(excel, source, FIXTURES / name, name, force=force)
        for name, source, parse in measured:
            excel.new_workbook()
            build_answers(excel, source, FIXTURES / name, name, force=force, parse=parse)
    print("done")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
