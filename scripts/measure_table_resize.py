"""Measure what Excel does to references and stored ranges when a table
gives up the rows at its bottom through ``ListObject.Resize``.

Remove Duplicates on a table resizes it this way once it has moved the rows
it keeps up, so these are the rules its references follow. Every case is a
sheet of its own holding a table over ``A1:C8`` whose data rows are 2 to 8,
or over ``A1:C9`` with a totals row on row 9, resized to ``A1:C5``: the
table keeps data rows 2 to 5, and a totals row moves up to row 6, rows 6
to 8 moving down a row to make way.

What is recorded, from the file saved before the resizes and the file
saved after:

- formulas beside the table reading it every way that tells the rules
  apart, with a totals row and without;
- every span of rows in column B, from row 1 to row 11, as a conditional
  format's range and as a validation's, with a totals row;
- conditional formats and validations of every kind of rule on the totals
  row, on one cell of it, and on it and the rows below, since whether a
  rule reads a cell relative to where it applies decides where its range
  goes.

Run this on a Windows machine with Excel installed:

    python -m pip install -e ".[dev]" --group live
    python scripts/measure_table_resize.py

It writes ``tests/fixtures/excel/table_resizes.json``, replacing what is
there, and nothing else.
"""

from __future__ import annotations

import datetime as dt
import html
import json
import os
import re
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FIXTURE = ROOT / "tests" / "fixtures" / "excel" / "table_resizes.json"

#: References beside the table, read with a totals row and without.
FORMULAS = [
    "=B2", "=B3", "=B4", "=B5", "=B6", "=B7", "=B8", "=B9", "=B10", "=C8", "=A6", "=D8", "=D9", "=C9",
    "=SUM(B3:B8)", "=SUM(B5:B8)", "=SUM(A2:B8)", "=SUM(B8:B10)", "=SUM(B2:B7)", "=SUM(A2:E8)", "=SUM(B6:B8)",
    "=SUM(B2:B8)", "=SUM(B7:B8)", "=SUM(B8:B9)", "=SUM(B9:B10)", "=SUM(B9:B12)", "=SUM(B7:B9)", "=SUM(B5:B9)",
    "=SUM(B2:B9)", "=SUM(A1:C9)", "=SUM(B6:B9)", "=SUM(B3:B6)", "=SUM(A1:C8)", "=SUM(B4:B10)", "=SUM(C2:C8)",
    "=SUM(B1:B8)", "=SUM(A8:C8)", "=SUM(B$2:B$8)", "=SUM(A8:E8)", "=SUM(A6:E6)", "=SUM(A9:E9)", "=SUM(A6:E9)",
    "=SUM(A5:E9)", "=SUM(A7:D7)", "=SUM(B8:D8)", "=SUM(C6:D9)", "=SUM(A6:D8)", "=SUM(A10:E10)", "=SUM(A1:E9)",
    "=SUM(A1:E12)", "=SUM(B6:E6)", "=SUM(C9:D9)", "=SUM(A8:D9)", "=SUM(A7:E9)", "=SUM(A5:E8)", "=SUM(A7:C9)",
    "=SUM(A6:C9)", "=SUM(A5:C9)", "=SUM(A9:C12)", "=SUM(A2:D8)", "=SUM(C2:D8)", "=SUM(A6:A8)", "=SUM(A1:D8)",
    "=SUM(A9:C9)", "=SUM(B9:B9)", "=SUM(A8:C9)", "=SUM(A2:C9)", "=SUM(A9:C10)", "=SUM(A5:C6)", "=SUM(B8:B8)",
    "=SUM(3:4)", "=SUM(B:B)", "=$B$9", "=SUM($B$2:$B$8)",
]  # fmt: skip

#: Conditional formats of each kind of rule, as VBA adds one to a range.
FORMATS = [
    'FormatConditions.Add Type:=1, Operator:=5, Formula1:="100"',
    'FormatConditions.Add Type:=1, Operator:=5, Formula1:="=$B9"',
    'FormatConditions.Add Type:=1, Operator:=5, Formula1:="=$B$9"',
    'FormatConditions.Add Type:=2, Formula1:="=$B9>0"',
    'FormatConditions.Add Type:=2, Formula1:="=$B$9>0"',
    'FormatConditions.Add Type:=2, Formula1:="=TRUE"',
    "FormatConditions.AddColorScale ColorScaleType:=2",
    "FormatConditions.AddTop10",
]
#: Validations of each kind, the same way.
VALIDATIONS = [
    'Validation.Add Type:=1, AlertStyle:=1, Operator:=1, Formula1:="0", Formula2:="9"',
    'Validation.Add Type:=7, AlertStyle:=1, Formula1:="=$B9>0"',
    'Validation.Add Type:=7, AlertStyle:=1, Formula1:="=$B$9>0"',
]
#: Where each kind of rule goes: the totals row, one cell of it, and it
#: and the rows below.
KIND_AREAS = ["A9:C9", "B9", "A9:C11"]
#: Every span of rows in column B, from row 1 to row 11.
SPANS = [f"B{top}:B{bottom}" for top in range(1, 12) for bottom in range(top, 12)]

_HELPERS = r'''
Private Function Table_(wb As Workbook, ByVal Totals As Boolean) As ListObject
    Dim ws As Worksheet, lo As ListObject, i As Long
    If wb.Worksheets.Count = 1 And wb.Worksheets(1).UsedRange.Address = "$A$1" And IsEmpty(wb.Worksheets(1).Range("A1")) Then
        Set ws = wb.Worksheets(1)
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    End If
    ws.Name = "Case" & wb.Worksheets.Count
    ws.Range("A1:C1").Value = Array("k", "n", "m")
    For i = 2 To 8
        ws.Cells(i, 1).Value = "k" & i
        ws.Cells(i, 2).Value = i
        ws.Cells(i, 3).Value = "m" & i
        ws.Cells(i, 5).Value = "o" & i
    Next i
    Set lo = ws.ListObjects.Add(1, ws.Range("A1:C8"), , 1)
    lo.Name = "T" & ws.Name
    If Totals Then lo.ShowTotals = True
    Set Table_ = lo
End Function

Private Sub Formulas(ws As Worksheet, ByVal Start As Long, ByVal Texts As String)
    Dim parts() As String, i As Long
    parts = Split(Texts, "@")
    For i = 0 To UBound(parts)
        ws.Cells(Start + i, 8).Formula = parts(i)
    Next i
End Sub

Private Sub Resize_(wb As Workbook)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        ws.ListObjects(1).Resize ws.Range("A1:C5")
    Next ws
End Sub
'''


def _cases() -> list[tuple[str, bool, str]]:
    """Each case: what it sets up, whether its table has a totals row, and
    the VBA that sets it up on ``ws``. The formulas go twenty to a line,
    since VBA caps how long a line may be."""
    texts = "\n    ".join(
        f'Formulas ws, {start + 1}, "{"@".join(FORMULAS[start : start + 20])}"' for start in range(0, len(FORMULAS), 20)
    )
    cases = [("formulas", totals, texts) for totals in (False, True)]
    cases += [("format", True, f'ws.Range("{span}").{FORMATS[0]}') for span in SPANS]
    cases += [("validation", True, f'ws.Range("{span}").{VALIDATIONS[0]}') for span in SPANS]
    cases += [("format", True, f'ws.Range("{area}").{rule}') for rule in FORMATS for area in KIND_AREAS]
    cases += [("validation", True, f'ws.Range("{area}").{rule}') for rule in VALIDATIONS for area in KIND_AREAS]
    return cases


def _module(cases: list[tuple[str, bool, str]]) -> str:
    """The measuring module: the cases in routines of forty, since VBA caps
    how long one routine may be, then the resizes, each workbook saved."""
    chunks = [cases[start : start + 40] for start in range(0, len(cases), 40)]
    routines = [
        f"Private Sub Build{index}(wb As Workbook)\n    Dim ws As Worksheet\n"
        + "".join(
            f"    Set ws = Table_(wb, {'True' if totals else 'False'}).Parent\n    {line}\n" for _, totals, line in chunk
        )
        + "End Sub\n"
        for index, chunk in enumerate(chunks)
    ]
    calls = "".join(f"    Build{index} wb\n" for index in range(len(chunks)))
    return (
        _HELPERS
        + "\n".join(routines)
        + "\nPublic Function Measure(ByVal Before As String, ByVal After As String) As String\n"
        + "    Dim wb As Workbook\n    Set wb = Workbooks.Add\n"
        + calls
        + "    Application.DisplayAlerts = False\n    wb.RemovePersonalInformation = True\n"
        + "    wb.SaveAs Filename:=Before, FileFormat:=51\n    Resize_ wb\n"
        + "    wb.SaveAs Filename:=After, FileFormat:=51\n    wb.Close SaveChanges:=False\n"
        + '    Application.DisplayAlerts = True\n    Measure = "ok"\nEnd Function\n'
    )


def _sheets(path: Path) -> dict[str, str]:
    """Each sheet's XML by name."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
        targets.update({rid: target for target, rid in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
        found: dict[str, str] = {}
        for name, rid in re.findall(r'<sheet name="([^"]+)" sheetId="\d+" r:id="([^"]+)"', workbook):
            found[name] = package.read("xl/" + targets[rid].lstrip("/").removeprefix("xl/")).decode("utf-8")
    return found


def _formulas(xml: str) -> list[str]:
    return [html.unescape(text) for _, text in re.findall(r'<c r="H(\d+)"[^>]*><f>([^<]*)</f>', xml)]


def _formats(xml: str) -> list[tuple[str, str, list[str]]]:
    """Each conditional format's range, its rule's type and formulas."""
    return [
        (sqref, kind, [html.unescape(text) for text in re.findall(r"<formula>([^<]*)</formula>", body)])
        for sqref, kind, body in re.findall(
            r'<conditionalFormatting sqref="([^"]*)"><cfRule type="(\w+)"(.*?)</conditionalFormatting>', xml
        )
    ]


def _validations(xml: str) -> list[tuple[str, list[str]]]:
    """Each validation's range and formulas."""
    return [
        (sqref, [html.unescape(text) for text in re.findall(r"<formula\d>([^<]*)</formula\d>", body)])
        for sqref, body in re.findall(r'<dataValidation [^>]*sqref="([^"]*)"[^>]*>(.*?)</dataValidation>', xml)
    ]


def main() -> int:
    try:
        from pyvbaharness import ExcelSession
        from pyvbaharness.session import HarnessConfig
    except ImportError:
        print('pyvbaharness is not installed. Install the live extra:\n    python -m pip install -e ".[dev]" --group live',
              file=sys.stderr)
        return 2
    cases = _cases()
    wait = float(os.environ.get("LIVE_EXCEL_LOCK_WAIT", "0"))
    with tempfile.TemporaryDirectory() as scratch:
        before_path, after_path = Path(scratch) / "before.xlsx", Path(scratch) / "after.xlsx"
        with ExcelSession(HarnessConfig(lock_wait_s=wait)) as excel:
            excel.new_workbook()
            today = dt.date.today()
            result = excel.run_vba(_module(cases), proc="Measure", args=(str(before_path), str(after_path)), timeout=900)
            if result.outcome != "passed":
                print(f"Excel refused the measurement: {result!r}", file=sys.stderr)
                return 1
        before, after = _sheets(before_path), _sheets(after_path)
    formulas: list[dict[str, object]] = []
    formats: list[dict[str, object]] = []
    validations: list[dict[str, object]] = []
    for index, (kind, totals, _) in enumerate(cases, start=1):
        name = f"Case{index}"
        if kind == "formulas":
            pairs = zip(_formulas(before[name]), _formulas(after[name]), strict=True)
            formulas += [{"totals": totals, "before": old, "after": new} for old, new in pairs]
        elif kind == "format":
            (old, rule, texts), (new, _, _) = _formats(before[name])[0], _formats(after[name])[0]
            formats.append({"type": rule, "formulas": texts, "before": old, "after": new})
        else:
            (old, texts), (new, _) = _validations(before[name])[0], _validations(after[name])[0]
            validations.append({"formulas": texts, "before": old, "after": new})
    record = {
        "measured": today.isoformat(),
        "table": {"columns": "A:C", "first": 2, "last": 8, "kept": 5, "totals": 9},
        "formulas": formulas,
        "formats": formats,
        "validations": validations,
    }
    FIXTURE.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    print(f"wrote {FIXTURE.relative_to(ROOT)} ({len(cases)} cases)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
