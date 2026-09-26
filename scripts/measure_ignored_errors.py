"""Measure how Excel keeps and writes the errors a sheet ignores.

Excel's Ignore Error marks a cell so that one rule of error checking no
longer puts a green triangle on it, and a sheet keeps the marks in its
``<ignoredErrors>``. The library's reading and writing of them are held to
this, in two parts:

- Writes. Each scenario is a list of steps, a rule ignored or no longer
  ignored in a cell, taken on a sheet of its own through
  ``Range.Errors(i).Ignore``. Excel then reports which rules each cell
  ignores, and the workbook is saved, so the markup Excel wrote for the
  scenario is recorded beside what it held.
- Reads. Each variant is an ``<ignoredErrors>`` put into a copy of
  ``empty.xlsx`` with text in B2:D4. Excel opens it, reports which rules
  each cell ignores, and saves it again.

Run this on a Windows machine with Excel installed:

    python -m pip install -e ".[dev]" --group live
    python scripts/measure_ignored_errors.py

It writes ``tests/fixtures/excel/ignored_errors.json``, replacing what is
there, and nothing else.
"""

from __future__ import annotations

import datetime as dt
import json
import os
import re
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FIXTURE = ROOT / "tests" / "fixtures" / "excel" / "ignored_errors.json"
EMPTY = ROOT / "tests" / "fixtures" / "excel" / "empty.xlsx"

#: The rules in the order ``Range.Errors`` numbers them, as a sheet names
#: them.
RULES = ["evalError", "twoDigitTextYear", "numberStoredAsText", "formula", "formulaRange", "unlockedFormula",
         "emptyCellReference", "listDataValidation", "calculatedColumn", "misleadingFormat"]  # fmt: skip

#: Each step is "cell rule", with " 0" after it to stop ignoring the rule.
#: Rules are numbered as Range.Errors numbers them: 3 is a number stored as
#: text, 1 an error value.
WRITES = {
    "single": "B2 3",
    "down": "B2 3;B3 3",
    "across": "B2 3;C2 3",
    "gap": "B2 3;D2 3",
    "square": "B2 3;C2 3;B3 3;C3 3",
    "column first": "B2 3;B3 3;C2 3;C3 3",
    "ell down": "B2 3;B3 3;C2 3",
    "ell across": "B2 3;C2 3;C3 3",
    "ell other": "B2 3;C2 3;B3 3",
    "upward": "B3 3;B2 3",
    "leftward": "C2 3;B2 3",
    "diagonal": "B2 3;C3 3",
    "far": "B2 3;B100 3;Z2 3",
    "bridge": "B2 3;D2 3;C2 3",
    "extend last": "B2 3;D2 3;F2 3;C2 3;E2 3",
    "choose last": "B3 3;C2 3;C3 3",
    "choose last again": "C2 3;B3 3;C3 3",
    "stack": "B2 3;C2 3;D2 3;B3 3;C3 3;D3 3",
    "merge before": "B2 3;C2 3;B4 3;C4 3;B3 3;C3 3",
    "merge apart": "B2 3;C2 3;Z9 3;B3 3;C3 3",
    "ring": "B2 3;C2 3;D2 3;B3 3;D3 3;B4 3;C4 3;D4 3",
    "middle": "B2 3;B3 3;B4 3;C3 3",
    "two cells two rules": "B2 3;B3 1",
    "split for a rule": "B2 3;B3 3;B3 1",
    "rules apart": "B2 3;B3 3;B4 3;C2 1;C3 1;C4 1;D2 3;D2 1",
    "join an entry": "B2 3;B3 3;D2 1;D3 1;D2 3;B3 1",
    "drop a rule": "B2 3;B3 3;B4 3;B2 1;B3 1;B3 1 0",
    "order of entries": "D2 1;B2 3;C2 2",
    "misleading": "B2 10",
    "every rule alone": "B2 1;D2 2;F2 3;H2 4;J2 5;L2 6;N2 7;P2 8;R2 9;T2 10",
    "empty cell": "!F2 1",
    "undone": "B2 3;B2 3 0",
    "undone in a range": "B2 3;B3 3;B4 3;B3 3 0",
    "undone again": "B2 3;B3 3;B4 3;B3 3 0;B3 3",
    "undone then another": "B2 3;B3 3;B4 3;B3 3 0;B3 1",
    "undone in place": "B2 3;B3 3;B4 3;D2 3;B3 3 0",
    "undone at a corner": "B2 3;C2 3;B3 3;C3 3;B2 3 0",
    "undone in the middle": "B2 3;C2 3;D2 3;C2 3 0",
    "hole": "B2 3;C2 3;D2 3;B3 3;C3 3;D3 3;B4 3;C4 3;D4 3;C3 3 0",
    "hole in a column": "B2 3;B3 3;B4 3;C2 3;C3 3;C4 3;D2 3;D3 3;D4 3;C2 3 0",
    "undone alone": "B2 3;B3 3;B4 3;C3 3;C3 3 0",
    "never ignored": "B2 1 0",
    "another rule undone": "B2 3;B2 1 0",
    # Excel loses marks here: a cell alone in the first entry taking
    # another rule.
    "second rule": "B2 3;B2 2",
    "second rule both": "B2 3;B3 1;B2 1;B3 3",
    "second rule then undone": "B2 3;B2 1;B2 3 0",
    "masks move": "B2 3;B3 3;B4 3;B2 1;B3 1;B4 1",
}

#: Each variant is the content of an <ignoredErrors> a file holds.
READS = {
    "overlap": '<ignoredError sqref="B2:B3" numberStoredAsText="1"/><ignoredError sqref="B3" evalError="1"/>',
    "overlap reversed": '<ignoredError sqref="B3" evalError="1"/><ignoredError sqref="B2:B3" numberStoredAsText="1"/>',
    "scattered": '<ignoredError sqref="B2 B3 B4" numberStoredAsText="1"/>',
    "twins": '<ignoredError sqref="B2" numberStoredAsText="1"/><ignoredError sqref="B3" numberStoredAsText="1"/>',
    "no rule": '<ignoredError sqref="B2"/><ignoredError sqref="C2" numberStoredAsText="1"/>',
    "order": '<ignoredError sqref="C2" numberStoredAsText="1"/><ignoredError sqref="B2" evalError="1"/>',
    "repeat": '<ignoredError sqref="B2 B2" numberStoredAsText="1"/>',
    "unsorted": '<ignoredError sqref="D4 B2 C3" numberStoredAsText="1"/>',
    "whole sheet": '<ignoredError sqref="A1:XFD1048576" numberStoredAsText="1"/>',
    "side by side": '<ignoredError sqref="B2:B3 C2:C3" numberStoredAsText="1"/>',
    "overlapping ranges": '<ignoredError sqref="B2:C3 C3:D4" numberStoredAsText="1"/>',
    "attribute order": '<ignoredError numberStoredAsText="1" sqref="B2" evalError="1"/>',
    "true": '<ignoredError sqref="B2" numberStoredAsText="true"/>',
    "zero": '<ignoredError sqref="B2" numberStoredAsText="0" evalError="1"/>',
    "misleading": (
        '<ignoredError sqref="B2" xmlns:x16r3="http://schemas.microsoft.com/office/spreadsheetml/2018/08/main"'
        ' x16r3:misleadingFormat="1"/>'
    ),
}

_MEASURE = r'''
Public Function Writes(ByVal Spec As String, ByVal Target As String) As String
    Dim wb As Workbook, ws As Worksheet, scenarios() As String, parts() As String, steps() As String
    Dim i As Long, j As Long, bits() As String, out As String, cells As Collection, place As Variant
    Set wb = ActiveWorkbook
    scenarios = Split(Spec, "|")
    For i = 0 To UBound(scenarios)
        parts = Split(scenarios(i), "=")
        If i = 0 Then
            Set ws = wb.Worksheets(1)
        Else
            Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        End If
        ws.Name = "S" & i
        steps = Split(parts(1), ";")
        Set cells = New Collection
        For j = 0 To UBound(steps)
            bits = Split(steps(j), " ")
            If Left$(bits(0), 1) = "!" Then
                bits(0) = Mid$(bits(0), 2)
            ElseIf IsEmpty(ws.Range(bits(0)).Value) Then
                ws.Range(bits(0)).Formula = "'5"
            End If
            On Error Resume Next
            cells.Add bits(0), bits(0)
            On Error GoTo 0
        Next j
        For j = 0 To UBound(steps)
            bits = Split(steps(j), " ")
            If Left$(bits(0), 1) = "!" Then bits(0) = Mid$(bits(0), 2)
            ws.Range(bits(0)).Errors(CLng(bits(1))).Ignore = (UBound(bits) < 2)
        Next j
        out = out & parts(0) & "=" & State(ws, cells) & vbLf
    Next i
    Application.DisplayAlerts = False
    wb.SaveAs Target, 51
    Application.DisplayAlerts = True
    Writes = out
End Function

Public Function Reads(ByVal Folder As String, ByVal Names As String) As String
    Dim wb As Workbook, names_() As String, i As Long, r As Long, c As Long, out As String, cells As Collection
    names_ = Split(Names, "|")
    For i = 0 To UBound(names_)
        Set wb = Workbooks.Open(Folder & "\" & names_(i) & ".xlsx")
        Set cells = New Collection
        For r = 2 To 4
            For c = 2 To 4
                cells.Add wb.Worksheets(1).Cells(r, c).Address(False, False)
            Next c
        Next r
        out = out & names_(i) & "=" & State(wb.Worksheets(1), cells) & vbLf
        Application.DisplayAlerts = False
        wb.SaveAs Folder & "\" & names_(i) & " saved.xlsx", 51
        wb.Close SaveChanges:=False
        Application.DisplayAlerts = True
    Next i
    Reads = out
End Function

Private Function State(ws As Worksheet, cells As Collection) As String
    ' Each cell and the rules it ignores, by their numbers.
    Dim place As Variant, k As Long, mask As String, out As String
    For Each place In cells
        mask = ""
        For k = 1 To 10
            If ws.Range(place).Errors(k).Ignore Then mask = mask & IIf(mask = "", "", ",") & k
        Next k
        out = out & IIf(out = "", "", ";") & place & ":" & mask
    Next place
    State = out
End Function
'''


def _state(text: str) -> dict[str, list[str]]:
    """Each cell and the rules Excel said it ignores, from the macro's
    report; cells ignoring nothing left out."""
    found: dict[str, list[str]] = {}
    for item in text.split(";"):
        place, _, numbers = item.partition(":")
        if numbers:
            found[place] = [RULES[int(number) - 1] for number in numbers.split(",")]
    return found


def _steps(spec: str) -> list[list[object]]:
    found: list[list[object]] = []
    for step in spec.split(";"):
        bits = step.split(" ")
        found.append([bits[0].lstrip("!"), RULES[int(bits[1]) - 1], len(bits) < 3])
    return found


def _ignored_errors(package: Path, sheet: str) -> str | None:
    """The ``<ignoredErrors>`` Excel saved for the sheet named ``sheet``."""
    with zipfile.ZipFile(package) as source:
        workbook = source.read("xl/workbook.xml").decode("utf-8")
        relations = source.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        found = re.search(rf'<sheet name="{re.escape(sheet)}" sheetId="\d+" r:id="(rId\d+)"', workbook)
        assert found, sheet
        target = re.search(rf'<Relationship Id="{found.group(1)}"[^>]*Target="([^"]+)"', relations)
        assert target, found.group(1)
        xml = source.read("xl/" + target.group(1)).decode("utf-8")
    found = re.search(r"<ignoredErrors>.*?</ignoredErrors>", xml)
    return found.group(0) if found else None


def _build(variant: str, target: Path) -> None:
    """A copy of empty.xlsx with text in B2:D4 and ``variant`` for its
    ignored errors."""
    cells = "".join(
        f'<row r="{row}">' + "".join(f'<c r="{column}{row}" t="inlineStr"><is><t>5</t></is></c>' for column in "BCD")
        + "</row>"
        for row in (2, 3, 4)
    )  # fmt: skip
    with zipfile.ZipFile(EMPTY) as source, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                text = data.decode("utf-8").replace('<dimension ref="A1"/>', '<dimension ref="B2:D4"/>')
                text = text.replace("<sheetData/>", f"<sheetData>{cells}</sheetData>")
                text = text.replace("</worksheet>", f"<ignoredErrors>{variant}</ignoredErrors></worksheet>")
                data = text.encode("utf-8")
            out.writestr(item, data)


def main() -> int:
    try:
        from pyvbaharness import ExcelSession
        from pyvbaharness.session import HarnessConfig
    except ImportError:
        print('pyvbaharness is not installed. Install the live extra:\n    python -m pip install -e ".[dev]" --group live',
              file=sys.stderr)
        return 2
    wait = float(os.environ.get("LIVE_EXCEL_LOCK_WAIT", "0"))
    with tempfile.TemporaryDirectory() as scratch:
        folder = Path(scratch)
        for name, variant in READS.items():
            _build(variant, folder / f"{name}.xlsx")
        written = folder / "writes.xlsx"
        spec = "|".join(f"{name}={steps}" for name, steps in WRITES.items())
        with ExcelSession(HarnessConfig(lock_wait_s=wait)) as excel:
            excel.new_workbook()
            today = dt.date.today()
            writes = excel.run_vba(_MEASURE, proc="Writes", args=(spec, str(written)), timeout=600)
            reads = excel.run_vba(_MEASURE, proc="Reads", args=(str(folder), "|".join(READS)), timeout=600)
            for result in (writes, reads):
                if result.outcome != "passed":
                    print(f"Excel refused the measurement: {result!r}", file=sys.stderr)
                    return 1
        write_states = dict(line.split("=", 1) for line in str(writes.value).strip("\n").split("\n"))
        read_states = dict(line.split("=", 1) for line in str(reads.value).strip("\n").split("\n"))
        record: dict[str, object] = {
            "measured": today.isoformat(),
            "writes": [
                {
                    "name": name,
                    "steps": _steps(steps),
                    "ignored": _state(write_states[name]),
                    "saved": _ignored_errors(written, f"S{index}"),
                }
                for index, (name, steps) in enumerate(WRITES.items())
            ],
            "reads": [
                {
                    "name": name,
                    "file": f"<ignoredErrors>{variant}</ignoredErrors>",
                    "ignored": _state(read_states[name]),
                    "saved": _ignored_errors(folder / f"{name} saved.xlsx", "Sheet1"),
                }
                for name, variant in READS.items()
            ],
        }
    FIXTURE.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    print(f"wrote {FIXTURE.relative_to(ROOT)} ({len(WRITES)} scenarios, {len(READS)} files)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
