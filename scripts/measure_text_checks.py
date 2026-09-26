"""Measure which text Excel's error checking calls a number stored as text
or a date with a two-digit year.

The two rules of error checking that read a cell's text are held to this:
each string below is typed into a cell after an apostrophe, and
``Range.Errors`` is read for both rules. The strings are the forms the
rules were worked out from, a grid of numbers and month names joined every
way, the separators, digit counts and spellings that decide a borderline
case, and a few thousand random strings made from the same pieces with a
fixed seed, which the rules were checked against after they were found.

Run this on a Windows machine with Excel installed:

    python -m pip install -e ".[dev]" --group live
    python scripts/measure_text_checks.py

It writes ``tests/fixtures/excel/text_checks.json``, replacing what is
there, and nothing else. The locale is part of the result, and so is the
date: whether ``2/29`` names a day depends on the year it is measured in,
which the file records. The committed corpus was measured in en-US.
"""

from __future__ import annotations

import datetime as dt
import itertools
import json
import os
import random
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FIXTURE = ROOT / "tests" / "fixtures" / "excel" / "text_checks.json"

TOKENS = ["1", "13", "32", "99", "0", "2020", "100", "Jan"]
SEPARATORS = ["/", "-", " ", "  ", ",", ", ", ",  ", " ,", " , ", " /", "/ ", " / ", "  /", " -", "- ", ""]
MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october",
          "november", "december"]  # fmt: skip
NUMBERS = [
    "123", "-123", "+123", "(123)", "123-", "1,000", "1,0000", "12,34,567", "1,2", ",123", "0,123", "0010,000",
    "1,000.5", "1,00.5", "(1,000)", "$5", "-$5", "($5)", "5$", "$5%", "%$5", "5%", "5 %", "(5%)", "%5", "5%%",
    "1e5", "1E+308", "1E+309", "1e", "1.5e3", "1 1/2", "1 1/0", "1/0", "0 1/2", "1  1/2", "1,000 1/2", "1.5 1/2",
    "1 32767/1", "1 32768/1", "0123", ".5", "5.", "-.5", ".", "TRUE", "abc", " 123", "123 ", " ", "1:30",
    "25:00", "1:30 PM", "1/2", "Jan 1", "1/2020", "2020-01-01", "1.1.2020",
]  # fmt: skip


def corpus() -> list[str]:
    """Every string measured, each once."""
    found: list[str] = []
    found += [f"{a}/{b}" for a, b in itertools.product([*TOKENS, "x"], repeat=2)]
    found += [f"{a}/{b}/{c}" for a, b, c in itertools.product(TOKENS, repeat=3)]
    for first, second, third in [("1", "1", "99"), ("Jan", "1", "99"), ("1", "Jan", "99"), ("99", "Jan", "1")]:
        found += [f"{first}{one}{second}{two}{third}" for one, two in itertools.product(SEPARATORS, repeat=2)]
    for day in ["1", "01", "001", "0001", "257", "287", "288", "1025", "9999", "65537", "31", "32"]:
        found += [f"1/{day}/99", f"2/{day}/99", f"Jan {day} 99", f"{day} Jan 99", f"99 Jan {day}"]
    for year in ["99", "099", "0099", "9", "09", "009", "0", "00", "1999", "2099"]:
        found += [f"1/1/{year}", f"1/{year}", f"Jan {year}", f"{year} Jan", f"1 Jan {year}", f"{year} Jan 1"]
    for month in ["1", "01", "001", "12", "13"]:
        found += [f"{month}/1/99", f"{month}/99"]
    for name in ["Jan", "January", "jan", "JAN", "Janu", "Sept", "Sep", "Febr", "Ma", "Mar", "Ju", "Juli", "Mai",
                 "Dece", "Jan.", "Mon", "Monday"]:
        found += [f"{name} 99", f"1 {name} 99", f"{name} 1, 99", f"99 {name}", f"{name} 1"]
    for first, second in itertools.product(["Jan", "Feb", "Mar", "Dec"], repeat=2):
        found += [f"{first} {second} 99", f"99 {first} {second}", f"{first} 1 {second}"]
    for base in ["1/1/99", "Jan 99", "99 Jan", "1 1/0"]:
        found += [" " + base, base + " ", base + " 12:00", "12:00 " + base, base + " 1 PM", base + " 12"]
    found += ["2/29", "2/30", "4/31", "Feb 29", "29 Feb", "2/29/00", "2/29/01", "1/1/99/1", "1 1 1 99", "Jan 1st 99"]
    found += NUMBERS
    found += _random_strings(random.Random(20260925), 5000)
    return list(dict.fromkeys(text for text in found if text.strip()))


def _random_strings(rng: random.Random, count: int) -> list[str]:
    def number() -> str:
        roll = rng.random()
        if roll < 0.35:
            text = str(rng.randint(0, 12))
        elif roll < 0.6:
            text = str(rng.randint(13, 31))
        elif roll < 0.8:
            text = str(rng.randint(32, 99))
        elif roll < 0.9:
            text = str(rng.choice([rng.randint(100, 999), rng.randint(1000, 9999), rng.randint(1900, 2100)]))
        else:
            text = str(rng.randint(0, 99999))
        return "0" * rng.randint(1, 2) + text if rng.random() < 0.12 else text

    def word() -> str:
        if rng.random() < 0.15:
            return rng.choice(["mon", "monday", "x", "am", "pm", "st", "of"])
        month = rng.choice(MONTHS)
        text = month[: rng.randint(2, len(month))]
        style = rng.random()
        return text.upper() if style < 0.2 else text.capitalize() if style < 0.8 else text

    joins = SEPARATORS[:6] * 3 + SEPARATORS
    made: list[str] = []
    for _ in range(count):
        parts = [word() if rng.random() < 0.3 else number() for _ in range(rng.choice([1, 2, 2, 3, 3, 3, 3, 4]))]
        text = parts[0] + "".join(rng.choice(joins) + part for part in parts[1:])
        if rng.random() < 0.08:
            text = " " * rng.randint(1, 2) + text
        if rng.random() < 0.05:
            text += " "
        if rng.random() < 0.04:
            text += rng.choice([" 12:00", " 1 PM", "12:00"])
        made.append(text)
    return made


_MEASURE = r'''
Public Function Measure(ByVal Source As String) As String
    Dim whole As String, lines() As String, f As Integer, n As Long, r As Long
    Dim cells() As Variant, ws As Worksheet, out As String
    f = FreeFile
    Open Source For Binary Access Read As #f
    whole = Space$(LOF(f))
    Get #f, , whole
    Close #f
    lines = Split(whole, vbCrLf)
    n = UBound(lines)
    ReDim cells(1 To n, 1 To 1)
    For r = 1 To n
        cells(r, 1) = "'" & lines(r - 1)
    Next r
    Set ws = ActiveWorkbook.Worksheets(1)
    ws.Range(ws.Cells(1, 1), ws.Cells(n, 1)).Formula = cells
    For r = 1 To n
        out = out & IIf(ws.Cells(r, 1).Errors(3).Value, "1", "0") & IIf(ws.Cells(r, 1).Errors(2).Value, "1", "0")
    Next r
    Measure = out
End Function
'''


def main() -> int:
    try:
        from pyvbaharness import ExcelSession
        from pyvbaharness.session import HarnessConfig
    except ImportError:
        print('pyvbaharness is not installed. Install the live extra:\n    python -m pip install -e ".[dev]" --group live',
              file=sys.stderr)
        return 2
    texts = corpus()
    wait = float(os.environ.get("LIVE_EXCEL_LOCK_WAIT", "0"))
    with tempfile.TemporaryDirectory() as scratch:
        source = Path(scratch) / "texts.txt"
        # Bytes, so no line ending is translated on the way.
        source.write_bytes(("\r\n".join(texts) + "\r\n").encode("ascii"))
        with ExcelSession(HarnessConfig(lock_wait_s=wait)) as excel:
            excel.new_workbook()
            today = dt.date.today()
            result = excel.run_vba(_MEASURE, proc="Measure", args=(str(source),), timeout=900)
            if result.outcome != "passed":
                print(f"Excel refused the measurement: {result!r}", file=sys.stderr)
                return 1
    flags = str(result.value or "")
    if len(flags) != 2 * len(texts):
        print(f"expected {2 * len(texts)} flags, Excel answered {len(flags)}", file=sys.stderr)
        return 1
    cases = [[text, flags[2 * index] == "1", flags[2 * index + 1] == "1"] for index, text in enumerate(texts)]
    record = {"measured": today.isoformat(), "fields": ["text", "numberStoredAsText", "twoDigitTextYear"]}
    lines = [json.dumps(record)[:-1] + ', "cases": [']
    lines += [f"  {json.dumps(case)}," for case in cases[:-1]] + [f"  {json.dumps(cases[-1])}", "]}"]
    FIXTURE.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"wrote {FIXTURE.relative_to(ROOT)} ({len(cases)} strings)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
