"""Measure how Excel reads text as a date or a time.

VALUE, DATEVALUE and TIMEVALUE, and text met in arithmetic, read a date
or a time from text. They are held to this: each string below is put in a
cell as text, and the three functions of it are read back, each as the
exact double Excel calculated or as an error. The strings are grids of the
forms the rules were worked out from: times with hours, minutes and
seconds of every width, with AM or PM; numbers and month names joined by
every separator, with every width of day, month and year; a date and a
time in either order, and what may trail them; and random strings made
with fixed seeds, the last few thousand of them measured only after the
rules were found. A few are measured again in a workbook using the 1904
date system.

Run this on a Windows machine with Excel installed:

    python -m pip install -e ".[dev]" --group live
    python scripts/measure_date_texts.py

It writes ``tests/fixtures/excel/date_texts.json``, replacing what is
there, and nothing else: each string with VALUE of it, ``null`` for an
error, and whether DATEVALUE and TIMEVALUE read it. They read it as VALUE's
whole days and the rest, which is checked of every string before the file
is written. The locale is part of the result, and so is the date: a date
without a year is in the year it is measured in, which the file records.
The committed corpus was measured in en-US.
"""

from __future__ import annotations

import datetime as dt
import itertools
import json
import math
import os
import random
import struct
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FIXTURE = ROOT / "tests" / "fixtures" / "excel" / "date_texts.json"

MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october",
          "november", "december"]  # fmt: skip


def _times() -> list[str]:
    hours = ["0", "00", "1", "01", "9", "12", "23", "24", "25", "29", "99", "100", "0100", "999", "1000", "9999",
             "10000", "32767", "32768", "65536", "99999"]  # fmt: skip
    minutes = ["0", "00", "5", "05", "59", "60", "61", "99", "100", "003", "0005", "999", "1000", "9999", "10000",
               "99999"]  # fmt: skip
    found = [f"{h}:{m}" for h, m in itertools.product(hours, minutes)]
    hours = ["0", "1", "12", "23", "24", "25", "99", "1000", "9999", "10000"]
    minutes = ["0", "59", "60", "99", "100", "003", "9999", "10000"]
    seconds = ["0", "00", "5", "59", "60", "99", "100", "018", "999", "5.5", "59.999", "60.5", "99.9", "0.5", ".5",
               "5.", "100.5", "1.23456789"]  # fmt: skip
    found += [f"{h}:{m}:{s}" for h, m, s in itertools.product(hours, minutes, seconds)]
    minutes = ["0", "1", "59", "60", "93", "100", "1000", "9999", "10000"]
    seconds = ["0.5", "22.15", "59.9", "60.5", "99.99", "100.5", "15215.8", "5.", "0.0", "05.5", "005.5"]
    found += [f"{m}:{s}" for m, s in itertools.product(minutes, seconds)]
    times = ["0", "1", "11", "12", "13", "23", "24", "0:30", "1:30", "12:30", "13:30", "24:00", "1:30:45",
             "12:00:00.5", "100", "1:60", "1:99", "0:00"]  # fmt: skip
    markers = [" AM", " PM", "AM", "PM", " am", " pm", " A", " P", " a", " p", " a.m.", " p.m.", "  PM", " A M",
               " AMX", " Am"]  # fmt: skip
    found += [time + marker for time, marker in itertools.product(times, markers)]
    found += ["12 :30", "12: 30", "12 : 30", "12:30 ", "  12:30", "12:30:", "12::30", ":30", "12:", "1:2:3:4",
              "12:30:45:6", "12:30.5.5", "12.30", "12h30", "12:30:45.", "12:30:.5", "12:3a", "12:-30", "-12:30",
              "+12:30", "(12:30)", "12:30-", "12,30", "1:2.5:3", "AM", "PM 12:30", "AM 1", "12:30 PM PM",
              "12:30AM PM", "1:00 PM 2:00"]  # fmt: skip
    words = ["Jan", "Dec", "Sept", "Ja", "Mon"]
    found += [f"{w}:{n}" for w, n in itertools.product(words, ["5", "59", "60", "003", "05", "0"])]
    found += [f"{n}:{w}" for n, w in itertools.product(["0", "5", "23", "24", "25"], ["Jan", "Dec"])]
    found += ["May:1:2", "5:Jan:5", "Jan:Jan", "1:2:Jan", "Jan:5 PM", "5:Jan PM", "Jan:5:5.5", "1:Jan:59"]
    return found


def _numeric_dates() -> list[str]:
    firsts = ["0", "1", "01", "001", "2", "9", "12", "13", "2020"]
    seconds = ["0", "1", "01", "28", "29", "30", "31", "32", "99", "100", "0100", "999", "1000", "1899", "1900",
               "2020", "9999", "10000", "029", "0029", "00"]  # fmt: skip
    found = [f"{a}{s}{b}" for a, s, b in itertools.product(firsts, ["/", "-", " / ", " ", ".", ","], seconds)]
    months = ["1", "01", "001", "2", "12", "13", "2020", "0"]
    days = ["1", "01", "29", "30", "31", "0", "001", "32"]
    years = ["0", "00", "1", "29", "30", "99", "100", "0099", "999", "1899", "1900", "2020", "2021", "2024", "9999",
             "10000"]  # fmt: skip
    found += [f"{m}/{d}/{y}" for m, d, y in itertools.product(months, days, years)]
    separators = ["/", "-", " / ", " /", "/ ", " - ", " ", ".", ",", ", ", ""]
    for m, d, y in [("1", "15", "2020"), ("2", "29", "2020"), ("2", "29", "2021"), ("12", "31", "99"),
                    ("2020", "1", "15")]:  # fmt: skip
        found += [f"{m}{one}{d}{two}{y}" for one, two in itertools.product(separators, repeat=2)]
    found += ["2/29/1900", "2/29/2000", "2/29/2100", "2/29/2024", "2/29/2023", "2/29/0", "2/29/4", "2/29/1",
              "2/29/00", "2/29/04", "29-Feb-1900", "Feb 29, 2100", "29 Feb 2024", "2020-2-29", "2021-2-29",
              "1900-2-29", "1900-3-1", "1900-1-1", "1899-12-31", "1900-1-0", "12/31/9999", "1/1/10000",
              "12/31/9999 23:59:59", "12/31/9999 24:00", "12/31/9999 48:00", "1/0/1900", "1/1/1900 0:00"]  # fmt: skip
    for year, month, day, separator in itertools.product(
        ["2020", "1900", "1899", "9999", "0099", "10000", "99", "0"],
        ["1", "01", "13", "Jan", "January"],
        ["1", "15", "29", "31", "32"],
        ["-", "/", " "],
    ):
        found.append(f"{year}{separator}{month}{separator}{day}")
    found += [f"2020{s}{m}" for s, m in itertools.product(["-", "/", " "], ["1", "01", "13", "Jan"])]
    return found


def _named_dates() -> list[str]:
    words = ["Jan", "February", "Sept", "Ma", "Mon", "Decembe", "jul"]
    separators = ["", " ", "  ", "/", "-", " / ", " - ", ", ", ",", ".", ". ", " , "]
    numbers = ["1", "01", "001", "29", "31", "32", "0", "99", "100", "1900", "2020", "10000"]
    found = [f"{w}{s}{n}" for w, s, n in itertools.product(words, separators, numbers)]
    found += [f"{n}{s}{w}" for n, s, w in itertools.product(numbers, separators, words)]
    words = ["Jan", "jan", "JAN", "January", "Janu", "Ja", "Sept", "Sep", "Septemb", "Mai", "Juli", "Marc", "May",
             "Jun", "Jul", "Ju", "Ma", "Dec", "Decembe", "Decemberx", "Mon", "Monday", "Feb", "Febr", "February"]
    for word in words:
        found += [f"{word} 1", f"1 {word}", f"{word} 2020", f"1 {word} 2020", f"{word} 1, 2020", word]
    pairs = [(" ", " "), ("-", "-"), ("/", "/"), (" ", ", "), (", ", " "), (", ", ", "), (" ", " , "), ("", ""),
             (", ", "/"), (".", " ")]  # fmt: skip
    for day, word, year, (one, two) in itertools.product(["1", "29", "31", "32", "0"], ["Jan", "Feb", "Sept"],
                                                         ["0", "99", "100", "1900", "2020"], pairs):  # fmt: skip
        found += [f"{day}{one}{word}{two}{year}", f"{word}{one}{day}{two}{year}"]
    found += ["Jan Feb 2020", "1 Jan Feb", "Jan 1 Feb", "Jan Jan", "Jan 1 Jan", "Monday, January 1, 2020",
              "Mon, Jan 1, 2020", "Monday 1/1/2020", "1st Jan 2020", "Jan 1st, 2020", "Jan 1st", "1st/1/2020",
              "January 1st", "the 1st of January", "-1/1/2020", "+1/1/2020", "(1/1/2020)", "1/1/2020-", "1/1/-2020",
              "1/-1/2020", "2020-01-15T12:30", "2020-01-15 12:30", "2020-01-15T12:30:00Z", "20200115",
              "2020-01-15 12:30:45.123", "Jan. 1, 2020", "1.1.2020", "1. Jan 2020", "Jan '20", "'20",
              "1/1/'20"]  # fmt: skip
    return found


def _dates_with_times() -> list[str]:
    dates = ["1/1/2020", "Jan 1, 2020", "1-Jan-2020", "1/2020", "Jan 2020", "2020-1-1", "1/1", "15 Jan"]
    times = ["12:00", "12:00 PM", "1 PM", "25:00", "12:30:45.5", "93:22.15", "0:0", "Jan:5"]
    joins = ["{d} {t}", "{t} {d}", "{d}  {t}", "{d}{t}", "{d},{t}", "{d}, {t}", "{d}-{t}", "{d}/{t}", "{t}{d}",
             "{t}, {d}"]  # fmt: skip
    return [join.format(d=d, t=t) for d, t, join in itertools.product(dates, times, joins)]


def _edges() -> list[str]:
    """The borderline cases: leading zeros in an hour, words and a meridiem
    where a time's field would be, a fraction before another colon, what
    may trail a time or a date, and what may join a time and a date."""
    found = [f"{hour}:{minute}" for hour, minute in itertools.product(
        ["000", "00", "001", "010", "099", "0001", "0005", "0010", "0023", "0024", "0099", "0100", "0999", "00100",
         "0000"], ["00", "30", "59", "60"])]  # fmt: skip
    found += ["0010:30:15", "0100:59:59", "0100:60", "010:Jan", "0100:Jan", "0010:OCT", "0010:OCT:26", "00 PM",
              "01 PM", "001 PM", "010 PM", "0010 AM", "0100 PM", "01:30 PM", "001:30 PM", "0012:30 PM"]  # fmt: skip
    words = ["AM", "PM", "A", "P", "am", "Jan", "Mon", "x", "T"]
    for word in words:
        found += [f"6 {word} : 27", f"{word}:5", f"5:{word}", f"1:2:{word}", f"6 {word}:27", f"6{word}:27",
                  f"{word} 6:27", f"6:27 {word}", f"6: {word}", f"6 {word} : 27 : 5"]
    found += ["6 30 : 27", "6 30:27", "6:30 27", "12 30", "6 PM:27:5", "6 PM : 27 PM", "6 AM : 70", "6 PM : 27.5",
              "24 PM : 27", "13 PM : 27", "6 PM : 60", "6 PM 1/1/2020", "1/1/2020 6 PM : 27"]  # fmt: skip
    for fraction in ["5", "05", "500", "5000", "391", "0", "60", "99", "100", "1000", ""]:
        found += [f"1:2.{fraction}:3", f"1:2.{fraction}:", f"1:2.{fraction}:3:4", f"25:2.{fraction}:3",
                  f"1:60.{fraction}:3", f"1:2.{fraction}:99", f"1:2.{fraction}:3.5"]
    found += ["1:2.5.5", "1:2:3.5:4", "1:2.5 PM", "1:2.5:3 PM", "1:2.5:3 1/1/2020", "1/1/2020 1:2.5:3",
              "0:0.5:0", "12:30.25:00", "Jan:2.5:3", "1:Jan.5:3"]  # fmt: skip
    times = ["12:00", "41:23", "12:00:30", "12:00 PM", "1:23"]
    for date, time, stray in itertools.product(["June-6", "1/1/2020", "Jan 1", "1/1", "6-June"], times,
                                                [" 4", " 45", " 2020", " 4 5", " x", " Jan"]):  # fmt: skip
        found += [f"{date} {time}{stray}", f"{time}{stray}", f"{time} {date}{stray}"]
    found += ["4 12:00", "4 12:00 1/1/2020", "12:00 4 PM", "12:00 PM 4"]
    trails = ["/", ",", ", ", " -", ".", "- ", ":", "-", " /", " ,", "--", "::"]
    for base in ["12:30", "12", "12:30 PM", "1/1/2020 12:30", "1/1/2020", "Jan 1, 2020", "Jan 1", "1-Jan", "1/1",
                 "Jan", "12:30:45", "12:30:45.5"]:  # fmt: skip
        found += [base + trail for trail in trails]
    for time, date in itertools.product(["12:00", "12:00 PM", "25:00", "12:00:30.5", "12:", "12:00:"],
                                        ["1/1/2020", "Jan 1, 2020", "15 Jan", "2020-1-1", "1/2020"]):  # fmt: skip
        found += [f"{time}{join}{date}" for join in ["/", "-", " / ", " - ", ",", " ,", " , ", ",  ", " :", "  "]]
    found += ["1/1 12:00 2020", "Jan 12:00 1, 2020", "Jan 1 12:00, 2020", "1 Jan 12:00 2020", "12:00 1/1 12:00",
              "1/1/2020 12:00 1/1/2020", "12:00 12:00", "Jan 1,2020", "Jan 1 ,2020", "Jan ,1 2020", "Jan , 1 2020",
              "Jan 1, 2020 PM", "1 PM Jan 1", "Jan 1 1 PM", "1 Jan 1 PM", "Jan 2020 1 PM", "1/2020 1 PM"]  # fmt: skip
    return found


#: Measured again in a workbook using the 1904 date system.
DATE_1904 = ["1/1/1900", "12/31/1903", "1/1/1904", "1/2/1904", "2/29/1904", "1/1/2020", "12/31/9999", "1/1/10000",
             "Jan 1904", "1/1904", "Jan 1, 1904", "1-Jan-04", "1/1/04", "1/1/00", "12/31/03", "1/1", "12:00",
             "25:00", "1/1/1904 12:00", "12/31/1903 23:59", "1/1/1904 0:00", "2/29/1900", "3/1/1900", "1900-1-1",
             "1904-1-1", "Jan 5", "5 Jan", "Jan 99", "1/99", "12/31/9999 23:59:59", "12/31/9999 24:00", "0:00",
             "Jan:5", "31-Dec-1903", "1-Jan-1904"]  # fmt: skip


def _random_strings(rng: random.Random, count: int) -> list[str]:
    """Random numbers and words joined by separators: strings the rules
    were worked out against."""

    def number() -> str:
        roll = rng.random()
        if roll < 0.35:
            text = str(rng.randint(0, 12))
        elif roll < 0.6:
            text = str(rng.randint(13, 31))
        elif roll < 0.8:
            text = str(rng.randint(32, 99))
        elif roll < 0.92:
            text = str(rng.choice([rng.randint(100, 999), rng.randint(1000, 9999), rng.randint(1900, 2100)]))
        else:
            text = str(rng.randint(0, 99999))
        if rng.random() < 0.1:
            text = "0" * rng.randint(1, 2) + text
        if rng.random() < 0.05:
            text += "." + str(rng.randint(0, 999))
        return text

    def word() -> str:
        roll = rng.random()
        if roll < 0.12:
            return rng.choice(["AM", "PM", "am", "pm", "A", "P", "mon", "x", "st", "T"])
        month = rng.choice(MONTHS)
        text = month[: rng.randint(2, len(month))]
        style = rng.random()
        return text.upper() if style < 0.2 else text.capitalize() if style < 0.8 else text

    joins = ["/", "-", " ", ":", ", "] * 3 + ["  ", ",", ".", " / ", " - ", " : ", "", " , "]
    made: list[str] = []
    for _ in range(count):
        parts = [word() if rng.random() < 0.3 else number() for _ in range(rng.choice([2, 2, 3, 3, 3, 4, 4, 5]))]
        text = parts[0] + "".join(rng.choice(joins) + part for part in parts[1:])
        if rng.random() < 0.06:
            text = " " * rng.randint(1, 2) + text
        if rng.random() < 0.05:
            text += " "
        made.append(text)
    return made


def _random_dates_and_times(rng: random.Random, count: int) -> list[str]:
    """Random dates and times, alone or joined either way, with what may
    trail them: the strings the rules were checked against once found."""

    def number() -> str:
        roll = rng.random()
        if roll < 0.4:
            text = str(rng.randint(0, 12))
        elif roll < 0.65:
            text = str(rng.randint(13, 31))
        elif roll < 0.8:
            text = str(rng.randint(32, 99))
        elif roll < 0.95:
            text = str(rng.choice([rng.randint(100, 999), rng.randint(1900, 2100), rng.randint(1000, 9999)]))
        else:
            text = str(rng.randint(0, 99999))
        return "0" * rng.randint(1, 2) + text if rng.random() < 0.1 else text

    def month() -> str:
        name = rng.choice(MONTHS)
        text = name[: rng.randint(2, len(name))]
        style = rng.random()
        return text.upper() if style < 0.2 else text.capitalize() if style < 0.8 else text

    def date() -> str:
        parts = [month() if kind == "m" else number() for kind in rng.choice(["nn", "nnn", "mn", "nm", "nmn", "mnn"])]
        joins = ["/", "-", " ", ", ", " / ", "  ", ",", "."]
        return parts[0] + "".join(rng.choice(joins) + part for part in parts[1:])

    def time() -> str:
        fields = [number() if rng.random() < 0.93 else rng.choice([month(), "AM", "PM"]) for _ in range(rng.randint(1, 3))]
        text = fields[0] + "".join(rng.choice([":", ":", " :", ": ", " : "]) + field for field in fields[1:])
        if rng.random() < 0.2:
            text += "." + str(rng.randint(0, 9999))[: rng.randint(0, 4)]
        if len(fields) == 1 and rng.random() < 0.5:
            text += ":"
        if rng.random() < 0.3:
            text += rng.choice([" AM", " PM", " am", " p", "PM", " A", "  PM"])
        return text

    made: list[str] = []
    for _ in range(count):
        roll = rng.random()
        if roll < 0.3:
            text = date()
        elif roll < 0.5:
            text = time()
        elif roll < 0.75:
            text = date() + rng.choice([" ", "  ", ", ", "", "-"]) + time()
        else:
            text = time() + rng.choice([" ", ", ", "/", "-", "", " - "]) + date()
        if rng.random() < 0.1:
            text += rng.choice([" 4", " 2020", "/", "-", ":", ".", ", ", " x", ",", " 1/1"])
        made.append(" " + text if rng.random() < 0.05 else text)
    return made


def corpus() -> list[str]:
    """Every string measured, each once."""
    found = _times() + _numeric_dates() + _named_dates() + _dates_with_times() + _edges()
    found += _random_strings(random.Random(20260926), 3000)
    found += _random_dates_and_times(random.Random(20260929), 5000)
    return list(dict.fromkeys(text for text in found if text.strip()))


_MEASURE = r'''
Private Type Exact
    Value As Double
End Type

Private Type Octets
    B(0 To 7) As Byte
End Type

Public Function Measure(ByVal Source As String, ByVal Target As String, ByVal Source1904 As String, _
                        ByVal Target1904 As String) As String
    Dim wb As Workbook
    Fill ActiveWorkbook.Worksheets(1), Source, Target
    ' A workbook of its own for the 1904 date system, closed unsaved.
    Set wb = Workbooks.Add
    wb.Date1904 = True
    Fill wb.Worksheets(1), Source1904, Target1904
    wb.Close SaveChanges:=False
    Measure = "measured"
End Function

Private Sub Fill(ws As Worksheet, ByVal Source As String, ByVal Target As String)
    Dim whole As String, lines() As String, f As Integer, n As Long, r As Long
    Dim cells() As Variant, c As Long, row As String
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
    ws.Range(ws.Cells(1, 1), ws.Cells(n, 1)).Formula = cells
    ws.Range(ws.Cells(1, 2), ws.Cells(n, 2)).Formula = "=VALUE(A1)"
    ws.Range(ws.Cells(1, 3), ws.Cells(n, 3)).Formula = "=DATEVALUE(A1)"
    ws.Range(ws.Cells(1, 4), ws.Cells(n, 4)).Formula = "=TIMEVALUE(A1)"
    Application.Calculate
    f = FreeFile
    Open Target For Output As #f
    For r = 1 To n
        row = ""
        For c = 2 To 4
            row = row & IIf(c > 2, vbTab, "") & Bits(ws.Cells(r, c).Value2)
        Next c
        Print #f, row
    Next r
    Close #f
End Sub

Private Function Bits(v As Variant) As String
    Dim e As Exact, o As Octets, i As Integer, s As String
    If IsError(v) Then
        Bits = "#"
        Exit Function
    End If
    e.Value = CDbl(v)
    LSet o = e
    For i = 7 To 0 Step -1
        s = s & Right$("0" & Hex$(o.B(i)), 2)
    Next i
    Bits = s
End Function
'''


def _double(bits: str) -> float | None:
    return None if bits == "#" else struct.unpack(">d", bytes.fromhex(bits))[0]


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
        paths = [Path(scratch) / name for name in ("texts.txt", "values.txt", "texts1904.txt", "values1904.txt")]
        # Bytes, so no line ending is translated on the way.
        paths[0].write_bytes(("\r\n".join(texts) + "\r\n").encode("ascii"))
        paths[2].write_bytes(("\r\n".join(DATE_1904) + "\r\n").encode("ascii"))
        with ExcelSession(HarnessConfig(lock_wait_s=wait)) as excel:
            excel.new_workbook()
            today = dt.date.today()
            result = excel.run_vba(_MEASURE, proc="Measure", args=tuple(str(path) for path in paths), timeout=900)
            if result.outcome != "passed":
                print(f"Excel refused the measurement: {result!r}", file=sys.stderr)
                return 1
        rows = paths[1].read_text(encoding="ascii").splitlines()
        rows_1904 = paths[3].read_text(encoding="ascii").splitlines()
    if len(rows) != len(texts) or len(rows_1904) != len(DATE_1904):
        print(f"expected {len(texts)} and {len(DATE_1904)} rows, Excel wrote {len(rows)} and {len(rows_1904)}",
              file=sys.stderr)  # fmt: skip
        return 1
    record = {"measured": today.isoformat(), "fields": ["text", "VALUE", "dateOrTime"]}
    lines = [json.dumps(record)[:-1]]
    for name, strings, found in (("cases", texts, rows), ("cases_1904", DATE_1904, rows_1904)):
        cases: list[list[object]] = []
        for text, row in zip(strings, found, strict=True):
            value, whole, rest = (_double(bits) for bits in row.split("\t"))
            read = whole is not None or rest is not None
            if read and (value is None or whole != math.floor(value) or rest != value - math.floor(value)):
                print(f"DATEVALUE and TIMEVALUE of {text!r} are not VALUE's days and the rest", file=sys.stderr)
                return 1
            cases.append([text, value, read])
        lines[-1] += f', "{name}": ['
        lines += [f"  {json.dumps(case)}," for case in cases[:-1]] + [f"  {json.dumps(cases[-1])}", "]"]
    lines[-1] += "}"
    FIXTURE.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"wrote {FIXTURE.relative_to(ROOT)} ({len(texts)} strings, {len(DATE_1904)} in the 1904 system)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
