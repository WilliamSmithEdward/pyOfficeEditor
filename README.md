# pyOfficeEditor

Edit the document surface of Microsoft Office files in pure Python. No Office
installation, no COM, no dependencies.

Its sister project [pyOpenVBA](https://github.com/WilliamSmithEdward/pyOpenVBA)
edits the VBA project inside an Office file. This one edits the document: the
cell, the formula, the paragraph, the slide, the table, the query.

> **Status: early, and growing.** The Excel surface reads and writes cells,
> values, formulas, dates, sheets, formatting, merged ranges, tables, row and
> column dimensions, frozen panes and defined names, each verified against real
> Excel. Conditional formatting reads and writes every rule family Excel has.
> Data validation covers every rule Excel has, dropdowns included. Sheets
> protect, hide, colour their tabs and set their own view.
> Rows and columns can be inserted and deleted, with every reference in the
> workbook following or breaking exactly as Excel breaks it, `A:A` and `2:4`
> included. Nothing about a sheet makes that refuse any more: validation,
> protected ranges, sorts, scenarios, shapes, form controls and comments all
> move with it. Word, PowerPoint and Access follow, in that order.

```python
import datetime as dt
from pyofficeeditor.excel import Border, Dxf, Workbook, cell_is, gradient

with Workbook.open("orders.xlsx") as book:
    sheet = book["Data"]

    sheet["A2"].value        # 'North'                  stored as an index
    sheet["F2"].value        # datetime.date(2026, 1, 15)   stored as 46037
    sheet["D3"].formula      # 'B3*C3'                  stored nowhere at all
    sheet["B8"].value        # CellError('#DIV/0!')     not the text of one

    sheet["B2"].value = 200
    sheet["G1"].value = dt.date(2026, 7, 4)
    sheet["G2"].formula = "=SUM(B2:B5)"

    sheet.range("A1:F1").apply_font(bold=True)   # each cell keeps its own rest
    sheet["B2"].fill = "FFFF00"
    sheet["B3"].border = Border.all_sides("thin", "FF0000")

    sheet["A20"].value = "wide heading"
    sheet.merge("A20:C20")
    sheet["B20"].merged_range                    # RangeRef('A20:C20')

    table = sheet.add_table("Sales", "A1:F20", totals_row=True)
    table.column_names                           # from the header row
    table.data_range                             # excludes header and totals

    sheet.set_row_height(1, 24)                  # points, exact
    sheet.set_column_hidden(4, True)
    sheet.freeze_panes("B2")                     # pins row 1 and column A

    sheet.add_conditional_format(
        "B2:B20", cell_is("greaterThan", 100), dxf=Dxf.of(fill="FFC7CE", bold=True)
    )
    sheet.add_conditional_format("C2:C20", gradient())   # three-colour scale

    sheet.insert_rows(3, 2)                      # every reference follows
    sheet.delete_columns(5, 1)                   # SUM(E2:E9) -> #REF!

    summary = book.add_sheet("Summary", index=0)
    summary["A1"].formula = "=SUM(Data!D2:D5)"
    book.add_defined_name("Totals", "Data!$D$2:$D$5")
    book.rename_sheet("Data", "Q1 Data")     # formulas and names both follow
    book.save()
```

Each of those four reads has a plausible wrong answer that a naive
implementation gives instead: `6` for the string, `46037` for the date, an
empty formula for `D3`, and a string that compares equal to `"#DIV/0!"` for
the error. Getting them right is most of what the Excel modules do.

## Why this exists

openpyxl already sets a cell, python-docx a paragraph, python-pptx a slide.
Each covers one host. This library is aimed at the ground they leave
uncovered:

- **One API across four hosts**, including Access, which none of them touch.
- **The legacy binary formats**, `.xls`, `.doc` and `.ppt`, which pyOpenVBA
  deliberately treats as opaque.
- **The analysis layer**: Excel formula parsing and linting, legacy data
  connections, and Power Query.
- **Byte fidelity as a correctness property**, not a nicety. See below.

## Byte fidelity

Editing one cell of a worksheet must leave every other byte of the package
alone. Otherwise a one-cell change produces a diff nobody can review, and a
save that should be a no-op is not one.

That turns out to rule out the obvious building blocks, for reasons that are
measurable rather than theoretical.

**The XML.** Excel writes a worksheet as a declaration, a CRLF, then a single
line whose root element declares `mc:Ignorable="x14ac xr xr2 xr3"` *before* it
declares the `x14ac` and `xr` prefixes themselves. Round-tripping that through
`xml.etree.ElementTree` renames the prefixes to `ns0` and `ns1`, reorders the
declarations, and drops the CRLF. So parts are parsed into a tree where every
node keeps the source text it was cut from. Serializing a node nobody touched
copies those bytes; serializing a node that changed rebuilds it and recurses.
Editing one cell rewrites that cell's row and copies the rest.

**The container.** In a freshly authored workbook, `[Content_Types].xml`
carries a 520-byte extra field in its *local* ZIP header and none in the
central directory. `zipfile.ZipInfo.extra` exposes only the central copy, so
anything built on `zipfile`'s writer silently drops those bytes. The field is
the Microsoft Open Packaging Growth Hint (tag `0xa220`): padding Excel reserves
so it can grow a part in place. This library reads and writes the container
itself, field for field, and copies each untouched member's stored bytes
without inflating them.

The gate: reading a package and writing it back with nothing changed
reproduces the input exactly. It holds for Excel-authored `.xlsm`, `.xlsb` and
`.xlsx`, for openpyxl-authored `.xlsx`, and for archives whose members carry
data descriptors.

## Architecture

```
+--------------------------------------------------------+
| excel/        Workbook, Worksheet, Range, Cell         |
|   _reference  A1 notation, bijective base-26 columns   |
|   _values     the six cell encodings, and serial dates |
|   _styles     number formats, which is how a date is   |
|               told from a number                       |
|   _formats    fonts, fills, borders, alignment, as     |
|               immutable values                         |
|   _tables     ListObjects: their own parts and wiring  |
|   _dimensions widths, heights, hiding, frozen panes    |
|   _names      defined names, and the rules tables share |
|   _validation what a cell accepts, and the inverted    |
|               attribute behind its dropdown            |
|   _conditional  cfRules, and the compatibility formula |
|                 that makes them fire                   |
|   _dxf        differential formats: what a rule paints |
|   _addresses  the five notations an address is spelled  |
|               in, two of them zero-based                |
|   _rowcol     inserting and deleting rows and columns,  |
|               and moving everything that records a     |
|               cell address                             |
|   _tokens     a formula, broken into editable pieces    |
|   _formulas   shifting and breaking references, for    |
|               shared formulas, a renamed sheet, and    |
|               the #REF! a deletion leaves behind       |
|   _sharedstrings   the per-workbook string table       |
|   _schema     where a child element has to go          |
+--------------------------------------------------------+
| word / powerpoint / access   to follow, in that order  |
+--------------------------------------------------------+
| opc.py        Open Packaging Conventions               |
|   - parts, cached and flushed only when modified       |
|   - [Content_Types].xml: defaults and overrides        |
|   - relationships, resolved the way Office resolves    |
+--------------------------------------------------------+
| _xml.py       XML that reproduces its own source       |
|   - hand-written parser, no stdlib XML module          |
|   - per-node source spans and dirty propagation        |
|   - prefixes, attribute order and empty-tag form kept  |
+--------------------------------------------------------+
| _zip.py       the ZIP container, field for field       |
|   - local and central headers kept apart               |
|   - untouched members copied as stored bytes           |
+--------------------------------------------------------+
```

Each layer knows the one below it and not the one above. `_zip.py` knows
nothing about OOXML; `_xml.py` knows nothing about packages.

`docs/architecture.md` is the contributor reference.

## Untrusted input

A document is untrusted input, so the XML parser is deliberately less capable
than a conforming XML processor:

- `<!DOCTYPE` is refused outright. No DTD processing means no external entity
  resolution and no entity-expansion amplification.
- Only the five predefined entities and numeric character references resolve.
  Any other `&name;` raises.
- Nesting deeper than 256 elements raises instead of exhausting the stack.
- zip64 archives and compression methods other than stored and deflate are
  refused by name rather than guessed at.

None of the five restricts anything OOXML is allowed to contain.

## Install

```bash
pip install pyOfficeEditor
```

Python 3.10 or newer. No runtime dependencies.

## Three things Excel does that catch readers out

Each is measured, each is pinned by a test, and each gives a wrong answer
rather than an error if you miss it.

**A formula assigned to a range is stored once.** Excel writes the text on the
group's first cell and leaves the rest pointing at it by index:

```xml
<c r="D2"><f t="shared" ref="D2:D5" si="0">B2*C2</f><v>510</v></c>
<c r="D3"><f t="shared" si="0"/><v>1445</v></c>
```

D3's formula is not in the file. It is D2's, shifted down a row. Finding the
references to shift is the delicate part, because `LOG10(x)` contains `G10`,
`"A1"` is a string literal, and `'My Sheet A1'!B2` has a reference inside a
quoted sheet name.

**A date is a number, and only its number format says otherwise.**
`2026-01-15` is stored as `46037`. Confirming it is a date means following
`s="2"` to `cellXfs[2]`, its `numFmtId` to a format code, and the code to its
date tokens, skipping the quoted, escaped and bracketed parts that only look
like them: `#,##0 "days"` is not a date and `[h]:mm` is. Excel also numbers
dates as though 1900 were a leap year, so serial 60 is a 29 February that
never happened and is refused rather than reported as 1 March.

**A changed cell invalidates cached results.** A formula cell stores the value
it last evaluated to, so setting `B2` leaves `D2`'s cached `510` behind. A
workbook this library modified is saved with `fullCalcOnLoad` set and the
`calcChain` part dropped, so Excel recalculates on open. The live gate proves
it: after changing an input, Excel reports the recomputed number rather than
the stale one still written in the file.

**A worksheet's children are a sequence, not a set.** SpreadsheetML declares
them in order, and Excel refuses a file that breaks it rather than repairing
one. The natural thing to do with a missing element is append it, and that is
wrong whenever anything that must follow it is already there: a sheet Excel
authored starts with `sheetPr`, so a missing `dimension` does not go at the
front, and a workbook usually ends with `extLst`, so a missing `calcPr` does
not go at the end. `_schema.py` writes all three orders down.

**Formatting is shared, so it cannot be edited in place.** A cell carries an
index into `cellXfs`, and two hundred cells may carry the same one. Changing
that entry restyles all of them, which is never what "embolden this cell"
meant. So formats are immutable values: read the cell's, derive a new one,
and the workbook finds or appends the entry that matches. Two consequences
worth knowing:

- A cell with no `s` attribute is not unformatted. It uses `cellXfs[0]`,
  which names the workbook's default font. Resolving it to an empty format
  instead would make "add bold" silently change the typeface.
- A solid fill's colour goes in `fgColor`, not `bgColor`. The names suggest
  otherwise, and putting it in `bgColor` produces a cell that looks unfilled.

## Lower-level access

The packaging layer is public, for anything the host surfaces do not cover:

```python
from pyofficeeditor import OpcPackage

with OpcPackage.open("book.xlsx") as package:
    # Navigate the way Office does: by relationship, not by path.
    workbook_part = package.main_document_part()          # 'xl/workbook.xml'
    document = package.xml(workbook_part)
    document.root.require("sheets")
    package.save()          # every untouched part keeps its original bytes
```

## Development

```bash
python -m pip install -e ".[dev]"
python -m pytest -p no:randomly
pyright src tests
```

`-p no:randomly` keeps ordering reproducible. Strict Pyright must report zero
errors on `src` and `tests` before anything merges, and a behavior change lands
with its test in the same commit.

The suite needs no Office installation. It tests against six committed
Excel-authored packages, an `.xlsb` among them, and against openpyxl-authored
ones generated during the run, because a reader that only ever sees one
producer's output encodes that producer's habits as rules.

There is also a live gate, which is the only check that can prove Excel
accepts what this library writes. It drives real Excel through
[pyVBAharness](https://github.com/WilliamSmithEdward/pyVBAharness), opens an
edited workbook, and reads the cells back through Excel's own object model:

```bash
python -m pip install -e ".[dev]" --group live
python scripts/build_excel_fixtures.py
RUN_LIVE_EXCEL=1 python -m pytest -m live -p no:randomly
```

Windows and Excel only. Everything else in the suite runs anywhere.

## License

MIT. See [LICENSE.md](LICENSE.md).
