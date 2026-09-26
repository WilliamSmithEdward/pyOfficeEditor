# Excel fixtures

Two producers write OOXML, both legally and differently, so the suite tests
against both. Telling them apart is mechanical: an Excel-authored package
carries ZIP "version made by" 45 and a `0xa220` growth-hint extra field in
some local headers, and a library-authored one carries 20 and none.

## Authored by Excel, committed here

| File | Parts | What it is for |
|------|-------|----------------|
| `excel_authored_minimal.xlsm` | 9 | The smallest real Excel package: growth hints on five parts, `mc:Ignorable` on the worksheet root, a CRLF after the XML declaration. The byte-fidelity gate for `_zip`, `_xml` and `opc`. |
| `excel_authored_powerquery.xlsx` | 32 | A wider relationship graph and a deeper part tree, so path arithmetic and relationship resolution are exercised on something real. It also carries a UTF-16LE DataMashup part, which is how we know Office does not write UTF-8 everywhere. |
| `excel_authored_binary.xlsb` | 12 | The binary workbook format: five `.bin` parts including the worksheets. Proves the container layer does not assume its members are text. |

All three came from sibling projects by the same author, in
`pyOpenVBA`: `tests/live_excel_testing/xlsm_file_with_no_vba_entered_yet.xlsm`,
`examples/power_query_demo.xlsx`, and
`tests/live_excel_testing/test_macro_workbook.xlsb`. None is generated; do
not regenerate them with a library, because their value is that Excel wrote
them.

## Authored by a library, generated during the test run

`conftest.py` builds openpyxl workbooks in a temporary directory. They
cover what the committed fixtures do not: a `sharedStrings` part, formula
cells with cached values, and the ZIP data-descriptor shape a streaming
writer produces. Nothing is committed, so the suite stays runnable with no
fixtures to refresh.

## Authored by Excel, built by script

`scripts/build_excel_fixtures.py` drives real Excel through
[pyvbaharness](https://github.com/WilliamSmithEdward/pyVBAharness) to author
twenty more. These are committed like the rest: the script exists so a
fixture's *content* can be changed deliberately and reproduced, not so the
suite rebuilds them. Run it on a Windows machine with Excel; tests that need
a fixture skip when it is absent.

| File | What it carries |
|------|-----------------|
| `empty.xlsx` | The smallest thing Excel will save as `.xlsx`. |
| `sample.xlsx` | Shared strings, formulas, dates, booleans, an error cell, a merged range, escaped and space-padded text, two sheets. |
| `structures.xlsx` | Two ListObjects, one with a totals row and a calculated column; defined names at workbook and sheet scope; column widths; a hyperlink with its own external relationship. |
| `refused.xlsx` | Every element that used to make an insertion refuse, on one sheet: all five address notations, including the two zero-based ones and the two cases Excel checks for agreement. |
| `settings.xlsx` | Sheet protection, tab colour, view settings and the three visibility states. |
| `links.xlsx` | Hyperlinks of every kind, an autofilter with two sorts of criteria, and outline grouping on rows and columns. |
| `geometry.xlsx` | One shape of each `MsoAutoShapeType`, so the preset geometry table is held to what Excel wrote rather than to a transcription. |
| `filters.xlsx` | One autofilter criterion per sheet, with `filters_answers.json` recording what Excel answered and how many rows it hid. The hidden counts are the measurement: Excel stores a filter twice and does not recompute it on open. |
| `controls.xlsm` | One of every Forms control, wired to cells and ranges, with `controls_answers.json` beside it recording what Excel's object model answered for each. |
| `pictures.xlsx` | Pictures from `Shapes.AddPicture`: a PNG at its own size, one at 144 dots to the inch, the first stretched, and a GIF with alternative text, with `pictures_answers.json` recording what Excel's object model said of each. The recipe writes the images itself, from hex. |
| `comments.xlsx` | Notes, plain, on two lines, showing, resized, on A1 and with a word in bold, and one sharing its sheet's VML part with a button, with `comments_answers.json` recording what Excel's object model said of each. The author is whoever ran the script: Excel takes it from the Office user name, and VBA cannot set it. |
| `richtext.xlsx` | Seven cells of text in more than one font, made through `Characters`: bold, a colour, a bigger size, italic, underline, superscript, another typeface, a line break, and a cell whose own font is bold. `richtext_answers.json` records the font Excel reported at every change. |
| `escapes.xlsx` | Text XML cannot carry as it is, control characters, a lone carriage return, a CRLF and a literal `_x0041_`, everywhere a workbook keeps text: cells, formulas and their results, a formula naming a sheet called `a_x0041_b`, a note, a validation's messages, a hyperlink's tip, a table's headers, a page header and a defined name's comment. `escapes_answers.json` records each character Excel read back. The page header brings a printer settings part, for Microsoft Print to PDF. |
| `styles.xlsx` | Every built-in cell style Excel lists, each applied by Excel to a cell of its own beside its name, a style of the workbook's own, and Good and Currency put over formatting a cell already had. `styles_answers.json` records what Excel showed for each cell, its style, number format, font, fill, alignment and edges, and which aspects each style includes. |
| `charts.xlsx` | A column chart with a typed title and a line chart over two areas on the data sheet, a pie whose title is linked to a cell and a scatter chart on a second sheet, and a bar chart on a chart sheet. `charts_answers.json` records each chart's series formulas as Excel reported them, and every reference in each chart part as Excel wrote it, in the file as built and in the file Excel saved after each of nine edits: rows and columns inserted and deleted, and the data sheet renamed. |
| `chartkinds.xlsx` | One chart of each kind Excel's Insert Chart makes, added with `Shapes.AddChart2` and its default style from `Data!A1:C6`: column, bar, line, line with markers, pie, doughnut, scatter and area, and a column chart with a title typed in. `chartkinds_answers.json` records each chart's series formulas. The parts are what a chart added here is compared with. |
| `pivots.xlsx` | Two pivot tables reading `Data!A1:C11` through caches of their own, one on a sheet of its own and one on the data sheet where edits reach it. `pivots_answers.json` records, for the file as built and after each of nineteen edits Excel made, where each pivot table was, what its cache read and how many caches the workbook kept, read from the file Excel saved, or that Excel refused the edit. |
| `pivotdata.xlsx` | Twenty-six pivot tables in each layout Excel offers, over `Data!A1:F13` and `Data!H1:J5`: nested, across, tabular, outline, filtered, without totals or subtotals, with values down the side, renamed items and fields, hidden items, dates grouped by years and months, numbers in bins, a calculated field and a share of the total. Sheet Q holds 147 GETPIVOTDATA formulas reading them, each with Excel's answer cached beside it. Excel removed personal information as it saved. |
| `datatables.xlsx` | What-if data tables from `Range.Table` over a small loan model: rates down a column against two formulas, years along a row, both at once, a chain through another sheet and back, a formula for an input cell, a branch the input decides, and a blank, an error, text and a formula among the values tried. The model is built twice, on Model and on Moved with other years and principal, so setting one's inputs to the other's and recalculating is held to Excel's answer. Excel removed personal information as it saved. |
| `errorchecks.xlsx` | Cells for each of error checking's ten rules and their near misses, a sheet apiece: text that is a number or a date with a two-digit year, alone or in a formula; errors, NA() among them, unlocked formulas and misleading formats; inconsistent formulas and ranges that leave out a number, SUMIF, absolute rows and a name among them; references to empty cells in rows that store cells and rows that do not, and past the last used row; a table with a calculated column's exceptions and values its validation refuses, a blank one included, and another with a validation of every kind; and errors ignored. `errorchecks_answers.json` records every cell with the rules `Range.Errors` says catch it, ignored ones included. Excel removed personal information as it saved. |

`shapes.xlsm` is committed but has no recipe here: it predates the script.
It carries one of every shape a sheet can hold, with `shapes_answers.json`
beside it, and a Button wired to nothing, which is why `controls.xlsm`
exists.

A measured fixture is worth more than its workbook. `controls_answers.json`
is what settles a disagreement, because the reader is held to what Excel
said rather than to a reading of the markup: the current value alone is
spelled three ways, and the off state is -4146 rather than 0.

## Measured by script

Five corpora record what Excel did rather than what it wrote, and each case
in them is a test. All five came from Excel 16.0 build 20326 in en-US. The
first two scripts run beside other harness sessions, since they only run
hidden macros; the other three wait for the lock like the fixture builder.

| File | Built by | What it records |
|------|----------|-----------------|
| `formulas.xlsx` | `scripts/measure_formulas.py` | 10,958 formulas on 115 sheets, each sheet named for what it probes, and the value Excel calculated for each. The inputs, sixteen sheets such as `Numbers`, `Pairs` and `Powers`, were written by this library as exact doubles into a copy of `empty.xlsx`; Excel then typed in the formulas, calculated and saved. `About` records the build and the day it was measured, which a text naming a day without a year needs. `test_excel_formula_corpus.py` calculates every formula and holds it to Excel's result. |
| `number_formats.json` | `scripts/measure_number_formats.py` | The text Excel showed for 503 format codes, each under 21 to 73 values, 27,898 texts in all across both date systems, and the nine codes it refused; the code of all 164 built-in ids with its text for five sample values; and the text of 782 cells of the committed fixtures. `test_excel_numfmt.py` renders every one. |
| `filter_semantics.json` | `scripts/measure_filters.py` | 358 criteria, each applied by Excel to one of 31 columns built to trip it, with the rows it hid and the markup it stored. 261 went through the object model; 93 were written straight into a package and applied from the file, for markup the object model will not write; 4 are special cases, such as rows hidden by hand. `test_excel_filter_semantics.py` evaluates every one. |
| `text_checks.json` | `scripts/measure_text_checks.py` | 6,683 strings, each typed into a cell after an apostrophe, and whether error checking called it a number stored as text or a date with a two-digit year: a grid of numbers and month names joined every way, the separators, digit counts and spellings that settle a borderline case, and random strings made from the same pieces with a fixed seed. `test_excel_errorchecks.py` judges every one. |
| `date_texts.json` | `scripts/measure_date_texts.py` | 18,899 strings, each put in a cell as text, with VALUE of it as the exact double Excel calculated or an error, and whether DATEVALUE and TIMEVALUE read it, which they do as VALUE's whole days and the rest: times with fields of every width and AM or PM, numbers and month names joined by every separator, dates and times joined either way and what may trail them, and random strings made with fixed seeds, the last 4,945 measured only after the rules were found; and 35 more in a workbook using the 1904 date system. `test_excel_calc.py` reads every one. |

Rebuilding any of them replaces it, since a corpus has no content to keep:
what it holds is whatever Excel answers. The dates in
`filter_semantics.json` are relative to the day it was built, which it
records, and its tests measure "today" from that day; `formulas.xlsx` does
the same for the texts in it that name a day without a year,
`text_checks.json` for `2/29`, a day only in a leap year, and
`date_texts.json` for every date it reads without a year.

A fixture that already exists is left alone, because Excel stamps every part
with a fresh revision GUID and so never produces the same bytes twice.
Rebuilding a committed fixture churns it for nothing and buries the real
change in the diff. Pass `--force` when a fixture's *content* needs to change,
which is a deliberate act.

Two things `sample.xlsx` taught us, both of which would have produced a
quietly wrong cell layer. `test_opc.py` pins each one.

**A range-assigned formula is a shared formula.** Assigning `=B2*C2` to
D2:D5 in one statement stores the text once:

```xml
<c r="D2"><f t="shared" ref="D2:D5" si="0">B2*C2</f><v>510</v></c>
<c r="D3"><f t="shared" si="0"/><v>1445</v></c>
```

Three of those four cells carry a cached value and no formula text at all.
Reading `<f>`'s text per cell reports an empty formula for D3, D4 and D5;
the real formula has to be translated from the master cell's, shifting its
relative references. A formula typed into a single cell is stored plainly,
so both shapes occur in one sheet.

**Relationship order is not sheet order.** Excel wrote
`rId2 -> worksheets/sheet2.xml` *before* `rId1 -> worksheets/sheet1.xml`, so
indexing the worksheet relationships picks the second sheet. Order and names
live in `<sheets>` inside `xl/workbook.xml`, and each entry names its
relationship by `r:id`. `sheet_part_by_name` in `test_opc.py` is the correct
navigation.
