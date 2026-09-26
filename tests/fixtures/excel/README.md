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
thirty more. These are committed like the rest: the script exists so a
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
| `sorts.xlsx` | One sort to a sheet, saved before Excel sorted anything: values of every kind both ways, words a collation and case decide between, ties settled by a second key, a row's formulas of every reference with its formats, note and link, a validation, a conditional format and row heights that stay put, a range inside a wider table, no header, hidden and filtered rows, blanks, a range well past the data, shared formulas, and the merged cells and arrays Excel refuses; then tables, one with a totals row, a calculated column and a note, one away from A1 sorted by two keys with case matched, one filtered and one without a header row, a table a range sort takes whole, totals row and all, and a sheet's filter. `sorts_sorted.xlsx` is the same workbook once Excel had sorted every sheet through the Sort object of the range, the table or the filter, and `sorts_answers.json` records each sort and the error Excel refused it with. Excel removed personal information as it saved. |
| `duplicates.xlsx` | Sheet Pairs holds 92 pairs of values, each in a range of its own with a marker beside each value: numbers shown alike and apart, numbers against text and logicals, text a filter's collation takes for one or tells apart, logicals against their text, errors, blanks and empty text. The other sheets hold keys over two columns with cells outside the range and formulas reading into it; formulas, fills, notes and links on rows kept and removed; validation and conditional formats cut back by the rows cleared; no header; hidden rows; array formulas; a last row taken for a total, and one that is not; a range running past the data; and the merged cells and arrays Excel refuses. Then tables, which Excel cleans whole: one with formulas, notes, links, formats, validation, conditional formats and a chart around it, one with a totals row besides, and one with formats and validation across its totals row; ranges in part of a table, in one cell and in its totals row; a calculated column; rows hidden by hand and by a table's filter, one of them with a totals row, which Excel refuses; no header row; empty rows; a last row that is not a total; a sort the table records; nothing to remove; and ranges running past a table, which Excel refuses. Sheet Reader holds formulas reading two of the tables, and four defined names read them too. `duplicates_removed.xlsx` is the same workbook once Excel had removed every range's duplicates, and `duplicates_answers.json` records each range, the columns compared and the error Excel refused it with. Excel removed personal information as it saved. |
| `copies.xlsx` | A source block to a sheet and 44 copies Excel made of them through `Range.Copy Destination:=`: values of every kind with their formats; references of every kind, copied on the sheet, to another and up off it; merged cells in the source, partly in it, and at the destination whole and in part; notes, a note's box moved and sized, links, and notes and links pasted over; validation and conditional formats copied on the sheet and to another and cut back at the destination; a copy over its own source; arrays whole and in part, a spill with its cells and its cells alone, and pastes onto a spill's formula and onto a cell it spilled into; rows hidden by hand and under the sheet's filter and a table's, and a hidden column; a shared formula; destinations a multiple of the source and not; formats on blank cells; whole rows and columns; and pastes into a table and out of one. No thread: a workbook Excel wrote one in carries who wrote it. `copies_pasted.xlsx` is the same workbook once Excel had made every copy, in order, and `copies_answers.json` records each copy's source and destination and the error Excel refused it with. Excel removed personal information as it saved. |
| `paste_specials.xlsx` | 109 pastes Excel made through `Range.PasteSpecial`, and Paste Link through `Worksheet.Paste Link:=True`, each sheet a source and a destination with things of its own: a rich block of every kind of value with formats, a note, a link, merged cells, validation and conditional formats pasted by every choice, skipping blanks, transposed and with each operation; blanks, merged cells and arrays in the way; number formats; conditional formats beside the destination's; blanks skipped note by note, link by link and rule by rule; transposed references of every kind, in the source and out of it, pushed off the sheet and on another sheet; the operations over every kind of value and formula, and the numbers they write into a formula; whole rows and columns; row and column formats under blank cells; widths; links; Paste Link over formats, tiled, merged and filtered; a filter, tiles, a table and a shape. `paste_specials_pasted.xlsx` is the same workbook once Excel had made every paste, in order, the formulas an operation built entered again from their text and the workbook calculated, as opening it reads them; `paste_specials_answers.json` records each paste's source, destination, choice, operation, whether it skipped blanks and transposed, and the error Excel refused it with. Excel removed personal information as it saved. |
| `note_boxes.xlsx` | Notes whose grid changes under them, a case to a sheet: a column's width set left of a note's box, under it and right of it, made narrower, and three at once; a column hidden under the box, at its first column and left of it, and hidden and shown again; a row's height set above the box, under it and below it, and made small; a row hidden above it and under it; a filter, and folded groups of rows and of columns; and a note of each placement, and one shown, under several changes. `note_boxes_changed.xlsx` is the same workbook once Excel had made every change, in order, and `note_boxes_answers.json` records each change: the sheet, what changed, the columns or rows, and the width or height Excel was given. Excel removed personal information as it saved. |

`shapes.xlsm` is committed but has no recipe here: it predates the script.
It carries one of every shape a sheet can hold, with `shapes_answers.json`
beside it, and a Button wired to nothing, which is why `controls.xlsm`
exists.

A measured fixture is worth more than its workbook. `controls_answers.json`
is what settles a disagreement, because the reader is held to what Excel
said rather than to a reading of the markup: the current value alone is
spelled three ways, and the off state is -4146 rather than 0.

## Measured by script

Eight corpora record what Excel did rather than what it wrote, and each case
in them is a test. All eight came from Excel 16.0 build 20326 in en-US. The
first two scripts run beside other harness sessions, since they only run
hidden macros; the other six wait for the lock like the fixture builder.

| File | Built by | What it records |
|------|----------|-----------------|
| `formulas.xlsx` | `scripts/measure_formulas.py` | 10,958 formulas on 115 sheets, each sheet named for what it probes, and the value Excel calculated for each. The inputs, sixteen sheets such as `Numbers`, `Pairs` and `Powers`, were written by this library as exact doubles into a copy of `empty.xlsx`; Excel then typed in the formulas, calculated and saved. `About` records the build and the day it was measured, which a text naming a day without a year needs. `test_excel_formula_corpus.py` calculates every formula and holds it to Excel's result. |
| `number_formats.json` | `scripts/measure_number_formats.py` | The text Excel showed for 503 format codes, each under 21 to 73 values, 27,898 texts in all across both date systems, and the nine codes it refused; the code of all 164 built-in ids with its text for five sample values; and the text of 782 cells of the committed fixtures. `test_excel_numfmt.py` renders every one. |
| `filter_semantics.json` | `scripts/measure_filters.py` | 358 criteria, each applied by Excel to one of 31 columns built to trip it, with the rows it hid and the markup it stored. 261 went through the object model; 93 were written straight into a package and applied from the file, for markup the object model will not write; 4 are special cases, such as rows hidden by hand. `test_excel_filter_semantics.py` evaluates every one. |
| `text_checks.json` | `scripts/measure_text_checks.py` | 6,683 strings, each typed into a cell after an apostrophe, and whether error checking called it a number stored as text or a date with a two-digit year: a grid of numbers and month names joined every way, the separators, digit counts and spellings that settle a borderline case, and random strings made from the same pieces with a fixed seed. `test_excel_errorchecks.py` judges every one. |
| `date_texts.json` | `scripts/measure_date_texts.py` | 18,899 strings, each put in a cell as text, with VALUE of it as the exact double Excel calculated or an error, and whether DATEVALUE and TIMEVALUE read it, which they do as VALUE's whole days and the rest: times with fields of every width and AM or PM, numbers and month names joined by every separator, dates and times joined either way and what may trail them, and random strings made with fixed seeds, the last 4,945 measured only after the rules were found; and 35 more in a workbook using the 1904 date system. `test_excel_calc.py` reads every one. |
| `ignored_errors.json` | `scripts/measure_ignored_errors.py` | 47 sequences of steps, each a rule ignored or no longer ignored in a cell through `Range.Errors(i).Ignore` on a sheet of its own, with the rules Excel then said each cell ignored and the `<ignoredErrors>` it saved; and 15 files holding `<ignoredErrors>` of every shape, with what Excel read from each and what it saved again. Where Excel saved less than it held, the steps show its own loss. `test_excel_errorchecks.py` replays every one. |
| `sheet_names.json` | `scripts/measure_sheet_names.py` | 1,246 sheet names, each given to one sheet in turn, with what Excel then wrote for `=S!B3` and `=SUM(S:T!B3)`: logicals, names that read as references in A1 or R1C1 and near misses, digits and punctuation, letters past ASCII, and every character of eight Unicode blocks, inside a name and at its start. And 23 renames of a sheet a formula or a defined name names, alone or as an end or the middle of a 3D reference's span, to and from names that need quotes, with what each read before and after. `test_excel_sheets.py` holds every name and every rename to it. |
| `table_resizes.json` | `scripts/measure_table_resize.py` | 167 tables resized through `ListObject.Resize` from `A1:C8`, or from `A1:C9` with a totals row, to `A1:C5`, each on a sheet of its own, read from the files saved before and after: 74 formulas beside a table reading it every way that tells the rules apart, with a totals row and without; every span of rows in column B as a conditional format's range and as a validation's; and conditional formats and validations of each kind of rule on the totals row and below it. Remove Duplicates resizes a table this way, and `test_excel_table_resize.py` holds every case. |

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
