# pyOfficeEditor Architecture

The internal layering, module responsibilities, conventions and decisions.
This is the contributor reference; end users should read the
[README](../README.md).

---

## 1. Scope and the line against pyOpenVBA

Two libraries edit Office files without Office. The line between them is the
one Office itself draws:

| | [pyOpenVBA](https://github.com/WilliamSmithEdward/pyOpenVBA) | pyOfficeEditor |
|---|---|---|
| Owns | the VBA project and its form designers | the document surface |
| Reads and writes | `vbaProject.bin`, `dir`, `PROJECT`, module streams, MSForms designer streams | cells, formulas, styles, paragraphs, slides, tables, queries |
| Container | MS-CFB, and OPC only far enough to reach `vbaProject.bin` | OPC in full, and MS-CFB for the legacy binary formats |

Neither depends on the other. Both are pure Python with no runtime
dependencies, which is a deliberate duplication: this project reimplements
the OPC and container work rather than importing it, so the two can be
released and reasoned about independently.

### 1.1 What moves here

Power Query is pyOfficeEditor's, not pyOpenVBA's. `PowerQueryWorkbook` and
the SpreadsheetML helpers currently private to `pyopenvba.powerquery`
(`_sheets.py`: `CellRef`, `column_letter`, `sheet_part`, `occupied`, table
and defined-name writers, OPC relationship editing) belong to the document
surface. That is a breaking change to a published v5 API, so it needs
sequencing on the pyOpenVBA side rather than a silent removal, and it is not
done yet.

### 1.2 Host order

Excel is finished before Word starts: cells, ranges, formulas, formats,
tables, connections, Power Query and formula analysis. Then Word, then
PowerPoint, then Access. Not a thin slice across four hosts.

---

## 2. Layering

Each layer knows the layer below it and not the layer above.

```
+--------------------------------------------------------+
| excel/          the Excel surface                      |
|   workbook      sheets, shared parts, recalculation     |
|   worksheet     cells, rows, ranges, ordering rules     |
|   _values       the six cell encodings, serial dates    |
|   _tables       ListObjects: their own parts and wiring  |
|   _dimensions   widths, heights, hiding, frozen panes    |
|   _rowcol       inserting and deleting rows and columns,  |
|                 and moving everything that records a      |
|                 cell address                              |
|   _addresses    the five notations an address is spelled  |
|                 in, two of them zero-based                |
|   _tokens       a formula, broken into the pieces a        |
|                 transform may touch                        |
|   _names        defined names, and the naming rules       |
|                 tables share with them                    |
|   _validation   what a cell accepts; showDropDown is      |
|                 inverted and means hide it                |
|   _protection   sheet protection, whose flags are locks   |
|                 rather than permissions                   |
|   _pagesetup    margins in inches, headers as one string  |
|   _hyperlinks   links, whose address is in a relationship |
|   _filters      autofilters, and which rows they hide:    |
|                 Excel does not recompute one on open      |
|   _numfmt       what a number format shows, rounding and  |
|                 the fictional 1900 leap day included      |
|   _collate      Windows word sort, the order and the      |
|                 equality a filter compares text by        |
|   _shapes       shapes, and the grid a form control needs |
|   _comments     notes and threads, and the box each gets  |
|   _pictures     images, sized as Excel sizes them         |
|   _charts       what each chart plots, and chart sheets   |
|   _chartbuild   charts written as Excel's Insert Chart    |
|                 writes them                               |
|   _pivots       where each pivot table is, and what its   |
|                 cache reads                               |
|   _richtext     runs of text in several fonts, read as    |
|                 Excel shows them, not as marked up        |
|   _conditional  cfRules, and the compatibility formula    |
|                 that makes them fire                      |
|   _dxf          differential formats: what a rule paints  |
|   _styles       the five style tables; is this a date?   |
|   _formats      fonts, fills, borders, alignment, as     |
|                 immutable values                         |
|   _cellstyles   named styles, and Excel's definitions of |
|                 its own                                  |
|   _formulas     reference shifting for shared formulas, |
|                 and repointing a renamed sheet          |
|   _sharedstrings  the per-workbook string table         |
|   _xstring      text as SpreadsheetML spells it, with   |
|                 _xHHHH_ for what XML cannot carry       |
|   _reference    A1 notation, bijective base-26          |
|   _schema       the child order CT_Worksheet and         |
|                 CT_Workbook require                     |
+--------------------------------------------------------+
| word / powerpoint / access   to follow, in that order  |
+--------------------------------------------------------+
| opc.py          Open Packaging Conventions             |
|   - parts, cached as trees, flushed only when modified  |
|   - ContentTypes: Default by extension, Override by name|
|   - Relationships: per-part .rels, target resolution    |
+--------------------------------------------------------+
| _xml.py         XML that reproduces its own source     |
|   - hand-written parser, no stdlib XML module           |
|   - per-node source spans, dirty propagates to the root |
|   - prefixes, attribute order, empty-tag form preserved |
+--------------------------------------------------------+
| _zip.py         the ZIP container, field for field     |
|   - local and central headers kept apart                |
|   - untouched members copied as stored bytes            |
+--------------------------------------------------------+
| exceptions.py   the hierarchy every layer raises from  |
+--------------------------------------------------------+
```

### 2.1 Layer rules

- `_zip.py` knows nothing about XML or OOXML. It could be lifted out.
- `_xml.py` knows nothing about packages, parts or relationships.
- `opc.py` is the only module that touches the filesystem, apart from
  `Workbook.open`, which delegates to it. Adding a `pathlib` import to
  `_xml.py` or `_zip.py` is a red flag.
- A host surface must not reach past `opc.py` into `_zip.py`. If it needs
  something from the container, `opc.py` grows a method.
- Inside `excel/`, the four private modules are pure: `_reference`,
  `_styles`, `_formulas` and `_values` know about elements and strings, not
  about packages or files. `workbook.py` is the only one that does. So are
  `_numfmt`, `_collate` and `_filters`: the worksheet hands a filter its
  cells and the date, and the filter decides rows without touching it.
- `Worksheet` owns the part and does the work, addressed by `CellRef`.
  `Cell` and `Range` are views that delegate to it, so nothing is reachable
  only through a cell. That is why `Worksheet.get_value` and friends are
  public even though `Cell.value` is the pleasant way to call them.

---

## 3. Why the standard library was not enough

Both lower layers exist because a measurement said the obvious choice would
corrupt a document. Neither was a preference.

### 3.1 The XML layer

Excel writes a worksheet as a declaration, a CRLF, then one long line. Its
root element declares `mc:Ignorable="x14ac xr xr2 xr3"` *before* declaring
the `x14ac` and `xr` prefixes. Through `xml.etree.ElementTree` that becomes
`ns0`/`ns1`, the declarations reorder, the CRLF disappears, and
`standalone="yes"` is dropped. Every save then rewrites the whole part, so a
one-cell change is unreviewable and a no-op save is not one.

So every node keeps the source text it was cut from, and
`Element.to_xml()` copies that text unless the node is marked modified.
Marking a node modified walks up and marks its ancestors, because an
ancestor's span contains the child's; the ancestor then rebuilds, and its
*other* children still copy. Editing one cell rewrites that cell's row and
memcpies the rest of the sheet.

An element that is modified has its own start tag rebuilt with
single-space attribute separators, which is what Excel writes anyway. An
element that is not modified is byte-identical, odd spacing included.

### 3.2 The container layer

In a freshly authored workbook, `[Content_Types].xml` carries a 520-byte
extra field in its **local** ZIP header and none in the central directory.
`zipfile.ZipInfo.extra` exposes only the central copy, so anything built on
`zipfile`'s writer drops those bytes. The field is tag `0xa220`, the
Microsoft Open Packaging Growth Hint: a signature word, a padding word, then
reserved zero bytes Excel keeps so it can grow a part in place.

A member written by a streaming producer diverges the other way: with flag
bit 3 set, the local header zeroes the CRC and both sizes and only the
central directory and the trailing descriptor state the truth. Both copies
of all three fields are kept, which is why `local_crc` and friends exist and
are `None` when the two headers agree.

`zipfile`'s writer also re-deflates every member. Copying stored bytes
instead makes a no-op save provably a no-op and keeps a one-cell edit from
re-compressing a hundred megabytes of untouched parts.

### 3.3 Encodings

Office writes UTF-8 worksheets and UTF-16 custom XML. A workbook's
DataMashup part is UTF-16LE with a byte-order mark; its worksheets in the
same package are UTF-8. Both are detected, decoded and re-encoded with the
mark preserved. Anything outside the UTF-8 and UTF-16 families is refused by
name, those two being the encodings every XML processor must support.

---

## 4. Untrusted input

A document is input from outside, so the XML parser is deliberately less
capable than a conforming XML processor. None of these restricts anything
OOXML is allowed to contain.

| Construct | Behavior | Why |
|---|---|---|
| `<!DOCTYPE` | refused | no DTD means no external entity resolution (XXE) and no entity-expansion amplification (billion laughs) |
| `&name;` other than the five predefined | raises when decoded | resolving one would need a DTD |
| nesting past `MAX_DEPTH` (256) | refused | a hostile document must not exhaust the interpreter stack |
| zip64 | refused by name | reading the 32-bit fields anyway yields a plausible, wrong package |
| compression other than stored and deflate | refused by name | same |

Entity decoding is lazy: parsing a part with `&xxe;` in it succeeds, and
reading that text raises. That is deliberate, so a part carrying an unknown
entity somewhere harmless still round-trips.

---

## 5. Public API surface

Everything in `__all__` of [`src/pyofficeeditor/__init__.py`](../src/pyofficeeditor/__init__.py)
is supported. Everything else may change without notice, and a leading
underscore on a module name says so.

| Public name | Defined in | Purpose |
|---|---|---|
| `OpcPackage` | `opc.py` | an OPC package; context manager |
| `Relationship`, `Relationships` | `opc.py` | how Office navigates a package |
| `XmlDocument`, `Element` | `_xml.py` | the editable tree |
| `PyOfficeEditorError` and subclasses | `exceptions.py` | the error hierarchy |

The Excel surface is exported from `pyofficeeditor.excel`:

| Public name | Defined in | Purpose |
|---|---|---|
| `Workbook` | `excel/workbook.py` | a workbook; sheets, context manager |
| `Worksheet`, `Cell`, `Range` | `excel/worksheet.py` | the sheet and views onto it |
| `CellRef`, `RangeRef` | `excel/_reference.py` | A1 notation |
| `CellError` | `excel/_values.py` | an Excel error value, distinct from its text |
| `CellValue` | `excel/_values.py` | the union a cell can hold |
| `Styles`, `SharedStrings` | `excel/_styles.py`, `excel/_sharedstrings.py` | the shared parts |
| `column_letter`, `column_index` | `excel/_reference.py` | bijective base-26 |
| `MAX_ROW`, `MAX_COLUMN` | `excel/_reference.py` | Excel's real limits |
| `AutoFilter`, `FilterColumn`, `ValueFilter`, `CustomFilter`, `Comparison`, `FilterOperator`, `Top10Filter`, `DynamicFilter`, `DateGroup`, `OpaqueCriterion`, `criteria` | `excel/_filters.py` | a filter's criteria, validated as Excel validates them |
| `FilterOutcome` | `excel/_filters.py` | which rows applying a filter hid, showed and left alone |
| `Comment` | `excel/_comments.py` | a note on a cell: text, author, and whether it shows |
| `ThreadedComment`, `Reply` | `excel/_comments.py` | a conversation on a cell, and each answer in it |
| `TextRun` | `excel/_richtext.py` | part of a cell's text, and the font it shows in |
| `CellStyle` | `excel/_cellstyles.py` | a named cell style: its format, and what of a cell it sets |
| `Chart`, `ChartSeries` | `excel/_charts.py` | what a chart plots, and each series' references |
| `ChartSheet` | `excel/_charts.py` | a tab that is one chart and holds no cells |
| `PivotTable` | `excel/_pivots.py` | where a pivot table is, and what its cache was read from |
| `format_value` | `excel/_numfmt.py` | the text Excel shows for a value under a format code |

`opc.py` also exports the path helpers (`normalize_part_name`,
`rels_part_for`, `resolve_target`, `relative_target`) and the content-type
and relationship constants. They are stable but are not the recommended user
surface; a host facade should be.

---

## 6. Exception hierarchy

```
PyOfficeEditorError
    UnsupportedFormatError   a format or container feature we refuse
    ZipError                 malformed ZIP container
    XmlError                 malformed XML, or a construct we refuse
    PackageError             missing part, broken relationship, no content type
```

Rules:

- Every parser raises from this hierarchy, or `ValueError` for a
  caller-input bug. No bare `Exception`, and no `KeyError` reaches user code.
- A refusal names the part, member or element it is about. "Malformed
  package" with no name is not an acceptable message.
- `XmlError` carries a line, a column and an excerpt.

---

## 7. Conventions and house rules

- **No runtime dependencies.** The stdlib only. `pytest`, `pyright`, `ruff`
  and `openpyxl` are dev extras. `pyvbaharness` is deliberately *not* an
  extra: it sits in the PEP 735 `live` dependency group, which never reaches
  the built wheel's metadata. It drives a real Excel over COM, and a library
  whose first promise is no Office installation and no dependencies must not
  advertise it as something a consumer can install. Install it with
  `pip install -e ".[dev]" --group live`.
- **Python 3.10+.** `|` unions, `match` where natural, dataclasses for
  record types.
- **No silent corruption.** Any path that could leave a document
  inconsistent either succeeds completely or raises. Work happens on the
  in-memory package and the bytes are written once.
- **Preserve what you do not understand.** Growth hints, data descriptors,
  archive comments, unmodelled tails, comments and processing instructions
  are round-tripped verbatim.
- **Refuse rather than guess.** zip64, an unsupported compression method, a
  DOCTYPE, an unknown encoding: each is named and refused, because a
  plausible wrong answer is worse than an error.
- **ASCII only in user-facing strings.** Messages, warnings and CLI output:
  no emoji, no smart quotes, no en dashes.

---

## 8. Test layout

```
tests/
  conftest.py                   fixtures; openpyxl workbooks built in-run
  test_zip.py                   container round-trip, header fidelity,
                                malformed input
  test_xml.py                   source-preserving round-trip, scoped
                                rewriting, untrusted input
  test_opc.py                   path arithmetic, no-op save, relationships,
                                interop, and the SpreadsheetML format facts
  test_excel_reference.py       A1 notation, exhaustive at the boundaries
  test_excel_styles.py          number formats and date classification
  test_excel_sharedstrings.py   the string table, whitespace, rich text
  test_excel_workbook.py        the Excel surface end to end
  test_excel_sheets.py          adding, removing, renaming, reordering
  test_excel_formats.py         fonts, fills, borders, alignment
  test_excel_merges.py          merged ranges and range intersection
  test_excel_tables.py          ListObjects: parts, columns, wiring
  test_excel_dimensions.py      widths, heights, hiding, frozen panes
  test_excel_names.py           defined names and their scope indices
  test_excel_rowcol.py          the tokenizer, shifting, insertion
  test_excel_delete.py          #REF!, shrinking ranges, orphaned groups
  test_excel_conditional.py     cfRules, and what makes them fire
  test_excel_validation.py      what a cell accepts, and the dropdown
  test_excel_settings.py        protection, tab colour, view, visibility
  test_excel_pagesetup.py       margins, orientation, headers, print area
  test_excel_links.py           hyperlinks and outline grouping
  test_excel_filters.py         autofilters on sheets and tables, and the
                                rows they hide
  test_excel_filter_semantics.py  which rows each criterion keeps, case
                                by case against filter_semantics.json
  test_excel_numfmt.py          what a number format shows, against
                                number_formats.json
  test_excel_shapes.py          shapes, against what Excel said of them
  test_excel_controls.py        what a form control is wired to
  test_excel_shapes_write.py    adding, removing and rewiring shapes
  test_excel_comments.py        notes and threads, against what Excel said
  test_excel_pictures.py        pictures, against what Excel said of them
  test_excel_xstring.py         _xHHHH_ escapes, character by character,
                                against what Excel wrote and read
  test_excel_escapes.py         that text everywhere a workbook keeps it,
                                against escapes_answers.json
  test_excel_richtext.py        text in several fonts, against the fonts
                                Excel showed and the entries it wrote
  test_excel_cellstyles.py      named styles, against the cells Excel
                                styled itself
  test_excel_charts.py          charts and chart sheets, and every
                                reference after nine edits, against the
                                files Excel saved after the same edits
  test_excel_pivots.py          pivot tables through nineteen edits, and
                                the ones Excel refuses
  test_excel_chartbuild.py      charts added, against the parts Excel
                                wrote for the same charts
  test_excel_dxf.py             differential formats and the dxfs table
  test_excel_live_gate.py       real Excel, opt-in
  fixtures/excel/               twenty-one Excel-authored packages:
                                three sourced, eighteen scripted, and
                                two measured corpora; see its README
```

- **Always** run pytest with `-p no:randomly` to keep ordering reproducible.
- **Strict Pyright** must report 0 errors on `src` and `tests` before any
  merge: `pyright src tests`.
- **Ruff** must pass: `ruff check src tests scripts`.
- New behavior lands with its test in the same commit.
- The suite needs no Office installation. It runs against twenty-one committed
  Excel-authored packages, an `.xlsb` among them, and against
  openpyxl-authored ones generated during the run, because a reader that
  only ever sees one producer's output encodes that producer's habits as
  rules.
- Richer Excel-authored fixtures come from
  `scripts/build_excel_fixtures.py`, which drives real Excel through
  `pyvbaharness`. Tests needing them skip when they are absent. The two
  measured corpora come from `scripts/measure_number_formats.py` and
  `scripts/measure_filters.py` the same way.

### 8.1 The fidelity gate

Three assertions, and everything else rests on them:

1. Reading a package and writing it back unchanged reproduces the input
   bytes exactly.
2. Editing one part leaves every other part's stored bytes identical.
3. What this library writes, `zipfile` and openpyxl can read.

The first two hold for Excel-authored `.xlsm`, `.xlsb` and `.xlsx`, for
openpyxl-authored `.xlsx`, and for archives whose members carry data
descriptors.

The third is narrower, and the difference matters. `zipfile` reads every one
of them, because that half of the assertion is about the container and not
the format inside it. openpyxl reads `.xlsx` and `.xlsm` and has never
supported `.xlsb`, so it cannot stand as a reader for the binary workbook at
all; the test behind this assertion round-trips an openpyxl-authored `.xlsx`
through the XML layer and reads it back with openpyxl.

### 8.2 Format facts the Excel layer is built around

Each of these gives a plausible wrong answer rather than an error, so each is
pinned by a test that names the wrong answer.

**Shared formulas.** A formula assigned to a range is stored once, on the
first cell, as `<f t="shared" ref="D2:D5" si="0">B2*C2</f>`; the rest carry
`<f t="shared" si="0"/>` and a cached value, with no formula text. Reading
`<f>` per cell reports an empty formula for all but the master. The rest are
translated from it by `_formulas.shared_formula_for`. A formula entered into
one cell is stored plainly, so both shapes occur in the same sheet.

**Sheet order.** The order of worksheet relationships is not the order of
sheets, and Excel does reorder them in practice: in the sample fixture
`rId2 -> sheet2.xml` precedes `rId1 -> sheet1.xml`. Sheets are resolved
through `<sheets>` in `xl/workbook.xml`, whose entries carry the name, the
`sheetId`, and an `r:id` naming the relationship.

**Dates are numbers.** Only the number format distinguishes `46037` from a
date, through `s` to `cellXfs` to `numFmtId` to a format code. Classifying
the code means skipping quoted runs, backslash escapes and bracketed
sections, because `#,##0 "days"` is not a date and `[h]:mm` is.

**The 1900 leap-year bug is in the format.** Excel numbers dates as though
1900 were a leap year, reserving serial 60 for a 29 February that never
existed, so serials from 61 are one greater than the true day count. Serial
60 is refused rather than reported as 1 March.

**Cached results go stale.** A formula cell stores its last computed value,
so changing an input leaves a wrong number in the file. A modified workbook
is saved with `fullCalcOnLoad` set and `calcChain` dropped. The live gate
verifies Excel really does recalculate: after an edit it reports the
recomputed value and not the one still in the bytes.

**Ordering is not cosmetic.** Rows must be in ascending `r` order and a
row's cells in ascending column order. Excel refuses a worksheet that breaks
either rule rather than repairing it, so `_ensure_cell` and `_ensure_row`
insert in place instead of appending.

**Nor is element order.** `CT_Worksheet` and `CT_Workbook` are sequences, so
a missing element cannot simply be appended. A sheet Excel authored begins
with `sheetPr`, so a missing `dimension` belongs after it, not at the front;
a workbook usually ends with `extLst`, so a missing `calcPr` belongs before
it, not at the end. `_schema.py` holds both orders and
`insert_in_schema_order` consults them. Both mistakes were live in this
codebase before a test caught them.

**A sheet's name lives in four places.** Adding a sheet needs the part, its
content-type override, a relationship from the workbook part, and an entry in
`<sheets>`. Renaming one needs the entry, every formula that reads from the
sheet, and every defined name scoped to it. Renaming rewrites only the
formulas that carry text, which leaves a shared-formula group intact: writing
each follower its own text would destroy the group.

**`activeTab` is an index.** Removing or moving a sheet shifts it, so it is
adjusted to keep pointing at the sheet that was active. Leaving it past the
end makes Excel offer to repair the file.

**Formatting is shared and so is never mutated.** A cell carries an index
into `cellXfs`, and many cells carry the same one, so editing an entry
restyles every cell using it. `ensure_cell_format` finds a matching entry or
appends one, and leaves existing entries alone. The live gate checks this the
only way that means anything: it italicises `A1`, which shared its entry with
`B1`, and asks Excel whether `B1` stayed upright.

**A cell with no `s` uses `cellXfs[0]`, not nothing.** Entry 0 names the
workbook's default font. Resolving a missing `s` to an empty `CellFormat`
would make `with_font(bold=True)` produce a font with no name or size, and
Excel would render the default typeface instead of the workbook's. The live
gate asserts the typeface is unchanged after emboldening, which is the only
place that distinction is visible.

**A solid fill's colour lives in `fgColor`.** Not `bgColor`, despite the
names. Putting it in `bgColor` produces a cell that renders unfilled.

**Fill indices 0 and 1 are reserved** for `none` and `gray125`. Excel writes
both into every workbook whether or not anything uses them. A new fill is
appended from index 2, and the reserved pair is only created when the table
is empty: inserting them in front of existing entries would shift every
`fillId` in the workbook and repaint every cell.

**A merged range is two things.** An entry in `<mergeCells>`, and a block
where only the top-left cell holds a value. The covered cells are still
written, empty and carrying the anchor's style, because that is how a border
renders across the merge. Merging discards the covered values, which is what
Excel does: a covered cell is not displayed, so data left in one would be
invisible. Merges may not overlap, and Excel repairs such a worksheet rather
than rendering it, so an overlap is refused before anything is recorded.
Reading a covered cell gives `None`, so `Cell.merged_range` is how a caller
tells an empty cell from a covered one.

**A table is four wired-together pieces.** Its own part, a content-type
override, a relationship from the sheet, and a `<tablePart>` entry. Excel
ignores a table whose wiring is incomplete, or refuses the file, so none of
the four is optional. Three more details:

- `ref` spans the whole table *including* the totals row, while the
  `autoFilter` spans only the header and the data. Equal extents put a filter
  dropdown on the totals row.
- Whether a totals row exists is spelled two ways: `totalsRowCount="1"` when
  there is one, `totalsRowShown="0"` when there is not. Writing neither leaves
  Excel to guess.
- A column's name has to equal the text in its header cell. Excel reconciles
  the two on open by rewriting the part, so a mismatch looks like the
  library's edit being silently undone. `add_table` writes the resolved names
  back into the header cells for that reason, filling blanks with `Column1`
  and disambiguating duplicates the way Excel does.

Table names and ids are workbook-wide rather than per sheet, so both are
allocated from `Workbook` rather than from the sheet doing the adding. A
headerless table is not supported: Excel makes one by inserting a row above
the block, which shifts every row below it and every formula referring to
them, and that belongs with row insertion.

**A `<col>` entry covers a range, not a column.** `<col min="1" max="5"
width="10"/>` sets five columns at once, so giving column 3 its own width
means splitting that entry into up to three, copying every attribute to each
piece. Editing it in place would resize all five, which is the shared-state
trap cell formatting has in a different costume.

**A column width is not the number a person types.** VBA's `ColumnWidth = 18`
stores `18.6328125`; the offset was `+0.6328125` for every integer width
measured, and `8.43`, Excel's default, broke even that by landing on
`9.08984375`. The unit counts `0` glyphs in the default font plus padding, so
converting needs that font's maximum digit width in pixels, which is not in
the file. The stored number is exposed as-is, because a conversion right for
one font and quietly wrong for the rest is worse than none. Row heights carry
no such trap: they are points, and 24 stores as 24.

**A row height needs its companion flag, and a column needs a width.** An
`ht` without `customHeight="1"` is ignored, so the value looks like it never
took. A `width` is honoured with or without `customWidth="1"`, but a `<col>`
with no `width` at all is a column of width 0: an entry carrying only
`outlineLevel`, `collapsed`, `style` or nothing shows as hidden, and stays
hidden when its outline is expanded. So every entry this library makes for a
column at the standard width carries that width, the one Excel writes: the
sheet's `defaultColWidth` when it has one, and otherwise a width fixed by the
Normal font, measured per font. Grouping columns hid them outright until this
was measured.

**A filter is stored twice, and Excel trusts the second copy.** The criteria
live in `<autoFilter>`, and which rows they hide lives on each row's `hidden`
flag. Excel does not recompute a filter when the workbook opens: it lights
the dropdowns from the first and shows the rows the second allows. So setting
a filter here decides every row, and that needs Excel's own rules. A value
list compares the text a cell *shows*, which is why `_numfmt` renders number
formats as Excel does, and compares it by Windows word sort, which `_collate`
reproduces. Both are held to corpora measured from Excel,
`number_formats.json` and `filter_semantics.json`, because the specification
states neither. A row that hangs on something not modelled, a colour filter
or a formula whose cached result an edit has made stale, is left as it was
rather than guessed at, and `FilterOutcome` says which rows those were.

**A chart reads the workbook through formulas.** Everything a chart plots,
a series' name, categories and values, and a title linked to a cell, is a
sheet-qualified reference in a `<c:f>` of the chart's own part. A chart on
any sheet, or on a chart sheet, may read from any other, so inserting or
deleting rows on one sheet has to visit every chart part in the workbook,
and renaming a sheet likewise. Measured, the references move exactly as a
cell's formula does, down to `Data!#REF!` for one deleted outright. Excel's
object model is no witness for that last case: it reports the old address
until the file is reopened, and then refuses to report the series, so the
fixture records the references in the files Excel saved.

**A pivot table is two parts away from its cells.** Its location is in its
own part and the range its cache was read from is in the cache's, and
neither is in the sheet. Excel refuses an insertion or deletion that cuts
through a pivot table, and so does this library, rather than guess at a
layout Excel would not produce; one that takes all of it deletes it, and its
cache when nothing else reads from that, as Excel does.

**A named style is defined only once something uses it.** A workbook's
`cellStyles` names each style it has and points at the style's format in
`cellStyleXfs`, whose `apply...="0"` attributes say what the style leaves
out. Excel's own styles are not in a workbook until a cell uses one, when
Excel writes one out from a definition of its own, and so does this library,
from Excel's definitions as Excel wrote them. Those definitions name the
theme's typefaces rather than the Normal style's, so they follow a change
of theme and not a change of Normal. A cell in a style carries a copy of what
the style sets, and its own `apply...="1"` marks where it differs.

**Text is not stored as XML stores it.** XML cannot carry most control
characters, and an attribute turns a line break into a space. SpreadsheetML
spells such a character `_xHHHH_` instead, and escapes the underscore of any
literal text that would read as one, so `_x0041_` in a cell is stored
`_x005F_x0041_`. Excel writes uppercase hex in element text, where a line
break stays a line break, and lowercase hex in attributes, where it does not,
and reads either. Every place a workbook keeps text goes through
`_xstring`, because one that does not either writes a file Excel refuses or
reads back seven characters for one. A table's headers are the exception:
Excel replaces such a character there with U+FFFD rather than escaping it.

**Rich text is read by rules its markup does not state.** Excel writes the
first run of a cell's text with no font and gives the cell that run's font;
every later run carries its font in full. Other writers leave runs less
complete, and Excel reads them by rules measured here rather than found in
the specification. Runs with no font show in the cell's font until a run has
one, and in the workbook's default font after that. A run's font takes what
it leaves unsaid from the default font as well, so a run that says only
"bold" in a Courier cell shows in the default typeface, not in Courier.

**Inserting a row is not a local edit.** A cell's address is written into the
file in a dozen places, and missing one gives a workbook that opens cleanly
and points at the wrong cells, which no byte comparison catches. `_rowcol.py`
lists them and moves all of them: cell and row `r` attributes, every formula
in the *workbook* that reads from the sheet, shared-formula `ref`s, merges,
hyperlinks, the sheet and table filters, table extents, the dimension, page
breaks, and defined names.

Deciding which formulas move is why `_tokens.py` exists. A bare `A5` means
the sheet its formula lives on; `Data!A5` means that sheet. Inserting rows
into `Data` must move the second and leave the first, and only a token stream
that tracks the qualifier can tell them apart. A regex cannot, and neither
can it keep `LOG10` from looking like `G10`.

**`A:A` and `2:4` are references too.** A whole-axis reference is not a pair
of cells: `A:A` is not `A1:A1048576`, and Excel keeps the short form when it
rewrites a formula, so it gets its own token kind and its own value type
rather than being decomposed. Its ends move like a range's: inserting two
rows at 2 turns `2:4` into `4:6`, and deleting rows 2 to 4 turns it into
`#REF!` while `1:6` shrinks to `1:3`.

Leaving them alone was the original behaviour and it was the silent-
corruption case in miniature: after deleting rows 2 to 4, an untouched
`SUM(2:4)` sums three different rows and nothing reports an error. The
tokenizer tries a cell reference first, so `A1:A4` is still two cells, and
a lookaround keeps `TIME(2,4,0)`, `"2:4"` and `Table1[Units]` out.

**Nothing is refused.** There used to be a list of elements that made an
insertion raise rather than move everything else and leave them behind, which
was the honest answer while they were unmodelled. It is empty, and the
machinery is gone with it.

Two entries on it were wrong in opposite directions, which is the argument
for measuring rather than reasoning about a format. A background `<picture>`
was refused even though it tiles the sheet and names no cell at all.
Comments and legacy drawings were never on the list at all, so inserting a
row on a sheet with a comment moved the cells and left the comment behind.

`_addresses.py` holds the five notations a worksheet uses for an address:
`sqref` (space-separated ranges), `ref` (one range), `r` (one cell), a
drawing anchor's `<xdr:col>`/`<xdr:row>` pair, and a VML `<x:Anchor>`'s eight
numbers. The last two are **zero-based**, so treating them as one-based puts
every shape and comment one cell off and leaves the file perfectly valid.
`_rowcol.py` then drives most of the shifting from a table of
(container, entry, attribute, notation) rather than a function per feature.

**Some addresses are stored twice, and Excel checks.** A comment's cell is
in the comments part as a `ref` and again in its VML shape as
`<x:Row>`/`<x:Column>`, which is not the same thing as its `<x:Anchor>`.
Moving one and not the other does not produce a repair prompt: Excel refuses
to open the workbook. A modern data bar is the same shape of problem, with
its range in the `sqref` and again in an `x14` twin's `xm:sqref`.

**Conditional formatting is three things moving together.** The `sqref`, the
condition of a `cellIs` or `expression` rule (a real formula with real
references), and the compatibility formula of every other rule type. The
third is *rebuilt* rather than shifted: it names the top-left of the range by
construction, so it follows the block's new anchor. Move the range and leave
any one of the three and the rules highlight cells nobody asked about, with
no error anywhere.

**Deletion is not insertion run backwards.** It reuses the same inventory of
places a cell address is written, and adds three problems insertion does not
have.

The first is that a deleted reference does not move, it breaks. `A3` inside a
deleted row becomes `#REF!`. But `SUM(A1:A10)` over the same deletion becomes
`SUM(A1:A8)`: a range only *partly* deleted shrinks. So the two ends of a
range are decided together rather than one at a time, and only a range with
nothing left at all is replaced by `#REF!`. Deciding each end separately is
the plausible implementation and it is wrong in both directions, turning
survivable ranges into errors and breaking the ones that should shrink.

The second is that a shared formula lives once. `<f t="shared" si="0" ref=…>`
carries the text on the group's first cell and every other member has an
empty `<f si="0"/>` that derives from it. Delete that first cell and the rest
point at nothing, so any group about to lose its master is given its own text
first. The order matters more than it looks: a follower derives *from the
master's element*, so stripping the master's markers while still walking the
group leaves the remainder with nothing to derive from and they come out
blank. Every derived formula is read before anything is written. The live
gate caught this and the offline tests did not, because reading a formula
warms a cache that hid the bug; the regression test deletes without reading
anything first.

The third is that some things do not survive being shrunk. A merge reduced to
a single cell is not a merge, and Excel drops it. A table whose every column
is deleted is not a table. Both are removed rather than left behind as
degenerate. Deleting a table's *header* row is refused outright instead: the
column names have to equal the text in those cells, and Excel does not offer
the operation either.

**A defined name's scope is a position, not a name.** `localSheetId="0"`
means the first sheet in tab order, so adding, removing or moving a sheet
silently rescopes every name after it. `Workbook` resolves scope to a sheet
*name* on the way out and back to an index on the way in, and remaps every
index whenever the sheet order changes. A name scoped to a sheet that is
removed goes with it, which is what Excel does. Without that remapping the
corruption is invisible: the file opens, the name resolves, and it points at
the wrong sheet.

**A table name is a defined name**, so both follow one set of rules and one
check enforces them: no spaces, no operator characters, nothing that reads as
a cell reference, and not the `_xlnm.` prefix Excel reserves. Where the name
must be unique differs, though: a table name across the workbook, a defined
name only within its scope, so the same name may exist workbook-wide and on a
sheet at once, and the sheet-scoped one wins there.

**`activePane` decides where the cursor lands** after a freeze: `bottomRight`
when both axes are frozen, `bottomLeft` for rows only, `topRight` for columns
only. Get it wrong and the cursor sits in a pane the user cannot see. The
`<pane>` this library writes is byte-identical to Excel's for all four shapes.
Note that `ActiveWindow.SplitRow` and `SplitColumn` are *not* how to verify a
freeze: a workbook Excel froze itself reports 0 for both, so they discriminate
nothing. `Panes.Count` and where `VisibleRange` starts do.

### 8.3 The live gate

`test_excel_live_gate.py` is the only check that can prove Excel accepts what
this library writes. It drives real Excel through `pyvbaharness`, opens an
edited workbook and reads the cells back through Excel's own object model.
Opt in with `RUN_LIVE_EXCEL=1`; it needs Windows, Excel and the `live` extra.

**One session, shared by every live test.** The harness acquires its mutex
with a zero timeout, so creating a session per test races against the previous
one's teardown and fails intermittently with `SessionLockHeld`. Office
automation is sequential by contract, so the module holds one Excel instance
for all of its tests; it is also several times faster. If the lock is held by
unrelated work on the machine, the fixture skips rather than fails, since that
is not a defect in the code under test.

The harness creates and owns its own Excel process and kills it by recorded
pid on teardown, so an unrelated Excel on the machine is untouched. Confirming
that takes sampling the process list *during* a run: comparing before and
after shows nothing, because teardown has already killed the owned one.

If the lock is genuinely held, wait rather than passing `exclusive=False`,
which would contend with whatever is running. The mutex is abandoned-safe, so
a dead holder would have been granted to you already.

---

## 9. Files to keep up to date

When making the listed change, update **all** of these in the same commit.
This table is the canonical sync checklist.

| Change | Files to update |
|---|---|
| New public API symbol | `__init__.py` `__all__`, README, this file section 5 |
| New layer or module | this file sections 2 and 5, README architecture box |
| New refusal or safety limit | this file section 4, README "Untrusted input" |
| New exception type | `exceptions.py`, this file section 6 |
| New supported extension | the host facade, README table |
| New test file or fixture | this file section 8, `tests/fixtures/excel/README.md` |
| A new host surface | this file sections 2, 5 and 8, README architecture box and status |
| A measured format fact | the docstring of the module that acts on it, and this file if it changed a decision |
