# Changelog

All notable changes to pyOfficeEditor are documented here. This project
follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

While the version is below 1.0 the public API may change between minor
releases. What will not change is the byte-fidelity contract: reading a
package and writing it back unchanged reproduces the input exactly.

<!-- Link definitions live above the sections, not below them. The release
workflow extracts one section by reading to the next "## [" heading, so
anything trailing the file is swept into the oldest release's notes. -->

[Unreleased]: https://github.com/WilliamSmithEdward/pyOfficeEditor/compare/v0.2.2...HEAD
[0.2.2]: https://github.com/WilliamSmithEdward/pyOfficeEditor/compare/v0.2.1...v0.2.2
[0.2.1]: https://github.com/WilliamSmithEdward/pyOfficeEditor/compare/v0.2.0...v0.2.1
[0.2.0]: https://github.com/WilliamSmithEdward/pyOfficeEditor/compare/v0.1.1...v0.2.0
[0.1.1]: https://github.com/WilliamSmithEdward/pyOfficeEditor/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/WilliamSmithEdward/pyOfficeEditor/releases/tag/v0.1.0

## [Unreleased]

### Added

- **Autofilters, on a sheet and on a table.** Every criterion Excel's
  object model sets is read and written, colour and icon aside: value
  lists and blanks, date groups from a year down to a second, one or two
  comparisons with wildcards, top and bottom N or N percent, above and
  below average, and the periods measured from today, from yesterday to
  year to date, with Q1 to Q4 and each month. `Worksheet.auto_filter`,
  `set_auto_filter`, `apply_auto_filter` and `clear_auto_filter` work on
  the sheet's filter; `Table.auto_filter`, `set_table_filter`,
  `apply_table_filter` and `clear_table_filter` on a table's.
  `criteria(">=10")` builds a criterion from Excel's own `Criteria1`
  spelling and stores it the way Excel does.

  Setting one hides the rows it excludes, because Excel does not
  recompute a filter when the workbook opens. The criteria live in
  `<autoFilter>` and which rows are out of view lives on each row, and
  Excel trusts the second: a file carrying criteria with every `hidden`
  flag stripped opens with the dropdowns lit and every row showing.

  So which rows a criterion keeps is worked out here, and held to 358
  cases in `filter_semantics.json`, each a criterion Excel applied to a
  column built to trip it. A value list matches the text a cell shows,
  its number format applied, and compares the way Windows sorts words:
  case does not count, and `ae` equals `æ`. `<` and `>` compare numbers
  with a number and text with text, and a blank or an error never passes
  either. `<>5` hides the number 5 and keeps the text "5". Once a list
  has a date group, a date is matched by the groups alone, even when its
  text is in the list.

  What Excel refuses to open, or opens and ignores, is refused when the
  criterion is built: no comparisons or three, an empty value, a top ten
  of 0 or 501, an unknown period. So is a sheet filter overlapping a
  table, and a top ten or an average over a column holding an error,
  which Excel's object model refuses as well.

  A filter by colour, by icon or in markup not modelled here reads as an
  `OpaqueCriterion` and is written back unchanged. `FilterOutcome` says
  which rows were hidden and which shown, and which were left as they
  were: those only such a criterion could decide, and those decided by
  a formula whose cached result an edit has made stale.

- **What a cell shows.** `Cell.text` is the value under its number
  format, Excel's `Range.Text` in a column wide enough to hold it, and
  `format_value(value, code)` does the same for any value and format
  code. It is held to 27,898 texts Excel rendered for 503 format codes,
  covering sections and conditions, fractions, exponents, elapsed time
  and both date systems, and to all 164 built-in format ids and the 782
  cells of the committed fixtures. Under General, `=0.1+0.2` shows
  `0.3`, where Python's `str` gives `0.30000000000000004`; a value
  filter needs the former.

- **Notes on cells**, which the file and Excel's object model call
  comments: `Worksheet.comments`, `comment`, `set_comment` and
  `remove_comment`, and `Cell.comment`, read and set. A note is three
  things that have to agree: its text in the comments part, a box in the
  VML part, and `<legacyDrawing>` on the sheet. All three are written the
  way Excel wrote them when measured, down to the Tahoma 9 run the text
  sits in, the CRLF it stores a line break as, and the box 144 by 79
  pixels placed 15 right of the cell and 10 above it. Excel reads notes
  written here back with their text, author and visibility intact, beside
  a form control as well as alone.

  A note and a form control on one sheet share its VML part and one
  sequence of shape ids, and each sheet's VML part claims a block of ids
  of its own. Both are measured, and both were wrong for controls before:
  every new VML part claimed the first block, and a control added after a
  note would have taken the note's id.

- **Threaded comments**, Excel's newer kind, which its Review tab calls
  comments: `Worksheet.threaded_comments`, `threaded_comment`,
  `add_threaded_comment`, `add_threaded_reply` and
  `resolve_threaded_comment`, with `ThreadedComment` and `Reply`, and
  `remove_comment` taking a thread off as it takes a note. A thread is
  four things: its entries in the sheet's threads part, their authors in
  the person list the workbook shares, the placeholder note Excel keeps
  beside it for versions that cannot show threads, and that note's box.
  Each is written as Excel wrote it when measured, the placeholder's
  wording included, and kept up to date as replies are added.

  Excel lists replies written at one moment in the order of their ids,
  measured, so a reply here takes the id after the thread's last, as
  Excel's own do; a random one put the second reply first.

- **Pictures**: `Worksheet.add_picture` for a PNG, a JPEG or a GIF, from
  bytes or a path, `picture_data` for the bytes a picture shows, and
  `Shape.image` naming its media part. Written as Excel's
  `Shapes.AddPicture` writes one, and sized as it sizes one, measured:
  the image's pixels at 96 to the inch, a PNG's own density where it
  gives one, and a JPEG's ignored even where it gives one. An image shown
  twice is stored once, as Excel stores it, and removing the last picture
  of an image removes the image.

- **Text in more than one font in a cell**: `Worksheet.get_rich_text`
  and `set_rich_text`, and `Cell.rich_text`, with `TextRun` for a run of
  text and its font. Written as Excel writes it, measured: the cell takes
  the first run's font and that run carries none of its own, every later
  run carries its font written out in full, and a line break is stored as
  CRLF. For seven cells Excel made through `Characters`, bold, colour,
  size, underline, superscript, another typeface and a line break among
  them, the entries written here are Excel's byte for byte.

  Runs are read the way Excel shows them, which the markup does not say.
  Runs with no font of their own show in the cell's font until a run has
  one. After that, a run with no font shows in the workbook's default
  font, not the cell's. A run's font takes whatever it leaves unsaid from
  that default font too, and an unsaid colour is automatic. Excel never
  writes such runs itself, so the rule was measured on runs written by
  hand, and a live gate has Excel check all twelve cases against this
  library's reading.

- **What a form control is wired to.** `Shape.control` carries the linked
  cell, the list range, the current value and which kind of control it
  is, read from the part the sheet points at. `Shape` and `ShapeKind`
  are exported now too; they never were, so a consumer got the objects
  but could not name the type.

  The value is worth the measuring. It is spelled three different ways
  and none of them is a plain `val`: a check box and an option button
  write `checked="Checked"` and write nothing at all when off, a drop
  down and a list box write `sel` and leave `val` at 0, and only a
  spinner and a scroll bar write `val`. A reader that takes `val` alone
  answers 0 for a ticked box and 0 for a drop down with the second item
  chosen. Both look right. Excel's own answer for an unticked box is
  -4146 rather than 0.

- **`Worksheet.add_shape`** for an AutoShape, a text box or a line, and
  **`Worksheet.add_form_control`** for all nine Forms controls: Button,
  CheckBox, Drop, List, Radio, Spin, Scroll, GBox and Label. A sheet
  with no drawing part gets one, with its content type and relationship;
  a control gets four parts that have to agree.

- **`Worksheet.remove_shape`**, which takes a control's four parts with
  it rather than leaving a relationship pointing at a part that is gone.

- **`Worksheet.set_shape_macro`**, pointing a shape at a procedure or
  clearing it. A drawing shape keeps a bare name on its own element; a
  control keeps `[0]!Name` on the sheet and a second copy in the VML.
  Measured: Excel reads the sheet's and ignores the VML's, and a file
  where the two disagree opens cleanly. Both are written anyway.

  The bracketed number indexes the workbook holding the procedure and is
  not always 0. An existing one is kept rather than replaced.

### Fixed

- **Clearing the last cell in a row dropped the row's own settings.** The
  row element went with the cell, and with it the row's hidden flag,
  height, style and outline level, so a row a filter had hidden came back
  into view. The element stays now while it says anything.

- **Inserting or deleting columns left a filter's criteria on the wrong
  columns.** A criterion names its column by offset from the filter's
  left edge, and the edge moved while the offsets stayed. Each is
  renumbered now, in a sheet's filter, a table's and a custom view's, and
  one whose column is deleted goes with it. A sheet's or a table's filter
  is then applied again, as Excel does, so the rows only that criterion
  hid come back.

- **Grouping columns hid them.** A `<col>` entry without a `width` is a
  column of width 0, measured: Excel shows it hidden and no unfolding
  brings it back. `group_columns` wrote exactly that, and restoring a
  column's standard width or ungrouping it could leave one behind. Every
  entry made for a column at the standard width now carries it, as Excel
  writes it: the sheet's `defaultColWidth`, or the width the Normal font
  fixes, measured for eighteen fonts.

- **A folded group did not say so.** Excel marks the row or column that
  summarises a folded group `collapsed`, and records how deep the outline
  goes on the sheet; `group_rows(collapsed=True)` and its column twin
  hid the rows and wrote neither, so the outline's plus button did not
  show. Both are written now, the depth kept up to date by ungrouping,
  and an eighth level is refused, as Excel refuses it.

- **A column inserted inside a table made Excel refuse the workbook.** The
  table grew while its list of columns did not. It now takes the new
  column as Excel does, named `ColumnN` with the smallest `N` free, and
  writes the name into the header cell. A blank header is named the same
  way: `add_table` over `["", "Amount", ""]` gives `Column1` and
  `Column2`, where it gave `Column1` and `Column3`.

- **A new part's relationships were lost when the part was written
  again.** Writing a part dropped its relationships from the cache, and
  the next lookup made an empty `.rels` over the one not yet saved.
  Nothing written until now added a relationship to a part it had just
  made and then wrote again; the second picture on a new drawing would
  have been the first thing to.

- **A form control made Excel refuse a sheet with a table.** The schema
  order this library inserts elements by had no place for
  `<legacyDrawing>`, so it was appended after `<tableParts>` or `<extLst>`
  where a sheet had them, which Excel refuses rather than repairs. A
  drawing added after it landed after it too, in the wrong order again.

- **Deleting a row or column left a deleted cell's note half there.** Its
  text went from the comments part and its box stayed in the VML part.
  Both go now.

- **Threaded comments stayed put when rows or columns moved.** Their
  placeholder notes moved and the threads did not, leaving each thread on
  a cell whose note had gone. Threads now move with their cells, and go
  with a deleted one.

- **Text holding a control character made Excel refuse the workbook.**
  XML cannot carry most characters below U+0020, and they were written
  as they were: a cell set to `"a\x01b"` gave a file Excel would not
  open. SpreadsheetML spells such a character `_xHHHH_`, and text is now
  written the way Excel writes it, measured for every one of them, in
  cells, formulas, notes, threads, headers and footers, validations,
  hyperlinks, tables, sheet names, defined names, filters, conditional
  formats and number formats. A literal `_x0041_` is written with its
  underscore escaped, as Excel writes it, so it reads back as itself.

- **Text Excel had escaped read back escaped.** A carriage return Excel
  stores as `_x000D_`, which a cell holding text from Windows line ends
  has before every line break, read as those seven characters, and so
  did every other escape, `_x005F_` in front of a literal one included.
  Each is read as the character it spells now, in either case of hex
  digit, as Excel reads them.

- **A line break in a validation's message, a hyperlink's tip or a
  table's column name became a space.** An attribute cannot keep a line
  break written as it is, and Excel reads one as a space, measured. It
  is written as Excel writes it now, `_x000a_`, and a raw one in a file
  another program wrote reads as the space Excel shows.

- **A table's header holding a control character stayed as it was.**
  Excel puts U+FFFD in its place, in the column's name and in its cell,
  when it makes a table, measured for each such character. `add_table`
  does the same.

- **`Worksheet.rename` did not rename the sheet.** It changed the name
  the object reported and nothing in the file, so the two disagreed and
  every formula still named the old sheet. It does what
  `Workbook.rename_sheet` does now.

- **Plain text that matched text in several fonts took on its fonts.**
  Setting a cell to `"bold plain"` pointed it at an existing entry with
  the same words in two fonts. Plain text gets a plain entry now.

- **A line break in text of more than one font read with a stray
  carriage return.** Excel stores the break as CRLF inside a run, and an
  XML reader reads that as one line feed; this library kept both, so a
  cell Excel shows as `two` over `lines` read as `"two\r\nlines"`. Line
  ends in text are read as XML reads them now.

- **A cell formatted with `e` read as a number.** `e` is the era year,
  which Excel in en-US shows as the year: 0 under it shows 1900. The
  cell layer now reads such a cell as a date, and a test holds its idea
  of which codes are dates to the renderer's across every measured code.

### Verified

Four rules cost a refused workbook each to find, since Excel declines to
open a package it disagrees with rather than repairing it:

- a control's anchor writes `<from><xdr:col>`: the wrapper loses the
  prefix and its children keep it
- a connector is `prst="line"`, not the `rect` every other shape gets
- the VML spells a tick box `Checkbox` where its own part spells it
  `CheckBox`
- a sheet that has never held a control declares neither `xdr` nor
  `x14`, and a control's markup needs both

Two more are measured rather than assumed. A spinner or scroll bar with
no maximum is pinned at zero whatever value it holds, in a file that is
perfectly valid, so Excel's own defaults are written instead: 30000 and
100. And a control with a linked cell takes its state from that cell on
load, so a tick box stored ticked and linked to an empty cell opens
unticked. Setting both is setting the cell.

`controls.xlsm` is a new Excel-authored fixture carrying one of every
control, wired up, with `controls_answers.json` recording what Excel's
object model answered for each. The committed `shapes.xlsm` carries a
Button and nothing else, so it could prove none of this.

`filters.xlsx` is another, one criterion kind per sheet, with
`filters_answers.json` recording how many rows Excel hid on each, and
`comments.xlsx` a third, with notes of every shape and
`comments_answers.json` recording what Excel's object model said of each,
and `pictures.xlsx` a fourth, with `pictures_answers.json`, and
`richtext.xlsx` a fifth, with `richtext_answers.json` recording the font
Excel reported at every change in its text, and `escapes.xlsx` a sixth,
with text XML cannot carry as it is everywhere a workbook keeps text and
`escapes_answers.json` recording each character Excel read back.
`filter_semantics.json` and `number_formats.json` are new measured
corpora, rebuilt by `scripts/measure_filters.py` and
`scripts/measure_number_formats.py` on a machine with Excel; every case in
them is a test.

### Internal

- Four fixture recipes never saved. `refused`, `settings`, `links` and
  `geometry` ended without a `SaveAs`, so rebuilding any of them stopped
  at "Excel reported success but the file is not there". They save now;
  run into a scratch folder, all four write a workbook that opens. The
  committed files were left alone.

- The two builders of measured fixtures were one function written
  twice. There is one now, handed the parser for each fixture's reply.

- Text with a line break between words is written without
  `xml:space="preserve"`, as Excel writes it; only white space at either
  end needs it. A line break is written as CRLF everywhere text is, as
  Excel writes it.

- An attribute value set with a tab, a line feed or a carriage return is
  written as a character reference, the one spelling an XML reader keeps,
  and one read as it was written reads as the space XML makes it.

- Writing a sheet top to bottom was quadratic: each new row searched
  every row for its place. A row past the last is appended now, and
  20,000 rows of three cells take 0.7 seconds where they took 3.6.

- The kind of control being added is checked before any of its four
  parts is written. The VML is what knows an unknown kind is wrong, and
  noticing it there left behind a control part, a relationship and an
  anchor for a control that was never made.

- The reader and the writer share one vocabulary, Excel's own
  `objectType`, so a control read from a file writes straight back. A
  test pins all nine kinds through that round trip. Keying the writer on
  a second set of names is how a control silently becomes a Button in a
  file that opens cleanly, which is
  [pyOpenVBA#25](https://github.com/WilliamSmithEdward/pyOpenVBA/issues/25).

### Known limits

A shape's placement is close rather than exact, and the error grows with
how far across the sheet it sits. Excel reports a shape's position from
its anchor, and turning points back into a column needs the standard
font's maximum digit width, which the file does not carry. A shape put
at 300 points came back at 300 on one sheet and 298.5 on another whose
columns had been resized. Rows are exact while Excel's default row height
is the one the file records. At another display scaling a row with no
height of its own is taller or shorter, 15 points at 100% against the
14.5 a file written at 150% records, and a shape below it moves by the
difference.

An option button group stores its linked cell once, on the button marked
`firstButton`. Excel's object model resolves the group and answers that
cell for every member; `linked_cell` reports what the part says, which is
empty for all but the first, because what makes a group is not measured
here.

The text a cell shows is Excel's in en-US, the locale every corpus here
was measured in. A value list stores the text of the machine that set
it, so a filter set in another locale names values as that locale showed
them. Applied here, it decides its rows by the en-US text, and a value it
cannot find there hides the rows holding it.

A table without a header row cannot take a filter here. Excel turns the
header row back on to filter one, which inserts a row, and this library
does not.

A note's text is read and written as plain text. Formatting inside a note,
a bold word say, reads as its text without the bold, and a note written
here is in Excel's own note font.

A run of text whose font names a typeface and the theme's font scheme as
well shows in the theme's typeface, measured, and reads here as the
typeface it names. Excel never writes one.

The author of a threaded comment written here is a person no account
stands behind, as Excel writes someone who is not signed in. Excel names
the signed-in account in its own, which is also why no Excel-authored
thread is committed as a fixture: the tests hold the reader to Excel's
measured markup with the person replaced.

## [0.2.2] - 2026-09-20

### Fixed

- **`__version__` reported 0.1.0 on every release since.** It was a literal
  in `__init__.py` that nothing linked to the version in
  `pyproject.toml`, so 0.1.1, 0.2.0 and 0.2.1 all imported saying 0.1.0.
  The distribution metadata was right throughout; only the attribute was
  wrong, so `pip` and the PyPI page were never affected and anything
  reading `pyofficeeditor.__version__` was.

  It is read from the installed distribution now, which cannot drift, and
  `tests/test_version.py` fails if the attribute, the metadata and
  `pyproject.toml` ever disagree. That test is the thing that was missing:
  nothing checked, so nothing failed.

## [0.2.1] - 2026-09-19

### Added

- **The preset geometry each MsoAutoShapeType maps to**, and
  `Shape.auto_shape_type` reporting it back, measured by having Excel make
  one shape of each of 33 types and reading the drawing.

  This is an addition rather than a fix: pyOfficeEditor never carried the
  table. It is worth measuring rather than transcribing, because every
  plausible wrong answer is a real preset name belonging to some other
  shape. A table pairing 11 with `cross`, 12 with `star5`, 16 with `can`
  and 17 with `cube` looks right and is wrong on all four: Excel writes
  `plus`, `pentagon`, `foldedCorner` and `smileyFace`. Nothing complains,
  the file is valid, and the wrong shape appears. The five-pointed star is
  92, not 12.

## [0.2.0] - 2026-09-19

Conditional formatting, data validation, sheet protection, page setup,
hyperlinks, grouping and shapes. Nothing about a sheet makes inserting or
deleting rows refuse any more.

### Added

- **Conditional formatting**, every rule family Excel has: comparisons,
  expressions, the four text rules, blanks, errors, duplicates and uniques,
  top and bottom, above and below average, the ten time periods, colour
  scales, data bars and icon sets. With the differential formats they paint
  with.

- **Data validation**: lists typed in or pointed at a range, whole numbers,
  decimals, dates, times, text length and custom formulas, with error and
  input messages and all three alert styles.

- **Sheet protection**, with or without a password, and the twelve
  allowances a protected sheet can still grant.

- **How a sheet prints**: margins, orientation, paper, scaling, print
  options, headers and footers, the print area and the rows and columns
  repeated on every page.

- **Sheet appearance**: tab colour, gridlines, row and column headings,
  zoom, and the three visibility states.

- **Hyperlinks**, internal and external, and **outline grouping** on rows
  and columns.

- **Reading the shapes on a sheet**: AutoShapes, text boxes, lines, groups
  and form controls, with their geometry, text and the macro a click runs.

### Fixed

- **A comment was left behind when rows moved.** Inserting or deleting rows
  on a sheet carrying a comment moved the cells and left the comment on the
  cell it used to be on. Comments and legacy drawings were on neither the
  shifted list nor the refused one, so nothing caught it. Present in 0.1.0
  and 0.1.1.

- **Whole-axis references did not move.** `SUM(2:4)` survived a deletion of
  rows 2 to 4 unchanged, going on to sum three different rows, and an
  insertion left it alone where Excel makes it `SUM(4:6)`. Present in 0.1.0
  and 0.1.1.

- **`Typing :: Typed` without a `py.typed`**, so downstream type checkers
  ignored the annotations. Fixed in 0.1.1.

### Changed

- **Nothing about a sheet's contents makes an insertion or deletion
  refuse.** The list of elements that used to raise is empty and the
  machinery is gone. Data validation, protected ranges, ignored errors,
  saved sorts, data consolidations, scenarios, custom sheet views, shapes,
  form controls, embedded objects, comments and extension content all move
  now, along with the addresses they keep in drawing anchors, VML and the
  comments part.

  Two entries on that list were wrong in opposite directions: a background
  picture was refused although it names no cell at all, and comments were
  never refused although they needed to be.

- `pyvbaharness` is no longer a published extra. It drove real Excel over
  COM and had no business being installable from a library whose first
  promise is no Office installation; it is a PEP 735 dependency group now,
  which never reaches the wheel's metadata.

## [0.1.1] - 2026-09-19

Documentation only; the library is byte for byte the code 0.1.0 shipped.
It gets a release because the README is the project page on PyPI, where a
wrong claim stays visible until a new version replaces it.

### Fixed

- **What openpyxl covers.** The README named three libraries and three
  things they do in one sentence, leaving the reader to pair them up. Read
  the other way it says openpyxl sets a paragraph and a slide, which it
  does not: openpyxl is Excel only, python-docx Word only, python-pptx
  PowerPoint only. Each is now paired with its own host explicitly.

- **openpyxl and `.xlsb`.** The architecture doc listed three fidelity
  assertions and said they hold for `.xlsm`, `.xlsb` and `.xlsx`. The third
  is that `zipfile` and openpyxl can read what this library writes, and
  openpyxl has never supported `.xlsb`. `zipfile` does read all three,
  since its half of the assertion is about the container rather than the
  format inside it, and the test behind the openpyxl half round-trips an
  `.xlsx`. Each half is scoped to what it actually covers now.

- **A stale count.** Both documents said the suite runs against two
  committed Excel-authored packages. There are six.

## [0.1.0] - 2026-09-19

First release. Excel's document surface, in pure Python with no
dependencies, verified against real Excel rather than against another
library's idea of the format.

### Added

- **The container, XML and packaging layers.** A hand-written ZIP reader
  and writer, because `zipfile` drops the 520-byte local extra field Excel
  uses as a growth hint. A hand-written XML parser, because `xml.etree`
  destroys `mc:Ignorable` and prefix ordering. Every node keeps its own
  source text and only a modified one is reserialized, so an untouched part
  comes back byte for byte. OPC on top: parts, content types, and
  relationships resolved the way Office resolves them.

- **Cells.** Values, formulas, dates and errors across the six cell
  encodings. A date is a number plus a number format, so reading one as a
  date means reading the style table; serial 60 is the leap day 1900 never
  had and is refused rather than reported as 1 March.

- **Sheets.** Adding, removing, renaming and reordering, with formulas and
  defined names following a rename.

- **Formatting.** Fonts, fills, borders, alignment and protection as
  immutable values, sharing entries in the style tables the way Excel does.

- **Structures.** Merged ranges, tables (ListObjects) with their own parts
  and four-way wiring, column widths, row heights, hiding, frozen panes,
  and defined names at both scopes.

- **Inserting and deleting rows and columns**, moving every reference in
  the workbook that records a cell address: cell and row attributes,
  formulas on this sheet and on others, shared-formula groups, merges,
  hyperlinks, filters, table extents, the dimension, page breaks and
  defined names. A deleted reference becomes `#REF!` while a range only
  partly deleted shrinks, matching Excel on both. Whole-axis references
  like `A:A` and `2:4` move too.

  Anything on the sheet that addresses cells and is not yet modelled makes
  the operation refuse and names what it found, rather than moving
  everything else and leaving a workbook that opens cleanly and points at
  the wrong cells.

### Verified

- 1093 offline tests, and 14 live gates that open what this library wrote
  in real Excel and read the result back through Excel's own object model.
  Twenty-three format behaviours were measured against Excel rather than
  assumed, including two that intuition gets wrong: a range shrunk to a
  single cell keeps its range shape, `SUM(A1:A1)` and not `SUM(A1)`, and a
  broken cross-sheet reference keeps its qualifier, `SUM(Data!#REF!)`.

- Byte fidelity holds for Excel-authored `.xlsm`, `.xlsb` and `.xlsx`, for
  openpyxl-authored `.xlsx`, and for archives whose members carry data
  descriptors.

### Not here yet

Word, PowerPoint and Access. Conditional formatting, data validation and
the other elements on the refusal list. The formula analyzer and linter,
legacy data connections, and Power Query.
