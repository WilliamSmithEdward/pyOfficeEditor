# Changelog

All notable changes to pyOfficeEditor are documented here. This project
follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

While the version is below 1.0 the public API may change between minor
releases. What will not change is the byte-fidelity contract: reading a
package and writing it back unchanged reproduces the input exactly.

<!-- Link definitions live above the sections, not below them. The release
workflow extracts one section by reading to the next "## [" heading, so
anything trailing the file is swept into the oldest release's notes. -->

[Unreleased]: https://github.com/WilliamSmithEdward/pyOfficeEditor/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/WilliamSmithEdward/pyOfficeEditor/compare/v0.1.1...v0.2.0
[0.1.1]: https://github.com/WilliamSmithEdward/pyOfficeEditor/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/WilliamSmithEdward/pyOfficeEditor/releases/tag/v0.1.0

## [Unreleased]

Nothing yet.

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
