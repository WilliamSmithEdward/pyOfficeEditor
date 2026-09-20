# Changelog

All notable changes to pyOfficeEditor are documented here. This project
follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

While the version is below 1.0 the public API may change between minor
releases. What will not change is the byte-fidelity contract: reading a
package and writing it back unchanged reproduces the input exactly.

<!-- Link definitions live above the sections, not below them. The release
workflow extracts one section by reading to the next "## [" heading, so
anything trailing the file is swept into the oldest release's notes. -->

[Unreleased]: https://github.com/WilliamSmithEdward/pyOfficeEditor/compare/v0.1.1...HEAD
[0.1.1]: https://github.com/WilliamSmithEdward/pyOfficeEditor/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/WilliamSmithEdward/pyOfficeEditor/releases/tag/v0.1.0

## [Unreleased]

Nothing yet.

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
