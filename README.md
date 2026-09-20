# pyOfficeEditor

Edit the document surface of Microsoft Office files in pure Python. No Office
installation, no COM, no dependencies.

Its sister project [pyOpenVBA](https://github.com/WilliamSmithEdward/pyOpenVBA)
edits the VBA project inside an Office file. This one edits the document: the
cell, the formula, the paragraph, the slide, the table, the query.

> **Status: early.** The container, XML and packaging layers every host sits on
> are built and tested against bytes Excel wrote. The Excel surface is being
> built on them now. Word, PowerPoint and Access follow, in that order.

## Why this exists

openpyxl, python-docx and python-pptx already set a cell, a paragraph and a
slide. This library is aimed at the ground they leave uncovered:

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
| excel / word / powerpoint / access   host surfaces     |
|   (being built, in that order)                         |
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

## Current API

```python
from pyofficeeditor import OpcPackage

with OpcPackage.open("book.xlsx") as package:
    # Navigate the way Office does: by relationship, not by path.
    workbook_part = package.main_document_part()          # 'xl/workbook.xml'
    sheets = package.relationships(workbook_part).by_type(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    )
    sheet = package.xml(sheets[0].target_part)

    sheet.root.require("dimension").set("ref", "A1:D20")
    package.save()          # every other part keeps its original bytes
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

The suite needs no Office installation. It tests against two committed
Excel-authored packages and against openpyxl-authored ones generated during the
run, because a reader that only ever sees one producer's output encodes that
producer's habits as rules.

Richer Excel-authored fixtures are built on demand, with real Excel, through
[pyVBAharness](https://github.com/WilliamSmithEdward/pyVBAharness):

```bash
python -m pip install -e ".[dev,live]"
python scripts/build_excel_fixtures.py
```

Tests needing those fixtures skip when they are absent.

## License

MIT. See [LICENSE.md](LICENSE.md).
