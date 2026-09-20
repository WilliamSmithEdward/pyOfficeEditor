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
| excel / word / powerpoint / access   host surfaces     |
|   - not built yet; Excel is next                       |
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
- `opc.py` is the only module that touches the filesystem. Adding a
  `pathlib` import to `_xml.py` or `_zip.py` is a red flag.
- A host surface must not reach past `opc.py` into `_zip.py`. If it needs
  something from the container, `opc.py` grows a method.

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

- **No runtime dependencies.** The stdlib only. `pytest`, `pyright` and
  `openpyxl` are dev extras; `pyvbaharness` is a `live` extra.
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
  conftest.py        fixtures; openpyxl workbooks built during the run
  test_zip.py        container round-trip, header fidelity, malformed input
  test_xml.py        source-preserving round-trip, scoped rewriting, safety
  test_opc.py        path arithmetic, no-op save, relationships, interop,
                     and the SpreadsheetML format facts the Excel layer owes
  fixtures/excel/    three committed Excel-authored packages, plus two built
                     on demand by real Excel; see its README
```

- **Always** run pytest with `-p no:randomly` to keep ordering reproducible.
- **Strict Pyright** must report 0 errors on `src` and `tests` before any
  merge: `pyright src tests`.
- **Ruff** must pass: `ruff check src tests scripts`.
- New behavior lands with its test in the same commit.
- The suite needs no Office installation. It runs against two committed
  Excel-authored packages and against openpyxl-authored ones generated
  during the run, because a reader that only ever sees one producer's output
  encodes that producer's habits as rules.
- Richer Excel-authored fixtures come from
  `scripts/build_excel_fixtures.py`, which drives real Excel through
  `pyvbaharness`. Tests needing them skip when they are absent.

### 8.1 The fidelity gate

Three assertions, and everything else rests on them:

1. Reading a package and writing it back unchanged reproduces the input
   bytes exactly.
2. Editing one part leaves every other part's stored bytes identical.
3. What this library writes, `zipfile` and openpyxl can read.

They hold for Excel-authored `.xlsm`, `.xlsb` and `.xlsx`, for
openpyxl-authored `.xlsx`, and for archives whose members carry data
descriptors.

### 8.2 Format facts the Excel layer owes

Two gates in `test_opc.py` pin SpreadsheetML behavior that a cell-level API
gets wrong by default. They live there until the Excel module exists, then
move.

**Shared formulas.** A formula assigned to a range is stored once, on the
first cell, as `<f t="shared" ref="D2:D5" si="0">B2*C2</f>`; the rest carry
`<f t="shared" si="0"/>` and a cached value, with no formula text. Reading
`<f>` per cell reports an empty formula for all but the master. The rest
must be translated from it, shifting relative references. A formula entered
into one cell is stored plainly, so both shapes occur in the same sheet.

**Sheet order.** The order of worksheet relationships is not the order of
sheets, and Excel does reorder them in practice. Sheets are resolved through
`<sheets>` in `xl/workbook.xml`, whose entries carry the name, the
`sheetId`, and an `r:id` naming the relationship.

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
| A measured format fact | the docstring of the module that acts on it, and this file if it changed a decision |
