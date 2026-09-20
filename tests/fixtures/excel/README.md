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

## Authored by Excel, built on demand

`scripts/build_excel_fixtures.py` drives real Excel through
[pyvbaharness](https://github.com/WilliamSmithEdward/pyVBAharness) to author
three more. Run it on a Windows machine with Excel; tests that need them skip
when they are absent.

| File | What it carries |
|------|-----------------|
| `empty.xlsx` | The smallest thing Excel will save as `.xlsx`. |
| `sample.xlsx` | Shared strings, formulas, dates, booleans, an error cell, a merged range, escaped and space-padded text, two sheets. |
| `structures.xlsx` | Two ListObjects, one with a totals row and a calculated column; defined names at workbook and sheet scope; column widths; a hyperlink with its own external relationship. |

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
