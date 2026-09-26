# Changelog

All notable changes to pyOfficeEditor are documented here. This project
follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

While the version is below 1.0 the public API may change between minor
releases. What will not change is the byte-fidelity contract: reading a
package and writing it back unchanged reproduces the input exactly.

<!-- Link definitions live above the sections, not below them. The release
workflow extracts one section by reading to the next "## [" heading, so
anything trailing the file is swept into the oldest release's notes. -->

[Unreleased]: https://github.com/WilliamSmithEdward/pyOfficeEditor/compare/v0.3.0...HEAD
[0.3.0]: https://github.com/WilliamSmithEdward/pyOfficeEditor/compare/v0.2.2...v0.3.0
[0.2.2]: https://github.com/WilliamSmithEdward/pyOfficeEditor/compare/v0.2.1...v0.2.2
[0.2.1]: https://github.com/WilliamSmithEdward/pyOfficeEditor/compare/v0.2.0...v0.2.1
[0.2.0]: https://github.com/WilliamSmithEdward/pyOfficeEditor/compare/v0.1.1...v0.2.0
[0.1.1]: https://github.com/WilliamSmithEdward/pyOfficeEditor/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/WilliamSmithEdward/pyOfficeEditor/releases/tag/v0.1.0

## [Unreleased]

### Added

- **Changing a shape in place.** `Worksheet.update_shape` moves, resizes
  and renames a shape, changes its text, sets its alt text and whether it
  is hidden, and rewires a Forms control's linked cell and list range,
  changing all four of a control's parts together. New text keeps the font, size,
  colour and alignment of the text it replaces, in the drawing and in a
  control's VML. A side left out keeps where the anchor has it, which is
  where Excel draws the shape. `Worksheet.cell_origin` gives a cell's
  top-left corner in points, to put a shape on that cell.
  [#4](https://github.com/WilliamSmithEdward/pyOfficeEditor/issues/4)

- **More of what a shape is.** `Shape.cells` is the range its anchor
  covers, `Shape.alt_text` its alternative text and `Shape.hidden` whether
  it shows. A Forms control's comes from its VML, because Excel marks every
  control's drawing twin hidden, shown or not. An ActiveX control reads as
  kind `"activeX"`, one with no drawing twin included, and `remove_shape`,
  `update_shape` and `set_shape_macro` refuse it rather than take its
  binary part apart. An embedded OLE object reads as kind `"oleObject"`,
  `Shape.Type` 7, and `remove_shape` and `update_shape` refuse it too. A
  Forms control inside a group now reads as a control.

- **ODDFPRICE, ODDFYIELD, ODDLPRICE and ODDLYIELD**, the prices and yields
  of bonds with an odd first or last period, measured against Excel on
  1,500 random bonds of each kind and a few hundred designed ones.
  Microsoft's formulas hold, with Excel's own quasi-coupon dates and day
  counts: each date a period after the one before, so one clipped to the
  28th stays there; on basis 0 an odd last period's lengths count every
  month end as the 30th; a first coupon has to be one of the maturity's
  coupon dates, and an odd first period exactly one period long is
  refused; a long one to a month-end first coupon is discounted a period
  more, unless settlement falls on a month end too or shares its month
  with a quasi-coupon date.
  ODDLPRICE and ODDLYIELD give Excel's bits on every bond, ODDFPRICE on
  all but four a unit in the last place away. ODDFYIELD does on all but
  216 bonds with a long odd first period, and comes within 1e-10 of
  Excel's answer on those.

- **GROUPBY and PIVOTBY**, measured against Excel on 437 formulas, all of
  which match but one Excel will not take: grouping by one field or
  several, totals and subtotals above or below, keys sorted either way or
  groups sorted by a result, filters, headers given or left to Excel, and
  several functions side by side or stacked. PERCENTOF, or a LAMBDA of two
  parameters, gets every row's values as its second argument, or in
  PIVOTBY the rows relative_to names: the cell's column, its row, all of
  them, or its parent column or row. Excel's own ways are kept. Keys match
  as SORT matches them, and each shows the spelling of its first leaf's
  first row. PIVOTBY sorts rows by a result only beside a total column,
  and columns only where a group has a column of its own. One function
  over several value columns puts PIVOTBY's column totals in the columns
  Excel puts them in, which are the wrong ones, by the rule measured for
  where each lands.

- **Formulas that read another workbook calculate.** They used to keep
  the values Excel cached. Excel keeps a copy of every linked cell a
  formula read, in the link's own part, and calculates with it while the
  other workbook is closed; that copy is now read the same way, names and
  ranges across sheets included, a cell it lacks being blank. Measured in
  Excel with the linked workbook closed: SUMIF, COUNTIF and the rest of
  their family, the database functions and OFFSET are `#VALUE!` over its
  ranges, INDIRECT naming it is `#REF!`, and CELL, SHEET, SHEETS and
  ISFORMULA are `#N/A`, while RANK, SUBTOTAL, COUNTBLANK, ROW and the
  others work as ever. A workbook the file keeps no link to still leaves
  the formula its cached value.

- **What-if data tables calculate.** A table Excel's Data Table makes, in
  one variable down a column or along a row or in two, used to keep the
  values Excel cached. Each cell is now calculated as Excel calculates
  it: the value tried is put in the input cell, and the table's formula
  calculated again with everything that reads that cell, whatever sheet
  it is on; the rest of the model is calculated once. Held to a fixture
  Excel authored, where changing the model's inputs and recalculating
  gives, on all 64 cells, what Excel gave for those inputs. A table whose
  input cell was deleted keeps its values.

- **GETPIVOTDATA**, which reads a value from a pivot table's report, held
  to Excel by 147 formulas in a fixture Excel authored, over pivot tables
  in each layout it offers. The report is read from the pivot table's
  definition, row by row and column by column, and the value is the cell
  Excel wrote there. Measured: a value answers to its own name or the
  field it summarizes, a field to the name the report gives it, and an
  item to its given name or the text it is shown as, or, as a number, to
  its value. A field left out means its total, where the report shows
  one; failing that, the only line at the items named. A filter can be
  named only as the item it is set to. An empty cell is 0 and a cell the
  report does not show is `#REF!`. Dates grouped by years or months take
  a number's whole part, and bins take a number at or below their start.

- **BAHTTEXT**, an amount in Thai words, measured against Excel on 284
  amounts and matching every one. A 1 in the ones reads et after any
  other digit, those before a million included; satang are rounded as
  ROUND rounds, from the fifteen significant digits the words also
  spell; and a negative amount that rounds to nothing keeps its minus.

- **A function passed by name, and optional LAMBDA parameters.** A file
  writes `=BYROW(A1:B9,SUM)` as `_xleta.SUM`, and a LAMBDA's `[y]` as
  `_xlop.y`; both now calculate, ISOMITTED telling an optional parameter
  left out.

### Changed

- **An inserted row is formatted like the row above it, and an inserted
  column like the column to its left, as Excel's Insert does.** Measured
  against Excel, the new row takes the height, row style and outline level
  of the one above, and each of its cells' styles, empty; a new column
  takes the width, style and outline level of its neighbour the same way.
  A conditional format, a data validation or a protected range that ends
  on the row above grows over the new rows, and a sparkline there is copied
  down, reading its data that much further down. Nothing is hidden, no value
  comes across, and a row inserted at row 1 or a column at A takes nothing.
  A visible row or column that lands in a collapsed group opens it, as Excel
  clears the group's folded mark. `insert_rows` and `insert_columns` used to
  insert plain rows and columns, and `copy_format=False` still does, even
  inside a run of columns sharing one width.

### Fixed

- **Removing a chart left the chart behind.** Its part, its relationship
  and its content-type entry now go, and so do the style and colour parts
  Excel gives every chart it makes.
  [#3](https://github.com/WilliamSmithEdward/pyOfficeEditor/issues/3)

- **Removing a group left its members' parts behind.** A picture's image,
  a chart, and a Forms control's record, part and VML shape now go with
  the group. Every relationship only the removed shape used goes, a
  hyperlink's included, and each part it reached goes once nothing else
  uses it.
  [#5](https://github.com/WilliamSmithEdward/pyOfficeEditor/issues/5)

- **A group member's name reached the whole group.** `remove_shape` on a
  name only a group member carried removed the group. A sheet lists
  groups, not their members, so that name now raises `KeyError`, and a new
  shape may not take it.

- **Removing the last Forms control on a sheet made Excel refuse the
  file.** The records went, and the `mc:AlternateContent` Excel wraps them
  in stayed behind empty. It now goes with them.

- **A control still called by its default name kept its VML.** Excel
  writes the VML of a "Button 1" with its number as its id and no
  `o:spid`, and only the `o:spid` was looked for. `remove_shape` left that
  VML shape behind and Excel went on showing the button, and
  `set_shape_macro` missed the VML's copy of the macro. Renaming such a
  control now gives its VML the `o:spid` Excel gives a renamed one.

- **A note or a control added beside a header or footer picture made
  Excel refuse the file.** The picture is VML related the same way, in a
  part Excel lists first, and the note or control went into it. Notes and
  controls now keep to the sheet's own VML part.

- **Removing an embedded OLE object left Excel showing it.** Only its
  hidden drawing shape went, and Excel draws the object from its record
  and its VML. `remove_shape` now refuses it, as it refuses an ActiveX
  control.

- **A shape's position went stale when rows or columns moved.** Inserting
  or deleting them moved a shape's anchor but not its own transform, which
  is where a shape's box is read from, so `top` and `left` kept their old
  values. The transform now moves with the anchor, by the distance the
  anchor moved on the sheet's own grid; a group's members stay in the
  group's own coordinates, as Excel leaves them.

- **Every drawn object moved as if it sized with its cells.** Excel gives
  each one a placement, and measured, each takes an edit its own way. A
  picture as Excel inserts one moves without sizing, so a row inserted or
  deleted inside it stretched or shrank it where Excel keeps its size, and
  an object set not to move went with the cells where Excel leaves it. Each
  now moves by the placement its drawing anchor, its record on the sheet
  and its VML all state. An edge whose row is deleted lands on the boundary
  at offset zero, and a bottom or right edge lying on the line where rows
  or columns go in stays, both as Excel has them.

- **A deletion that took all of a control's rows made Excel refuse the
  file.** The control's anchors turned upside down. A shape, chart, group
  or Forms control that moves and sizes with its cells now goes when all
  of its rows or columns do, parts and all, as Excel deletes it. A deletion
  that would take an ActiveX control or an embedded object whole is
  refused, having changed nothing.

- **A note's box stretched when a row went in at its cell.** Measured,
  the box follows the cell it annotates, by as many rows or columns as the
  cell moves, and keeps its size whatever the rows inside it do.

- **IPMT, PPMT, CUMIPMT and CUMPRINC differed from Excel in the last
  digits.** They were computed exactly and rounded once. Measured, Excel
  forms each from the pieces PMT uses, through its own accurate logarithm
  and exponential: PPMT discounts the loan's share by the payments left,
  and the cumulative functions are closed forms over the span. All four
  now give Excel's bits on 2,640 loans each. IPMT and PPMT also take a
  period up to the term plus one and refuse a rate of -1 or below, and
  CUMIPMT and CUMPRINC check a fractional start or end as given, then
  round the start up and the end down, all as Excel does; they used to
  truncate both first.

- **A LAMBDA called without one of its arguments still ran.** Measured,
  that is `#VALUE!`, as an argument too many already was. An argument
  left empty, as in `f(1,)`, counts as omitted.

- **A parameter that takes only a range took an array whole.** Measured,
  Excel takes the array one item at a time, and no item is a range, so the
  result is an array of `#VALUE!` the array's size. That holds for
  COUNTIF, SUMIF and the other conditional functions, COUNTBLANK, RANK,
  ROW, COLUMN, SUBTOTAL, OFFSET, CELL, AREAS, ISFORMULA, FORMULATEXT and
  the database functions. They gave one `#VALUE!`, but COUNTBLANK counted
  the array's blanks and RANK ranked within it. It is why a LAMBDA calling
  COUNTBLANK makes a GROUPBY `#VALUE!` whole. COUNTBLANK of a union is
  `#VALUE!` too.

- **TAKE and DROP of a range gave an array.** Measured, they give a range,
  so COUNTIF, SUMIF's sum range, ROW and CELL take their result as they
  would the range itself. A reference of several areas is `#VALUE!`.

- **TRIMRANGE refused an array.** It trims an array's blank edges as it
  trims a range's. An array all blank is `#VALUE!`.

- **SHEET and SHEETS gave `#VALUE!` for what is not a range.** Excel gives
  `#N/A` for a number, a logical or an array, and SHEETS for a sheet's
  name too. An error passes through.

- **COUNTA left an array's blank items out.** Every item of an array
  counts, blank or not; only a range's blank cells are left out.

- **TEXTJOIN with ignore_empty left out kept the empty items.** Left out,
  it is TRUE. A blank cell given for it is FALSE, and a range of more than
  one cell is `#VALUE!`.

- **VSTACK and HSTACK given an error on its own gave that error.** The
  error is one item of the stack, as a function given on its own is.

- **TRANSPOSE took a whole range in a formula written before dynamic
  arrays.** Measured, Excel cuts the range to the formula's own row or
  column first, as it does a single value, so `=SUM(TRANSPOSE(A1:A3))`
  in row 5 is `#VALUE!`. In a dynamic-array formula it still takes the
  range whole.

- **Text read as a number misread thousands separators, fractions and a
  currency sign beside a percent sign.** VALUE, and text in arithmetic,
  took only groups of exactly three digits after a comma, so `1,0000`
  was not a number, and took `0,123` and `$5%`, which Excel refuses.
  Measured on 957 strings, which now match Excel but for `1 -1/2`, a
  date to Excel: every group after the first has three digits or more
  and the first is not all zeros, a fraction may follow grouped digits,
  as in `1,000 1/2`, but not a decimal point, its numerator and
  denominator are at most 32767, and a currency sign and a percent sign
  together are not a number.

## [0.3.0] - 2026-09-23

### Added

- **A formula engine.** `Workbook.calculate()` calculates every formula
  and writes each result into its cell as Excel's own recalculation
  would; `Worksheet.evaluate("SUM(B2:B9)", at="C1")` gives what a formula
  would give in a cell without putting it there. Formulas read one
  another in whatever order their inputs allow, on an explicit stack, so
  a column of thousands of running totals calculates without deep
  recursion; a circular reference keeps the values Excel cached.

  Behind them are a parser for formula text as a file stores it, sheet
  qualifiers, 3D references, whole rows and columns, structured
  references, array constants and the reference operators included, and
  an evaluator with Excel's precedence (`-2^2` is 4, `2^3^2` is 64), its
  implicit intersection of a range where one value is wanted, array
  formulas, defined names, tables, `INDIRECT` and `OFFSET`, and 493 of
  Excel's 525 functions: math and trigonometry, statistics and the
  distributions, financial, engineering, text, information, lookup and
  reference, dates and times, the database functions, the dynamic array
  functions but `GROUPBY` and `PIVOTBY`, and `LET` and `LAMBDA` with its
  helpers.

  Every result is held to Excel, in `formulas.xlsx`: 10,958 formulas Excel
  typed in and calculated over inputs written here as exact doubles, and
  the engine gives the value Excel cached for each, to the last bit where
  the corpus can tell. That settled rules no reference states. A number
  becomes text with fifteen digits, plain while it fits in twenty
  characters. `=` compares numbers read to fifteen digits, so
  `0.1+0.2=0.3`. A formula's last `+` or `-` gives 0 when the result is
  under eight units in the last place of its left operand, and so does
  SUM's last addition, while `(0.1+0.2-0.3)` keeps its rounding error.
  Text reads as a number the way typing it would, `"$1,000"`, `"(5)"`,
  `"1 1/2"`, `"12:30 PM"` and `"Jan 15, 2020"` among them, fifteen digits
  cut, not rounded. `^` is square-and-multiply for a whole exponent and
  takes a negative one's reciprocal in an x87 register, SIN, COS and TAN
  reduce with pi to 64 bits, and ROUNDUP adds its step in floating point:
  `ROUNDUP(2.675,2)` is `2.6799999999999997`.

  Arithmetic is the x87's, as Excel's is: every `+`, `-`, `*`, `/` and
  square root rounds to a 64-bit mantissa and then to a double, and EXP
  and LN follow the x87's own instructions to the bit. Below 1, e^x - 1
  is Kahan's `(u-1)*x/LN(u)` with `u=EXP(x)`, where the logarithm undoes
  u's rounding: SINH is `(a-b)/2` and TANH `(a-b)/(a+b+2)` on
  `a=e^x-1` and `b=e^-x-1`, COTH and CSCH their reciprocals, and
  WEIBULL.DIST's distribution is `-(e^-t-1)`. PMT forms `w=1-(1+r)^-n`
  as `-(e^-x-1)` with `x=n*ln(1+r)`, where Excel's own ln(1+r) below
  0.375 in size sums `2*atanh(r/(2+r))` as a series and above takes LN;
  the payment is `-((pv+fv)/w-fv)*r`, with `1/(1/r+1)` for r when
  payments fall at the start of each period. Where the corpus
  pinned an algorithm down the engine uses it: GAMMALN from 8 up is
  Stirling's series to ten terms, and below 8 is reduced into [2, 3)
  by the logarithm of a product; the normal density multiplies
  `EXP(-z*z/2)` by 1/sqrt(2 pi) as a double; ERF, ERFC and NORM.S.DIST
  square their argument to a double before the error function, which is
  most of how far Excel's tail values are from the exact ones; LINEST
  takes a Householder QR of a column of ones followed by the centred
  data; FV and PV form the annuity factor as `(1+r*type)*(((1+r)^n-1)/r)`,
  which decides their bits where the payment and the balance cancel; and
  RATE and IRR are Excel's own secant iterations, RATE in the rate and IRR
  in the discount rate `r/(1+r)`, each from the guess and a point 0.001
  away, stopping where a step and its residual both come under 1e-7, so
  each stops short of the root exactly where Excel does. IRR gives up on a
  guess after 200 steps or a repeated residual and starts again from 0.1,
  as Excel does. XIRR is not Newton's method but halving: from the guess
  it doubles outward until the root is bracketed, then halves until the
  bracket is narrower than 1e-8 of `0.5+|r|` and XNPV is under 1e-8 of
  the discounted amounts' sizes, the documented 0.000001 percent. PRICE,
  DURATION and MDURATION count the days to the next coupon as what is
  left of the period, E - A, on actual/360 and actual/365 too, where
  COUPDAYSNC counts days on the calendar. PRICE adds the coupons, then
  the redemption, then takes off the accrued interest, formed from the
  rate as `A/E*rate*100/f`; DURATION times each coupon as PRICE does,
  `index+DSC/E`, but the redemption as `DSC/E+N-1`, which rounds another
  way, and MDURATION divides it by `1+y/f`. YIELD starts from the
  textbook estimate with a quarter of the discount taken off its
  denominator, `(100*r*Y+B)/(100*Y-B*Y/2-B/4)`, where `B=100-price` and
  Y is the years to maturity, then takes Newton steps on PRICE with its
  exact slope. Once a step would be under 1e-10 it answers with the yield
  it would step from, up to 1e-10 short of the root, as Excel does. With
  one coupon left it is the documented closed form, with the period on
  every actual basis the calendar's and the days to redemption counted
  by the basis. Where
  Excel's own approximation is not yet reproduced, as for
  the error and incomplete gamma and beta functions and GAMMALN from 0.7
  to 3, the engine gives the double nearest the exact value, and the
  corpus test records how many units in the last place each function
  may be from Excel's. Sums run in order, never through Python's `sum`,
  which compensates from Python 3.12 on and would make a result depend
  on the interpreter.

  A function the engine does not have, a reference to another workbook
  or a data table keeps its cell's cached value, and so does everything
  that reads it; the `Calculation` that `calculate()` returns lists
  them, and only a complete calculation makes the cached values trusted
  again. `Worksheet.evaluate` raises `UnsupportedFormulaError` instead.

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

- **Named cell styles**: `Workbook.cell_styles`, `cell_style` and
  `add_cell_style`, `Worksheet.get_cell_style` and `set_cell_style`,
  `Cell.style` and `Range.apply_style`, with `CellStyle`. All 46 of
  Excel's own styles, Good to 60% - Accent6, are defined from Excel's
  definitions the first time a cell uses one, as Excel defines them, and
  Excel shows every one applied here exactly as it shows it applied by
  itself. The definitions do not follow the Normal style, measured: with
  Normal in Arial 10, Good is still in the theme's body font at 11 points
  and Title in its heading font at 18, so the typefaces come from the
  theme.

  A style sets some aspects of a cell and not others, the Style dialog's
  "Style includes" boxes, and applying one replaces only those, measured:
  Good over a bold, centred cell showing two decimals gives Good's font and
  fill and keeps the decimals and the centring. A cell's own `apply...`
  flags now mark what differs from its style, as Excel marks them.

- **Charts, read**: `Worksheet.charts` and `ChartSheet.chart` give each
  chart's kind, its series with where each takes its name, categories and
  values, and its title, as `Chart` and `ChartSeries`. A series' `formula`
  is the one Excel's formula bar shows, `=SERIES(Data!$B$1,...,1)`, the
  same character for character for every chart in `charts.xlsx`.

- **Charts, added**: `Worksheet.add_chart` for a column, bar, line, line
  with markers, pie, doughnut, scatter or area chart of a block of cells,
  on the same sheet or another. Written as Excel's Insert Chart writes
  one, measured kind by kind: for each, the chart part is Excel's own byte
  for byte, cached values and all, bar the ids Excel draws at random, and
  Excel reports every chart added here as it reports its own. A block's
  series run down its columns when it is taller than it is wide and along
  its rows otherwise, square blocks included, as Excel lays them out,
  measured; series past the sixth take the default palette's variations
  in the order Excel gave them to twenty.

- **Pivot tables, read**: `Worksheet.pivot_tables` gives each one's name,
  the block of the sheet it fills, and the range, table or name its cache
  was read from, as `PivotTable`.

- **Chart sheets**, `Workbook.chart_sheets` and `chart_sheet`, as
  `ChartSheet`. They keep their place among the tabs, in `sheet_names` and
  in the positions a defined name's scope and the active tab count by, and
  can be renamed, moved and removed. `sheets`, iteration and `book[0]` are
  the worksheets, and `Workbook.active` may be a chart sheet.

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

- **Inserting or deleting rows or columns left charts reading the old
  cells.** A chart's references were not moved at all, so a series kept
  reading the rows it read before an insertion above it. Every chart in the
  workbook moves now, on any sheet or a chart sheet, as Excel moves it,
  measured edit by edit: a series grows with rows inserted inside it and
  shrinks with rows deleted from it, and one deleted outright is written
  `Data!#REF!`, as is a title linked to a deleted cell. Renaming a sheet
  renames it in every chart too.

- **Inserting or deleting rows or columns left pivot tables behind.** A
  pivot table's place on its sheet and its cache's source range were not
  moved, so the table claimed cells its data no longer filled and the cache
  read the wrong rows. Both move now, as Excel moves them, measured edit by
  edit. As Excel does, an insertion inside a pivot table or a deletion of
  part of one is refused, since Excel refuses it too; inserting at its first
  row or column moves it whole, and deleting all its rows or columns
  deletes it, and its cache when no other pivot table reads from that. A
  source deleted outright keeps its address, as Excel keeps it, and a
  renamed sheet is renamed in every cache that reads from it.

- **A chart sheet was opened as a worksheet.** Writing a cell or inserting
  a row there would have put cells in a part that holds a chart.

- **Removing a sheet left its drawing, charts, comments and the like
  behind** in the package, unused. Every part only that sheet used goes
  with it now, as it does when Excel saves; a picture another sheet also
  shows stays.

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
`pictures.xlsx` a fourth, with `pictures_answers.json`, and
`richtext.xlsx` a fifth, with `richtext_answers.json` recording the font
Excel reported at every change in its text, and `escapes.xlsx` a sixth,
with text XML cannot carry as it is everywhere a workbook keeps text and
`escapes_answers.json` recording each character Excel read back, and
`styles.xlsx` a seventh, every built-in cell style applied by Excel, with
`styles_answers.json` recording what Excel showed for each cell and what
each style includes, and `charts.xlsx` an eighth, with charts on three
kinds of sheet and `charts_answers.json` recording every reference in them
as Excel wrote it, before any edit and after each of nine. Those are read
from the files Excel saved: its object model reports a deleted reference by
its old address, and once the file is reopened refuses to report the series
at all. `pivots.xlsx` is a ninth, two pivot tables reading one range, with
`pivots_answers.json` recording where each was and what its cache read
after each of nineteen edits, or that Excel refused the edit, and
`chartkinds.xlsx` a tenth, one chart of each kind Excel's Insert Chart
makes, the markup a chart added here is held to.
`filter_semantics.json` and `number_formats.json` are new measured
corpora, rebuilt by `scripts/measure_filters.py` and
`scripts/measure_number_formats.py` on a machine with Excel; every case
in them is a test. `formulas.xlsx` is a third, rebuilt by
`scripts/measure_formulas.py`: every formula in it is calculated by the
engine and has to give what Excel cached.

Measuring it turned up one more refusal: Excel will not open a file whose
formula holds a number it cannot hold, such as `1E-320` or `1.8E+308`,
although a cell's value that small opens.

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

A chart added here has no chart-style or colour part, the two Excel keeps
for its Chart Styles gallery; everything it looks like is in its own part,
as for a chart made through `ChartObjects.Add`. Dates along its axis are
written as text categories, where Excel would give the chart a date axis.

The author of a threaded comment written here is a person no account
stands behind, as Excel writes someone who is not signed in. Excel names
the signed-in account in its own, which is also why no Excel-authored
thread is committed as a fixture: the tests hold the reader to Excel's
measured markup with the person replaced.

The formula engine calculates in the en-US locale the corpus was measured
in, which decides the dates text reads as: `"1/2/2020"` is 2 January. A
formula that spills is calculated over the block its file records, which
is not grown or shrunk to fit. Thirty-two of Excel's functions are not
implemented, and a cell calling one keeps its cached value: those that
reach outside the workbook, `WEBSERVICE`, `FILTERXML`, `RTD`,
`STOCKHISTORY`, `TRANSLATE`, `DETECTLANGUAGE`, `IMAGE`, `PY`, `CALL`,
`REGISTER.ID` and the seven cube functions, and `GETPIVOTDATA`,
`GROUPBY`, `PIVOTBY`, the four `FORECAST.ETS` functions, the four
odd-period bond functions, `BAHTTEXT`, `EUROCONVERT`, `INFO` and
`PHONETIC`.

Where Excel's own approximation is not yet reproduced, a result may be a
few units in the last place from Excel's, and the corpus test names each
such function with the most it may differ by: ASIN below 0.35, where
Excel's is not odd, IPMT, PPMT, CUMIPMT and CUMPRINC, the error and gamma
functions, most of the statistical distributions and the tests and
intervals built on them, LINEST past three points and TREND with it, and
GEOMEAN. When IRR's secant fails
from the guess and again from 0.1, Excel finds the root a third way not
yet reproduced; the engine halves down to it, within 1e-11 of Excel's
answer.

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
