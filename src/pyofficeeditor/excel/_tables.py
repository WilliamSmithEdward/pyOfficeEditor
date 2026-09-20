"""Tables, which Excel's object model calls ListObjects.

A table is not a property of a range. It is a part of its own, wired to the
sheet four ways, and all four have to line up or Excel either ignores the
table or refuses the file::

    xl/tables/table1.xml            the table itself
    [Content_Types].xml             an Override naming its content type
    xl/worksheets/_rels/sheet1.xml.rels   a relationship from the sheet
    xl/worksheets/sheet1.xml        <tableParts><tablePart r:id="rId2"/>

The part Excel writes for a three-column table with a totals row::

    <table id="1" name="SalesTable" displayName="SalesTable"
           ref="A1:C5" totalsRowCount="1">
      <autoFilter ref="A1:C4"/>
      <tableColumns count="3">
        <tableColumn id="1" name="Region" totalsRowLabel="Total"/>
        <tableColumn id="2" name="Units" totalsRowFunction="sum"/>
        <tableColumn id="3" name="Revenue" totalsRowFunction="sum">
          <calculatedColumnFormula>B2*10</calculatedColumnFormula>
        </tableColumn>
      </tableColumns>
      <tableStyleInfo name="TableStyleMedium2" showRowStripes="1" .../>
    </table>

Three details in that are easy to get wrong.

``ref`` spans the whole table *including* the totals row, while the
``autoFilter`` spans only the header and the data. Giving them the same
extent puts a filter dropdown on the totals row.

Whether a totals row exists is spelled two different ways. With one,
``totalsRowCount="1"``; without one, ``totalsRowShown="0"``. Writing neither
leaves Excel to guess.

A column's name is not free text: it has to equal the text in the header
cell. Excel reconciles them on open by rewriting the part, so a mismatch
looks like the library's edit being silently undone.
"""

from __future__ import annotations

from dataclasses import dataclass

from pyofficeeditor._xml import Element, XmlDocument
from pyofficeeditor.excel._names import MAX_NAME_LENGTH, check_name
from pyofficeeditor.excel._reference import CellRef, RangeRef

NS_SPREADSHEETML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

CT_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
RT_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"

#: The order ``CT_Table`` requires its children to appear in.
TABLE_CHILD_ORDER: tuple[str, ...] = (
    "autoFilter",
    "sortState",
    "tableColumns",
    "tableStyleInfo",
    "extLst",
)

#: ``ST_TotalsRowFunction``.
TOTALS_FUNCTIONS = frozenset(
    {
        "none",
        "sum",
        "min",
        "max",
        "average",
        "count",
        "countNums",
        "stdDev",
        "var",
        "custom",
    }
)

#: The longest table name Excel accepts.  Same as a defined name's, because
#: a table name is one.
MAX_TABLE_NAME_LENGTH = MAX_NAME_LENGTH


def check_table_name(name: str, *, taken: set[str] | None = None) -> None:
    """Refuse a table name Excel would refuse.

    A table name *is* a defined name: formulas refer to it, so it follows the
    same rules, and it is unique across the whole workbook rather than per
    sheet. The rules live in :mod:`pyofficeeditor.excel._names` so both kinds
    of name are checked by one piece of code.
    """
    check_name(name, taken=taken, what="table", unique_within="the workbook")


@dataclass(frozen=True)
class TableStyle:
    """Which banded-style options a table shows."""

    name: str | None = "TableStyleMedium2"
    show_first_column: bool = False
    show_last_column: bool = False
    show_row_stripes: bool = True
    show_column_stripes: bool = False

    @classmethod
    def read(cls, element: Element | None) -> TableStyle:
        if element is None:
            return cls(name=None, show_row_stripes=False)
        return cls(
            name=element.get("name"),
            show_first_column=element.get("showFirstColumn") in ("1", "true"),
            show_last_column=element.get("showLastColumn") in ("1", "true"),
            show_row_stripes=element.get("showRowStripes") in ("1", "true"),
            show_column_stripes=element.get("showColumnStripes") in ("1", "true"),
        )

    def write(self) -> Element:
        element = Element.create("tableStyleInfo")
        if self.name is not None:
            element.set("name", self.name)
        element.set("showFirstColumn", "1" if self.show_first_column else "0")
        element.set("showLastColumn", "1" if self.show_last_column else "0")
        element.set("showRowStripes", "1" if self.show_row_stripes else "0")
        element.set("showColumnStripes", "1" if self.show_column_stripes else "0")
        return element


class TableColumn:
    """One column of a table, as a view over its element."""

    __slots__ = ("_element", "_table")

    def __init__(self, table: Table, element: Element) -> None:
        self._table = table
        self._element = element

    @property
    def id(self) -> int:
        """The column's own identifier, which is stable rather than
        positional: moving a column does not renumber it."""
        raw = self._element.get("id")
        try:
            return int(raw) if raw is not None else 0
        except ValueError:
            return 0

    @property
    def name(self) -> str:
        return self._element.get("name") or ""

    @property
    def totals_label(self) -> str | None:
        """The literal text in this column's totals cell, if it has one."""
        return self._element.get("totalsRowLabel")

    @property
    def totals_function(self) -> str | None:
        """The aggregate shown in this column's totals cell, such as
        ``'sum'``."""
        return self._element.get("totalsRowFunction")

    @property
    def calculated_formula(self) -> str | None:
        """The formula filling the whole column, if it has one.

        Stored once for the column rather than per cell, the same idea as a
        shared formula.
        """
        child = self._element.child("calculatedColumnFormula")
        return None if child is None else (child.text or None)

    def __repr__(self) -> str:
        return f"<TableColumn {self.name!r} id={self.id}>"


class Table:
    """A table on one worksheet."""

    def __init__(self, sheet: object, part_name: str, document: XmlDocument) -> None:
        self._sheet = sheet
        self._part_name = part_name
        self._document = document
        self._root = document.root

    # -- identity -------------------------------------------------------

    @property
    def part_name(self) -> str:
        return self._part_name

    @property
    def document(self) -> XmlDocument:
        return self._document

    @property
    def id(self) -> int:
        raw = self._root.get("id")
        try:
            return int(raw) if raw is not None else 0
        except ValueError:
            return 0

    @property
    def name(self) -> str:
        return self._root.get("name") or ""

    @property
    def display_name(self) -> str:
        """The name shown in Excel's UI.

        Excel keeps it equal to :attr:`name` and it is what formulas use, so
        renaming sets both.
        """
        return self._root.get("displayName") or self.name

    # -- extent ---------------------------------------------------------

    @property
    def ref(self) -> RangeRef:
        """The whole table, header and totals row included."""
        raw = self._root.get("ref")
        if raw is None:
            raise ValueError(f"{self._part_name} has no ref, so it covers nothing.")
        return RangeRef.parse(raw).normalized

    @property
    def header_row_count(self) -> int:
        raw = self._root.get("headerRowCount")
        if raw is None:
            return 1
        try:
            return int(raw)
        except ValueError:
            return 1

    @property
    def totals_row_count(self) -> int:
        raw = self._root.get("totalsRowCount")
        if raw is None:
            return 0
        try:
            return int(raw)
        except ValueError:
            return 0

    @property
    def has_totals_row(self) -> bool:
        return self.totals_row_count > 0

    @property
    def header_row(self) -> int | None:
        """The row the column names sit in, or ``None`` for a headerless
        table."""
        return self.ref.top if self.header_row_count else None

    @property
    def totals_row(self) -> int | None:
        return self.ref.bottom if self.has_totals_row else None

    @property
    def data_range(self) -> RangeRef | None:
        """The rows between the header and the totals, or ``None`` when the
        table holds no data rows."""
        block = self.ref
        top = block.top + self.header_row_count
        bottom = block.bottom - self.totals_row_count
        if top > bottom:
            return None
        return RangeRef(CellRef(top, block.left), CellRef(bottom, block.right))

    @property
    def filter_ref(self) -> RangeRef | None:
        """What the autofilter covers, which excludes the totals row.

        Giving the filter the table's full extent puts a dropdown on the
        totals row, which is why the two are separate.
        """
        element = self._root.child("autoFilter")
        if element is None:
            return None
        raw = element.get("ref")
        if raw is None:
            return None
        try:
            return RangeRef.parse(raw).normalized
        except ValueError:
            return None

    # -- columns --------------------------------------------------------

    @property
    def columns(self) -> list[TableColumn]:
        container = self._root.child("tableColumns")
        if container is None:
            return []
        return [TableColumn(self, entry) for entry in container.children_named("tableColumn")]

    @property
    def column_names(self) -> list[str]:
        return [column.name for column in self.columns]

    def column(self, name: str) -> TableColumn:
        for column in self.columns:
            if column.name.casefold() == name.casefold():
                return column
        raise KeyError(
            f"table {self.name!r} has no column {name!r}. It has: {', '.join(self.column_names)}"
        )

    # -- style ----------------------------------------------------------

    @property
    def style(self) -> TableStyle:
        return TableStyle.read(self._root.child("tableStyleInfo"))

    @style.setter
    def style(self, style: TableStyle) -> None:
        existing = self._root.child("tableStyleInfo")
        if existing is not None:
            self._root.remove(existing)
        insert_table_child(self._root, style.write())

    def __repr__(self) -> str:
        return f"<Table {self.name!r} {self.ref.a1} columns={len(self.columns)}>"


def insert_table_child(parent: Element, element: Element) -> None:
    """Put a child of ``<table>`` where ``CT_Table``'s sequence wants it."""
    from pyofficeeditor.excel._schema import insert_in_schema_order

    insert_in_schema_order(parent, element, TABLE_CHILD_ORDER)


def build_table_part(
    *,
    identifier: int,
    name: str,
    ref: RangeRef,
    column_names: list[str],
    has_totals_row: bool,
    style: TableStyle,
) -> XmlDocument:
    """A new table part.

    ``ref`` is the whole table; the autofilter gets the header and data rows
    only, since a filter dropdown does not belong on a totals row. Excel
    stamps its parts with revision GUIDs, which are extension attributes and
    are left out here.
    """
    root = Element.create(
        "table",
        {
            "xmlns": NS_SPREADSHEETML,
            "id": str(identifier),
            "name": name,
            "displayName": name,
            "ref": ref.a1,
        },
    )
    # The two spellings are not interchangeable: Excel writes the count when
    # there is a totals row and the shown flag when there is not.
    if has_totals_row:
        root.set("totalsRowCount", "1")
    else:
        root.set("totalsRowShown", "0")

    filter_bottom = ref.bottom - (1 if has_totals_row else 0)
    filter_ref = RangeRef(CellRef(ref.top, ref.left), CellRef(filter_bottom, ref.right))
    root.append(Element.create("autoFilter", {"ref": filter_ref.a1}))

    columns = Element.create("tableColumns", {"count": str(len(column_names))})
    for index, column_name in enumerate(column_names, start=1):
        columns.append(Element.create("tableColumn", {"id": str(index), "name": column_name}))
    root.append(columns)
    root.append(style.write())
    return XmlDocument(root)


def unique_column_names(wanted: list[str]) -> list[str]:
    """Column names with blanks filled and duplicates disambiguated.

    A table's column names have to be non-empty and distinct, and Excel
    invents ``Column1`` and appends a digit rather than refusing the table,
    so doing the same here keeps a caller from producing a file Excel then
    silently rewrites.
    """
    out: list[str] = []
    seen: set[str] = set()
    for index, raw in enumerate(wanted, start=1):
        name = raw.strip() or f"Column{index}"
        candidate = name
        suffix = 1
        while candidate.casefold() in seen:
            suffix += 1
            candidate = f"{name}{suffix}"
        seen.add(candidate.casefold())
        out.append(candidate)
    return out


__all__ = [
    "CT_TABLE",
    "MAX_TABLE_NAME_LENGTH",
    "NS_SPREADSHEETML",
    "RT_TABLE",
    "TABLE_CHILD_ORDER",
    "TOTALS_FUNCTIONS",
    "Table",
    "TableColumn",
    "TableStyle",
    "build_table_part",
    "check_table_name",
    "insert_table_child",
    "unique_column_names",
]
