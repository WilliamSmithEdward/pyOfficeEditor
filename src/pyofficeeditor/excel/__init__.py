"""The Excel surface: workbooks, worksheets, ranges and cells.

    from pyofficeeditor.excel import Workbook

    with Workbook.open("orders.xlsx") as book:
        sheet = book["Data"]
        sheet["B2"].value            # 120
        sheet["A2"].value            # 'North'      a shared string, resolved
        sheet["F2"].value            # date(2026, 1, 15)   a number, until its
                                     #   number format is followed
        sheet["D3"].formula          # 'B3*C3'      derived from its shared
                                     #   formula's master, which is D2
        sheet["B2"].value = 200
        book.save()

Reading any of those four cells naively gives a different, plausible, wrong
answer: an index instead of text, a serial instead of a date, an empty
formula instead of a translated one. The modules under this package exist to
get each of them right, and each is tested against bytes real Excel wrote.
"""

from __future__ import annotations

from pyofficeeditor.excel._conditional import (
    Cfvo,
    ColorScale,
    ConditionalFormatting,
    ConditionalRule,
    DataBar,
    IconSet,
    average,
    bar,
    begins_with,
    cell_is,
    contains_text,
    duplicates,
    during,
    ends_with,
    expression,
    gradient,
    icons,
    is_blank,
    is_error,
    not_contains_text,
    top,
    uniques,
)
from pyofficeeditor.excel._dxf import Dxf, DxfFill, DxfFont
from pyofficeeditor.excel._formats import (
    Alignment,
    Border,
    BorderStyle,
    CellFormat,
    Color,
    Fill,
    Font,
    Protection,
    Side,
)
from pyofficeeditor.excel._pagesetup import (
    HeaderFooter,
    HeaderFooterText,
    PageMargins,
    PageSetup,
    PrintOptions,
)
from pyofficeeditor.excel._protection import SheetProtection, hash_password
from pyofficeeditor.excel._reference import (
    MAX_COLUMN,
    MAX_ROW,
    CellRef,
    RangeRef,
    column_index,
    column_letter,
)
from pyofficeeditor.excel._sharedstrings import SharedStrings
from pyofficeeditor.excel._styles import Styles
from pyofficeeditor.excel._tables import Table, TableColumn, TableStyle
from pyofficeeditor.excel._validation import DataValidation, ErrorStyle, ValidationOperator, ValidationType
from pyofficeeditor.excel._values import CellError, CellValue, DateOutOfRangeError
from pyofficeeditor.excel.workbook import Workbook
from pyofficeeditor.excel.worksheet import Cell, Range, Worksheet

__all__ = [
    "MAX_COLUMN",
    "MAX_ROW",
    "Alignment",
    "Border",
    "BorderStyle",
    "Cell",
    "CellError",
    "CellFormat",
    "CellRef",
    "CellValue",
    "Cfvo",
    "Color",
    "ColorScale",
    "ConditionalFormatting",
    "ConditionalRule",
    "DataBar",
    "DataValidation",
    "DateOutOfRangeError",
    "Dxf",
    "DxfFill",
    "DxfFont",
    "ErrorStyle",
    "Fill",
    "Font",
    "HeaderFooter",
    "HeaderFooterText",
    "IconSet",
    "PageMargins",
    "PageSetup",
    "PrintOptions",
    "Protection",
    "Range",
    "RangeRef",
    "SharedStrings",
    "SheetProtection",
    "Side",
    "Styles",
    "Table",
    "TableColumn",
    "TableStyle",
    "ValidationOperator",
    "ValidationType",
    "Workbook",
    "Worksheet",
    "average",
    "bar",
    "begins_with",
    "cell_is",
    "column_index",
    "column_letter",
    "contains_text",
    "duplicates",
    "during",
    "ends_with",
    "expression",
    "gradient",
    "hash_password",
    "icons",
    "is_blank",
    "is_error",
    "not_contains_text",
    "top",
    "uniques",
]
