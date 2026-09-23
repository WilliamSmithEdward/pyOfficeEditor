"""Where a child element has to go.

SpreadsheetML declares the contents of a worksheet and of a workbook as
*sequences*, not as unordered bags. An element in the wrong place makes Excel
refuse the file rather than repair it, and the mistake is easy to make: the
natural thing to do with a missing element is append it, and appending is
wrong whenever anything that must follow it is already there.

Two cases in this library hit it directly. A worksheet Excel authored begins
with ``sheetPr``, so inserting a missing ``dimension`` at the front breaks the
sequence. And a workbook usually ends with ``extLst``, so appending a missing
``calcPr`` puts it after the element it must precede.

So both orders are written down here, and insertion consults them.
"""

from __future__ import annotations

from pyofficeeditor._xml import Element, local_name

#: The order ``CT_Worksheet`` requires its children to appear in.
WORKSHEET_CHILD_ORDER: tuple[str, ...] = (
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "sheetCalcPr",
    "sheetProtection",
    "protectedRanges",
    "scenarios",
    "autoFilter",
    "sortState",
    "dataConsolidate",
    "customSheetViews",
    "mergeCells",
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "drawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
)

#: The order ``CT_Workbook`` requires its children to appear in.
WORKBOOK_CHILD_ORDER: tuple[str, ...] = (
    "fileVersion",
    "fileSharing",
    "workbookPr",
    "workbookProtection",
    "bookViews",
    "sheets",
    "functionGroups",
    "externalReferences",
    "definedNames",
    "calcPr",
    "oleSize",
    "customWorkbookViews",
    "pivotCaches",
    "smartTagPr",
    "smartTagTypes",
    "webPublishing",
    "fileRecoveryPr",
    "webPublishObjects",
    "extLst",
)


#: The order ``CT_Stylesheet`` requires its children to appear in.
STYLESHEET_CHILD_ORDER: tuple[str, ...] = (
    "numFmts",
    "fonts",
    "fills",
    "borders",
    "cellStyleXfs",
    "cellXfs",
    "cellStyles",
    "dxfs",
    "tableStyles",
    "colors",
    "extLst",
)


#: The order ``CT_SheetPr`` requires its children to appear in.
SHEET_PR_CHILD_ORDER: tuple[str, ...] = (
    "tabColor",
    "outlinePr",
    "pageSetUpPr",
)

#: ``CT_AutoFilter``: every filter column before the sort state and the
#: extensions. Excel refuses a workbook with a column after the sort state.
AUTO_FILTER_CHILD_ORDER: tuple[str, ...] = (
    "filterColumn",
    "sortState",
    "extLst",
)


def insert_in_schema_order(parent: Element, element: Element, order: tuple[str, ...]) -> None:
    """Put ``element`` where ``order`` says it belongs among its siblings.

    It lands before the first existing child that has to follow it. A name
    the order does not mention is appended, which is the safe direction,
    since the extension elements that genuinely sort last are named.
    """
    try:
        position = order.index(local_name(element.name))
    except ValueError:
        parent.append(element)
        return
    for existing in parent.elements():
        try:
            other = order.index(local_name(existing.name))
        except ValueError:
            continue
        if other > position:
            parent.insert_before(existing, element)
            return
    parent.append(element)


def ensure_child(parent: Element, name: str, order: tuple[str, ...]) -> Element:
    """The named child, created in its schema position if it is absent."""
    existing = parent.child(name)
    if existing is not None:
        return existing
    created = Element.create(name)
    insert_in_schema_order(parent, created, order)
    return created


__all__ = [
    "AUTO_FILTER_CHILD_ORDER",
    "SHEET_PR_CHILD_ORDER",
    "STYLESHEET_CHILD_ORDER",
    "WORKBOOK_CHILD_ORDER",
    "WORKSHEET_CHILD_ORDER",
    "ensure_child",
    "insert_in_schema_order",
]
