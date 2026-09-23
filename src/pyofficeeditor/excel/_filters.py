"""The autofilter on a sheet: what it keeps, and what that hides.

A filter is stored twice, and this is the whole reason the module has an
evaluator in it rather than just a reader and a writer.

**Excel does not re-apply a filter when the workbook opens.** The criteria
live in ``<autoFilter>`` and which rows are out of view lives on each
``<row>`` as ``hidden="1"``, and Excel trusts the second. A file carrying
criteria with every ``hidden`` stripped opens with the dropdowns lit and
every row showing. So writing criteria means working out which rows they
keep and hiding the rest, and that has to agree with what Excel would keep.

The evaluator is held to what Excel kept, measured one criterion at a time
over columns of every kind of cell (``tests/fixtures/excel/
filter_semantics.json``). What that showed, beyond the documentation:

**A value list matches the text Excel shows**, the number format applied,
fill dropped and the cell's outer spaces trimmed, under Excel's collation:
case ignored, accents not. ``5`` keeps the number 5 and the text "5" but
not 5 shown as ``5.00``, and a date matches as the text its format gives.
A stored value is compared as written, so one with outer spaces of its own
matches nothing.

**A comparison is typed by its value.** A number compares with numbers
only, dates among them; TRUE or FALSE with booleans only; anything else
with text only, by collation. Blanks and errors pass no ordering.
``equal`` without wildcards is a value list; ``notEqual`` excludes only a
cell of the value's own type that equals it. Two values are special
whatever the operator: an empty one keeps the blanks, a single space keeps
everything else. A wildcard value matches text cells only, and a text cell
it does not match counts as greater, never less.

**Top ten and averages are worked out again** from the column each time
the filter is applied, whatever the file stored, unless the column holds
an error: then a top ten keeps every row and an average uses the stored
value, a missing one being zero. Relative date periods are likewise worked
out again from today, and date groups and periods only look at numbers
shown with a date format.

**Some criteria cannot be evaluated here**: a colour or an icon filter,
and markup this module cannot model. They are carried verbatim as an
:class:`OpaqueCriterion`, and a row only they could hide is left as it is.
"""

from __future__ import annotations

import datetime as dt
import math
import re
from collections.abc import Iterable, Sequence
from dataclasses import dataclass
from typing import Literal, get_args

from pyofficeeditor._xml import Element, XmlDocument
from pyofficeeditor.excel import _collate
from pyofficeeditor.excel._numfmt import date_parts
from pyofficeeditor.excel._values import CellError, datetime_to_serial, format_number
from pyofficeeditor.excel._xstring import decode, encode_attribute

# ---------------------------------------------------------------------------
# Vocabulary
# ---------------------------------------------------------------------------

#: ``ST_FilterOperator``: how one ``<customFilter>`` compares.
FilterOperator = Literal[
    "equal",
    "notEqual",
    "greaterThan",
    "greaterThanOrEqual",
    "lessThan",
    "lessThanOrEqual",
]
OPERATORS: tuple[FilterOperator, ...] = get_args(FilterOperator)

#: Excel's own spelling of each operator, as a criterion string writes it.
_SYMBOLS: dict[str, FilterOperator] = {
    "<>": "notEqual",
    ">=": "greaterThanOrEqual",
    "<=": "lessThanOrEqual",
    "=": "equal",
    ">": "greaterThan",
    "<": "lessThan",
}

#: How finely a date group matches, coarsest first.
DateGrouping = Literal["year", "month", "day", "hour", "minute", "second"]
GROUPINGS: tuple[DateGrouping, ...] = get_args(DateGrouping)

#: ``ST_DynamicFilterType``, less its ``null``.
DynamicType = Literal[
    "aboveAverage", "belowAverage",
    "today", "yesterday", "tomorrow",
    "thisWeek", "lastWeek", "nextWeek",
    "thisMonth", "lastMonth", "nextMonth",
    "thisQuarter", "lastQuarter", "nextQuarter",
    "thisYear", "lastYear", "nextYear", "yearToDate",
    "Q1", "Q2", "Q3", "Q4",
    "M1", "M2", "M3", "M4", "M5", "M6", "M7", "M8", "M9", "M10", "M11", "M12",
]
DYNAMIC_TYPES: tuple[DynamicType, ...] = get_args(DynamicType)
AVERAGES: tuple[DynamicType, ...] = ("aboveAverage", "belowAverage")
#: Periods measured from today, stored with the bounds they had then.
RELATIVE_PERIODS: tuple[DynamicType, ...] = DYNAMIC_TYPES[2:18]

#: Where Excel stores a date group since 2017: an extension of the
#: ``<filterColumn>`` rather than the ``<filters>`` inside it.
DATE_GROUP_EXTENSION = "{1AD28BCE-077C-4C59-8B6E-1921CE8616D4}"

#: Largest counts Excel accepts in a top ten.
MAX_TOP_ITEMS = 500
MAX_TOP_PERCENT = 100

#: What a comparison can compare with. Anything but text is stored the way
#: Excel stores it, a date as its serial.
ComparisonValue = str | int | float | dt.date | dt.datetime | dt.time

_NUMBER = re.compile(r"^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$")
_NOT_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")


def _check_text(value: str, what: str) -> None:
    if _NOT_XML.search(value):
        raise ValueError(f"{what} {value!r} holds a character XML cannot carry.")


# ---------------------------------------------------------------------------
# Criteria
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Comparison:
    """One ``<customFilter>``: an operator and the value it compares with.

    The value may be text, a number, a boolean or a date or time; all but
    text are stored the way Excel stores them, a date as its serial number
    in the workbook's date system. Text is stored as written: wildcards
    ``*`` and ``?`` in it make it a pattern, ``~`` escapes either.
    """

    operator: FilterOperator = "equal"
    value: ComparisonValue = ""

    def __post_init__(self) -> None:
        if self.operator not in OPERATORS:
            raise ValueError(
                f"{self.operator!r} is not a filter operator; use one of {', '.join(OPERATORS)}."
            )
        value: object = self.value
        if isinstance(value, str):
            _check_text(value, "comparison value")
        elif isinstance(value, (int, float)) and not isinstance(value, bool):
            if not math.isfinite(value):
                raise ValueError(f"{value!r} cannot be compared with; Excel has no such number.")
        elif not isinstance(value, (bool, dt.date, dt.datetime, dt.time)):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise TypeError(
                f"a comparison value is text, a number, a boolean or a date, not {type(value).__name__}."
            )


@dataclass(frozen=True)
class CustomFilter:
    """One or two comparisons, joined by and or by or.

    Excel refuses a workbook whose custom filter has no comparison and
    ignores one with more than two, so both are refused here.
    """

    comparisons: tuple[Comparison, ...]
    #: ``and="1"``: both must hold. Or is the default and Excel writes
    #: nothing for it; a single comparison is unaffected either way.
    require_all: bool = False

    def __post_init__(self) -> None:
        comparisons = tuple(self.comparisons)
        object.__setattr__(self, "comparisons", comparisons)
        if not 1 <= len(comparisons) <= 2:
            raise ValueError(
                f"a custom filter holds one or two comparisons, not {len(comparisons)}: Excel refuses "
                "a workbook with none and ignores more than two."
            )
        for one in comparisons:
            if not isinstance(one, Comparison):  # pyright: ignore[reportUnnecessaryIsInstance]
                raise TypeError(f"{one!r} is not a Comparison.")


@dataclass(frozen=True)
class DateGroup:
    """A date Excel's dropdown picked, at a grouping level.

    ``DateGroup("month", 2026, 1)`` keeps every date in January 2026;
    ``DateGroup("hour", 2026, 1, 5, 10)`` every time from 10:00 to 10:59
    that day. The fields down to the grouping are required and the ones
    below it must be absent.
    """

    grouping: DateGrouping
    year: int
    month: int | None = None
    day: int | None = None
    hour: int | None = None
    minute: int | None = None
    second: int | None = None

    def __post_init__(self) -> None:
        if self.grouping not in GROUPINGS:
            raise ValueError(f"{self.grouping!r} is not a date grouping; use one of {', '.join(GROUPINGS)}.")
        depth = GROUPINGS.index(self.grouping)
        fields = self.fields
        for index, (name, lowest, highest) in enumerate(_GROUP_FIELDS):
            value = fields[index]
            if index <= depth:
                if value is None:
                    raise ValueError(f"a {self.grouping} group needs its {name}.")
                if isinstance(value, bool) or not isinstance(value, int) or not lowest <= value <= highest:  # pyright: ignore[reportUnnecessaryIsInstance]
                    raise ValueError(f"{name} {value!r} is out of range {lowest} to {highest}.")
            elif value is not None:
                raise ValueError(f"a {self.grouping} group has no {name}; drop it or group more finely.")
        if depth >= 2:
            try:
                dt.date(self.year, self.month or 1, self.day or 1)
            except ValueError as error:
                raise ValueError(f"{self.year}-{self.month}-{self.day} is not a day.") from error

    @property
    def fields(self) -> tuple[int | None, ...]:
        return (self.year, self.month, self.day, self.hour, self.minute, self.second)

    @classmethod
    def of(cls, moment: dt.date | dt.datetime, grouping: DateGrouping) -> DateGroup:
        """The group a date or time falls in at a level."""
        parts: tuple[int, ...] = (moment.year, moment.month, moment.day)
        if isinstance(moment, dt.datetime):
            parts += (moment.hour, moment.minute, moment.second)
        depth = GROUPINGS.index(grouping)
        if depth >= len(parts):
            raise ValueError(f"a date has no {grouping}; pass a datetime to group by one.")
        values = [*parts[: depth + 1], *([None] * (5 - depth))]
        return cls(grouping, *values)  # type: ignore[arg-type]

    def contains(self, parts: tuple[int, int, int, int, int, int]) -> bool:
        """Whether a moment, as its six parts, falls in this group."""
        depth = GROUPINGS.index(self.grouping)
        return all(parts[index] == self.fields[index] for index in range(depth + 1))


_GROUP_FIELDS = (
    ("year", 1900, 9999), ("month", 1, 12), ("day", 1, 31),
    ("hour", 0, 23), ("minute", 0, 59), ("second", 0, 59),
)


@dataclass(frozen=True)
class ValueFilter:
    """Values to keep, and optionally the blanks and some date groups.

    A value is the text a cell shows, as Excel displays it: ``"1.40"``
    under ``0.00``, ``"1/5/2026"`` under ``m/d/yyyy``. Case is ignored.
    The cell's outer spaces are trimmed before comparing and the value's
    are not, so a value with outer spaces matches nothing.

    Excel refuses a workbook holding an empty value and ignores a list
    with nothing in it, so both are refused here.
    """

    values: tuple[str, ...] = ()
    #: ``blank="1"``: keep empty cells and cells holding only spaces.
    blank: bool = False
    date_groups: tuple[DateGroup, ...] = ()

    def __post_init__(self) -> None:
        if isinstance(self.values, str):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise TypeError("values is a tuple of texts; wrap a single value as (value,).")
        values = tuple(self.values)
        groups = tuple(self.date_groups)
        object.__setattr__(self, "values", values)
        object.__setattr__(self, "date_groups", groups)
        for value in values:
            if not isinstance(value, str):  # pyright: ignore[reportUnnecessaryIsInstance]
                raise TypeError(
                    f'a filter value is the text the cell shows, such as "1.40", not {type(value).__name__}.'
                )
            if not value:
                raise ValueError("a filter value cannot be empty; Excel refuses the workbook. Use blank=True.")
            _check_text(value, "filter value")
        for group in groups:
            if not isinstance(group, DateGroup):  # pyright: ignore[reportUnnecessaryIsInstance]
                raise TypeError(f"{group!r} is not a DateGroup.")
        if not values and not self.blank and not groups:
            raise ValueError("a value filter needs a value, blank=True or a date group; Excel ignores an empty one.")


@dataclass(frozen=True)
class Top10Filter:
    """The top or bottom N numbers of a column, or N percent of them.

    ``threshold`` is Excel's ``filterVal``, the number the ranking came
    down to. It is worked out from the column whenever the filter is
    applied, as Excel does, and stored for whoever reads the file.
    """

    count: int = 10
    percent: bool = False
    top: bool = True
    threshold: float | None = None

    def __post_init__(self) -> None:
        count: object = self.count
        if isinstance(count, bool) or not isinstance(count, int):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise TypeError(f"a top ten count is a whole number, not {count!r}.")
        highest = MAX_TOP_PERCENT if self.percent else MAX_TOP_ITEMS
        if not 1 <= count <= highest:
            raise ValueError(f"a top ten keeps 1 to {highest}{' percent' if self.percent else ''}, not {count}.")
        if self.threshold is not None and not math.isfinite(self.threshold):
            raise ValueError(f"threshold {self.threshold!r} is not a number.")


@dataclass(frozen=True)
class DynamicFilter:
    """A filter Excel works out from the column or from today.

    Above and below average store the average as ``value``; a relative
    date period such as ``thisMonth`` stores its bounds as ``value`` and
    ``maximum``, the second excluded. ``Q1`` to ``Q4`` and ``M1`` to
    ``M12`` store nothing and match their months in any year. All of it is
    worked out again whenever the filter is applied.
    """

    kind: DynamicType
    value: float | None = None
    maximum: float | None = None

    def __post_init__(self) -> None:
        if self.kind not in DYNAMIC_TYPES:
            raise ValueError(f"{self.kind!r} is not a dynamic filter; Excel refuses a workbook holding one.")
        for number in (self.value, self.maximum):
            if number is not None and not math.isfinite(number):
                raise ValueError(f"{number!r} is not a number.")


@dataclass(frozen=True)
class OpaqueCriterion:
    """A criterion carried as the markup it was read from.

    A colour or icon filter, or markup this module cannot model. It is
    written back exactly as read and never evaluated, except that a
    criterion Excel itself ignores, such as three comparisons, keeps every
    row, as it does in Excel.
    """

    markup: str
    #: Whether Excel ignores it when applying the filter.
    ignored: bool = False

    @property
    def kind(self) -> str:
        """The name of the element the criterion is, such as ``colorFilter``."""
        found = re.match(r"\s*<([A-Za-z_][\w.:-]*)", self.markup)
        return "" if found is None else found.group(1).rpartition(":")[2]


#: What a column can be filtered by.
Criterion = ValueFilter | CustomFilter | Top10Filter | DynamicFilter | OpaqueCriterion


@dataclass(frozen=True)
class FilterColumn:
    """One column's criterion, by its offset into the filtered range."""

    #: Zero-based, counted from the left of :attr:`AutoFilter.ref` rather
    #: than from column A.
    column: int = 0
    criterion: Criterion | None = None
    #: Whether this column's dropdown arrow is hidden.
    hide_dropdown: bool = False

    def __post_init__(self) -> None:
        if isinstance(self.column, bool) or not isinstance(self.column, int) or self.column < 0:  # pyright: ignore[reportUnnecessaryIsInstance]
            raise ValueError(f"a filter column is an offset of zero or more into the range, not {self.column!r}.")


@dataclass(frozen=True)
class AutoFilter:
    """A sheet's autofilter: the range it covers and what it keeps."""

    ref: str = ""
    columns: tuple[FilterColumn, ...] = ()

    def column(self, offset: int) -> FilterColumn | None:
        for entry in self.columns:
            if entry.column == offset:
                return entry
        return None

    @property
    def filtering(self) -> bool:
        """Whether any column has a criterion Excel acts on, which is when
        it reports the sheet as filtered."""
        return any(_acts(entry.criterion) for entry in self.columns)


def _acts(criterion: Criterion | None) -> bool:
    if criterion is None:
        return False
    return not (isinstance(criterion, OpaqueCriterion) and criterion.ignored)


def criteria(first: str, second: str | None = None, *, require_all: bool = True) -> Criterion:
    """A criterion from Excel's own criteria strings, stored as Excel stores it.

    ``criteria(">=10")``, ``criteria("=North")``, ``criteria("=a*")`` and
    ``criteria(">5", "<100")`` build what ``Range.AutoFilter`` builds from
    the same ``Criteria1`` and ``Criteria2``: an equality without wildcards
    becomes a value list, trimmed; ``"="`` alone the blanks and ``"<>"``
    alone everything else; a date written ``m/d/yyyy`` or ``yyyy-mm-dd`` in
    an ordering is compared as a date. Two equalities joined by or merge
    into one value list, as Excel merges them.
    """
    one = _parse_criterion(first)
    if second is None:
        return one if isinstance(one, ValueFilter) else CustomFilter((one,))
    other = _parse_criterion(second)
    if not require_all and isinstance(one, ValueFilter) and isinstance(other, ValueFilter):
        return ValueFilter(one.values + other.values, blank=one.blank or other.blank)
    return CustomFilter((_as_comparison(one), _as_comparison(other)), require_all=require_all)


def _parse_criterion(text: str) -> ValueFilter | Comparison:
    operator: FilterOperator = "equal"
    rest = text
    for symbol, name in _SYMBOLS.items():
        if text.startswith(symbol):
            operator, rest = name, text[len(symbol) :]
            break
    if operator == "equal":
        if not rest.strip(" "):
            return ValueFilter(blank=True)
        if _collate.has_wildcards(rest):
            return Comparison("equal", rest)
        return ValueFilter((_collate.unescape(rest).strip(" "),))
    if operator == "notEqual" and not rest:
        # Excel's spelling of "not blank".
        return Comparison("notEqual", " ")
    return Comparison(operator, _typed(rest))


def _typed(text: str) -> ComparisonValue:
    """A criterion value as a number or a date when it reads as one."""
    if _NUMBER.match(text):
        number = float(text)
        return int(number) if number.is_integer() and abs(number) < 1e15 else number
    for pattern in ("%m/%d/%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(text, pattern).date()
        except ValueError:
            continue
    return text


def _as_comparison(part: ValueFilter | Comparison) -> Comparison:
    if isinstance(part, Comparison):
        return part
    if part.blank:
        return Comparison("equal", "")
    return Comparison("equal", part.values[0])


# ---------------------------------------------------------------------------
# Reading
# ---------------------------------------------------------------------------


def read_auto_filter(element: Element) -> AutoFilter:
    """An ``<autoFilter>`` element, read."""
    return AutoFilter(
        ref=element.get("ref") or "",
        columns=tuple(read_filter_column(entry) for entry in element.children_named("filterColumn")),
    )


def read_filter_column(entry: Element) -> FilterColumn:
    """One ``<filterColumn>``, its criterion modelled where it can be."""
    try:
        offset = max(int(entry.get("colId") or "0"), 0)
    except ValueError:
        offset = 0
    return FilterColumn(
        column=offset,
        criterion=_read_criterion(entry),
        hide_dropdown=entry.get("hiddenButton") in ("1", "true"),
    )


def _read_criterion(entry: Element) -> Criterion | None:
    children = list(entry.elements())
    if not children:
        return None
    markup = "".join(child.to_xml() for child in children)
    if len(children) == 1:
        try:
            modelled = _model(children[0])
        except (ValueError, TypeError):
            modelled = None
        if modelled is not None:
            return modelled
    return OpaqueCriterion(markup, ignored=_ignored(children))


def _model(child: Element) -> Criterion | None:
    """The criterion one child element is, or None if it is not modelled."""
    name = child.name.rpartition(":")[2]
    if name == "filters":
        if any(key not in ("blank",) for key in child.attributes):
            return None
        return _value_filter(child)
    if name == "customFilters":
        if any(key != "and" for key in child.attributes):
            return None
        comparisons = [
            Comparison(_operator(one.get("operator")), decode(one.get("val") or ""))
            for one in child.children_named("customFilter")
        ]
        return CustomFilter(tuple(comparisons), require_all=child.get("and") in ("1", "true"))
    if name == "top10":
        count = float(child.get("val") or "nan")
        if not count.is_integer():
            return None
        threshold = child.get("filterVal")
        return Top10Filter(
            count=int(count),
            percent=child.get("percent") in ("1", "true"),
            # Excel writes nothing for the top, which is the default.
            top=child.get("top") not in ("0", "false"),
            threshold=None if threshold is None else float(threshold),
        )
    if name == "dynamicFilter":
        if any(key not in ("type", "val", "maxVal") for key in child.attributes):
            return None
        kind = child.get("type") or ""
        if kind not in DYNAMIC_TYPES:
            return None
        low, high = child.get("val"), child.get("maxVal")
        return DynamicFilter(
            kind,  # type: ignore[arg-type]
            value=None if low is None else float(low),
            maximum=None if high is None else float(high),
        )
    if name == "extLst":
        return _extension_filter(child)
    return None


def _value_filter(container: Element) -> ValueFilter:
    values: list[str] = []
    groups: list[DateGroup] = []
    for one in container.elements():
        kind = one.name.rpartition(":")[2]
        if kind == "filter":
            values.append(decode(one.get("val") or ""))
        elif kind == "dateGroupItem":
            groups.append(_date_group(one))
        else:
            raise ValueError(f"unexpected {one.name} in a value filter")
    return ValueFilter(tuple(values), blank=container.get("blank") in ("1", "true"), date_groups=tuple(groups))


def _date_group(item: Element) -> DateGroup:
    grouping = item.get("dateTimeGrouping") or ""
    if grouping not in GROUPINGS:
        raise ValueError(f"{grouping!r} is not a date grouping")
    fields = [item.get(name) for name, _, _ in _GROUP_FIELDS]
    return DateGroup(grouping, *(None if raw is None else int(raw) for raw in fields))  # type: ignore[arg-type]


def _extension_filter(extensions: Element) -> ValueFilter | None:
    """The value filter Excel keeps in the date-group extension, if that
    is the only thing in the extension list."""
    items = list(extensions.elements())
    if len(items) != 1 or items[0].get("uri") != DATE_GROUP_EXTENSION:
        return None
    columns = list(items[0].elements())
    if len(columns) != 1 or columns[0].name.rpartition(":")[2] != "filterColumn":
        return None
    containers = list(columns[0].elements())
    if len(containers) != 1 or containers[0].name.rpartition(":")[2] != "filters":
        return None
    return _value_filter(containers[0])


def _ignored(children: list[Element]) -> bool:
    """Whether Excel applies the filter as if the column had no criterion:
    an empty value list, or three comparisons or more."""
    if len(children) != 1:
        return False
    child = children[0]
    name = child.name.rpartition(":")[2]
    if name == "filters":
        return not child.attributes and not list(child.elements())
    if name == "customFilters":
        return len(list(child.children_named("customFilter"))) > 2
    return False


def _operator(raw: str | None) -> FilterOperator:
    if raw is None:
        return "equal"
    if raw in OPERATORS:
        return raw  # type: ignore[return-value]
    raise ValueError(f"{raw!r} is not a filter operator")


# ---------------------------------------------------------------------------
# Writing
# ---------------------------------------------------------------------------


def filter_column_element(entry: FilterColumn, *, epoch_1904: bool = False) -> Element:
    """A ``<filterColumn>`` for a column and its criterion."""
    element = Element.create("filterColumn", {"colId": str(entry.column)})
    if entry.hide_dropdown:
        element.set("hiddenButton", "1")
    criterion = entry.criterion
    if criterion is None:
        return element
    if isinstance(criterion, OpaqueCriterion):
        holder = XmlDocument.parse(f"<holder>{criterion.markup}</holder>".encode()).root
        for node in list(holder.children):
            element.append(node)
        return element
    element.append(criterion_element(criterion, epoch_1904=epoch_1904))
    return element


def criterion_element(criterion: Criterion, *, epoch_1904: bool = False) -> Element:
    """The element a modelled criterion is stored as.

    A date group is written as ``<dateGroupItem>`` inside ``<filters>``,
    which Excel reads and applies; Excel itself writes an extension
    instead, and reading accepts both.
    """
    if isinstance(criterion, ValueFilter):
        element = Element.create("filters")
        if criterion.blank:
            element.set("blank", "1")
        for value in criterion.values:
            element.append(Element.create("filter", {"val": encode_attribute(value)}))
        for group in criterion.date_groups:
            attributes = {
                name: str(value)
                for (name, _, _), value in zip(_GROUP_FIELDS, group.fields, strict=True)
                if value is not None
            }
            attributes["dateTimeGrouping"] = group.grouping
            element.append(Element.create("dateGroupItem", attributes))
        return element
    if isinstance(criterion, CustomFilter):
        element = Element.create("customFilters")
        if criterion.require_all:
            element.set("and", "1")
        for comparison in criterion.comparisons:
            attributes = {"val": encode_attribute(stored_value(comparison.value, epoch_1904=epoch_1904))}
            if comparison.operator != "equal":
                attributes = {"operator": comparison.operator, **attributes}
            element.append(Element.create("customFilter", attributes))
        return element
    if isinstance(criterion, Top10Filter):
        attributes: dict[str, str] = {}
        if not criterion.top:
            attributes["top"] = "0"
        if criterion.percent:
            attributes["percent"] = "1"
        attributes["val"] = str(criterion.count)
        if criterion.threshold is not None:
            attributes["filterVal"] = format_number(criterion.threshold)
        return Element.create("top10", attributes)
    if isinstance(criterion, DynamicFilter):
        attributes = {"type": criterion.kind}
        if criterion.value is not None:
            attributes["val"] = format_number(criterion.value)
        if criterion.maximum is not None:
            attributes["maxVal"] = format_number(criterion.maximum)
        return Element.create("dynamicFilter", attributes)
    raise TypeError(f"{criterion!r} has no element of its own; it is written from its markup.")


def stored_value(value: ComparisonValue, *, epoch_1904: bool = False) -> str:
    """A comparison value as the text Excel stores for it."""
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (dt.date, dt.datetime, dt.time)):
        return format_number(datetime_to_serial(value, epoch_1904=epoch_1904))
    return format_number(value)


# ---------------------------------------------------------------------------
# Evaluating
# ---------------------------------------------------------------------------

#: What a cell holds, as a filter needs it: a date is its serial number.
CellScalar = str | float | bool | CellError | None


@dataclass(frozen=True)
class FilterCell:
    """One cell as a filter sees it."""

    value: CellScalar
    #: The text Excel shows for it, fill left out.
    text: str = ""
    #: Whether it is a number shown with a date or time format.
    is_date: bool = False
    #: Whether its value cannot be trusted, such as a formula result that
    #: may be out of date. No criterion decides a row on such a cell.
    unknown: bool = False

    def __post_init__(self) -> None:
        # One numeric type, so a comparison need not ask which.
        if isinstance(self.value, int) and not isinstance(self.value, bool):
            object.__setattr__(self, "value", float(self.value))


#: Per cell: keep, hide, or cannot say.
Verdict = bool | None


@dataclass(frozen=True)
class FilterOutcome:
    """What applying a filter did to the rows of its range."""

    #: Rows a criterion excludes, now hidden.
    hidden: tuple[int, ...] = ()
    #: Rows every criterion keeps, now showing.
    shown: tuple[int, ...] = ()
    #: Rows only a criterion this library cannot evaluate could exclude,
    #: or whose deciding cell holds a formula result that may be out of
    #: date. They are left hidden or shown as they were.
    undecided: tuple[int, ...] = ()
    #: Offsets of the columns whose criterion could not be evaluated for
    #: at least one row.
    unevaluated: tuple[int, ...] = ()


def keeps(
    criterion: Criterion | None,
    cells: Sequence[FilterCell],
    *,
    today: dt.date,
    epoch_1904: bool = False,
) -> list[Verdict]:
    """Whether a criterion keeps each cell's row, applied as Excel applies it."""
    if criterion is None:
        return [True] * len(cells)
    if isinstance(criterion, OpaqueCriterion):
        return [True if criterion.ignored else None] * len(cells)
    if isinstance(criterion, ValueFilter):
        return [None if cell.unknown else _in_list(criterion, cell, epoch_1904) for cell in cells]
    if isinstance(criterion, CustomFilter):
        texts = [stored_value(one.value, epoch_1904=epoch_1904) for one in criterion.comparisons]
        verdicts: list[Verdict] = []
        for cell in cells:
            if cell.unknown:
                verdicts.append(None)
                continue
            results = [
                _compares(one.operator, text, cell)
                for one, text in zip(criterion.comparisons, texts, strict=True)
            ]
            verdicts.append(all(results) if criterion.require_all else any(results))
        return verdicts
    if isinstance(criterion, Top10Filter):
        return _top(criterion, cells)
    return _dynamic(criterion, cells, today=today, epoch_1904=epoch_1904)


def decide(verdicts: Sequence[Sequence[Verdict]]) -> list[Verdict]:
    """Rows across every column: hidden if any criterion hides it, shown if
    every criterion keeps it, and undecided if only an unevaluated one
    could hide it."""
    decided: list[Verdict] = []
    for row in zip(*verdicts, strict=True):
        if any(verdict is False for verdict in row):
            decided.append(False)
        elif any(verdict is None for verdict in row):
            decided.append(None)
        else:
            decided.append(True)
    return decided


def resolve(
    criterion: Criterion | None,
    cells: Sequence[FilterCell],
    *,
    today: dt.date,
    epoch_1904: bool = False,
    strict: bool = True,
) -> Criterion | None:
    """A criterion with what Excel works out when it applies one filled
    in: a top ten's threshold, an average, a date period's bounds.

    Excel's object model refuses a top ten or an average over a column
    holding an error, so that is refused here, before anything is written.
    Re-applying one already stored is not refused, as Excel does not
    refuse it: ``strict=False`` leaves such a criterion as it was.
    """
    aggregate = isinstance(criterion, Top10Filter) or (
        isinstance(criterion, DynamicFilter) and criterion.kind in AVERAGES
    )
    if aggregate and any(isinstance(cell.value, CellError) for cell in cells):
        if not strict:
            return criterion
        raise ValueError(
            "Excel refuses a top ten or an average over a column holding an error; clear the error first."
        )
    if isinstance(criterion, Top10Filter):
        threshold = _threshold(criterion, [cell for cell in cells if not cell.unknown])
        return Top10Filter(criterion.count, criterion.percent, criterion.top, threshold)
    if isinstance(criterion, DynamicFilter):
        if criterion.kind in AVERAGES:
            numbers = _numbers(cell for cell in cells if not cell.unknown)
            mean = _mean(numbers)
            return DynamicFilter(criterion.kind, mean)
        if criterion.kind in RELATIVE_PERIODS:
            low, high = period_bounds(criterion.kind, today)
            return DynamicFilter(
                criterion.kind,
                datetime_to_serial(low, epoch_1904=epoch_1904),
                datetime_to_serial(high, epoch_1904=epoch_1904),
            )
    return criterion


def is_blank(cell: FilterCell) -> bool:
    """Excel's blanks: an empty cell, and text that is empty or all spaces."""
    value = cell.value
    return value is None or (isinstance(value, str) and not value.strip(" "))


def _numbers(cells: Iterable[FilterCell]) -> list[float]:
    return [cell.value for cell in cells if isinstance(cell.value, float)]


def _mean(numbers: Sequence[float]) -> float | None:
    """The average of a column, added in order. Python's ``sum`` compensates
    for rounding from 3.12 on, so a mean taken with it, and the rows compared
    against that mean, would depend on the interpreter."""
    if not numbers:
        return None
    total = 0.0
    for number in numbers:
        total += number
    return total / len(numbers)


def _shown(cell: FilterCell) -> str:
    """The text a filter matches: what the cell shows, outer spaces trimmed."""
    return cell.text.strip(" ")


def _in_list(criterion: ValueFilter, cell: FilterCell, epoch_1904: bool) -> bool:
    if criterion.blank and is_blank(cell):
        return True
    if criterion.date_groups and cell.is_date and isinstance(cell.value, float):
        # Once a list has date groups, a date is theirs alone: its text
        # matching a value keeps nothing, as Excel's dropdown shows dates
        # as a tree and everything else as a list.
        parts = date_parts(cell.value, epoch_1904=epoch_1904)
        return parts is not None and any(group.contains(parts) for group in criterion.date_groups)
    shown = _shown(cell)
    return bool(shown) and any(_collate.equal(value, shown) for value in criterion.values)


def _held(number: float) -> float:
    """A number as Excel holds it, fifteen significant digits."""
    return float(format(number, ".15g"))


def _compares(operator: FilterOperator, stored: str, cell: FilterCell) -> bool:
    """Whether one stored comparison keeps a cell."""
    if stored == "":
        return is_blank(cell)
    if stored == " ":
        return not is_blank(cell)
    value = cell.value
    if _collate.has_wildcards(stored):
        if not isinstance(value, str):
            return operator == "notEqual"
        # A match counts as equal and a miss as greater, never less.
        return _holds(operator, 0 if _collate.wildcard_match(stored, _shown(cell)) else 1)
    if operator == "equal":
        return _collate.equal(_collate.unescape(stored), _shown(cell))
    number = float(stored) if _NUMBER.match(stored) else None
    boolean = stored.upper() if stored.upper() in ("TRUE", "FALSE") else None
    if operator == "notEqual":
        if number is not None:
            return not (isinstance(value, float) and _held(value) == _held(number))
        if boolean is not None:
            return not (isinstance(value, bool) and value == (boolean == "TRUE"))
        return not (isinstance(value, str) and _collate.equal(_collate.unescape(stored), _shown(cell)))
    if number is not None:
        if not isinstance(value, float):
            return False
        left, right = _held(value), _held(number)
        return _holds(operator, (left > right) - (left < right))
    if boolean is not None:
        if not isinstance(value, bool):
            return False
        return _holds(operator, int(value) - int(boolean == "TRUE"))
    if not isinstance(value, str):
        return False
    return _holds(operator, _collate.compare(_shown(cell), stored))


def _holds(operator: FilterOperator, sign: int) -> bool:
    """Whether an operator holds, given how the cell compares with the value."""
    if operator == "equal":
        return sign == 0
    if operator == "notEqual":
        return sign != 0
    if operator == "greaterThan":
        return sign > 0
    if operator == "greaterThanOrEqual":
        return sign >= 0
    if operator == "lessThan":
        return sign < 0
    return sign <= 0


def _threshold(criterion: Top10Filter, cells: Sequence[FilterCell]) -> float | None:
    numbers = sorted(_numbers(cells), reverse=criterion.top)
    if not numbers:
        return None
    wanted = criterion.count
    if criterion.percent:
        wanted = max(1, len(numbers) * criterion.count // 100)
    return numbers[min(wanted, len(numbers)) - 1]


def _top(criterion: Top10Filter, cells: Sequence[FilterCell]) -> list[Verdict]:
    if any(isinstance(cell.value, CellError) for cell in cells):
        # Excel cannot rank past an error and keeps every row.
        return [True] * len(cells)
    if any(cell.unknown for cell in cells):
        return [None] * len(cells)
    threshold = _threshold(criterion, cells)
    verdicts: list[Verdict] = []
    for cell in cells:
        value = cell.value
        if threshold is None or not isinstance(value, float):
            verdicts.append(False)
        else:
            verdicts.append(value >= threshold if criterion.top else value <= threshold)
    return verdicts


def _dynamic(
    criterion: DynamicFilter, cells: Sequence[FilterCell], *, today: dt.date, epoch_1904: bool
) -> list[Verdict]:
    kind = criterion.kind
    if kind in AVERAGES:
        if any(isinstance(cell.value, CellError) for cell in cells):
            # Excel cannot average past an error and uses what was stored.
            mean: float | None = criterion.value or 0.0
        elif any(cell.unknown for cell in cells):
            return [None] * len(cells)
        else:
            numbers = _numbers(cells)
            mean = _mean(numbers)
        verdicts: list[Verdict] = []
        for cell in cells:
            value = cell.value
            if mean is None or not isinstance(value, float):
                verdicts.append(False)
            else:
                verdicts.append(value > mean if kind == "aboveAverage" else value < mean)
        return verdicts

    if kind in RELATIVE_PERIODS:
        low, high = period_bounds(kind, today)
        low_serial = datetime_to_serial(low, epoch_1904=epoch_1904)
        high_serial = datetime_to_serial(high, epoch_1904=epoch_1904)

        def inside(value: float) -> bool:
            return low_serial <= value < high_serial
    else:
        months = _period_months(kind)

        def inside(value: float) -> bool:
            parts = date_parts(value, epoch_1904=epoch_1904)
            return parts is not None and parts[1] in months

    return [
        None if cell.unknown else (cell.is_date and isinstance(cell.value, float) and inside(cell.value))
        for cell in cells
    ]


def _period_months(kind: str) -> frozenset[int]:
    if kind.startswith("Q"):
        quarter = int(kind[1:])
        return frozenset(range(3 * quarter - 2, 3 * quarter + 1))
    return frozenset({int(kind[1:])})


def period_bounds(kind: DynamicType, today: dt.date) -> tuple[dt.date, dt.date]:
    """The first day of a relative date period and the day after its last,
    measured from ``today``. Weeks start on Sunday."""
    one_day = dt.timedelta(days=1)
    if kind in ("today", "yesterday", "tomorrow"):
        day = today + {"today": 0, "yesterday": -1, "tomorrow": 1}[kind] * one_day
        return day, day + one_day
    if kind.endswith("Week"):
        sunday = today - dt.timedelta(days=(today.weekday() + 1) % 7)
        start = sunday + {"this": 0, "last": -7, "next": 7}[kind[:4]] * one_day
        return start, start + 7 * one_day
    if kind == "yearToDate":
        return dt.date(today.year, 1, 1), today + one_day
    shift = {"this": 0, "last": -1, "next": 1}[kind[:4]]
    if kind.endswith("Month"):
        return _month_start(today.year, today.month + shift), _month_start(today.year, today.month + shift + 1)
    if kind.endswith("Quarter"):
        first = 3 * ((today.month - 1) // 3) + 1 + 3 * shift
        return _month_start(today.year, first), _month_start(today.year, first + 3)
    year = today.year + shift
    return dt.date(year, 1, 1), dt.date(year + 1, 1, 1)


def _month_start(year: int, month: int) -> dt.date:
    """The first of a month counted from January of a year, carrying over."""
    carry, index = divmod(month - 1, 12)
    return dt.date(year + carry, index + 1, 1)


__all__ = [
    "AVERAGES",
    "DATE_GROUP_EXTENSION",
    "DYNAMIC_TYPES",
    "GROUPINGS",
    "OPERATORS",
    "RELATIVE_PERIODS",
    "AutoFilter",
    "CellScalar",
    "Comparison",
    "ComparisonValue",
    "Criterion",
    "CustomFilter",
    "DateGroup",
    "DateGrouping",
    "DynamicFilter",
    "DynamicType",
    "FilterCell",
    "FilterColumn",
    "FilterOperator",
    "FilterOutcome",
    "OpaqueCriterion",
    "Top10Filter",
    "ValueFilter",
    "Verdict",
    "criteria",
    "criterion_element",
    "decide",
    "filter_column_element",
    "is_blank",
    "keeps",
    "period_bounds",
    "read_auto_filter",
    "read_filter_column",
    "resolve",
    "stored_value",
]
