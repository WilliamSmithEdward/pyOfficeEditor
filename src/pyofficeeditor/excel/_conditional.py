"""Conditional formatting: the rules, and the ranges they cover.

A rule is a ``<cfRule>`` inside a ``<conditionalFormatting sqref="...">``.
The formatting it paints is not inline: ``dxfId`` indexes the workbook's
``dxfs`` table, which :mod:`pyofficeeditor.excel._dxf` handles.

**The compatibility formula is load-bearing.** Excel writes a ``<formula>``
beside the rules whose condition is expressed in attributes, such as
``containsText`` with ``operator`` and ``text``. It looks redundant, and it
is not: a ``containsText`` rule written without it sits in the file, opens
with no repair, is counted by ``FormatConditions.Count``, and never fires.
Measured against Excel by writing the same rule twice, once each way, and
reading ``DisplayFormat.Interior.Color`` back: the one with the formula
painted its cell and the one without left it white. So every template below
was copied from Excel's own output rather than reconstructed.

**The formula anchors to the top-left of the first area.** ``I7:I9`` gets
``...,I7)``, and a multi-area ``K3:K5 M8:M9`` gets ``...,K3)``, not the
top-left of the bounding box. The reference is relative, so Excel offsets it
for every other cell in the range.

**Text is escaped twice, differently.** The ``text`` attribute is XML, so a
quote in the search term is written ``&quot;``. Inside the formula the term
is a string literal, so the same quote is doubled instead. Searching for
``a"b`` gives the attribute ``text="a&quot;b"`` beside the formula
``SEARCH("a""b",H2)``.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import Literal

from pyofficeeditor._xml import Element, XmlDocument
from pyofficeeditor.excel._addresses import parse_sqref
from pyofficeeditor.excel._formats import Color
from pyofficeeditor.excel._reference import CellRef, RangeRef

#: ``ST_CfType``.
RuleType = Literal[
    "expression",
    "cellIs",
    "colorScale",
    "dataBar",
    "iconSet",
    "top10",
    "uniqueValues",
    "duplicateValues",
    "containsText",
    "notContainsText",
    "beginsWith",
    "endsWith",
    "containsBlanks",
    "notContainsBlanks",
    "containsErrors",
    "notContainsErrors",
    "timePeriod",
    "aboveAverage",
]

#: ``ST_ConditionalFormattingOperator``.
Operator = Literal[
    "lessThan",
    "lessThanOrEqual",
    "equal",
    "notEqual",
    "greaterThanOrEqual",
    "greaterThan",
    "between",
    "notBetween",
    "containsText",
    "notContains",
    "beginsWith",
    "endsWith",
]

#: ``ST_TimePeriod``.
TimePeriod = Literal[
    "today",
    "yesterday",
    "tomorrow",
    "last7Days",
    "thisWeek",
    "lastWeek",
    "nextWeek",
    "thisMonth",
    "lastMonth",
    "nextMonth",
]

#: The operator each text rule type writes. The two differ for "does not
#: contain": the type is ``notContainsText`` and the operator ``notContains``.
TEXT_OPERATORS: dict[str, Operator] = {
    "containsText": "containsText",
    "notContainsText": "notContains",
    "beginsWith": "beginsWith",
    "endsWith": "endsWith",
}

#: Excel's own compatibility formulas for the text rules, with ``{anchor}``
#: for the top-left cell and ``{text}`` for the search term with its quotes
#: already doubled.
_TEXT_FORMULAS: dict[str, str] = {
    "containsText": 'NOT(ISERROR(SEARCH("{text}",{anchor})))',
    "notContainsText": 'ISERROR(SEARCH("{text}",{anchor}))',
    "beginsWith": 'LEFT({anchor},LEN("{text}"))="{text}"',
    "endsWith": 'RIGHT({anchor},LEN("{text}"))="{text}"',
}

#: The same, for the rules that take no operand.
_PRESENCE_FORMULAS: dict[str, str] = {
    "containsBlanks": "LEN(TRIM({anchor}))=0",
    "notContainsBlanks": "LEN(TRIM({anchor}))>0",
    "containsErrors": "ISERROR({anchor})",
    "notContainsErrors": "NOT(ISERROR({anchor}))",
}

#: And for every time period. These are Excel's, verbatim, including the
#: ``0-1``/``0+1`` in the month ones and the mix of FLOOR and ROUNDDOWN.
_TIME_PERIOD_FORMULAS: dict[str, str] = {
    "today": "FLOOR({anchor},1)=TODAY()",
    "yesterday": "FLOOR({anchor},1)=TODAY()-1",
    "tomorrow": "FLOOR({anchor},1)=TODAY()+1",
    "last7Days": "AND(TODAY()-FLOOR({anchor},1)<=6,FLOOR({anchor},1)<=TODAY())",
    "thisWeek": (
        "AND(TODAY()-ROUNDDOWN({anchor},0)<=WEEKDAY(TODAY())-1,"
        "ROUNDDOWN({anchor},0)-TODAY()<=7-WEEKDAY(TODAY()))"
    ),
    "lastWeek": (
        "AND(TODAY()-ROUNDDOWN({anchor},0)>=(WEEKDAY(TODAY())),"
        "TODAY()-ROUNDDOWN({anchor},0)<(WEEKDAY(TODAY())+7))"
    ),
    "nextWeek": (
        "AND(ROUNDDOWN({anchor},0)-TODAY()>(7-WEEKDAY(TODAY())),"
        "ROUNDDOWN({anchor},0)-TODAY()<(15-WEEKDAY(TODAY())))"
    ),
    "thisMonth": "AND(MONTH({anchor})=MONTH(TODAY()),YEAR({anchor})=YEAR(TODAY()))",
    "lastMonth": (
        "AND(MONTH({anchor})=MONTH(EDATE(TODAY(),0-1)),"
        "YEAR({anchor})=YEAR(EDATE(TODAY(),0-1)))"
    ),
    "nextMonth": (
        "AND(MONTH({anchor})=MONTH(EDATE(TODAY(),0+1)),"
        "YEAR({anchor})=YEAR(EDATE(TODAY(),0+1)))"
    ),
}

#: Rule types whose ``<formula>`` states the condition itself rather than
#: restating an attribute. These are never regenerated.
CONDITION_FORMULA_TYPES = frozenset({"cellIs", "expression"})


def compatibility_formula(
    rule_type: str, anchor: CellRef, *, text: str | None, time_period: str | None
) -> str | None:
    """The ``<formula>`` Excel writes beside an attribute-driven rule.

    ``None`` for a rule that needs none. See the module docstring: leaving
    this out gives a rule that is present, valid and inert.
    """
    reference = anchor.a1
    if rule_type in _TEXT_FORMULAS:
        if text is None:
            return None
        return _TEXT_FORMULAS[rule_type].format(anchor=reference, text=_quote(text))
    if rule_type in _PRESENCE_FORMULAS:
        return _PRESENCE_FORMULAS[rule_type].format(anchor=reference)
    if rule_type == "timePeriod" and time_period is not None:
        template = _TIME_PERIOD_FORMULAS.get(time_period)
        return None if template is None else template.format(anchor=reference)
    return None


def _quote(text: str) -> str:
    """A search term as it appears inside a formula's string literal.

    Excel doubles an embedded quote there, which is a different escaping
    from the XML the ``text`` attribute needs.
    """
    return text.replace('"', '""')


@dataclass(frozen=True)
class Cfvo:
    """A threshold on a colour scale, data bar or icon set.

    ``kind`` is ``min``, ``max``, ``num``, ``percent``, ``percentile`` or
    ``formula``. ``min`` and ``max`` carry no value.
    """

    kind: str
    value: str | None = None
    greater_than_or_equal: bool | None = None

    @classmethod
    def read(cls, element: Element) -> Cfvo:
        raw = element.get("gte")
        return cls(
            kind=element.get("type") or "min",
            value=element.get("val"),
            greater_than_or_equal=None if raw is None else raw in ("1", "true"),
        )

    def write(self) -> Element:
        element = Element.create("cfvo", {"type": self.kind})
        if self.value is not None:
            element.set("val", self.value)
        if self.greater_than_or_equal is not None:
            element.set("gte", "1" if self.greater_than_or_equal else "0")
        return element


@dataclass(frozen=True)
class ColorScale:
    """A two- or three-colour gradient."""

    values: tuple[Cfvo, ...] = ()
    colors: tuple[Color, ...] = ()

    @classmethod
    def three(cls, low: str = "F8696B", middle: str = "FFEB84", high: str = "63BE7B") -> ColorScale:
        """Excel's own default three-colour scale."""
        return cls(
            values=(Cfvo("min"), Cfvo("percentile", "50"), Cfvo("max")),
            colors=tuple(Color.from_rgb(c) for c in (low, middle, high)),
        )

    @classmethod
    def two(cls, low: str = "FF7128", high: str = "FFEF9C") -> ColorScale:
        return cls(
            values=(Cfvo("min"), Cfvo("max")),
            colors=tuple(Color.from_rgb(c) for c in (low, high)),
        )

    @classmethod
    def read(cls, element: Element) -> ColorScale:
        return cls(
            values=tuple(Cfvo.read(node) for node in element.children_named("cfvo")),
            colors=tuple(
                color
                for node in element.children_named("color")
                if (color := Color.read(node)) is not None
            ),
        )

    def write(self) -> Element:
        element = Element.create("colorScale")
        for value in self.values:
            element.append(value.write())
        for color in self.colors:
            element.append(color.write())
        return element


@dataclass(frozen=True)
class DataBar:
    """A bar drawn inside the cell.

    Excel pairs a modern data bar with an ``x14`` twin in the sheet's
    ``extLst``, carrying its own copy of the range. This library writes the
    plain form, which older and newer Excel both render.
    """

    minimum: Cfvo = field(default_factory=lambda: Cfvo("min"))
    maximum: Cfvo = field(default_factory=lambda: Cfvo("max"))
    color: Color = field(default_factory=lambda: Color.from_rgb("638EC6"))
    show_value: bool = True

    @classmethod
    def read(cls, element: Element) -> DataBar:
        values = [Cfvo.read(node) for node in element.children_named("cfvo")]
        color = Color.read(element.child("color"))
        raw = element.get("showValue")
        return cls(
            minimum=values[0] if values else Cfvo("min"),
            maximum=values[1] if len(values) > 1 else Cfvo("max"),
            color=color if color is not None else Color.from_rgb("638EC6"),
            show_value=True if raw is None else raw in ("1", "true"),
        )

    def write(self) -> Element:
        element = Element.create("dataBar")
        if not self.show_value:
            element.set("showValue", "0")
        element.append(self.minimum.write())
        element.append(self.maximum.write())
        element.append(self.color.write())
        return element


@dataclass(frozen=True)
class IconSet:
    """Arrows, traffic lights and the rest."""

    icons: str = "3TrafficLights1"
    values: tuple[Cfvo, ...] = ()
    show_value: bool = True
    reverse: bool = False

    @classmethod
    def three(cls, icons: str = "3TrafficLights1") -> IconSet:
        return cls(
            icons=icons,
            values=(Cfvo("percent", "0"), Cfvo("percent", "33"), Cfvo("percent", "67")),
        )

    @classmethod
    def read(cls, element: Element) -> IconSet:
        show = element.get("showValue")
        reverse = element.get("reverse")
        return cls(
            icons=element.get("iconSet") or "3TrafficLights1",
            values=tuple(Cfvo.read(node) for node in element.children_named("cfvo")),
            show_value=True if show is None else show in ("1", "true"),
            reverse=reverse in ("1", "true"),
        )

    def write(self) -> Element:
        element = Element.create("iconSet")
        if self.icons != "3TrafficLights1":
            element.set("iconSet", self.icons)
        if not self.show_value:
            element.set("showValue", "0")
        if self.reverse:
            element.set("reverse", "1")
        for value in self.values:
            element.append(value.write())
        return element


@dataclass(frozen=True)
class ConditionalRule:
    """One ``<cfRule>``.

    Mirrors ``CT_CfRule``, which is itself one element with many optional
    attributes, rather than splitting into a class per rule type: the file
    format does not make that distinction and a round trip has to reproduce
    whatever combination it finds.
    """

    kind: RuleType
    priority: int = 1
    dxf_id: int | None = None
    stop_if_true: bool = False
    operator: Operator | None = None
    text: str | None = None
    time_period: TimePeriod | None = None
    rank: int | None = None
    percent: bool = False
    bottom: bool = False
    above_average: bool | None = None
    equal_average: bool = False
    std_dev: int | None = None
    formulas: tuple[str, ...] = ()
    color_scale: ColorScale | None = None
    data_bar: DataBar | None = None
    icon_set: IconSet | None = None
    #: The rule's own ``<extLst>``, kept as source text.
    #:
    #: A data bar Excel wrote carries one holding an ``x14:id`` GUID, which
    #: names the ``x14:cfRule`` twin in the sheet's ``extLst``. Dropping it
    #: would leave that twin pointing at a rule with no id, so it is carried
    #: through verbatim rather than modelled: this library does not author
    #: x14 rules and has no reason to rewrite one.
    extension: str | None = None

    @classmethod
    def read(cls, element: Element) -> ConditionalRule:
        scale = element.child("colorScale")
        bar = element.child("dataBar")
        icons = element.child("iconSet")
        extension = element.child("extLst")
        above = element.get("aboveAverage")
        return cls(
            kind=_as_rule_type(element.get("type")),
            priority=_as_int(element.get("priority")) or 1,
            dxf_id=_as_int(element.get("dxfId")),
            stop_if_true=element.get("stopIfTrue") in ("1", "true"),
            operator=_as_operator(element.get("operator")),
            text=element.get("text"),
            time_period=_as_time_period(element.get("timePeriod")),
            rank=_as_int(element.get("rank")),
            percent=element.get("percent") in ("1", "true"),
            bottom=element.get("bottom") in ("1", "true"),
            above_average=None if above is None else above in ("1", "true"),
            equal_average=element.get("equalAverage") in ("1", "true"),
            std_dev=_as_int(element.get("stdDev")),
            formulas=tuple(node.text or "" for node in element.children_named("formula")),
            color_scale=None if scale is None else ColorScale.read(scale),
            data_bar=None if bar is None else DataBar.read(bar),
            icon_set=None if icons is None else IconSet.read(icons),
            extension=None if extension is None else extension.to_xml(),
        )

    def write(self) -> Element:
        # Attributes go in the order CT_CfRule declares them, which is also
        # the order Excel emits: type, dxfId, priority, stopIfTrue,
        # aboveAverage, percent, bottom, operator, text, timePeriod, rank,
        # stdDev, equalAverage.
        element = Element.create("cfRule", {"type": self.kind})
        if self.dxf_id is not None:
            element.set("dxfId", str(self.dxf_id))
        element.set("priority", str(self.priority))
        if self.stop_if_true:
            element.set("stopIfTrue", "1")
        if self.above_average is not None and not self.above_average:
            element.set("aboveAverage", "0")
        if self.percent:
            element.set("percent", "1")
        if self.bottom:
            element.set("bottom", "1")
        if self.operator is not None:
            element.set("operator", self.operator)
        if self.text is not None:
            element.set("text", self.text)
        if self.time_period is not None:
            element.set("timePeriod", self.time_period)
        if self.rank is not None:
            element.set("rank", str(self.rank))
        if self.std_dev is not None:
            element.set("stdDev", str(self.std_dev))
        if self.equal_average:
            element.set("equalAverage", "1")
        for formula in self.formulas:
            node = Element.create("formula")
            node.set_text(formula)
            element.append(node)
        for child in (self.color_scale, self.data_bar, self.icon_set):
            if child is not None:
                element.append(child.write())
        if self.extension is not None:
            element.append(XmlDocument.parse(self.extension.encode()).root)
        return element

    def anchored_at(self, anchor: CellRef) -> ConditionalRule:
        """The same rule with its compatibility formula rebuilt for a range.

        A rule whose ``<formula>`` states its own condition, ``cellIs`` and
        ``expression``, is returned untouched: that formula is the rule.
        """
        if self.kind in CONDITION_FORMULA_TYPES:
            return self
        formula = compatibility_formula(
            self.kind, anchor, text=self.text, time_period=self.time_period
        )
        if formula is None:
            return self if not self.formulas else replace(self, formulas=())
        return replace(self, formulas=(formula,))


@dataclass(frozen=True)
class ConditionalFormatting:
    """One ``<conditionalFormatting>``: a set of ranges and the rules on it.

    ``sqref`` holds several ranges separated by spaces, which is how Excel
    records a rule applied to a multi-area selection.
    """

    ranges: tuple[RangeRef, ...] = ()
    rules: tuple[ConditionalRule, ...] = ()

    @property
    def anchor(self) -> CellRef:
        """The cell a compatibility formula points at: the top-left of the
        first range, not of the bounding box."""
        if not self.ranges:
            return CellRef(1, 1)
        return self.ranges[0].start

    @property
    def sqref(self) -> str:
        return " ".join(block.a1 for block in self.ranges)

    @classmethod
    def read(cls, element: Element) -> ConditionalFormatting:
        return cls(
            ranges=parse_sqref(element.get("sqref") or ""),
            rules=tuple(ConditionalRule.read(node) for node in element.children_named("cfRule")),
        )

    def write(self) -> Element:
        element = Element.create("conditionalFormatting", {"sqref": self.sqref})
        for rule in self.rules:
            element.append(rule.write())
        return element


# ----------------------------------------------------------------------
# Constructors
#
# Module functions rather than classmethods: the fields mirror CT_CfRule,
# so ``color_scale``, ``above_average`` and ``time_period`` are already
# taken, and a constructor sharing a name with a field would shadow it.
# ----------------------------------------------------------------------


def cell_is(operator: Operator, value: object, other: object | None = None) -> ConditionalRule:
    """Compare the cell against one value, or two for ``between``.

    The values become the rule's formulas, so a string that is not a number
    must be quoted the way a formula quotes it.
    """
    formulas = [_operand(value)]
    if other is not None:
        formulas.append(_operand(other))
    if operator in ("between", "notBetween") and len(formulas) != 2:
        raise ValueError(f"{operator!r} needs two values, not one.")
    return ConditionalRule(kind="cellIs", operator=operator, formulas=tuple(formulas))


def expression(formula: str) -> ConditionalRule:
    """Format where a formula is true.

    The formula is written relative to the top-left of the range, so anchor
    the parts that must not move: ``=$A1>5`` follows the row and holds the
    column.
    """
    return ConditionalRule(kind="expression", formulas=(formula.lstrip("="),))


def contains_text(text: str) -> ConditionalRule:
    return _text_rule("containsText", text)


def not_contains_text(text: str) -> ConditionalRule:
    return _text_rule("notContainsText", text)


def begins_with(text: str) -> ConditionalRule:
    return _text_rule("beginsWith", text)


def ends_with(text: str) -> ConditionalRule:
    return _text_rule("endsWith", text)


def is_blank(*, blank: bool = True) -> ConditionalRule:
    return ConditionalRule(kind="containsBlanks" if blank else "notContainsBlanks")


def is_error(*, error: bool = True) -> ConditionalRule:
    return ConditionalRule(kind="containsErrors" if error else "notContainsErrors")


def duplicates() -> ConditionalRule:
    return ConditionalRule(kind="duplicateValues")


def uniques() -> ConditionalRule:
    return ConditionalRule(kind="uniqueValues")


def top(rank: int = 10, *, percent: bool = False, bottom: bool = False) -> ConditionalRule:
    """The highest ``rank`` values, or the lowest with ``bottom=True``."""
    if rank < 1:
        raise ValueError(f"rank {rank} is not a count; it has to be at least 1.")
    return ConditionalRule(kind="top10", rank=rank, percent=percent, bottom=bottom)


def average(
    *, above: bool = True, or_equal: bool = False, standard_deviations: int | None = None
) -> ConditionalRule:
    """Above or below the range's average.

    ``standard_deviations`` narrows it to values that far out, which is what
    Excel's "1 std dev above" offers.
    """
    return ConditionalRule(
        kind="aboveAverage",
        # The attribute is written only when it is false; absent means above,
        # which is the default Excel relies on.
        above_average=None if above else False,
        equal_average=or_equal,
        std_dev=standard_deviations,
    )


def during(period: TimePeriod) -> ConditionalRule:
    """A date rule: today, last week, next month and the rest."""
    if period not in _TIME_PERIOD_FORMULAS:
        raise ValueError(
            f"{period!r} is not a time period Excel has; expected one of "
            f"{', '.join(sorted(_TIME_PERIOD_FORMULAS))}."
        )
    return ConditionalRule(kind="timePeriod", time_period=period)


def gradient(scale: ColorScale | None = None) -> ConditionalRule:
    """A colour scale. Defaults to Excel's own three-colour one."""
    return ConditionalRule(kind="colorScale", color_scale=scale or ColorScale.three())


def bar(data_bar: DataBar | None = None) -> ConditionalRule:
    """A data bar drawn in the cell."""
    return ConditionalRule(kind="dataBar", data_bar=data_bar or DataBar())


def icons(icon_set: IconSet | None = None) -> ConditionalRule:
    """An icon set. Defaults to three traffic lights."""
    return ConditionalRule(kind="iconSet", icon_set=icon_set or IconSet.three())


def _text_rule(kind: RuleType, text: str) -> ConditionalRule:
    return ConditionalRule(kind=kind, operator=TEXT_OPERATORS[kind], text=text)


def _operand(value: object) -> str:
    """A comparison value as formula text.

    A number goes in bare and a string is quoted, because ``cellIs`` puts
    its operand in a ``<formula>`` where bare text would be read as a name.
    A string that already looks like a formula is left alone.
    """
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        return repr(value) if isinstance(value, float) else str(value)
    text = str(value)
    if text.startswith("="):
        return text[1:]
    if text.startswith('"') and text.endswith('"'):
        return text
    return '"' + text.replace('"', '""') + '"'


def _as_int(raw: str | None) -> int | None:
    if raw is None:
        return None
    try:
        return int(raw)
    except ValueError:
        return None


def _as_rule_type(raw: str | None) -> RuleType:
    # An unknown type is kept as written rather than refused: the rule is
    # copied back unchanged and Excel decides what it means.
    return raw if raw else "expression"  # type: ignore[return-value]


def _as_operator(raw: str | None) -> Operator | None:
    return raw  # type: ignore[return-value]


def _as_time_period(raw: str | None) -> TimePeriod | None:
    return raw  # type: ignore[return-value]


__all__ = [
    "CONDITION_FORMULA_TYPES",
    "TEXT_OPERATORS",
    "Cfvo",
    "ColorScale",
    "ConditionalFormatting",
    "ConditionalRule",
    "DataBar",
    "IconSet",
    "Operator",
    "RuleType",
    "TimePeriod",
    "average",
    "bar",
    "begins_with",
    "cell_is",
    "compatibility_formula",
    "contains_text",
    "duplicates",
    "during",
    "ends_with",
    "expression",
    "gradient",
    "icons",
    "is_blank",
    "is_error",
    "not_contains_text",
    "parse_sqref",
    "top",
    "uniques",
]
