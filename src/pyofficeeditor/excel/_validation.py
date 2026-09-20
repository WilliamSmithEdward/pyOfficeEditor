"""Data validation: what a cell will accept, and what it says when it won't.

A ``<dataValidation>`` names its ranges in an ``sqref`` and its bounds in one
or two ``<formula>`` children, the same shape a conditional format uses. Four
things about it are measured from Excel rather than read off the schema, and
each is a way to produce a file that opens and behaves wrong.

**``showDropDown`` is inverted.** The attribute reads like "show the
dropdown" and means the opposite: Excel writes ``showDropDown="1"`` for a
list whose in-cell dropdown is turned *off*, and writes no attribute at all
for the ordinary list that has one. Measured by setting
``Validation.InCellDropdown`` both ways and reading the part back. This class
calls the field :attr:`hide_dropdown` so the name says what the value does.

**A literal list is a quoted string.** ``<formula1>"red,green,blue"</formula1>``
with the quotes inside the text, against a bare ``$A$1:$A$2`` when the list
points at a range. Writing the items unquoted gives Excel a formula that
names three cells it cannot find.

**A date or a time is a serial.** ``2026-01-01`` is stored as ``46023`` and
``09:00`` as ``0.375``, so a validation built from a :class:`datetime.date`
converts on the way in exactly as a cell value does.

**The defaults are absent, not written.** ``errorStyle`` is missing for
``stop`` and ``operator`` is missing for ``between``, so a reader that
expects them present reports the wrong rule for every validation Excel
wrote with its defaults.
"""

from __future__ import annotations

import datetime as dt
from collections.abc import Sequence
from dataclasses import dataclass, replace
from typing import Literal

from pyofficeeditor._xml import Element
from pyofficeeditor.excel._addresses import parse_sqref
from pyofficeeditor.excel._reference import RangeRef
from pyofficeeditor.excel._values import datetime_to_serial

#: ``ST_DataValidationType``.
ValidationType = Literal[
    "none", "whole", "decimal", "list", "date", "time", "textLength", "custom"
]

#: ``ST_DataValidationOperator``. ``between`` is the default and is written
#: only when it is not the one in use.
ValidationOperator = Literal[
    "between",
    "notBetween",
    "equal",
    "notEqual",
    "lessThan",
    "lessThanOrEqual",
    "greaterThan",
    "greaterThanOrEqual",
]

#: ``ST_DataValidationErrorStyle``. ``stop`` is the default and unwritten.
ErrorStyle = Literal["stop", "warning", "information"]


@dataclass(frozen=True)
class DataValidation:
    """What a range of cells will accept.

    Build one with the constructors rather than by hand: :meth:`any_of`,
    :meth:`whole_number`, :meth:`decimal`, :meth:`date`, :meth:`time`,
    :meth:`text_length` and :meth:`custom` set the type, operator and
    formulas together, which is where the combinations go wrong.
    """

    kind: ValidationType = "none"
    operator: ValidationOperator | None = None
    formula1: str | None = None
    formula2: str | None = None
    allow_blank: bool = True
    #: Whether the in-cell dropdown is suppressed. The file spells this
    #: ``showDropDown``, inverted; see the module docstring.
    hide_dropdown: bool = False
    error_style: ErrorStyle | None = None
    error_title: str | None = None
    error_message: str | None = None
    prompt_title: str | None = None
    prompt_message: str | None = None
    show_error: bool = True
    show_prompt: bool = True
    ranges: tuple[RangeRef, ...] = ()
    #: The revision id Excel stamps on each one, carried so an untouched
    #: validation writes back unchanged. Out of the comparison because it
    #: identifies an edit, not a rule.
    uid: str | None = None

    # ------------------------------------------------------------------
    # Constructors
    # ------------------------------------------------------------------

    @classmethod
    def any_of(
        cls, items: Sequence[str] | str, *, hide_dropdown: bool = False, **rest: object
    ) -> DataValidation:
        """A list, either of literal values or from a range.

        A sequence becomes the quoted, comma-joined form Excel uses for a
        typed-in list. A string is taken as a reference to the cells holding
        the items.
        """
        if isinstance(items, str):
            formula = items.lstrip("=")
        else:
            values = list(items)
            for value in values:
                if "," in value:
                    raise ValueError(
                        f"{value!r} contains a comma, and a typed-in list is stored as one "
                        f"comma-separated string, so Excel would read it as two items. Put "
                        f"the list in cells and pass the range instead."
                    )
                if '"' in value:
                    raise ValueError(
                        f"{value!r} contains a quote, which a typed-in list cannot hold. "
                        f"Put the list in cells and pass the range instead."
                    )
            formula = '"' + ",".join(values) + '"'
        return cls(kind="list", formula1=formula, hide_dropdown=hide_dropdown, **rest)  # type: ignore[arg-type]

    @classmethod
    def whole_number(
        cls, minimum: object = None, maximum: object = None, *, operator: ValidationOperator | None = None, **rest: object
    ) -> DataValidation:
        return _bounded("whole", minimum, maximum, operator, rest)

    @classmethod
    def decimal(
        cls, minimum: object = None, maximum: object = None, *, operator: ValidationOperator | None = None, **rest: object
    ) -> DataValidation:
        return _bounded("decimal", minimum, maximum, operator, rest)

    @classmethod
    def date(
        cls, earliest: object = None, latest: object = None, *, operator: ValidationOperator | None = None, **rest: object
    ) -> DataValidation:
        """Bounds as dates, stored as the serials Excel keeps."""
        return _bounded("date", earliest, latest, operator, rest)

    @classmethod
    def time(
        cls, earliest: object = None, latest: object = None, *, operator: ValidationOperator | None = None, **rest: object
    ) -> DataValidation:
        return _bounded("time", earliest, latest, operator, rest)

    @classmethod
    def text_length(
        cls, minimum: object = None, maximum: object = None, *, operator: ValidationOperator | None = None, **rest: object
    ) -> DataValidation:
        return _bounded("textLength", minimum, maximum, operator, rest)

    @classmethod
    def custom(cls, formula: str, **rest: object) -> DataValidation:
        """Accept a value when a formula is true.

        The formula is relative to the top-left of the range, as a
        conditional format's is.
        """
        return cls(kind="custom", formula1=formula.lstrip("="), **rest)  # type: ignore[arg-type]

    # ------------------------------------------------------------------
    # Messages
    # ------------------------------------------------------------------

    def with_error(
        self, message: str, *, title: str | None = None, style: ErrorStyle = "stop"
    ) -> DataValidation:
        """The same rule, with what Excel says when a value is refused.

        ``stop`` refuses the value, ``warning`` and ``information`` let the
        user keep it.
        """
        return replace(
            self,
            error_message=message,
            error_title=title,
            error_style=None if style == "stop" else style,
            show_error=True,
        )

    def with_prompt(self, message: str, *, title: str | None = None) -> DataValidation:
        """The same rule, with the tip shown when the cell is selected."""
        return replace(self, prompt_message=message, prompt_title=title, show_prompt=True)

    @property
    def items(self) -> tuple[str, ...] | None:
        """A typed-in list's values, or ``None`` when it is not one.

        A list pointed at a range returns ``None`` as well: its items are in
        cells, and reading them is the caller's business.
        """
        if self.kind != "list" or not self.formula1:
            return None
        text = self.formula1
        if not (text.startswith('"') and text.endswith('"') and len(text) >= 2):
            return None
        return tuple(text[1:-1].split(","))

    @property
    def sqref(self) -> str:
        return " ".join(block.a1 for block in self.ranges)

    # ------------------------------------------------------------------
    # The part
    # ------------------------------------------------------------------

    @classmethod
    def read(cls, element: Element) -> DataValidation:
        return cls(
            kind=_as_kind(element.get("type")),
            # Absent means between, which is the default Excel relies on.
            operator=_as_operator(element.get("operator")),
            formula1=_child_text(element, "formula1"),
            formula2=_child_text(element, "formula2"),
            allow_blank=element.get("allowBlank") in ("1", "true"),
            hide_dropdown=element.get("showDropDown") in ("1", "true"),
            error_style=_as_style(element.get("errorStyle")),
            error_title=element.get("errorTitle"),
            error_message=element.get("error"),
            prompt_title=element.get("promptTitle"),
            prompt_message=element.get("prompt"),
            show_error=element.get("showErrorMessage") in ("1", "true"),
            show_prompt=element.get("showInputMessage") in ("1", "true"),
            ranges=parse_sqref(element.get("sqref") or ""),
            uid=element.get("xr:uid"),
        )

    def write(self) -> Element:
        # The attribute order CT_DataValidation declares, which is also the
        # order Excel emits.
        element = Element.create("dataValidation")
        if self.kind != "none":
            element.set("type", self.kind)
        if self.error_style is not None and self.error_style != "stop":
            element.set("errorStyle", self.error_style)
        if self.operator is not None and self.operator != "between":
            element.set("operator", self.operator)
        if self.allow_blank:
            element.set("allowBlank", "1")
        if self.hide_dropdown:
            # Inverted: this is what suppresses the dropdown.
            element.set("showDropDown", "1")
        if self.show_prompt:
            element.set("showInputMessage", "1")
        if self.show_error:
            element.set("showErrorMessage", "1")
        for name, value in (
            ("errorTitle", self.error_title),
            ("error", self.error_message),
            ("promptTitle", self.prompt_title),
            ("prompt", self.prompt_message),
        ):
            if value is not None:
                element.set(name, value)
        element.set("sqref", self.sqref)
        if self.uid is not None:
            element.set("xr:uid", self.uid)
        for name, formula in (("formula1", self.formula1), ("formula2", self.formula2)):
            if formula is None:
                continue
            node = Element.create(name)
            node.set_text(formula)
            element.append(node)
        return element


def _bounded(
    kind: ValidationType,
    first: object,
    second: object,
    operator: ValidationOperator | None,
    rest: dict[str, object],
) -> DataValidation:
    """A rule with one or two bounds, picking the operator to match."""
    if first is None:
        raise ValueError(f"a {kind} validation needs at least one bound.")
    chosen = operator
    if chosen is None:
        chosen = "between" if second is not None else "greaterThanOrEqual"
    if chosen in ("between", "notBetween") and second is None:
        raise ValueError(f"{chosen!r} needs two bounds, not one.")
    return DataValidation(
        kind=kind,
        operator=chosen,
        formula1=_bound(first),
        formula2=None if second is None else _bound(second),
        **rest,  # type: ignore[arg-type]
    )


def _bound(value: object) -> str:
    """A bound as the formula text Excel stores.

    A date or a time becomes its serial, which is what the file holds and
    what Excel shows back as a date.
    """
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return _number(datetime_to_serial(value))
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        return _number(value)
    text = str(value)
    return text[1:] if text.startswith("=") else text


def _number(value: float) -> str:
    if float(value) == int(value):
        return str(int(value))
    return repr(float(value))


def _child_text(element: Element, name: str) -> str | None:
    child = element.child(name)
    return None if child is None else (child.text or "")


def _as_kind(raw: str | None) -> ValidationType:
    return raw if raw else "none"  # type: ignore[return-value]


def _as_operator(raw: str | None) -> ValidationOperator | None:
    return raw  # type: ignore[return-value]


def _as_style(raw: str | None) -> ErrorStyle | None:
    return raw  # type: ignore[return-value]

__all__ = [
    "DataValidation",
    "ErrorStyle",
    "ValidationOperator",
    "ValidationType",
]
