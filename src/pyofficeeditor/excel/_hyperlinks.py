"""Hyperlinks, whose destination is mostly not in the element.

**An external link's URL lives in a relationship.** ``<hyperlink ref="E2"
r:id="rId1" tooltip="go there"/>`` carries no address at all: the target is
the ``.rels`` entry ``rId1`` names, with ``TargetMode="External"``. Reading
the element alone gives a link with nowhere to go, which is why this type
carries the resolved address and the worksheet wires the relationship.

**An internal link has no relationship.** ``<hyperlink ref="E3"
location="L!A1"/>`` points inside the workbook and names its destination in
``location``, so the two kinds are told apart by which of the two is
present rather than by anything in the element's name.

**They can carry both.** A URL with a fragment writes the relationship for
the page and ``location`` for the part after the ``#``.

**A link covers a range.** ``ref`` is one range, not one cell, so a link
laid over ``G2:G4`` is a single entry rather than three.
"""

from __future__ import annotations

from dataclasses import dataclass

from pyofficeeditor._xml import Element
from pyofficeeditor.excel._reference import RangeRef

#: The relationship type an external hyperlink uses.
RT_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"


@dataclass(frozen=True)
class Hyperlink:
    """A link on a cell or a range.

    Exactly one of :attr:`target` and :attr:`location` is usually set:
    ``target`` for somewhere outside the workbook, ``location`` for a cell
    or a defined name inside it. A URL with a fragment sets both.
    """

    ref: RangeRef
    #: The address, for a link that leaves the workbook. Resolved from the
    #: relationship rather than read off the element.
    target: str | None = None
    #: Where inside the workbook it points, such as ``Sheet1!A1``.
    location: str | None = None
    #: The text Excel shows in place of the cell's own, where it differs.
    display: str | None = None
    tooltip: str | None = None
    #: The relationship that holds the target, for a link read from a file.
    relationship_id: str | None = None
    uid: str | None = None

    @property
    def is_external(self) -> bool:
        return self.target is not None

    @classmethod
    def read(cls, element: Element, *, target: str | None = None) -> Hyperlink:
        """One entry. ``target`` comes from the relationship, resolved by
        the caller, because the element does not hold it."""
        raw = element.get("ref") or "A1"
        try:
            ref = RangeRef.parse(raw)
        except ValueError:
            ref = RangeRef.parse("A1")
        return cls(
            ref=ref,
            target=target,
            location=element.get("location"),
            display=element.get("display"),
            tooltip=element.get("tooltip"),
            relationship_id=element.get("r:id"),
            uid=element.get("xr:uid"),
        )

    def write(self) -> Element:
        element = Element.create("hyperlink", {"ref": self.ref.a1})
        if self.relationship_id is not None:
            element.set("r:id", self.relationship_id)
        if self.location is not None:
            element.set("location", self.location)
        if self.tooltip is not None:
            element.set("tooltip", self.tooltip)
        if self.display is not None:
            element.set("display", self.display)
        if self.uid is not None:
            element.set("xr:uid", self.uid)
        return element


__all__ = ["RT_HYPERLINK", "Hyperlink"]
