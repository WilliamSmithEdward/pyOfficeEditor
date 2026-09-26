"""The collation Excel's filter and sort compare text by, against what
Excel answered.

Measured on a column of text: an autofilter keeping one value shows which
others Excel takes for equal to it, and a sort with case ignored and with
case matched shows where each goes. These are the characters the order's
word sort alone does not settle: a full-width and a raised digit, a
zero-width space, and final sigma.
"""

from __future__ import annotations

import pytest

from pyofficeeditor.excel._collate import case_sort_key, equal, sort_key

#: The column filtered and sorted, in the order it was written.
VALUES = ["1", "\uff11", "\u00b2", "2", "ab", "a\u200bb", "\u03c3", "\u03c2", "\u03a3"]

#: What a filter keeping each value showed.
FILTERED = {
    "1": ["1"],
    "2": ["2"],
    "ab": ["ab"],
    "\u03c3": ["\u03c3", "\u03c2", "\u03a3"],
}

#: The order a sort left the column in, the same whether case was ignored
#: or matched.
SORTED = ["1", "\uff11", "2", "\u00b2", "a\u200bb", "ab", "\u03c3", "\u03c2", "\u03a3"]


@pytest.mark.parametrize(("kept", "shown"), FILTERED.items())
def test_a_filter_takes_for_equal_what_excel_takes(kept: str, shown: list[str]) -> None:
    assert [value for value in VALUES if equal(value, kept)] == shown


def test_a_sort_orders_them_as_excel_orders_them() -> None:
    assert sorted(VALUES, key=sort_key) == SORTED


def test_with_case_matched_too() -> None:
    """Final sigma goes between sigma and its capital, as a title-case
    letter does between its lowercase and its capital."""
    assert sorted(VALUES, key=case_sort_key) == SORTED
    assert sorted(["\u03c2", "\u03a3", "\u03c3"], key=case_sort_key) == ["\u03c3", "\u03c2", "\u03a3"]
