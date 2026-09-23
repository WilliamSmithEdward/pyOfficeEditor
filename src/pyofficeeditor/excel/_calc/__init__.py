"""Formulas as Excel reads and calculates them.

The rest of the package treats a formula as text to carry: it shifts the
references in it when rows move and renames a sheet inside it, and that is
all. This package reads one as Excel does, into a tree, and calculates it.

- :mod:`.lexer` breaks formula text into tokens, references whole
- :mod:`.nodes` is the tree
- :mod:`.parser` builds the tree with Excel's precedence, which is not the
  usual one: ``-2^2`` is 4 and ``2^3^2`` is 64
"""

from __future__ import annotations

from pyofficeeditor.excel._calc.parser import FormulaSyntaxError, parse

__all__ = ["FormulaSyntaxError", "parse"]
