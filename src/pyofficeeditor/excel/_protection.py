"""Sheet protection: what a protected sheet still lets you do.

**Every flag names a lock, not a permission.** ``<sheetProtection
formatCells="0"/>`` means formatting cells is *allowed*. Measured: calling
``Protect(AllowFormattingCells:=True, AllowSorting:=True)`` made Excel write
``sheet="1" formatCells="0" sort="0"``, so the attribute asserts the
restriction and ``0`` lifts it. Most of them are blocked when absent, which
is why a plain ``<sheetProtection sheet="1"/>`` locks almost everything.

The two selection flags go the other way round: ``selectLockedCells`` and
``selectUnlockedCells`` are *absent* on an ordinary protected sheet, where
selection is allowed, and set to ``1`` to forbid it. So the direction cannot
be read off the attribute name; each one carries its own default here.

This class exposes ``allow_*`` fields, because that is the question a caller
is actually asking, and does the inversion on the way out.

**A password is a hash, and a modern one.** Excel writes
``algorithmName="SHA-512"`` with a base64 ``saltValue``, a ``hashValue`` and
``spinCount="100000"``. The hash is SHA-512 over the salt followed by the
password in UTF-16LE, then that many rounds of SHA-512 over the previous
digest followed by the round number as a little-endian uint32. Reproduced
byte for byte against a workbook Excel protected with ``secret``.

Protection is a deterrent, not a security control: the hash guards the
workbook's own UI and anything that can read the file can remove it. Do not
use it to keep a secret.
"""

from __future__ import annotations

import base64
import hashlib
import os
from dataclasses import dataclass
from typing import Literal

from pyofficeeditor._xml import Element

#: The hash Excel uses, and the one this writes.
DEFAULT_ALGORITHM = "SHA-512"

#: How many rounds Excel spins it.
DEFAULT_SPIN_COUNT = 100000

#: ``algorithmName`` values and the :mod:`hashlib` name behind each.
ALGORITHMS: dict[str, str] = {
    "SHA-512": "sha512",
    "SHA-384": "sha384",
    "SHA-256": "sha256",
    "SHA-1": "sha1",
    "MD5": "md5",
}

#: Each ``allow_`` field, the attribute behind it, and whether the action is
#: blocked when the attribute is absent. The last column is the whole reason
#: this table exists: it is not the same for every flag.
_FLAGS: tuple[tuple[str, str, bool], ...] = (
    ("allow_format_cells", "formatCells", True),
    ("allow_format_columns", "formatColumns", True),
    ("allow_format_rows", "formatRows", True),
    ("allow_insert_columns", "insertColumns", True),
    ("allow_insert_rows", "insertRows", True),
    ("allow_insert_hyperlinks", "insertHyperlinks", True),
    ("allow_delete_columns", "deleteColumns", True),
    ("allow_delete_rows", "deleteRows", True),
    ("allow_sort", "sort", True),
    ("allow_autofilter", "autoFilter", True),
    ("allow_pivot_tables", "pivotTables", True),
    # Selection is permitted on an ordinary protected sheet, so these two
    # are absent when allowed and "1" when forbidden.
    ("allow_select_locked_cells", "selectLockedCells", False),
    ("allow_select_unlocked_cells", "selectUnlockedCells", False),
)

Algorithm = Literal["SHA-512", "SHA-384", "SHA-256", "SHA-1", "MD5"]


def hash_password(
    password: str,
    *,
    salt: bytes | None = None,
    algorithm: Algorithm = DEFAULT_ALGORITHM,
    spin_count: int = DEFAULT_SPIN_COUNT,
) -> tuple[str, str]:
    """The ``hashValue`` and ``saltValue`` Excel would write, base64 encoded.

    A random salt is generated unless one is given, which is what makes two
    workbooks protected with the same password look different.
    """
    name = ALGORITHMS.get(algorithm)
    if name is None:
        raise ValueError(
            f"{algorithm!r} is not an algorithm Excel names; expected one of "
            f"{', '.join(sorted(ALGORITHMS))}."
        )
    if salt is None:
        salt = os.urandom(16)
    digest = hashlib.new(name, salt + password.encode("utf-16-le")).digest()
    for round_number in range(spin_count):
        digest = hashlib.new(name, digest + round_number.to_bytes(4, "little")).digest()
    return base64.b64encode(digest).decode(), base64.b64encode(salt).decode()


@dataclass(frozen=True)
class SheetProtection:
    """What a protected sheet allows.

    The defaults match a bare ``Worksheet.Protect`` in Excel: the cells, the
    drawing objects and the scenarios are all locked, and everything else a
    user might do is refused except selecting cells.
    """

    #: Whether the cells are locked at all. This is the file's ``sheet``
    #: attribute, and without it the rest does nothing.
    contents: bool = True
    objects: bool = True
    scenarios: bool = True

    allow_format_cells: bool = False
    allow_format_columns: bool = False
    allow_format_rows: bool = False
    allow_insert_columns: bool = False
    allow_insert_rows: bool = False
    allow_insert_hyperlinks: bool = False
    allow_delete_columns: bool = False
    allow_delete_rows: bool = False
    allow_sort: bool = False
    allow_autofilter: bool = False
    allow_pivot_tables: bool = False
    allow_select_locked_cells: bool = True
    allow_select_unlocked_cells: bool = True

    #: The password's hash, as the file carries it. Set through
    #: :meth:`with_password` rather than by hand.
    algorithm_name: str | None = None
    hash_value: str | None = None
    salt_value: str | None = None
    spin_count: int | None = None

    @property
    def has_password(self) -> bool:
        return self.hash_value is not None

    def with_password(
        self,
        password: str,
        *,
        algorithm: Algorithm = DEFAULT_ALGORITHM,
        spin_count: int = DEFAULT_SPIN_COUNT,
        salt: bytes | None = None,
    ) -> SheetProtection:
        """The same protection, behind a password.

        A deterrent rather than a secret: anything that can read the file
        can remove this.
        """
        if not password:
            raise ValueError("an empty password protects nothing; leave it off instead.")
        digest, encoded = hash_password(
            password, salt=salt, algorithm=algorithm, spin_count=spin_count
        )
        from dataclasses import replace

        return replace(
            self,
            algorithm_name=algorithm,
            hash_value=digest,
            salt_value=encoded,
            spin_count=spin_count,
        )

    def matches(self, password: str) -> bool:
        """Whether a password is the one this was protected with."""
        if self.hash_value is None or self.salt_value is None:
            return False
        digest, _ = hash_password(
            password,
            salt=base64.b64decode(self.salt_value),
            algorithm=self.algorithm_name or DEFAULT_ALGORITHM,  # type: ignore[arg-type]
            spin_count=self.spin_count or DEFAULT_SPIN_COUNT,
        )
        return digest == self.hash_value

    @classmethod
    def read(cls, element: Element) -> SheetProtection:
        values: dict[str, object] = {
            "contents": element.get("sheet") in ("1", "true"),
            "objects": element.get("objects") in ("1", "true"),
            "scenarios": element.get("scenarios") in ("1", "true"),
            "algorithm_name": element.get("algorithmName"),
            "hash_value": element.get("hashValue"),
            "salt_value": element.get("saltValue"),
            "spin_count": _as_int(element.get("spinCount")),
        }
        for field_name, attribute, blocked_by_default in _FLAGS:
            raw = element.get(attribute)
            blocked = blocked_by_default if raw is None else raw in ("1", "true")
            values[field_name] = not blocked
        return cls(**values)  # type: ignore[arg-type]

    def write(self) -> Element:
        element = Element.create("sheetProtection")
        # The password comes first, as Excel writes it.
        for name, value in (
            ("algorithmName", self.algorithm_name),
            ("hashValue", self.hash_value),
            ("saltValue", self.salt_value),
            ("spinCount", None if self.spin_count is None else str(self.spin_count)),
        ):
            if value is not None:
                element.set(name, value)
        for name, flag in (
            ("sheet", self.contents),
            ("objects", self.objects),
            ("scenarios", self.scenarios),
        ):
            if flag:
                element.set(name, "1")
        for field_name, attribute, blocked_by_default in _FLAGS:
            blocked = not getattr(self, field_name)
            # Only when it differs from what absence already means.
            if blocked != blocked_by_default:
                element.set(attribute, "1" if blocked else "0")
        return element


def _as_int(raw: str | None) -> int | None:
    if raw is None:
        return None
    try:
        return int(raw)
    except ValueError:
        return None


__all__ = [
    "ALGORITHMS",
    "DEFAULT_ALGORITHM",
    "DEFAULT_SPIN_COUNT",
    "Algorithm",
    "SheetProtection",
    "hash_password",
]
