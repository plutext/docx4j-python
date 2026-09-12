"""``Parts``: a package's parts by name. docx4j ``Parts``.

CR-002 decided question 3: a ``Mapping[PartName, Part]`` wrapper rather than a
``dict`` subclass, so that case-insensitivity is enforced in one place and
``add``/``remove`` can keep the part and the map in step. Lookup takes a
:class:`PartName` or a plain string.
"""

from __future__ import annotations

from collections.abc import Iterator, Mapping

from docx4j_py.openpackaging.part_name import PartName
from docx4j_py.openpackaging.parts.part import Part

__all__ = ["Parts"]


class Parts(Mapping[PartName, Part]):
    """The parts of a package, keyed case-insensitively by name."""

    __slots__ = ("_by_key",)

    def __init__(self) -> None:
        """An empty map."""
        self._by_key: dict[str, Part] = {}

    # -- Mapping -----------------------------------------------------------

    def __getitem__(self, key: PartName | str) -> Part:
        """The part with this name; ``KeyError`` when there is none."""
        return self._by_key[PartName.of(key).key]

    def __iter__(self) -> Iterator[PartName]:
        """Iterate the part names, in the order they were added."""
        return (part.part_name for part in self._by_key.values())

    def __len__(self) -> int:
        """The number of parts."""
        return len(self._by_key)

    def __contains__(self, key: object) -> bool:
        """Whether a part of this name is in the package."""
        if isinstance(key, (PartName, str)):
            return PartName.of(key).key in self._by_key
        return False

    # -- mutation ----------------------------------------------------------

    def add(self, part: Part) -> None:
        """Put a part in, replacing any part of the same name."""
        self._by_key[part.part_name.key] = part

    def remove(self, key: PartName | str) -> Part | None:
        """Take a part out and return it, or None."""
        return self._by_key.pop(PartName.of(key).key, None)

    def rename(self, old: PartName | str, part: Part) -> None:
        """Re-key a part whose name changed. Called by ``Part.part_name``'s setter."""
        self._by_key.pop(PartName.of(old).key, None)
        self._by_key[part.part_name.key] = part

    # -- convenience -------------------------------------------------------

    def parts(self) -> list[Part]:
        """Every part, in the order they were added."""
        return list(self._by_key.values())

    def of_type(self, cls: type) -> list[Part]:
        """Every part that is an instance of a class."""
        return [part for part in self._by_key.values() if isinstance(part, cls)]

    def __repr__(self) -> str:
        """``Parts(18)``."""
        return f"Parts({len(self._by_key)})"
