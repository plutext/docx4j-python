"""``Part``: the base of everything in a package. docx4j ``openpackaging.parts.Part``.

CR-002 section 5.4. A part has a name, a content type, the relationship type it
is the target of, the package it belongs to, its own relationships part if it
has one, and the relationships that point at it. Two abstract questions each
subclass answers: :attr:`bytes_for_save` (what to write) and :attr:`is_loaded`
(whether it carries content of its own or still leans on the source container).

:class:`RelationshipSource` is docx4j's ``Base``: a part, or the package itself,
that can own relationships. It is a protocol rather than a base class because
:class:`~docx4j_py.openpackaging.packages.opc_package.OpcPackage` is not a part.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx4j_py.openpackaging.exceptions import Docx4JException, InvalidFormatException
from docx4j_py.openpackaging.part_name import PartName

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.openpackaging.packages.opc_package import OpcPackage
    from docx4j_py.openpackaging.parts.relationships_part import (
        AddPartBehaviour,
        RelationshipsPart,
    )

__all__ = ["Part"]

# Set by relationships_part.py once it is imported, so that Part can create a
# relationships part without importing the module that subclasses XmlPart, which
# subclasses Part. docx4j has no such cycle because Java resolves it at link
# time; the TypeScript engine uses the same injected factory.
_rels_factory: Any = None


def _set_relationships_part_factory(factory: Any) -> None:
    """Called by ``relationships_part`` at import; breaks the import cycle."""
    global _rels_factory
    _rels_factory = factory


class Part:
    """A part of a package. docx4j ``Part``.

    Attributes:
        part_name: the part's :class:`PartName`.
        content_type: the string from ``[Content_Types].xml``.
        relationship_type: the ``Type`` of the relationship that reached it.
        package: the package it belongs to, once it has been added to one.
        relationships_part: its *own* ``.rels``, if it has relationships.
        owning_relationship_part: the ``.rels`` that first loaded it.
        source_relationships: every relationship that targets it. A header used
            by two sections, or an image placed twice, has more than one.
    """

    __slots__ = (
        "_part_name",
        "content_type",
        "owning_relationship_part",
        "package",
        "relationship_type",
        "relationships_part",
        "source_relationships",
    )

    def __init__(
        self,
        part_name: PartName | str,
        content_type: str = "",
        relationship_type: str = "",
    ) -> None:
        """Build a part. It belongs to no package until something adds it."""
        self._part_name = PartName.of(part_name)
        self.content_type = content_type
        self.relationship_type = relationship_type
        self.package: OpcPackage | None = None
        self.relationships_part: RelationshipsPart | None = None
        self.owning_relationship_part: RelationshipsPart | None = None
        self.source_relationships: list[Any] = []

    # -- identity ----------------------------------------------------------

    @property
    def part_name(self) -> PartName:
        """The part's name."""
        return self._part_name

    @part_name.setter
    def part_name(self, value: PartName | str) -> None:
        """Rename the part, re-keying it in the package if it is in one."""
        new = PartName.of(value)
        old = self._part_name
        self._part_name = new
        if self.package is not None and old != new:
            self.package.parts.rename(old, self)

    @property
    def is_relationships_part(self) -> bool:
        """Whether this part *is* a ``.rels``; those never own relationships."""
        return False

    @property
    def compress(self) -> bool:
        """Whether a zip should deflate this part. Overridden by the binary parts."""
        return True

    # -- relationships -----------------------------------------------------

    @property
    def source_relationship(self) -> Any:
        """The first relationship that targets this part, or None."""
        return self.source_relationships[0] if self.source_relationships else None

    def get_relationships_part(self, create_if_absent: bool = False) -> RelationshipsPart | None:
        """This part's own relationships part, created on demand.

        docx4j ``Base.getRelationshipsPart(boolean)``. A relationships part
        never has one of its own, and answers None.
        """
        if self.is_relationships_part:
            return None
        if self.relationships_part is None and create_if_absent:
            if _rels_factory is None:  # pragma: no cover - import order guard
                raise Docx4JException("RelationshipsPart is not initialised")
            self.relationships_part = _rels_factory(self)
        return self.relationships_part

    def add_target_part(
        self,
        target: Part,
        mode: AddPartBehaviour | str | None = None,
        rel_id: str | None = None,
    ) -> Any:
        """Add a part as a target of this one. docx4j ``Base.addTargetPart``.

        Registers the target in the package and in the content type manager, and
        writes a relationship in *this* part's relationships part, creating it if
        it does not exist yet.

        Returns:
            The ``Relationship`` that was written or reused.
        """
        if self.package is None:
            raise InvalidFormatException(
                "Package not set; if you are adding part2 to part1, add part1 first."
            )
        if self.is_relationships_part:
            raise InvalidFormatException("A relationships part cannot have targets")
        rels = self.get_relationships_part(True)
        assert rels is not None
        rel = rels.add_part(target, mode, rel_id)
        self.set_part_shortcut(target, target.relationship_type)
        return rel

    def set_part_shortcut(self, part: Part, relationship_type: str) -> bool:
        """Keep a typed reference to a well-known target. docx4j ``setPartShortcut``.

        The base does nothing and says so; ``MainDocumentPart`` and the package
        classes override it to fill in ``style_definitions_part`` and friends.
        """
        return False

    # -- bytes -------------------------------------------------------------

    def _source_bytes(self) -> bytes | None:
        """The bytes this part was loaded from, if it came from a container."""
        package = self.package
        store = package.source_part_store if package is not None else None
        if store is None:
            return None
        name = self._part_name.store_name
        if store.has(name):
            return store.load(name)
        from urllib.parse import unquote

        decoded = unquote(name)
        if decoded != name and store.has(decoded):
            return store.load(decoded)
        return None

    @property
    def bytes_for_save(self) -> bytes:
        """The bytes to write on save. docx4j ``ZipPartStore.save*Part``."""
        raise NotImplementedError

    @property
    def is_loaded(self) -> bool:
        """Whether the part carries content of its own rather than the container's."""
        raise NotImplementedError

    def __repr__(self) -> str:
        """``MainDocumentPart('/word/document.xml')``."""
        return f"{type(self).__name__}({self._part_name.name!r})"
