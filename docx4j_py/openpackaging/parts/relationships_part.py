"""``RelationshipsPart``: a ``.rels``, over the generated model.

CR-002 section 5.2, docx4j
``org.docx4j.openpackaging.parts.relationships.RelationshipsPart``. An
:class:`~docx4j_py.openpackaging.parts.xml_part.XmlPart` whose content is the
generated ``docx4j_py.relationships.Relationships`` (CR-002 section 9), so the
relationship graph is typed like everything else.

Every relationships part is always unmarshalled --- loading a package walks it
--- and therefore always re-serialised on save, which is docx4j's behaviour too.
"""

from __future__ import annotations

import enum
from typing import TYPE_CHECKING, Any

from docx4j_py.openpackaging.content_types import ContentTypes
from docx4j_py.openpackaging.exceptions import (
    InvalidFormatException,
    InvalidOperationException,
)
from docx4j_py.openpackaging.part_name import PACKAGE_RELS_NAME, PartName
from docx4j_py.openpackaging.parts.namespaces import Namespaces
from docx4j_py.openpackaging.parts.part import Part, _set_relationships_part_factory
from docx4j_py.openpackaging.parts.xml_part import XmlPart

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.relationships import Relationship, Relationships

__all__ = ["AddPartBehaviour", "RelationshipsPart", "is_external"]

RELATIONSHIPS_ROOT = f"{{{Namespaces.RELATIONSHIPS}}}Relationships"


class AddPartBehaviour(enum.Enum):
    """What ``add_part`` does when the package already holds that part name.

    docx4j's enum, with its three values and docx4j's semantics:

    ``OVERWRITE_IF_NAME_EXISTS``
        the default: the new part replaces the old one under the same name, and
        an existing relationship with that target is reused rather than
        duplicated.
    ``REUSE_EXISTING``
        the part already in the package wins and is returned; nothing is
        written. What you want when adding the same image twice.
    ``RENAME_IF_NAME_EXISTS``
        the new part is given the first free name of the form
        ``image2.png``, ``image3.png``, ... and added under that.
    """

    OVERWRITE_IF_NAME_EXISTS = "overwrite"
    REUSE_EXISTING = "reuse"
    RENAME_IF_NAME_EXISTS = "rename"


def is_external(rel: Any) -> bool:
    """Whether a relationship's target is a URI rather than a part.

    ``TargetMode`` is a generated enum (``RelationshipTargetMode.EXTERNAL``),
    because the schema declares it as an inline enumeration; the string form is
    accepted too, so a hand-built relationship works either way.
    """
    mode = getattr(rel, "target_mode", None)
    if mode is None:
        return False
    return getattr(mode, "value", mode) == "External"


class RelationshipsPart(XmlPart["Relationships"]):
    """A ``.rels`` part. docx4j ``RelationshipsPart``."""

    __slots__ = ("_next_id", "source_p")

    def __init__(
        self,
        source: Any = None,
        part_name: PartName | str | None = None,
    ) -> None:
        """Build the relationships of `source`, or of the package at ``/_rels/.rels``."""
        if part_name is None:
            part_name = (
                PartName.rels_for(source.part_name) if source is not None else PACKAGE_RELS_NAME
            )
        super().__init__(
            part_name,
            ContentTypes.RELATIONSHIPS_PART,
            "",
            RELATIONSHIPS_ROOT,
        )
        #: The part, or package, whose relationships these are. docx4j ``sourceP``.
        self.source_p = source
        self._next_id = 1
        if source is not None:
            self.package = getattr(source, "package", None) or (
                source if hasattr(source, "parts") else None
            )

    model_class_path = "docx4j_py.relationships.Relationships"
    default_namespace = Namespaces.RELATIONSHIPS

    @property
    def is_relationships_part(self) -> bool:
        """True. A ``.rels`` never has relationships of its own."""
        return True

    # -- construction ------------------------------------------------------

    @classmethod
    def create_package_rels(cls) -> RelationshipsPart:
        """An empty ``/_rels/.rels``."""
        from docx4j_py.relationships import Relationships

        part = cls(None, PACKAGE_RELS_NAME)
        part.set_contents(Relationships())
        return part

    @classmethod
    def create_for(cls, source: Any) -> RelationshipsPart:
        """An empty relationships part for a part. docx4j ``createRelationshipsPartForPart``."""
        from docx4j_py.relationships import Relationships

        part = cls(source)
        part.set_contents(Relationships())
        package = getattr(source, "package", None)
        if package is not None:
            part.package = package
            package.content_type_manager.add_default_content_type(
                "rels", ContentTypes.RELATIONSHIPS_PART
            )
        return part

    # -- the graph ---------------------------------------------------------

    @property
    def source_name(self) -> PartName:
        """The name of the part these relationships belong to; ``/`` for the package."""
        source = self.source_p
        if source is None:
            return PartName.ROOT
        name = getattr(source, "part_name", None)
        return name if isinstance(name, PartName) else PartName.ROOT

    @property
    def is_package_relationship_part(self) -> bool:
        """Whether this is ``/_rels/.rels``."""
        return self.source_name == PartName.ROOT

    @property
    def relationships(self) -> Relationships:
        """The typed root. docx4j ``getRelationships()``."""
        return self.contents

    @property
    def list(self) -> list[Relationship]:
        """The live list of relationships, in document order."""
        return self.relationships.relationship

    def __len__(self) -> int:
        """The number of relationships."""
        return len(self.list)

    def __iter__(self):
        """Iterate the relationships in document order."""
        return iter(self.list)

    def get_relationship_by_id(self, rel_id: str) -> Relationship | None:
        """The relationship with this ``Id``, or None."""
        for rel in self.list:
            if rel.id == rel_id:
                return rel
        return None

    def get_relationship_by_type(self, relationship_type: str) -> Relationship | None:
        """The first relationship of this ``Type``, or None."""
        for rel in self.list:
            if rel.type_value == relationship_type:
                return rel
        return None

    def get_relationships_by_type(self, relationship_type: str) -> list[Relationship]:
        """Every relationship of this ``Type``, in document order."""
        return [rel for rel in self.list if rel.type_value == relationship_type]

    def resolve_target(self, rel: Relationship) -> PartName:
        """The part name an internal relationship's target resolves to."""
        return PartName.resolve(self.source_name, rel.target or "")

    def get_part(self, rel: Relationship | str) -> Part | None:
        """The part a relationship targets; None for an external target."""
        resolved = self.get_relationship_by_id(rel) if isinstance(rel, str) else rel
        if resolved is None or is_external(resolved):
            return None
        package = self.package
        if package is None:
            return None
        return package.get_part(self.resolve_target(resolved))

    def get_rel(self, part_name: PartName | str) -> Relationship | None:
        """The first relationship here that targets a part."""
        wanted = PartName.of(part_name)
        for rel in self.list:
            if not is_external(rel) and self.resolve_target(rel) == wanted:
                return rel
        return None

    def is_a_target(self, part_name: PartName | str) -> bool:
        """Whether a part is the target of any relationship here."""
        return self.get_rel(part_name) is not None

    # -- ids ---------------------------------------------------------------

    def get_next_id(self) -> str:
        """The next free ``rIdN``. docx4j ``getNextId``."""
        used = {rel.id for rel in self.list}
        while True:
            candidate = f"rId{self._next_id}"
            self._next_id += 1
            if candidate not in used:
                return candidate

    def reset_id_allocator(self) -> None:
        """Recompute the id counter from the relationships. docx4j ``resetIdAllocator``."""
        highest = 0
        for rel in self.list:
            rel_id = rel.id or ""
            if rel_id.startswith("rId"):
                try:
                    highest = max(highest, int(rel_id[3:]))
                except ValueError:
                    continue
        self._next_id = highest + 1

    # -- mutation ----------------------------------------------------------

    def add_relationship(self, rel: Relationship) -> str:
        """Add a relationship, allocating an id when it has none. Returns the id.

        Raises:
            InvalidOperationException: the id is already used here, which
                docx4j refuses rather than silently renaming.
        """
        if not rel.id:
            rel.id = self.get_next_id()
        elif self.get_relationship_by_id(rel.id) is not None:
            raise InvalidOperationException(
                f"Refusing to add another rel with id {rel.id}. Target is {rel.target}"
            )
        self.list.append(rel)
        return rel.id

    def add_external_relationship(
        self, relationship_type: str, target: str, rel_id: str | None = None
    ) -> Relationship:
        """Add an external relationship: a hyperlink, a linked image."""
        from docx4j_py.relationships import Relationship, RelationshipTargetMode

        rel = Relationship(
            id=rel_id or "",
            type_value=relationship_type,
            target=target,
            target_mode=RelationshipTargetMode.EXTERNAL,
        )
        self.add_relationship(rel)
        return rel

    def add_part(
        self,
        part: Part,
        mode: AddPartBehaviour | str | None = None,
        rel_id: str | None = None,
    ) -> Relationship:
        """Add a part as a target of these relationships. docx4j ``addPart``.

        Registers the part in the package and its content type in the manager,
        writes a relationship whose target is relative to the source part, and
        returns that relationship. `mode` is :class:`AddPartBehaviour` and
        defaults to ``OVERWRITE_IF_NAME_EXISTS``, as docx4j's does.
        """
        from docx4j_py.relationships import Relationship

        behaviour = _behaviour(mode)
        package = self.package
        if package is None:
            raise InvalidFormatException("This relationships part has no package")

        existing = package.get_part(part.part_name)
        if existing is not None:
            if behaviour is AddPartBehaviour.REUSE_EXISTING:
                part = existing
            elif behaviour is AddPartBehaviour.RENAME_IF_NAME_EXISTS:
                part.part_name = self._new_part_name(part.part_name)

        target = PartName.relativize(self.source_name, part.part_name)
        exists_already = next((r for r in self.list if r.target == target), None)

        if exists_already is not None:
            if behaviour is AddPartBehaviour.REUSE_EXISTING:
                return exists_already
            if behaviour is AddPartBehaviour.RENAME_IF_NAME_EXISTS:
                raise InvalidFormatException(
                    "Found existing rel, and yet the constructed part name should be "
                    "globally unique"
                )
            if existing is not None:  # OVERWRITE, and the name was taken
                exists_already.type_value = part.relationship_type
                self.load_part(part, exists_already)
                return exists_already

        rel = Relationship(
            id=rel_id or "",
            type_value=part.relationship_type,
            target=target,
        )
        self.add_relationship(rel)
        package.content_type_manager.add_content_type(part.part_name, part.content_type)
        self.load_part(part, rel)
        return rel

    def load_part(self, part: Part, source_relationship: Relationship) -> None:
        """Link a part into the package as the target of a relationship here."""
        package = self.package
        if package is None:
            raise InvalidFormatException("This relationships part has no package")
        part.owning_relationship_part = self
        if source_relationship not in part.source_relationships:
            part.source_relationships.append(source_relationship)
        part.package = package
        package.parts.add(part)

    def remove_relationship(self, rel: Relationship | PartName | str) -> None:
        """Remove a relationship, by object or by the part it targets."""
        if isinstance(rel, (PartName, str)):
            resolved = self.get_rel(rel)
        else:
            resolved = rel
        if resolved is None:
            return
        try:
            self.list.remove(resolved)
        except ValueError:
            return
        if not is_external(resolved) and self.package is not None:
            target = self.package.get_part(self.resolve_target(resolved))
            if target is not None and resolved in target.source_relationships:
                target.source_relationships.remove(resolved)

    def remove_relationships_by_type(self, relationship_type: str) -> None:
        """Remove every relationship of a type."""
        for rel in self.get_relationships_by_type(relationship_type):
            self.remove_relationship(rel)

    def remove_part(self, part_name: PartName | str) -> list[PartName]:
        """Remove a part, its relationship, and what only it related to.

        Returns:
            The names removed, the part's own first.
        """
        wanted = PartName.of(part_name)
        removed: list[PartName] = []
        package = self.package
        part = package.get_part(wanted) if package is not None else None
        if part is None or package is None:
            return removed
        self.remove_relationship(wanted)
        if part.relationships_part is not None:
            removed.extend(part.relationships_part.remove_parts())
        package.parts.remove(wanted)
        removed.append(wanted)
        return removed

    def remove_parts(self) -> list[PartName]:
        """Remove every part these relationships target, recursively."""
        removed: list[PartName] = []
        for rel in list(self.list):
            if is_external(rel):
                continue
            name = self.resolve_target(rel)
            part = self.package.get_part(name) if self.package is not None else None
            if part is not None and len(part.source_relationships) <= 1:
                removed.extend(self.remove_part(name))
            else:
                self.remove_relationship(rel)
        return removed

    def import_strict(self) -> None:
        """Rewrite Strict relationship types as Transitional. docx4j ``importStrict``."""
        strict_extended = (
            "http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties"
        )
        for rel in self.list:
            value = rel.type_value or ""
            if value == strict_extended:
                rel.type_value = Namespaces.PROPERTIES_EXTENDED
            elif value.startswith(Namespaces.NAMESPACE_PREFIX_STRICT):
                rel.type_value = Namespaces.NAMESPACE_PREFIX_TRANSITIONAL + value[41:]

    # -- helpers -----------------------------------------------------------

    def _new_part_name(self, proposed: PartName) -> PartName:
        """The first free ``image2.png``-style name. docx4j ``getNewPartName``."""
        package = self.package
        assert package is not None
        name = proposed.name
        dot = name.rfind(".")
        if dot > name.rfind("/"):
            prefix, suffix = name[:dot], name[dot:]
        else:
            prefix, suffix = name, ""
        index = 2
        while True:
            candidate = PartName.of(f"{prefix}{index}{suffix}")
            if package.get_part(candidate) is None:
                return candidate
            index += 1

    def __repr__(self) -> str:
        """``RelationshipsPart('/word/_rels/document.xml.rels', 11 rels)``."""
        try:
            count = len(self.list)
        except Exception:  # noqa: BLE001 - repr must not raise
            count = -1
        return f"RelationshipsPart({self.part_name.name!r}, {count} rels)"


def _behaviour(mode: AddPartBehaviour | str | None) -> AddPartBehaviour:
    if mode is None:
        return AddPartBehaviour.OVERWRITE_IF_NAME_EXISTS
    if isinstance(mode, AddPartBehaviour):
        return mode
    try:
        return AddPartBehaviour(mode)
    except ValueError:
        return AddPartBehaviour[mode]


_set_relationships_part_factory(RelationshipsPart.create_for)
