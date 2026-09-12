"""Writing a package to a container. docx4j ``io3.Save``.

CR-002 section 5.5, and the order is docx4j's:

1. ``[Content_Types].xml`` **first**, as Word writes it;
2. ``/_rels/.rels``;
3. every part reachable through relationships, depth first, each written once,
   and each part's own ``.rels`` after it (skipped when it holds nothing, as
   docx4j skips it).

What each part contributes is :attr:`Part.bytes_for_save`: a relationships part
and an unmarshalled :class:`XmlPart` are marshalled; everything else is the
bytes it was loaded with, copied from the source store **without being decoded**.
There is no dirty tracking (decided question 2): a part the caller unmarshalled
is re-marshalled even if nothing in it changed, which is docx4j's rule and is
the safe one, and the cost is bounded by what the caller chose to unmarshal.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from docx4j_py.openpackaging.content_types import ContentTypes
from docx4j_py.openpackaging.exceptions import Docx4JException
from docx4j_py.openpackaging.part_name import CONTENT_TYPES_NAME
from docx4j_py.openpackaging.parts.namespaces import Namespaces
from docx4j_py.openpackaging.parts.part import Part
from docx4j_py.openpackaging.parts.relationships_part import (
    RelationshipsPart,
    is_external,
)
from docx4j_py.openpackaging.stores import PartSink

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.openpackaging.packages.opc_package import OpcPackage

__all__ = ["save_package"]


def save_package(package: OpcPackage, sink: PartSink) -> Any:
    """Write a package to a sink and return whatever the sink produces."""
    sink.put(
        CONTENT_TYPES_NAME,
        package.content_type_manager.to_bytes(),
        content_type=ContentTypes.CONTENT_TYPES_PART,
    )
    saver = _Saver(package, sink)
    saver.save_rels(package.relationships_part)
    return sink.finish()


class _Saver:
    """The recursion. docx4j's ``Save`` methods, as one object."""

    __slots__ = ("handled", "package", "sink")

    def __init__(self, package: OpcPackage, sink: PartSink) -> None:
        self.package = package
        self.sink = sink
        self.handled: set[str] = set()

    def save_rels(self, rels: RelationshipsPart) -> None:
        """Write a relationships part, then every part it targets."""
        self.sink.put(
            rels.part_name.store_name,
            rels.bytes_for_save,
            content_type=ContentTypes.RELATIONSHIPS_PART,
        )
        for rel in rels.list:
            if rel.type_value == Namespaces.HYPERLINK or is_external(rel):
                continue
            name = rels.resolve_target(rel)
            part = self.package.get_part(name)
            if part is None:
                raise Docx4JException(
                    f"Part {name} (target of {rel.id} in {rels.part_name}) is not in the package"
                )
            self.save_part(part)

    def save_part(self, part: Part) -> None:
        """Write one part and, after it, its own relationships."""
        key = part.part_name.key
        if key in self.handled:
            return
        self.handled.add(key)
        try:
            data = part.bytes_for_save
        except Docx4JException:
            raise
        except Exception as exc:  # noqa: BLE001 - name the part
            raise Docx4JException(f"Problem saving part {part.part_name}: {exc}") from exc
        self.sink.put(
            part.part_name.store_name,
            data,
            content_type=part.content_type,
            compress=part.compress,
        )
        own = part.relationships_part
        if own is not None and len(own.list) > 0:
            self.save_rels(own)
