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

Before any of it, the package's ``save_hooks`` run (:func:`run_save_hooks`,
CR-019 section 4.9): what the content API keeps for the save, such as the
glossary entries a placeholder names, goes into the parts then.
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

__all__ = ["run_save_hooks", "save_package"]


def save_package(package: OpcPackage, sink: PartSink, *, normalise: bool | None = None) -> Any:
    """Write a package to a sink and return whatever the sink produces.

    `normalise` (CR-018 section 4.7): write each value a parse kept as its text
    in Word's normal form, for this save only
    (:func:`~docx4j_py.openpackaging.normalise.normalised_package`). None is the
    package's ``load_options.normalise_on_save``. What it rewrote is
    ``package.last_save``.
    """
    from docx4j_py.openpackaging.normalise import normalised_package

    if normalise is None:
        normalise = bool(getattr(getattr(package, "load_options", None), "normalise_on_save", True))
    added = run_save_hooks(package)
    with normalised_package(package, enabled=normalise) as report:
        sink.put(
            CONTENT_TYPES_NAME,
            package.content_type_manager.to_bytes(),
            content_type=ContentTypes.CONTENT_TYPES_PART,
        )
        saver = _Saver(package, sink)
        saver.save_rels(package.relationships_part)
        result = sink.finish()
    report.added = added
    if hasattr(package, "last_save"):
        package.last_save = report
    return result


def run_save_hooks(package: OpcPackage) -> list[str]:
    """Run the package's :attr:`~OpcPackage.save_hooks`, in order, before anything is written; what they added.

    CR-019 section 4.9: the content API's writes that wait for the save (a
    glossary entry a placeholder names) reach the package here, so that this
    layer imports nothing from it. Each hook is given the package and returns
    the names of the parts it added, or None.
    """
    added: list[str] = []
    for hook in tuple(getattr(package, "save_hooks", None) or ()):
        names = hook(package)
        if names:
            added.extend(str(name) for name in names)
    return added


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
