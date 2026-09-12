"""Loading a package from a container. docx4j ``io3.Load3``.

CR-002 section 5.5. The order is docx4j's exactly:

1. ``[Content_Types].xml``;
2. ``/_rels/.rels``, unmarshalled;
3. the ``officeDocument`` relationship, whose target's content type picks the
   package class;
4. every part reachable through relationships, depth first, each created once
   through the :class:`PartRegistry` and linked into the package;
5. the custom XML data parts indexed by the ``ds:itemID`` of their properties
   part.

**Parts not reachable through a relationship are not loaded.** That is docx4j's
rule and Word's: a zip entry no relationship names is not part of the document.
It is still in the source store, so a round trip that touches nothing does not
carry it --- which is also what Word does when it rewrites a file.

**Nothing is unmarshalled here.** Loading a package costs the central directory
and the relationship parts; every other part's content waits for ``.contents``.
"""

from __future__ import annotations

import logging
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

from docx4j_py.openpackaging.content_types import ContentTypeManager
from docx4j_py.openpackaging.exceptions import Docx4JException, InvalidFormatException
from docx4j_py.openpackaging.part_name import CONTENT_TYPES_NAME, PartName
from docx4j_py.openpackaging.parts.default_xml_part import CustomXmlDataStoragePart
from docx4j_py.openpackaging.parts.docprops import (
    DS_NS,
    CustomXmlDataStoragePropertiesPart,
)
from docx4j_py.openpackaging.parts.namespaces import (
    MAIN_PART_RELATIONSHIP_TYPES,
    Namespaces,
)
from docx4j_py.openpackaging.parts.part import Part
from docx4j_py.openpackaging.parts.registry import PartRegistry, default_part_registry
from docx4j_py.openpackaging.parts.relationships_part import (
    RelationshipsPart,
    is_external,
)
from docx4j_py.openpackaging.stores import PartStore

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.openpackaging.packages.opc_package import OpcPackage

__all__ = ["LoadOptions", "load_package"]

logger = logging.getLogger("docx4j_py.openpackaging")

_SCHEME = __import__("re").compile(r"^[a-zA-Z][a-zA-Z0-9+.-]*:")


def _looks_absolute(target: str | None) -> bool:
    """Whether a relationship target is an absolute URI rather than a part name."""
    return bool(target) and _SCHEME.match(target) is not None


@dataclass(slots=True)
class LoadOptions:
    """How to load a package. CR-002 section 5.5.

    Attributes:
        strict: raise on the first piece of content the parser skips, instead of
            reporting it on the part. Off by default (decided question 5):
            lenient with the report is what real documents need, and strict is
            for tests and for callers who prefer to fail.
        mce_preprocess: resolve ``mc:AlternateContent`` before unmarshalling, as
            docx4j does. Off by default (decided question 1): the model holds
            both branches, so resolving is a *view* (``mce="resolve"`` on the
            traversal functions) and not a rewrite.
        registry: the part classes to use; docx4j's table by default.
    """

    strict: bool = False
    mce_preprocess: bool = False
    registry: PartRegistry | None = field(default=None)

    @property
    def part_registry(self) -> PartRegistry:
        """:attr:`registry`, or docx4j's table."""
        return self.registry if self.registry is not None else default_part_registry


def load_package(
    store: PartStore,
    create_package: Callable[[str | None], OpcPackage],
    options: LoadOptions | None = None,
) -> OpcPackage:
    """Load a package from a container. docx4j ``Load3.get()``.

    Args:
        store: where the bytes come from.
        create_package: given the main part's content type, returns the empty
            package object to fill in. This is how the kind of package is
            sniffed without this module knowing the package classes.
        options: see :class:`LoadOptions`.

    Returns:
        The package, with every reachable part created but none unmarshalled.
    """
    options = options or LoadOptions()

    if not store.has(CONTENT_TYPES_NAME):
        raise InvalidFormatException(f"{CONTENT_TYPES_NAME} is missing from this package")
    ctm = ContentTypeManager.parse(store.load(CONTENT_TYPES_NAME))

    package_rels = RelationshipsPart.create_package_rels()
    if not store.has("_rels/.rels"):
        raise InvalidFormatException("_rels/.rels appears to be missing from this package")
    package_rels.set_bytes(store.load("_rels/.rels"))

    main_rel = next(
        (r for r in package_rels.list if r.type_value in MAIN_PART_RELATIONSHIP_TYPES),
        None,
    )
    if main_rel is None:
        raise InvalidFormatException("No relationship of type officeDocument")
    main_part_name = PartName.resolve(PartName.ROOT, main_rel.target or "")

    package = create_package(ctm.get_content_type(main_part_name))
    package.content_type_manager = ctm
    package.source_part_store = store
    package.load_options = options
    package.relationships_part = package_rels
    package_rels.source_p = package
    package_rels.package = package

    is_strict = any(r.type_value == Namespaces.DOCUMENT_STRICT for r in package_rels.list)
    if is_strict:
        package.was_strict = True
        package_rels.import_strict()

    loader = _Loader(package, store, options)
    loader.add_parts_from_relationships(package, package_rels)
    loader.register_custom_xml_data_storage_parts()
    return package


class _Loader:
    """The recursion. docx4j's ``Load3`` methods, as one object."""

    __slots__ = ("handled", "options", "package", "registry", "store")

    def __init__(self, package: OpcPackage, store: PartStore, options: LoadOptions) -> None:
        self.package = package
        self.store = store
        self.options = options
        self.registry = options.part_registry
        self.handled: set[str] = set()

    def add_parts_from_relationships(self, source: Any, rels: RelationshipsPart) -> None:
        """Create a part for every internal relationship, in document order."""
        for rel in list(rels.list):
            try:
                self._one(source, rels, rel)
            except Docx4JException:
                raise
            except Exception as exc:  # noqa: BLE001 - name the rels part
                raise Docx4JException(
                    f"Failed to add parts from the relationships of {rels.part_name}: {exc}"
                ) from exc

    def _one(self, source: Any, rels: RelationshipsPart, rel: Any) -> None:
        if (
            rel.type_value == Namespaces.HYPERLINK
            or is_external(rel)
            or _looks_absolute(rel.target)
        ):
            # A hyperlink is a URI even when it is not marked External, and an
            # external target is not a part. docx4j skips the first two; the
            # third is ours, because a producer that writes an absolute URI and
            # forgets ``TargetMode="External"`` would otherwise make the whole
            # package unloadable, and docx4j throws on exactly that.
            return

        part_name = rels.resolve_target(rel)
        if part_name.key in self.handled:
            # A part can be the target of several relationships: a header used
            # by two sections, an image placed twice. Record the extra
            # relationship and do not walk it again.
            existing = self.package.get_part(part_name)
            if existing is not None:
                source.set_part_shortcut(existing, rel.type_value)
                existing.source_relationships.append(rel)
            return

        part = self._create(part_name, rel, rels)
        source.set_part_shortcut(part, rel.type_value)
        if not part.relationship_type:
            part.relationship_type = rel.type_value or ""
        rels.load_part(part, rel)
        self.handled.add(part_name.key)

        rels_name = PartName.rels_for(part_name).store_name
        if self.store.has(rels_name):
            own = RelationshipsPart(part)
            own.set_bytes(self.store.load(rels_name))
            own.package = self.package
            part.relationships_part = own
            if self.package.was_strict:
                own.import_strict()
            self.add_parts_from_relationships(part, own)

    def _create(self, part_name: PartName, rel: Any, rels: RelationshipsPart) -> Part:
        from urllib.parse import unquote

        store_name = part_name.store_name
        if not self.store.has(store_name) and not self.store.has(unquote(store_name)):
            raise Docx4JException(
                f"For source {rels.source_name}, cannot find part {part_name} "
                f"from rel {rel.id}={rel.target}"
            )
        content_type = self.package.content_type_manager.get_content_type(part_name)
        return self.registry.create_part(part_name, content_type, rel)

    def register_custom_xml_data_storage_parts(self) -> None:
        """Index the custom XML data parts by their ``ds:itemID``.

        docx4j ``Load.registerCustomXmlDataStorageParts``. The id is read off the
        properties part's root element, so neither part is unmarshalled.
        """
        for part in self.package.parts.parts():
            if not isinstance(part, CustomXmlDataStoragePart):
                continue
            rels = part.relationships_part
            if rels is None:
                continue
            rel = rels.get_relationship_by_type(Namespaces.CUSTOM_XML_DATA_STORAGE_PROPERTIES)
            props = rels.get_part(rel) if rel is not None else None
            if props is None:
                continue
            item_id = _item_id(props)
            if not item_id:
                continue
            part.item_id = item_id.lower()
            self.package.custom_xml_data_storage_parts[part.item_id] = part


def _item_id(props: Part) -> str | None:
    if isinstance(props, CustomXmlDataStoragePropertiesPart):
        return props.item_id
    try:
        from lxml import etree

        root = etree.fromstring(props.bytes_for_save)
    except Exception:  # noqa: BLE001 - a broken properties part is not fatal
        logger.warning("%s: could not read ds:itemID", props.part_name)
        return None
    return root.get(f"{{{DS_NS}}}itemID") or root.get("itemID")
