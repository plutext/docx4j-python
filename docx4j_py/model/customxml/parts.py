"""``CustomXmlPart`` and ``pkg.custom_xml_parts``: the data behind the bindings.

CR-003 section 3.7. Office JS's ``Word.CustomXmlPart`` and
``Word.CustomXmlPartCollection`` over the ``/customXml/itemN.xml`` parts CR-002
already loads as lxml trees, with:

* ``get_item(id)`` **case- and brace-insensitive**, because Word writes the
  store item id one way in ``w:dataBinding/@w:storeItemID`` and another in
  ``ds:itemID`` --- three of the twenty bindings in
  ``samples/invoice2013.docx`` name the same part in lower case;
* ``add(xml)`` following docx4j's ``AbstractMigrator.addPropertiesPart``: an
  ``/customXml/itemN.xml`` with its ``/customXml/itemPropsN.xml``, a
  brace-wrapped upper-case UUID as the ``ds:itemID``, ``ds:schemaRefs`` from the
  document element's namespace, and **the relationship from the main document
  part**, because Word silently drops a custom XML part the package alone
  relates to;
* reading nothing. ``items``, ``get_xml()``, ``select_nodes`` and every node
  read leave each part byte for byte; only a mutation calls :meth:`touch`.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from lxml import etree

from docx4j_py.model.content.errors import BindingError, ContentError
from docx4j_py.model.customxml.nodes import (
    CustomXmlNode,
    CustomXmlPrefixMappingCollection,
    namespaces_of,
    node_for,
    require_node,
)
from docx4j_py.model.customxml.xpath import parse_prefix_mappings, select

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.customxml.template import FillResult, Skeleton

__all__ = [
    "BUILT_IN_NAMESPACES",
    "CustomXmlPart",
    "CustomXmlPartCollection",
    "custom_xml_parts_of",
    "normalise_id",
]

#: The namespaces Word's own property stores use. A part in one of them is
#: ``built_in``; the three well-known **store item ids** are deliberately *not*
#: special-cased (CR-003 section 4), so a binding to one resolves only when the
#: package really holds that part.
BUILT_IN_NAMESPACES: frozenset[str] = frozenset(
    {
        "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
        "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
        "http://schemas.microsoft.com/office/2006/coverPageProps",
        "http://purl.org/dc/elements/1.1/",
    }
)

#: ``ds:``, the custom XML data storage properties namespace.
_DS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"


def normalise_id(item_id: str | None) -> str:
    """A store item id with its braces off and lower-cased, for comparison."""
    return (item_id or "").replace("{", "").replace("}", "").lower()


class CustomXmlPart:
    """One ``/customXml/itemN.xml``. Office JS ``Word.CustomXmlPart``."""

    __slots__ = ("_prefixes", "collection", "id", "part", "schema_collection", "unlinked")

    def __init__(
        self,
        part: Any,
        collection: CustomXmlPartCollection,
        item_id: str,
        schema_refs: tuple[str, ...] = (),
    ) -> None:
        """Build the view over a data storage part and its properties part."""
        #: The engine part, a ``CustomXmlDataStoragePart``.
        self.part = part
        #: The collection it belongs to.
        self.collection = collection
        #: The ``ds:itemID``, exactly as the properties part spells it.
        self.id = item_id
        #: The ``ds:schemaRef`` URIs. Office JS ``schemaCollection``.
        self.schema_collection: tuple[str, ...] = tuple(schema_refs)
        self._prefixes: CustomXmlPrefixMappingCollection | None = None
        #: The XPaths of the mappings :meth:`delete` unlinked, once it has run.
        self.unlinked: tuple[str, ...] = ()

    # -- identity ----------------------------------------------------------

    def __eq__(self, other: object) -> bool:
        """Two views of the same part are equal (CR-003 section 3.1)."""
        return isinstance(other, CustomXmlPart) and other.part is self.part

    def __hash__(self) -> int:
        """Hashes by the part's identity."""
        return hash(id(self.part))

    def __repr__(self) -> str:
        """``<CustomXmlPart /customXml/item1.xml {5D7B…} 'invoice'>``."""
        root = self.document_element
        name = root.base_name if root is not None else ""
        return f"<CustomXmlPart {self.part.part_name} {self.id} {name!r}>"

    # -- what it holds -----------------------------------------------------

    @property
    def tree(self) -> Any:
        """The part's root element, parsed on first access. Extension."""
        return self.part.tree

    @property
    def document_element(self) -> CustomXmlNode | None:
        """The root node, or None for a part that will not parse. Office JS ``documentElement``."""
        try:
            root = self.part.tree
        except Exception:  # noqa: BLE001 - a part that will not parse has no document element
            return None
        return CustomXmlNode(root, self)

    @property
    def namespace_uri(self) -> str:
        """The document element's namespace, or ``""``. Office JS ``namespaceUri``.

        ``samples/invoice2013.docx``'s ``<invoice>`` has none, which is why its
        twenty bindings carry an empty ``w:prefixMappings``.
        """
        root = self.document_element
        return root.namespace_uri if root is not None else ""

    @property
    def built_in(self) -> bool:
        """Whether this is one of Word's own property stores. Office JS ``builtIn``.

        By **namespace**, as the TypeScript engine decided: the three well-known
        store item ids are not special-cased (CR-003 section 4).
        """
        return self.namespace_uri in BUILT_IN_NAMESPACES

    @property
    def namespace_manager(self) -> CustomXmlPrefixMappingCollection:
        """The part's own prefixes. Office JS ``namespaceManager``.

        Built once from the document element's declarations, with an unprefixed
        ``xmlns`` given the first free ``nsN`` name, because XPath 1.0 cannot use
        a default namespace and ``ns0`` is what Word writes.
        """
        if self._prefixes is None:
            root = self.document_element
            element = root.element if root is not None else None
            self._prefixes = CustomXmlPrefixMappingCollection(namespaces_of(element))
        return self._prefixes

    def get_xml(self) -> str:
        """The part as an XML string. Office JS ``getXml``.

        Synchronous and a plain string, where Office JS returns a
        ``ClientResult``: a custom XML part is a tree here, not something a
        marshaller has to be asked for (CR-003 section 3.7).
        """
        root = self.document_element
        if root is None:
            return ""
        return etree.tostring(root.element, encoding="unicode")

    def set_xml(self, xml: str) -> None:
        """Replace the whole part. Office JS ``setXml``.

        Raises:
            BindingError: `xml` is not well formed.
        """
        try:
            root = etree.fromstring(xml.encode("utf-8") if isinstance(xml, str) else xml)
        except Exception as error:
            raise BindingError(
                f"a custom XML part needs a well-formed document element: {error}",
                code="binding.not_well_formed",
                hint="pass the XML of the whole part, root element included",
            ) from error
        self.part.set_tree(root)
        self._prefixes = None

    def touch(self) -> None:
        """Mark the part for re-marshalling. What every mutation here calls."""
        mark = getattr(self.part, "mark_modified", None)
        if mark is not None:
            mark()

    @property
    def is_modified(self) -> bool:
        """Whether this part will be re-marshalled on save. Extension."""
        return bool(getattr(self.part, "is_modified", False))

    # -- selecting ---------------------------------------------------------

    def mappings_for(self, namespace_mappings: str | dict[str, str] | None) -> dict[str, str]:
        """The prefix mapping an XPath call should use.

        Word's ``xmlns:ns0='urn:invoice'`` string, a mapping as it is, or ---
        for None --- the part's own :attr:`namespace_manager`.
        """
        if namespace_mappings is None:
            return self.namespace_manager.to_map()
        if isinstance(namespace_mappings, dict):
            return dict(namespace_mappings)
        return parse_prefix_mappings(namespace_mappings)

    def select_from(
        self,
        context: Any,
        xpath: str,
        namespace_mappings: str | dict[str, str] | None = None,
    ) -> list[CustomXmlNode]:
        """Every node an XPath selects from a context element. Extension."""
        found = select(context, xpath, self.mappings_for(namespace_mappings))
        out: list[CustomXmlNode] = []
        for result in found:
            node = node_for(result, self)
            if node is not None:
                out.append(node)
        return out

    def select_nodes(
        self, xpath: str, namespace_mappings: str | dict[str, str] | None = None
    ) -> list[CustomXmlNode]:
        """Every node an XPath selects. Office JS ``selectNodes``."""
        root = self.document_element
        if root is None:
            return []
        return self.select_from(root.element, xpath, namespace_mappings)

    def select_single_node(
        self, xpath: str, namespace_mappings: str | dict[str, str] | None = None
    ) -> CustomXmlNode | None:
        """The first node an XPath selects, or None. Office JS ``selectSingleNode``."""
        found = self.select_nodes(xpath, namespace_mappings)
        return found[0] if found else None

    def _require(
        self, xpath: str, namespace_mappings: str | dict[str, str] | None, *, one: bool = True
    ) -> CustomXmlNode:
        return require_node(
            self.select_nodes(xpath, namespace_mappings),
            xpath,
            f"custom XML part {self.part.part_name}",
            one=one,
        )

    # -- editing through an XPath ------------------------------------------

    def insert_element(
        self,
        xpath: str,
        xml: str,
        namespace_mappings: str | dict[str, str] | None = None,
        index: int | None = None,
    ) -> CustomXmlNode:
        """Add a child to the element an XPath selects. Office JS ``insertElement``.

        `namespace_mappings` comes **third**, before `index`, so that the common
        call --- a part with one namespace, or none --- is
        ``insert_element("/invoice/lines", "<lineitem/>")``. Office JS's
        desktop-only signature puts it second; the departure is recorded in
        CR-003 section 17 and the member is excluded from the subset test's
        argument check.
        """
        return self._require(xpath, namespace_mappings).append_child_node(xml, index=index)

    def update_element(
        self, xpath: str, xml: str, namespace_mappings: str | dict[str, str] | None = None
    ) -> CustomXmlNode:
        """Replace the element an XPath selects. Office JS ``updateElement``."""
        node = self._require(xpath, namespace_mappings)
        node.set_xml(xml)
        return node

    def delete_element(
        self, xpath: str, namespace_mappings: str | dict[str, str] | None = None
    ) -> None:
        """Remove the element an XPath selects. Office JS ``deleteElement``."""
        self._require(xpath, namespace_mappings).delete()

    def insert_attribute(
        self,
        xpath: str,
        name: str,
        value: str,
        namespace_mappings: str | dict[str, str] | None = None,
    ) -> CustomXmlNode:
        """Set an attribute on the element an XPath selects. Office JS ``insertAttribute``."""
        return self.update_attribute(xpath, name, value, namespace_mappings)

    def update_attribute(
        self,
        xpath: str,
        name: str,
        value: str,
        namespace_mappings: str | dict[str, str] | None = None,
    ) -> CustomXmlNode:
        """Set an attribute, whether or not it is there. Office JS ``updateAttribute``.

        A prefixed `name` is resolved through `namespace_mappings`, then through
        the part's own namespace manager; an unresolved prefix is written as the
        literal name, which is what a document that declares it inline expects.
        """
        node = self._require(xpath, namespace_mappings)
        if node.node_type != "Element":
            raise BindingError(
                f"{xpath!r} does not select an element",
                code="binding.not_an_element",
                hint="an attribute belongs on an element; select one",
            )
        qualified = self._qualify(name, namespace_mappings)
        node.element.set(qualified, value)
        self.touch()
        return CustomXmlNode(node.element, self, attribute=qualified)

    def delete_attribute(
        self, xpath: str, name: str, namespace_mappings: str | dict[str, str] | None = None
    ) -> None:
        """Remove an attribute. Office JS ``deleteAttribute``."""
        node = self._require(xpath, namespace_mappings)
        node.element.attrib.pop(self._qualify(name, namespace_mappings), None)
        node.element.attrib.pop(name, None)
        self.touch()

    def _qualify(self, name: str, namespace_mappings: str | dict[str, str] | None) -> str:
        prefix, _, local = name.partition(":")
        if not local:
            return name
        uri = self.mappings_for(namespace_mappings).get(prefix) or (
            self.namespace_manager.lookup_namespace(prefix)
        )
        return f"{{{uri}}}{local}" if uri else name

    # -- the part itself ---------------------------------------------------

    def delete(self) -> None:
        """Remove the part, its properties part and the mappings that name it.

        Office JS ``delete``. Only the mappings this call can **reach** are
        unlinked: the controls of the parts already unmarshalled. A document
        whose main part was never read keeps a ``w:dataBinding`` naming a part
        that is gone, which Word reads as an unbound control (CR-003 section 17).
        """
        self.collection.remove(self)

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary: what a tool result says about a part."""
        return {
            "id": self.id,
            "part_name": str(self.part.part_name),
            "namespace_uri": self.namespace_uri,
            "built_in": self.built_in,
            "schema_collection": list(self.schema_collection),
        }


# ---------------------------------------------------------------------------
# the collection
# ---------------------------------------------------------------------------


class CustomXmlPartCollection:
    """``pkg.custom_xml_parts``. Office JS ``Word.CustomXmlPartCollection``."""

    __slots__ = ("_views", "package")

    def __init__(self, package: Any) -> None:
        """Build the collection over a package; nothing is parsed yet."""
        #: The package --- a ``WordprocessingMLPackage`` or a trial one.
        self.package = package
        self._views: dict[int, CustomXmlPart] = {}

    # -- the parts ---------------------------------------------------------

    @property
    def items(self) -> list[CustomXmlPart]:
        """Every custom XML part, in the order the package holds them."""
        out: list[CustomXmlPart] = []
        for part in self._storage_parts():
            out.append(self._view_for(part))
        return out

    def _storage_parts(self) -> list[Any]:
        """The engine parts, a trial's copies where there is a trial."""
        found = getattr(self.package, "custom_xml_data_storage_parts", None) or {}
        trial_part = getattr(self.package, "trial_part", None)
        parts = []
        for part in found.values():
            parts.append(trial_part(part) if trial_part is not None else part)
        return parts

    def _view_for(self, part: Any) -> CustomXmlPart:
        key = id(part)
        view = self._views.get(key)
        if view is None:
            item_id, schema_refs = _properties_of(part)
            view = CustomXmlPart(part, self, item_id, schema_refs)
            self._views[key] = view
        return view

    def get_item(self, item_id: str) -> CustomXmlPart | None:
        """The part with a store item id, or None. Office JS ``getItem``.

        Case-insensitive and brace-insensitive both ways, because Word writes
        ``{5D7BA57F-…}`` in one binding and ``{5d7ba57f-…}`` in the next.
        Office JS hands out a proxy that errors later; here it is None.
        """
        wanted = normalise_id(item_id)
        for part in self.items:
            if normalise_id(part.id) == wanted:
                return part
        return None

    def get_by_namespace(self, namespace_uri: str) -> list[CustomXmlPart]:
        """Every part whose document element is in a namespace. Office JS ``getByNamespace``."""
        return [part for part in self.items if part.namespace_uri == namespace_uri]

    def __len__(self) -> int:
        """How many custom XML parts the document has."""
        return len(self._storage_parts())

    def __iter__(self) -> Any:
        """Iterates the parts."""
        return iter(self.items)

    def __repr__(self) -> str:
        """``<CustomXmlPartCollection 1 part>``."""
        count = len(self)
        return f"<CustomXmlPartCollection {count} part{'' if count == 1 else 's'}>"

    # -- adding and removing -----------------------------------------------

    def add(self, xml: str, *, schema_refs: list[str] | None = None) -> CustomXmlPart:
        """Add a custom XML part and its properties part. Office JS ``add``.

        docx4j's ``AbstractMigrator.addPropertiesPart``, step for step: the data
        part goes in as ``/customXml/itemN.xml`` with the first free N, the
        properties part beside it as ``itemPropsN.xml``, the ``ds:itemID`` is a
        brace-wrapped **upper-case** UUID, ``ds:schemaRefs`` carries the document
        element's namespace unless `schema_refs` says otherwise, and the
        relationship is written **from the main document part** --- Word drops a
        custom XML part that only the package relates to.

        The UUID is drawn from ``pkg.id_generator()``, so a seeded package makes
        the same bytes twice (CR-003 section 3.4). In a
        :func:`~docx4j_py.model.content.trial.dry_run` the part, its properties
        part, both relationships and both content types are **undone** on the way
        out (CR-003 section 14.4).

        Raises:
            BindingError: `xml` has no document element.
            ContentError: the package has no main document part to relate from.
        """
        from docx4j_py.model.content.picture import note_added_part
        from docx4j_py.openpackaging.parts.default_xml_part import (
            CustomXmlDataStoragePart,
        )
        from docx4j_py.openpackaging.parts.docprops import (
            CustomXmlDataStoragePropertiesPart,
        )
        from docx4j_py.openpackaging.parts.relationships_part import AddPartBehaviour

        try:
            root = etree.fromstring(xml.encode("utf-8") if isinstance(xml, str) else xml)
        except Exception as error:
            raise BindingError(
                f"a custom XML part needs a document element: {error}",
                code="binding.not_well_formed",
                hint="pass the XML of the whole part, root element included",
            ) from error

        package = getattr(self.package, "_package", self.package)
        source = package.main_document_part
        if source is None:
            raise ContentError(
                "this package has no main document part to relate a custom XML part from",
                code="binding.no_main_part",
                hint="Word drops a custom XML part that only the package relates to",
            )

        data_part = CustomXmlDataStoragePart(_free_item_name(package))
        data_part.set_tree(root)
        had_data_type = package.content_type_manager.get_override_content_type(data_part.part_name)
        data_rel = source.add_target_part(data_part, AddPartBehaviour.RENAME_IF_NAME_EXISTS)

        item_id = "{" + _uuid(self.package) + "}"
        props = CustomXmlDataStoragePropertiesPart(_props_name(data_part.part_name))
        props.set_tree(_datastore_item(item_id, schema_refs, root))
        had_props_type = package.content_type_manager.get_override_content_type(props.part_name)
        props_rel = data_part.add_target_part(props, AddPartBehaviour.RENAME_IF_NAME_EXISTS)

        data_part.item_id = item_id.lower()
        package.custom_xml_data_storage_parts[data_part.item_id] = data_part

        note_added_part(
            self.package, props, props_rel, data_part, added_content_type=had_props_type is None
        )
        note_added_part(
            self.package, data_part, data_rel, source, added_content_type=had_data_type is None
        )

        view = CustomXmlPart(data_part, self, item_id, tuple(_schema_refs(schema_refs, root)))
        self._views[id(data_part)] = view
        return view

    def remove(self, view: CustomXmlPart) -> None:
        """Remove a part and its properties part; :meth:`CustomXmlPart.delete` is the name.

        Raises:
            ContentError: in a dry run, which cannot un-delete a part.
        """
        if getattr(self.package, "_package", None) is not None:
            raise ContentError(
                "a dry run cannot un-delete a custom XML part",
                code="dry_run.delete_part",
                hint="delete the part on the real package, or preview a fill instead",
            )
        package = self.package
        part = getattr(view.part, "_wrapped", view.part)
        unlinked = self._unlink(view)
        owner = part.owning_relationship_part
        if owner is None:
            main = package.main_document_part
            owner = main.relationships_part if main is not None else None
        removed = owner.remove_part(part.part_name) if owner is not None else []
        for name in removed:
            package.content_type_manager.remove_override_content_type(name)
        package.custom_xml_data_storage_parts.pop(getattr(part, "item_id", None) or "", None)
        self._views.pop(id(view.part), None)
        view.unlinked = tuple(unlinked)

    def _unlink(self, view: CustomXmlPart) -> list[str]:
        """Delete the mappings of every control that names this part, and say which."""
        from docx4j_py.model.customxml.bindings import controls_of

        wanted = normalise_id(view.id)
        out: list[str] = []
        for control in controls_of(self.package):
            mapping = control.xml_mapping
            if mapping.is_mapped and normalise_id(mapping.store_item_id) == wanted:
                out.append(mapping.xpath)
                mapping.delete()
        return out

    # -- the bindings ------------------------------------------------------

    def apply_bindings(self) -> Any:
        """Push the custom XML into the controls. docx4j ``BindingHandler.applyBindings``."""
        from docx4j_py.model.customxml.bindings import apply_bindings

        return apply_bindings(self.package)

    def update_from_content_controls(self) -> Any:
        """Write the controls back into the custom XML. docx4j ``UpdateXmlFromDocumentSurface``."""
        from docx4j_py.model.customxml.bindings import update_from_content_controls

        return update_from_content_controls(self.package)

    def describe(self) -> Skeleton:
        """The data this template wants: docx4j-mcp's ``describe_template``."""
        from docx4j_py.model.customxml.template import describe_template

        return describe_template(self.package)

    def fill(self, data: dict[str, Any] | str) -> FillResult:
        """Set the nodes and apply the bindings: docx4j-mcp's ``fill_template``."""
        from docx4j_py.model.customxml.template import fill_template

        return fill_template(self.package, data)


# ---------------------------------------------------------------------------
# the properties part, read without unmarshalling anything
# ---------------------------------------------------------------------------


def _properties_of(part: Any) -> tuple[str, tuple[str, ...]]:
    """A data part's ``ds:itemID`` and ``ds:schemaRef`` URIs.

    Read off a throwaway parse of the properties part's **bytes**, as the loader
    does, so that neither part is unmarshalled and the properties part goes out
    exactly as it came in.
    """
    from docx4j_py.openpackaging.parts.namespaces import Namespaces

    real = getattr(part, "_wrapped", part)
    fallback = (getattr(real, "item_id", None) or "").strip()
    rels = getattr(real, "relationships_part", None)
    if rels is None:
        return fallback, ()
    rel = rels.get_relationship_by_type(Namespaces.CUSTOM_XML_DATA_STORAGE_PROPERTIES)
    props = rels.get_part(rel) if rel is not None else None
    if props is None:
        return fallback, ()
    try:
        root = etree.fromstring(props.xml)
    except Exception:  # noqa: BLE001 - a broken properties part describes nothing
        return fallback, ()
    item_id = root.get(f"{{{_DS}}}itemID") or root.get("itemID") or fallback
    refs = tuple(
        node.get(f"{{{_DS}}}uri") or node.get("uri") or ""
        for node in root.iter(f"{{{_DS}}}schemaRef")
    )
    return item_id, tuple(uri for uri in refs if uri)


def _datastore_item(item_id: str, schema_refs: list[str] | None, root: Any) -> Any:
    """The ``ds:datastoreItem`` of a new properties part."""
    element = etree.Element(f"{{{_DS}}}datastoreItem", nsmap={"ds": _DS})
    element.set(f"{{{_DS}}}itemID", item_id)
    refs = _schema_refs(schema_refs, root)
    container = etree.SubElement(element, f"{{{_DS}}}schemaRefs")
    for uri in refs:
        ref = etree.SubElement(container, f"{{{_DS}}}schemaRef")
        ref.set(f"{{{_DS}}}uri", uri)
    return element


def _schema_refs(schema_refs: list[str] | None, root: Any) -> list[str]:
    """The URIs a new properties part declares: the caller's, or the root's namespace."""
    if schema_refs is not None:
        return list(schema_refs)
    tag = getattr(root, "tag", "")
    if isinstance(tag, str) and tag.startswith("{"):
        return [tag[1:].partition("}")[0]]
    return []


def _free_item_name(package: Any) -> str:
    """``/customXml/itemN.xml``, the first N free (docx4j's ``RENAME_IF_NAME_EXISTS``)."""
    from docx4j_py.openpackaging.part_name import PartName

    index = 1
    while package.get_part(PartName.of(f"/customXml/item{index}.xml")) is not None:
        index += 1
    return f"/customXml/item{index}.xml"


def _props_name(data_name: Any) -> str:
    """``/customXml/itemPropsN.xml`` for ``/customXml/itemN.xml``."""
    import re

    match = re.search(r"item(\d*)\.xml$", str(data_name))
    return f"/customXml/itemProps{match.group(1) if match else ''}.xml"


def _uuid(package: Any) -> str:
    """A brace-free upper-case UUID from the package's generator (CR-003 section 3.4)."""
    import uuid as uuid_module

    generator = getattr(package, "id_generator", None)
    if generator is None:
        return str(uuid_module.uuid4()).upper()
    rng = generator(derive_from=("custom-xml",))
    return str(uuid_module.UUID(int=rng.getrandbits(128), version=4)).upper()


def custom_xml_parts_of(package: Any) -> CustomXmlPartCollection:
    """The package's collection, made once and kept on it."""
    found = getattr(package, "_custom_xml_parts", None)
    if found is None:
        found = CustomXmlPartCollection(package)
        package._custom_xml_parts = found
    return found
