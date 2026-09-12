"""``XmlPart[T]``: a part whose content is a typed object. docx4j ``JaxbXmlPart<E>``.

CR-002 section 5.4 and 5.6. ``T`` is the root class from the generated model
(``Document``, ``Styles``, ``Settings``, ...).

Three rules, and they are the whole of the part's behaviour:

**Lazy.** ``part.contents`` unmarshals on first access and not before, exactly as
docx4j's ``getContents()`` does. Loading a package therefore costs the zip's
central directory and the relationship parts, and nothing else.

**Never silent.** The unmarshal uses a *per-call* :class:`ParserConfig` --- one
per thread, CR-001 section 14 --- that is lenient and carries the fork's
skipped-content report. What lenience cost is kept on the part as
:attr:`skipped` and logged at warning level with the part's name;
``LoadOptions(strict=True)`` raises instead.

**Untouched means byte for byte.** :attr:`bytes_for_save` returns the source
bytes unless the part was unmarshalled or had bytes set on it. Only a part the
caller actually touched is re-serialised, which is what makes a round trip of a
document full of markup this model has never seen safe (CR-002 section 1).

Serialisation is CR-002 section 5.6's other half: the prefix table, numeric
booleans, and a root that declares every prefix ``mc:Ignorable`` names --- from
the table where it knows one, and from the *source* root's declarations where it
does not, which is why the part remembers them.
"""

from __future__ import annotations

import logging
from typing import Any, Generic, TypeVar

from lxml import etree

from docx4j_py.namespaces import PREFIXES
from docx4j_py.openpackaging.exceptions import Docx4JException
from docx4j_py.openpackaging.part_name import PartName
from docx4j_py.openpackaging.parts.part import Part

__all__ = ["XML_DECLARATION", "XmlPart"]

T = TypeVar("T")

logger = logging.getLogger("docx4j_py.openpackaging")

#: What Word writes at the top of every XML part, CRLF and all.
XML_DECLARATION = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_MC_IGNORABLE = f"{{{_MC_NS}}}Ignorable"

_LXML_PARSER = etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)

#: The serialiser's prefix table: docx4j's, minus ``xml``, which is bound by
#: definition and invalid to declare.
_NS_MAP: dict[str, str] = {p: u for p, u in PREFIXES.items() if p != "xml"}


class XmlPart(Part, Generic[T]):
    """A part whose content is a typed object of the model."""

    __slots__ = (
        "_bytes",
        "_contents",
        "_options",
        "_source_ns_map",
        "root_name",
        "skipped",
    )

    def __init__(
        self,
        part_name: PartName | str,
        content_type: str = "",
        relationship_type: str = "",
        root_name: str | None = None,
    ) -> None:
        """Build the part; `root_name` is the qualified name of its root element."""
        super().__init__(part_name, content_type, relationship_type)
        self._contents: Any = None
        self._bytes: bytes | None = None
        #: ``{namespace}localName`` of the root, when the class alone does not say.
        self.root_name = root_name
        self._source_ns_map: dict[str, str] = {}
        #: What the fork's parser skipped, per entry. Empty is the good answer.
        self.skipped: list[Any] = []

    # -- the class to unmarshal as -----------------------------------------

    #: The dotted path of the model class this part's content is, as a
    #: **string**: ``"docx4j_py.wml.Document"``. A string rather than the class
    #: itself because importing anything under ``docx4j_py`` imports the whole
    #: generated model (CR-001 section 13.5 point 3, 0.6 s), and the engine must
    #: import in milliseconds on top of it. :attr:`model_class` resolves it on
    #: first use and caches it on the class.
    model_class_path: str | None = None

    #: How ``xsd:boolean`` is spelled when this part is written. ``"numeric"``
    #: is CR-001 decided question 5 and is right for WordprocessingML, where
    #: Word writes ``w:val="1"``. The document-properties parts override it with
    #: ``"words"``, because Word writes ``<ScaleCrop>false</ScaleCrop>`` there:
    #: the two spellings are the same ``xsd:boolean``, and matching Word per
    #: part is what makes a round trip byte-comparable rather than merely
    #: equivalent.
    bool_format: str = "numeric"

    #: A namespace to bind as the *default* (prefix-less) one when this part is
    #: written, instead of taking its prefix from the table. Only the
    #: relationships parts use it, because Word writes
    #: ``<Relationships xmlns="...">`` and not ``<rel:Relationships>``
    #: (CR-002 section 5.2).
    default_namespace: str | None = None

    @property
    def model_class(self) -> type | None:
        """The model class, imported on first use. None for a generic part."""
        cls = type(self)
        cached = cls.__dict__.get("_model_class")
        if cached is not None:
            return cached
        path = self.model_class_path
        if path is None:
            return None
        from importlib import import_module

        module_name, _, class_name = path.rpartition(".")
        value = getattr(import_module(module_name), class_name)
        cls._model_class = value  # type: ignore[attr-defined]
        return value

    # -- contents ----------------------------------------------------------

    @property
    def is_unmarshalled(self) -> bool:
        """True once the contents were unmarshalled or set. docx4j ``isUnmarshalled``."""
        return self._contents is not None

    @property
    def is_loaded(self) -> bool:
        """True when the part carries content of its own, tree or bytes."""
        return self._contents is not None or self._bytes is not None

    @property
    def contents(self) -> T:
        """The typed content, unmarshalled on first access. docx4j ``getContents()``.

        Raises:
            Docx4JException: the part has no content and no source container, or
                no model class to unmarshal as, or `strict` was asked for and
                something was skipped.
        """
        if self._contents is None:
            self._unmarshal()
        return self._contents  # type: ignore[return-value]

    def get_contents(self) -> T:
        """docx4j's spelling of :attr:`contents`, for code that reads across."""
        return self.contents

    def set_contents(self, value: T) -> None:
        """Replace the content and wire its parent pointers. docx4j ``setContents``."""
        from docx4j_py.child import link_parents
        from docx4j_py.runtime import context

        self._contents = value
        self._bytes = None
        if value is not None:
            link_parents(value, context=context())

    def set_bytes(self, data: bytes) -> None:
        """Replace the content with raw XML; unmarshalled on the next access."""
        self._bytes = bytes(data)
        self._contents = None
        self.skipped = []

    def set_xml(self, xml: str) -> None:
        """Replace the content with an XML string."""
        self.set_bytes(xml.encode("utf-8"))

    # -- unmarshalling -----------------------------------------------------

    def _raw_bytes(self) -> bytes:
        data = self._bytes if self._bytes is not None else self._source_bytes()
        if data is None:
            raise Docx4JException(
                f"Part {self.part_name} has no content: set contents, or load it from a container"
            )
        return data

    def _unmarshal(self) -> None:
        from docx4j_py.child import link_parents
        from docx4j_py.runtime import context, parser

        clazz = self.model_class
        if clazz is None:
            raise Docx4JException(
                f"Part {self.part_name} ({self.content_type}) has no model class; "
                "it is a generic XML part, use .tree"
            )
        data = self._raw_bytes()
        self._source_ns_map = _root_ns_map(data)

        options = self._load_options()
        if options is not None and options.mce_preprocess:
            from docx4j_py.openpackaging.mce import mce_preprocess_bytes

            data = mce_preprocess_bytes(data)

        xml_parser = parser(lenient=True, skipped_report=True)
        try:
            obj = xml_parser.from_bytes(data, clazz)
        except Exception as exc:  # noqa: BLE001 - the part name is the useful half
            raise Docx4JException(f"Problem with part {self.part_name}: {exc}") from exc

        report = xml_parser.skipped
        self.skipped = list(getattr(report, "items", ()) or ())
        if self.skipped:
            strict = options is not None and options.strict
            for item in self.skipped:
                logger.warning("%s: skipped %s", self.part_name, item)
            if strict:
                raise Docx4JException(
                    f"{self.part_name}: {len(self.skipped)} item(s) were skipped "
                    f"and strict=True: {self.skipped[:5]}"
                )
        link_parents(obj, context=context())
        self._contents = obj
        self._bytes = None

    def _load_options(self) -> Any:
        package = self.package
        return getattr(package, "load_options", None) if package is not None else None

    # -- serialising -------------------------------------------------------

    @property
    def xml(self) -> bytes:
        """The bytes as they will be written. docx4j's part of ``Save``.

        The source bytes when the part was never touched; otherwise the tree
        re-serialised with the prefix table, numeric booleans and the
        ``mc:Ignorable`` reconciliation of CR-002 section 5.6.
        """
        if self._contents is not None:
            return self._marshal(self._contents)
        if self._bytes is not None:
            return self._bytes
        data = self._source_bytes()
        if data is None:
            raise Docx4JException(f"Part {self.part_name} has no content and no source container")
        return data

    @property
    def bytes_for_save(self) -> bytes:
        """:attr:`xml`."""
        return self.xml

    def get_xml(self) -> str:
        """:attr:`xml` decoded, for reading and for tests."""
        return self.xml.decode("utf-8")

    def _marshal(self, obj: Any) -> bytes:
        from docx4j_py.runtime import serializer

        ns_map = _NS_MAP
        if self.default_namespace is not None:
            ns_map = {p: u for p, u in _NS_MAP.items() if u != self.default_namespace}
            ns_map[None] = self.default_namespace  # type: ignore[index]
        text = serializer(bool_format=self.bool_format).render(obj, ns_map)
        root = etree.fromstring(text.encode("utf-8"), _LXML_PARSER)

        wanted = self.root_name or _element_name(obj)
        if wanted is not None and root.tag != wanted:
            root = _rename(root, wanted)

        # Drop what the tree does not use, then put back the prefixes
        # mc:Ignorable names --- they must be declared even though nothing in
        # the tree is in them, or Word rejects the file.
        etree.cleanup_namespaces(root)
        ns_map = dict(root.nsmap)
        for prefix in _ignorable_prefixes(root):
            if prefix in ns_map:
                continue
            uri = PREFIXES.get(prefix) or self._source_ns_map.get(prefix)
            if uri is not None:
                ns_map[prefix] = uri
            else:
                logger.warning(
                    "%s: mc:Ignorable names the prefix %r, which neither the "
                    "prefix table nor the source declared; it is left undeclared",
                    self.part_name,
                    prefix,
                )
        if ns_map != root.nsmap:
            root = _redeclare(root, ns_map)

        return XML_DECLARATION + etree.tostring(root, encoding="utf-8")

    # -- introspection -----------------------------------------------------

    @property
    def source_ns_map(self) -> dict[str, str]:
        """The namespace declarations the source root carried, prefix -> URI."""
        return dict(self._source_ns_map)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _element_name(obj: Any) -> str | None:
    from docx4j_py.traversal import element_name

    return element_name(obj)


def _root_ns_map(data: bytes) -> dict[str, str]:
    """The namespace declarations on a part's root element.

    Parsed with ``iterparse`` and abandoned at the root's start tag, so this
    costs the root element and not the document.
    """
    import io

    try:
        for _event, element in etree.iterparse(
            io.BytesIO(data),
            events=("start",),
            resolve_entities=False,
            huge_tree=True,
            recover=True,
        ):
            return {p: u for p, u in element.nsmap.items() if p is not None}
    except etree.XMLSyntaxError:
        return {}
    return {}


def _ignorable_prefixes(root: etree._Element) -> set[str]:
    """Every prefix named by an ``mc:Ignorable`` anywhere in the tree."""
    found: set[str] = set()
    for node in root.iter():
        if not isinstance(node.tag, str):
            continue
        value = node.get(_MC_IGNORABLE)
        if value:
            found.update(value.split())
    return found


def _rename(root: etree._Element, name: str) -> etree._Element:
    """A copy of `root` under a different element name, children moved over."""
    new = etree.Element(name, nsmap=root.nsmap)
    for key, value in root.attrib.items():
        new.set(key, value)
    new.text = root.text
    for child in list(root):
        new.append(child)
    return new


def _redeclare(root: etree._Element, ns_map: dict[str, str]) -> etree._Element:
    """A copy of `root` whose namespace declarations are exactly `ns_map`."""
    new = etree.Element(root.tag, nsmap=ns_map)
    for key, value in root.attrib.items():
        new.set(key, value)
    new.text = root.text
    for child in list(root):
        new.append(child)
    return new
