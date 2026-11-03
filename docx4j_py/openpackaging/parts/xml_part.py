"""``XmlPart[T]``: a part whose content is a typed object. docx4j ``JaxbXmlPart<E>``.

CR-002 section 5.4 and 5.6. ``T`` is the root class from the generated model
(``Document``, ``Styles``, ``Settings``, ...).

Three rules, and they are the whole of the part's behaviour:

**Lazy.** ``part.contents`` unmarshals on first access and not before, exactly as
docx4j's ``getContents()`` does. Loading a package therefore costs the zip's
central directory and the relationship parts, and nothing else. A part that
``read_contents()`` has already parsed is not parsed again: ``contents`` adopts
that tree (CR-002 section 16).

**Never silent.** The unmarshal uses a *per-call* :class:`ParserConfig` --- one
per thread, CR-001 section 14 --- that is lenient and carries the fork's
skipped-content report. What lenience cost is kept on the part as
:attr:`skipped` and logged at warning level with the part's name;
``LoadOptions(strict=True)`` raises instead. Where the report names a ``w:br``
outside a ``w:r``, or an element inside a text element, the part is read as
Word reads it (CR-030, :mod:`~docx4j_py.openpackaging.breaks`): the bytes are
repaired and parsed again, and the part's :attr:`warnings` says so.

**Refused, not mended.** A part whose XML is not well formed raises
``Docx4JException`` naming it and lxml's reason when its content is first read
(CR-031, docx4j CR-065 D1): the fork's parser never recovers
(:data:`docx4j_py.runtime.RECOVER`), and the failure is kept on the part
(:attr:`XmlPart.unreadable`), raised again by every later read without another
parse. The load stays lazy, and a save writes such a part's bytes as read.

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
from dataclasses import dataclass
from typing import Any, Generic, TypeVar

from lxml import etree

from docx4j_py.namespaces import PREFIXES
from docx4j_py.openpackaging.exceptions import Docx4JException
from docx4j_py.openpackaging.part_name import PartName
from docx4j_py.openpackaging.parts.part import Part

__all__ = ["XML_DECLARATION", "XmlPart", "syntax_error_of"]

T = TypeVar("T")

logger = logging.getLogger("docx4j_py.openpackaging")

#: What Word writes at the top of every XML part, CRLF and all.
XML_DECLARATION = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_MC_IGNORABLE = f"{{{_MC_NS}}}Ignorable"

#: The serialiser's prefix table: docx4j's, minus ``xml``, which is bound by
#: definition and invalid to declare.
_NS_MAP: dict[str, str] = {p: u for p, u in PREFIXES.items() if p != "xml"}


@dataclass(frozen=True, slots=True)
class _ParseSettings:
    """What a parse of a part reads from its package. CR-002 section 16.2.

    :meth:`XmlPart._unmarshal` reads the package's load options through this
    and nothing else, and a private tree keeps the settings it was parsed
    under, so that :attr:`XmlPart.contents` adopts it only where a live parse
    now would read the same. CR-012's ``mc_load_mode``, on the load options
    or on the package, is the third field (CR-012 Phase A): a tree parsed in
    one mode is never adopted in the other.
    """

    #: ``LoadOptions.strict``: a parse that skipped content raises.
    strict: bool = False
    #: ``LoadOptions.mce_preprocess``: ``mc:AlternateContent`` resolved first.
    mce_preprocess: bool = False
    #: ``pkg.mc_load_mode`` (CR-012 Phase A): ``"preserve"`` or ``"word"``.
    mc_load_mode: str = "preserve"

    @classmethod
    def of(cls, package: Any) -> _ParseSettings:
        """What a parse of a part of `package` reads now; the defaults for a part with no package."""
        options = getattr(package, "load_options", None) if package is not None else None
        if options is None:
            return _DEFAULT_SETTINGS
        mode = getattr(package, "mc_load_mode", None) or getattr(options, "mc_load_mode", "preserve")
        return cls(
            strict=bool(options.strict),
            mce_preprocess=bool(options.mce_preprocess),
            mc_load_mode=mode,
        )


_DEFAULT_SETTINGS = _ParseSettings()


class XmlPart(Part, Generic[T]):
    """A part whose content is a typed object of the model."""

    __slots__ = (
        "_bytes",
        "_contents",
        "_empty_value",
        "_options",
        "_private_contents",
        "_private_settings",
        "_source_ns_map",
        "_unreadable",
        "_unreadable_settings",
        "kept",
        "resolved_alternate_content",
        "root_name",
        "skipped",
        "warnings",
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
        self._private_contents: Any = None
        #: The settings the private tree was parsed under (CR-002 section 16.2).
        self._private_settings: _ParseSettings | None = None
        self._bytes: bytes | None = None
        #: The ``Docx4JException`` of the read that failed, and the settings it was read under (CR-031 D8).
        self._unreadable: Docx4JException | None = None
        self._unreadable_settings: _ParseSettings | None = None
        #: ``{namespace}localName`` of the root, when the class alone does not say.
        self.root_name = root_name
        self._source_ns_map: dict[str, str] = {}
        #: Whether the last parse's source holds an empty attribute value (``=""``): an empty
        #: rsid binds without a record, so this, as well as :attr:`kept`, sends the part through
        #: the save's typed rows (CR-021 D6, :mod:`~docx4j_py.openpackaging.normalise`).
        self._empty_value = False
        #: What the fork's parser skipped, per entry. Empty is the good answer.
        self.skipped: list[Any] = []
        #: The values the last parse bound as the text they were read with (CR-018 Q16): the fork's
        #: ``SkippedReport.kept``, each a ``KeptValue`` (``reason`` ``"exact"`` for a value the exact
        #: parse kept on purpose, such as ``w:val=" false"``; ``"invalid"`` for one no rule converts,
        #: such as ``"off"`` or ``"Yellow"``). Not skipped content: the text is in the tree, a save
        #: writes it back (normalised where CR-018 section 4.7 says, :mod:`~docx4j_py.openpackaging.normalise`),
        #: and :mod:`docx4j_py.model.values` reads it as Word does. Empty is the usual answer.
        self.kept: list[Any] = []
        #: Each ``mc:AlternateContent`` the last parse took out in ``"word"`` mode
        #: (CR-012 Phase A), outermost first: a
        #: :class:`~docx4j_py.openpackaging.mc_load_mode.ResolvedAlternateContent`,
        #: the element whole and where it stood. Empty in ``"preserve"`` mode.
        #: docx4j ``JaxbXmlPart.getResolvedAlternateContent()``.
        self.resolved_alternate_content: list[Any] = []
        #: What the last parse read otherwise than as written, one record a kind, at most one a part (CR-030 Q2):
        #: a :class:`~docx4j_py.openpackaging.breaks.BrOutsideRun` where it found a ``w:br`` outside a ``w:r``,
        #: or a text element holding elements, and read them as Word reads them. Set by every parse, as
        #: :attr:`skipped` is, cleared by :meth:`set_bytes`; empty is the usual answer. core-ts's
        #: ``XmlPart.warnings``.
        self.warnings: list[Any] = []

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
    def unreadable(self) -> Docx4JException | None:
        """Why the part's content could not be read, or None. docx4j ``JaxbXmlPart``'s ``unreadable``.

        The ``Docx4JException`` the read that failed raised (CR-031 D8): its
        XML is not well formed ("Problem with part /word/document.xml: Premature
        end of data in tag t line 2, line 2, column 351", lxml's
        ``XMLSyntaxError`` the deepest ``__cause__``), or the parse failed
        otherwise. A later :attr:`contents` or :meth:`read_contents` raises it
        again without parsing, unless the parse settings have changed since
        (``strict``, ``mce_preprocess``, ``mc_load_mode``); :meth:`set_contents`,
        :meth:`set_bytes` and :meth:`set_xml` clear it. None for a part never
        read, one read, and one whose content was replaced.
        """
        return self._unreadable

    @property
    def is_content_unreadable(self) -> bool:
        """Whether the last read of the content failed. docx4j ``isContentUnreadable()``.

        Tells a part that could not be read from one with no content yet (a
        part made in code before :meth:`set_contents`), for which
        :attr:`contents` raises too.
        """
        return self._unreadable is not None

    @property
    def contents(self) -> T:
        """The typed content, unmarshalled on first access. docx4j ``getContents()``.

        Where :meth:`read_contents` has already parsed the part, under the load
        options a parse now would use, that private tree is **adopted** rather
        than the bytes parsed a second time (CR-002 section 16): it becomes the
        live tree, the part is unmarshalled from then on, and the private slot
        and any set bytes are cleared, exactly as an unmarshal leaves them. Its
        parent pointers are wired, and :attr:`skipped` and the source's
        namespace declarations are what that parse of the same bytes found. A
        private tree parsed under other options (``strict``, ``mce_preprocess``,
        the package's ``mc_load_mode``) is dropped and the bytes are parsed
        afresh.

        Adoption depends on one rule: **a private reader never writes into the
        tree it reads**, since what it wrote would now be saved.
        ``tests/openpackaging/test_private_readers.py`` holds every reader to it.

        Raises:
            Docx4JException: the part has no content and no source container, or
                no model class to unmarshal as, or `strict` was asked for and
                something was skipped; or its content cannot be read, its XML
                not well formed (CR-031): :attr:`unreadable`, raised again
                without a parse.
        """
        if self._contents is None:
            if self._private_contents is not None and self._private_settings == self._parse_settings():
                self._adopt()
            else:
                self._unmarshal()
        return self._contents  # type: ignore[return-value]

    def get_contents(self) -> T:
        """docx4j's spelling of :attr:`contents`, for code that reads across."""
        return self.contents

    def read_contents(self) -> T:
        """The typed content, **without** marking the part unmarshalled.

        CR-002 section 6.1. A resolution utility --- the property resolver, the
        numbering definitions, the font selector --- has to read the styles,
        numbering, theme and font table parts of every document it is built
        over. Reading :attr:`contents` would mark each of them for
        re-marshalling, so a document whose styles nobody edited would lose its
        byte-for-byte round trip the moment something asked for an effective
        property.

        The tree is unmarshalled once into a private slot and cached there, with
        the load options it was parsed under. Once the part is unmarshalled for
        real, the live tree is what this returns. :attr:`contents` **adopts**
        the private tree as that live tree where the options are unchanged, so
        a read followed by an edit parses the part once (CR-002 section 16).

        So the tree this returns may become the one that is saved, and a reader
        must **never write into it**: what it wrote would be saved, as an edit
        the caller never made. Every caller in this package keeps the rule, and
        ``tests/openpackaging/test_private_readers.py`` holds each to it.
        """
        if self._contents is not None:
            self._private_contents = None
            self._private_settings = None
            return self._contents  # type: ignore[return-value]
        if self._private_contents is None:
            self._unmarshal(private=True)
        return self._private_contents  # type: ignore[return-value]

    def set_contents(self, value: T) -> None:
        """Replace the content and wire its parent pointers. docx4j ``setContents``."""
        from docx4j_py.child import link_parents
        from docx4j_py.runtime import context

        self._private_contents = None
        self._private_settings = None
        self._contents = value
        self._bytes = None
        self._unreadable = None
        self._unreadable_settings = None
        if value is not None:
            link_parents(value, context=context())

    def set_bytes(self, data: bytes) -> None:
        """Replace the content with raw XML; unmarshalled on the next access."""
        self._bytes = bytes(data)
        self._contents = None
        self._private_contents = None
        self._private_settings = None
        self.skipped = []
        self.kept = []
        self._empty_value = False
        self.resolved_alternate_content = []
        self.warnings = []
        self._unreadable = None
        self._unreadable_settings = None

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

    def source_tree(self) -> T | None:
        """A fresh tree of the part's **source** bytes, parsed as a parse now would, held by nothing.

        CR-020 Phase B: the save's check (:func:`docx4j_py.validator.check_for_save`)
        compares what a part holds now with what its source held, to tell what a
        session introduced from what the file already had. The part is not
        touched: not its tree, its private tree, its reports nor its bytes. None
        for a part with no source container (one made in code), and for a source
        that cannot be read, its XML not well formed (CR-031 D9): it held
        nothing readable, so every finding in the part now is the session's.
        """
        data = self._source_bytes()
        if data is None or self.model_class is None:
            return None
        try:
            obj, _parser, _resolved, _warnings = self._parse(data, self._parse_settings())
        except Docx4JException:
            return None
        return obj

    def _parse(self, data: bytes, settings: _ParseSettings) -> tuple[Any, Any, list[Any], list[Any]]:
        """Parse `data` under `settings`: the tree, the parser (its skipped report), the ``mc`` content resolved, and the warnings.

        After the ``mc`` passes and the fork's parse, a part whose skipped
        report names a ``w:br`` outside a ``w:r``, or an element in a text
        element, is read as Word reads it (CR-030 section 4.1): its bytes go
        through :func:`~docx4j_py.openpackaging.breaks.repair_bytes`, and the
        fork parses the result, whose tree and report are the part's. A clean
        part pays the test of an empty report.

        XML that is not well formed is refused by whichever parser meets it
        first, the ``mc`` passes' or the fork's, which never recovers (CR-031
        D2): ``Docx4JException`` "Problem with part <name>: <lxml's message>",
        lxml's ``XMLSyntaxError`` the deepest ``__cause__``.
        """
        from docx4j_py.runtime import parser

        clazz = self.model_class
        resolved: list[Any] = []
        try:
            if settings.mc_load_mode == "word":
                # CR-012 Phase A: block-, row- and cell-level mc:AlternateContent
                # replaced by the branch Word reads; a part without the element
                # pays one substring test
                from docx4j_py.openpackaging.mc_load_mode import splice_word

                data, resolved = splice_word(data)
            if settings.mce_preprocess:
                from docx4j_py.openpackaging.mce import mce_preprocess_bytes

                data = mce_preprocess_bytes(data)

            xml_parser = parser(lenient=True, skipped_report=True)
            obj = xml_parser.from_bytes(data, clazz)
        except Exception as exc:  # noqa: BLE001 - the part name is the useful half
            raise Docx4JException(f"Problem with part {self.part_name}: {_reason(exc)}") from exc
        warnings: list[Any] = []
        items = getattr(xml_parser.skipped, "items", None)
        if items:
            # CR-030: docx4j's parse, repair and parse again, keyed on the report rather than on a failure
            from docx4j_py.openpackaging import breaks

            if breaks.triggers(items):
                repaired, found, texts = breaks.repair_bytes(data)
                if found or texts:
                    warnings.append(breaks.BrOutsideRun(str(self.part_name), found, texts))
                if repaired is not data:
                    xml_parser = parser(lenient=True, skipped_report=True)
                    try:
                        obj = xml_parser.from_bytes(repaired, clazz)
                    except Exception as exc:
                        raise Docx4JException(f"Problem with part {self.part_name}: {_reason(exc)}") from exc
        if xml_parser.config.on_bind is None:
            # No post-bind hook on this config, so the parse left the parent
            # pointers unwired and the second walk is still needed. With the
            # hook --- which `runtime.parser_config` fills in --- every object
            # was linked as it was built (CR-002 section 12.14).
            from docx4j_py.child import link_parents
            from docx4j_py.runtime import context

            link_parents(obj, context=context())
        return obj, xml_parser, resolved, warnings

    def _unmarshal(self, private: bool = False) -> None:
        clazz = self.model_class
        if clazz is None:
            raise Docx4JException(
                f"Part {self.part_name} ({self.content_type}) has no model class; "
                "it is a generic XML part, use .tree"
            )
        settings = self._parse_settings()
        if self._unreadable is not None and self._unreadable_settings == settings:
            # docx4j's unreadable: the read that failed, raised again without a second parse (CR-031 D8)
            raise self._unreadable.with_traceback(None)
        data = self._raw_bytes()
        self._source_ns_map = _root_ns_map(data)
        # an empty attribute value, either quote and any spacing round "=": two substring tests
        self._empty_value = b'""' in data or b"''" in data

        try:
            obj, xml_parser, resolved, warnings = self._parse(data, settings)
        except Docx4JException as error:
            self._unreadable = error
            self._unreadable_settings = settings
            raise
        self._unreadable = None
        self._unreadable_settings = None

        report = xml_parser.skipped
        self.skipped = list(getattr(report, "items", ()) or ())
        self.kept = list(getattr(report, "kept", ()) or ())
        self.resolved_alternate_content = resolved
        self.warnings = warnings
        for warning in warnings:
            # once a part, as docx4j's WARN (its CR-073 D3)
            logger.warning("%s", warning.message)
        if self.skipped:
            for item in self.skipped:
                logger.warning("%s: skipped %s", self.part_name, item)
            if settings.strict:
                raise Docx4JException(
                    f"{self.part_name}: {len(self.skipped)} item(s) were skipped "
                    f"and strict=True: {self.skipped[:5]}"
                )

        if private:
            self._private_contents = obj
            self._private_settings = settings
            return
        self._contents = obj
        self._bytes = None
        self._private_contents = None
        self._private_settings = None

    def _adopt(self) -> None:
        """The private tree made the live one: what :meth:`_unmarshal` would leave, without the parse.

        :attr:`skipped`, :attr:`kept`, :attr:`resolved_alternate_content`,
        :attr:`warnings` and ``_source_ns_map`` are already that parse's: the
        private parse set them from the same bytes under the same settings.
        """
        self._contents = self._private_contents
        self._bytes = None
        self._private_contents = None
        self._private_settings = None

    def _parse_settings(self) -> _ParseSettings:
        """What a parse now reads from the package."""
        return _ParseSettings.of(self.package)

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
        """The object as the bytes of a part: one render, one serialisation.

        The fork's CR-001 Phase A. The render goes straight into an lxml tree
        (:func:`~docx4j_py.runtime.tree_serializer`), the namespace
        reconciliation happens **in place** on that tree, and ``tostring`` is
        the only serialisation. Before Phase A this rendered to a string,
        re-parsed it, and copied the whole tree into a new root to change its
        namespace map --- 42% of a marshal of a 771 KB part.
        """
        from docx4j_py import mc_declarations
        from docx4j_py.openpackaging.normalise import normalise_render
        from docx4j_py.runtime import tree_serializer

        ns_map = _NS_MAP
        if self.default_namespace is not None:
            ns_map = {p: u for p, u in _NS_MAP.items() if u != self.default_namespace}
            ns_map[None] = self.default_namespace  # type: ignore[index]
        root = tree_serializer(bool_format=self.bool_format).render(obj, ns_map).getroot()

        wanted = self.root_name or _element_name(obj)
        if wanted is not None and root.tag != wanted:
            # The namespace is declared on the root already: every prefix of the
            # table is, before the cleanup below, so the rename keeps its prefix.
            root.tag = wanted

        # A normalising save writes the enumerations and colours in Word's normal
        # form (CR-018 section 4.7, Phase B), found by name in the render, so that
        # the tree keeps what was read; outside such a save this does nothing.
        normalise_render(root, str(self.part_name))

        # Drop what the tree does not use, then put back the prefixes
        # mc:Ignorable names --- they must be declared even though nothing in
        # the tree is in them, or Word rejects the file.
        etree.cleanup_namespaces(root)
        declared = root.nsmap
        missing: dict[str, str] = {}
        for prefix in sorted(_ignorable_prefixes(root)):
            if prefix in declared:
                continue
            # the source's binding first, then the table's (docx4j CR-052's D3)
            uri = self._source_ns_map.get(prefix) or PREFIXES.get(prefix)
            if uri is not None:
                missing[prefix] = uri
            else:
                logger.warning(
                    "%s: mc:Ignorable names the prefix %r, which neither the "
                    "prefix table nor the source declared; it is left undeclared",
                    self.part_name,
                    prefix,
                )
        # A prefix only an mc:Choice's Requires (or its own mc:Ignorable or
        # mc:MustUnderstand) names is no element's or attribute's, so the
        # cleanup took its declaration: on the root or on the Choice, as Word
        # places it (CR-011 Q3 and Q11, docx4j CR-052).
        declarations = mc_declarations.plan(root, obj, missing, self.part_name)
        missing.update(declarations.root)
        if missing:
            # ``top_nsmap`` declares them on the root, ``keep_ns_prefixes`` stops
            # the cleanup taking them straight back off again.
            etree.cleanup_namespaces(root, top_nsmap=missing, keep_ns_prefixes=list(missing))
        declarations.apply()

        data = etree.tostring(root, encoding="utf-8")
        return XML_DECLARATION + mc_declarations.finish(data, declarations)

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


def syntax_error_of(error: BaseException | None) -> etree.XMLSyntaxError | None:
    """lxml's ``XMLSyntaxError``, `error` or one of its causes (``__cause__``), or None: why XML was not well formed.

    The engine's refusal of such a part is a ``Docx4JException`` from the
    fork's ``ParserError``, from lxml's error (CR-031 Q8), or from lxml's error
    directly where an ``mc`` pass met it first. Its ``position`` is the line
    and the column.
    """
    seen: set[int] = set()
    while error is not None and id(error) not in seen:
        if isinstance(error, etree.XMLSyntaxError):
            return error
        seen.add(id(error))
        error = error.__cause__
    return None


def _reason(error: BaseException) -> str:
    """Why a parse failed: lxml's message where the XML is not well formed, else the error's own.

    lxml's ``msg``, without the ``(<string>, line 2)`` its ``str()`` appends
    for a source with no file name (CR-031 section 4.2).
    """
    syntax = syntax_error_of(error)
    if syntax is not None and syntax.msg:
        return str(syntax.msg)
    return str(error)


def _root_ns_map(data: bytes) -> dict[str, str]:
    """The namespace declarations on a part's root element.

    Parsed with ``iterparse`` and abandoned at the root's start tag, so this
    costs the root element and not the document. It never recovers (CR-031
    D4): lxml yields the root's start event before an error later in the
    bytes, so a part whose root start tag is well formed gives the same, and
    one whose root start tag is broken gives ``{}``.
    """
    import io

    try:
        for _event, element in etree.iterparse(
            io.BytesIO(data),
            events=("start",),
            resolve_entities=False,
            huge_tree=True,
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
