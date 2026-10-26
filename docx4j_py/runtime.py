"""The shared runtime: contexts, parsers, serialisers, ``warm_up``.

CR-001 section 8. Three things live here:

* the **shared ``XmlContext``** and the factories that make a parser and a
  serialiser configured the way this project needs them (lenient but never
  silent, numeric booleans, docx4j's prefixes);
* :func:`warm_up`, which parses an embedded representative part so that the
  lazily built class metadata exists before the first real document arrives;
* :func:`text_element` and :func:`needs_preserve`, the two helpers the
  generated ``el`` modules call for the text-carrying classes.

**Thread safety.** :func:`context` is shared on purpose: ``XmlContext`` only
ever *adds* to its caches, and the metadata it builds for a class is
deterministic, so two threads racing to build the same class waste a little
work and agree on the result (see ``scripts/threads.py``). A **parser** is not
shared: ``ParserConfig`` carries the per-parse ``skipped`` report and the
parser carries per-parse state, so :func:`parser` returns a new one each time
and a worker thread must keep its own. Serialisers hold no per-call state and
may be shared, but :func:`serializer` also returns a new one, which costs
nothing.
"""

from __future__ import annotations

import threading
import time
from collections.abc import Iterable
from enum import Enum
from importlib import import_module, resources
from typing import Any

from docx4j_xsdata.formats.dataclass.context import XmlContext
from docx4j_xsdata.formats.dataclass.parsers import XmlParser
from docx4j_xsdata.formats.dataclass.parsers.config import ParserConfig
from docx4j_xsdata.formats.dataclass.serializers import XmlSerializer
from docx4j_xsdata.formats.dataclass.serializers.config import SerializerConfig
from docx4j_xsdata.formats.dataclass.serializers.tree import TreeSerializer

__all__ = [
    "BRANCH_CHILD_NAMES",
    "EXACT_TYPES",
    "LEXICAL_TYPES",
    "branch_child_class",
    "container_binding",
    "context",
    "needs_preserve",
    "parser",
    "parser_config",
    "root_element_class",
    "serializer",
    "serializer_config",
    "text_element",
    "tree_serializer",
    "warm_up",
    "wildcard_names",
]

_context: XmlContext | None = None
_context_lock = threading.Lock()


def context() -> XmlContext:
    """Return the process-wide :class:`XmlContext`.

    One context per process is the point of CR-001 section 8: the metadata it
    builds is the expensive part and it is immutable once built.
    """
    global _context
    if _context is None:
        with _context_lock:
            if _context is None:
                _context = XmlContext()
    return _context


#: The value types whose values keep the spelling they were read with (the fork's
#: ``ParserConfig(lexical_types=)``, stage 12; CR-012 Phase C): VML's ``xsd:float``
#: and ``xsd:base64Binary`` attributes, which Word spells otherwise than the
#: converter would, and any other ``float`` or ``bytes`` value the same way.
LEXICAL_TYPES: frozenset[type] = frozenset({float, bytes})

#: The value types read exactly as written (the fork's ``ParserConfig(exact_types=)``,
#: stage 13, item 29; CR-018 section 4.1a): an ``xsd:boolean`` (Word's ST_OnOff in
#: docx4j's copy of the schema) converts only from exactly ``true``, ``false``,
#: ``1`` or ``0``, and an enumeration only from exactly a member's value. Anything
#: else is bound as the text it was read with and written back as read, so that
#: ``w:b w:val=" false"``, which Word reads as on, is not trimmed to ``False``.
#: Since the fork's stage 14 (item 32, docx4j-xsdata 26.2.4; CR-021 phase 0) an
#: ``int`` too converts only from ``[+-]?[0-9]+`` as written: XSD's whitespace
#: (``" 240 "``) is kept on purpose, underscores, other scripts' digits and Unicode
#: whitespace (``"1_000"``, ``"٢٤٠"``) with the warning, where Python's ``int()``
#: bound each as 240 (CR-021 section 2.2). Word refuses a document holding any of
#: them, so :func:`docx4j_py.model.values.integer` reads each as unset and a save
#: drops it (:mod:`docx4j_py.openpackaging.normalise`).
EXACT_TYPES: frozenset[type] = frozenset({bool, Enum, int})


def parser_config(
    *, lenient: bool = True, skipped_report: bool = True, **kwargs: Any
) -> ParserConfig:
    """A :class:`ParserConfig` for OOXML: lenient, but never silent.

    ``lenient`` keeps the promise of CR-001 section 1 --- a document Word wrote
    and no schema describes still loads --- and ``skipped_report`` is the fork's
    record of what lenience cost, so the caller can fail on it.

    ``on_bind`` is the fork's post-bind hook, and it is what wires the parent
    pointers: :func:`docx4j_py.child.link_children` runs on every object the
    parser builds, as it is built, instead of ``link_parents`` walking the
    finished tree a second time (CR-002 section 12.14). A caller who wants the
    plain parser may pass ``on_bind=None``; the tree then needs ``link_parents``
    before anything reads ``.parent``.

    ``on_namespaces`` is the fork's other hook (CR-011 Q3): each ``mc:Choice``
    the parser builds, typed or a wildcard, has its ``Requires`` recorded with
    the namespace URIs its prefixes resolve to in scope there
    (:func:`docx4j_py.child.record_bindings`), as are its ``mc:Ignorable`` and
    ``mc:MustUnderstand``,, so that a branch is selected by
    URI (WD REQ-303) and a save can declare a prefix only ``Requires`` names.

    ``wildcard_class`` is the fork's third (stage 11, CR-012 Phase A): a
    ``w:sdt`` straight in an ``mc:Choice`` or ``mc:Fallback`` is bound as the
    container the ``mc:AlternateContent`` stands in binds one
    (:func:`branch_child_class`), as docx4j's ``McBranchSdt`` binds it; and
    since CR-012 Phase A2 a ``w:tr`` or ``w:tc`` there too, as docx4j binds
    them by name. Since CR-012 Phase B the hook is asked about docx4j's 61
    WordprocessingML root elements as well (:func:`wildcard_names`): in a
    branch each is bound as its container binds it, and in any other lax
    wildcard, or a branch whose container does not take it, as its root
    class, which is what JAXB's ``@XmlRootElement`` gives docx4j.

    ``lexical_types`` is the fork's stage 12 (CR-012 Phase C): a ``float`` or
    ``bytes`` value the converter would write in another spelling than its
    source's keeps the source's (``XmlLexicalFloat``, ``XmlLexicalBytes``), so
    that VML's ``o:spt="202"`` is not saved ``202.0`` and Word's ``o:gfxdata``,
    wrapped at 76 characters, is not saved on one line. A value set in code is
    written in the converter's spelling, as before.

    ``exact_types`` is the fork's stages 13 and 14 (CR-018 section 4.1a, CR-021
    phase 0, :data:`EXACT_TYPES`): a ``bool``, an enumeration or an ``int`` value
    converts only from exactly a lexical form, with no XSD whitespace collapse.
    Anything else keeps its text: ``" false"`` and ``" 240 "`` (which XSD would
    have trimmed) silently, ``"off"``, ``"Yellow"``, ``"720.5"`` or ``"1_000"``
    with the usual ``ConverterWarning``. The skipped report lists both apart from the
    skipped items (``SkippedReport.kept``; ``part.kept``), and
    :mod:`docx4j_py.model.values` reads them as Word does. **Every parse in this
    package goes through this function** (``tests/test_parser_configs.py``), so
    that no parse reads ``" false"`` as ``False``.

    **One config per thread.** The report is mutable state on the config and is
    cleared at the start of every parse.
    """
    kwargs.setdefault("lexical_types", LEXICAL_TYPES)
    kwargs.setdefault("exact_types", EXACT_TYPES)
    if "on_bind" not in kwargs:
        from functools import partial

        from docx4j_py.child import link_children

        kwargs["on_bind"] = partial(link_children, context=context())
    if "on_namespaces" not in kwargs:
        from docx4j_py.child import MC_CHOICE, record_bindings

        kwargs["on_namespaces"] = record_bindings
        kwargs.setdefault("namespaces_for", frozenset({MC_CHOICE}))
    if "wildcard_class" not in kwargs:
        kwargs["wildcard_class"] = branch_child_class
    if kwargs["wildcard_class"] is not None and "wildcard_class_for" not in kwargs:
        kwargs["wildcard_class_for"] = wildcard_names()

    return ParserConfig(
        fail_on_unknown_properties=not lenient,
        fail_on_unknown_attributes=not lenient,
        skipped_report=skipped_report,
        **kwargs,
    )


def parser(*, lenient: bool = True, skipped_report: bool = True, **kwargs: Any) -> XmlParser:
    """A new :class:`XmlParser` on the shared context. Never share one."""
    return XmlParser(
        context=context(),
        config=parser_config(lenient=lenient, skipped_report=skipped_report, **kwargs),
    )


def serializer_config(*, pretty: bool = False, **kwargs: Any) -> SerializerConfig:
    """A :class:`SerializerConfig` for OOXML: numeric booleans, no declaration.

    ``bool_format="numeric"`` is CR-001 open question 5, decided: Word writes
    ``w:val="1"``, so writing ``true`` would make every saved document diff
    against its own source (3,034 attributes over ``samples/``).

    ``element_names`` (the fork's stage 11) names an object that no field
    names: the ``w:sdt`` kinds :func:`branch_child_class` binds in a branch
    are written ``w:sdt``, not under their types' names (``CT_SdtRun``), and
    a branch's ``Tr`` and ``Tc`` ``w:tr`` and ``w:tc`` (CR-012 Phase A2).
    Since CR-012 Phase B, every class the hook can give for one of the root
    elements is written under that element's name (:func:`_branch_element_names`):
    a ``Drawing`` as ``w:drawing``, an ``RT`` as ``w:t``, with no ``xsi:type``.
    A field's own name always wins, so nothing written through a field
    changes.
    """
    kwargs.setdefault("bool_format", "numeric")
    kwargs.setdefault("xml_declaration", False)
    if "element_names" not in kwargs:
        kwargs["element_names"] = _branch_element_names()
    if pretty:
        kwargs.setdefault("indent", "  ")
    return SerializerConfig(**kwargs)


def serializer(*, pretty: bool = False, **kwargs: Any) -> XmlSerializer:
    """A new :class:`XmlSerializer` on the shared context."""
    return XmlSerializer(context=context(), config=serializer_config(pretty=pretty, **kwargs))


def tree_serializer(*, pretty: bool = False, **kwargs: Any) -> TreeSerializer:
    """A new :class:`TreeSerializer` on the shared context, configured as :func:`serializer`.

    The twin of :func:`serializer` that renders an **lxml tree** instead of a
    string: the same :func:`serializer_config`, and the same events --- the fork
    builds both writers on ``EventGenerator``, so ``bool_format`` and the rest
    of the config mean exactly what they mean for a string render, and only the
    last step differs (``LxmlTreeBuilder`` keeps the tree the sax handler built,
    where ``LxmlEventWriter`` serialises it). ``xml_declaration`` has no meaning
    for a tree and is ignored.

    This is what :meth:`XmlPart._marshal` uses: a part has namespace work to do
    on the tree afterwards, and going through a string would mean serialising
    twice and parsing once (the fork's CR-001 Phase A).
    """
    return TreeSerializer(context=context(), config=serializer_config(pretty=pretty, **kwargs))


# ---------------------------------------------------------------------------
# a markup-compatibility branch's children (CR-012 Phase A)
# ---------------------------------------------------------------------------

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

#: The container-bound names :func:`branch_child_class` is asked about.
#:
#: * ``w:sdt``, a content control (CR-012 Phase A). Its four classes share one
#:   element name, and only the block-level one (``SdtBlock``) is global, so a
#:   lookup by name alone gives the block class at every level.
#: * ``w:tr`` and ``w:tc``, a row among a table's rows and a cell among a row's
#:   cells (CR-012 Phase A2, docx4j CR-051 phase 2). ``w:tc`` is declared only
#:   inside ``CT_Row`` and its controls, so a lookup by name leaves a
#:   cell-level branch's cell a wildcard tree that no walk of cells reads;
#:   ``w:tr`` is global as ``Tr``, which a row-level branch's container binds
#:   it to as well, and it is asked so that both are bound by the one rule.
#:
#: Since CR-012 Phase B the hook is also asked about the root elements
#: (``docx4j_py.el_index.ROOT_ELEMENTS``); :func:`wildcard_names` is the whole
#: set.
BRANCH_CHILD_NAMES = frozenset({f"{{{_W_NS}}}sdt", f"{{{_W_NS}}}tr", f"{{{_W_NS}}}tc"})

#: The classes of ``mc:AlternateContent`` and its branches: what stands
#: between a branch's child and the container the element stands in.
_MC_WRAPPERS = frozenset({"AlternateContent", "AlternateContentChoice", "AlternateContentFallback"})

_branch_names: dict[type, str] | None = None
_wildcard_names: frozenset[str] | None = None
_root_classes: dict[str, type] | None = None


def wildcard_names() -> frozenset[str]:
    """Every name the parser asks :func:`branch_child_class` about: the fork's ``wildcard_class_for``.

    :data:`BRANCH_CHILD_NAMES` and, since CR-012 Phase B, the 61 root elements
    of ``docx4j_py.el_index.ROOT_ELEMENTS``: docx4j's WordprocessingML
    ``@XmlRootElement`` names that this model does not declare globally,
    derived from docx4j's own classes (``codegen/names/root_elements.json``).
    A name not in the set is bound by the lookup by name alone: a branch's
    ``w:ins``, ``w:instrText`` or ``w:delText`` stays a wildcard tree, as docx4j
    keeps it a DOM element. Read from the generated index, which holds strings
    only, so asking imports no model class.
    """
    global _wildcard_names
    if _wildcard_names is None:
        from docx4j_py.el_index import ROOT_ELEMENTS

        _wildcard_names = BRANCH_CHILD_NAMES | frozenset(ROOT_ELEMENTS)
    return _wildcard_names


def _resolve(path: str) -> type:
    module, _, name = path.rpartition(".")
    return getattr(import_module(module), name)


def root_element_class(qname: str) -> type | None:
    """The root class of one of docx4j's root elements, or None: step (b) of :func:`branch_child_class`.

    JAXB binds a lax wildcard's child by its ``@XmlRootElement`` class, so in
    docx4j a ``w:drawing`` anywhere a lax wildcard receives one is a
    ``Drawing``. ``docx4j_py.el_index.ROOT_ELEMENTS`` names each root's class
    as a string, the class the ``el`` table builds for the name; it is
    resolved the first time it is asked for, never at import, since the
    engine imports no model class at module level.
    """
    global _root_classes
    table = _root_classes
    if table is None:
        from docx4j_py.el_index import ROOT_ELEMENTS

        table = _root_classes = {qname: _resolve(path) for qname, path in ROOT_ELEMENTS.items()}
    return table.get(qname)


def branch_child_class(qname: str, ancestors: tuple[type, ...]) -> type | None:
    """The class to bind a lax wildcard's child to, by its context and then by its name.

    The fork parser's ``wildcard_class`` hook (``ParserConfig``, stage 11),
    asked about :func:`wildcard_names`. `ancestors` is the classes enclosing
    the element, innermost first, the class whose wildcard receives it first.
    Three steps (CR-012 section 4.2 item 1):

    * **(a) in a branch, the container's binding** (:func:`container_binding`):
      docx4j's ``McBranchSdt.kindFor`` (its CR-051 section 11.3), the rule
      Phases A and A2 established for ``w:sdt``, ``w:tr`` and ``w:tc`` and
      Phase B extends to the root elements --- a branch's child is read as the
      container's own child is, so a ``w:hyperlink`` in a simple field's
      branch is the simple field's ``CtSimpleFieldHyperlink``, as it is
      outside a branch;
    * **(b) elsewhere, the root class** (:func:`root_element_class`): in any
      other lax wildcard, and in a branch whose container does not take the
      name, as JAXB binds an ``@XmlRootElement`` name;
    * **(c) otherwise None**, and the lookup by name applies: a ``w:sdt`` or
      ``w:tr`` outside a branch is the global ``SdtBlock`` or ``Tr``.
    """
    found = container_binding(qname, ancestors)
    if found is not None:
        return found
    return root_element_class(qname)


def container_binding(qname: str, ancestors: tuple[type, ...]) -> type | None:
    """Step (a) of :func:`branch_child_class`: the class a branch's container binds `qname` to.

    Past the markup-compatibility wrappers, the first other class in
    `ancestors` is the container the ``mc:AlternateContent`` stands in, and
    the class its own content binds `qname` to is the answer: ``SdtBlock`` in
    a body, a cell or a block-level control, ``SdtRun`` in a paragraph,
    ``CTSdtRow`` among a table's rows, ``CTSdtCell`` among a row's cells;
    ``Tr`` among a table's rows and ``Tc`` among a row's cells, in a row- or
    cell-level control's content and custom XML too (CR-012 Phase A2); and
    since Phase B ``Drawing`` or ``RT`` in a run, ``Tbl`` in a body,
    ``CtRTab`` in a math run. None where that container takes no such
    element, or the wildcard is not a branch's. The rule is the schema's own
    binding, read from the metadata, not a table.
    """
    if not ancestors or not _is_mc_wrapper(ancestors[0]):
        return None
    for clazz in ancestors:
        if _is_mc_wrapper(clazz):
            continue
        for var in context().build(clazz).get_children(qname):
            if var.clazz is not None:
                return var.clazz
        return None
    return None


def _is_mc_wrapper(clazz: type) -> bool:
    return clazz.__module__ == "docx4j_py.mce" and clazz.__name__ in _MC_WRAPPERS


def element_names_table(pairs: Iterable[tuple[type, str]]) -> dict[type, str]:
    """The serialiser's ``element_names`` from ``(class, qualified name)`` pairs: a function, or an error.

    Each class is written under one name wherever no field names it, so a
    class claimed by two names is refused rather than one of them chosen.

    Raises:
        ValueError: a class paired with two names.
    """
    names: dict[type, str] = {}
    for cls, qname in pairs:
        current = names.setdefault(cls, qname)
        if current != qname:
            raise ValueError(
                f"{cls.__module__}.{cls.__qualname__} is claimed by {current} and {qname}: "
                "the serialiser's element names must be a function, class to name"
            )
    return names


def _branch_element_names() -> dict[type, str]:
    """The serialiser's ``element_names``: each class the branch hook may bind, under its element's name.

    * the four ``w:sdt`` classes as ``w:sdt``, and ``Tr`` and ``Tc`` as
      ``w:tr`` and ``w:tc`` (``Tc``'s own name is its type's, ``CT_Tc``);
    * since CR-012 Phase B, each root class under its name
      (``docx4j_py.el_index.ROOT_ELEMENTS``), and each other class step (a)
      can give for a root under that name (``ROOT_ELEMENT_BINDINGS``, the
      generator's list of what every container that can hold an
      ``mc:AlternateContent`` binds it to).

    Built by :func:`element_names_table`, so a class claimed by two names
    raises here, the first time a serialiser is configured.
    """
    global _branch_names
    if _branch_names is None:
        from docx4j_py.el_index import ROOT_ELEMENT_BINDINGS, ROOT_ELEMENTS
        from docx4j_py.wml import CTSdtCell, CTSdtRow, SdtBlock, SdtRun, Tc, Tr

        pairs: list[tuple[type, str]] = [
            (cls, f"{{{_W_NS}}}sdt") for cls in (SdtBlock, SdtRun, CTSdtRow, CTSdtCell)
        ]
        pairs += [(Tr, f"{{{_W_NS}}}tr"), (Tc, f"{{{_W_NS}}}tc")]
        pairs += [(_resolve(path), qname) for qname, path in ROOT_ELEMENTS.items()]
        pairs += [
            (_resolve(path), qname)
            for qname, paths in ROOT_ELEMENT_BINDINGS.items()
            for path in paths
        ]
        _branch_names = element_names_table(pairs)
    return _branch_names


# ---------------------------------------------------------------------------
# text elements
# ---------------------------------------------------------------------------


def needs_preserve(text: str) -> bool:
    """True when ``xml:space="preserve"`` has to be written for `text`.

    docx4j's rule, and the one ``builders/wml.mts`` `t()` uses: leading or
    trailing whitespace, or two whitespace characters in a row, are the cases an
    XML processor is free to collapse.
    """
    if not text:
        return False
    if text[0].isspace() or text[-1].isspace():
        return True
    previous = False
    for ch in text:
        space = ch.isspace()
        if space and previous:
            return True
        previous = space
    return False


def text_element(
    cls: type,
    text: str | None,
    kwargs: dict[str, Any],
    text_field: str,
    space_field: str | None,
) -> Any:
    """Build a text-carrying element, setting ``xml:space`` when it is needed.

    CR-001 decided question 6: the ``el`` functions of the text-carrying classes
    --- and only those --- take the text as one positional argument.
    """
    if text is not None:
        if text_field in kwargs:
            raise TypeError(
                f"{cls.__name__}() got the text both positionally and as {text_field}="
            )
        kwargs[text_field] = text
        if space_field is not None and space_field not in kwargs and needs_preserve(text):
            kwargs[space_field] = "preserve"
    return cls(**kwargs)


# ---------------------------------------------------------------------------
# warm up
# ---------------------------------------------------------------------------

#: The parts ``warm_up`` parses, in order.
WARM_UP_PARTS: tuple[tuple[str, str], ...] = (
    ("warmup_document.xml", "docx4j_py.wml.Document"),
    ("warmup_styles.xml", "docx4j_py.wml.Styles"),
)

_warmed = False


def warm_up(*, force: bool = False) -> float:
    """Build the class metadata by parsing an embedded representative part.

    ``XmlContext`` builds a class's metadata on first use, so the first real
    document after start-up pays for every class it touches. This parses a small
    ``document.xml`` and ``styles.xml`` that between them hold paragraphs, runs,
    direct and style formatting, a table, a drawing, ``w14`` attributes and an
    ``mc:AlternateContent``, which is most of what a Word document is made of.

    Returns:
        The seconds spent. Zero if the process is already warm and `force` is
        not set.
    """
    global _warmed
    if _warmed and not force:
        return 0.0

    started = time.perf_counter()
    files = resources.files("docx4j_py.resources")
    for name, target in WARM_UP_PARTS:
        module_name, _, class_name = target.rpartition(".")
        from importlib import import_module

        clazz = getattr(import_module(module_name), class_name)
        data = (files / name).read_bytes()
        parser().from_bytes(data, clazz)
    _warmed = True
    return time.perf_counter() - started
