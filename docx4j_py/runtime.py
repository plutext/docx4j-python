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
from importlib import resources
from typing import Any

from docx4j_xsdata.formats.dataclass.context import XmlContext
from docx4j_xsdata.formats.dataclass.parsers import XmlParser
from docx4j_xsdata.formats.dataclass.parsers.config import ParserConfig
from docx4j_xsdata.formats.dataclass.serializers import XmlSerializer
from docx4j_xsdata.formats.dataclass.serializers.config import SerializerConfig
from docx4j_xsdata.formats.dataclass.serializers.tree import TreeSerializer

__all__ = [
    "BRANCH_CHILD_NAMES",
    "branch_child_class",
    "context",
    "needs_preserve",
    "parser",
    "parser_config",
    "serializer",
    "serializer_config",
    "text_element",
    "tree_serializer",
    "warm_up",
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
    them by name.

    **One config per thread.** The report is mutable state on the config and is
    cleared at the start of every parse.
    """
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
    if kwargs["wildcard_class"] is not None:
        kwargs.setdefault("wildcard_class_for", BRANCH_CHILD_NAMES)

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

#: The names :func:`branch_child_class` is asked about.
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
BRANCH_CHILD_NAMES = frozenset({f"{{{_W_NS}}}sdt", f"{{{_W_NS}}}tr", f"{{{_W_NS}}}tc"})

#: The classes of ``mc:AlternateContent`` and its branches: what stands
#: between a branch's child and the container the element stands in.
_MC_WRAPPERS = frozenset({"AlternateContent", "AlternateContentChoice", "AlternateContentFallback"})

_branch_names: dict[type, str] | None = None


def branch_child_class(qname: str, ancestors: tuple[type, ...]) -> type | None:
    """The class to bind a branch's child to: the one its container binds that name to.

    The fork parser's ``wildcard_class`` hook (``ParserConfig``, stage 11),
    docx4j's ``McBranchSdt.kindFor`` (its CR-051 section 11.3). `ancestors` is
    the classes enclosing the element, innermost first, the class whose
    wildcard receives it (an ``mc:Choice`` or ``mc:Fallback``) first. Past the
    markup-compatibility wrappers, the first other class is the container the
    ``mc:AlternateContent`` stands in, and the class its own content binds
    `qname` to is the answer: ``SdtBlock`` in a body, a cell or a block-level
    control, ``SdtRun`` in a paragraph, ``CTSdtRow`` among a table's rows,
    ``CTSdtCell`` among a row's cells; ``Tr`` among a table's rows and ``Tc``
    among a row's cells, in a row- or cell-level control's content and custom
    XML too (CR-012 Phase A2). None where that container takes no such
    element, or the wildcard is not a branch's; the lookup by name then
    applies.
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


def _branch_element_names() -> dict[type, str]:
    """The serialiser's ``element_names``: each class the branch hook may bind, under its element's name.

    The four ``w:sdt`` classes as ``w:sdt``, and ``Tr`` and ``Tc`` as ``w:tr``
    and ``w:tc`` (``Tc``'s own name is its type's, ``CT_Tc``).
    """
    global _branch_names
    if _branch_names is None:
        from docx4j_py.wml import CTSdtCell, CTSdtRow, SdtBlock, SdtRun, Tc, Tr

        names = {cls: f"{{{_W_NS}}}sdt" for cls in (SdtBlock, SdtRun, CTSdtRow, CTSdtCell)}
        names[Tr] = f"{{{_W_NS}}}tr"
        names[Tc] = f"{{{_W_NS}}}tc"
        _branch_names = names
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
