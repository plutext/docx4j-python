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

__all__ = [
    "context",
    "needs_preserve",
    "parser",
    "parser_config",
    "serializer",
    "serializer_config",
    "text_element",
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

    **One config per thread.** The report is mutable state on the config and is
    cleared at the start of every parse.
    """
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
    """
    kwargs.setdefault("bool_format", "numeric")
    kwargs.setdefault("xml_declaration", False)
    if pretty:
        kwargs.setdefault("indent", "  ")
    return SerializerConfig(**kwargs)


def serializer(*, pretty: bool = False, **kwargs: Any) -> XmlSerializer:
    """A new :class:`XmlSerializer` on the shared context."""
    return XmlSerializer(context=context(), config=serializer_config(pretty=pretty, **kwargs))


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
