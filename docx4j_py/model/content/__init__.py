"""The content API: ``Body``, ``Paragraph``, ``Range`` and ``Font``.

CR-003 Phase B. Office JS's object model and verbs, in Python's conventions
(section 3.1), over the docx4j tree that CR-001 and CR-002 already hold:

    >>> from docx4j_py import create_package
    >>> pkg = create_package()
    >>> title = pkg.body.insert_paragraph("Report", style="Heading 1")
    >>> title.alignment = "Centered"
    >>> pkg.save("report.docx")                        # doctest: +SKIP

A view is a plain object over the tree, created on access and cheap to throw
away: ``paragraph.element`` is the ``P``, ``body.content`` is the live
``ChildList``, and nothing is cached, so an edit made through the tree is
visible through the views and the other way round.

Importing this module **registers** ``body`` on the parts that have one ---
``MainDocumentPart``, ``HeaderPart``, ``FooterPart``, ``FootnotesPart``,
``EndnotesPart``, ``CommentsPart`` --- and on ``WordprocessingMLPackage``. The
registration goes this way round on purpose (CR-003 section 5): the content
module imports the parts layer, and the parts layer never imports the content
module.

Everything but the error hierarchy is imported lazily, through a module
``__getattr__`` (PEP 562), so that :mod:`docx4j_py.wml.builders` can take
:class:`~docx4j_py.model.content.errors.BuilderError` from here without the
views importing the builders back.
"""

from __future__ import annotations

from docx4j_py.model.content.errors import (
    AddressError,
    BuilderError,
    ContentError,
    Docx4JError,
    InvalidTargetError,
    SpanError,
    StyleError,
)

__all__ = [
    "BUILT_IN_STYLES",
    "AddressError",
    "Alignment",
    "Block",
    "Body",
    "BreakType",
    "BuilderError",
    "ContentError",
    "Docx4JError",
    "Font",
    "InsertLocation",
    "InvalidTargetError",
    "Paragraph",
    "Range",
    "Segment",
    "SpanError",
    "StyleError",
    "UnderlineType",
    "body_of",
    "built_in_of",
    "grapheme_clusters",
    "segments_of",
]

#: Public name -> the module it lives in. Imported on first use.
_LAZY: dict[str, str] = {
    "Block": "docx4j_py.model.content.body",
    "Body": "docx4j_py.model.content.body",
    "body_of": "docx4j_py.model.content.body",
    "Paragraph": "docx4j_py.model.content.paragraph",
    "Range": "docx4j_py.model.content.range",
    "Font": "docx4j_py.model.content.font",
    "Segment": "docx4j_py.model.content.text_model",
    "segments_of": "docx4j_py.model.content.text_model",
    "grapheme_clusters": "docx4j_py.model.content.text_model",
    "BUILT_IN_STYLES": "docx4j_py.model.content.styles",
    "built_in_of": "docx4j_py.model.content.styles",
    "Alignment": "docx4j_py.model.content.enums",
    "BreakType": "docx4j_py.model.content.enums",
    "InsertLocation": "docx4j_py.model.content.enums",
    "UnderlineType": "docx4j_py.model.content.enums",
}


def __getattr__(name: str) -> object:
    """Import a view, an enum or a helper on first use."""
    target = _LAZY.get(name)
    if target is None:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    from importlib import import_module

    value = getattr(import_module(target), name)
    globals()[name] = value
    return value


def _body_view(target: object) -> object:
    """``body_of``, imported on first use so that registering costs nothing."""
    from docx4j_py.model.content.body import body_of

    return body_of(target)


_PART_BODY = property(
    _body_view,
    doc=(
        "The content-API :class:`~docx4j_py.model.content.Body` over this "
        "part's block-level content (CR-003 Phase B). The typed ``w:body`` "
        "itself is ``part.contents.body``, or ``part.body_element``."
    ),
)

_PACKAGE_BODY = property(
    _body_view,
    doc=(
        "The main document part's :class:`~docx4j_py.model.content.Body` "
        "(CR-003: ``pkg.body.insert_paragraph('Hello World')``)."
    ),
)


def register() -> None:
    """Give the parts and the package their ``body``. Called on import.

    Idempotent, and the only thing this package does at import time beyond
    defining the error hierarchy. The parts layer is imported here; it never
    imports this one (CR-003 section 5).
    """
    from docx4j_py.openpackaging.packages.wordprocessingml_package import (
        WordprocessingMLPackage,
    )
    from docx4j_py.openpackaging.parts.xml_part import XmlPart

    if not isinstance(getattr(XmlPart, "body", None), property):
        XmlPart.body = _PART_BODY  # type: ignore[attr-defined]
    if not isinstance(getattr(WordprocessingMLPackage, "body", None), property):
        WordprocessingMLPackage.body = _PACKAGE_BODY  # type: ignore[attr-defined]


register()
