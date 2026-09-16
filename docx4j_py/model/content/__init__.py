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
    "ChangeReport",
    "ContentError",
    "Description",
    "Docx4JError",
    "Font",
    "InsertLocation",
    "InvalidTargetError",
    "Outline",
    "OutlineEntry",
    "OutlineSection",
    "OutlineStats",
    "PageSetup",
    "Paragraph",
    "PartInfo",
    "Range",
    "SearchHit",
    "Segment",
    "SpanError",
    "StyleError",
    "StyleInfo",
    "TextExcerpt",
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
    # CR-003 Phase D, the agent surface
    "ChangeReport": "docx4j_py.model.content.reports",
    "Outline": "docx4j_py.model.content.reports",
    "OutlineEntry": "docx4j_py.model.content.reports",
    "OutlineSection": "docx4j_py.model.content.reports",
    "OutlineStats": "docx4j_py.model.content.reports",
    "SearchHit": "docx4j_py.model.content.reports",
    "TextExcerpt": "docx4j_py.model.content.reports",
    "Description": "docx4j_py.model.content.describe",
    "PageSetup": "docx4j_py.model.content.describe",
    "PartInfo": "docx4j_py.model.content.describe",
    "StyleInfo": "docx4j_py.model.content.describe",
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


# ---------------------------------------------------------------------------
# the package's half of the agent surface (CR-003 section 3.4, Phase D)
# ---------------------------------------------------------------------------


def _package_outline(self: object, **options: object) -> object:
    """The whole document's outline: the body, then the headers and footers.

    Args:
        **options: ``depth``, ``max_chars``, ``headings_only`` and ``limit``,
            as :meth:`docx4j_py.model.content.Body.outline`.
    """
    from docx4j_py.model.content.reports import package_outline

    return package_outline(self, **options)  # type: ignore[arg-type]


def _package_describe(self: object) -> object:
    """What an agent can use: the styles, the page, the parts, the authors.

    Reads every part it needs as **bytes**, with lxml, so nothing is
    unmarshalled and an untouched part stays byte for byte
    (:mod:`docx4j_py.model.content.describe`).
    """
    from docx4j_py.model.content.describe import describe_package

    return describe_package(self)


def _package_element_at(self: object, address: str) -> object:
    """The block at an address, in any of the document's bodies."""
    from docx4j_py.model.content.addresses import element_at

    return element_at(self.body, address)  # type: ignore[attr-defined]


def _package_paragraph_at(self: object, address: str | None = None, **options: object) -> object:
    """The paragraph at an address, containing some text, or with a paraId."""
    from docx4j_py.model.content.addresses import paragraph_at

    return paragraph_at(self.body, address, **options)  # type: ignore[attr-defined,arg-type]


def _package_find(self: object, text: str, **options: object) -> object:
    """Every match in the main document part, as hits with addresses."""
    from docx4j_py.model.content.reports import find_in

    return find_in(self.body, text, **options)  # type: ignore[attr-defined,arg-type]


def _package_bodies(self: object) -> list:
    """Every body of the document: the main part, the headers, the footers, the notes."""
    from docx4j_py.model.content.addresses import package_bodies

    return list(package_bodies(self))


def _package_dry_run(self: object) -> object:
    """``with pkg.dry_run() as trial:`` --- edits on a copy, then thrown away."""
    from docx4j_py.model.content.trial import dry_run

    return dry_run(self)


#: What :func:`register` puts on ``WordprocessingMLPackage``, beyond ``body``.
_PACKAGE_MEMBERS: dict[str, object] = {
    "outline": _package_outline,
    "describe": _package_describe,
    "element_at": _package_element_at,
    "paragraph_at": _package_paragraph_at,
    "find": _package_find,
    "bodies": _package_bodies,
    "dry_run": _package_dry_run,
}


def register() -> None:
    """Give the parts and the package their content API. Called on import.

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
    for name, member in _PACKAGE_MEMBERS.items():
        if getattr(WordprocessingMLPackage, name, None) is None:
            setattr(WordprocessingMLPackage, name, member)


register()
