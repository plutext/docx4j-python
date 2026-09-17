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
    BindingError,
    BuilderError,
    ContentError,
    Docx4JError,
    InvalidTargetError,
    SpanError,
    StyleError,
    TrackedChangeError,
)

__all__ = [
    "BUILT_IN_STYLES",
    "AddressError",
    "Alignment",
    "Author",
    "BindingError",
    "BindingInfo",
    "BindingResult",
    "Block",
    "Body",
    "BreakType",
    "BuilderError",
    "ChangeReport",
    "ChangeTracker",
    "ChangeTracking",
    "CheckboxContentControl",
    "ComboBoxContentControl",
    "Comment",
    "ContentControl",
    "ContentControlListItem",
    "ContentError",
    "CustomXmlNode",
    "CustomXmlPart",
    "CustomXmlPartCollection",
    "CustomXmlPrefixMappingCollection",
    "DatePickerContentControl",
    "Description",
    "Docx4JError",
    "DropDownListContentControl",
    "FillResult",
    "Font",
    "GroupContentControl",
    "InlinePicture",
    "InsertLocation",
    "InvalidTargetError",
    "List",
    "ListContentControl",
    "ListItem",
    "Outline",
    "OutlineEntry",
    "OutlineSection",
    "OutlineStats",
    "PageSetup",
    "Paragraph",
    "PartInfo",
    "PictureContentControl",
    "Range",
    "RepeatInfo",
    "RepeatingSectionContentControl",
    "SearchHit",
    "Segment",
    "Skeleton",
    "SpanError",
    "StyleError",
    "StyleInfo",
    "Table",
    "TableCell",
    "TableRow",
    "TextExcerpt",
    "TrackedChange",
    "TrackedChangeError",
    "TrackedChangeType",
    "UnderlineType",
    "XmlMapping",
    "body_of",
    "built_in_of",
    "cells_of",
    "grapheme_clusters",
    "rows_of",
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
    # CR-003 Phase C, the tables, the pictures and the controls
    "Table": "docx4j_py.model.content.table",
    "TableRow": "docx4j_py.model.content.table",
    "TableCell": "docx4j_py.model.content.table",
    "rows_of": "docx4j_py.model.content.text_model",
    "cells_of": "docx4j_py.model.content.text_model",
    "InlinePicture": "docx4j_py.model.content.picture",
    "ContentControl": "docx4j_py.model.content.controls",
    # CR-003 Phase H, the lists
    "List": "docx4j_py.model.content.lists",
    "ListItem": "docx4j_py.model.content.lists",
    # CR-003 Phase G, the comments
    "Author": "docx4j_py.model.content.comments",
    "Comment": "docx4j_py.model.content.comments",
    # CR-003 Phase F, change tracking
    "ChangeTracker": "docx4j_py.model.content.tracking",
    "ChangeTracking": "docx4j_py.model.content.enums",
    "TrackedChange": "docx4j_py.model.content.tracked_change",
    "TrackedChangeType": "docx4j_py.model.content.enums",
    # CR-003 Phase E, custom XML and the typed kinds
    "BindingError": "docx4j_py.model.content.errors",
    "CustomXmlNode": "docx4j_py.model.customxml",
    "CustomXmlPart": "docx4j_py.model.customxml",
    "CustomXmlPartCollection": "docx4j_py.model.customxml",
    "CustomXmlPrefixMappingCollection": "docx4j_py.model.customxml",
    "XmlMapping": "docx4j_py.model.customxml",
    "CheckboxContentControl": "docx4j_py.model.customxml",
    "ComboBoxContentControl": "docx4j_py.model.customxml",
    "ContentControlListItem": "docx4j_py.model.customxml",
    "DatePickerContentControl": "docx4j_py.model.customxml",
    "DropDownListContentControl": "docx4j_py.model.customxml",
    "GroupContentControl": "docx4j_py.model.customxml",
    "ListContentControl": "docx4j_py.model.customxml",
    "PictureContentControl": "docx4j_py.model.customxml",
    "RepeatingSectionContentControl": "docx4j_py.model.customxml",
    "Skeleton": "docx4j_py.model.customxml",
    "BindingInfo": "docx4j_py.model.customxml",
    "BindingResult": "docx4j_py.model.customxml",
    "FillResult": "docx4j_py.model.customxml",
    "RepeatInfo": "docx4j_py.model.customxml",
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


def _package_to_markdown(self: object, **options: object) -> str:
    """The main document part as markdown (CR-003 section 3.5, Phase K).

    docx4j-mcp's ``docx_to_markdown``, in one call: the body's blocks, with the
    footnotes it refers to as GFM footnotes at the end.

    Args:
        **options: ``addresses``, ``view`` (``"accepted"`` or ``"markup"``) and
            ``max_chars``, as :meth:`docx4j_py.model.content.Body.to_markdown`.
    """
    return self.body.to_markdown(**options)  # type: ignore[attr-defined,no-any-return]


def _package_markdown_budget(self: object, max_chars: object = None, **options: object) -> object:
    """:meth:`to_markdown` with the flag: ``TextExcerpt(text, chars, truncated)``."""
    return self.body.markdown_budget(max_chars, **options)  # type: ignore[attr-defined]


def _package_insert_markdown(self: object, markdown: str, **options: object) -> object:
    """``markdown_to_docx``: insert markdown into the main document part's body."""
    return self.body.insert_markdown(markdown, **options)  # type: ignore[attr-defined]


def _package_author(self: object) -> object:
    """Who this package's comments and tracked changes are by (CR-003 section 3.9).

    ``pkg.author = Author("Claude", initials="C")``. There is no signed-in user
    here, so the package carries the identity; a package that has not been given
    one writes :data:`~docx4j_py.model.content.comments.DEFAULT_AUTHOR`,
    ``Author("docx4j-python")``. **Phase F's tracked changes read the same
    setting.**
    """
    from docx4j_py.model.content.comments import DEFAULT_AUTHOR

    found = self._author  # type: ignore[attr-defined]
    return found if found is not None else DEFAULT_AUTHOR


def _set_package_author(self: object, value: object) -> None:
    """Set the identity; a bare string is taken as the author's name."""
    from docx4j_py.model.content.comments import Author

    if isinstance(value, str):
        value = Author(value)
    if value is not None and not isinstance(value, Author):
        from docx4j_py.model.content.errors import ContentError

        raise ContentError(
            f"pkg.author takes an Author or a name, not {type(value).__name__}",
            code="author.invalid",
            hint="pkg.author = Author('Claude', initials='C')",
        )
    self._author = value  # type: ignore[attr-defined]


def _package_change_tracking_mode(self: object) -> object:
    """Whether this package's edits are tracked (CR-003 section 3.8).

    ``"Off"``, ``"TrackAll"`` or ``"TrackMineOnly"``, over ``w:trackRevisions``
    in ``/word/settings.xml``. **Reading does not unmarshal the settings part**
    --- the flag is read with lxml from the bytes the part would be saved as,
    as ``describe().tracking_on`` reads it --- so a document whose mode is only
    read keeps that part byte for byte. Setting it does unmarshal the part, and
    creates one with its relationship and content type when the document has
    none; ``"TrackMineOnly"`` is stored as ``"TrackAll"``, because
    ``w:trackRevisions`` is a flag and a file cannot tell the two apart.
    """
    from docx4j_py.model.content.tracking import mode_of

    return mode_of(self)


def _set_package_change_tracking_mode(self: object, value: object) -> None:
    """Turn tracking on or off; the part touched goes into the ``ChangeReport``.

    One report per call, as every mutating call has (CR-003 decided question 4),
    with ``/word/settings.xml`` in ``parts_touched`` when the setter really
    wrote it and ``text_after`` the mode now in force.
    """
    from docx4j_py.model.content.reports import recording
    from docx4j_py.model.content.tracking import set_mode

    try:
        body = self.body  # type: ignore[attr-defined]
    except Exception:  # noqa: BLE001 - a package with no body still takes a mode
        body = None
    if body is None:
        set_mode(self, None if value is None else str(value))
        return
    with recording(body, "change_tracking_mode") as change:
        touched = set_mode(self, None if value is None else str(value))
        change.text(after=str(value))
        parts = getattr(change, "parts", None)
        if parts is not None:
            for name in touched:
                if name not in parts:
                    parts.append(name)


def _package_compatibility_mode(self: object) -> object:
    """Which Word this document says it is written for (CR-003 section 3.4).

    An ``int``: 11 is Word 2003, 12 Word 2007, 14 Word 2010, 15 Word 2013 ---
    and Word 2016, 2019 and 365, which all write 15. **12 when the document
    declares nothing**, which is what Word assumes, and why Word shows
    *Compatibility Mode* in the title bar for such a document; a document this
    library creates declares 15.

    **Reading does not unmarshal the settings part**: ``w:compatSetting`` is
    read with lxml from the bytes the part would be saved as, as
    ``change_tracking_mode`` is, and cached on the package. Setting it unmarshals
    the part, writes or replaces the setting, and records a ``ChangeReport``
    naming ``/word/settings.xml``. Setting it does **not** add or remove any
    other markup: it says which Word to target, and a feature a lower mode
    cannot carry is reported in that call's ``ChangeReport.warnings``
    (:func:`~docx4j_py.model.content.compatibility.warn_below`).
    """
    from docx4j_py.model.content.compatibility import mode_of

    return mode_of(self)


def _set_package_compatibility_mode(self: object, value: object) -> None:
    """Declare a Word version: 11, 12, 14 or 15. The part touched is reported."""
    from docx4j_py.model.content.compatibility import set_mode
    from docx4j_py.model.content.reports import recording

    try:
        body = self.body  # type: ignore[attr-defined]
    except Exception:  # noqa: BLE001 - a package with no body still takes a mode
        body = None
    if body is None:
        set_mode(self, value)
        return
    with recording(body, "compatibility_mode") as change:
        touched = set_mode(self, value)
        change.text(after=str(int(value)))  # type: ignore[call-overload]
        parts = getattr(change, "parts", None)
        if parts is not None:
            for name in touched:
                if name not in parts:
                    parts.append(name)


def _package_tracked_change_date(self: object) -> object:
    """The date a new revision carries, or None for the wall clock (section 3.8).

    Fixing it is what makes a tracked edit byte-reproducible, as ``id_seed``
    does for the ids: ``pkg.tracked_change_date = datetime(2026, 9, 17, tzinfo=UTC)``.
    """
    return self._tracked_change_date  # type: ignore[attr-defined]


def _set_package_tracked_change_date(self: object, value: object) -> None:
    """Fix the ``w:date`` of every new revision, or None to use the wall clock."""
    import datetime as _datetime

    if value is not None and not isinstance(value, _datetime.datetime):
        from docx4j_py.model.content.errors import ContentError

        raise ContentError(
            f"pkg.tracked_change_date takes a datetime or None, not {type(value).__name__}",
            code="tracking.date_invalid",
            hint="pkg.tracked_change_date = datetime.datetime(2026, 9, 17, tzinfo=datetime.UTC)",
        )
    self._tracked_change_date = value  # type: ignore[attr-defined]


def _package_get_tracked_changes(self: object) -> list:
    """Every tracked change in the main document part, in document order."""
    return list(self.body.get_tracked_changes())  # type: ignore[attr-defined]


def _package_custom_xml_parts(self: object) -> object:
    """The document's custom XML parts (CR-003 section 3.7, Phase E).

    ``pkg.custom_xml_parts.get_item(id)``, ``.add(xml)``, ``.apply_bindings()``,
    ``.describe()`` and ``.fill(data)``, over
    :mod:`docx4j_py.model.customxml`, which is imported here on **first use** so
    that a caller who never touches a content control pays nothing for it.
    Reading the collection unmarshals nothing: a part that is only read is
    written back byte for byte.
    """
    from docx4j_py.model.customxml.parts import custom_xml_parts_of

    return custom_xml_parts_of(self)


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
    # CR-003 Phase K, section 3.5
    "to_markdown": _package_to_markdown,
    "markdown_budget": _package_markdown_budget,
    "insert_markdown": _package_insert_markdown,
    # CR-003 Phase F, section 3.8
    "get_tracked_changes": _package_get_tracked_changes,
}

#: The properties :func:`register` installs, which ``_PACKAGE_MEMBERS`` cannot
#: hold because a property is not a function. CR-003 Phase G, section 3.9.
_PACKAGE_PROPERTIES: dict[str, property] = {
    "author": property(_package_author, _set_package_author, doc=_package_author.__doc__),
    # CR-003 Phase F, section 3.8
    "change_tracking_mode": property(
        _package_change_tracking_mode,
        _set_package_change_tracking_mode,
        doc=_package_change_tracking_mode.__doc__,
    ),
    "tracked_change_date": property(
        _package_tracked_change_date,
        _set_package_tracked_change_date,
        doc=_package_tracked_change_date.__doc__,
    ),
    # CR-003 section 3.4, added 2026-09-17
    "compatibility_mode": property(
        _package_compatibility_mode,
        _set_package_compatibility_mode,
        doc=_package_compatibility_mode.__doc__,
    ),
    # CR-003 Phase E, section 3.7
    "custom_xml_parts": property(
        _package_custom_xml_parts, doc=_package_custom_xml_parts.__doc__
    ),
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
    for name, member in _PACKAGE_PROPERTIES.items():
        if not isinstance(getattr(WordprocessingMLPackage, name, None), property):
            setattr(WordprocessingMLPackage, name, member)


register()
