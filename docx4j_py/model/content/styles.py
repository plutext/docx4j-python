"""Style ids, display names and Office JS's ``Word.Style`` values.

CR-003 section 4, the rule the TypeScript engine had to correct itself into
(its CR-002 section 11):

``style``
    the style's **display name** --- ``"Heading 1"``, ``"Table Grid"``, and in
    Word localised. Read from the styles part's ``w:name`` when that part is
    unmarshalled (**nothing is unmarshalled for a read**), with Word's stored
    lower-case built-in names (``heading 1``, ``toc 1``, ``annotation text``)
    mapped to their display names through :data:`BUILT_IN_STYLES`, which is
    docx4j's ``KnownStyles.xml``; otherwise derived from the id by inserting
    spaces, which is exact for the built-ins because Word derives the id from
    the English display name by removing them.
``style_built_in``
    the ``Word.Style`` value --- ``"Heading1"``, ``"Toc1"`` for Word's id
    ``TOC1`` --- or ``"Other"`` for a style that is not built in, which cannot
    be set, as in Office JS.
``style_id``
    the docx4j-named extension: what is really in ``w:pStyle`` / ``w:tblStyle``.

Setting accepts any of the three. A name that the document's styles part does
not define raises :class:`~docx4j_py.model.content.errors.StyleError` listing
the five closest names, **unless** it resolves to a built-in style, in which
case the setter **adds the definition** --- docx4j's ``KnownStyles.xml``, which
is what :func:`ensure_style` does --- before writing the id, and the call's
``ChangeReport`` lists ``/word/styles.xml`` in ``parts_touched``.

Phase B wrote such an id as it stands, on the premise that "Word creates the
definition when it opens the document". That premise is false, as Word showed
on 2026-09-17: a ``w:pStyle`` naming a style ``styles.xml`` does not define
falls back to Normal, and a ``w:tblStyle`` to Table Normal, silently and with
no repair prompt. CR-003 section 14.9 records the correction. Two cases still
write the id as it stands, because there is nothing to add it to or nothing to
add:

* a document with **no styles part at all**;
* the seven modern table styles (``PlainTable1`` ... ``ListTable1Light``) that
  Office JS's ``Word.Style`` has and docx4j's ``KnownStyles.xml``, which
  predates them, does not.

``style_id`` is the escape hatch and stays raw: setting it writes the id
exactly as given, defined or not. Validating means unmarshalling the styles
part and defining means re-marshalling it; **a read does neither**, and only a
setter that has to add a definition unmarshals anything.
"""

from __future__ import annotations

import dataclasses
import difflib
import re
from typing import Any

from docx4j_py.child import ChildList, deep_copy, link_parents

__all__ = [
    "BUILT_IN_STYLES",
    "CUSTOM_STYLE_XML",
    "WORD_STYLE_VALUES",
    "BuiltInStyle",
    "built_in_of",
    "define_built_in",
    "display_name_of",
    "ensure_style",
    "id_of_built_in",
    "style_id_of",
    "style_ids_of",
    "style_name_of",
    "styles_of",
]


@dataclasses.dataclass(frozen=True, slots=True)
class BuiltInStyle:
    """One of Word's built-in styles, as docx4j's ``KnownStyles.xml`` has it."""

    style_id: str
    """``w:styleId``: ``"Heading1"``, ``"TOC1"``, ``"TableNormal"``."""
    name: str
    """The display name Word shows: ``"Heading 1"``, ``"TOC 1"``, ``"Normal Table"``."""
    built_in: str | None
    """The Office JS ``Word.Style`` value, or None when Office JS has no name for it."""
    type: str
    """``paragraph``, ``character``, ``table`` or ``numbering``."""
    stored_name: str | None = None
    """``w:name`` as Word writes it, when it differs from the display name."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view."""
        return dataclasses.asdict(self)


#: Word's built-in styles: docx4j's ``KnownStyles.xml`` (a copy of which is
#: ``docx4j_py/resources/KnownStyles.xml``), plus the seven modern table styles
#: ``Word.Style`` names that it predates. One committed list, as section 4 asks.
BUILT_IN_STYLES: tuple[BuiltInStyle, ...] = (
    BuiltInStyle("Normal", "Normal", "Normal", "paragraph", None),
    BuiltInStyle("Heading1", "Heading 1", "Heading1", "paragraph", "heading 1"),
    BuiltInStyle("Heading2", "Heading 2", "Heading2", "paragraph", "heading 2"),
    BuiltInStyle("Heading3", "Heading 3", "Heading3", "paragraph", "heading 3"),
    BuiltInStyle("Heading4", "Heading 4", "Heading4", "paragraph", "heading 4"),
    BuiltInStyle("Heading5", "Heading 5", "Heading5", "paragraph", "heading 5"),
    BuiltInStyle("Heading6", "Heading 6", "Heading6", "paragraph", "heading 6"),
    BuiltInStyle("Heading7", "Heading 7", "Heading7", "paragraph", "heading 7"),
    BuiltInStyle("Heading8", "Heading 8", "Heading8", "paragraph", "heading 8"),
    BuiltInStyle("Heading9", "Heading 9", "Heading9", "paragraph", "heading 9"),
    BuiltInStyle("DefaultParagraphFont", "Default Paragraph Font", None, "character", None),
    BuiltInStyle("TableNormal", "Normal Table", None, "table", None),
    BuiltInStyle("NoList", "No List", None, "numbering", None),
    BuiltInStyle("111111", "Outline List 2", None, "numbering", None),
    BuiltInStyle("1ai", "Outline List 1", None, "numbering", None),
    BuiltInStyle("Heading1Char", "Heading 1 Char", None, "character", None),
    BuiltInStyle("Heading2Char", "Heading 2 Char", None, "character", None),
    BuiltInStyle("Heading3Char", "Heading 3 Char", None, "character", None),
    BuiltInStyle("Heading4Char", "Heading 4 Char", None, "character", None),
    BuiltInStyle("Heading5Char", "Heading 5 Char", None, "character", None),
    BuiltInStyle("Heading6Char", "Heading 6 Char", None, "character", None),
    BuiltInStyle("Heading7Char", "Heading 7 Char", None, "character", None),
    BuiltInStyle("Heading8Char", "Heading 8 Char", None, "character", None),
    BuiltInStyle("Heading9Char", "Heading 9 Char", None, "character", None),
    BuiltInStyle("ArticleSection", "Outline List 3", None, "numbering", None),
    BuiltInStyle("BalloonText", "Balloon Text", None, "paragraph", None),
    BuiltInStyle("BalloonTextChar", "Balloon Text Char", None, "character", None),
    BuiltInStyle("Bibliography", "Bibliography", "Bibliography", "paragraph", None),
    BuiltInStyle("BlockText", "Block Text", None, "paragraph", None),
    BuiltInStyle("BodyText", "Body Text", None, "paragraph", None),
    BuiltInStyle("BodyTextChar", "Body Text Char", None, "character", None),
    BuiltInStyle("BodyText2", "Body Text 2", None, "paragraph", None),
    BuiltInStyle("BodyText2Char", "Body Text 2 Char", None, "character", None),
    BuiltInStyle("BodyText3", "Body Text 3", None, "paragraph", None),
    BuiltInStyle("BodyText3Char", "Body Text 3 Char", None, "character", None),
    BuiltInStyle("BodyTextFirstIndent", "Body Text First Indent", None, "paragraph", None),
    BuiltInStyle("BodyTextFirstIndentChar", "Body Text First Indent Char", None, "character", None),
    BuiltInStyle("BodyTextIndent", "Body Text Indent", None, "paragraph", None),
    BuiltInStyle("BodyTextIndentChar", "Body Text Indent Char", None, "character", None),
    BuiltInStyle("BodyTextFirstIndent2", "Body Text First Indent 2", None, "paragraph", None),
    BuiltInStyle("BodyTextFirstIndent2Char", "Body Text First Indent 2 Char", None, "character", None),
    BuiltInStyle("BodyTextIndent2", "Body Text Indent 2", None, "paragraph", None),
    BuiltInStyle("BodyTextIndent2Char", "Body Text Indent 2 Char", None, "character", None),
    BuiltInStyle("BodyTextIndent3", "Body Text Indent 3", None, "paragraph", None),
    BuiltInStyle("BodyTextIndent3Char", "Body Text Indent 3 Char", None, "character", None),
    BuiltInStyle("BookTitle", "Book Title", "BookTitle", "character", None),
    BuiltInStyle("Caption", "Caption", "Caption", "paragraph", "caption"),
    BuiltInStyle("Closing", "Closing", None, "paragraph", None),
    BuiltInStyle("ClosingChar", "Closing Char", None, "character", None),
    BuiltInStyle("ColorfulGrid", "Colorful Grid", None, "table", None),
    BuiltInStyle("CommentText", "Comment Text", None, "paragraph", "annotation text"),
    BuiltInStyle("CommentTextChar", "Comment Text Char", None, "character", None),
    BuiltInStyle("CommentSubject", "Comment Subject", None, "paragraph", "annotation subject"),
    BuiltInStyle("CommentSubjectChar", "Comment Subject Char", None, "character", None),
    BuiltInStyle("Date", "Date", None, "paragraph", None),
    BuiltInStyle("DateChar", "Date Char", None, "character", None),
    BuiltInStyle("DocumentMap", "Document Map", None, "paragraph", None),
    BuiltInStyle("DocumentMapChar", "Document Map Char", None, "character", None),
    BuiltInStyle("E-mailSignature", "E-mail Signature", None, "paragraph", None),
    BuiltInStyle("E-mailSignatureChar", "E-mail Signature Char", None, "character", None),
    BuiltInStyle("EndnoteText", "Endnote Text", "EndnoteText", "paragraph", "endnote text"),
    BuiltInStyle("EndnoteTextChar", "Endnote Text Char", None, "character", None),
    BuiltInStyle("EnvelopeAddress", "Envelope Address", None, "paragraph", "envelope address"),
    BuiltInStyle("EnvelopeReturn", "Envelope Return", None, "paragraph", "envelope return"),
    BuiltInStyle("Footer", "Footer", "Footer", "paragraph", "footer"),
    BuiltInStyle("FooterChar", "Footer Char", None, "character", None),
    BuiltInStyle("FootnoteText", "Footnote Text", "FootnoteText", "paragraph", "footnote text"),
    BuiltInStyle("FootnoteTextChar", "Footnote Text Char", None, "character", None),
    BuiltInStyle("Header", "Header", "Header", "paragraph", "header"),
    BuiltInStyle("HeaderChar", "Header Char", None, "character", None),
    BuiltInStyle("HTMLAddress", "HTML Address", None, "paragraph", None),
    BuiltInStyle("HTMLAddressChar", "HTML Address Char", None, "character", None),
    BuiltInStyle("HTMLPreformatted", "HTML Preformatted", None, "paragraph", None),
    BuiltInStyle("HTMLPreformattedChar", "HTML Preformatted Char", None, "character", None),
    BuiltInStyle("Index1", "Index 1", None, "paragraph", "index 1"),
    BuiltInStyle("Index2", "Index 2", None, "paragraph", "index 2"),
    BuiltInStyle("Index3", "Index 3", None, "paragraph", "index 3"),
    BuiltInStyle("Index4", "Index 4", None, "paragraph", "index 4"),
    BuiltInStyle("Index5", "Index 5", None, "paragraph", "index 5"),
    BuiltInStyle("Index6", "Index 6", None, "paragraph", "index 6"),
    BuiltInStyle("Index7", "Index 7", None, "paragraph", "index 7"),
    BuiltInStyle("Index8", "Index 8", None, "paragraph", "index 8"),
    BuiltInStyle("Index9", "Index 9", None, "paragraph", "index 9"),
    BuiltInStyle("IndexHeading", "Index Heading", None, "paragraph", "index heading"),
    BuiltInStyle("IntenseQuote", "Intense Quote", "IntenseQuote", "paragraph", None),
    BuiltInStyle("IntenseQuoteChar", "Intense Quote Char", None, "character", None),
    BuiltInStyle("List", "List", None, "paragraph", None),
    BuiltInStyle("List2", "List 2", None, "paragraph", None),
    BuiltInStyle("List3", "List 3", None, "paragraph", None),
    BuiltInStyle("List4", "List 4", None, "paragraph", None),
    BuiltInStyle("List5", "List 5", None, "paragraph", None),
    BuiltInStyle("ListBullet", "List Bullet", None, "paragraph", None),
    BuiltInStyle("ListBullet2", "List Bullet 2", None, "paragraph", None),
    BuiltInStyle("ListBullet3", "List Bullet 3", None, "paragraph", None),
    BuiltInStyle("ListBullet4", "List Bullet 4", None, "paragraph", None),
    BuiltInStyle("ListBullet5", "List Bullet 5", None, "paragraph", None),
    BuiltInStyle("ListContinue", "List Continue", None, "paragraph", None),
    BuiltInStyle("ListContinue2", "List Continue 2", None, "paragraph", None),
    BuiltInStyle("ListContinue3", "List Continue 3", None, "paragraph", None),
    BuiltInStyle("ListContinue4", "List Continue 4", None, "paragraph", None),
    BuiltInStyle("ListContinue5", "List Continue 5", None, "paragraph", None),
    BuiltInStyle("ListNumber", "List Number", None, "paragraph", None),
    BuiltInStyle("ListNumber2", "List Number 2", None, "paragraph", None),
    BuiltInStyle("ListNumber3", "List Number 3", None, "paragraph", None),
    BuiltInStyle("ListNumber4", "List Number 4", None, "paragraph", None),
    BuiltInStyle("ListNumber5", "List Number 5", None, "paragraph", None),
    BuiltInStyle("ListParagraph", "List Paragraph", "ListParagraph", "paragraph", None),
    BuiltInStyle("MacroText", "Macro Text", None, "paragraph", "macro"),
    BuiltInStyle("MacroTextChar", "Macro Text Char", None, "character", None),
    BuiltInStyle("MessageHeader", "Message Header", None, "paragraph", None),
    BuiltInStyle("MessageHeaderChar", "Message Header Char", None, "character", None),
    BuiltInStyle("NoSpacing", "No Spacing", "NoSpacing", "paragraph", None),
    BuiltInStyle("NormalWeb", "Normal (Web)", None, "paragraph", None),
    BuiltInStyle("NormalIndent", "Normal Indent", None, "paragraph", None),
    BuiltInStyle("NoteHeading", "Note Heading", None, "paragraph", None),
    BuiltInStyle("NoteHeadingChar", "Note Heading Char", None, "character", None),
    BuiltInStyle("PlainText", "Plain Text", None, "paragraph", None),
    BuiltInStyle("PlainTextChar", "Plain Text Char", None, "character", None),
    BuiltInStyle("Quote", "Quote", "Quote", "paragraph", None),
    BuiltInStyle("QuoteChar", "Quote Char", None, "character", None),
    BuiltInStyle("Salutation", "Salutation", None, "paragraph", None),
    BuiltInStyle("SalutationChar", "Salutation Char", None, "character", None),
    BuiltInStyle("Signature", "Signature", None, "paragraph", None),
    BuiltInStyle("SignatureChar", "Signature Char", None, "character", None),
    BuiltInStyle("Subtitle", "Subtitle", "Subtitle", "paragraph", None),
    BuiltInStyle("SubtitleChar", "Subtitle Char", None, "character", None),
    BuiltInStyle("TableofAuthorities", "Table of Authorities", None, "paragraph", "table of authorities"),
    BuiltInStyle("TableofFigures", "Table of Figures", None, "paragraph", "table of figures"),
    BuiltInStyle("Title", "Title", "Title", "paragraph", None),
    BuiltInStyle("TitleChar", "Title Char", None, "character", None),
    BuiltInStyle("TOAHeading", "TOA Heading", None, "paragraph", "toa heading"),
    BuiltInStyle("TOC1", "TOC 1", "Toc1", "paragraph", "toc 1"),
    BuiltInStyle("TOC2", "TOC 2", "Toc2", "paragraph", "toc 2"),
    BuiltInStyle("TOC3", "TOC 3", "Toc3", "paragraph", "toc 3"),
    BuiltInStyle("TOC4", "TOC 4", "Toc4", "paragraph", "toc 4"),
    BuiltInStyle("TOC5", "TOC 5", "Toc5", "paragraph", "toc 5"),
    BuiltInStyle("TOC6", "TOC 6", "Toc6", "paragraph", "toc 6"),
    BuiltInStyle("TOC7", "TOC 7", "Toc7", "paragraph", "toc 7"),
    BuiltInStyle("TOC8", "TOC 8", "Toc8", "paragraph", "toc 8"),
    BuiltInStyle("TOC9", "TOC 9", "Toc9", "paragraph", "toc 9"),
    BuiltInStyle("TOCHeading", "TOC Heading", "TocHeading", "paragraph", None),
    BuiltInStyle("CommentReference", "Comment Reference", None, "character", "annotation reference"),
    BuiltInStyle("Emphasis", "Emphasis", "Emphasis", "character", None),
    BuiltInStyle("EndnoteReference", "Endnote Reference", "EndnoteReference", "character", "endnote reference"),
    BuiltInStyle("FollowedHyperlink", "FollowedHyperlink", None, "character", None),
    BuiltInStyle("FootnoteReference", "Footnote Reference", "FootnoteReference", "character", "footnote reference"),
    BuiltInStyle("HTMLAcronym", "HTML Acronym", None, "character", None),
    BuiltInStyle("HTMLCite", "HTML Cite", None, "character", None),
    BuiltInStyle("HTMLCode", "HTML Code", None, "character", None),
    BuiltInStyle("HTMLDefinition", "HTML Definition", None, "character", None),
    BuiltInStyle("HTMLKeyboard", "HTML Keyboard", None, "character", None),
    BuiltInStyle("HTMLSample", "HTML Sample", None, "character", None),
    BuiltInStyle("HTMLTypewriter", "HTML Typewriter", None, "character", None),
    BuiltInStyle("HTMLVariable", "HTML Variable", None, "character", None),
    BuiltInStyle("Hyperlink", "Hyperlink", "Hyperlink", "character", None),
    BuiltInStyle("IntenseEmphasis", "Intense Emphasis", "IntenseEmphasis", "character", None),
    BuiltInStyle("IntenseReference", "Intense Reference", "IntenseReference", "character", None),
    BuiltInStyle("LineNumber", "Line Number", None, "character", "line number"),
    BuiltInStyle("PageNumber", "Page Number", None, "character", "page number"),
    BuiltInStyle("PlaceholderText", "Placeholder Text", None, "character", None),
    BuiltInStyle("Strong", "Strong", "Strong", "character", None),
    BuiltInStyle("SubtleEmphasis", "Subtle Emphasis", "SubtleEmphasis", "character", None),
    BuiltInStyle("SubtleReference", "Subtle Reference", "SubtleReference", "character", None),
    BuiltInStyle("TableGrid", "Table Grid", "TableGrid", "table", None),
    BuiltInStyle("PlainTable1", "Plain Table 1", "PlainTable1", "table", None),
    BuiltInStyle("PlainTable2", "Plain Table 2", "PlainTable2", "table", None),
    BuiltInStyle("PlainTable3", "Plain Table 3", "PlainTable3", "table", None),
    BuiltInStyle("PlainTable4", "Plain Table 4", "PlainTable4", "table", None),
    BuiltInStyle("PlainTable5", "Plain Table 5", "PlainTable5", "table", None),
    BuiltInStyle("GridTable1Light", "Grid Table 1 Light", "GridTable1Light", "table", None),
    BuiltInStyle("ListTable1Light", "List Table 1 Light", "ListTable1Light", "table", None),
)

#: The Office JS ``Word.Style`` values, in Office JS's order, ``"Other"``
#: excluded (it is read, never set).
WORD_STYLE_VALUES: tuple[str, ...] = tuple(
    dict.fromkeys(s.built_in for s in BUILT_IN_STYLES if s.built_in is not None)
)

_BY_ID: dict[str, BuiltInStyle] = {s.style_id.lower(): s for s in BUILT_IN_STYLES}
_BY_BUILT_IN: dict[str, BuiltInStyle] = {
    s.built_in.lower(): s for s in BUILT_IN_STYLES if s.built_in is not None
}
_BY_NAME: dict[str, BuiltInStyle] = {s.name.lower(): s for s in BUILT_IN_STYLES}
_BY_STORED: dict[str, BuiltInStyle] = {
    (s.stored_name or s.name).lower(): s for s in BUILT_IN_STYLES
}

#: Where a lower-case id and a lower-case display name are the same thing.
_COMPACT: dict[str, BuiltInStyle] = {s.name.replace(" ", "").lower(): s for s in BUILT_IN_STYLES}

_SPACED = re.compile(r"(?<=[a-z0-9])(?=[A-Z])|(?<=[A-Za-z])(?=[0-9])")


def built_in_of(style_id: str) -> str:
    """Office JS's ``style_built_in`` for a style id: a ``Word.Style`` value or ``"Other"``.

    ``built_in_of("TOC1")`` is ``"Toc1"``, ``built_in_of("MyStyle")`` is
    ``"Other"``. Spaces are ignored, so a display name works too.
    """
    entry = _BY_ID.get(style_id.replace(" ", "").lower()) or _BY_BUILT_IN.get(
        style_id.replace(" ", "").lower()
    )
    if entry is not None and entry.built_in is not None:
        return entry.built_in
    return "Other"


def id_of_built_in(value: str) -> str:
    """The style id for a ``Word.Style`` value, in Word's spelling (``"TOC1"``).

    Leniently accepts the display name (``"Heading 1"``) as well, as the
    TypeScript engine does. ``"Other"`` cannot be set, as in Office JS.

    Raises:
        StyleError: `value` is ``"Other"``.
    """
    from docx4j_py.model.content.errors import StyleError

    if value == "Other":
        raise StyleError(
            "style_built_in cannot be set to 'Other'",
            code="style.other",
            hint="set style to the display name of a custom style, or style_id to its id",
        )
    compact = value.replace(" ", "")
    lower = compact.lower()
    entry = _BY_BUILT_IN.get(lower) or _BY_ID.get(lower) or _COMPACT.get(lower)
    return entry.style_id if entry is not None else compact


def display_name_of(style_id: str, stored_name: str | None = None) -> str:
    """The display name of a style id, given the ``w:name`` the document stores.

    The three cases of section 4, in order: a built-in id answers from
    :data:`BUILT_IN_STYLES`; a custom style answers with its own stored name;
    and with no stored name at all the id gets its spaces back
    (``"MyOwnStyle"`` to ``"My Own Style"``).
    """
    entry = _BY_ID.get(style_id.lower())
    if entry is not None:
        return entry.name
    if stored_name:
        return stored_name
    return _SPACED.sub(" ", style_id)


def styles_of(package: Any, *, unmarshal: bool = False) -> list[Any]:
    """The ``w:style`` list of the package's styles part.

    Empty unless the part is already unmarshalled: a *read* of ``style`` must
    not mark ``styles.xml`` for re-marshalling, which is what section 4 means by
    "nothing is unmarshalled for it". `unmarshal` is the opt-in a *write* takes,
    because a write has to validate the name against what the document defines.
    """
    part = getattr(package, "style_definitions_part", None)
    if part is None:
        return []
    if not unmarshal and not part.is_unmarshalled:
        return []
    try:
        styles = part.contents
    except Exception:  # noqa: BLE001 - a styles part that will not parse is no styles
        return []
    return list(getattr(styles, "style", None) or ())


def _stored_name(style: Any) -> str | None:
    return getattr(getattr(style, "name", None), "val", None)


def style_name_of(package: Any, style_id: str) -> str:
    """Office JS's ``style``: the display name of a style id, in this package."""
    for style in styles_of(package):
        if getattr(style, "style_id", None) == style_id:
            return display_name_of(style_id, _stored_name(style))
    return display_name_of(style_id)


def style_id_of(package: Any, name: str, *, validate: bool = False, define: bool = False) -> str:
    """The style id for what ``style`` accepts: a display name, a stored name or an id.

    Resolved against the styles part when it is unmarshalled (or when
    `validate` is on, which unmarshals it), and otherwise against
    :data:`BUILT_IN_STYLES` and by removing spaces.

    Args:
        package: the package whose styles part to resolve against; None for
            a body with no package.
        name: a display name (``"Heading 1"``), a stored name (``"heading 1"``)
            or an id (``"Heading1"``).
        validate: refuse a name the document does not define and that is not a
            built-in style.
        define: a built-in style the document does not define is **defined**
            (:func:`define_built_in`) rather than left dangling. Every setter
            passes it; ``style_id`` does not (section 14.9).

    Raises:
        StyleError: `validate` is on and nothing matches; the message lists the
            five closest names the document does define.
    """
    lower = name.lower()
    compact = name.replace(" ", "")
    styles = styles_of(package, unmarshal=validate)
    for style in styles:
        style_id = getattr(style, "style_id", None)
        if not style_id:
            continue
        stored = _stored_name(style)
        if (
            style_id.lower() == lower
            or (stored or "").lower() == lower
            or display_name_of(style_id, stored).lower() == lower
        ):
            return style_id

    entry = (
        _BY_NAME.get(lower)
        or _BY_STORED.get(lower)
        or _BY_ID.get(compact.lower())
        or _BY_BUILT_IN.get(compact.lower())
    )
    if entry is not None:
        return define_built_in(package, entry.style_id) if define else entry.style_id
    if not validate:
        return compact

    from docx4j_py.model.content.errors import StyleError

    known: list[str] = []
    for style in styles:
        style_id = getattr(style, "style_id", None)
        if style_id:
            known.append(display_name_of(style_id, _stored_name(style)))
    closest = difflib.get_close_matches(name, sorted(set(known)), n=5, cutoff=0.0)
    listed = ", ".join(repr(c) for c in closest) if closest else "none"
    raise StyleError(
        f"no style named {name!r} in this document; the closest are {listed}",
        hint=(
            "pass one of those names, a built-in style such as 'Heading 1', "
            "or set style_id to write the id as it stands"
        ),
    )


# ---------------------------------------------------------------------------
# defining a style the document lacks (CR-003 section 14.9; Phase K's machinery,
# moved here from docx4j_py.model.markdown.importer, where the rule does not live)
# ---------------------------------------------------------------------------

_W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

#: The two styles Word has no built-in equivalent of --- the ones markdown's
#: code spans and code blocks need --- exactly as Java ``ImportStyles`` writes
#: them. Everything else comes from ``KnownStyles.xml``.
CUSTOM_STYLE_XML: dict[str, str] = {
    "CodeChar": (
        f'<w:style {_W} w:type="character" w:styleId="CodeChar">'
        '<w:name w:val="Code Char"/>'
        "<w:rPr>"
        '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>'
        '<w:sz w:val="20"/><w:szCs w:val="20"/>'
        '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>'
        "</w:rPr>"
        "</w:style>"
    ),
    "SourceCode": (
        f'<w:style {_W} w:type="paragraph" w:styleId="SourceCode">'
        '<w:name w:val="Source Code"/>'
        '<w:basedOn w:val="Normal"/>'
        "<w:pPr>"
        "<w:keepLines/>"
        '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
        '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>'
        "</w:pPr>"
        "<w:rPr>"
        '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>'
        "<w:noProof/>"
        '<w:sz w:val="20"/><w:szCs w:val="20"/>'
        "</w:rPr>"
        "</w:style>"
    ),
}

_known_styles_cache: dict[str, Any] | None = None


def _known_styles() -> dict[str, Any]:
    """docx4j's ``KnownStyles.xml``, by style id, parsed once per process."""
    global _known_styles_cache  # one cache for the process
    if _known_styles_cache is None:
        from importlib import resources

        from docx4j_py.openpackaging.parts.wml import StyleDefinitionsPart

        part = StyleDefinitionsPart()
        part.set_bytes((resources.files("docx4j_py.resources") / "KnownStyles.xml").read_bytes())
        found: dict[str, Any] = {}
        for style in getattr(part.contents, "style", None) or ():
            style_id = getattr(style, "style_id", None)
            if style_id:
                found[str(style_id)] = style
        _known_styles_cache = found
    return _known_styles_cache


def style_ids_of(part: Any) -> set[str]:
    """The style ids a styles part defines, **without unmarshalling it**.

    Read from the part's bytes with lxml when the part is untouched, the way
    :mod:`docx4j_py.model.content.describe` reads, so that a style the document
    already has costs no re-marshalling.
    """
    ids: set[str] = set()
    if part.is_unmarshalled:
        for style in getattr(part.contents, "style", None) or ():
            style_id = getattr(style, "style_id", None)
            if style_id:
                ids.add(str(style_id))
    else:
        from lxml import etree

        w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        try:
            root = etree.fromstring(part.xml)
        except Exception:  # noqa: BLE001 - a styles part that will not parse defines none
            root = None
        if root is not None:
            for style in root.findall(f"{w}style"):
                style_id = style.get(f"{w}styleId")
                if style_id:
                    ids.add(style_id)
    return ids


def ensure_style(
    package: Any,
    style_id: str,
    *,
    touched: set[str] | None = None,
    defined: set[str] | None = None,
) -> str:
    """Make sure the document defines `style_id`, adding the definition if not.

    docx4j ``ImportStyles.ensureKnown`` / ``ensureCustom``: a style the
    template already defines is **never** replaced; one docx4j knows
    (``KnownStyles.xml``, 164 of them) is activated on demand; the two code
    styles Word has no equivalent of come from :data:`CUSTOM_STYLE_XML`.

    **This unmarshals ``/word/styles.xml``** when it has to add anything, which
    marks that part for re-marshalling on the next save. It reads without
    unmarshalling, so a style that is already there costs nothing.

    Args:
        package: the ``WordprocessingMLPackage``.
        style_id: ``"Heading1"``, ``"ListParagraph"``, ``"CodeChar"``, ...
        touched: a set the part name is added to when the part was changed.
        defined: the ids the part already has, if the caller is keeping them;
            a style added here is added to it. Computed from the part when not
            given, which costs one lxml parse.

    Returns:
        The style id, which is `style_id`.

    Raises:
        StyleError: the id is neither defined, nor known to docx4j, nor one of
            the two custom ones.
    """
    from docx4j_py.model.content.errors import StyleError

    part = getattr(package, "style_definitions_part", None) if package is not None else None
    if part is None:
        # no styles part: there is nowhere to put a definition, so the id goes
        # in as it stands and Word resolves it against its own latent styles
        return style_id
    if defined is None:
        defined = style_ids_of(part)
    if style_id in defined:
        return style_id

    definition = _known_styles().get(style_id)
    if definition is not None:
        new = deep_copy(definition)
    elif style_id in CUSTOM_STYLE_XML:
        from docx4j_py.wml import wml

        new = wml(CUSTOM_STYLE_XML[style_id], wrapper="styles")
    else:
        raise StyleError(
            f"this document does not define the style {style_id!r} and docx4j does not know it",
            code="style.not_creatable",
            hint="add the style to the document, or pass a style docx4j's KnownStyles.xml has",
        )
    styles = part.contents
    if styles.style is None:
        styles.style = ChildList([], owner=styles)
    styles.style.append(new)
    link_parents(new)
    new.parent = styles
    defined.add(style_id)
    if touched is not None:
        touched.add(str(part.part_name))
    return style_id


def define_built_in(package: Any, style_id: str) -> str:
    """Define `style_id` when it is a built-in style this document does not define.

    The Phase B correction of CR-003 section 14.9: Word does **not** materialise
    a definition for a dangling ``w:pStyle`` / ``w:tblStyle`` id, it falls back
    to Normal, so a setter writing a built-in id has to put the definition in
    first. What a setter calls; ``style_id`` deliberately does not.

    Leaves everything else alone --- a custom id (the ``style_id`` escape
    hatch), a document with no styles part, and the seven modern table styles
    ``KnownStyles.xml`` predates --- and never unmarshals the styles part unless
    it adds something. When it does add, ``/word/styles.xml`` goes into the
    :class:`~docx4j_py.model.content.reports.ChangeReport` being collected.

    Returns:
        `style_id`, unchanged; the side effect is the definition.
    """
    from docx4j_py.model.content.errors import StyleError

    if package is None or not style_id or style_id.lower() not in _BY_ID:
        return style_id
    touched: set[str] = set()
    try:
        ensure_style(package, style_id, touched=touched)
    except StyleError as error:
        if error.code != "style.not_creatable":
            raise
        # a Word.Style value docx4j's KnownStyles.xml has no definition for:
        # the seven modern table styles. Nothing to add, so write the id.
        return style_id
    if touched:
        from docx4j_py.model.content.reports import current_recorder

        parts = getattr(current_recorder(package), "parts", None)
        if parts is not None:
            for name in sorted(touched):
                if name not in parts:
                    parts.append(name)
    return style_id
