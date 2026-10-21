"""The block model (WD sections 7.4 and 8, REQ-706). CR-011 section 3.1.

docx4j's ``org.docx4j.model.text`` result classes --- ``TextResult``,
``TextBlock``, ``TextRun``, ``TextRow``, ``TextCell``, ``TextField``,
``TextNumbering``, ``TextRevision``, ``TextWarning``, ``TextComment`` --- as
core-ts CR-012 ports them. docx4j's enum values are kept verbatim as strings
(``"PARAGRAPH"``, ``"TEXT_BOX"``, ``"NOTE_REFERENCE"``), so that a block model
passed as JSON reads the same from every implementation.

A block's and a run's ``source`` (the tree objects it was read from) and a
run's ``office_offset`` are this package's (REQ-202's emulated-API exemption,
CR-011 Q4): they are not in :meth:`TextBlock.to_dict`, which holds no tree
object (REQ-706).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Literal

from docx4j_py.model.text.images import ImageDescription, TextImage
from docx4j_py.model.text.options import TextOptions

__all__ = [
    "TextBlock",
    "TextCell",
    "TextComment",
    "TextField",
    "TextNumbering",
    "TextResult",
    "TextRevision",
    "TextRow",
    "TextRun",
    "TextRunSource",
    "TextWarning",
]

#: What a block is (docx4j ``TextBlock.Kind``).
BlockKind = Literal["PARAGRAPH", "TABLE", "CHART"]
#: The story a block is in (WD section 5.1; docx4j ``TextBlock.Story``).
Story = Literal["MAIN", "TEXT_BOX", "FOOTNOTE", "ENDNOTE", "HEADER", "FOOTER", "COMMENT", "GLOSSARY", "SMARTART", "EMBEDDED"]
#: What a run is (docx4j ``TextRun.Kind``).
RunKind = Literal[
    "TEXT", "FIELD_CODE", "NOTE_REFERENCE", "NOTE_MARK", "COMMENT_REFERENCE", "LIST_LABEL", "IMAGE", "OBJECT", "PAGE_MARKER"
]


@dataclass(frozen=True, slots=True)
class TextRevision:
    """A run's revision in the ``markup`` view (REQ-706). docx4j ``TextRevision``.

    ``w:ins`` and ``w:moveTo`` are insertions, ``w:del`` and ``w:moveFrom``
    deletions, each with its author and its date as the package wrote it.
    """

    type: Literal["INSERTION", "DELETION"]
    author: str | None = None
    date: str | None = None

    def __str__(self) -> str:
        """``INSERTION by Author A 2026-10-01T00:00:00Z``."""
        return self.type + (f" by {self.author}" if self.author else "") + (f" {self.date}" if self.date else "")


@dataclass(frozen=True, slots=True)
class TextRunSource:
    """Where a run's text was read from: its ``w:r`` and its ``w:p``."""

    run: Any
    paragraph: Any


@dataclass(slots=True, eq=False)
class TextRun:
    """A piece of a block's content (REQ-706): document text, or a typed item the serialiser writes. docx4j ``TextRun``.

    A note reference is written ``[n]`` in plain text (REQ-707); here it is a
    field of the run, never text.
    """

    kind: RunKind
    #: The document's text for ``TEXT``, ``FIELD_CODE`` and ``LIST_LABEL``; empty for the others.
    text: str = ""
    #: The run's revision in the ``markup`` view; None in the other views and for an unrevised run.
    revision: TextRevision | None = None
    #: Hidden text included under ``hidden: true`` (REQ-507).
    hidden: bool = False
    #: A content control's placeholder text (REQ-508).
    placeholder: bool = False
    #: For a note reference or mark: ``"footnote"`` or ``"endnote"``.
    note_type: str | None = None
    #: For a note reference or mark: its number in reference order, from 1; 0 for a custom mark.
    note_number: int = 0
    #: For a note reference or mark: the note's ``w:id``.
    note_id: str | None = None
    #: For a comment reference: the comment's ``w:id``.
    comment_id: str | None = None
    #: For an object: its ProgID (``o:OLEObject/@ProgID``), or None (an ActiveX control's is in its part).
    prog_id: str | None = None
    #: For an image: what the document says of it (REQ-514).
    image: TextImage | None = field(default=None, repr=False)
    #: For an image: the describer's description, with its provenance; None for none (a decorative picture's).
    description: ImageDescription | None = None
    #: For a page marker: the number of the page that begins there, from 2 (REQ-708).
    page_number: int = 0
    #: Bold, italic, struck through, monospaced where its paragraph's style is not (docx4j CR-046 phase 3).
    bold: bool = False
    italic: bool = False
    strike: bool = False
    monospace: bool = False
    #: For text in a link: its target (a ``w:hyperlink``'s, or a HYPERLINK field's).
    link: str | None = None
    #: An equation (REQ-604), and a display one (``m:oMathPara``).
    math: bool = False
    display_math: bool = False
    #: Where the text was read from (not in the JSON).
    source: TextRunSource | None = field(default=None, repr=False)
    #: Where the run's text starts in its paragraph's Office JS ``Paragraph.text``
    #: (``Paragraph.text`` and ``get_text()`` of the content API, CR-011 Q4), where
    #: it contributes there; within the run the two texts agree character for
    #: character. None for a run only the model writes: a note reference or
    #: mark, a list label, a field code, a bound value, text ``Paragraph.text``
    #: does not hold. Text that ``Paragraph.text`` holds and the model leaves
    #: out (hidden text) is the gap between two runs' offsets.
    office_offset: int | None = None

    def same_format(self, other: TextRun) -> bool:
        """Whether two runs' text can be one run: the same formatting, link and equation."""
        return (
            self.bold == other.bold
            and self.italic == other.italic
            and self.strike == other.strike
            and self.monospace == other.monospace
            and self.link == other.link
            and self.math == other.math
            and self.display_math == other.display_math
        )

    def copy_format(self, other: TextRun) -> None:
        """Take another run's formatting, link and equation flags."""
        self.bold = other.bold
        self.italic = other.italic
        self.strike = other.strike
        self.monospace = other.monospace
        self.link = other.link
        self.math = other.math
        self.display_math = other.display_math

    def to_dict(self) -> dict[str, Any]:
        """The run as JSON: its kind, text, revision and flags; no tree object."""
        out: dict[str, Any] = {"kind": self.kind, "text": self.text}
        if self.revision is not None:
            out["revision"] = {"type": self.revision.type, "author": self.revision.author, "date": self.revision.date}
        for name in ("hidden", "placeholder", "bold", "italic", "strike", "monospace", "math", "display_math"):
            if getattr(self, name):
                out[name] = True
        for name in ("note_type", "note_id", "comment_id", "prog_id", "link"):
            value = getattr(self, name)
            if value is not None:
                out[name] = value
        if self.note_type is not None:
            out["note_number"] = self.note_number
        if self.kind == "IMAGE":
            if self.description is not None:
                out["description"] = {"text": self.description.text, "provenance": self.description.provenance}
            if self.image is not None and self.image.decorative:
                out["decorative"] = True
        if self.kind == "PAGE_MARKER":
            out["page_number"] = self.page_number
        return out

    def __repr__(self) -> str:
        return f"TextRun({self.kind}{' ' + repr(self.text) if self.text else ''}{' [' + str(self.revision) + ']' if self.revision else ''})"


@dataclass(frozen=True, slots=True)
class TextField:
    """A field in a block (REQ-505): its instruction, and its cached result, None where it has none. docx4j ``TextField``."""

    instr: str
    result: str | None
    #: True when the result was computed under ``fields: update`` (REQ-506) rather than cached.
    updated: bool = False


@dataclass(slots=True)
class TextNumbering:
    """A paragraph's numbering, as its effective ``w:numPr`` and the definitions give it (REQ-510, REQ-706).

    docx4j ``TextNumbering``. The label itself is the block's ``list_label``,
    under ``list_labels: true``.
    """

    num_id: str
    #: ``w:ilvl``, 0 to 8.
    level: int
    #: The level's ``w:numFmt``, or None.
    format: str | None
    #: The level's ``w:lvlText``, or None.
    level_text: str | None
    #: The level's ``w:start``: 1 where it has none.
    start: int
    #: The count the level shows, as the list labels count it (the 3 of "1.3."), or None.
    value: int | None = None
    #: The list the paragraph belongs to: its ``w:abstractNum``'s id, or None.
    abstract_num_id: str | None = None

    @property
    def bullet(self) -> bool:
        """A bullet level."""
        return self.format == "bullet"

    def to_dict(self) -> dict[str, Any]:
        """REQ-706's names."""
        return {
            "numId": self.num_id,
            "level": self.level,
            "format": self.format,
            "levelText": self.level_text,
            "start": self.start,
            "value": self.value,
            "abstractNumId": self.abstract_num_id,
        }


@dataclass(slots=True)
class TextCell:
    """A table cell of the block model: its blocks, its grid span, and its place in a vertical merge. docx4j ``TextCell``."""

    address: str
    blocks: list[TextBlock] = field(default_factory=list)
    grid_span: int = 1
    merge_continue: bool = False
    merge_restart: bool = False

    def to_dict(self) -> dict[str, Any]:
        out: dict[str, Any] = {"address": self.address, "blocks": [b.to_dict() for b in self.blocks]}
        if self.grid_span != 1:
            out["grid_span"] = self.grid_span
        if self.merge_continue:
            out["merge_continue"] = True
        if self.merge_restart:
            out["merge_restart"] = True
        return out


@dataclass(slots=True)
class TextRow:
    """A table row of the block model: its cells, and whether it is a header row (``w:tblHeader``). docx4j ``TextRow``."""

    address: str
    cells: list[TextCell] = field(default_factory=list)
    header: bool = False

    def to_dict(self) -> dict[str, Any]:
        out: dict[str, Any] = {"address": self.address, "cells": [c.to_dict() for c in self.cells]}
        if self.header:
            out["header"] = True
        return out


@dataclass(slots=True, eq=False)
class TextBlock:
    """A block of the block model (REQ-706): a paragraph, or a table whose cells hold blocks. docx4j ``TextBlock``.

    A paragraph's text is its runs' in the view; what the plain text adds
    (note numbers, a comment's CriticMarkup) is carried as runs and fields,
    not text (REQ-707).

    Addresses are the ordinal form of the unresolved document (REQ-405,
    REQ-519; core-ts CR-002 section 3.3's grammar): ``body/3``,
    ``body/4/1/0/2`` (a table's row 1, cell 0, block 2), ``header:rId3/0``,
    ``footnote:2/0``, ``comment:5/0``; a text box is addressed from its
    anchoring paragraph, ``body/3/textbox:0/1`` (core-ts CR-012's D4, WD
    0.27). A paragraph joined to the next in a view (a deleted mark accepted)
    takes the first paragraph's address and ``w14:paraId``.
    """

    kind: BlockKind
    story: Story
    address: str
    para_id: str | None = None
    #: The style id (``w:pStyle``); a joined paragraph's is the last one's.
    style: str | None = None
    runs: list[TextRun] = field(default_factory=list)
    fields: list[TextField] = field(default_factory=list)
    rows: list[TextRow] = field(default_factory=list)
    #: A hidden paragraph included under ``hidden: true``: its runs and mark are hidden (REQ-507).
    hidden: bool = False
    #: How many revisions (run-level revisions and paragraph marks) the block's paragraphs hold, in any view.
    revisions: int = 0
    indent_left: int | None = None
    indent_hanging: int | None = None
    indent_first_line: int | None = None
    #: The effective alignment (``w:jc``'s value).
    alignment: str | None = None
    #: Every text run is in a monospaced font, and at least one is.
    monospace: bool = False
    #: The list label from the numbering model, under ``list_labels: true`` (REQ-510).
    list_label: str | None = None
    #: The effective outline level, 0 to 8; None for body text.
    outline_level: int | None = None
    numbering: TextNumbering | None = None
    #: The effective ``w:contextualSpacing``; None where neither the paragraph nor its styles set it.
    contextual_spacing: bool | None = None
    #: The paragraph has a bottom border of its own.
    bottom_border: bool = False
    #: For a note's block: whose (``"footnote"``, ``"endnote"``), its number in reference order and its ``w:id``.
    note_type: str | None = None
    note_number: int = 0
    note_id: str | None = None
    #: The section the block ends or is in (the main story only).
    section: int = 0
    #: For a table: always written in REQ-605's simple form, whatever its shape
    #: (an embedded workbook's sheet, docx4j's; none is read here, CR-011 Q7).
    simple_form: bool = False
    #: What the block was read from (not in the JSON): a paragraph's ``w:p``
    #: (a joined block's first), a table's ``w:tbl``.
    source: Any = field(default=None, repr=False)

    @property
    def placeholder(self) -> bool:
        """The block holds placeholder text (REQ-508)."""
        return any(r.placeholder for r in self.runs)

    @property
    def textbox(self) -> bool:
        """A block in a text box (REQ-518)."""
        return self.story == "TEXT_BOX"

    @property
    def decorative(self) -> bool:
        """A paragraph whose only content is one or more pictures marked decorative (REQ-514, REQ-706)."""
        found = False
        for r in self.runs:
            if r.kind == "IMAGE":
                if r.image is None or not r.image.decorative:
                    return False
                found = True
            elif r.kind == "TEXT" and r.text:
                return False
        return found

    @property
    def text(self) -> str:
        """The block's text in the view: its text runs and field codes; a table's rows a line apart, cells a tab apart.

        No list label, note number or comment.
        """
        if self.kind == "TABLE":
            return "\n".join(
                "\t".join(" ".join(b.text for b in cell.blocks) for cell in row.cells) for row in self.rows
            )
        return "".join(r.text for r in self.runs if r.kind in ("TEXT", "FIELD_CODE"))

    def to_dict(self) -> dict[str, Any]:
        """The block model's JSON (REQ-706): its fields, its text and its flags; never a tree object."""
        out: dict[str, Any] = {"kind": self.kind, "story": self.story, "address": self.address}
        if self.para_id is not None:
            out["paraId"] = self.para_id
        out["style"] = self.style
        out["text"] = self.text
        out["placeholder"] = self.placeholder
        out["textbox"] = self.textbox
        out["decorative"] = self.decorative
        out["hidden"] = self.hidden
        out["revisions"] = self.revisions
        if self.indent_left is not None or self.indent_hanging is not None or self.indent_first_line is not None:
            out["indent"] = {"left": self.indent_left, "hanging": self.indent_hanging, "firstLine": self.indent_first_line}
        if self.alignment is not None:
            out["alignment"] = self.alignment
        out["monospace"] = self.monospace
        if self.outline_level is not None:
            out["outline_level"] = self.outline_level
        if self.numbering is not None:
            out["numbering"] = self.numbering.to_dict()
        if self.list_label is not None:
            out["list_label"] = self.list_label
        if self.note_type is not None:
            out["note"] = {"type": self.note_type, "number": self.note_number, "id": self.note_id}
        if self.kind == "PARAGRAPH":
            out["runs"] = [r.to_dict() for r in self.runs]
            if self.fields:
                out["fields"] = [{"instr": f.instr, "result": f.result} for f in self.fields]
        else:
            out["rows"] = [r.to_dict() for r in self.rows]
        return out

    def __repr__(self) -> str:
        return f"TextBlock({self.address} {self.story} {self.text!r})"


@dataclass(slots=True, eq=False)
class TextComment:
    """A comment (REQ-504): its author, date and blocks, the comment it replies to, whether it is resolved. docx4j ``TextComment``."""

    id: str
    author: str | None = None
    initials: str | None = None
    #: The date as the package wrote it.
    date: str | None = None
    #: The ``w:id`` of the comment this one replies to (``w15:paraIdParent``).
    parent_id: str | None = None
    #: Resolved (``w15:done``).
    resolved: bool = False
    blocks: list[TextBlock] = field(default_factory=list)
    #: The replies, in the comments part's order.
    replies: list[TextComment] = field(default_factory=list)

    @property
    def text(self) -> str:
        """The comment's text: its blocks a line apart."""
        return "\n".join(b.text for b in self.blocks)

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "author": self.author,
            "initials": self.initials,
            "date": self.date,
            "parent_id": self.parent_id,
            "resolved": self.resolved,
            "blocks": [b.to_dict() for b in self.blocks],
            "replies": [r.id for r in self.replies],
        }


@dataclass(frozen=True, slots=True)
class TextWarning:
    """Where a result is not the whole document (WD section 8, REQ-801). docx4j ``TextWarning``.

    A kind, the address of the block it is in where there is one, and its
    selector: for ``mce.branches_differ`` the branch selected, for
    ``altchunk.not_expanded`` the chunk's content type, for ``content.not_read``
    the part's name. Never the document's text (REQ-802).
    """

    #: Both branches of an ``mc:AlternateContent`` hold text, and the texts differ (REQ-307).
    MCE_BRANCHES_DIFFER = "mce.branches_differ"
    #: A field with no cached result (REQ-505).
    FIELD_NO_RESULT = "field.no_result"
    #: A data-bound control whose surface differs from its bound value (REQ-509).
    SDT_BINDING_DIFFERS = "sdt.binding_differs"
    #: An ``altChunk`` not expanded (REQ-513).
    ALTCHUNK_NOT_EXPANDED = "altchunk.not_expanded"
    #: A part that cannot be read: a text box that would lose content as typed, a
    #: chart, a SmartArt data part, an embedded package or a part inside it; it
    #: contributes nothing (REQ-511, REQ-512, REQ-801).
    CONTENT_NOT_READ = "content.not_read"
    #: An image the describer failed to describe: it has no description (REQ-514, REQ-801).
    IMAGE_NOT_DESCRIBED = "image.not_described"

    kind: str
    address: str | None = None
    #: For ``mce.branches_differ``: ``mc:Choice Requires="wps"`` or ``mc:Fallback``.
    branch: str | None = None
    #: For ``altchunk.not_expanded``: the chunk's content type as the package declares it.
    content_type: str | None = None
    #: For ``content.not_read``: the part's name (for a part inside an embedded
    #: package, the embedded package part's); for ``image.not_described``, the
    #: image part's.
    part_name: str | None = None

    @property
    def selector(self) -> str | None:
        """REQ-801's selector: the branch, the content type or the part's name, whichever the kind has."""
        return self.branch or self.content_type or self.part_name

    def to_dict(self) -> dict[str, Any]:
        return {"kind": self.kind, "address": self.address, "selector": self.selector}

    def __str__(self) -> str:
        return self.kind + (f" at {self.address}" if self.address else "") + (f" ({self.selector})" if self.selector else "")


@dataclass(slots=True, eq=False)
class TextResult:
    """A text in the block model (WD sections 7.4 and 8): blocks, comments, warnings and metadata. docx4j ``TextResult``.

    :func:`~docx4j_py.model.text.plain_text` writes plain text from it
    (REQ-202): no reader of the model reads the tree again.
    """

    #: The metadata's key for the settings' ``w:updateFields`` (REQ-506, REQ-801).
    FIELDS_UPDATE_ON_OPEN = "fields_update_on_open"
    #: The metadata's key for the core properties (REQ-515): each property's element name to its value.
    CORE = "core"
    #: The metadata's key for the application properties (REQ-515), the statistics left out.
    APP = "app"
    #: The metadata's key for the custom properties (REQ-515): each property's name to its typed value.
    CUSTOM = "custom"

    #: The blocks in reading order (REQ-501, REQ-502, REQ-518, REQ-606). Comments are not among them.
    blocks: list[TextBlock]
    #: The comments the view includes, by ``w:id``, in the comments part's order (REQ-504).
    comments: dict[str, TextComment]
    #: The warnings raised (REQ-801), each once.
    warnings: list[TextWarning]
    #: Whatever the options (REQ-801, REQ-515): ``fields_update_on_open``, and
    #: the document properties under :attr:`CORE`, :attr:`APP` and :attr:`CUSTOM`,
    #: each a map, where the package has that part. Property values are document
    #: text, in the text stream only under ``metadata: true``.
    metadata: dict[str, Any]
    #: The options the result was made with.
    options: TextOptions

    def to_dict(self) -> dict[str, Any]:
        """The block model's JSON."""
        return {
            "blocks": [b.to_dict() for b in self.blocks],
            "comments": {k: c.to_dict() for k, c in self.comments.items()},
            "warnings": [w.to_dict() for w in self.warnings],
            "metadata": _json(self.metadata),
            "options": self.options.to_dict(),
        }


def _json(value: Any) -> Any:
    """A metadata value as JSON holds it: a map's values in turn, a decimal as its text."""
    if isinstance(value, dict):
        return {k: _json(v) for k, v in value.items()}
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)
