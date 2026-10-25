"""``ContentControl``: Office JS ``Word.ContentControl`` over ``w:sdt``.

CR-003 section 3.2, Phase C --- the **reads** and ``delete`` --- and Phase E,
which added the ``w:sdtPr`` properties (``appearance``, ``color``,
``cannot_delete``, ``cannot_edit``, ``remove_when_edited``,
``placeholder_text``), the :class:`~docx4j_py.model.customxml.XmlMapping` and
the seven typed sub-views, all in :mod:`docx4j_py.model.customxml` and reached
from here.

``w:sdt`` is one element name with four types, and the model gives each its own
class (``SdtBlock``, ``SdtRun``, ``CTSdtRow``, ``CTSdtCell``). :attr:`form`
reports which, falling back to **what the content holds** for the block class,
because a run-level control parsed in a ``w:body``'s scope comes back as
``SdtBlock`` (CR-001 section 10.4).

The rules of CR-003 section 4 that land here:

* :attr:`type` is ``"RichText"`` when ``w:sdtPr`` names no kind and
  ``"PlainText"`` for ``w:text`` --- Phase A's
  :func:`~docx4j_py.wml.sdt.sdt_kind_of`, unchanged;
* :meth:`get_range` is **exact** for a run-level control and the first
  paragraph's range for the block, row and cell forms;
* a row- or cell-level control **refuses** ``insert_paragraph`` at ``"Start"``
  or ``"End"`` and ``insert_text(..., location="Replace")``, naming what it
  holds.
"""

from __future__ import annotations

import dataclasses
from collections.abc import Collection
from typing import TYPE_CHECKING, Any

from docx4j_py.child import ChildList, is_alternate_content, link_parents
from docx4j_py.model.content.addresses import ordinal_of
from docx4j_py.model.content.enums import RangeLocation, TextLocation
from docx4j_py.model.content.errors import ContentError, InvalidTargetError, SpanError
from docx4j_py.model.content.reports import recording
from docx4j_py.model.content.table import Table
from docx4j_py.model.content.text_model import (
    PARAGRAPH_BREAKS,
    W_P,
    W_SDT,
    W_TBL,
    W_TC,
    W_TR,
    block_children_of,
    drop_unwritable,
    explode_mapped,
    for_writer,
    text_of_view,
    text_run,
)
from docx4j_py.model.values import effective, is_on, on_off
from docx4j_py.traversal import element_name, run_items_of, text_of
from docx4j_py.wml import (
    CTSdtCell,
    CTSdtRow,
    P,
    SdtBlock,
    SdtContentBlock,
    SdtRun,
    el,
    to_xml,
)
from docx4j_py.wml.sdt import W14_NS, W15_NS, W_NS, sdt_kind_of, sdt_property

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.body import Body
    from docx4j_py.model.content.paragraph import Paragraph
    from docx4j_py.model.content.range import Range

__all__ = [
    "SDT_PR_ORDER",
    "ContentControl",
    "ContentControlForm",
    "block_control_of",
    "control_view_of",
    "controls_in",
    "controls_in_paragraph",
    "controls_of_types",
    "holds_one_paragraph",
    "is_bound_block_control",
    "paragraph_lines_of",
    "parent_control_of",
    "plain_text_lines_of",
]

#: Which of the four ``w:sdt`` forms a control is (CR-003 section 3.2).
ContentControlForm = str

#: The four container classes, so a walk can recognise a control of any form
#: without asking its element name (all four are ``w:sdt``).
_SDT_CLASSES: tuple[type, ...] = (SdtBlock, SdtRun, CTSdtRow, CTSdtCell)

#: How far up the parent chain the paragraph of a run-level control is looked for.
_MAX_DEPTH = 64

#: How much of the text a ``repr`` shows.
_PREVIEW = 30

#: ``w15:appearance`` values, to and from Office JS's ``ContentControlAppearance``.
_APPEARANCE_IN = {"boundingBox": "BoundingBox", "tags": "Tags", "hidden": "Hidden"}
_APPEARANCE_OUT = {name: stored for stored, name in _APPEARANCE_IN.items()}

#: The order Word writes the ``w:sdtPr`` children in. The schema is a choice
#: with no order at all, but Word is happier reading its own, so
#: :meth:`ContentControl.put_property` inserts by this table and anything not in
#: it goes just before the element that types the control.
SDT_PR_ORDER: tuple[str, ...] = (
    f"{{{W_NS}}}rPr",
    f"{{{W_NS}}}alias",
    f"{{{W_NS}}}tag",
    f"{{{W_NS}}}id",
    f"{{{W_NS}}}lock",
    f"{{{W_NS}}}placeholder",
    f"{{{W_NS}}}temporary",
    f"{{{W_NS}}}showingPlcHdr",
    f"{{{W_NS}}}dataBinding",
    f"{{{W15_NS}}}dataBinding",
    f"{{{W15_NS}}}appearance",
    f"{{{W15_NS}}}color",
)

#: The ``w:sdtPr`` children that say what **kind** of control it is; they come
#: last, which is where Word writes them.
_KIND_NAMES = frozenset(
    {
        f"{{{W_NS}}}{name}"
        for name in (
            "text",
            "richText",
            "picture",
            "docPartObj",
            "docPartList",
            "comboBox",
            "dropDownList",
            "date",
            "group",
            "citation",
            "bibliography",
            "equation",
        )
    }
    | {
        f"{{{W14_NS}}}checkbox",
        f"{{{W14_NS}}}entityPicker",
        f"{{{W15_NS}}}repeatingSection",
        f"{{{W15_NS}}}repeatingSectionItem",
    }
)


def _position_for(items: list, name: str) -> int:
    """Where a new ``w:sdtPr`` child goes: by :data:`SDT_PR_ORDER`, kinds last."""
    if name in _KIND_NAMES:
        return len(items)
    rank = SDT_PR_ORDER.index(name) if name in SDT_PR_ORDER else len(SDT_PR_ORDER)
    for index, item in enumerate(items):
        other = element_name(item) or ""
        if other in _KIND_NAMES:
            return index
        other_rank = SDT_PR_ORDER.index(other) if other in SDT_PR_ORDER else len(SDT_PR_ORDER)
        if other_rank > rank:
            return index
    return len(items)


#: The character style Word gives a placeholder's runs (docx4j's, and :mod:`~docx4j_py.model.customxml.bindings`').
PLACEHOLDER_STYLE = "PlaceholderText"


def _placeholder_run(text: str, run_properties: Any) -> Any:
    """docx4j's placeholder run: the prompt, styled ``PlaceholderText``."""
    from docx4j_py.child import deep_copy
    from docx4j_py.wml import R
    from docx4j_py.wml import t as text_element

    properties = deep_copy(run_properties) if run_properties is not None else el.rPr()
    properties.r_style = el.rStyle(val=PLACEHOLDER_STYLE)
    return R(r_pr=properties, content=ChildList([text_element(text)]))


#: Word's filler for a control emptied in the bare form: five spaces under ``w:showingPlcHdr`` (CR-019 section 11, D10).
FILLER = "     "

#: The inner start markers a control tracked as inserted or deleted ends its content with (check 30).
_TRACKED_HOLDER_STARTS = frozenset({f"{{{W_NS}}}customXmlInsRangeStart", f"{{{W_NS}}}customXmlDelRangeStart"})

#: A tracked insertion's inner start marker, before which text into an empty run control goes.
_CUSTOM_XML_INS_RANGE_START = f"{{{W_NS}}}customXmlInsRangeStart"


@dataclasses.dataclass(slots=True)
class _Taken:
    """A placeholder's runs taken out of a control, and the list and object that held them (CR-019)."""

    #: The runs, in document order.
    runs: list
    #: The live list they came out of: a run-level control's content, a block control's first paragraph's.
    holder: list
    #: The object whose list `holder` is, for a parent pointer.
    owner: Any


class ContentControl:
    """A subset of Office JS ``Word.ContentControl`` over a ``w:sdt``."""

    __slots__ = ("container", "element", "parent_body")

    def __init__(self, element: Any, container: list, parent_body: Body) -> None:
        """Build the view over a ``w:sdt``, the list holding it and its body."""
        #: The ``w:sdt``: a ``SdtBlock``, ``SdtRun``, ``CTSdtRow`` or ``CTSdtCell``.
        self.element = element
        #: The live list holding it.
        self.container = container
        #: The nearest :class:`~docx4j_py.model.content.Body`; a read-only
        #: one in a kept ``mc:AlternateContent`` branch (CR-012 Phase A).
        self.parent_body = parent_body.body_for(element) if parent_body is not None else parent_body

    # -- identity ----------------------------------------------------------

    def __eq__(self, other: object) -> bool:
        """Two views of the same ``w:sdt`` are equal (CR-003 section 3.1)."""
        return isinstance(other, ContentControl) and other.element is self.element

    def __hash__(self) -> int:
        """Hashes by the element's identity."""
        return hash(id(self.element))

    def __str__(self) -> str:
        """The control's text."""
        return self.text

    def __repr__(self) -> str:
        """``<ContentControl body/5 RichText 'customer' 'Acme Ltd'>``."""
        text = self.text.replace("\n", " ")
        preview = text[:_PREVIEW] + "…" if len(text) > _PREVIEW else text
        where = self.ordinal
        tag = self.tag or self.title
        label = f" {tag!r}" if tag else ""
        return f"<ContentControl{' ' + where if where else ''} {self.type}{label} {preview!r}>"

    @property
    def name(self) -> str:
        """``"w:sdt"``, as :class:`~docx4j_py.model.content.body.Block` reports it."""
        return "w:sdt"

    @property
    def index(self) -> int:
        """This control's position in its container, or -1 if it is gone."""
        for index, item in enumerate(self.container):
            if item is self.element:
                return index
        return -1

    # -- the tree ----------------------------------------------------------

    @property
    def sdt_pr(self) -> Any:
        """The ``w:sdtPr``, or None. Extension."""
        return getattr(self.element, "sdt_pr", None)

    @property
    def content(self) -> list:
        """The control's ``w:sdtContent`` list; empty when it has none."""
        items = block_children_of(self.element)
        if items is not None:
            return items
        inner = getattr(self.element, "sdt_content", None)
        return run_items_of(inner) or [] if inner is not None else []

    @property
    def form(self) -> str:
        """``"block"``, ``"run"``, ``"row"`` or ``"cell"``.

        The class decides for three of the four; for ``SdtBlock`` the
        **content** decides, because a run-level control parsed in a
        ``w:body``'s scope is typed ``SdtBlock`` (CR-001 section 10.4).
        """
        if isinstance(self.element, CTSdtRow):
            return "row"
        if isinstance(self.element, CTSdtCell):
            return "cell"
        if isinstance(self.element, SdtRun):
            return "run"
        items = self.content
        for _ in range(_MAX_DEPTH):
            if not items:
                break
            # a marker says nothing of the level: Word's save of core-ts check 52 case 82 opens a block
            # control's content with a w:commentRangeStart (CR-013 Phase B); the first item that is no
            # marker decides, the first item where there is none
            item = next((item for item in items if element_name(item) not in _MARKER_NAMES), items[0])
            if is_alternate_content(item):
                # a kept block-level mc:AlternateContent: what Word's branch holds says
                # (CR-012 Phase A2; it read "run" before)
                items = block_children_of(item) or []
                continue
            name = element_name(item)
            if name == W_TR:
                return "row"
            if name == W_TC:
                return "cell"
            return "block" if _is_block(name) else "run"
        return "block"

    @property
    def type(self) -> str:
        """The Office JS ``Word.ContentControlType`` (``"RichText"`` for an untyped one)."""
        return sdt_kind_of(self.sdt_pr)

    @property
    def tag(self) -> str:
        """``w:tag``, the machine-readable name; ``""`` when there is none."""
        return self._property("tag")

    @property
    def title(self) -> str:
        """``w:alias``, which Word's dialog calls the title; ``""`` when there is none."""
        return self._property("alias")

    @property
    def id(self) -> int:
        """``w:id``, the number Word gives a control; 0 when it has none."""
        value = getattr(sdt_property(self.sdt_pr, "id", W_NS), "val", None)
        return int(value) if isinstance(value, int) else 0

    @property
    def text(self) -> str:
        """The control's text as Office JS's ``contentControl.text`` gives it.

        CR-013 Phase A (core-ts check 51): a block control's is its content's
        ``Body.text``, its paragraphs a mark ``\\r`` apart; a run-level
        control's is the span of its paragraph's ``Paragraph.text`` it covers,
        in Office JS's characters, hidden text left out. A row- or cell-level
        control's is its content's ``Body.text`` too, in Word's table form,
        and a control whose content ends with a row or a table leaves the
        row's ``\\r\\n`` out (core-ts check 53; CR-013 Phase B): a row
        control's text is its cells a ``\\t`` apart.

        A control **showing its placeholder** (``w:showingPlcHdr``) reads
        ``""``, as Office JS reads it (CR-016 Phase B; core-ts check 53 cases 34
        and 35, check 58 case 11), and so does a control inside one; its
        prompt is :attr:`placeholder_text`.

        Office JS's text, not the text model's (CR-011 Q4; WD REQ-202's exemption for a reader of an
        emulated API): :func:`docx4j_py.model.text.extract` is the text model, and each of its runs
        records its ``office_offset`` in this text. It reads Word's branch of a kept
        ``mc:AlternateContent`` (Q4 revised, CR-011 section 15).

        A bound control's write-back into its custom XML reads :attr:`docx4j_text`.
        """
        if self.form == "run":
            span = self._run_span()
            if span is not None:
                paragraph, start, end = span
                return paragraph.text[start:end]
            inner = getattr(self.element, "sdt_content", None)
            return text_of_view(inner) if inner is not None else ""
        # a row control's cells a \t apart (core-ts check 53: ROW15A\tROW15B), a cell control's paragraphs
        # a \r apart; a control whose content ends with a row (a row control, or a table last) leaves the
        # row's end out (check 53; CR-013 Phase B)
        return self.body().text.removesuffix("\r\n")

    @property
    def docx4j_text(self) -> str:
        """The control's text in docx4j's characters, as 0.1.0's :attr:`text` read it. Extension.

        CR-013 section 4.1: what a bound control writes back into its custom XML
        node (:func:`~docx4j_py.model.customxml.bindings.update_from_control`),
        so that the node holds docx4j's text, as core-ts's D4 has it --- a break
        ``\\n``, the hyphens U+2011 and U+00AD, a block control's paragraphs a
        ``\\n`` apart --- and never one of Office JS's characters, U+000B
        among them, which XML cannot hold. A run-level control's is
        :func:`~docx4j_py.traversal.text_of`'s.
        """
        if self.form == "run":
            inner = getattr(self.element, "sdt_content", None)
            return text_of(inner) if inner is not None else ""
        return "\n".join(
            text_of_view(paragraph.element, characters="docx4j")
            for paragraph in self.body().iter_paragraphs()
        )

    # -- addresses ---------------------------------------------------------

    @property
    def address(self) -> str:
        """The ordinal address (``"body/5"``); CR-003 section 3.4."""
        return self.ordinal

    @property
    def ordinal(self) -> str:
        """The ordinal address, ``""`` when the control is no longer in its body."""
        return ordinal_of(self.parent_body, self.element) or ""

    # -- what it holds -----------------------------------------------------

    def body(self) -> Body:
        """A :class:`Body` over the control's content. Extension.

        The four ``w:sdtContent`` levels are invisible to an address (CR-003
        section 4), so this body carries the **control's** prefix and its
        children are numbered under it: ``body/5/0`` is the control's first
        block.
        """
        return self.parent_body.sub(self.element, self.address or self.parent_body.prefix)

    @property
    def paragraphs(self) -> list[Paragraph]:
        """Every paragraph in the control; for a run-level one, the paragraph it is in."""
        if self.form == "run":
            paragraph = self.parent_paragraph()
            return [paragraph] if paragraph is not None else []
        return self.body().paragraphs

    @property
    def tables(self) -> list[Table]:
        """The tables in the control; for a row or cell control, the table it is part of."""
        form = self.form
        if form == "block":
            return self.body().tables
        if form in ("row", "cell"):
            table = self._ancestor_table()
            return [table] if table is not None else []
        return []

    @property
    def content_controls(self) -> list[ContentControl]:
        """The controls nested in this one, in document order."""
        out: list[ContentControl] = []
        if self.form == "run":
            _collect_run_controls(self.element, self.parent_body, out)
        else:
            _collect(self.content, self.parent_body, out)
        return out

    def get_content_controls(self, *, types: Collection[str] | None = None) -> list[ContentControl]:
        """Office JS ``getContentControls`` (WordApi 1.5): :attr:`content_controls`, those of `types` where given.

        `types` are ``Word.ContentControlType`` values, as :attr:`type` reports
        them (``"PlainText"``, ``"RichText"``, ``"CheckBox"``, ...). Word's
        ``contentControls`` lists rich-text controls only; this package's
        :attr:`content_controls` lists every kind, a recorded departure (CR-013
        section 11, G1; core-ts check 55b), so this is how code that asks for
        plain-text controls finds them, as Word's does.
        """
        return controls_of_types(self.content_controls, types)

    @property
    def inline_pictures(self) -> list[Any]:
        """Every inline picture in the control, in document order."""
        return [picture for paragraph in self.paragraphs for picture in paragraph.inline_pictures]

    def parent_paragraph(self) -> Paragraph | None:
        """The paragraph a run-level control is in, or None. Extension."""
        current: Any = self.element
        for _ in range(_MAX_DEPTH):
            if current is None:
                return None
            if isinstance(current, P):
                return self.parent_body.paragraph_for(current)
            current = getattr(current, "parent", None)
        return None

    # -- ranges ------------------------------------------------------------

    def get_range(self, location: RangeLocation = "Whole") -> Range:
        """The control's range.

        **Exact** for a run-level control --- the offsets its runs cover in the
        paragraph. For the block, row and cell forms one range over its
        paragraphs, from the first one's start to the last one's end (CR-013
        Phase B; core-ts's at ``ef57729``): ``"Whole"`` ending with the last
        one's mark, as Office JS's does (core-ts check 51:
        ``CCP1\\rCCP2\\rCCP3\\r``), ``"Content"`` without it; a row-level
        control's, or one whose content ends with a table, ``"Content"``
        ending with the last cell's text and ``"Whole"`` with ``\\t\\t``, as
        Word 365 gives them (core-ts check 53). ``"Start"`` is the first
        paragraph's start and ``"End"`` the last one's end; ``"After"`` is
        after the range's end.

        Its own ``"Content"`` and ``"Whole"`` ranges, of a block- or run-level
        control, know the control (CR-019 section 4.4, core-ts's D4): a write
        through one asks the control first (:meth:`_write_through`,
        :meth:`_delete_through`), since a run-level control showing its
        placeholder, or holding nothing, has no width of its own. A range made from one (``expand_to``,
        ``get_range``) is an ordinary range.

        Raises:
            ContentError: a run-level control that is in no paragraph, or a
                control that holds no paragraph at all.
        """
        found = self._range_of(location)
        if location in ("Content", "Whole") and self.form in ("run", "block"):
            # its own 'Content' and 'Whole' ranges know the control, for a write into one with no width
            # (CR-019 section 4.4, core-ts's D4 and Range.ofControl)
            found.of_control = (self, location)
        return found

    def _range_of(self, location: RangeLocation) -> Range:
        """:meth:`get_range`'s range, before it is told its control."""
        from docx4j_py.model.content.range import Range

        if self.form == "run":
            span = self._run_span()
            if span is None:
                raise ContentError(
                    "this run-level content control is not in a paragraph",
                    code="control.no_paragraph",
                    hint="a run-level w:sdt gets its range from the w:p it sits in",
                )
            paragraph, start, end = span
            if location == "Start":
                return Range(paragraph, start, start)
            if location == "End":
                return Range(paragraph, end, end)
            if location == "After":
                return Range(paragraph, start, end).get_range("After")
            return Range(paragraph, start, end)
        paragraphs = self.paragraphs
        if not paragraphs:
            raise ContentError(
                f"this {self.form}-level content control holds no paragraph",
                code="control.no_paragraph",
                hint="insert one with insert_paragraph(), or read its text",
            )
        first, last = paragraphs[0], paragraphs[-1]
        if location == "Start":
            return first.get_range("Start")
        if location == "End":
            return last.get_range("End")
        length = len(last.text)
        if self.form == "row" or _ends_with_table(self.content):
            if location == "Whole":
                return Range(first, 0, length, end_paragraph=last, mark=True, end_mark="\t\t", whole=True)
            whole = Range(first, 0, length, end_paragraph=last, whole=True)
            return whole.get_range("After") if location == "After" else whole
        if len(paragraphs) == 1:
            return first.get_range(location)
        whole = Range(first, 0, length, end_paragraph=last, mark=location in ("Whole", "After"), whole=True)
        return whole.get_range("After") if location == "After" else whole

    def search(self, text: str, **options: Any) -> list[Range]:
        """Every match of `text` inside the control, as ranges.

        Matches are found in Office JS's text, not the text model's (CR-011 Q4; WD REQ-202's exemption for a reader of an
        emulated API): :func:`docx4j_py.model.text.extract` is the text model, and each of its runs
        records its ``office_offset`` in this text, which reads Word's branch of a run- or
        paragraph-level ``mc:AlternateContent`` (Q4 revised, CR-011 section 15).
        """
        if self.form != "run":
            return self.body().search(text, **options)
        span = self._run_span()
        if span is None:
            return []
        paragraph, start, end = span
        from docx4j_py.model.content.range import Range

        return Range(paragraph, start, end).search(text, **options)

    # -- editing -----------------------------------------------------------

    def insert_text(self, text: str, *, location: TextLocation = "End") -> Range:
        """Text at the start, the end, or in place of the control's text.

        **A control showing its placeholder** (``w:showingPlcHdr``), block- or
        run-level, takes the text in the placeholder's place whatever
        `location` says, as Word's does (core-ts check 58 cases 20 to 22, check
        63's run 2 cases 1, 2, 13 and 14, check 64 cases 1 to 3 and 10 to 12,
        check 67; CR-016 Phase B, CR-019 sections 4.2 and 11): the write is
        :meth:`_write_placeholder`'s, which clears ``w:showingPlcHdr``, writes
        the text plain in the placeholder's runs' place (in the bare form in
        the placeholder run's own look, its style left off), in one paragraph
        tracked or not, only the insertion tracked, and writes a bound
        control's node, in either form. A refusal below leaves the control
        showing its placeholder.

        ``\\n``, ``\\r`` and ``\\r\\n`` make a paragraph each in a block,
        row or cell control, as every writer through ``Paragraph.splice`` makes
        them (CR-013 section 11, E1), except in a ``w:text`` control, bound or not,
        and a bound block one Word reads as plain text: there they are a
        ``w:br`` in a multi-line control and nothing in a single-line one, a line,
        page or column break dropped with them (F3; docx4j-core-ts CR-002
        sections 60 and 61 at ``6285dfa``; :func:`plain_text_lines_of`). In a
        run-level control a paragraph break is refused (core-ts's F4), and a
        single-line one drops a line break, its node following it (F3, check
        55b case 8). U+0002 and the other characters XML 1.0 excludes are
        dropped (section 11, E3). A bound control's node takes the text, block
        or run (CR-019 section 4.7, core-ts's D7).

        Raises:
            InvalidTargetError: a row- or cell-level control asked to
                ``"Replace"``; the message names what it holds.
            SpanError: ``range.splits_control``, a run-level control given a
                ``\\n`` or a ``\\r``; nothing has changed.
        """
        form = self.form
        if form in ("row", "cell") and location == "Replace":
            raise self._holds(
                f"a {form}-level content control", "replace the text of its cells instead"
            )
        text = drop_unwritable(text)
        if self.shows_placeholder_content:
            # the text in the placeholder's place, wherever it is put (CR-019 section 4.2)
            return self._write_placeholder(text)
        return self._insert_text(text, location, form)

    def _refuse_run_text(self, text: str) -> None:
        """A run-level control's refusals of :meth:`insert_text`, before anything changes."""
        if PARAGRAPH_BREAKS.search(text):
            raise SpanError(
                "insert_text refused: a paragraph break (\\n or \\r) written into a run-level content "
                "control would cut the control in two (CR-013 section 11, E1; docx4j-core-ts F4)",
                code="range.splits_control",
                hint="write a line break (U+000B) instead, or write the paragraph outside the control",
            )
        if self.parent_paragraph() is None:
            raise ContentError(
                "this run-level content control is not in a paragraph",
                code="control.no_paragraph",
                hint="a run-level w:sdt is edited through the w:p it sits in",
            )

    def _take_property(self, local_name: str, namespace: str) -> tuple[int, Any] | None:
        """Remove every ``w:sdtPr`` child of a name; the first one's ``(index, element)``, to put it back, or None."""
        items = self.properties()
        for index, item in enumerate(items):
            name = element_name(item) or ""
            uri, _, local = name.rpartition("}")
            if local == local_name and uri.lstrip("{") == namespace:
                del items[index]
                self.remove_property(local_name, namespace)
                return index, item
        return None

    def _content_runs(self) -> list[Any]:
        """The runs of the control's content, a block control's paragraphs' too (core-ts ``contentRuns``)."""
        from docx4j_py.model.content.text_model import runs_of

        if self.form == "run":
            inner = getattr(self.element, "sdt_content", None)
            return runs_of(inner) if inner is not None else []
        return [run for paragraph in self.paragraphs for run in paragraph.runs]

    def _insert_text(self, text: str, location: TextLocation, form: str) -> Range:
        """:meth:`insert_text` once the placeholder rule has run."""
        if form != "run":
            out = self.body().insert_text(text, location=location)
            # every text write into a bound control writes its node (CR-019 section 4.7, core-ts's D7),
            # a block control's too, as core-ts's insertText does
            self.write_through()
            return out
        if PARAGRAPH_BREAKS.search(text):
            # Word, given one, moved a bound plain-text control to block level (core-ts
            # check 55b, cases 3 and 4); core-ts's F4 refuses it, and so does this
            raise SpanError(
                "insert_text refused: a paragraph break (\\n or \\r) written into a run-level content "
                "control would cut the control in two (CR-013 section 11, E1; docx4j-core-ts F4)",
                code="range.splits_control",
                hint="write a line break (U+000B) instead, or write the paragraph outside the control",
            )
        if plain_text_lines_of(self) == "single":
            # a single-line plain-text control drops a line break, as Word does (F3; check 55b case 8)
            text = for_writer(text, "single-line")
        span = self._run_span()
        if span is None:
            raise ContentError(
                "this run-level content control is not in a paragraph",
                code="control.no_paragraph",
                hint="a run-level w:sdt is edited through the w:p it sits in",
            )
        paragraph, start, end = span
        with recording(self.parent_body, "insert_text") as change:
            change.text(before=self.text)
            if location in ("Start", "End"):
                # ``splice`` at a boundary extends the *neighbouring* run, which
                # at the control's own boundary is a run outside it. Text asked
                # for at the start or the end of a control belongs inside it, as
                # Word puts it, so the control's own items are edited directly.
                out = self._extend(paragraph, text, at_start=location == "Start")
            else:
                out = paragraph.splice(start, end, text)
            change.touched(paragraph)
            change.text(after=text)
            # CR-003 section 4, and section 16.6: the write-through sits
            # **outside** the tracked branch. Word refreshes a bound control
            # from its data on open, and a ``w:del`` in the control does not
            # change the data, so the node takes the new text whatever the
            # change-tracking mode is.
            self.write_through()
            return out

    def _extend(self, paragraph: Paragraph, text: str, *, at_start: bool) -> Range:
        """Put text at the start or the end of a run control's **own** content."""
        from docx4j_py.model.content.range import Range
        from docx4j_py.model.content.text_model import set_text

        inside = [segment for segment in paragraph.segments() if self._inside(segment.run)]
        edge = (inside[0] if at_start else inside[-1]) if inside else None
        if edge is not None and edge.branch is not None:
            # the control's text begins or ends in a kept branch's (CR-011 Phase F): the
            # text goes beside the element, inside the control, into no branch
            item, owner, index = edge.site()
            run = text_run(text)
            owner.insert(index if at_start else index + 1, run)
            link_parents(run)
            run.parent = getattr(item, "parent", None)
            offset = edge.start if at_start else edge.end
            return Range(paragraph, offset, offset + len(text))
        segments = [segment for segment in inside if segment.editable]
        if segments:
            target = segments[0] if at_start else segments[-1]
            at = 0 if at_start else len(target.text)
            set_text(target.item, target.text[:at] + text + target.text[at:])
            offset = target.start + at
            # each of Office JS's characters the element it reads from (CR-013 Q2)
            explode_mapped(paragraph.segments(), offset, offset + len(text))
            return Range(paragraph, offset, offset + len(text))
        # an empty control: the text becomes its first run
        items = self.content
        run = text_run(text)
        items.insert(0, run)
        link_parents(run)
        span = self._run_span()
        start = span[1] if span is not None else 0
        return Range(paragraph, start, start + len(text))

    def _inside(self, run: Any) -> bool:
        """Whether a run sits inside this control."""
        current = run
        for _ in range(32):
            if current is None:
                return False
            if current is self.element:
                return True
            current = getattr(current, "parent", None)
        return False

    def insert_paragraph(self, text: str = "", *, location: str = "End", **options: Any):
        """A paragraph in, before or after the control.

        ``"Start"`` and ``"End"`` put it inside a block control; ``"Before"``
        and ``"After"`` beside it, which is the only thing a run-level control
        offers. Into a block control **showing its placeholder** (CR-019
        sections 4.2 and 11, core-ts's D1): the placeholder goes, its paragraph
        left empty, the flag cleared, and the new paragraph follows it, in
        either form, as Word's (core-ts check 64 cases 6 and 27, check 63's run
        2 cases 3 and 29).

        Raises:
            InvalidTargetError: a row- or cell-level control asked for
                ``"Start"`` or ``"End"``; the message names what it holds.
        """
        form = self.form
        if form in ("row", "cell") and location in ("Start", "End"):
            holds = "rows" if form == "row" else "cells"
            raise self._holds(
                f"a {form}-level content control holds {holds}, not paragraphs",
                "insert into a cell's body",
            )
        text = drop_unwritable(text)  # CR-013 section 11, E3
        if location in ("Start", "End") and form == "block" and self.is_showing_placeholder:
            # the placeholder goes, its paragraph left empty, the flag cleared, and the new paragraph
            # follows, in either form (core-ts check 64 cases 6 and 27, check 63's run 2 cases 3 and 29;
            # CR-019 sections 4.2 and 11, core-ts's D1 at d84f3da)
            self._take_placeholder_runs(form)
            self.remove_property("showingPlcHdr", W_NS)
        if form == "run" or location in ("Before", "After"):
            where = "Before" if location == "Before" else "After"
            paragraph = self.parent_paragraph()
            if paragraph is not None:
                return paragraph.insert_paragraph(text, location=where, **options)
            element = P(content=ChildList([text_run(text)]))
            link_parents(element)
            view = self.parent_body.insert_element(element, location=where, target=self)
            if options.get("style"):
                view.style = options["style"]
            return view
        return self.body().insert_paragraph(text, location=location, **options)

    def delete(self, *, keep_content: bool = True) -> None:
        """Remove the control, keeping what it held by default (Office JS ``delete``).

        ``keep_content=True`` puts the control's children where the control
        was, as Word's "remove content control" does; False removes them with
        it. Deleting a content control is **not** a tracked change (CR-003
        section 4), and this phase writes no revision markup at all.

        **In a table cell**, where the block control held the paragraph the cell
        ends with, that paragraph stays in the control's place, emptied, its
        properties kept, outside any control, as Word keeps a cell's last
        paragraph (CR-016 Phase A; check 52 case 95). A cell left ending in a
        table any other way is given docx4j's repair.
        """
        from docx4j_py.model.content.addresses import assign_para_id
        from docx4j_py.model.content.table_rules import (
            end_with_paragraph,
            ends_with_paragraph,
            paragraph_end_of,
            wrapping_cell_of,
        )

        with recording(self.parent_body, "delete") as change:
            index = self.index
            if index < 0:
                return
            change.touched(self.address)
            change.text(before=self.text, after="" if not keep_content else self.text)
            owner = getattr(self.container, "owner", None)
            cell = wrapping_cell_of(owner) if self.form == "block" else None
            end = paragraph_end_of(cell) if cell is not None else None
            items = list(self.content) if keep_content else []
            self.container[index : index + 1] = items
            if items:
                for item in items:
                    link_parents(item)
                    if owner is not None:
                        item.parent = owner
            if cell is not None and not ends_with_paragraph(cell):
                if end is not None:
                    end.element.content.clear()
                    self.container.insert(index, end.element)
                    if owner is not None:
                        end.element.parent = owner
                else:
                    change.created(assign_para_id(self.parent_body, end_with_paragraph(cell)))

    # -- writing into a control showing its placeholder (CR-019 Phase A) ----

    @property
    def placeholder_entry(self) -> str | None:
        """The glossary entry the control's ``w:placeholder`` names; None where it has none. Extension.

        CR-019 section 4.1 (core-ts ``placeholderEntry`` at ``d84f3da``, its
        CR-002 sections 63.9 and 63.17): ``w:placeholder/w:docPart/@w:val``,
        ``""`` for a ``w:placeholder`` naming none. With any entry, present in
        the glossary or not, the control is in **Word's form** (core-ts checks
        64 and 65); with None, in the **bare form**, docx4j's binding's. Writes
        are the same in both forms (core-ts check 63's run 2, CR-019 section
        11): the form decides only what a control emptied shows (its entry, or
        Word's five-space filler, :meth:`_show_emptied`) and what
        :attr:`placeholder_text` reads (CR-019 Phase B).
        """
        from docx4j_py.model.content.placeholders import placeholder_entry_of

        return placeholder_entry_of(self.sdt_pr)

    @property
    def default_placeholder_entry(self) -> Any:
        """Word's default placeholder entry for this control's kind, a :class:`~.placeholders.DefaultEntry`, or None. Extension.

        core-ts ``defaultPlaceholderEntry`` (``b12e0c5``, its section 63.22;
        docx4j CR-063's D8): the text kinds' ``DefaultPlaceholder_-1854013440``,
        the date picker's, the lists', a building block gallery's and a
        repeating section item's (:data:`~.placeholders.DEFAULT_ENTRIES`); None
        for a picture, a check box, a repeating section and the kinds Word was
        not measured making.
        """
        from docx4j_py.model.content.placeholders import default_entry_for

        return default_entry_for(self.sdt_pr)

    @property
    def shows_words_form_placeholder(self) -> bool:
        """Whether an empty control shows a placeholder in Word's form. Extension.

        core-ts ``showsWordsFormPlaceholder`` (``b12e0c5``; D8, docx4j CR-063's
        D3 and D8): a kind Word gives a default entry
        (:attr:`default_placeholder_entry`), or one naming an entry the
        glossary holds. The binding pass gives an empty value such a control's
        placeholder (:meth:`show_placeholder`), and any other docx4j's run.
        """
        if self.default_placeholder_entry is not None:
            return True
        name = self.placeholder_entry
        glossary = self._glossary
        return bool(name) and glossary is not None and glossary.has(name)

    @property
    def _glossary(self) -> Any:
        """The package's :class:`~.placeholders.GlossaryEntries`, where the control is in a package (a trial's its own)."""
        from docx4j_py.model.content.placeholders import glossary_entries_of

        return glossary_entries_of(getattr(self.parent_body, "package", None))

    def _glossary_part_name(self) -> str:
        """The glossary document part's name, as a report names it: the package's, or the one a save would make."""
        from docx4j_py.model.content.placeholders import GLOSSARY_DOCUMENT_NAME

        package = getattr(self.parent_body, "package", None)
        main = getattr(package, "main_document_part", None) if package is not None else None
        part = getattr(main, "glossary_document_part", None) if main is not None else None
        return str(part.part_name) if part is not None else GLOSSARY_DOCUMENT_NAME

    def show_placeholder(self) -> None:
        """Show the control's placeholder: a copy of its glossary entry's content, with ``w:showingPlcHdr``. Extension.

        core-ts ``showPlaceholder`` at ``b12e0c5`` (its CR-002 sections 63.12,
        63.18, 63.21 and 63.22; CR-019 section 4.8, D8). The entry is the one
        ``w:placeholder`` names (:attr:`placeholder_entry`). A control naming
        none, or one the glossary lacks, is made to name **Word's default entry
        for its kind** (:attr:`default_placeholder_entry`: the text kinds'
        ``DefaultPlaceholder_-1854013440``, a date picker's "Click or tap to
        enter a date.", a list's "Choose an item."), which is written into the
        glossary on save where it lacks it (D9), and nothing is written under
        the missing name, as docx4j's CR-063 has it (Word's save drops a
        reference no entry holds, core-ts check 65). A kind Word gives no
        default entry (a picture, a check box) keeps the name it has, and with
        no entry to show, docx4j's run (``PLACEHOLDER_TEXT`` in the
        ``PlaceholderText`` style over the control's run properties), with no
        reference.

        What the control shows is a copy of each of the entry's paragraphs'
        run content (:func:`~.placeholders.entry_content_of`), every run as the
        glossary holds it, the control's own ``w:sdtPr/w:rPr`` not over it, as
        Word brings an entry back (core-ts check 68 cases 21 to 24, its section
        63.21; until CR-019's port of ``c275a2a`` the control's properties went
        over each run, docx4j's ``createPlaceholder`` order). Where no glossary
        entry is there to copy, the default text's run alone, in the
        ``PlaceholderText`` style. A block control takes every paragraph, but a bound one
        (:func:`is_bound_block_control`) only the first; the further paragraphs
        take a copy of the control's first paragraph's ``w:pPr``, as docx4j's
        binding gives them. A run-level control takes the first paragraph's
        runs. The entry's own paragraph properties stay in the glossary.
        """
        from docx4j_py.child import deep_copy
        from docx4j_py.model.content.placeholders import (
            entry_content_of,
            name_placeholder_entry,
        )

        with recording(self.parent_body, "show_placeholder") as change:
            glossary = self._glossary
            fallback = self.default_placeholder_entry
            name = self.placeholder_entry or ""
            # a kind's default entry where it names none or one the glossary lacks (docx4j CR-063's D8); a glossary is
            # read on first use, so core-ts's "until the glossary is read, a name is kept" does not arise here
            missing = name == "" or (glossary is not None and not glossary.has(name))
            if missing and fallback is not None:
                name = fallback.name
                name_placeholder_entry(self, name)
            if (
                fallback is not None
                and name == fallback.name
                and glossary is not None
                and glossary.require(name, fallback.text)
            ):
                change.part(self._glossary_part_name())
            body = glossary.body_of(name) if glossary is not None and name else None
            if body is not None:
                paragraphs = entry_content_of(body)
            elif fallback is not None:
                # the kind's default text as its entry holds it, the control's run properties not under it (check 68)
                paragraphs = [[_placeholder_run(fallback.text, None)]]
            else:
                # no entry to show: docx4j's run, over the control's run properties (CR-063's D8 (a))
                from docx4j_py.model.customxml.bindings import PLACEHOLDER_TEXT

                paragraphs = [[_placeholder_run(PLACEHOLDER_TEXT, self.run_properties)]]
            change.text(before=self._shown_text(), after="")
            self.set_bound_content(paragraphs[0] if paragraphs else [], True)
            change.touched(self.address)
            # a block control shows an entry's further paragraphs too, where it may hold more than one
            if self.form != "block" or len(paragraphs) < 2 or is_bound_block_control(self):
                return
            content = self.content
            first = next((item for item in content if isinstance(item, P)), None)
            at = (next(i for i, item in enumerate(content) if item is first) + 1) if first is not None else len(content)
            p_pr = first.p_pr if first is not None else None
            owner = getattr(content, "owner", None) or getattr(self.element, "sdt_content", None)
            for runs in paragraphs[1:]:
                paragraph = P(p_pr=deep_copy(p_pr) if p_pr is not None else None, content=ChildList(runs))
                link_parents(paragraph)
                content.insert(at, paragraph)
                paragraph.parent = owner
                at += 1

    @property
    def shows_placeholder_content(self) -> bool:
        """Whether a text write into this control replaces its placeholder: a block or run control showing it. Extension.

        core-ts ``showsPlaceholderContent`` (CR-019 section 4.2). A row or cell
        control's placeholder is left to the writes of its cells.
        """
        return self.is_showing_placeholder and self.form in ("block", "run")

    def _write_placeholder(self, text: str, *, breaks: str = "paragraph") -> Range:
        """Text written into this control while it shows its placeholder, by any route (CR-019 sections 4.2 and 11).

        core-ts ``writePlaceholder`` (``c275a2a``; its D1, D2 (b), D5 and
        D11). The routes: :meth:`insert_text`, its ``"Content"`` range's
        ``insert_text`` (:meth:`_write_through`), and a block control's
        paragraph's ``insert_text``, ``text = ...`` and ``splice``. **One rule
        for both forms**: the placeholder is no content (core-ts checks 64 and
        67, check 63's run 2). Its runs come out untracked (a block control's
        other paragraphs go), the text takes their place in the first
        paragraph, ``w:showingPlcHdr`` is cleared, and under tracking only the
        insertion is recorded. **The text takes the control's own run
        properties, and nothing of the placeholder's** (D11, core-ts check 68
        cases 1 to 20; :meth:`_take_control_properties`): each run the write
        leaves in the control a copy of ``w:sdtPr/w:rPr``, and plain where the
        control has none. Until CR-019 Phase B's port of ``c275a2a`` the bare
        form's text took the first placeholder run's properties less the
        ``PlaceholderText`` style (CR-016 Phase B's rule), which check 68's
        italic placeholder (cases 1 to 4) showed is not Word's.

        One paragraph, tracked as untracked (D5, core-ts's F1 and F3): the flag
        is cleared, so the tracked write reads as text. A run-level control's
        refusals come first, before anything changes. `breaks` is
        :meth:`Paragraph.splice`'s, ``"line"`` for the one-paragraph writers.
        A bound control's node takes the text (D7, Word's behaviour in both
        forms since check 63's run 2).

        Raises:
            SpanError: ``range.splits_control``, a paragraph break into a
                run-level control; nothing has changed.
            ContentError: ``control.no_paragraph``, a run-level control in no
                paragraph.
        """
        form = self.form
        text = drop_unwritable(text)
        with recording(self.parent_body, "insert_text") as change:
            if form == "run":
                # refused before anything changes: a paragraph break into a run-level control (F4)
                self._refuse_run_text(text)
            # the report says what the write replaced: the prompt (CR-016 section 11.6 item 6)
            change.text(before=self._shown_text())
            self._take_placeholder_runs(form)
            self.remove_property("showingPlcHdr", W_NS)
            if form == "run":
                paragraph = self.parent_paragraph()
                assert paragraph is not None  # _refuse_run_text refused one in no paragraph
                out = self._insert_into_empty(paragraph, text)
                change.touched(paragraph)
            else:
                first = self._first_paragraph()
                out = first.splice(0, 0, text, breaks=breaks)
                change.touched(out.paragraph, out.end_paragraph)
            self._take_control_properties()
            change.text(after=text)
            self.write_through()
            return out

    def _shown_text(self) -> str:
        """The text the control's runs show, deleted runs and the placeholder rule aside: a block's paragraphs ``\\r`` apart."""
        if self.form == "run":
            inner = getattr(self.element, "sdt_content", None)
            return text_of_view(inner) if inner is not None else ""
        return "\r".join(text_of_view(paragraph.element) for paragraph in self.paragraphs)

    def _run_content(self) -> Any:
        """A run-level control's ``w:sdtContent``, made where it has none."""
        from docx4j_py.wml import CTSdtContentRun

        inner = getattr(self.element, "sdt_content", None)
        if inner is None:
            inner = CTSdtContentRun(content=ChildList([]))
            self.element.sdt_content = inner
            inner.parent = self.element
        return inner

    def _first_paragraph(self) -> Paragraph:
        """A block control's first paragraph, made where its content holds none (core-ts ``firstParagraph``)."""
        content = self.content
        first = next((item for item in content if isinstance(item, P)), None)
        if first is None:
            first = P(content=ChildList([]))
            content.append(first)
            link_parents(first)
            first.parent = getattr(content, "owner", None) or getattr(self.element, "sdt_content", None)
        return self.body().paragraph_for(first)

    def _take_placeholder_runs(self, form: str) -> _Taken:
        """The placeholder's runs taken out of the content (core-ts ``takePlaceholderRuns``).

        A run-level control's, or a block control's paragraphs', the first
        paragraph kept, emptied, and any other going. What held them is kept,
        for Word's filler to go into (:meth:`_show_filler`).
        """
        from docx4j_py.wml import R

        def take(items: list) -> list:
            runs = [item for item in items if isinstance(item, R)]
            for run in runs:
                del items[next(i for i, item in enumerate(items) if item is run)]
            return runs

        if form == "run":
            owner = self._run_content()
            holder = run_items_of(owner)
            assert holder is not None  # a w:sdtContent's own list
            return _Taken(take(holder), holder, owner)
        content = self.content
        paragraphs = [item for item in content if isinstance(item, P)]
        runs = [run for paragraph in paragraphs for run in take(paragraph.content)]
        for paragraph in paragraphs[1:]:
            del content[next(i for i, item in enumerate(content) if item is paragraph)]
        first = paragraphs[0] if paragraphs else self._first_paragraph().element
        return _Taken(runs, first.content, first)

    def _take_control_properties(self) -> None:
        """Every run a write into the placeholder left in the control takes a copy of its ``w:sdtPr/w:rPr`` (D11).

        core-ts ``writePlaceholder``'s last step at ``c275a2a`` (its CR-002
        sections 63.20 and 63.21; check 68 cases 1 to 20, Word 365 with Design
        Mode off): the control's own run properties, in both forms and at
        every location, the placeholder run's not taken (its italic, cases 1 to
        4); where the control has none, the runs are left as the write made
        them, plain. Every run in the content is the write's: the placeholder's
        went first.
        """
        from docx4j_py.child import deep_copy

        properties = self.run_properties
        if properties is None:
            return
        for run in self._content_runs():
            run.r_pr = deep_copy(properties)
            run.r_pr.parent = run

    def _show_filler(self) -> None:
        """Word's filler for a control emptied in the bare form: the flag on over a run of five spaces (CR-019 section 11).

        core-ts ``showFiller`` (``d84f3da``, its D10 (a)): where no glossary
        entry gives the control a placeholder, Word shows one run of five
        spaces, in the control's own run properties (``w:sdtPr/w:rPr``,
        deep-copied), with ``w:showingPlcHdr`` on (core-ts check 63's run 2
        cases 6, 17, 32 and 43; check 52 case 13, whose Word's filler is bold
        because Word also gave that control a ``w:sdtPr/w:rPr`` taken from the
        deleted text, which is not followed). Whatever runs the content holds go
        first, as the placeholder's do, and the filler goes last in what held
        them, before an inner ``w:customXmlInsRangeStart`` or ``DelRangeStart``
        a tracked control ends with (core-ts ``pushBeforeMarker``). Nothing is
        marked under tracking. The spaces read as no text while the flag is
        on (CR-016 Phase B), and the control is never left a run-level
        placeholder over an empty ``w:sdtContent``, which Word refused to
        leave Design Mode over (core-ts section 63.16, control 53).
        """
        from docx4j_py.child import deep_copy

        taken = self._take_placeholder_runs(self.form)
        properties = self.run_properties
        run = text_run(FILLER, deep_copy(properties) if properties is not None else None)
        link_parents(run)
        holder = taken.holder
        last = element_name(holder[-1]) if holder else None
        holder.insert(len(holder) - 1 if last in _TRACKED_HOLDER_STARTS else len(holder), run)
        run.parent = taken.owner
        self.put_property(el.showingPlcHdr())

    def _show_emptied(self) -> None:
        """After an untracked ``"Content"`` delete that left the control holding no text (CR-019 section 11).

        core-ts ``showEmptied`` (``aefc41f``, its D10 (a) and section 63.18).
        A block or run control not showing its placeholder whose text is now
        ``""``: in Word's form it shows its entry again
        (:meth:`show_placeholder`; core-ts check 66 case 3), one naming an
        entry the glossary lacks its kind's default (``b12e0c5``); in the bare
        form, and for a control of a kind Word gives no default entry naming
        one the glossary lacks, Word's filler (:meth:`_show_filler`; core-ts
        check 52 case 13).
        """
        if self.is_showing_placeholder or self.form not in ("block", "run") or self.text != "":
            return
        name = self.placeholder_entry
        glossary = self._glossary
        if name is not None and (
            self.default_placeholder_entry is not None or (glossary is not None and glossary.has(name))
        ):
            # Word's form shows its glossary entry again (core-ts check 66 case 3; CR-019 section 4.8)
            self.show_placeholder()
            return
        self._show_filler()

    def _insert_into_empty(self, paragraph: Paragraph, text: str) -> Range:
        """Text into a run-level control holding no run: a new run in its content (core-ts ``insertIntoEmpty``).

        In a ``w:ins`` under tracking (Word: check 30, new-control-typed),
        placed before the control's trailing ``w:customXmlInsRangeStart`` where
        the control is itself an insertion. A single-line plain-text control
        drops a line break (F3).
        """
        from docx4j_py.model.content.range import Range

        if plain_text_lines_of(self) == "single":
            text = for_writer(text, "single-line")
        owner = self._run_content()
        items = run_items_of(owner)
        assert items is not None  # a w:sdtContent's own list
        tracker = paragraph.change_tracker
        run = text_run(text)
        link_parents(run)
        item = tracker.ins([run]) if tracker is not None else run
        at = next((i for i, x in enumerate(items) if element_name(x) == _CUSTOM_XML_INS_RANGE_START), len(items))
        items.insert(at, item)
        item.parent = owner
        span = self._run_span()
        start, end = (span[1], span[2]) if span is not None else (0, len(text))
        return Range(paragraph, start, end)

    def _write_through(self, part: str, text: str, location: str) -> Range | None:
        """A write through the control's own ``"Content"`` or ``"Whole"`` range that the control takes (CR-019 section 4.4).

        core-ts ``writeThrough`` (its D3 and D4, F2):

        * the control showing its placeholder: :meth:`_write_placeholder`
          (core-ts check 63 cases 4, 5, 15 and 16; check 64 cases 4, 5, 13 and
          14);
        * an empty run-level control: the text inside it (check 64 cases 21
          and 42);
        * at ``"Start"`` or ``"Before"``, a run-level control holding text:
          inside it, before its text, as core-ts writes it, where the range's
          own write joins the run before (CR-019 Q4; unmeasured);
        * ``"Whole"`` at ``"Replace"``, a run-level control showing its
          placeholder: :meth:`_take_whole`, the text in its place (D3; check 63
          cases 18 and 44).

        None where the range's own write applies.
        """
        if part == "Whole":
            if location != "Replace" or self.form != "run" or not self.is_showing_placeholder:
                return None
            return self._take_whole(text)
        if self.shows_placeholder_content:
            return self._write_placeholder(text)
        if self.form != "run":
            return None
        place = self._run_place()
        if place is None:
            return None
        paragraph, start, _end, empty = place
        if not empty and location not in ("Start", "Before"):
            return None
        text = drop_unwritable(text)
        with recording(self.parent_body, "insert_text") as change:
            self._refuse_run_text(text)
            change.text(before=self.text)
            if empty:
                out = self._insert_into_empty(paragraph, text)
            elif paragraph.change_tracker is not None:
                # the w:ins before the control's first run, inside it, in its look (Paragraph.splice's prefer)
                out = paragraph.splice(start, start, text, prefer="after")
            else:
                out = self._extend(paragraph, text, at_start=True)
            change.touched(paragraph)
            change.text(after=text)
            self.write_through()
            return out

    def _delete_through(self, part: str) -> bool:
        """A delete through the control's own ``"Content"`` or ``"Whole"`` range that the control takes (CR-019 section 4.4).

        core-ts ``deleteThrough`` (``d84f3da``). A ``"Content"`` delete of a
        control showing its placeholder keeps the control and the flag: in the
        bare form over Word's filler, five spaces, nothing marked under
        tracking (:meth:`_show_filler`, D10 (a); core-ts check 63's run 2 cases
        6, 17, 32 and 43); in Word's form the placeholder is no content, and
        nothing changes (unmeasured, section 4.10). A ``"Whole"`` delete takes
        a run-level control showing its placeholder, as its Replace does (D3,
        unmeasured). False where the range's own delete applies.
        """
        if part == "Whole":
            if self.form != "run" or not self.is_showing_placeholder:
                return False
            self._take_whole()
            return True
        if not self.shows_placeholder_content:
            return False
        with recording(self.parent_body, "delete") as change:
            if self.placeholder_entry is not None:
                return True
            change.text(before=self._shown_text(), after="")
            self._show_filler()
            change.touched(self.address)
        return True

    def _take_whole(self, text: str = "") -> Range:
        """Takes this control, showing its placeholder, whole (CR-019 section 4.3, core-ts's D3 and ``takeWhole``).

        A Replace or delete over a control showing its placeholder, as Word
        takes one (core-ts check 63 cases 12, 18 to 20, 26, 38, 44 to 46 and
        52). **Untracked** it goes, a block control leaving its first
        paragraph, emptied, in its place; `text` goes where a run-level
        control stood. **Tracked** it is marked deleted in Word's form
        (:func:`~.tracking.mark_holder_deleted`: ``w:customXmlDelRange``
        markers around it and inside its content), nothing inside marked
        deleted, its placeholder gone unmarked and the flag cleared (core-ts's
        D3 at ``d84f3da``; check 63's run 2 cases 38, 44 to 46 and 52); `text`
        goes first, before the control, as Word puts a Replace's insertion
        first. A bound control taken leaves its node as it was: nothing is
        bound any more.

        Returns:
            The range of the text, or the collapsed place the control stood.
        """
        from docx4j_py.model.content.range import Range

        form = self.form
        with recording(self.parent_body, "insert_text" if text else "delete") as change:
            change.text(before="", after=text)
            tracker = self.parent_body.change_tracker
            if form == "run":
                span = self._run_span()
                if span is None:
                    raise ContentError(
                        "this run-level content control is not in a paragraph",
                        code="control.no_paragraph",
                        hint="a run-level w:sdt is edited through the w:p it sits in",
                    )
                paragraph, start, _end = span
                change.touched(paragraph)
                if tracker is None:
                    index = self.index
                    if index >= 0:
                        del self.container[index]
                    return paragraph.splice(start, start, text) if text else Range(paragraph, start, start)
                out = paragraph.splice(start, start, text) if text else Range(paragraph, start, start)
                self._mark_whole_deleted(tracker, form)
                return out
            first = self._first_paragraph()
            if tracker is not None:
                self._mark_whole_deleted(tracker, form)
                return Range(first, 0, 0)
            index = self.index
            self._take_placeholder_runs(form)
            element = first.element
            owner = getattr(self.element, "parent", None)
            if index >= 0:
                self.container[index : index + 1] = [element]
            element.parent = owner
            view = self.parent_body.paragraph_for(element)
            change.touched(view)
            return Range(view, 0, 0)

    def _mark_whole_deleted(self, tracker: Any, form: str) -> None:
        """The control marked deleted with Word's markers, its placeholder gone, unmarked, the flag cleared (:meth:`_take_whole`).

        core-ts ``markWholeDeleted`` at ``d84f3da`` (check 63's run 2 cases 38
        and 44 to 46). Rejected, the control comes back empty, without the
        flag, as Word's markup implies (unmeasured).
        """
        from docx4j_py.model.content.tracking import mark_holder_deleted

        mark_holder_deleted(tracker, self.element, self.container, lambda: self._take_placeholder_runs(form))
        self.remove_property("showingPlcHdr", W_NS)

    # -- the w:sdtPr properties (CR-003 section 3.7, Phase E) ---------------

    @property
    def placeholder_text(self) -> str | None:
        """The placeholder Word shows for the empty control: its glossary entry's text, or None.

        Office JS ``placeholderText`` (CR-019 section 4.8, Q2; core-ts at
        ``aefc41f``, its D8). **Read:** the text of the glossary entry the
        control's ``w:placeholder`` names (:attr:`placeholder_entry`), its
        paragraphs ``\r`` apart, as Word reads it (core-ts check 65); None in
        the bare form (no ``w:placeholder``) or where the glossary lacks the
        entry, as Word's ``null``. The glossary is read privately
        (``pkg.glossary_entries``), so reading leaves it byte for byte. Until
        CR-019 Phase B this read the control's runs, and ``""`` where it was not
        showing its placeholder.

        **Written:** the control's kind's default text
        (:attr:`default_placeholder_entry`; ``"Click or tap here to enter
        text."`` for a text control, and for a kind with none), or ``""``,
        names that kind's default entry (``DefaultPlaceholder_-1854013440``
        for text; core-ts ``b12e0c5``); any other text a new entry of its
        own, named as Word names one (32 upper-case hex digits), which the save
        writes into the glossary, making the glossary where the file has none
        (D9). The entry named before stays in the glossary. A control showing
        its placeholder, or holding nothing, shows the text as the entry holds
        it, one run in the ``PlaceholderText`` style (the control's own run
        properties not under it, check 68), with ``w:showingPlcHdr``; a control holding content keeps it, and only its
        reference changes (until Phase B it refused, ``binding.holds_content``).
        """
        name = self.placeholder_entry
        if name is None:
            return None
        glossary = self._glossary
        return glossary.text_of(name) if glossary is not None else None

    @placeholder_text.setter
    def placeholder_text(self, value: str) -> None:
        """Name an entry holding the text; a control showing its placeholder, or holding nothing, shows it."""
        from docx4j_py.model.content.placeholders import (
            DEFAULT_ENTRIES,
            name_placeholder_entry,
            new_entry_name,
        )

        # the kind's default text, or "", names its default entry (core-ts b12e0c5, section 63.22); any other text a
        # new entry of its own
        fallback = self.default_placeholder_entry or DEFAULT_ENTRIES["text"]
        own = value not in ("", fallback.text)
        glossary = self._glossary
        name = fallback.name
        if own:
            # a name the glossary holds no entry of: a trial's generator starts where the real package's did
            package = getattr(self.parent_body, "package", None)
            name = new_entry_name(package)
            while glossary is not None and glossary.has(name):
                name = new_entry_name(package)
        shown = value if own else fallback.text
        with recording(self.parent_body, "placeholder_text") as change:
            change.touched(self.address)
            empty = self.is_showing_placeholder or self.text == ""
            name_placeholder_entry(self, name)
            if glossary is not None and glossary.require(name, shown):
                change.part(self._glossary_part_name())
            if empty:
                # the entry's run, as the glossary will hold it: the control's run properties aren't under it (check
                # 68; core-ts c275a2a)
                self.set_bound_content([_placeholder_run(shown, None)], True)
            change.text(after=shown)

    @property
    def is_showing_placeholder(self) -> bool:
        """Whether ``w:showingPlcHdr`` is set, and not ``w:val="0"``. Extension.

        ``w:showingPlcHdr`` is an on/off property, so ``w:val="0"`` (or ``"false"``,
        or ``"off"``) turns it off, as core-ts's ``isShowingPlaceholder`` reads it
        and as Word reads an on/off value (CR-018).
        """
        item = sdt_property(self.sdt_pr, "showingPlcHdr", W_NS)
        return is_on(item)

    @property
    def showing_placeholder_text(self) -> bool:
        """Whether the control is showing its placeholder text rather than a value.

        Office JS ``showingPlaceholderText`` (WordApiDesktop 1.3, read-only):
        :attr:`is_showing_placeholder` under Office JS's name, as core-ts's
        ``222cc20`` has it (CR-003 section 24).
        """
        return self.is_showing_placeholder

    @property
    def appearance(self) -> str:
        """``w15:appearance``: ``"BoundingBox"``, ``"Tags"`` or ``"Hidden"``.

        Office JS ``appearance``. ``"BoundingBox"`` when the control says
        nothing, which is what Word draws.
        """
        value = getattr(sdt_property(self.sdt_pr, "appearance", W15_NS), "val", None)
        value = getattr(value, "value", value)
        return _APPEARANCE_IN.get(str(value), "BoundingBox") if value else "BoundingBox"

    @appearance.setter
    def appearance(self, value: str) -> None:
        """Set how Word draws the control."""
        from docx4j_py.model.content.errors import BindingError
        from docx4j_py.w15 import el as w15_el

        stored = _APPEARANCE_OUT.get(str(value))
        if stored is None:
            raise BindingError(
                f"not a content-control appearance: {value!r}",
                code="binding.bad_appearance",
                hint="one of BoundingBox, Tags, Hidden",
            )
        with self._property_change("appearance", value, needs="w15:appearance"):
            self.put_property(w15_el.appearance(val=stored))

    @property
    def color(self) -> str:
        """``w15:color``, the colour of the control's tags, as ``#RRGGBB``.

        Office JS ``color``; ``""`` when the control has none.
        """
        value = getattr(sdt_property(self.sdt_pr, "color", W15_NS), "val", None)
        value = getattr(value, "value", value)
        if not value:
            return ""
        # ``w15:color/@w:val`` is ``ST_HexColorAuto``, a union with ``xs:hexBinary``,
        # so a parsed document hands the colour back as bytes and a document
        # this API built hands back the string it was given.
        text = value.hex().upper() if isinstance(value, bytes) else str(value)
        return text if text.startswith("#") else f"#{text}"

    @color.setter
    def color(self, value: str) -> None:
        """Set the colour; ``""`` removes it. ``#RRGGBB`` or ``RRGGBB``."""
        from docx4j_py.w15 import el as w15_el

        with self._property_change("color", value, needs="w15:color"):
            if not value:
                self.remove_property("color", W15_NS)
                return
            self.put_property(w15_el.color(val=value.lstrip("#").upper()))

    @property
    def cannot_delete(self) -> bool:
        """Whether Word refuses to delete the control. Office JS ``cannotDelete``.

        ``w:lock`` is one attribute with four values, so this and
        :attr:`cannot_edit` are the two halves of it.
        """
        return self._lock() in ("sdtLocked", "sdtContentLocked")

    @cannot_delete.setter
    def cannot_delete(self, value: bool) -> None:
        """Lock or unlock the control itself."""
        with self._property_change("cannot_delete", value):
            self._set_lock(delete=bool(value), edit=self.cannot_edit)

    @property
    def cannot_edit(self) -> bool:
        """Whether Word refuses to edit the content. Office JS ``cannotEdit``."""
        return self._lock() in ("contentLocked", "sdtContentLocked")

    @cannot_edit.setter
    def cannot_edit(self, value: bool) -> None:
        """Lock or unlock the content."""
        with self._property_change("cannot_edit", value):
            self._set_lock(delete=self.cannot_delete, edit=bool(value))

    @property
    def remove_when_edited(self) -> bool:
        """``w:temporary``: Word removes the control once it is filled in.

        Office JS ``removeWhenEdited``.
        """
        element = sdt_property(self.sdt_pr, "temporary", W_NS)
        if element is None:
            return False
        return is_on(element)

    @remove_when_edited.setter
    def remove_when_edited(self, value: bool) -> None:
        """Set or clear ``w:temporary``."""
        with self._property_change("remove_when_edited", value):
            if value:
                self.put_property(el.temporary())
            else:
                self.remove_property("temporary", W_NS)

    def _property_change(self, name: str, value: Any, *, needs: str = "") -> Any:
        """``with self._property_change("color", value):`` --- one report per setter.

        CR-003 decided question 4: every mutating call records a
        :class:`~docx4j_py.model.content.reports.ChangeReport`, and a
        ``w:sdtPr`` property setter is one. ``tag`` and ``title`` have no
        setters, so there is nothing to record for them (section 17.9).

        Args:
            needs: what to name in a compatibility warning when the property is
                Word 2013's and the document's mode is below 15 (section 17.11);
                ``""`` for a property every Word understands.
        """
        from contextlib import ExitStack

        stack = ExitStack()
        change = stack.enter_context(recording(self.parent_body, f"control.{name}"))
        change.touched(self.address)
        change.text(after="" if value is None else str(value))
        if needs:
            from docx4j_py.model.content.compatibility import warn_below

            warn_below(getattr(self.parent_body, "package", None), change, needs)
        return stack

    def _lock(self) -> str:
        value = getattr(sdt_property(self.sdt_pr, "lock", W_NS), "val", None)
        value = getattr(value, "value", value)
        return str(value) if value else ""

    def _set_lock(self, *, delete: bool, edit: bool) -> None:
        if delete and edit:
            self.put_property(el.lock(val="sdtContentLocked"))
        elif delete:
            self.put_property(el.lock(val="sdtLocked"))
        elif edit:
            self.put_property(el.lock(val="contentLocked"))
        else:
            self.remove_property("lock", W_NS)

    # -- the properties list ------------------------------------------------

    def properties(self) -> ChildList:
        """The ``w:sdtPr``'s live choice list, created if the control has none."""
        from docx4j_py.wml import SdtPr

        pr = self.sdt_pr
        if pr is None:
            pr = SdtPr(content=ChildList([]))
            self.element.sdt_pr = pr
            link_parents(self.element)
        items = pr.content
        if not isinstance(items, ChildList):
            items = ChildList(list(items or []), owner=pr)
            pr.content = items
        return items

    def put_property(self, element: Any) -> None:
        """Add or replace a ``w:sdtPr`` child, in the order Word writes them.

        Extension, and the one place a property setter goes through: the model
        keeps every ``w:sdtPr`` child in one choice list (docx4j's
        ``getRPrOrAliasOrLock``), so "set the lock" is "replace the ``w:lock``".
        """
        name = element_name(element) or ""
        items = self.properties()
        for index, item in enumerate(list(items)):
            if element_name(item) == name:
                items[index] = element
                return
        items.insert(_position_for(items, name), element)

    def remove_property(self, local_name: str, namespace: str | tuple[str, ...] = W_NS) -> int:
        """Remove every ``w:sdtPr`` child of a name; returns how many went. Extension."""
        wanted = frozenset([namespace] if isinstance(namespace, str) else namespace)
        items = self.properties()
        gone = 0
        for item in list(items):
            name = element_name(item) or ""
            uri, _, local = name.rpartition("}")
            if local == local_name and uri.lstrip("{") in wanted:
                items.remove(item)
                gone += 1
        return gone

    @property
    def run_properties(self) -> Any:
        """``w:sdtPr/w:rPr``: the formatting bound content takes. Extension."""
        return sdt_property(self.sdt_pr, "rPr", W_NS)

    @property
    def is_multi_line(self) -> bool:
        """``w:text/@w:multiLine``: whether a plain-text control takes newlines."""
        text = sdt_property(self.sdt_pr, "text", W_NS)
        if text is None:
            return self.type != "PlainText"
        return on_off(effective(text, "multi_line"))  # absent: single-line (CR-059)

    # -- the XML mapping and the typed kinds --------------------------------

    @property
    def xml_mapping(self) -> Any:
        """The control's binding to a custom XML node. Office JS ``xmlMapping``."""
        from docx4j_py.model.customxml.mapping import XmlMapping

        return XmlMapping(self)

    @property
    def checkbox_content_control(self) -> Any:
        """The checkbox view, or None. Office JS ``checkboxContentControl``."""
        return self._kind_view("CheckBox")

    @property
    def date_picker_content_control(self) -> Any:
        """The date-picker view, or None. Office JS ``datePickerContentControl``."""
        return self._kind_view("DatePicker")

    @property
    def drop_down_list_content_control(self) -> Any:
        """The drop-down view, or None. Office JS ``dropDownListContentControl``."""
        return self._kind_view("DropDownList")

    @property
    def combo_box_content_control(self) -> Any:
        """The combo-box view, or None. Office JS ``comboBoxContentControl``."""
        return self._kind_view("ComboBox")

    @property
    def picture_content_control(self) -> Any:
        """The picture view, or None. Office JS ``pictureContentControl``."""
        return self._kind_view("Picture")

    @property
    def repeating_section_content_control(self) -> Any:
        """The repeating-section view, or None. Office JS ``repeatingSectionContentControl``."""
        return self._kind_view("RepeatingSection")

    @property
    def group_content_control(self) -> Any:
        """The group view, or None. Office JS ``groupContentControl``."""
        return self._kind_view("Group")

    def _kind_view(self, kind: str) -> Any:
        from docx4j_py.model.customxml.kinds import kind_view

        return kind_view(self, kind)

    # -- writing content into a control (CR-003 section 4) ------------------

    def set_bound_content(self, runs: list, showing_placeholder: bool = False) -> None:
        """Replace what the control shows with `runs`. Extension.

        docx4j's ``BindingTraverser.applyBoundContent``, with the departure
        CR-003 section 4 records: the runs go into the **first paragraph** of a
        block, row or cell control, keeping its ``w:pPr`` and everything around
        it, rather than reducing a ``w:tr`` or a ``w:tbl`` to one cell.
        """
        items = self.content
        if self.form == "run":
            items[:] = list(runs)
        else:
            paragraph = next((item for item in items if isinstance(item, P)), None)
            if paragraph is None:
                from docx4j_py.wml import p as paragraph_builder

                paragraph = paragraph_builder()
                items.append(paragraph)
            holder = run_items_of(paragraph)
            if holder is not None:
                holder[:] = list(runs)
        link_parents(self.element)
        if showing_placeholder:
            self.put_property(el.showingPlcHdr())
            # the placeholder run names PlaceholderText, which a document made here lacks: defined in the main
            # document's styles, as core-ts's setBoundContent requires it (CR-019 section 12)
            from docx4j_py.model.content.placeholders import require_placeholder_style

            require_placeholder_style(getattr(self.parent_body, "package", None))
        else:
            self.remove_property("showingPlcHdr", W_NS)

    def set_checkbox_glyph(self, symbol: str, font: str) -> None:
        """Write the run Word draws a checkbox from. Extension.

        Word draws the box from the control's **content**, not from
        ``w14:checked``, so ticking a box writes both.
        """
        from docx4j_py.child import deep_copy
        from docx4j_py.wml import R, t

        properties = deep_copy(self.run_properties) if self.run_properties is not None else el.rPr()
        properties.r_fonts = el.rFonts(ascii=font, h_ansi=font, east_asia=font, hint="eastAsia")
        run = R(r_pr=properties, content=ChildList([t(symbol)]))
        self.set_bound_content([run])

    def write_through(self) -> bool:
        """Push this control's text into the custom XML node it is bound to.

        CR-003 section 4: Word refreshes a bound control from its data on open,
        so an edit that only touched ``w:sdtContent`` would be lost. Called by
        :meth:`insert_text`, outside the tracked branch (section 16.6). A
        container, a picture and an explicitly rich-text control are skipped, as
        :meth:`~docx4j_py.model.customxml.CustomXmlPartCollection.update_from_content_controls`
        skips them.

        Returns:
            Whether a node was written.
        """
        from docx4j_py.model.customxml.bindings import write_control_to_node

        return write_control_to_node(self)

    def insert_copy_after(self) -> ContentControl:
        """Copy this control and put the copy straight after it. Extension.

        What Word's ``+`` does to a repeating section's item: a structural deep
        copy, with a fresh ``w:id`` from
        :func:`~docx4j_py.wml.sdt.next_sdt_id`.
        """
        from docx4j_py.child import deep_copy
        from docx4j_py.wml.sdt import next_sdt_id

        index = self.index
        if index < 0:
            raise ContentError(
                "this content control is no longer in its container",
                code="control.detached",
                hint="read it again from body.content_controls",
            )
        with recording(self.parent_body, "insert_copy_after") as change:
            copy = deep_copy(self.element)
            part = getattr(self.parent_body, "part", None)
            root = getattr(part, "contents", None) if part is not None else None
            container = self.container
            container.insert(index + 1, copy)
            link_parents(copy)
            view = ContentControl(copy, container, self.parent_body)
            identifier = next_sdt_id(root if root is not None else self.parent_body.container)
            view.put_property(el.id_(val=identifier))
            change.touched(view.address)
            return view

    # -- comments (CR-003 section 3.9, Phase G) -----------------------------

    def get_comments(self) -> list[Any]:
        """The comments anchored inside this control (CR-003 section 3.2).

        Extension: Office JS's ``ContentControl`` has no ``getComments``, but
        CR-003 section 3.2 lists it, and a bound control an agent has just
        filled in is exactly where a comment about it belongs.
        """
        from docx4j_py.model.content.comments import comments_of

        return comments_of(self)

    # -- output ------------------------------------------------------------

    def get_xml(self) -> str:
        """The control as XML, with docx4j's prefixes. Extension.

        Only ``SdtBlock`` is a global element declaration, so the other three
        forms are named explicitly (CR-003 section 10.4).
        """
        if isinstance(self.element, SdtBlock):
            return to_xml(self.element)
        return to_xml(self.element, name=f"{{{W_NS}}}sdt")

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary: what a tool result says about a control."""
        return {
            "kind": "control",
            "address": self.address,
            "ordinal": self.ordinal,
            "type": self.type,
            "form": self.form,
            "tag": self.tag,
            "title": self.title,
            "id": self.id,
            "text": self.text,
        }

    # -- helpers -----------------------------------------------------------

    def _property(self, local_name: str) -> str:
        value = getattr(sdt_property(self.sdt_pr, local_name, W_NS), "val", None)
        value = getattr(value, "value", value)
        return str(value) if value is not None else ""

    def _holds(self, what: str, hint: str) -> InvalidTargetError:
        return InvalidTargetError(
            f"{what}; it holds {self.text!r}",
            code="control.wrong_form",
            hint=hint,
        )

    def _ancestor_table(self) -> Table | None:
        """The table a row- or cell-level control is part of."""
        from docx4j_py.model.content.text_model import W_TBL

        current = getattr(self.element, "parent", None)
        for _ in range(_MAX_DEPTH):
            if current is None:
                return None
            if element_name(current) == W_TBL:
                container = block_children_of(getattr(current, "parent", None) or ())
                if container is None:
                    container = self.parent_body.content
                return Table(current, container, self.parent_body)
            current = getattr(current, "parent", None)
        return None

    def _run_span(self) -> tuple[Paragraph, int, int] | None:
        """The offsets a run-level control covers in its paragraph."""
        found = self._run_place()
        return None if found is None else found[:3]

    def _run_place(self) -> tuple[Paragraph, int, int, bool] | None:
        """:meth:`_run_span`, and whether the control holds nothing the walk reads (core-ts ``runSpan``'s ``empty``)."""
        from docx4j_py.model.content.text_model import segments_of

        paragraph = self.parent_paragraph()
        if paragraph is None:
            return None

        element = self.element
        at = 0

        def visit(item: Any, offset: int, *_rest: Any) -> None:
            nonlocal at
            if item is element:
                at = offset

        start: int | None = None
        end = 0
        # each hidden run's item and each object as a place of no text (``objects``), so that a
        # control holding only those --- one showing its placeholder (CR-016 Phase B), a picture
        # control --- is found where it stands, collapsed; a control holding nothing at all stands
        # where the walk passes it, after the text before it (core-ts's runSpan; CR-019 section 4.4),
        # where it was found at 0
        for segment in segments_of(paragraph.element, hidden=paragraph.hidden_runs(), objects=True, visit=visit):
            if not self._inside(segment.run):
                continue
            if start is None:
                start = segment.start
            end = segment.end
        if start is None:
            return paragraph, at, at, True
        return paragraph, start, end, False


# ---------------------------------------------------------------------------
# finding controls
# ---------------------------------------------------------------------------

_BLOCK_NAMES = frozenset(
    {
        f"{{{W_NS}}}p",
        f"{{{W_NS}}}tbl",
        W_SDT,
        f"{{{W_NS}}}customXml",
        f"{{{W_NS}}}bookmarkStart",
        f"{{{W_NS}}}bookmarkEnd",
        f"{{{W_NS}}}altChunk",
    }
)


#: What a control's content may open with that tells nothing of its level (:attr:`ContentControl.form`).
_MARKER_NAMES = frozenset(
    f"{{{W_NS}}}{local}"
    for local in (
        "bookmarkStart",
        "bookmarkEnd",
        "commentRangeStart",
        "commentRangeEnd",
        "permStart",
        "permEnd",
        "proofErr",
        "moveFromRangeStart",
        "moveFromRangeEnd",
        "moveToRangeStart",
        "moveToRangeEnd",
        "customXmlInsRangeStart",
        "customXmlInsRangeEnd",
        "customXmlDelRangeStart",
        "customXmlDelRangeEnd",
    )
)


def _ends_with_table(content: list) -> bool:
    """Whether a control's content's last block, as a reader reads it, is a table (core-ts's ``BLOCK_ITEMS``, ``ef57729``)."""
    items = content
    for _ in range(_MAX_DEPTH):
        last = next(
            (
                item
                for item in reversed(items)
                if element_name(item) in (W_P, W_TBL, W_TR, W_TC) or is_alternate_content(item)
            ),
            None,
        )
        if last is None:
            return False
        if is_alternate_content(last):
            items = block_children_of(last) or []
            continue
        return element_name(last) == W_TBL
    return False


def _is_block(name: str | None) -> bool:
    """Whether an element name is block-level content, for :attr:`form`."""
    return name in _BLOCK_NAMES


def _collect(items: list, body: Body, out: list[ContentControl], start: int = 0, stop: int | None = None) -> None:
    """Every control in a block list, in document order, nested ones included.

    `start` and `stop` take ``items[start:stop]`` only, each control still held
    by the live `items` (CR-013 Phase B's ``Range.content_controls``).
    """
    for item in items if start == 0 and stop is None else items[start:stop]:
        if isinstance(item, _SDT_CLASSES):
            control = ContentControl(item, items, body)
            out.append(control)
            children = block_children_of(item)
            if children is not None:
                _collect(children, body, out)
            continue
        if isinstance(item, P):
            _collect_run_controls(item, body, out)
            continue
        children = block_children_of(item)
        if children is not None:
            _collect(children, body, out)


def _collect_run_controls(holder: Any, body: Body, out: list[ContentControl]) -> None:
    """Every run-level control in a run holder, in order, nested ones included."""
    items = run_items_of(holder)
    if items is None:
        return
    for item in items:
        if isinstance(item, _SDT_CLASSES):
            out.append(ContentControl(item, items, body))
            _collect_run_controls(item, body, out)
            continue
        if element_name(item) != f"{{{W_NS}}}r":
            _collect_run_controls(item, body, out)


def controls_in(body: Body) -> list[ContentControl]:
    """Every content control in a body, in document order, nested ones included."""
    out: list[ContentControl] = []
    _collect(body.content, body, out)
    return out


def controls_in_paragraph(paragraph: Paragraph) -> list[ContentControl]:
    """The run-level controls in a paragraph, in order, nested ones included."""
    out: list[ContentControl] = []
    _collect_run_controls(paragraph.element, paragraph.parent_body, out)
    return out


def is_bound_block_control(control: ContentControl) -> bool:
    """Whether Word holds a block control to one paragraph and reads it as plain text. Extension.

    docx4j-core-ts CR-002 section 61, as its Word check 56 measured it (``38dcae9``,
    H1, superseding D1; built there as ``isBoundBlockControl`` at ``6285dfa``): a
    block-level control with a ``w:dataBinding`` that has no type element, a
    ``w:text`` or a ``w:richText``, and is not a container
    (:func:`~docx4j_py.model.customxml.bindings.is_container`) --- save one whose
    node holds a flat OPC package, Word's own rich-text binding (not measured).
    ``w15:dataBinding`` isn't this, as in core-ts. Word can't open a file where
    such a control holds a second block (CR-013 section 10.10; check 56's four
    files), and reads every one as plain text on open.
    """
    from docx4j_py.model.customxml.bindings import is_container

    if control.form != "block":
        return False
    sdt_pr = control.sdt_pr
    if sdt_property(sdt_pr, "dataBinding", W_NS) is None:
        return False
    plain = sdt_property(sdt_pr, "text", W_NS) is not None
    rich = sdt_property(sdt_pr, "richText", W_NS) is not None
    untyped = control.type == "RichText" and not rich
    if not (plain or rich or untyped):
        return False
    return not is_container(control) and not _node_holds_flat_opc(control)


def plain_text_lines_of(control: ContentControl) -> str | None:
    """How a control holds lines, as Word reads it: ``"multi"``, ``"single"``, or None for any other. Extension.

    core-ts's ``plainTextLinesOf`` (``6285dfa``): a ``w:text`` control, bound or
    not, block or run, is multi-line where ``w:multiLine`` is set and single-line
    where not (CR-013 section 11, item 2 of the review's four); a bound block
    control Word reads as plain text (:func:`is_bound_block_control`) is
    single-line, since Word gives one with no type or ``w:richText`` a
    ``<w:text/>`` on open (check 56; H4). :attr:`ContentControl.is_multi_line`
    and :attr:`ContentControl.type` still read the file as it is: their reading
    of such a control is CR-014's.
    """
    text = sdt_property(control.sdt_pr, "text", W_NS)
    if text is not None:
        return "multi" if on_off(effective(text, "multi_line")) else "single"
    return "single" if is_bound_block_control(control) else None


def block_control_of(paragraph: Paragraph) -> ContentControl | None:
    """The block-level content control holding a paragraph directly, or None. Extension (core-ts's ``blockControl``)."""
    content = getattr(paragraph.element, "parent", None)
    sdt = getattr(content, "parent", None) if isinstance(content, SdtContentBlock) else None
    if not isinstance(sdt, SdtBlock):
        return None
    owner = getattr(sdt, "parent", None)
    container = block_children_of(owner) if owner is not None else None
    if container is None or not any(item is sdt for item in container):
        container = [sdt]
    return ContentControl(sdt, container, paragraph.parent_body)


def paragraph_lines_of(paragraph: Paragraph) -> str | None:
    """How the block control holding a paragraph directly takes lines: :func:`plain_text_lines_of`'s, or None. Extension.

    Where it answers, a ``\\n`` or ``\\r`` written into the paragraph makes no
    paragraph (CR-013 section 11): a ``w:br`` in a multi-line control, and in a
    single-line one nothing, a line break, a page break and a column break with
    it (F3).
    """
    control = block_control_of(paragraph)
    return plain_text_lines_of(control) if control is not None else None


def controls_of_types(controls: list[ContentControl], types: Collection[str] | None) -> list[ContentControl]:
    """The controls whose :attr:`ContentControl.type` is one of `types`, or all of them (core-ts's ``controlsOfTypes``)."""
    if not types:
        return list(controls)
    wanted = {str(kind) for kind in types}
    return [control for control in controls if control.type in wanted]


#: The namespace of a flat OPC package, which Word's rich-text binding stores in its node.
_FLAT_OPC_NS = "http://schemas.microsoft.com/office/2006/xmlPackage"


def _node_holds_flat_opc(control: ContentControl) -> bool:
    """Whether the node a control is bound to holds a flat OPC package (core-ts's H1 leaves such a control as it is)."""
    try:
        node = control.xml_mapping.custom_xml_node
    except Exception:  # noqa: BLE001 - an unusable binding is no package
        return False
    text = getattr(node, "text", None) if node is not None else None
    return bool(text) and (_FLAT_OPC_NS in text or "pkg:package" in text)


def holds_one_paragraph(control: ContentControl) -> bool:
    """Whether a block control holds one paragraph, or no block at all (core-ts ``holdsOneParagraph``; CR-019 section 4.5).

    A block is a paragraph, a table, a block control, custom XML, an
    ``w:altChunk`` or a kept ``mc:AlternateContent``; a bookmark or another
    marker beside the paragraph is none.
    """
    blocks = [
        item
        for item in control.content
        if element_name(item) in (W_P, W_TBL, W_SDT, f"{{{W_NS}}}customXml", f"{{{W_NS}}}altChunk")
        or is_alternate_content(item)
    ]
    return not blocks or (len(blocks) == 1 and isinstance(blocks[0], P))


def control_view_of(element: Any, body: Body) -> ContentControl | None:
    """The view of a ``w:sdt`` in the live list holding it, found from its parent pointer; None where it is in none."""
    parent = getattr(element, "parent", None)
    if parent is None:
        return None
    for container in (block_children_of(parent), run_items_of(parent)):
        if container is not None and any(item is element for item in container):
            return ContentControl(element, container, body)
    return None


def parent_control_of(element: Any, body: Body) -> ContentControl | None:
    """The innermost control an element is inside, or None. Extension."""
    current = getattr(element, "parent", None)
    for _ in range(_MAX_DEPTH):
        if current is None:
            return None
        if isinstance(current, _SDT_CLASSES):
            container = block_children_of(getattr(current, "parent", None) or ())
            if container is None or not any(item is current for item in container):
                container = run_items_of(getattr(current, "parent", None)) or body.content
            return ContentControl(current, container, body)
        current = getattr(current, "parent", None)
    return None
