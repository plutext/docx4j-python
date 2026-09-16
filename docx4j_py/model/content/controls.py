"""``ContentControl``: Office JS ``Word.ContentControl`` over ``w:sdt``.

CR-003 section 3.2, Phase C --- the **reads** and ``delete``. The typed kinds
(checkbox, date, drop-down), the bindings, ``appearance``, ``color``,
``cannot_delete``, ``placeholder_text`` and ``insert_content_control`` are Phase
E; what is here is what a reader and a text editor need.

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

from typing import TYPE_CHECKING, Any

from docx4j_py.child import ChildList, link_parents
from docx4j_py.model.content.addresses import ordinal_of
from docx4j_py.model.content.enums import RangeLocation, TextLocation
from docx4j_py.model.content.errors import ContentError, InvalidTargetError
from docx4j_py.model.content.reports import recording
from docx4j_py.model.content.table import Table
from docx4j_py.model.content.text_model import (
    W_SDT,
    W_TC,
    W_TR,
    block_children_of,
    segments_of,
)
from docx4j_py.traversal import element_name, run_items_of, text_of
from docx4j_py.wml import (
    CTSdtCell,
    CTSdtRow,
    P,
    SdtBlock,
    SdtRun,
    to_xml,
)
from docx4j_py.wml.sdt import W_NS, sdt_kind_of, sdt_property

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.body import Body
    from docx4j_py.model.content.paragraph import Paragraph
    from docx4j_py.model.content.range import Range

__all__ = [
    "ContentControl",
    "ContentControlForm",
    "controls_in",
    "controls_in_paragraph",
    "parent_control_of",
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


class ContentControl:
    """A subset of Office JS ``Word.ContentControl`` over a ``w:sdt``."""

    __slots__ = ("container", "element", "parent_body")

    def __init__(self, element: Any, container: list, parent_body: Body) -> None:
        """Build the view over a ``w:sdt``, the list holding it and its body."""
        #: The ``w:sdt``: a ``SdtBlock``, ``SdtRun``, ``CTSdtRow`` or ``CTSdtCell``.
        self.element = element
        #: The live list holding it.
        self.container = container
        #: The nearest :class:`~docx4j_py.model.content.Body`.
        self.parent_body = parent_body

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
        for item in self.content:
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
        """The control's text, a line per paragraph (docx4j ``TextUtils``).

        A run-level control holds run content, which is one line;
        :func:`~docx4j_py.traversal.text_of` reads a ``w:sdtContent`` as a run
        holder and would join a block control's paragraphs into one, so the
        other three forms answer through their :meth:`body`.
        """
        if self.form == "run":
            inner = getattr(self.element, "sdt_content", None)
            return text_of(inner) if inner is not None else ""
        return self.body().text

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
        paragraph --- and the first paragraph's range for the block, row and
        cell forms, because a ``Range`` here is within one paragraph (CR-003
        section 4).

        Raises:
            ContentError: a run-level control that is in no paragraph, or a
                control that holds no paragraph at all.
        """
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
            return Range(paragraph, start, end)
        paragraphs = self.paragraphs
        if not paragraphs:
            raise ContentError(
                f"this {self.form}-level content control holds no paragraph",
                code="control.no_paragraph",
                hint="insert one with insert_paragraph(), or read its text",
            )
        return paragraphs[0].get_range(location)

    def search(self, text: str, **options: Any) -> list[Range]:
        """Every match of `text` inside the control, as ranges."""
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

        Raises:
            InvalidTargetError: a row- or cell-level control asked to
                ``"Replace"``; the message names what it holds.
        """
        form = self.form
        if form in ("row", "cell") and location == "Replace":
            raise self._holds(
                f"a {form}-level content control", "replace the text of its cells instead"
            )
        if form != "run":
            return self.body().insert_text(text, location=location)
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
                change.touched(paragraph)
                change.text(after=text)
                return out
            out = paragraph.splice(start, end, text)
            change.touched(paragraph)
            change.text(after=text)
            return out

    def _extend(self, paragraph: Paragraph, text: str, *, at_start: bool) -> Range:
        """Put text at the start or the end of a run control's **own** content."""
        from docx4j_py.model.content.range import Range
        from docx4j_py.model.content.text_model import set_text
        from docx4j_py.wml import R, t

        segments = [
            segment
            for segment in segments_of(paragraph.element)
            if segment.editable and self._inside(segment.run)
        ]
        if segments:
            target = segments[0] if at_start else segments[-1]
            at = 0 if at_start else len(target.text)
            set_text(target.item, target.text[:at] + text + target.text[at:])
            offset = target.start + at
            return Range(paragraph, offset, offset + len(text))
        # an empty control: the text becomes its first run
        items = self.content
        run = R(content=ChildList([t(text)]))
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
        offers.

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
        if form == "run" or location in ("Before", "After"):
            where = "Before" if location == "Before" else "After"
            paragraph = self.parent_paragraph()
            if paragraph is not None:
                return paragraph.insert_paragraph(text, location=where, **options)
            from docx4j_py.wml import p as paragraph_builder

            view = self.parent_body.insert_element(
                paragraph_builder(text), location=where, target=self
            )
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
        """
        with recording(self.parent_body, "delete") as change:
            index = self.index
            if index < 0:
                return
            change.touched(self.address)
            change.text(before=self.text, after="" if not keep_content else self.text)
            items = list(self.content) if keep_content else []
            self.container[index : index + 1] = items
            if items:
                owner = getattr(self.container, "owner", None)
                for item in items:
                    link_parents(item)
                    if owner is not None:
                        item.parent = owner

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
        paragraph = self.parent_paragraph()
        if paragraph is None:
            return None

        start: int | None = None
        end = 0
        for segment in segments_of(paragraph.element):
            if not self._inside(segment.run):
                continue
            if start is None:
                start = segment.start
            end = segment.end
        if start is None:
            return paragraph, 0, 0
        return paragraph, start, end


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


def _is_block(name: str | None) -> bool:
    """Whether an element name is block-level content, for :attr:`form`."""
    return name in _BLOCK_NAMES


def _collect(items: list, body: Body, out: list[ContentControl]) -> None:
    """Every control in a block list, in document order, nested ones included."""
    for item in items:
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
