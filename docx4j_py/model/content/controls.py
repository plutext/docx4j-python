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


def _placeholder_run(text: str, run_properties: Any) -> Any:
    """docx4j's placeholder run: the prompt, styled ``PlaceholderText``."""
    from docx4j_py.child import deep_copy
    from docx4j_py.wml import R
    from docx4j_py.wml import t as text_element

    properties = deep_copy(run_properties) if run_properties is not None else el.rPr()
    properties.r_style = el.rStyle(val="PlaceholderText")
    return R(r_pr=properties, content=ChildList([text_element(text)]))


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

    # -- the w:sdtPr properties (CR-003 section 3.7, Phase E) ---------------

    @property
    def placeholder_text(self) -> str:
        """The greyed-out prompt the control shows while it is empty.

        Office JS ``placeholderText``. Read: the control's own text while
        ``w:showingPlcHdr`` is set, and ``""`` otherwise. Written: the text
        becomes the control's content, styled ``PlaceholderText``, and
        ``w:showingPlcHdr`` goes on --- but **only on a control that holds
        nothing**, because overwriting real content with a prompt is silent data
        loss (CR-003 section 4). Word keeps placeholder text in a glossary
        document part, which this phase does not create; ``w:placeholder`` is
        left exactly as it was.

        Raises:
            BindingError: the control already holds content.
        """
        return self.text if self.is_showing_placeholder else ""

    @placeholder_text.setter
    def placeholder_text(self, value: str) -> None:
        """Set the prompt; refuses over content."""
        from docx4j_py.model.content.errors import BindingError

        if self.text.strip() and not self.is_showing_placeholder:
            raise BindingError(
                f"this content control holds content ({self.text[:30]!r}), "
                "so its placeholder cannot be set over it",
                code="binding.holds_content",
                hint="clear the control first, or set its text with insert_text()",
            )
        with recording(self.parent_body, "placeholder_text") as change:
            change.touched(self.address)
            self.set_bound_content([_placeholder_run(value, self.run_properties)], True)
            change.text(after=value)

    @property
    def is_showing_placeholder(self) -> bool:
        """Whether ``w:showingPlcHdr`` is set. Extension."""
        return sdt_property(self.sdt_pr, "showingPlcHdr", W_NS) is not None

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
        self._set_lock(delete=bool(value), edit=self.cannot_edit)

    @property
    def cannot_edit(self) -> bool:
        """Whether Word refuses to edit the content. Office JS ``cannotEdit``."""
        return self._lock() in ("contentLocked", "sdtContentLocked")

    @cannot_edit.setter
    def cannot_edit(self, value: bool) -> None:
        """Lock or unlock the content."""
        self._set_lock(delete=self.cannot_delete, edit=bool(value))

    @property
    def remove_when_edited(self) -> bool:
        """``w:temporary``: Word removes the control once it is filled in.

        Office JS ``removeWhenEdited``.
        """
        element = sdt_property(self.sdt_pr, "temporary", W_NS)
        if element is None:
            return False
        value = getattr(element, "val", None)
        value = getattr(value, "value", value)
        return value is None or str(value).lower() not in ("0", "false")

    @remove_when_edited.setter
    def remove_when_edited(self, value: bool) -> None:
        """Set or clear ``w:temporary``."""
        if value:
            self.put_property(el.temporary())
        else:
            self.remove_property("temporary", W_NS)

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
        value = getattr(text, "multi_line", None)
        value = getattr(value, "value", value)
        return value is not None and str(value).lower() not in ("0", "false")

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
