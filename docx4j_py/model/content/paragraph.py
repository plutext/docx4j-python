"""``Paragraph``: Office JS ``Word.Paragraph`` over a ``w:p``.

CR-003 section 3.2. A light view: it holds the ``w:p`` and the list that
contains it, nothing is cached, and every write mutates the tree in place and
links the parents of what it adds (CR-003 section 10.4: a constructor cannot
link a single-valued field, so a view that assigns one calls ``link_parents``).

The rules of section 4 that land here:

* ``style`` is the display name, ``style_built_in`` the ``Word.Style`` value and
  ``style_id`` the docx4j-named extension (:mod:`.styles`);
* ``font`` is over the **runs**, not the paragraph mark;
* ``alignment`` is ``"Unknown"`` when there is no ``w:jc``, and maps ``start`` /
  ``end`` as well as ``left`` / ``right`` (section 3.12);
* ``outline_level`` is ``w:outlineLvl + 1``, ``10`` when absent;
* the indents read the transitional (``w:left``) and the strict (``w:start``)
  attribute names and write whichever the paragraph already uses;
* a one-``w:p`` fragment given to :meth:`Paragraph.insert_xml` at ``"Start"`` or
  ``"End"`` has its runs merged into this paragraph, as Word's paste does.
"""

from __future__ import annotations

import contextlib
import dataclasses
from collections.abc import Callable, Iterator
from typing import TYPE_CHECKING, Any

from docx4j_py.child import ChildList, deep_copy, link_parents
from docx4j_py.model.content.addresses import ordinal_of, para_id_address
from docx4j_py.model.content.enums import (
    AlignmentValue,
    BreakTypeValue,
    ParagraphLocation,
    RangeLocation,
    TextLocation,
    TextView,
)
from docx4j_py.model.content.errors import ContentError, SpanError
from docx4j_py.model.content.font import Font
from docx4j_py.model.content.reports import (
    TextExcerpt,
    moved_by_delete,
    recording,
)
from docx4j_py.model.content.styles import (
    built_in_of,
    define_built_in,
    id_of_built_in,
    style_id_of,
    style_name_of,
)
from docx4j_py.model.content.text_model import (
    LINE_CHARACTERS,
    PARAGRAPH_BREAKS,
    canonical_view,
    drop_unwritable,
    explode_mapped,
    for_writer,
    hidden_runs_of,
    in_showing_placeholder,
    markup_text,
    office_matches,
    paragraph_mark_pieces,
    placeholder_control_of,
    refuse_branch_text,
    refuse_equation,
    refuse_mark_replacement,
    require_view,
    runs_of,
    segments_of,
    set_text,
    split_at,
    text_items,
    text_run,
)
from docx4j_py.model.property_resolver import resolver_of
from docx4j_py.namespaces import WML_NS
from docx4j_py.traversal import element_name, run_items_of
from docx4j_py.wml import P, R, br, el, t, to_xml, wml

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.body import Body
    from docx4j_py.model.content.range import Range

__all__ = ["Paragraph", "replace_matches"]

#: Twips per point: the unit every Office JS measurement here is in.
TWIPS_PER_POINT = 20

_JC_TO_ALIGNMENT: dict[str, str] = {
    "left": "Left",
    "start": "Left",
    "center": "Centered",
    "right": "Right",
    "end": "Right",
    "both": "Justified",
    "distribute": "Justified",
}
_ALIGNMENT_TO_JC: dict[str, str] = {
    "Left": "left",
    "Centered": "center",
    "Right": "right",
    "Justified": "both",
}

#: How much of the text a ``repr`` shows.
_PREVIEW = 30


def _value(obj: Any) -> Any:
    """``w:val`` as a plain value, whether the model made it an enum or not."""
    val = getattr(obj, "val", None)
    return getattr(val, "value", val)


class Paragraph:
    """A subset of Office JS ``Word.Paragraph`` over a ``w:p``."""

    __slots__ = ("container", "element", "parent_body")

    def __init__(self, element: P, container: list, parent_body: Body) -> None:
        """Build the view over a ``w:p``, the list holding it and its body."""
        #: The ``w:p``.
        self.element = element
        #: The live list holding the element: a body's, a cell's, a control's.
        self.container = container
        #: The nearest :class:`~docx4j_py.model.content.Body`; a read-only
        #: one in a kept ``mc:AlternateContent`` branch (CR-012 Phase A).
        self.parent_body = parent_body.body_for(element) if parent_body is not None else parent_body

    # -- identity ----------------------------------------------------------

    def __eq__(self, other: object) -> bool:
        """Two views of the same ``w:p`` are equal (CR-003 section 3.1)."""
        return isinstance(other, Paragraph) and other.element is self.element

    def __hash__(self) -> int:
        """Hashes by the element's identity, so views can go in a set."""
        return hash(id(self.element))

    def __str__(self) -> str:
        """The paragraph's text."""
        return self.text

    def __repr__(self) -> str:
        """``<Paragraph body/3 'Chapter 2: The …'>``."""
        text = self.text
        preview = text[:_PREVIEW] + "…" if len(text) > _PREVIEW else text
        where = self.ordinal
        return f"<Paragraph{' ' + where if where else ''} {preview!r}>"

    @property
    def index(self) -> int:
        """This paragraph's position in its container, or -1 if it is gone."""
        for position, item in enumerate(self.container):
            if item is self.element:
                return position
        return -1

    @property
    def p(self) -> P:
        """The ``w:p``; docx4j's name for :attr:`element`."""
        return self.element

    # -- text --------------------------------------------------------------

    @property
    def text(self) -> str:
        """The text as Office JS's ``paragraph.text`` gives it in Word's default view.

        CR-013 Phase A (Q1; docx4j-core-ts CR-012 phase 4, held to Word 365 by
        checks 48, 48b and 51): the accepted text, hidden text left out (a run's
        own ``w:vanish``, else its styles'), in Word's characters --- a tab
        ``\\t``, a line break U+000B, a page break U+000C, a column break U+000E,
        a ``w:cr`` U+000D, a ``w:ptab`` ``0``, the hyphens U+001E and U+001F, any
        ``w:sym`` ``(``, a note reference U+0002 and an equation as Word draws
        it (:mod:`~docx4j_py.model.content.text_model`'s table). Split it into
        lines with care: ``str.splitlines()`` also splits at U+000B, U+000C and
        U+001E. Every offset a ``Range`` or ``search`` takes counts in it.

        Office JS's text, not the text model's (CR-011 Q4; WD REQ-202's exemption for a reader of an
        emulated API): :func:`docx4j_py.model.text.extract` is the text model, and each of its runs
        records its ``office_offset`` in this text, agreeing under CR-013 Q11's character map
        (:func:`~.text_model.office_as_model`). A run- or paragraph-level
        ``mc:AlternateContent`` reads as the branch Word shows, as Office JS reads
        it (Q4 revised, CR-011 section 15; core-ts check 49); an edit through that
        text is refused (``address.read_only``), and an edit beside it works.

        Setting it replaces the content, every such element with it, hidden runs
        too: ``paragraph.text = paragraph.text`` drops hidden text, as core-ts's
        does. **A run-level content control showing its placeholder** reads as
        no text (CR-016 Phase B), so the write would take it and its binding
        with it: the write is **refused** (``SpanError`` ``range.holds_placeholder``,
        tracked or not, nothing changed), and the control is written into with
        :meth:`ContentControl.insert_text` instead (CR-016 section 7, risk 2). It writes one paragraph (CR-013 section 11, E2): each of Office JS's
        characters is written as the element it reads from (Q2;
        :func:`~.text_model.text_items`), ``\\n`` and ``\\r\\n`` a ``w:br``
        and ``\\r`` a ``w:cr``. U+0002 and the other characters XML 1.0 excludes
        are dropped (E3), so ``paragraph.text = paragraph.text`` on a paragraph
        holding a note reference loses the reference, as core-ts's does. With
        change tracking on, a paragraph whose text holds a branch's is refused
        (``address.read_only``), since its deletion would be marked in Word's
        branch alone.
        """
        return "".join(segment.text for segment in self.segments())

    @text.setter
    def text(self, value: str) -> None:
        value = drop_unwritable(value)
        self._refuse_placeholder_controls("Paragraph.text = ...")
        with recording(self.parent_body, "set_text") as change:
            change.text(before=self.text, after=value)
            if self.change_tracker is not None:
                # tracked, replacing the text is a deletion and an insertion,
                # not a new run: the splice writes both (CR-003 section 3.8); the
                # hidden runs go too, as the untracked setter takes them (CR-013);
                # one paragraph, its breaks w:br (section 11, E2)
                self.splice(0, len(self.text), value, whole=True, breaks="line")
                change.touched(self)
                return
            from docx4j_py.model.content.controls import paragraph_lines_of

            # one paragraph (E2), or as a plain-text control takes lines (section 11's item 2, F3)
            lines = paragraph_lines_of(self)
            mode = {"multi": "multi-line", "single": "single-line"}.get(lines or "", "one-paragraph")
            r_pr = None
            first = self._visible_first_run()
            if first is not None and getattr(first, "r_pr", None) is not None:
                # the first visible run's look: a hidden one's would hide the text (CR-013 section 11)
                r_pr = deep_copy(first.r_pr)
            run = R(content=ChildList(text_items(for_writer(value, mode))))
            if r_pr is not None:
                run.r_pr = r_pr
                r_pr.parent = run
            self.element.content[:] = [run]
            link_parents(self.element)
            change.touched(self)

    def _refuse_placeholder_controls(self, operation: str) -> None:
        """Refuse a write that replaces the whole content where a run-level control shows its placeholder.

        CR-016 Phase B (its section 7, risk 2): such a control reads as no text
        (core-ts CR-012 section 18), so a write that replaces the paragraph's
        content would take it, and its binding, with nothing read to say so.

        Raises:
            SpanError: ``range.holds_placeholder``; nothing has changed.
        """
        for run in runs_of(self.element):
            control = placeholder_control_of(run, self.element)
            if control is None:
                continue
            from docx4j_py.wml.sdt import sdt_property

            tag = getattr(sdt_property(getattr(control, "sdt_pr", None), "tag", WML_NS), "val", None)
            named = f" (tagged {tag!r})" if tag else ""
            raise SpanError(
                f"{operation} refused: this paragraph holds a content control showing its placeholder{named}, "
                "which reads as no text, and the write would take it (CR-016 Phase B)",
                code="range.holds_placeholder",
                hint="write into the control with ContentControl.insert_text, or delete the control first",
            )

    def text_since(self, index: int) -> str:
        """The text of this paragraph and of those a split put in front of it since it stood at `index`. Extension.

        A write through :meth:`splice` that made paragraphs leaves them in
        front of this one (CR-013 section 11, E1); their texts a ``\\r`` apart,
        as Office JS reads text across paragraphs, are what it wrote over. With
        no split, :attr:`text`.
        """
        position = _index_of(self.container, self.element)
        if index < 0 or position <= index:
            return self.text
        views = [Paragraph(item, self.container, self.parent_body) for item in self.container[index:position]]
        return "\r".join([*(view.text for view in views if isinstance(view.element, P)), self.text])

    def get_text(self, *, view: TextView = "accepted", max_chars: int | None = None) -> str:
        """The text in one of the views of a tracked document.

        In ``accepted`` and ``original``, Office JS's text in Office JS's characters (CR-013 Phase A;
        :attr:`text`), not the text model's (CR-011 Q4; WD REQ-202's exemption for a reader of an
        emulated API): :func:`docx4j_py.model.text.extract` is the text model, and each of its runs
        records its ``office_offset`` in this text. Hidden text is left out in both views. A run- or paragraph-level
        ``mc:AlternateContent`` reads as the branch Word shows, its deletions out in
        ``accepted`` and its insertions out in ``original`` (Q4 revised, section 15).

        ``"accepted"`` (the default) is :attr:`text`; ``"original"`` reads the
        document as it was before the tracked changes (CR-003 section 3.8).
        `max_chars` is the budget of CR-003 section 3.4; it truncates and does
        not say so, which is what :meth:`text_budget` is for.

        ``"markup"`` is the text model's view (CR-011 Phase B): the paragraph's
        runs as stored, each revision written as CriticMarkup and each comment
        at its reference (:func:`~.text_model.markup_text`); hidden text left
        out.

        Raises:
            ContentError: any other view (``text.view_invalid``).
        """
        view = canonical_view(view)  # type: ignore[assignment]
        require_view(view)
        out = self._text_in(view)
        return out if max_chars is None else out[:max_chars]

    def _text_in(self, view: str) -> str:
        if view == "markup":
            return markup_text(self, text_boxes=False)
        return "".join(segment.text for segment in self.segments(view=view))

    def text_budget(
        self, max_chars: int | None = None, *, view: TextView = "accepted"
    ) -> TextExcerpt:
        """:meth:`get_text` with the flag: ``TextExcerpt(text, chars, truncated)``."""
        view = canonical_view(view)  # type: ignore[assignment]
        require_view(view)
        full = self._text_in(view)
        if max_chars is None or len(full) <= max_chars:
            return TextExcerpt(full, len(full), False)
        return TextExcerpt(full[:max_chars], len(full), True)

    @property
    def runs(self) -> list[R]:
        """Every ``w:r``, direct and nested in hyperlinks, controls and insertions."""
        return runs_of(self.element)

    def segments(self, *, view: TextView = "accepted", objects: bool = False) -> list[Any]:
        """The text segments with their offsets in :attr:`text`. Extension; see :mod:`.text_model`.

        Hidden runs are left out (:meth:`hidden_runs`); `objects` lists each
        drawn object and each hidden run's item as a segment of no text where it
        stands (:func:`~.text_model.segments_of`).
        """
        return segments_of(self.element, view=view, hidden=self.hidden_runs(), objects=objects)

    def hidden_runs(self) -> Any:
        """Which of this paragraph's runs are hidden: a run's own ``w:vanish``, else its styles'. Extension.

        CR-013 Phase A (:func:`~.text_model.hidden_runs_of`): the text model's
        rule, over the package's property resolver; a paragraph with no package
        reads a run's own ``w:vanish`` only.
        """
        package = getattr(self.parent_body, "package", None)
        return hidden_runs_of(resolver_of(package) if package is not None else None, self.element)

    @property
    def mark_text(self) -> str:
        """The character Office JS gives this paragraph's mark at the end of its ``'Whole'`` range. Extension.

        ``\\t`` for the last paragraph of a table cell, the cell's end,
        whatever its mark; ``""`` for a mark tracked as deleted, which Word's
        default view shows as nothing (CR-016 Phase B; core-ts check 58 cases 6
        to 10); and ``\\r`` for any other (core-ts check 48; CR-013 Phase A).
        """
        from docx4j_py.model.content.story import mark_text_of

        return mark_text_of(self.element)

    def refuse_branch_text(self, start: int, end: int, operation: str = "this edit") -> None:
        """Refuse an edit of ``[start, end)`` that would change text read from a kept branch. Extension.

        CR-011 Phase F: a span holding text a run- or paragraph-level
        ``mc:AlternateContent`` gives this paragraph (the branch Word shows), or
        an insertion point strictly inside it, is refused; an insertion at either
        end of it goes beside the element.

        Raises:
            AddressError: ``address.read_only``.
        """
        from docx4j_py.model.content.reports import current_recorder

        verb = getattr(current_recorder(getattr(self.parent_body, "package", None)), "operation", None)
        refuse_branch_text(self.element, start, end, verb or operation, segments=self.segments())

    def refuse_equation(self, start: int, end: int, *, deleting: bool = False) -> None:
        """Refuse an edit of ``[start, end)`` that would cut an equation, or replace one tracked (CR-013 Phase A). Extension.

        `deleting`: the edit is a deletion, which takes an equation whole under
        change tracking in Word's form (CR-013 section 13.7, D2); a tracked
        replacement taking one is refused.

        Raises:
            SpanError: ``range.cuts_equation``.
            TrackedChangeError: ``tracking.equation``.
        """
        tracked = self.change_tracker is not None
        refuse_equation(self.segments(), start, end, tracked=("delete" if deleting else "replace") if tracked else False)

    def holds_branch_text(self) -> bool:
        """Whether this paragraph's text holds text read from a kept branch (CR-011 Phase F). Extension."""
        return any(segment.branch is not None and segment.text for segment in self.segments())

    def refuse_whole_text(self, operation: str) -> None:
        """Refuse a write to all of this paragraph's text when it holds a kept branch's (CR-011 Phase F)."""
        if self.holds_branch_text():
            self.refuse_branch_text(0, len(self.text), operation)

    # -- recording a format change (CR-003 section 3.4) --------------------

    @contextlib.contextmanager
    def formatting(self, name: str, read: Callable[[], Any]) -> Iterator[None]:
        """Record a ``"format"`` :class:`ChangeReport` around a property write.

        `read` is called before and after, so the report carries the values
        either side (``"alignment=Left"`` to ``"alignment=Centered"``) and an
        agent can see what it did without reading the document back.
        """
        with recording(self.parent_body, "format") as change:
            # the value before is recorded *now*, so that a setter written over
            # another one (``style`` over ``style_id``) wins: ``text_before`` is
            # set once, and the outer call gets there first
            change.text(before=f"{name}={read()}")
            yield
            self._settle_final_mark()
            change.touched(self)
            change.text(after=f"{name}={read()}")

    def _settle_final_mark(self) -> None:
        """After a write: the old final mark's record kept only while it changes something (check 29)."""
        tracker = self.change_tracker
        if tracker is None or not tracker.inserted(self.element):
            return
        from docx4j_py.model.content.tracking import settle_final_mark

        settle_final_mark(tracker, self.element, self.container)

    # -- style -------------------------------------------------------------

    @property
    def style_id(self) -> str:
        """``w:pStyle``, docx4j's view of a style; ``"Normal"`` when there is none.

        Extension: Office JS has ``style`` and ``style_built_in`` only.
        """
        p_pr = self.element.p_pr
        value = _value(getattr(p_pr, "p_style", None)) if p_pr is not None else None
        return value or "Normal"

    @style_id.setter
    def style_id(self, value: str) -> None:
        with self.formatting("style_id", lambda: self.style_id):
            p_pr = self._p_pr()
            if value in ("", "Normal"):
                p_pr.p_style = None
                return
            p_pr.p_style = el.pStyle(val=value)
            p_pr.p_style.parent = p_pr

    @property
    def style(self) -> str:
        """The style's display name (``"Heading 1"``), as Office JS reports it.

        Read from the styles part's ``w:name`` when that part is already
        unmarshalled, and derived from the id otherwise; **a read never
        unmarshals it** (CR-003 section 4). Setting accepts a display name, the
        name Word stores, or an id, and refuses a name the document does not
        define unless it is a built-in style, whose definition it **adds** to
        ``/word/styles.xml`` (CR-003 section 14.9): Word renders a dangling id
        as Normal, so writing one would lose the style silently.
        """
        return style_name_of(self.parent_body.package, self.style_id)

    @style.setter
    def style(self, value: str) -> None:
        with self.formatting("style", lambda: self.style):
            self.style_id = style_id_of(
                self.parent_body.package, value, validate=True, define=True
            )

    @property
    def style_built_in(self) -> str:
        """The ``Word.Style`` value (``"Heading1"``), or ``"Other"``, as Office JS."""
        return built_in_of(self.style_id)

    @style_built_in.setter
    def style_built_in(self, value: str) -> None:
        with self.formatting("style_built_in", lambda: self.style_built_in):
            self.style_id = define_built_in(self.parent_body.package, id_of_built_in(value))

    # -- paragraph properties ---------------------------------------------

    def _p_pr(self) -> Any:
        """The ``w:pPr``, created and linked when absent. Every setter goes here.

        The **one** accessor every paragraph-property setter shares, which is
        what makes it the place a tracked write records ``w:pPrChange`` with the
        properties as they stood before it (CR-003 section 3.8). It is recorded
        once: a second write keeps the first original. A paragraph this author
        inserted records nothing --- its properties are this author's too ---
        unless it is the one an append left holding the container's old final
        mark, whose properties are that mark's (check 29, CR-010 section 12).
        """
        p_pr = self.element.p_pr
        if p_pr is None:
            p_pr = el.pPr()
            self.element.p_pr = p_pr
            p_pr.parent = self.element
        tracker = self.change_tracker
        if tracker is not None:
            if not tracker.inserted(self.element):
                tracker.record_p_pr_change(p_pr)
            else:
                from docx4j_py.model.content.tracking import holds_final_mark

                if holds_final_mark(self.element, self.container):
                    tracker.record_p_pr_change(p_pr)
        return p_pr

    def _ind(self) -> Any:
        p_pr = self._p_pr()
        if p_pr.ind is None:
            p_pr.ind = el.ind()
            p_pr.ind.parent = p_pr
        return p_pr.ind

    def _spacing(self) -> Any:
        p_pr = self._p_pr()
        if p_pr.spacing is None:
            p_pr.spacing = el.spacing()
            p_pr.spacing.parent = p_pr
        return p_pr.spacing

    def _read_p_pr(self, *, direct: bool = False) -> Any:
        """The ``w:pPr`` a **getter** reads: the effective one unless asked otherwise.

        Every setter goes through :meth:`_p_pr` instead, which is the
        paragraph's own and is where a write belongs.
        """
        return self.element.p_pr if direct else self.effective_p_pr

    @property
    def _ind_or_none(self) -> Any:
        p_pr = self._read_p_pr()
        return getattr(p_pr, "ind", None) if p_pr is not None else None

    @property
    def _spacing_or_none(self) -> Any:
        p_pr = self._read_p_pr()
        return getattr(p_pr, "spacing", None) if p_pr is not None else None

    @property
    def alignment(self) -> AlignmentValue:
        """Office JS ``Word.Alignment``; ``"Unknown"`` when there is no ``w:jc``.

        ``start`` and ``end`` map like ``left`` and ``right``, so a document
        written by a right-to-left Word reads the same (CR-003 section 3.12).
        """
        p_pr = self._read_p_pr()
        value = _value(getattr(p_pr, "jc", None)) if p_pr is not None else None
        return _JC_TO_ALIGNMENT.get(value or "", "Unknown")  # type: ignore[return-value]

    @alignment.setter
    def alignment(self, value: AlignmentValue) -> None:
        with self.formatting("alignment", lambda: self.alignment):
            p_pr = self._p_pr()
            jc = _ALIGNMENT_TO_JC.get(str(value))
            if jc is None:
                p_pr.jc = None
                return
            p_pr.jc = el.jc(val=jc)
            p_pr.jc.parent = p_pr

    def _indent(self, transitional: str, strict: str) -> float:
        ind = self._ind_or_none
        if ind is None:
            return 0.0
        value = getattr(ind, transitional, None)
        if value is None:
            value = getattr(ind, strict, None)
        return 0.0 if value is None else float(value) / TWIPS_PER_POINT

    def _set_indent(self, transitional: str, strict: str, points: float) -> None:
        ind = self._ind()
        twips = round(points * TWIPS_PER_POINT)
        # Write whichever name the paragraph already uses; the transitional one
        # otherwise, because that is what Word writes (CR-003 section 3.12).
        if getattr(ind, transitional, None) is None and getattr(ind, strict, None) is not None:
            setattr(ind, strict, twips)
        else:
            setattr(ind, transitional, twips)

    @property
    def left_indent(self) -> float:
        """The left indent in points (``w:ind/@w:left`` or ``@w:start``)."""
        return self._indent("left", "start")

    @left_indent.setter
    def left_indent(self, points: float) -> None:
        self._set_indent("left", "start", points)

    @property
    def right_indent(self) -> float:
        """The right indent in points (``w:ind/@w:right`` or ``@w:end``)."""
        return self._indent("right", "end")

    @right_indent.setter
    def right_indent(self, points: float) -> None:
        self._set_indent("right", "end", points)

    @property
    def first_line_indent(self) -> float:
        """The first-line indent in points; negative for a hanging indent."""
        ind = self._ind_or_none
        if ind is None:
            return 0.0
        if ind.hanging is not None:
            return -float(ind.hanging) / TWIPS_PER_POINT
        return 0.0 if ind.first_line is None else float(ind.first_line) / TWIPS_PER_POINT

    @first_line_indent.setter
    def first_line_indent(self, points: float) -> None:
        ind = self._ind()
        if points < 0:
            ind.hanging = round(-points * TWIPS_PER_POINT)
            ind.first_line = None
        else:
            ind.first_line = round(points * TWIPS_PER_POINT)
            ind.hanging = None

    @property
    def space_before(self) -> float:
        """The space above the paragraph in points (``w:spacing/@w:before``)."""
        spacing = self._spacing_or_none
        value = getattr(spacing, "before", None) if spacing is not None else None
        return 0.0 if value is None else float(value) / TWIPS_PER_POINT

    @space_before.setter
    def space_before(self, points: float) -> None:
        self._spacing().before = round(points * TWIPS_PER_POINT)

    @property
    def space_after(self) -> float:
        """The space below the paragraph in points (``w:spacing/@w:after``)."""
        spacing = self._spacing_or_none
        value = getattr(spacing, "after", None) if spacing is not None else None
        return 0.0 if value is None else float(value) / TWIPS_PER_POINT

    @space_after.setter
    def space_after(self, points: float) -> None:
        self._spacing().after = round(points * TWIPS_PER_POINT)

    @property
    def line_spacing(self) -> float:
        """Line spacing in points; ``0`` when the paragraph sets none.

        ``w:line`` is 240ths of a line when ``w:lineRule`` is ``auto`` (or
        absent), which Office JS reports as the equivalent of 12-point lines,
        and twips otherwise. Setting writes ``w:lineRule="exact"``.
        """
        spacing = self._spacing_or_none
        line = getattr(spacing, "line", None) if spacing is not None else None
        if line is None:
            return 0.0
        rule = getattr(spacing, "line_rule", None)
        rule = getattr(rule, "value", rule)
        if rule in (None, "auto"):
            return float(line) / 240 * 12
        return float(line) / TWIPS_PER_POINT

    @line_spacing.setter
    def line_spacing(self, points: float) -> None:
        spacing = self._spacing()
        spacing.line = round(points * TWIPS_PER_POINT)
        spacing.line_rule = "exact"

    @property
    def outline_level(self) -> int:
        """``w:outlineLvl + 1``, 1 to 9; ``10`` for body text, as Office JS."""
        p_pr = self._read_p_pr()
        value = _value(getattr(p_pr, "outline_lvl", None)) if p_pr is not None else None
        return 10 if value is None else int(value) + 1

    @outline_level.setter
    def outline_level(self, level: int) -> None:
        p_pr = self._p_pr()
        if level >= 10 or level < 1:
            p_pr.outline_lvl = None
            return
        p_pr.outline_lvl = el.outlineLvl(val=level - 1)
        p_pr.outline_lvl.parent = p_pr

    # -- lists (CR-003 section 3.10, Phase H) ------------------------------

    @property
    def is_list_item(self) -> bool:
        """Office JS ``isListItem``: whether Word paints a list label for this.

        True for a paragraph the **emulator** numbers, so a paragraph numbered
        through its paragraph style is one too, and a ``w:numPr`` naming a
        ``w:numId`` of 0, or one the document does not define, is not.
        """
        from docx4j_py.model.content.lists import is_list_item

        return is_list_item(self)

    @property
    def list(self) -> Any:
        """Office JS ``list``: the :class:`~docx4j_py.model.content.List` this is in, or None."""
        from docx4j_py.model.content.lists import list_of

        return list_of(self)

    @property
    def list_item(self) -> Any:
        """Office JS ``listItem``: the :class:`~docx4j_py.model.content.ListItem`, or None.

        ``None`` rather than Office JS's ``listItemOrNullObject`` distinction:
        CR-003 section 3.1 has no null objects.
        """
        from docx4j_py.model.content.lists import list_item_of

        return list_item_of(self)

    def start_new_list(self, *, kind: str = "Number", like: Any = None) -> Any:
        """Office JS ``startNewList``: a new list with this paragraph as its first item.

        Args:
            kind: ``"Number"`` (the default) or ``"Bullet"``; which of docx4j's
                two default definitions the new ``w:abstractNum`` is copied from.
                Extension: Office JS's ``startNewList`` always makes a numbered
                list.
            like: an existing list to restart --- a new ``w:num`` over the same
                ``w:abstractNum`` with a ``w:startOverride`` of 1, which is what
                Word's *Restart numbering at 1* writes. Extension.

        Raises:
            ContentError: this paragraph is already a list item.
        """
        from docx4j_py.model.content.lists import start_new_list

        return start_new_list(self, kind=kind, like=like)  # type: ignore[arg-type]

    def attach_to_list(self, list_id: int, level: int = 0) -> None:
        """Office JS ``attachToList``: make this paragraph an item of an existing list.

        Writes ``w:numPr`` through :meth:`_p_pr`, so with tracking on it records
        a ``w:pPrChange`` --- what Word writes for a numbering change (CR-003
        section 16.7).
        """
        from docx4j_py.model.content.lists import attach_to_list

        attach_to_list(self, list_id, level)

    def detach_from_list(self) -> None:
        """Office JS ``detachFromList``: this paragraph is no longer a list item."""
        from docx4j_py.model.content.lists import detach_from_list

        detach_from_list(self)

    # -- formatting --------------------------------------------------------

    @property
    def font(self) -> Font:
        """The formatting of the runs (CR-003 section 4: over the runs).

        Reads report the **effective** formatting --- the document defaults,
        the paragraph style, the run's character style and its direct
        formatting, through
        :class:`~docx4j_py.model.property_resolver.PropertyResolver` --- of
        **every run of the paragraph's text**, as Office JS does (CR-016 Phase
        B; core-ts ``getFont`` at ``222cc20``, its CR-002 section 59): the runs
        of ``get_range("Content")``, which are :attr:`text`'s in Word's default
        view, so no deleted run, no hidden run and no equation's math runs. A
        property where they differ reads ``None``, ``""`` (the name, the colour,
        the highlight) or ``"Mixed"`` (the underline). With no such run, the
        runs of a control showing its placeholder, which Word's font reads
        though its text doesn't (core-ts check 52 case 13); with none of those,
        the paragraph mark's run properties (core-ts check 52 case 60). An
        automatic colour reads ``#000000``. Writes go to every run and are
        direct as they always were; a read splits nothing. :meth:`get_font`
        with ``direct=True`` reads the runs' own ``w:rPr`` by the same rule.
        """
        return self.get_font()

    def get_font(self, *, direct: bool = False) -> Font:
        """:attr:`font`, or the **direct** formatting when `direct` is given.

        CR-002 section 15.1: a caller reproducing markup wants what the runs
        themselves say, which is what these reads were before the resolver
        landed.
        """
        from docx4j_py.model.content.font import mark_run_properties

        return Font(
            lambda: runs_of(self.element),
            scope="paragraph",
            record=self.formatting,
            tracking=self.font_tracking,
            context=self.resolution_context,
            direct=direct,
            guard=lambda: self.refuse_whole_text("a font write over the paragraph"),
            reads=self._font_runs,
            mark=lambda: mark_run_properties(self.element),
        )

    def _font_runs(self) -> list[Any]:
        """The runs :attr:`font` reads: its text's, else a placeholder's (core-ts ``getFont``'s ``holders(false)``).

        The runs ``get_range("Content").runs`` gives --- each run holding a
        segment of :attr:`text`, once, in order, an equation's math runs left
        out --- read from one walk of the segments.
        """
        runs: list[Any] = []
        seen: set[int] = set()
        for segment in self.segments():
            if segment.math or segment.end <= 0 or id(segment.run) in seen:
                continue
            seen.add(id(segment.run))
            runs.append(segment.run)
        if runs:
            return runs
        return [run for run in runs_of(self.element) if in_showing_placeholder(run)]

    def resolution_context(self) -> tuple[Any, Any]:
        """The package's property resolver and this paragraph's ``w:pPr``.

        ``(None, None)`` where there is no package --- a detached body --- so
        that every view over it reads direct formatting and nothing raises.
        """
        return resolver_of(self.parent_body.package), self.element.p_pr

    @property
    def effective_p_pr(self) -> Any:
        """The ``w:pPr`` which actually applies, or the paragraph's own.

        The resolved object is **shared and cached** by the resolver when the
        paragraph states no direct formatting: copy it before changing it.
        """
        resolver = resolver_of(self.parent_body.package)
        p_pr = self.element.p_pr
        return p_pr if resolver is None else resolver.get_effective_ppr(p_pr)

    @property
    def effective_paragraph_mark_r_pr(self) -> Any:
        """The ``w:rPr`` of the paragraph mark which actually applies, or its own."""
        resolver = resolver_of(self.parent_body.package)
        p_pr = self.element.p_pr
        if resolver is None:
            return getattr(p_pr, "r_pr", None)
        return resolver.get_effective_paragraph_mark_rpr(p_pr)

    def font_tracking(self) -> Any:
        """What :class:`Font` needs to record ``w:rPrChange``, or None (section 3.8).

        The tracker and the runs that are **this author's own insertion**: a run
        inside a ``w:ins`` of ours takes the formatting outright, because there
        is nothing in the document yet to record as the original.
        """
        from docx4j_py.model.content.tracking import FontTracking, revision_of

        tracker = self.change_tracker
        if tracker is None:
            return None
        own = {
            id(run)
            for run in runs_of(self.element)
            if tracker.own_insertion(revision_of(run))
        }
        return FontTracking(tracker, frozenset(own))

    @property
    def change_tracker(self) -> Any:
        """The package's tracker while ``change_tracking_mode`` is on, else None; its body's, bound to its part."""
        return self.parent_body.change_tracker

    # -- the stable handles ------------------------------------------------

    @property
    def para_id(self) -> str | None:
        """``w14:paraId``, the stable address Word gives paragraphs (section 3.4)."""
        return self.element.para_id

    @para_id.setter
    def para_id(self, value: str | None) -> None:
        self.element.para_id = value

    @property
    def address(self) -> str:
        """The stable handle: ``"w14:5A2B1C3D"`` when there is a paraId, else the ordinal.

        CR-003 section 3.4's paraId-first rule. A paraId survives every edit,
        an insert before this paragraph included; an ordinal does not, and a
        :class:`~docx4j_py.model.content.reports.ChangeReport` says when one
        moved.
        """
        para_id = self.element.para_id
        return para_id_address(para_id) if para_id else self.ordinal

    @property
    def ordinal(self) -> str:
        """The ordinal address (``"body/3"``, ``"body/4/0/1/0"``), always.

        ``""`` for a paragraph that is no longer in its body.
        """
        return ordinal_of(self.parent_body, self.element) or ""

    @property
    def parent_table_cell(self) -> Any:
        """The cell this paragraph is in, or None (Office JS ``parentTableCell``).

        The walk is up the parent pointers and goes through a cell-level
        ``w:sdt`` without naming it, so a paragraph inside an OpenDoPE repeat
        answers with the cell Word shows.
        """
        from docx4j_py.model.content.table import cell_of

        return cell_of(self.element, self.parent_body)

    @property
    def table_nesting_level(self) -> int:
        """How many tables deep the paragraph is (Office JS ``tableNestingLevel``).

        0 outside a table, 1 in a cell, 2 in a nested table's cell, counted
        through :attr:`parent_table_cell` (CR-013 Phase B; core-ts's at
        ``ef57729``).
        """
        level = 0
        cell = self.parent_table_cell
        for _ in range(256):
            if cell is None:
                break
            level += 1
            cell = cell.parent_table.parent_table_cell
        return level

    @property
    def parent_content_control(self) -> Any:
        """The innermost content control this paragraph is in, or None."""
        from docx4j_py.model.content.controls import parent_control_of

        return parent_control_of(self.element, self.parent_body)

    @property
    def content_controls(self) -> list[Any]:
        """The run-level controls in this paragraph, in order, nested ones included."""
        from docx4j_py.model.content.controls import controls_in_paragraph

        return controls_in_paragraph(self)

    def get_content_controls(self, *, types: Any = None) -> list[Any]:
        """Office JS ``getContentControls`` (WordApi 1.5): :attr:`content_controls`, those of `types` where given.

        `types` are ``Word.ContentControlType`` values (``"PlainText"``,
        ``"RichText"``, ...), as a control's ``type`` reports them. Word's
        ``contentControls`` lists rich-text controls only, where this package's
        lists every kind, a recorded departure (CR-013 section 11, G1).
        """
        from docx4j_py.model.content.controls import controls_of_types

        return controls_of_types(self.content_controls, types)

    @property
    def inline_pictures(self) -> list[Any]:
        """The inline pictures in this paragraph, in order (Office JS ``inlinePictures``)."""
        from docx4j_py.model.content.picture import pictures_of

        return pictures_of(self)

    def insert_inline_picture(
        self,
        data: bytes,
        *,
        location: TextLocation = "End",
        width: float | None = None,
        height: float | None = None,
        alt_text_description: str = "",
        alt_text_title: str | None = None,
        name: str | None = None,
    ) -> Any:
        """A picture in this paragraph (Office JS ``insertInlinePicture``).

        The image part, the relationship and the sizing are
        :meth:`Body.insert_inline_picture`'s, and so is what a
        :meth:`~docx4j_py.model.content.trial.dry_run` does with the part it
        adds; the difference is where the run goes.

        Args:
            data: the image's bytes; ``insert_inline_picture_from_base64`` is
                the base64 twin.
            location: ``"Start"``, ``"End"`` (the default) or ``"Replace"``,
                which puts the picture in place of the paragraph's content.
            width: the width in points; the height follows the aspect ratio.
            height: the height in points.
            alt_text_description: ``wp:docPr/@descr``, Word's alt text.
            alt_text_title: ``wp:docPr/@title``; not written when None.
            name: ``wp:docPr/@name``; ``"Picture N"`` by default.
        """
        from docx4j_py.model.content.picture import insert_picture_into_paragraph

        return insert_picture_into_paragraph(
            self,
            data,
            location=location,
            width=width,
            height=height,
            alt_text_description=alt_text_description,
            alt_text_title=alt_text_title,
            name=name,
        )

    def insert_inline_picture_from_base64(
        self, base64: str, *, location: TextLocation = "End", **options: Any
    ) -> Any:
        """:meth:`insert_inline_picture` from base64 (Office JS's own spelling)."""
        import base64 as base64_module

        return self.insert_inline_picture(
            base64_module.b64decode(base64), location=location, **options
        )

    def insert_ooxml(self, ooxml: str, *, location: str = "After") -> list[Any]:
        """Word's ``insertOoxml`` beside or into this paragraph (CR-003 section 4).

        A flat OPC ``pkg:package`` or a bare fragment, as
        :meth:`Body.insert_ooxml`; the paragraph-level rule is
        :meth:`insert_xml`'s, verbatim --- a fragment of exactly one ``w:p``
        inserted at ``"Start"`` or ``"End"`` has its **runs merged into this
        paragraph**, as Word's paste does.

        Args:
            ooxml: the ``pkg:package`` document, or the fragment.
            location: ``"Before"``, ``"After"`` (the default), ``"Start"``,
                ``"End"`` or ``"Replace"``.

        Returns:
            The views of what was inserted --- ``[self]`` when the runs were
            merged in.
        """
        from docx4j_py.model.content.ooxml import insert_ooxml_into_paragraph

        return insert_ooxml_into_paragraph(self, ooxml, location=location)

    # -- editing -----------------------------------------------------------

    def insert_text(self, text: str, *, location: TextLocation = "End") -> Range:
        """Insert text at the start, the end, or in place of the whole text.

        ``\\n``, ``\\r`` and ``\\r\\n`` each make a new paragraph, as
        Word's ``insertText`` does (CR-013 section 11, E1; :meth:`splice`): this
        view goes on viewing the paragraph that keeps the mark, the last.
        Tracked, a ``"Replace"`` writes the ``w:ins`` first and the ``w:del``
        after it, as Word does (CR-013 Phase C; :meth:`Range.insert_text`).

        Args:
            text: what to insert.
            location: ``"Start"``, ``"End"`` (the default) or ``"Replace"``.

        Returns:
            The :class:`~docx4j_py.model.content.Range` of the inserted text; where
            it made paragraphs, a range across them, as Office JS's (CR-013
            Phase B).
        """
        with recording(self.parent_body, "insert_text") as change:
            before = self.text
            length = len(before)
            if location in ("Start", "End", "Replace"):
                start, end = {
                    "Start": (0, 0),
                    "End": (length, length),
                    "Replace": (0, length),
                }[location]
                index = _index_of(self.container, self.element)
                if location == "Replace" and self.change_tracker is not None:
                    # a tracked replacement taking an equation is refused, unmeasured, as
                    # core-ts's Paragraph.insertText refuses it (CR-013 section 13.7, D2)
                    self.refuse_equation(start, end)
                    # tracked, the w:ins first and the w:del after it, as Word writes a
                    # replacement (CR-013 Phase C; docx4j-core-ts check 54 case 4)
                    out = self.get_range("Content").insert_text(text, location="Replace")
                elif location == "Replace":
                    # the whole content: its hidden runs and pictures with it (core-ts's WHOLE)
                    out = self.splice(start, end, text, whole=True)
                else:
                    # the paragraph's own edge: before its leading hidden runs, after its trailing ones
                    out = self.splice(start, end, text, edge="start" if location == "Start" else "end")
                change.touched(self)
                change.text(before=before, after=self.text_since(index))
                return out
        raise ContentError(
            f"insert_text at paragraph level takes Start, End or Replace, not {location!r}",
            code="location.invalid",
            hint="use Before or After on insert_paragraph, or a Range for a span",
        )

    def insert_paragraph(
        self,
        text: str = "",
        *,
        location: ParagraphLocation = "After",
        style: str | None = None,
    ) -> Paragraph:
        """A new paragraph before or after this one, with this one's properties.

        Word copies the paragraph properties when it splits a paragraph, minus
        the mark's run properties, the section break and any ``w:pPrChange``;
        so does this.

        Args:
            text: the text of the new paragraph; ``""`` leaves it empty. One
                paragraph whatever it holds (CR-013 section 11, E2): ``\\n`` and
                ``\\r\\n`` a ``w:br``, ``\\r`` a ``w:cr``, each other of Office
                JS's characters its element (Q2), and U+0002 and the other
                characters XML 1.0 excludes dropped (E3).
            location: ``"Before"`` or ``"After"`` (the default).
            style: a style name or id for the new paragraph.
        """
        text = drop_unwritable(text)
        with recording(self.parent_body, "insert_paragraph") as change:
            new = P(content=ChildList([_text_run(text)] if text else []))
            if self.element.p_pr is not None:
                p_pr = deep_copy(self.element.p_pr)
                # the mark's own run properties, the section break and any recorded
                # property change belong to this paragraph, not to the new one
                p_pr.r_pr = None
                p_pr.sect_pr = None
                p_pr.p_pr_change = None
                new.p_pr = p_pr
                link_parents(new)
                p_pr.parent = new
            view = self.parent_body.insert_element(new, location=location, target=self)
            if style is not None:
                view.style = style
            change.text(after=text)
            return view  # type: ignore[return-value]

    def insert_break(
        self,
        type: BreakTypeValue = "Page",
        *,
        location: str = "End",
    ) -> None:
        """Insert a break (Office JS ``Word.Paragraph.insertBreak``).

        Args:
            type: ``"Page"`` or ``"Line"``.
            location: ``"Start"`` or ``"End"`` puts the break in this
                paragraph; ``"Before"`` or ``"After"`` puts it in a new one.
        """
        with recording(self.parent_body, "insert_break") as change:
            item = br("page" if type == "Page" else None)
            tracker = self.change_tracker
            if location in ("Before", "After"):
                new = self.insert_paragraph("", location=location)  # type: ignore[arg-type]
                new.element.content.append(R(content=ChildList([item])))
                link_parents(new.element)
                if tracker is not None:
                    from docx4j_py.model.content.tracking import wrap_new_runs

                    wrap_new_runs(tracker, new.element)
                return
            run: Any = R(content=ChildList([item]))
            if tracker is not None:
                link_parents(run)
                run = tracker.ins([run])
            if location == "Start":
                self.element.content.insert(0, run)
            elif location == "End":
                self.element.content.append(run)
            else:
                raise ContentError(
                    f"insert_break takes Start, End, Before or After, not {location!r}",
                    code="location.invalid",
                    hint="Start and End put the break in this paragraph",
                )
            link_parents(self.element)
            change.touched(self)

    def insert_xml(self, xml: str, *, location: str = "After") -> list[Any]:
        """Insert a WordprocessingML fragment (CR-003 section 4).

        A fragment of exactly one ``w:p`` inserted at ``"Start"`` or ``"End"``
        has its **runs merged into this paragraph**, as Word's paste does;
        anything else is inserted as blocks before or after it. ``"Replace"``
        puts the content where this paragraph is and deletes it.

        Args:
            xml: one or more block-level elements, written as they appear
                inside ``document.xml``; the prefixes are declared for you.
            location: ``"Before"``, ``"After"`` (the default), ``"Start"``,
                ``"End"`` or ``"Replace"``.

        Returns:
            The views of what was inserted --- this paragraph, when the runs
            were merged into it.
        """
        with recording(self.parent_body, "insert_xml") as change:
            elements = wml.all(xml, wrapper="body")
            if not elements:
                return []
            only = elements[0] if len(elements) == 1 and isinstance(elements[0], P) else None
            if only is not None and location in ("Start", "End"):
                before = self.text
                self.insert_items_at(
                    0 if location == "Start" else len(self.text), list(only.content)
                )
                change.touched(self)
                change.text(before=before, after=self.text)
                return [self]
            if location == "Replace":
                if self.change_tracker is not None:
                    self.refuse_whole_text("insert_xml")
                self.parent_body.insert_element(elements, location="Before", target=self)
                views = [self.parent_body.view_for(e) for e in elements]
                self.delete()
                return views
            where = "Before" if location == "Start" else "After" if location == "End" else location
            self.parent_body.insert_element(elements, location=where, target=self)
            return [self.parent_body.view_for(e) for e in elements]

    def insert_markdown(self, markdown: str, *, location: str = "After") -> list[Any]:
        """Insert markdown next to this paragraph (CR-003 section 3.5).

        ``"Before"`` and ``"After"`` (the default) put the blocks beside this
        paragraph. ``"Start"`` and ``"End"`` are accepted as well and follow
        :meth:`insert_xml`'s rule of CR-003 section 4: a fragment that is
        exactly one paragraph has its **runs merged into this one**, as Word's
        paste does; anything else goes before or after.

        What the call may touch beyond this part, and what it does with an
        image or an HTML block, is :meth:`Body.insert_markdown`'s docstring.

        Returns:
            The inserted views --- ``[self]`` when the runs were merged in.
        """
        from docx4j_py.model.markdown import blocks_for

        with recording(self.parent_body, "insert_markdown") as change:
            elements, touched = blocks_for(self.parent_body, markdown, change)
            parts = getattr(change, "parts", None)
            if parts is not None:
                for name in sorted(touched):
                    if name not in parts:
                        parts.append(name)
            if not elements:
                return []
            only = elements[0] if len(elements) == 1 and isinstance(elements[0], P) else None
            if only is not None and location in ("Start", "End"):
                before = self.text
                self.insert_items_at(
                    0 if location == "Start" else len(self.text), list(only.content)
                )
                change.touched(self)
                change.text(before=before, after=self.text)
                return [self]
            where = "Before" if location == "Start" else "After" if location == "End" else location
            self.parent_body.insert_element(elements, location=where, target=self)
            return [self.parent_body.view_for(element) for element in elements]

    def to_markdown(self, *, addresses: bool = False, view: str = "accepted") -> str:
        """This paragraph as Markdown: its lines in its story's, the text model's (CR-011 Phase D).

        Its block form as :meth:`Body.to_markdown` writes it: a heading is
        ``# ...``, a list item has its marker and indent (``3. ...``,
        ``4. 2. ...``), a quote ``> ...``; a paragraph in a table cell is
        written on its own. A footnote reference is ``[^1]``, but the
        definitions belong to the document, so only :meth:`Body.to_markdown`
        writes them. Trailing white space is left out.

        Args:
            addresses: put :attr:`address` in an HTML comment on its own line
                before the paragraph.
            view: ``"accepted"``, ``"original"`` or ``"markup"`` (CriticMarkup).
        """
        from docx4j_py.model.markdown import paragraph_markdown

        return paragraph_markdown(self, addresses=addresses, view=view)

    def search(self, text: str, **options: Any) -> list[Range]:
        """Every match of `text` in this paragraph, as ranges.

        Matches are found in Office JS's text, not the text model's (CR-011 Q4; WD REQ-202's exemption for a reader of an
        emulated API): :func:`docx4j_py.model.text.extract` is the text model, and each of its runs
        records its ``office_offset`` in this text, which reads Word's branch of a run- or
        paragraph-level ``mc:AlternateContent`` (Q4 revised, CR-011 section 15).

        Matches are found on the paragraph's whole text, so one **spans runs**
        freely; only formatting a match splits them. Office JS's rules (CR-013
        Phase A; :func:`~.text_model.office_matches`): hidden text is not there,
        so a query matches across it; a picture is U+FFFC, which ``^g`` matches;
        an optional hyphen is passed over unless the query names it (``^-``);
        Word's Find codes ``^f ^e ^l ^m ^t ^~ ^- ^g ^^``; a plain ``-`` matches
        a non-breaking hyphen too.

        Args:
            text: what to look for.
            **options: ``match_case``, ``match_whole_word``, ``match_wildcards``
                and ``limit``, as :meth:`Body.search`.
        """
        from docx4j_py.model.content.range import Range

        if not options.get("match_wildcards") and paragraph_mark_pieces(text) is not None:
            # Word's paragraph mark, ^p, is searched in the paragraph's whole range, which
            # holds its mark (CR-013 Phase C; core-ts CR-014 phase 3)
            return self.get_range("Whole").search(text, **options)
        limit = options.pop("limit", None)
        hits = office_matches(self.segments(objects=True), text, **options)
        if limit is not None:
            hits = hits[:limit]
        return [Range(self, start, end) for start, end in hits]

    def replace_text(self, find: str, replace: str, **options: Any) -> int:
        """Replace every match, last first so the offsets stay valid.

        ``search`` then ``insert_text(replace, location="Replace")`` from the
        last match to the first, so that the offsets of the ones still to do
        stay valid --- **tracked** when the package's ``change_tracking_mode``
        is on, in which case each replacement is a ``w:del`` and a ``w:ins``.
        The report names the pair rather than the text either side, because one
        call may touch many paragraphs (CR-003 section 16). A `replace` holding
        ``\\n``, ``\\r`` or ``\\r\\n`` makes a paragraph of each, as
        :meth:`insert_text` does (CR-013 section 11, E1); every match's split is
        checked with the rest before the first replacement.

        Args:
            find: what to look for.
            replace: what to put there.
            **options: ``match_case``, ``match_whole_word``,
                ``match_wildcards`` and ``limit``, as :meth:`search`.

        Returns:
            How many were replaced.

        Raises:
            SpanError: ``range.replaces_mark``, a query holding ``^p``, Word's
                paragraph mark, whose replacement is unmeasured (CR-013 Phase C).
        """
        refuse_mark_replacement("Paragraph", find, options)
        replace = drop_unwritable(replace)
        with recording(self.parent_body, "replace_text") as change:
            change.text(before=find)
            matches = self.search(find, **options)
            # every match checked before the first is replaced (CR-011 Phase F),
            # a paragraph break's split among the checks (CR-013 section 11, E1)
            segments = self.segments()
            tracked = self.change_tracker is not None
            for match in matches:
                refuse_branch_text(self.element, match.start, match.end, "replace_text", segments=segments)
                refuse_equation(segments, match.start, match.end, tracked=tracked)
                self.refuse_paragraph_breaks(match.start, match.end, replace, "replace_text")
            replace_matches(matches, replace)
            if matches:
                change.touched(self)
            change.text(after=replace)
            return len(matches)

    def find(self, text: str, *, context: int = 40, limit: int = 20, **options: Any) -> list[Any]:
        """Every match in this paragraph, as hits with addresses and context.

        Matches are found in Office JS's text, not the text model's (CR-011 Q4; WD REQ-202's exemption for a reader of an
        emulated API): :func:`docx4j_py.model.text.extract` is the text model, and each of its runs
        records its ``office_offset`` in this text, which reads Word's branch of a run- or
        paragraph-level ``mc:AlternateContent`` (Q4 revised, CR-011 section 15).

        :class:`~docx4j_py.model.content.reports.SearchHit`\\ s for a tool
        result, where :meth:`search` gives
        :class:`~docx4j_py.model.content.range.Range`\\ s for code.
        """
        from docx4j_py.model.content.reports import hit_for

        hits = office_matches(self.segments(objects=True), text, **options)
        return [hit_for(self, start, end, context=context) for start, end in hits[:limit]]

    # -- comments (CR-003 section 3.9, Phase G) -----------------------------

    def get_comments(self) -> list[Any]:
        """The comments anchored in this paragraph (Office JS ``getComments``)."""
        from docx4j_py.model.content.comments import comments_of

        return comments_of(self)

    def insert_comment(self, text: str) -> Any:
        """Comment on the whole paragraph (Office JS ``insertComment``).

        The paragraph's whole range, so the markers sit around all of its text;
        :meth:`Range.insert_comment` comments on a span. What the call writes,
        and what it creates when the document has no comment parts, is
        :meth:`Range.insert_comment`'s docstring.
        """
        return self.get_range().insert_comment(text)

    # -- change tracking (CR-003 section 3.8, Phase F) ----------------------

    def get_tracked_changes(self) -> list[Any]:
        """The tracked changes in this paragraph, in document order (Office JS).

        The paragraph's own ``w:pPrChange`` and mark revisions first, then the
        run-level ``w:ins``, ``w:del``, ``w:moveFrom``, ``w:moveTo`` and
        ``w:rPrChange`` in document order.
        """
        from docx4j_py.model.content.tracked_change import tracked_changes_of_paragraph

        return tracked_changes_of_paragraph(self)

    def get_range(self, location: RangeLocation = "Whole") -> Range:
        """A :class:`~docx4j_py.model.content.Range` over this paragraph.

        ``"Whole"`` holds the paragraph's mark too, so its ``text`` ends with
        :attr:`mark_text` (``\\r``, or ``\\t`` for a cell's last paragraph), as
        Office JS's does (core-ts check 48; CR-013 Phase A); ``"Content"`` ends
        with the paragraph's text. The mark is no offset: every edit and search
        works on the text before it. Both hold the paragraph's whole content
        (:attr:`Range.whole`, core-ts's ``whole``; CR-013 Phase B).
        ``"After"`` is collapsed at the start of the paragraph after it.
        """
        from docx4j_py.model.content.range import Range

        length = len(self.text)
        if location == "Start":
            return Range(self, 0, 0)
        if location == "End":
            return Range(self, length, length)
        if location == "After":
            return Range(self, 0, length, mark=True).get_range("After")
        return Range(self, 0, length, mark=location == "Whole", whole=True)

    def delete(self) -> None:
        """Remove the paragraph from its container.

        **While the package tracks changes** the paragraph stays: its content
        becomes a ``w:del`` and its mark is marked deleted
        (``w:pPr/w:rPr/w:del``), which is what Word shows until the change is
        accepted --- unless the whole paragraph was this author's own
        insertion, which Word simply takes back.

        **The paragraph a table cell ends with**, as Word reads it, stays where
        removing it would leave the cell without one: emptied, its properties
        kept, and no block shifts (CR-016 Phase A; check 52 cases 88 and 89).
        Tracked, its text is deleted and its mark is never marked deleted, as
        Word keeps a cell's last paragraph (cases 90 and 91).

        Raises:
            TrackedChangeError: tracked, and this paragraph's mark is already
                marked deleted.
        """
        self._delete(keep_cell_end=True)

    def _delete(self, *, keep_cell_end: bool) -> None:
        """:meth:`delete`; with `keep_cell_end` False a cell's last paragraph goes as any other.

        For ``Body.clear``'s callers that put their own paragraph in next
        (core-ts's ``remove``, ``Paragraph.mts`` at ``222cc20``).
        """
        from docx4j_py.model.content.table_rules import (
            ends_cell,
            ends_with_paragraph,
            wrapping_cell_of,
        )

        with recording(self.parent_body, "delete") as change:
            index = self.index
            if index < 0:
                return
            cell_end = keep_cell_end and ends_cell(self.element)
            tracker = self.change_tracker
            if tracker is not None:
                change.touched(self)
                change.text(before=self.text, after="")
                self._tracked_delete(tracker, cell_end=cell_end)
                return
            change.touched(self)
            change.text(before=self.text, after="")
            moved = moved_by_delete(self.parent_body, self.container, index)
            cell = wrapping_cell_of(getattr(self.element, "parent", None)) if cell_end else None
            del self.container[index]
            if cell is not None and not ends_with_paragraph(cell):
                # the cell's last paragraph, emptied, where it was (check 52 cases 88 and 89)
                self.element.content.clear()
                self.container.insert(index, self.element)
                return
            change.shifted(moved)

    def _tracked_delete(self, tracker: Any, *, cell_end: bool = False) -> None:
        """The tracked half of :meth:`delete`: the content, then the mark; a cell's last mark never marked."""
        from docx4j_py.model.content.tracking import (
            mark_deleted,
            prune_paragraph_properties,
        )

        self.refuse_whole_text("delete")
        if mark_deleted(self.element):
            from docx4j_py.model.content.errors import TrackedChangeError

            raise TrackedChangeError(
                "this paragraph's mark is already marked deleted",
                code="tracking.already_deleted",
                hint="accept or reject the deletion before deleting the paragraph again",
            )
        from docx4j_py.model.content.tracking import insertion_mark_of

        length = len(self.text)
        if length or self.holds_objects():
            self._delete_text(tracker, 0, length, whole=True)
        mark = insertion_mark_of(self.element, self.container)
        ours = (
            mark is not None
            and tracker.own_paragraph(mark)
            and (mark is self.element or tracker.inserted(self.element))
        )
        if ours and not self.element.content:
            if cell_end and self._ends_its_cell_alone():
                # a cell's last paragraph stays, whoever inserted it (CR-016 Phase A)
                return
            # our own insertion, taken back whole: the paragraph simply goes, and
            # so does the mark that recorded it --- which is the **previous**
            # paragraph's when this one is the last of its container (16.10)
            if mark is not self.element:
                from docx4j_py.model.content.tracking import hand_back_final_mark

                mark.p_pr.r_pr.ins = None
                prune_paragraph_properties(mark)
                hand_back_final_mark(tracker, self.element, mark)
            index = self.index
            if index >= 0:
                del self.container[index]
            return
        if cell_end:
            # a cell's last mark: Word deletes the paragraph's text and never marks the mark (check 52
            # cases 90 and 91)
            return
        tracker.mark_paragraph_deleted(self.element)

    def _ends_its_cell_alone(self) -> bool:
        """Whether this paragraph's cell, without it, would end with no paragraph, as Word reads it."""
        from docx4j_py.model.content.table_rules import (
            ends_with_paragraph,
            wrapping_cell_of,
        )

        cell = wrapping_cell_of(getattr(self.element, "parent", None))
        index = self.index
        if cell is None or index < 0:
            return False
        del self.container[index]
        try:
            return not ends_with_paragraph(cell)
        finally:
            self.container.insert(index, self.element)

    def get_xml(self) -> str:
        """The paragraph as XML, with docx4j's prefixes. Extension."""
        return to_xml(self.element)

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary: what a tool result says about a paragraph.

        A list item carries ``list_item`` as well: its level, the label Word
        paints and its place among its siblings (CR-003 Phase H).
        """
        out: dict[str, Any] = {
            "address": self.address,
            "ordinal": self.ordinal,
            "text": self.text,
            "style": self.style,
            "style_id": self.style_id,
            "style_built_in": self.style_built_in,
            "alignment": self.alignment,
            "outline_level": self.outline_level,
            "para_id": self.para_id,
            "index": self.index,
            "runs": len(self.runs),
        }
        item = self.list_item
        if item is not None:
            out["list_item"] = item.to_dict()
        return out

    # -- docx4j names ------------------------------------------------------

    def to_api_script(self, **options: Any) -> str:
        """The content-API calls that would produce this paragraph (CR-003 section 3.11).

        Python source against a name ``body``, executable as it stands; see
        :func:`docx4j_py.model.content.api_script.to_api_script` for the
        options (``variable``, ``location``, ``pictures``, ``limit``,
        ``addresses``).
        """
        from docx4j_py.model.content.api_script import to_api_script

        return to_api_script(self, **options)

    def get_content(self) -> list:
        """docx4j: the ``w:p``'s content list."""
        return self.element.content

    # -- the editing primitives, shared with Range -------------------------

    def split_at(self, offset: int, *, prefer: str = "back") -> int:
        """Split the run at `offset` so that ``[offset, ...)`` begins a run.

        The offset is snapped to a grapheme boundary first (CR-003 section
        3.12); the offset actually used is returned.
        """
        made: list = []
        used = split_at(self.element, offset, prefer=prefer, made=made, hidden=self.hidden_runs())
        self.remember_split(made)
        return used

    def remember_split(self, pairs: list) -> None:
        """Tell the package which two runs a split made. Extension, CR-003 16.12.

        `pairs` are ``(head, tail)`` or ``(head, tail, item)``, where `item` is
        the head's ``w:t`` when the split is what put ``xml:space`` on it. A
        rejected revision puts the halves back side by side, and this is what
        says they are **those** halves: runs a document keeps apart are never
        joined, and neither is a half beside a run it never belonged to.
        """
        if not pairs:
            return
        package = self.parent_body.package
        store = getattr(package, "_split_runs", None)
        if store is None:
            return
        spaces = getattr(package, "_split_spaces", None)
        for pair in pairs:
            head, tail = pair[0], pair[1]
            store[id(tail)] = id(head)
            item = pair[2] if len(pair) > 2 else None
            if item is not None and spaces is not None:
                spaces.add(id(item))

    def insert_content_control(self, kind: str = "RichText") -> Any:
        """Wrap this paragraph in a content control, in place.

        Office JS ``Paragraph.insertContentControl`` (CR-003 section 3.7). The
        container is the same length afterwards: the ``w:p`` is inside a
        ``w:sdt`` where it was.
        """
        from docx4j_py.model.customxml.insert import insert_content_control_in_paragraph

        return insert_content_control_in_paragraph(self, kind)

    def splice(
        self,
        start: int,
        end: int,
        text: str,
        *,
        whole: bool | tuple[bool, bool] = False,
        breaks: str = "paragraph",
        edge: str | None = None,
        prefer: str = "before",
    ) -> Range:
        """Replace the text in ``[start, end)`` and return the range of the new text.

        Runs are edited in place: a ``w:t`` loses or gains characters, a tab or
        a break inside the span goes, a run left empty goes. The new text takes
        the place --- and the formatting --- of the first replaced character, as
        Word does; where there is nothing to replace it extends the ``w:t`` at
        the offset, and failing that becomes a new run with the neighbouring
        run's formatting.

        Office JS's characters in `text` are written as the elements they read
        from (CR-013 Q2; :func:`~.text_model.text_items`), each in the run the
        text went into; U+0002 and the other characters XML 1.0 excludes are
        dropped (CR-013 section 11, E3). Hidden text strictly inside the span
        goes with it, as Word takes it (core-ts's rule; CR-013 Phase A); at the
        span's edges only when `whole`, or when the span is the paragraph's
        whole text. An equation is taken whole and untracked, or not at all
        (:func:`~.text_model.refuse_equation`).

        **Paragraph breaks** (CR-013 section 11, E1; docx4j-core-ts check 55,
        cases 2 to 4 and 48 to 49). With `breaks` ``"paragraph"`` (the default)
        each ``\\n`` and each ``\\r`` makes a new paragraph, ``\\r\\n`` one,
        as Word's ``insertText`` does: the text is written without them, and the
        paragraph is split after each piece (:func:`~.split.split_paragraph`).
        Each new paragraph is in front, in a copy of the properties with a mark
        of its own (marked inserted under change tracking); this ``w:p`` keeps
        the last piece, its mark, its section break and its ``w14:paraId``, as
        Word keeps them, so this view goes on viewing the last. A hyperlink, a
        revision or a smart tag the text is in is cut in two; a text that would
        go into a run-level content control or a field is refused
        (``range.splits_control``, ``range.splits_field``) before anything
        changes. Inside a bound block control (no type, ``w:text`` or
        ``w:richText``), which Word reads as plain text, a break is a ``w:br`` in a
        multi-line ``w:text`` one and nothing in any other (docx4j-core-ts CR-002
        section 61, ``80a933d``, widened by its check 56, ``38dcae9``;
        :func:`~.controls.bound_plain_text_control`). With `breaks`
        ``"line"`` (the one-paragraph writers, section 11's E2) ``\\n`` and
        ``\\r\\n`` are a ``w:br`` and ``\\r`` a ``w:cr``.

        `whole` may name the span's edges one by one, ``(start, end)``: what a
        range across paragraphs holds of its first and last paragraph (core-ts's
        ``SpanEdges``; CR-013 Phase C). A comment's reference in the span goes
        with the text, and the comment with it once the write is done: untracked
        the reference run goes, tracked it goes into the ``w:del`` (core-ts
        CR-014 section 12.6, Word's check 54 cases 27 and 28). `prefer`
        ``"after"`` puts a tracked insertion before the run at `start`, in its
        look, where ``"before"`` (the default) puts it after the run ending
        there: a tracked replacement's ``w:ins`` goes before the text it
        replaces (CR-013 Phase C).

        Returns:
            The range of the new text; where it made paragraphs, a range across
            them, from the first new paragraph to this ``w:p``, as Office JS's
            is (core-ts check 55: ``x\\ry``; CR-013 Phase B).
        """
        text = drop_unwritable(text)
        if LINE_CHARACTERS.search(text):
            lines = self._lines_at(start, end, edge)
            if lines is not None:
                where, kind = lines
                if where == "run" and PARAGRAPH_BREAKS.search(text):
                    _refuse_run_control_break()
                if kind == "single":
                    # a single-line plain-text control takes no line at all (F3)
                    text = for_writer(text, "single-line")
                elif kind == "multi" and where == "block":
                    # a multi-line w:text control, bound or not: each break a w:br (section 11's item 2)
                    text = for_writer(text, "multi-line")
                return self._splice_text(start, end, text, whole=whole, edge=edge, prefer=prefer)
        if breaks == "line":
            text = for_writer(text, "one-paragraph")
        elif PARAGRAPH_BREAKS.search(text):
            return self._splice_paragraphs(start, end, text, whole=whole, edge=edge, prefer=prefer)
        return self._splice_text(start, end, text, whole=whole, edge=edge, prefer=prefer)

    def _lines_at(self, start: int, end: int, edge: str | None = None) -> tuple[str, str | None] | None:
        """The content control that decides how a text's lines go where :meth:`splice` puts it: ``(where, kind)``.

        ``("run", kind)`` where the text lands inside a run-level control, `kind`
        its :func:`~.controls.plain_text_lines_of` (None for a rich-text one); else
        ``("block", kind)`` for the block control holding this paragraph directly
        where it takes lines (:func:`~.controls.paragraph_lines_of`); else None.
        core-ts's ``runControlOf`` and ``blockControl`` (``6285dfa``).
        """
        from docx4j_py.model.content.controls import (
            ContentControl,
            paragraph_lines_of,
            plain_text_lines_of,
        )
        from docx4j_py.model.content.split import holders_above
        from docx4j_py.wml import SdtRun

        anchor, _how = self._landing(start, end, edge)
        if anchor is not None:
            for holder in holders_above(anchor, self.element):
                if isinstance(holder, SdtRun):
                    owner = getattr(holder, "parent", None)
                    container = run_items_of(owner) if owner is not None else None
                    control = ContentControl(holder, container or [holder], self.parent_body)
                    return "run", plain_text_lines_of(control)
        kind = paragraph_lines_of(self)
        return ("block", kind) if kind is not None else None

    def _splice_paragraphs(
        self,
        start: int,
        end: int,
        text: str,
        *,
        whole: bool | tuple[bool, bool],
        edge: str | None = None,
        prefer: str = "before",
    ) -> Range:
        """:meth:`splice` of a text holding paragraph breaks: the text without them, then a split after each piece."""
        from docx4j_py.model.content.range import Range
        from docx4j_py.model.content.reports import current_recorder
        from docx4j_py.model.content.split import split_paragraph

        verb = getattr(current_recorder(getattr(self.parent_body, "package", None)), "operation", None) or "this edit"
        pieces = PARAGRAPH_BREAKS.split(text)
        joined = "".join(pieces)
        length = len(self.text)
        start = max(0, min(start, length))
        end = max(start, min(end, length))
        self.refuse_paragraph_breaks(start, end, text, verb, edge=edge, prefer=prefer)
        out = self._splice_text(start, end, joined, whole=whole, edge=edge, prefer=prefer)
        start = out.start
        edge = start + len(joined)
        consumed = 0
        position = start
        container = self.container
        index = _index_of(container, self.element)
        for piece in pieces[:-1]:
            position += len(piece)
            # after the text's last character the places of no text after it stay with what follows
            keep = "tail" if joined and position == edge else "head"
            split_paragraph(self, position - consumed, keep=keep, operation=verb)
            consumed = position
        # the text's range runs from its first piece, in the first paragraph the splits made, to
        # its last, in this w:p (CR-013 Phase B; core-ts check 55: Word's range reads x\ry)
        if len(pieces) > 1 and 0 <= index < len(container) and container[index] is not self.element:
            first = Paragraph(container[index], container, self.parent_body)
            return Range(first, start, len(pieces[-1]), end_paragraph=self)
        return Range(self, 0, len(pieces[-1]))

    def refuse_paragraph_breaks(
        self,
        start: int,
        end: int,
        text: str,
        operation: str = "this edit",
        *,
        edge: str | None = None,
        prefer: str = "before",
    ) -> None:
        """Refuse, before anything changes, a splice of `text` whose paragraph breaks can't be made there. Extension.

        CR-013 section 11, E1: where :meth:`splice` would put the text --- its
        landing --- must not be inside a run-level content control (core-ts's
        F4) or a field; a text of breaks alone is checked as a split at the
        place. Nothing to check for a text with no break, or in a block control
        that takes lines (a ``w:text`` one, or a bound one Word reads as plain
        text), where a break makes no paragraph (:meth:`_lines_at`).

        Raises:
            SpanError: ``range.splits_control``, ``range.splits_field``,
                ``range.cuts_equation``.
            AddressError: ``address.read_only``.
        """
        from docx4j_py.model.content.split import refuse_place, refuse_split

        text = drop_unwritable(text)
        if not PARAGRAPH_BREAKS.search(text):
            return
        segments = self.segments()
        if any(segment.branch is not None for segment in segments):
            self.refuse_branch_text(start, end, operation)
        refuse_equation(segments, start, end, tracked=self.change_tracker is not None)
        lines = self._lines_at(start, end, edge)
        if lines is not None:
            if lines[0] == "run":
                _refuse_run_control_break(operation)
            return
        if PARAGRAPH_BREAKS.sub("", text):
            anchor, how = self._landing(start, end, edge, prefer=prefer)
            refuse_place(self.element, anchor, how, operation)
        else:
            refuse_split(self, end, keep="head", operation=operation)

    def _landing(self, start: int, end: int, edge: str | None = None, *, prefer: str = "before") -> tuple[Any, str]:
        """Where :meth:`_splice_text` puts a text, read without changing anything: ``(anchor, how)``.

        `how` is ``"into"`` (the text extends `anchor`, a ``w:t``), ``"after"``
        or ``"before"`` (a new run beside `anchor`); `anchor` None is the
        paragraph's own list, ``"start"`` or ``"end"``. It follows
        :meth:`_splice_text` and :meth:`_insert_tracked` rule for rule, so that a
        paragraph break is refused before the text is written.
        """
        from docx4j_py.model.content.tracking import revision_of

        listed = self.segments(objects=True)
        segments = [segment for segment in listed if not segment.object]
        length = segments[-1].end if segments else 0
        start = max(0, min(start, length))
        end = max(start, min(end, length))
        whole = start == 0 and end == length and end > start
        hidden = [segment for segment in listed if segment.object and _in_span(segment.start, start, end, whole)]
        tracker = self.change_tracker
        beside = _hidden_edge(listed, start, end, edge)
        if beside is not None:
            # the paragraph's edge beside its hidden runs (:meth:`_place_at_hidden_edge`)
            segment, how = beside
            site = segment.site()[0]
            revision = revision_of(site)
            return (revision.element if revision is not None else site), how
        if tracker is None:
            if end > start or hidden:
                first = next((s for s in segments if s.start <= start < s.end), None)
                if first is not None and first.editable:
                    return first.item, "into"
                segments = _after_deletion(segments, start, end)
            return _untracked_landing(segments, start)
        if end > start or hidden:
            deleted = [
                s
                for s in listed
                if (s.object and _in_span(s.start, start, end, whole))
                or (not s.object and s.end > start and s.start < end and s.text)
            ]
            others = [s for s in deleted if not tracker.own_insertion(revision_of(s.run))]
            if others:
                # the w:ins follows the last w:del, which is where that text's run was
                return others[-1].item, "after"
            segments = _after_deletion(segments, start, end)
        return _tracked_landing(tracker, segments, start, prefer)

    def _splice_text(
        self,
        start: int,
        end: int,
        text: str,
        *,
        whole: bool | tuple[bool, bool] = False,
        edge: str | None = None,
        prefer: str = "before",
    ) -> Range:
        """:meth:`splice` of a text the paragraph breaks have been taken out of.

        `edge` ``"start"`` or ``"end"`` is the paragraph's own edge (its
        ``insert_text`` at ``"Start"`` or ``"End"``): text put there goes before
        the hidden runs at the start, or after those at the end, not into them
        (CR-013 section 11, the review's item 1). A drawn object in the span goes
        with it, as a hidden run does (core-ts CR-002 section 48).
        """
        from docx4j_py.model.content.range import Range

        listed = self.segments(objects=True)
        segments = [segment for segment in listed if not segment.object]
        length = segments[-1].end if segments else 0
        start = max(0, min(start, length))
        end = max(start, min(end, length))
        whole = _edges_of(whole, start, end, length)
        if segments and any(segment.branch is not None for segment in segments):
            self.refuse_branch_text(start, end)

        tracker = self.change_tracker
        # under tracking a deletion takes an equation whole in Word's form, and a replacement
        # doesn't take one at all (CR-013 section 13.7, D2)
        refuse_equation(segments, start, end, tracked=False if tracker is None else "replace" if text else "delete")
        # the places of no text in the span: hidden runs' items (CR-013 Phase A) and
        # drawn objects (core-ts CR-002 section 48; CR-013 section 11), which go with it
        hidden = [
            segment
            for segment in listed
            if segment.object and _in_span(segment.start, start, end, whole)
        ]
        # a comment's reference in the span goes with the text, and its comment with it
        # (core-ts CR-014 section 12.6, Word's check 54 cases 27 and 28)
        references = self._comment_references_in(start, end, whole) if end > start or whole else []
        if tracker is not None:
            # the paragraph primitive's replacement is the w:del first and the w:ins
            # after it (CR-003 section 4); a Range's tracked Replace puts the w:ins
            # first, as Word does, through span_edits (CR-013 Phase C)
            deletions = (
                self._delete_text(tracker, start, end, whole=whole)
                if end > start or hidden or references
                else []
            )
            if text:
                self._insert_tracked(
                    tracker, start, text, deletions[-1] if deletions else None, edge=edge, prefer=prefer
                )
                explode_mapped(self.segments(), start, start + len(text))
            return Range(self, start, start + len(text))

        inserted = False
        if end > start or hidden or references:
            first = next((s for s in segments if s.start <= start < s.end), None)
            to_remove = list(hidden)
            for segment in segments:
                if segment.end <= start or segment.start >= end:
                    continue
                if segment.editable:
                    from_ = max(0, start - segment.start)
                    to = min(len(segment.text), end - segment.start)
                    middle = text if segment is first and not inserted else ""
                    if segment is first:
                        inserted = True
                    value = segment.text[:from_] + middle + segment.text[to:]
                    if value == "":
                        to_remove.append(segment)
                    else:
                        set_text(segment.item, value)
                else:
                    to_remove.append(segment)
            for segment in reversed(to_remove):
                for position, item in enumerate(segment.owner):
                    if item is segment.item:
                        del segment.owner[position]
                        break
            if references:
                from docx4j_py.model.content.span_edits import queue_comment_removal

                for run, owner, identifier in references:
                    at = _index_of(owner, run)
                    if at >= 0:
                        del owner[at]
                    queue_comment_removal(self.parent_body, identifier)
            self._remove_empty_runs()
            if inserted and text:
                explode_mapped(self.segments(), start, start + len(text))
                return Range(self, start, start + len(text))
            segments = self.segments()

        if text and self._place_at_hidden_edge(start, end, text, edge):
            explode_mapped(self.segments(), start, start + len(text))
            return Range(self, start, start + len(text))
        if text:
            target = (
                next((s for s in segments if s.editable and s.start < start < s.end), None)
                or next((s for s in segments if s.editable and s.end == start), None)
                or next((s for s in segments if s.editable and s.start == start), None)
            )
            if target is not None:
                at = start - target.start
                set_text(target.item, target.text[:at] + text + target.text[at:])
            else:
                before = next((s for s in reversed(segments) if s.end <= start), None)
                after = next((s for s in segments if s.start >= start), None)
                neighbour = before or after
                r_pr = None
                source = neighbour.run if neighbour is not None and not neighbour.math else None
                if source is None:
                    # no run beside it: the first run's look, unless that run is hidden
                    # (core-ts's visibleFirstRunRPr, 6285dfa; CR-013 section 11)
                    source = self._visible_first_run()
                if source is not None and getattr(source, "r_pr", None) is not None:
                    r_pr = deep_copy(source.r_pr)
                run = R(content=ChildList([t(text)]))
                if r_pr is not None:
                    run.r_pr = r_pr
                    r_pr.parent = run
                # beside the neighbour's run, or beside the kept element its text
                # is read from, never into a branch (CR-011 Phase F)
                if before is not None:
                    item, owner, index = before.site()
                    owner.insert(index + 1, run)
                    run.parent = getattr(item, "parent", None) or self.element
                elif after is not None:
                    item, owner, index = after.site()
                    owner.insert(index, run)
                    run.parent = getattr(item, "parent", None) or self.element
                elif start == 0:
                    self.element.content.insert(0, run)
                else:
                    self.element.content.append(run)
                link_parents(run)
            explode_mapped(self.segments(), start, start + len(text))
        return Range(self, start, start + len(text))

    def insert_items_at(self, offset: int, items: list) -> None:
        """Insert run-level items at a text offset, splitting the run there.

        The parents of what is inserted are linked. This is what merging a
        one-paragraph fragment into this paragraph is built on.
        """
        if not items:
            return
        self.refuse_branch_text(offset, offset, "this insertion")
        refuse_equation(self.segments(), offset, offset)
        offset = self.split_at(offset)
        segments = self.segments()
        after = next((s for s in segments if s.start >= offset), None)
        before = next((s for s in reversed(segments) if s.end <= offset), None)
        neighbour = after or before
        if neighbour is None:
            target = self.element.content
            at = 0 if offset == 0 else len(target)
            owner: Any = self.element
        else:
            # beside the run, or beside the kept element its text is read from (CR-011 Phase F)
            item, target, index = neighbour.site()
            at = index if after is not None else index + 1
            owner = getattr(item, "parent", None) or self.element
        for offset_in_items, item in enumerate(items):
            target.insert(at + offset_in_items, item)
            link_parents(item)
            item.parent = owner

    # -- tracked editing (CR-003 section 3.8, Phase F) ----------------------

    def delete_text_tracked(
        self, tracker: Any, start: int, end: int, *, whole: bool | tuple[bool, bool] = False
    ) -> list[Any]:
        """Turn the text in ``[start, end)`` into a ``w:del``. Extension.

        The tracked half of :meth:`splice`, reachable on its own because a
        **deleted table row** deletes the content of every cell through exactly
        this primitive (CR-003 section 16.10). `whole` takes the hidden text at
        the span's edges too (CR-013 Phase A).
        """
        return self._delete_text(tracker, start, end, whole=whole)

    def holds_hidden_text(self) -> bool:
        """Whether a hidden run of this paragraph holds text (CR-013 Phase A). Extension."""
        return any(segment.hidden for segment in self.segments(objects=True))

    def holds_objects(self) -> bool:
        """Whether this paragraph holds a place of no text: a hidden run's item or a drawn object. Extension.

        What a deletion of the whole paragraph takes even where it has no text
        (CR-013 Phase A's hidden runs; core-ts CR-002 section 48's pictures).
        """
        return any(segment.object for segment in self.segments(objects=True))

    def _visible_first_run(self) -> Any:
        """The first run that is not hidden, or None: whose look new text with no run beside it takes.

        core-ts's ``visibleFirstRunRPr`` (``6285dfa``): text put into a paragraph
        whose every run is hidden took the first run's look, ``w:vanish`` with
        it, so it read as nothing and escaped the writers' map (CR-013 section
        11, measured at ``e79a8e8``).
        """
        hidden = self.hidden_runs()
        return next((run for run in runs_of(self.element) if not hidden(run)), None)

    def _place_at_hidden_edge(
        self, start: int, end: int, text: str, edge: str | None, *, tracker: Any = None
    ) -> bool:
        """Text put at an edge where hidden runs stand: after those at the end, before those at the start.

        CR-013 section 11 (the review's item 1): the paragraph's ``"End"`` is
        after everything in it, hidden text too, and its ``"Start"`` before; in a
        paragraph with no visible text an insertion goes at the end. The new run
        takes the look of the visible text beside it, or none, never a hidden
        run's. Beside a revision the run goes beside the revision, as typed text
        does. Returns whether it placed the text.
        """
        from docx4j_py.model.content.tracking import copy_r_pr, revision_of

        listed = self.segments(objects=True)
        found = _hidden_edge(listed, start, end, edge)
        if found is None:
            return False
        segment, how = found
        visible = [s for s in listed if not s.object and not s.math and s.text]
        neighbour = (visible[-1] if how == "after" else visible[0]) if visible else None
        r_pr = copy_r_pr(getattr(neighbour.run, "r_pr", None)) if neighbour is not None else None
        item: Any = _run_of(text, r_pr)
        if tracker is not None:
            item = tracker.ins([item])
        site, owner, index = segment.site()
        revision = revision_of(site)
        if revision is not None:
            site, owner = revision.element, revision.owner
            index = _index_of(owner, site)
        parent = getattr(site, "parent", None) or self.element
        owner.insert(index + 1 if how == "after" else index, item)
        link_parents(item)
        item.parent = parent
        return True

    def _delete_text(
        self, tracker: Any, start: int, end: int, *, whole: bool | tuple[bool, bool] = False
    ) -> list[_Anchor]:
        """The text in ``[start, end)`` becomes a deletion; returns the ``w:del``\\ s.

        The runs holding it are isolated, their ``w:t`` turned into
        ``w:delText``, and each run of consecutive ones moved into one
        ``w:del``. Text **this author** had inserted is simply removed, as Word
        does, rather than nested in a ``w:del``. A hidden run strictly inside
        the span is deleted with it (at its edges too when `whole`), as Word
        puts it inside the ``w:del`` (CR-013 Phase A). A comment's reference
        run in the span goes into the ``w:del`` where it stands among the runs,
        as Word puts it (docx4j-core-ts check 54 case 28, its CR-014 section
        12.6; CR-013 Phase C), so that accepting the deletion takes the comment;
        taken back with this author's own insertion, it takes the comment at
        once. An equation in the span is deleted where it stands, in Word's form
        --- a ``w:del`` in each math run and each structure's ``m:ctrlPr``
        (:func:`~.tracking.mark_equation_deleted`; CR-013 section 13.7, D2;
        docx4j-core-ts check 57) --- the runs either side in ``w:del``\\ s of
        their own; one holding a revision of its own goes into the ``w:del``
        whole.
        """
        from docx4j_py.model.content.tracking import (
            equation_revision_in,
            mark_equation_deleted,
            revision_of,
            to_deleted_text,
        )

        self.refuse_branch_text(start, end, "a tracked deletion")
        refuse_equation(self.segments(), start, end, tracked="delete")
        targets = self._isolate(start, end, whole=whole)
        for segment in targets:
            tracker.assert_editable(revision_of(segment.run))

        runs: list[_Target] = []
        equations: list[Any] = []
        for segment in targets:
            if (
                segment.math
                and not tracker.own_insertion(revision_of(segment.run))
                and equation_revision_in(segment.item) is None
            ):
                equations.append(segment.item)
                continue
            if runs and runs[-1].run is segment.run:
                continue
            revision = revision_of(segment.run)
            runs.append(
                _Target(
                    run=segment.run,
                    owner=segment.run_owner,
                    own=tracker.own_insertion(revision),
                    revision=revision,
                    parent=getattr(segment.run, "parent", None),
                )
            )
        references = self._comment_references_in(start, end, whole)
        if references:
            # each in document order among the runs, so that it joins their w:del
            order = _run_order(self.element)
            for run, owner, _identifier in references:
                if any(target.run is run for target in runs):
                    continue
                revision = revision_of(run)
                target = _Target(
                    run=run,
                    owner=owner,
                    own=tracker.own_insertion(revision),
                    revision=revision,
                    parent=getattr(run, "parent", None),
                )
                key = order.get(id(run), -1)
                at = next((i for i, other in enumerate(runs) if order.get(id(other.run), -1) > key), len(runs))
                runs.insert(at, target)

        for equation in equations:
            mark_equation_deleted(tracker, equation)
        groups: list[list[_Target]] = []
        for target in runs:
            group = groups[-1] if groups else None
            previous = group[-1] if group else None
            if (
                group is not None
                and previous is not None
                and previous.own == target.own
                and previous.owner is target.owner
                and _index_of(target.owner, target.run)
                == _index_of(previous.owner, previous.run) + 1
            ):
                group.append(target)
            else:
                groups.append([target])

        made: list[_Anchor] = []
        for group in reversed(groups):
            owner = group[0].owner
            at = _index_of(owner, group[0].run)
            if at < 0:
                continue
            taken = list(owner[at : at + len(group)])
            del owner[at : at + len(group)]
            if group[0].own:
                # taking back our own insertion: the runs go, and an emptied
                # w:ins with them; a comment whose reference goes with them goes too
                taken_references = _comment_reference_ids(taken)
                if taken_references:
                    from docx4j_py.model.content.span_edits import queue_comment_removal

                    for identifier in taken_references:
                        queue_comment_removal(self.parent_body, identifier)
                revision = group[0].revision
                if revision is not None and not revision.items:
                    index = _index_of(revision.owner, revision.element)
                    if index >= 0:
                        del revision.owner[index]
                continue
            for run in taken:
                to_deleted_text(run)
            deletion = tracker.deletion(taken)
            owner.insert(at, deletion)
            revision = group[0].revision
            # the parent captured before the move: ``tracker.deletion`` has just
            # made the runs children of the w:del, so asking the run now would
            # make the w:del its own runs' sibling's child (section 16.9)
            parent = (
                revision.element if revision is not None else (group[0].parent or self.element)
            )
            deletion.parent = parent
            made.insert(0, _Anchor(owner=owner, element=deletion, parent=parent))
        return made

    def _insert_tracked(
        self,
        tracker: Any,
        at: int,
        text: str,
        anchor: _Anchor | None = None,
        *,
        edge: str | None = None,
        prefer: str = "before",
    ) -> None:
        """Text inserted at `at` as a ``w:ins``, extending one of ours where it can.

        A run of this author's own that is already inside a ``w:ins`` is
        extended rather than nested in another one, as Word does. `anchor` is
        the ``w:del`` of a replacement, which the insertion must follow.
        `prefer` ``"after"`` looks at the run starting at `at` first: the
        ``w:ins`` goes before it, in its look (core-ts's ``insertTracked``
        ``prefer``; a tracked replacement's insertion, CR-013 Phase C).
        """
        from docx4j_py.model.content.tracking import copy_r_pr, revision_of

        if anchor is not None:
            source = runs_of(anchor.element, view="original")
            # the first deleted run's look, unless it was hidden: a hidden one's would hide
            # the new text (CR-013 section 11, the review's item 1)
            first = source[0] if source and not self.hidden_runs()(source[0]) else None
            r_pr = copy_r_pr(getattr(first, "r_pr", None))
            wrapper = tracker.ins([_run_of(text, r_pr)])
            index = _index_of(anchor.owner, anchor.element)
            anchor.owner.insert(index + 1 if index >= 0 else len(anchor.owner), wrapper)
            wrapper.parent = anchor.parent
            return

        self.refuse_branch_text(at, at, "a tracked insertion")
        refuse_equation(self.segments(), at, at)
        if self._place_at_hidden_edge(at, at, text, edge, tracker=tracker):
            return
        at = self.split_at(at)
        segments = self.segments()
        before = next((s for s in reversed(segments) if s.end <= at), None)
        after = next((s for s in segments if s.start >= at), None)
        if (
            before is not None
            and after is not None
            and before.run is after.run
            and before.branch is None
            and not after.math
            and after.index > 0
        ):
            # a tab, a break or a symbol sharing the run with the text after it: the run cut between them, so
            # that the insertion stands at `at` and not beside the whole run (found by the landing audit's
            # probe between a tab and its run's text; CR-013 section 14.9)
            self._split_run_before(after)
            segments = self.segments()
            before = next((s for s in reversed(segments) if s.end <= at), None)
            after = next((s for s in segments if s.start >= at), None)
        sides = ((before, "after"), (after, "before"))
        for segment, side in sides if prefer != "after" else tuple(reversed(sides)):
            if segment is None or segment.branch is not None:
                # text read from a kept branch is never extended (CR-011 Phase F)
                continue
            revision = revision_of(segment.run)
            if not tracker.own_insertion(revision):
                continue
            if segment.editable:
                value = (
                    segment.text + text if side == "after" else text + segment.text
                )
                set_text(segment.item, value)
            else:
                run = _run_of(text, copy_r_pr(getattr(segment.run, "r_pr", None)))
                index = segment.run_index + (1 if side == "after" else 0)
                segment.run_owner.insert(index, run)
                link_parents(run)
                run.parent = revision.element if revision is not None else self.element
            return

        neighbour = (after or before) if prefer == "after" else (before or after)
        if self._insert_between_halves(tracker, text, before, after, neighbour):
            return
        # beside the run, or beside the kept element its text is read from (CR-011 Phase F)
        site, site_owner, _site_index = neighbour.site() if neighbour is not None else (None, None, -1)
        revision = revision_of(site) if site is not None else None
        tracker.assert_editable(revision)
        source = neighbour.run if neighbour is not None else None
        if source is None:
            # the first run's look, unless it is hidden (core-ts's visibleFirstRunRPr)
            source = self._visible_first_run()
        wrapper = tracker.ins([_run_of(text, copy_r_pr(getattr(source, "r_pr", None)))])
        if neighbour is None:
            if at == 0:
                self.element.content.insert(0, wrapper)
            else:
                self.element.content.append(wrapper)
            wrapper.parent = self.element
            return
        owner = revision.owner if revision is not None else site_owner
        item = revision.element if revision is not None else site
        parent = (
            getattr(revision.element, "parent", None)
            if revision is not None
            else getattr(site, "parent", None)
        ) or self.element
        index = _index_of(owner, item)
        owner.insert(index + (1 if neighbour is before else 0), wrapper)
        wrapper.parent = parent

    def split_runs_around(self, start: int, end: int) -> tuple[int, int]:
        """Split runs so that every run holding text in ``[start, end)`` holds nothing else. Extension.

        core-ts's ``splitRunsAround`` (``Paragraph.mts`` at its CR-012 phase 4,
        ``9307bdb``): the offsets snapped to grapheme boundaries as
        :meth:`split_at` snaps them (the start back, the end forward), the runs
        cut there, and a tab, a break or a drawing that shares a run with the
        span's edge moved into a run of its own. What a ``font`` write formats
        and where a comment's markers go, so that neither reaches a character
        outside the span (CR-013 Phase C; the comment's markers went a whole
        run wide until then). Returns the offsets used.
        """
        start = self.split_at(start, prefer="back")
        end = self.split_at(end, prefer="forward") if end > start else start
        if end > start:
            self._isolate(start, end)
        return start, end

    def _comment_references_in(
        self, start: int, end: int, whole: bool | tuple[bool, bool] = False
    ) -> list[tuple[Any, list, int]]:
        """The runs of comments' references in ``[start, end)``: ``(run, owner, comment id)``.

        Each a run holding nothing but a reference, where the walk passes it,
        strictly inside the span or at an edge `whole` holds (:func:`_in_span`);
        not in a kept branch, which is read-only. A reference has no text, so a
        deletion's segments never hold it, and a comment whose reference a
        deletion took stayed: Word's deletion takes it with the text
        (docx4j-core-ts check 54 cases 27 and 28, its ``commentReferencesIn``,
        CR-014 section 12.6; CR-013 Phase C).
        """
        out: list[tuple[Any, list, int]] = []

        def visit(item: Any, at: int, owner: Any, run: Any, run_owner: Any, branch: Any) -> None:
            if run is None or branch is not None or element_name(item) != _W_COMMENT_REFERENCE:
                return
            if not _in_span(at, start, end, whole) or any(found is run for found, _owner, _id in out):
                return
            if not all(element_name(child) == _W_COMMENT_REFERENCE for child in run.content or ()):
                return
            identifier = getattr(item, "id", None)
            if isinstance(identifier, int) and isinstance(run_owner, list):
                out.append((run, run_owner, identifier))

        segments_of(self.element, hidden=self.hidden_runs(), visit=visit)
        return out

    def _insert_between_halves(self, tracker: Any, text: str, before: Any, after: Any, neighbour: Any) -> bool:
        """A tracked insertion strictly inside another author's ``w:ins`` or ``w:moveTo``: the revision split, the text between.

        Word's form for text typed inside someone else's revision (docx4j-core-ts
        check 20 case 01, check 21; its ``insertBetweenHalves``): the revision's
        items from the place on go into a second one of the same, under a new
        ``w:id``, and this author's ``w:ins`` stands between the two. Until CR-013
        Phase C the ``w:ins`` went after the whole revision (section 11.1.9);
        with a tracked replacement's insertion first, the text would have gone
        before it. Only where both sides' runs stand directly in the one
        revision. Returns whether it placed the text.
        """
        from docx4j_py.model.content.comments import _close_insertion
        from docx4j_py.model.content.tracking import copy_r_pr, revision_of

        if before is None or after is None or before.branch is not None or after.branch is not None:
            return False
        first, second = revision_of(before.run), revision_of(after.run)
        if first is None or second is None or first.element is not second.element:
            return False
        if first.kind not in ("ins", "moveTo") or tracker.own_insertion(first):
            return False
        items = run_items_of(first.element)
        if items is None or before.run_owner is not items or after.run_owner is not items:
            return False
        index = _index_of(items, after.run)
        if index <= 0:
            return False
        source = neighbour.run if neighbour is not None else None
        wrapper = tracker.ins([_run_of(text, copy_r_pr(getattr(source, "r_pr", None)))])
        return _close_insertion(self, first.element, index, [wrapper])

    def _isolate(self, start: int, end: int, *, whole: bool | tuple[bool, bool] = False) -> list[Any]:
        """Split runs so that each one holding text in ``[start, end)`` holds nothing else.

        :meth:`split_at` does the ``w:t`` boundaries; this also moves the items
        of a run that straddles a boundary --- a tab, a break, a drawing --- out
        into runs of their own, so that a partly deleted run is not wholly
        deleted. A hidden run's items strictly inside the span (at its edges
        too when `whole`) are among what is returned (CR-013 Phase A).
        """

        def inside() -> list[Any]:
            return [
                segment
                for segment in self.segments(objects=True)
                if (segment.object and _in_span(segment.start, start, end, whole))
                or (not segment.object and segment.start >= start and segment.end <= end and segment.text)
            ]

        self.split_at(start)
        self.split_at(end, prefer="forward")
        for _guard in range(10_000):  # pragma: no branch - the loop always settles
            found = inside()
            split = False
            seen: set[int] = set()
            for segment in found:
                if id(segment.run) in seen:
                    continue
                seen.add(id(segment.run))
                if segment.math:
                    # in no run: taken whole or refused (core-ts's isolate)
                    continue
                of_run = [s for s in found if s.run is segment.run]
                first, last = of_run[0], of_run[-1]
                if first.index > 0:
                    self._split_run_before(first)
                    split = True
                    break
                if last.index < len(last.owner) - 1:
                    self._split_run_after(last)
                    split = True
                    break
            if not split:
                return found
        return inside()  # pragma: no cover - 10,000 splits in one paragraph

    def _split_run_before(self, segment: Any) -> None:
        """Move the items before `segment` into a run of their own, in front of it."""
        from docx4j_py.model.content.tracking import copy_r_pr

        head = list(segment.owner[: segment.index])
        del segment.owner[: segment.index]
        run = _run_items(head, copy_r_pr(getattr(segment.run, "r_pr", None)))
        _carry_attributes(segment.run, run)
        segment.run_owner.insert(segment.run_index, run)
        run.parent = getattr(segment.run, "parent", None) or self.element
        self.remember_split([(run, segment.run)])

    def _split_run_after(self, segment: Any) -> None:
        """Move the items after `segment` into a run of their own, behind it."""
        from docx4j_py.model.content.tracking import copy_r_pr

        tail = list(segment.owner[segment.index + 1 :])
        if not tail:
            return
        del segment.owner[segment.index + 1 :]
        run = _run_items(tail, copy_r_pr(getattr(segment.run, "r_pr", None)))
        _carry_attributes(segment.run, run)
        segment.run_owner.insert(segment.run_index + 1, run)
        run.parent = getattr(segment.run, "parent", None) or self.element
        self.remember_split([(segment.run, run)])

    def _remove_empty_runs(self) -> None:
        """Drop runs an edit emptied, and revisions left holding nothing."""
        from docx4j_py.model.content.text_model import _REVISIONS, RUN_HOLDER_NAMES
        from docx4j_py.traversal import element_name, run_items_of

        def prune(items: list) -> None:
            for position in range(len(items) - 1, -1, -1):
                element = items[position]
                name = element_name(element)
                if isinstance(element, R):
                    if not element.content:
                        del items[position]
                    continue
                if name in RUN_HOLDER_NAMES or name in _REVISIONS:
                    nested = run_items_of(element)
                    if nested is None:
                        continue
                    prune(nested)
                    if not nested and name in _REVISIONS:
                        del items[position]

        prune(self.element.content)


def _hidden_edge(listed: list[Any], start: int, end: int, edge: str | None) -> tuple[Any, str] | None:
    """The hidden segment text at a paragraph's edge goes beside, and which side: for :meth:`Paragraph._place_at_hidden_edge`.

    ``"end"`` at the text's end: after the last hidden run there; ``"start"`` at
    0: before the first. With no visible text at all, any insertion is at the
    end. None where no hidden run stands at that edge.
    """
    if start != end:
        return None
    length = max((s.end for s in listed if not s.object), default=0)
    hidden = [s for s in listed if s.hidden]
    if not hidden:
        return None
    if edge is None and length == 0:
        edge = "end"
    if edge == "end" and start == length:
        trailing = [s for s in hidden if s.start == length]
        return (trailing[-1], "after") if trailing else None
    if edge == "start" and start == 0:
        leading = [s for s in hidden if s.start == 0]
        return (leading[0], "before") if leading else None
    return None


def _refuse_run_control_break(operation: str = "this write") -> None:
    """Refuse a ``\\n`` or ``\\r`` written into a run-level content control (core-ts's F4)."""
    raise SpanError(
        f"{operation} refused: a paragraph break (\\n or \\r) written into a run-level content control "
        "would cut the control in two (CR-013 section 11, E1; docx4j-core-ts F4)",
        code="range.splits_control",
        hint="write a line break (U+000B) instead, or write the paragraph outside the control",
    )


def _after_deletion(segments: list[Any], start: int, end: int) -> list[Any]:
    """The segments as they will stand once ``[start, end)`` is deleted, offsets shifted: for :meth:`Paragraph._landing`.

    What the deletion keeps: everything outside the span, and the part outside
    it of an editable item it cuts; an item of no editable text inside goes
    whole, as :meth:`Paragraph._splice_text` removes it.
    """
    shift = end - start
    out: list[Any] = []
    for segment in segments:
        if segment.end <= start:
            out.append(segment)
        elif segment.start >= end:
            out.append(dataclasses.replace(segment, start=segment.start - shift, end=segment.end - shift))
        elif segment.editable:
            text = segment.text[: max(0, start - segment.start)] + segment.text[max(0, end - segment.start) :]
            if text:
                at = min(segment.start, start)
                out.append(dataclasses.replace(segment, text=text, start=at, end=at + len(text)))
    return out


def _untracked_landing(segments: list[Any], start: int) -> tuple[Any, str]:
    """Where :meth:`Paragraph._splice_text` puts a text at a collapsed `start`, untracked: its target rule."""
    target = (
        next((s for s in segments if s.editable and s.start < start < s.end), None)
        or next((s for s in segments if s.editable and s.end == start), None)
        or next((s for s in segments if s.editable and s.start == start), None)
    )
    if target is not None:
        return target.item, "into"
    before = next((s for s in reversed(segments) if s.end <= start), None)
    after = next((s for s in segments if s.start >= start), None)
    if before is not None:
        return before.site()[0], "after"
    if after is not None:
        return after.site()[0], "before"
    return None, "start" if start == 0 else "end"


def _tracked_landing(tracker: Any, segments: list[Any], at: int, prefer: str = "before") -> tuple[Any, str]:
    """Where :meth:`Paragraph._insert_tracked` puts a ``w:ins`` at `at`: this author's insertion extended, or beside."""
    from docx4j_py.model.content.tracking import revision_of

    inside = next((s for s in segments if s.editable and s.start < at < s.end), None)
    before = inside or next((s for s in reversed(segments) if s.end <= at), None)
    after = inside or next((s for s in segments if s.start >= at), None)
    sides = ((before, "after"), (after, "before"))
    for segment, side in sides if prefer != "after" else tuple(reversed(sides)):
        if segment is None or segment.branch is not None:
            continue
        if not tracker.own_insertion(revision_of(segment.run)):
            continue
        return (segment.item, "into") if segment.editable else (segment.run, side)
    neighbour = (after or before) if prefer == "after" else (before or after)
    if neighbour is None:
        return None, "start" if at == 0 else "end"
    site = neighbour.site()[0]
    revision = revision_of(site)
    if revision is not None:
        return revision.element, "after" if neighbour is before else "before"
    if inside is not None:
        # split there first: the w:ins goes between the halves, where the w:t was
        return inside.item, "into"
    return site, "after" if neighbour is before else "before"


def replace_matches(matches: list[Range], replace: str) -> None:
    """Each match's text replaced, last first, each a ``Replace`` of :meth:`Range.insert_text`.

    A replacement holding a paragraph break leaves the text before it in a new
    paragraph in front (CR-013 section 11, E1: the ``w:p`` keeps the last), so
    the matches still to do, all before it, move there with their offsets.
    """
    for position in range(len(matches) - 1, -1, -1):
        match = matches[position]
        paragraph = match.paragraph
        container = paragraph.container
        index = _index_of(container, paragraph.element)
        match.insert_text(replace, location="Replace")
        if 0 <= index < len(container) and container[index] is not paragraph.element:
            first = Paragraph(container[index], container, paragraph.parent_body)
            for earlier in matches[:position]:
                if earlier.paragraph.element is paragraph.element:
                    earlier.paragraph = first


def _in_span(at: int, start: int, end: int, whole: bool | tuple[bool, bool]) -> bool:
    """Whether a place of no text at `at` is in the span: strictly inside it, or at an edge of a whole one.

    A hidden run stands between two characters, so one at a span's start or end
    offset may be inside it or outside it, and offsets alone cannot say which.
    A paragraph's whole text holds the hidden runs at both of its edges; a span
    of text, such as a search found, holds only those strictly inside it
    (core-ts ``objectInSpan`` and ``WHOLE`` at ``9307bdb``). `whole` may say it
    edge by edge, ``(start, end)``, as core-ts's ``SpanEdges`` does: a range
    across paragraphs holds its first paragraph's end and its last's start
    (CR-013 Phase C).
    """
    if isinstance(whole, tuple):
        held_start, held_end = whole
        return start < at < end or (held_start and at == start) or (held_end and at == end)
    return start < at < end or (whole and start <= at <= end)


def _edges_of(whole: bool | tuple[bool, bool], start: int, end: int, length: int) -> bool | tuple[bool, bool]:
    """A span's edges as :func:`_in_span` reads them: both where it is the paragraph's whole text."""
    if start == 0 and end == length and end > start:
        return True
    if isinstance(whole, tuple):
        return True if all(whole) else whole
    return whole


#: A comment's reference, ``w:commentReference``.
_W_COMMENT_REFERENCE = f"{{{WML_NS}}}commentReference"


def _run_order(element: Any) -> dict[int, int]:
    """Every element among a paragraph's runs, in document order, by ``id()``: where a deletion puts a reference run."""
    order: dict[int, int] = {}

    def visit(items: list | None) -> None:
        for item in items or ():
            order[id(item)] = len(order)
            if not isinstance(item, R):
                visit(run_items_of(item))

    visit(run_items_of(element))
    return order


def _comment_reference_ids(runs: list) -> list[int]:
    """The ids of the comment references among runs."""
    out: list[int] = []
    for run in runs:
        for child in getattr(run, "content", None) or ():
            if element_name(child) == _W_COMMENT_REFERENCE and isinstance(getattr(child, "id", None), int):
                out.append(child.id)
    return out


def _text_run(text: str) -> R:
    """A run of the text, each of Office JS's characters its element (CR-013 Q2, :func:`~.text_model.text_items`)."""
    return text_run(text)


# ---------------------------------------------------------------------------
# the tracked primitives' helpers (CR-003 section 3.8, Phase F)
# ---------------------------------------------------------------------------


def _index_of(items: list, element: Any) -> int:
    """The index of an element **by identity**; ``-1`` when it is not there."""
    for index, item in enumerate(items):
        if item is element:
            return index
    return -1


def _run_of(text: str, r_pr: Any = None) -> R:
    """One ``w:r`` of exactly this text, with the formatting given."""
    return _run_items([t(text)], r_pr)


def _carry_attributes(source: R, made: R) -> None:
    """Copy a run's own attributes onto a run split off from it (section 16.12)."""
    from docx4j_py.model.content.tracking import RUN_ATTRIBUTES

    for name in RUN_ATTRIBUTES:
        value = getattr(source, name, None)
        if value is not None:
            setattr(made, name, value)


def _run_items(items: list, r_pr: Any = None) -> R:
    """One ``w:r`` of these run children, with the formatting given, parents linked."""
    run = R(content=ChildList(list(items)))
    if r_pr is not None:
        run.r_pr = r_pr
    link_parents(run)
    return run


@dataclasses.dataclass(frozen=True, slots=True)
class _Target:
    """One run a tracked deletion is about to move into a ``w:del``."""

    run: Any
    owner: list
    own: bool
    revision: Any = None
    parent: Any = None
    """What held the run **before** the move, captured now: moving it into the
    ``w:del`` rewrites ``run.parent``, and the ``w:del`` (and the ``w:ins`` that
    follows it) belong where the run was (CR-003 section 16.9)."""


@dataclasses.dataclass(frozen=True, slots=True)
class _Anchor:
    """Where a ``w:ins`` goes when it must follow the ``w:del`` of a replacement."""

    owner: list
    element: Any
    parent: Any
