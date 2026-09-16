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

from typing import TYPE_CHECKING, Any

from docx4j_py.child import ChildList, deep_copy, link_parents
from docx4j_py.model.content.enums import (
    AlignmentValue,
    BreakTypeValue,
    ParagraphLocation,
    RangeLocation,
    TextLocation,
    TextView,
)
from docx4j_py.model.content.errors import ContentError
from docx4j_py.model.content.font import Font
from docx4j_py.model.content.styles import (
    built_in_of,
    id_of_built_in,
    style_id_of,
    style_name_of,
)
from docx4j_py.model.content.text_model import (
    find_all,
    runs_of,
    search_pattern,
    segments_of,
    set_text,
    split_at,
    text_of_view,
)
from docx4j_py.wml import P, R, br, el, t, to_xml, wml

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.body import Body
    from docx4j_py.model.content.range import Range

__all__ = ["Paragraph"]

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
        #: The nearest :class:`~docx4j_py.model.content.Body`.
        self.parent_body = parent_body

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
        """``<Paragraph body/3 'Chapter 2: The …'>``; the address form is Phase D."""
        text = self.text
        preview = text[:_PREVIEW] + "…" if len(text) > _PREVIEW else text
        index = self.index
        where = f" {self.parent_body.prefix}/{index}" if index >= 0 else ""
        return f"<Paragraph{where} {preview!r}>"

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
        """The text: runs joined, tabs as ``\\t``, breaks as ``\\n``, deletions out."""
        return text_of_view(self.element)

    @text.setter
    def text(self, value: str) -> None:
        r_pr = None
        runs = runs_of(self.element)
        if runs and getattr(runs[0], "r_pr", None) is not None:
            r_pr = deep_copy(runs[0].r_pr)
        run = R(content=ChildList([t(value)] if value else []))
        if r_pr is not None:
            run.r_pr = r_pr
            r_pr.parent = run
        self.element.content[:] = [run]
        link_parents(self.element)

    def get_text(self, *, view: TextView = "accepted") -> str:
        """The text in one of the two views of a tracked document.

        ``"accepted"`` (the default) is :attr:`text`; ``"original"`` reads the
        document as it was before the tracked changes (CR-003 section 3.8).
        """
        return text_of_view(self.element, view=view)

    @property
    def runs(self) -> list[R]:
        """Every ``w:r``, direct and nested in hyperlinks, controls and insertions."""
        return runs_of(self.element)

    def segments(self, *, view: TextView = "accepted") -> list[Any]:
        """The text segments with their offsets. Extension; see :mod:`.text_model`."""
        return segments_of(self.element, view=view)

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
        define unless it is a built-in style.
        """
        return style_name_of(self.parent_body.package, self.style_id)

    @style.setter
    def style(self, value: str) -> None:
        self.style_id = style_id_of(self.parent_body.package, value, validate=True)

    @property
    def style_built_in(self) -> str:
        """The ``Word.Style`` value (``"Heading1"``), or ``"Other"``, as Office JS."""
        return built_in_of(self.style_id)

    @style_built_in.setter
    def style_built_in(self, value: str) -> None:
        self.style_id = id_of_built_in(value)

    # -- paragraph properties ---------------------------------------------

    def _p_pr(self) -> Any:
        """The ``w:pPr``, created and linked when absent. Every setter goes here."""
        p_pr = self.element.p_pr
        if p_pr is None:
            p_pr = el.pPr()
            self.element.p_pr = p_pr
            p_pr.parent = self.element
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

    @property
    def _ind_or_none(self) -> Any:
        p_pr = self.element.p_pr
        return getattr(p_pr, "ind", None) if p_pr is not None else None

    @property
    def _spacing_or_none(self) -> Any:
        p_pr = self.element.p_pr
        return getattr(p_pr, "spacing", None) if p_pr is not None else None

    @property
    def alignment(self) -> AlignmentValue:
        """Office JS ``Word.Alignment``; ``"Unknown"`` when there is no ``w:jc``.

        ``start`` and ``end`` map like ``left`` and ``right``, so a document
        written by a right-to-left Word reads the same (CR-003 section 3.12).
        """
        p_pr = self.element.p_pr
        value = _value(getattr(p_pr, "jc", None)) if p_pr is not None else None
        return _JC_TO_ALIGNMENT.get(value or "", "Unknown")  # type: ignore[return-value]

    @alignment.setter
    def alignment(self, value: AlignmentValue) -> None:
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
        p_pr = self.element.p_pr
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

    # -- formatting --------------------------------------------------------

    @property
    def font(self) -> Font:
        """The direct formatting of the runs (CR-003 section 4: over the runs).

        Reads report the first run; writes go to every run. Not the paragraph
        mark's ``w:pPr/w:rPr``, which is a different thing and is what Word
        calls the paragraph mark's font.
        """
        return Font(lambda: runs_of(self.element), scope="paragraph")

    # -- the stable handles ------------------------------------------------

    @property
    def para_id(self) -> str | None:
        """``w14:paraId``, the stable address Word gives paragraphs (section 3.4)."""
        return self.element.para_id

    @para_id.setter
    def para_id(self, value: str | None) -> None:
        self.element.para_id = value

    @property
    def parent_table_cell(self) -> Any:
        """The cell this paragraph is in, or None. ``TableCell`` is Phase C.

        Wired now so that the member exists and code written against it keeps
        working; until Phase C it is None even inside a table.
        """
        return None

    # -- editing -----------------------------------------------------------

    def insert_text(self, text: str, *, location: TextLocation = "End") -> Range:
        """Insert text at the start, the end, or in place of the whole text.

        Args:
            text: what to insert.
            location: ``"Start"``, ``"End"`` (the default) or ``"Replace"``.

        Returns:
            The :class:`~docx4j_py.model.content.Range` of the inserted text.
        """
        length = len(self.text)
        if location == "Start":
            return self.splice(0, 0, text)
        if location == "End":
            return self.splice(length, length, text)
        if location == "Replace":
            return self.splice(0, length, text)
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
            text: the text of the new paragraph; ``""`` leaves it empty.
            location: ``"Before"`` or ``"After"`` (the default).
            style: a style name or id for the new paragraph.
        """
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
        item = br("page" if type == "Page" else None)
        if location in ("Before", "After"):
            new = self.insert_paragraph("", location=location)  # type: ignore[arg-type]
            new.element.content.append(R(content=ChildList([item])))
            link_parents(new.element)
            return
        run = R(content=ChildList([item]))
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
        elements = wml.all(xml, wrapper="body")
        if not elements:
            return []
        only = elements[0] if len(elements) == 1 and isinstance(elements[0], P) else None
        if only is not None and location in ("Start", "End"):
            self.insert_items_at(0 if location == "Start" else len(self.text), list(only.content))
            return [self]
        if location == "Replace":
            self.parent_body.insert_element(elements, location="Before", target=self)
            views = [self.parent_body.view_for(e) for e in elements]
            self.delete()
            return views
        where = "Before" if location == "Start" else "After" if location == "End" else location
        self.parent_body.insert_element(elements, location=where, target=self)
        return [self.parent_body.view_for(e) for e in elements]

    def search(self, text: str, **options: Any) -> list[Range]:
        """Every match of `text` in this paragraph, as ranges.

        Matches are found on the paragraph's whole text, so one **spans runs**
        freely; only formatting a match splits them.

        Args:
            text: what to look for.
            **options: ``match_case``, ``match_whole_word``, ``match_wildcards``
                and ``limit``, as :meth:`Body.search`.
        """
        from docx4j_py.model.content.range import Range

        limit = options.pop("limit", None)
        hits = find_all(self.text, search_pattern(text, **options))
        if limit is not None:
            hits = hits[:limit]
        return [Range(self, start, end) for start, end in hits]

    def replace_text(self, find: str, replace: str, **options: Any) -> int:
        """Replace every match, last first so the offsets stay valid.

        Returns:
            How many were replaced.
        """
        matches = self.search(find, **options)
        for match in reversed(matches):
            match.insert_text(replace, location="Replace")
        return len(matches)

    def get_range(self, location: RangeLocation = "Whole") -> Range:
        """A :class:`~docx4j_py.model.content.Range` over this paragraph."""
        from docx4j_py.model.content.range import Range

        length = len(self.text)
        if location == "Start":
            return Range(self, 0, 0)
        if location == "End":
            return Range(self, length, length)
        return Range(self, 0, length)

    def delete(self) -> None:
        """Remove the paragraph from its container."""
        index = self.index
        if index >= 0:
            del self.container[index]

    def get_xml(self) -> str:
        """The paragraph as XML, with docx4j's prefixes. Extension."""
        return to_xml(self.element)

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary: what a tool result says about a paragraph."""
        return {
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

    # -- docx4j names ------------------------------------------------------

    def get_content(self) -> list:
        """docx4j: the ``w:p``'s content list."""
        return self.element.content

    # -- the editing primitives, shared with Range -------------------------

    def split_at(self, offset: int, *, prefer: str = "back") -> int:
        """Split the run at `offset` so that ``[offset, ...)`` begins a run.

        The offset is snapped to a grapheme boundary first (CR-003 section
        3.12); the offset actually used is returned.
        """
        return split_at(self.element, offset, prefer=prefer)

    def splice(self, start: int, end: int, text: str) -> Range:
        """Replace the text in ``[start, end)`` and return the range of the new text.

        Runs are edited in place: a ``w:t`` loses or gains characters, a tab or
        a break inside the span goes, a run left empty goes. The new text takes
        the place --- and the formatting --- of the first replaced character, as
        Word does; where there is nothing to replace it extends the ``w:t`` at
        the offset, and failing that becomes a new run with the neighbouring
        run's formatting.
        """
        from docx4j_py.model.content.range import Range

        length = len(self.text)
        start = max(0, min(start, length))
        end = max(start, min(end, length))

        segments = self.segments()
        inserted = False
        if end > start:
            first = next((s for s in segments if s.start <= start < s.end), None)
            to_remove = []
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
            self._remove_empty_runs()
            if inserted and text:
                return Range(self, start, start + len(text))
            segments = self.segments()

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
                source = neighbour.run if neighbour is not None else None
                if source is None:
                    runs = runs_of(self.element)
                    source = runs[0] if runs else None
                if source is not None and getattr(source, "r_pr", None) is not None:
                    r_pr = deep_copy(source.r_pr)
                run = R(content=ChildList([t(text)]))
                if r_pr is not None:
                    run.r_pr = r_pr
                    r_pr.parent = run
                if before is not None:
                    before.run_owner.insert(before.run_index + 1, run)
                    run.parent = getattr(before.run, "parent", None) or self.element
                elif after is not None:
                    after.run_owner.insert(after.run_index, run)
                    run.parent = getattr(after.run, "parent", None) or self.element
                elif start == 0:
                    self.element.content.insert(0, run)
                else:
                    self.element.content.append(run)
                link_parents(run)
        return Range(self, start, start + len(text))

    def insert_items_at(self, offset: int, items: list) -> None:
        """Insert run-level items at a text offset, splitting the run there.

        The parents of what is inserted are linked. This is what merging a
        one-paragraph fragment into this paragraph is built on.
        """
        if not items:
            return
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
            target = neighbour.run_owner
            at = neighbour.run_index if after is not None else neighbour.run_index + 1
            owner = getattr(neighbour.run, "parent", None) or self.element
        for offset_in_items, item in enumerate(items):
            target.insert(at + offset_in_items, item)
            link_parents(item)
            item.parent = owner

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


def _text_run(text: str) -> R:
    """A run of the text, with tabs and newlines as ``w:tab`` and ``w:br``."""
    from docx4j_py.wml import r as run_builder

    return run_builder(text)
