"""The list definitions and the counters. docx4j ``model.listnumbering``.

Three classes and two of state, each one docx4j's:

:class:`ListLevel`
    one ``w:lvl``: its ``w:start``, ``w:numFmt``, ``w:lvlText``,
    ``w:lvlRestart``, ``w:isLgl``, the ``w:pStyle`` it is linked to, the font its
    ``w:rPr`` names and the ``w:ind`` its ``w:pPr`` states. An instance's level
    is a copy of the abstract one with the ``w:lvlOverride/w:lvl`` applied over
    it (:meth:`ListLevel.set_overrides`).
:class:`AbstractListNumberingDefinition`
    one ``w:abstractNum``, and the ``w:numStyleLink`` it may take its levels from.
:class:`ListNumberingDefinition`
    one ``w:num``: the abstract definition it names, with its ``w:lvlOverride``\\ s
    applied, and the counting --- :meth:`increment_counter` and
    :meth:`current_number_string`.
:class:`NumberingState`
    the counters of **one story of one traversal**, keyed by the *referencing*
    ``w:abstractNum`` and the level, which is the sharing rule Word applies
    (docx4j CR-014 probes P1 and P2, measured): every ``w:num`` over one abstract
    definition continues one sequence, and a ``w:numStyleLink`` definition is a
    list of its own.
:class:`NumberingStates`
    one state per story --- the body, the headers and footers together, the
    footnotes, the endnotes, the comments (CR-014 probe P7).

The definitions are read from **lxml**, not from the typed model: CR-003 Phase K
promised that reading a document never unmarshals the numbering part, and a
document whose labels are only read is still written back byte for byte. What a
level's ``w:ind`` contributes therefore comes out as :class:`Indent`, twips in a
frozen dataclass, rather than as a typed ``w:ind`` (CR-003 section 18).
"""

from __future__ import annotations

import dataclasses
from typing import Any

from docx4j_py.model.listnumbering.formats import format_value

__all__ = [
    "LEVELS",
    "AbstractListNumberingDefinition",
    "Counter",
    "Indent",
    "ListLevel",
    "ListNumberingDefinition",
    "NumberingState",
    "NumberingStates",
    "StyleNumbering",
    "read_definitions",
]

#: How many levels a list definition may have (ECMA-376: ``w:ilvl`` 0 to 8).
LEVELS = 9

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _w(local: str) -> str:
    return f"{_W}{local}"


def _val(element: Any, local: str) -> str | None:
    """``w:<local>/@w:val`` of a child element, or None."""
    if element is None:
        return None
    child = element.find(_w(local))
    return None if child is None else child.get(_w("val"))


def _int(text: str | None) -> int | None:
    if text is None:
        return None
    try:
        return int(text)
    except ValueError:
        return None


def _on(element: Any, local: str) -> bool:
    """Whether a boolean child element is present and not turned off."""
    if element is None:
        return False
    child = element.find(_w(local))
    if child is None:
        return False
    value = child.get(_w("val"))
    return value not in ("0", "false", "off")


# ---------------------------------------------------------------------------
# the indent a level contributes
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True, slots=True)
class Indent:
    """A level's ``w:pPr/w:ind``, in twips. What docx4j returns as a typed ``Ind``."""

    left: int | None = None
    right: int | None = None
    hanging: int | None = None
    first_line: int | None = None

    @classmethod
    def of(cls, lvl: Any) -> Indent | None:
        """The ``w:ind`` of a ``w:lvl``'s ``w:pPr``, or None where it states none."""
        if lvl is None:
            return None
        p_pr = lvl.find(_w("pPr"))
        ind = None if p_pr is None else p_pr.find(_w("ind"))
        if ind is None:
            return None
        return cls(
            left=_int(ind.get(_w("left")) or ind.get(_w("start"))),
            right=_int(ind.get(_w("right")) or ind.get(_w("end"))),
            hanging=_int(ind.get(_w("hanging"))),
            first_line=_int(ind.get(_w("firstLine"))),
        )

    def to_dict(self) -> dict[str, int]:
        """The JSON-ready view: the attributes that are stated, in twips."""
        return {
            name: value
            for name, value in dataclasses.asdict(self).items()
            if value is not None
        }


# ---------------------------------------------------------------------------
# the counters
# ---------------------------------------------------------------------------


class Counter:
    """One level's count in one :class:`NumberingState`. docx4j ``ListLevel.Counter``."""

    __slots__ = ("encountered_already", "reset_pending", "value")

    def __init__(self, value: int = 0) -> None:
        """A counter at `value`, not yet used."""
        self.value = value
        #: The level has been counted in this state; its start value is spent.
        self.encountered_already = False
        #: A shallower level reset it: it holds its start value and the next use
        #: takes that value rather than incrementing (docx4j CR-014 probe P8).
        self.reset_pending = False

    def increment(self) -> None:
        """One more item at this level."""
        self.value += 1

    def copy(self) -> Counter:
        """An independent copy; what :meth:`NumberingState.copy` is made of."""
        other = Counter(self.value)
        other.encountered_already = self.encountered_already
        other.reset_pending = self.reset_pending
        return other

    def __repr__(self) -> str:
        """``<Counter 3>``, with the two flags when they are set."""
        flags = "" if self.encountered_already else " (unused)"
        return f"<Counter {self.value}{flags}{' (reset)' if self.reset_pending else ''}>"


class NumberingState:
    """The counters of one story of one traversal. docx4j ``NumberingState``.

    Not thread-safe and not shared: one state belongs to one walk of one story,
    so that two exports at once do not interleave and a footer's list does not
    continue the body's.
    """

    __slots__ = ("_counters", "_start_overrides")

    def __init__(self) -> None:
        """An empty state: every list starts again."""
        self._counters: dict[str, Counter] = {}
        self._start_overrides: set[str] = set()

    def counter(self, abstract_num_id: str, ilvl: str, initial: int | None) -> Counter:
        """The counter of a level of an abstract list, created at `initial`."""
        key = f"{abstract_num_id}/{ilvl}"
        found = self._counters.get(key)
        if found is None:
            found = Counter(0 if initial is None else initial)
            self._counters[key] = found
        return found

    def start_override_applied(self, num_id: str, ilvl: str) -> bool:
        """Whether this ``w:num``'s ``w:startOverride`` has been spent at a level."""
        return f"{num_id}/{ilvl}" in self._start_overrides

    def mark_start_override_applied(self, num_id: str, ilvl: str) -> None:
        """Remember that it has."""
        self._start_overrides.add(f"{num_id}/{ilvl}")

    def reset(self) -> None:
        """Every list starts again, as at the head of a story."""
        self._counters.clear()
        self._start_overrides.clear()

    @property
    def is_empty(self) -> bool:
        """Nothing has been numbered in this state yet."""
        return not self._counters and not self._start_overrides

    def copy(self) -> NumberingState:
        """An independent copy: what :meth:`Emulator.peek` numbers against."""
        other = NumberingState()
        other._counters = {key: counter.copy() for key, counter in self._counters.items()}
        other._start_overrides = set(self._start_overrides)
        return other

    def __repr__(self) -> str:
        """``<NumberingState 2 counters>``."""
        return f"<NumberingState {len(self._counters)} counters>"


class NumberingStates:
    """One :class:`NumberingState` per story. docx4j ``NumberingStates``.

    Word numbers each story from its own counters (docx4j CR-014 probe P7,
    measured): the body is one story; a section's header and footer share one;
    the footnotes part is one, the endnotes part another, the comments part a
    third; and the body's count runs past all of them untouched.
    """

    __slots__ = ("_by_part", "_headers_footers", "_main")

    def __init__(self) -> None:
        """A state for the main story, and the rest made on demand."""
        self._main = NumberingState()
        self._headers_footers: NumberingState | None = None
        self._by_part: dict[int, NumberingState] = {}

    @property
    def main(self) -> NumberingState:
        """The main document's story."""
        return self._main

    def for_part(self, part: Any) -> NumberingState:
        """The story a part's paragraphs number in."""
        kind = _story_kind(part)
        if kind == "main":
            return self._main
        if kind == "headers_footers":
            if self._headers_footers is None:
                self._headers_footers = NumberingState()
            return self._headers_footers
        key = id(part)
        found = self._by_part.get(key)
        if found is None:
            found = NumberingState()
            self._by_part[key] = found
        return found

    @staticmethod
    def new_story() -> NumberingState:
        """A fresh story: a text box's, or one comment's."""
        return NumberingState()


def _story_kind(part: Any) -> str:
    """``"main"``, ``"headers_footers"`` or the part's own story."""
    name = type(getattr(part, "_wrapped", part)).__name__
    if name in ("HeaderPart", "FooterPart"):
        return "headers_footers"
    if name in ("FootnotesPart", "EndnotesPart", "CommentsPart"):
        return name
    return "main"


# ---------------------------------------------------------------------------
# the definitions
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True, slots=True)
class StyleNumbering:
    """What a paragraph style contributes to numbering: its ``w:numPr``, its base.

    Read from ``/word/styles.xml`` without unmarshalling it. `num_id` and `ilvl`
    are what the style itself states; :meth:`Emulator.effective` follows
    `based_on` for whichever of them it does not.
    """

    style_id: str
    based_on: str | None = None
    num_id: str | None = None
    ilvl: str | None = None
    default: bool = False
    ind: Indent | None = None


class ListLevel:
    """One ``w:lvl``. docx4j ``ListLevel``; the count lives in a state."""

    __slots__ = (
        "abstract_element",
        "abstract_num_id",
        "font",
        "has_start_override",
        "id",
        "ind",
        "is_bullet",
        "is_lgl",
        "level_text",
        "lvl_restart",
        "num_fmt",
        "override_element",
        "owner_num_id",
        "p_style",
        "start_value",
    )

    def __init__(self, element: Any = None, ilvl: str | None = None) -> None:
        """Read a level of an abstract definition from its ``w:lvl``."""
        #: The ``w:abstractNum/w:lvl`` this level was read from.
        self.abstract_element = element
        #: The instance's ``w:lvlOverride/w:lvl`` for this level, or None.
        self.override_element: Any = None
        #: ``w:ilvl``, as a string ("0" to "8").
        self.id = ilvl if ilvl is not None else (element.get(_w("ilvl")) if element is not None else "0")
        #: The referencing ``w:abstractNum``: the counter's key.
        self.abstract_num_id = ""
        #: The ``w:num`` an instance level belongs to; None for an abstract one.
        self.owner_num_id: str | None = None
        #: ``w:start`` **less one**, since the first item increments it.
        self.start_value = 0
        self.lvl_restart: int | None = None
        self.level_text: str | None = None
        self.font: str | None = None
        self.num_fmt: str | None = None
        self.is_bullet = False
        self.is_lgl = False
        self.p_style: str | None = None
        self.ind: Indent | None = None
        self.has_start_override = False
        if element is not None:
            self._read(element)

    def _read(self, element: Any) -> None:
        """Take from a ``w:lvl`` whatever it states; leave the rest alone."""
        start = _int(_val(element, "start"))
        if start is not None:
            # one less than the user set it to, since every fetch increments
            self.start_value = start - 1
        restart = _int(_val(element, "lvlRestart"))
        if restart is not None:
            self.lvl_restart = restart
        level_text = _val(element, "lvlText")
        if level_text is not None:
            self.level_text = level_text
        r_pr = element.find(_w("rPr"))
        fonts = None if r_pr is None else r_pr.find(_w("rFonts"))
        if fonts is not None and fonts.get(_w("hAnsi")):
            self.font = fonts.get(_w("hAnsi"))
        num_fmt = _val(element, "numFmt")
        if num_fmt is not None:
            self.num_fmt = num_fmt
            self.is_bullet = num_fmt == "bullet"
        p_style = _val(element, "pStyle")
        if p_style is not None:
            self.p_style = p_style
        if _on(element, "isLgl"):
            self.is_lgl = True
        ind = Indent.of(element)
        if ind is not None:
            self.ind = ind

    def copy(self) -> ListLevel:
        """A copy for an instance definition. docx4j's copy constructor."""
        other = ListLevel.__new__(ListLevel)
        for name in ListLevel.__slots__:
            setattr(other, name, getattr(self, name))
        return other

    def set_overrides(self, element: Any) -> None:
        """Apply this instance's ``w:lvlOverride/w:lvl``. docx4j ``setOverrides``."""
        self.override_element = element
        self._read(element)

    def set_start_value(self, value: int) -> None:
        """A ``w:startOverride``: applied on the first use of this ``w:num``."""
        self.start_value = value
        self.has_start_override = True

    @property
    def ilvl(self) -> int:
        """:attr:`id` as an ``int``."""
        try:
            return int(self.id)
        except (TypeError, ValueError):  # pragma: no cover - a malformed w:ilvl
            return 0

    def counter(self, state: NumberingState) -> Counter:
        """This level's counter in the given state."""
        return state.counter(self.abstract_num_id, self.id, self.start_value)

    def increment_counter(self, state: NumberingState) -> None:
        """One more item at this level. docx4j ``incrementCounter(state)``.

        The shared counter takes this level's start value the first time it is
        met in the state, and again the first time this instance level's
        ``w:num`` is met there when that ``w:num`` overrides the start ---
        deferred until then, since otherwise earlier numbering over the same
        abstract list would use it.
        """
        counter = self.counter(state)
        override_pending = (
            self.has_start_override
            and self.owner_num_id is not None
            and not state.start_override_applied(self.owner_num_id, self.id)
        )
        if override_pending or not counter.encountered_already:
            counter.value = self.start_value
            counter.encountered_already = True
            counter.reset_pending = False
            if self.owner_num_id is not None:
                state.mark_start_override_applied(self.owner_num_id, self.id)
        if counter.reset_pending:
            # the reset already placed the counter at its start value
            counter.reset_pending = False
            return
        counter.increment()

    def reset_counter(self, state: NumberingState) -> None:
        """A shallower level was used: show the start value until next used.

        docx4j CR-014 probe P8: before this the counter went to start - 1 and a
        deeper label printed it, so a level-2 item straight after a level-0 one
        read "2.0.1" where Word reads "2.1.1".
        """
        counter = self.counter(state)
        counter.value = self.start_value + 1
        counter.reset_pending = True

    def restarts_after(self, shallower_ilvl: int) -> bool:
        """Whether using `shallower_ilvl` restarts this level (ECMA-376 17.9.11).

        Without ``w:lvlRestart`` any shallower level restarts it; ``w:val="0"``
        means none does; ``w:val="n"`` means levels 1..n (so ilvl 0..n-1) do and
        deeper ones do not. Measured, docx4j CR-014 probe P8.
        """
        if self.lvl_restart is None:
            return True
        if self.lvl_restart <= 0:
            return False
        return shallower_ilvl <= self.lvl_restart - 1

    def current_value_formatted(self, state: NumberingState, where: str | None = None) -> str:
        """The count at this level, in this level's ``w:numFmt``."""
        return format_value(self.num_fmt, self.counter(state).value, where)

    def current_value_unformatted(self, state: NumberingState) -> str:
        """The count at this level, as a decimal (what ``w:isLgl`` shows)."""
        return str(self.counter(state).value)

    def __repr__(self) -> str:
        """``<ListLevel 0 decimal '%1.'>``."""
        return f"<ListLevel {self.id} {self.num_fmt} {self.level_text!r}>"


class AbstractListNumberingDefinition:
    """One ``w:abstractNum``. docx4j ``AbstractListNumberingDefinition``."""

    __slots__ = ("element", "id", "levels", "linked_style_id", "style_link")

    def __init__(self, element: Any) -> None:
        """Read the definition and its levels from a ``w:abstractNum``."""
        #: The ``w:abstractNum``.
        self.element = element
        #: ``w:abstractNumId``.
        self.id = element.get(_w("abstractNumId")) or ""
        #: The levels, keyed by ``w:ilvl`` ("0" to "8").
        self.levels: dict[str, ListLevel] = {}
        #: ``w:numStyleLink``: the numbering style this takes its levels from.
        self.linked_style_id = _val(element, "numStyleLink")
        #: ``w:styleLink``: the inverse pointer, read and not acted on, as docx4j.
        self.style_link = _val(element, "styleLink")
        self._read_levels(element)

    def _read_levels(self, element: Any) -> None:
        for lvl in element.findall(_w("lvl")):
            level = ListLevel(lvl)
            # the REFERENCING abstract list: a w:numStyleLink definition counts
            # on its own (docx4j CR-014 probe P2)
            level.abstract_num_id = self.id
            self.levels[level.id] = level

    def update_from_linked_style(self, linked: AbstractListNumberingDefinition) -> None:
        """The second pass for a ``w:numStyleLink`` definition: take its levels."""
        if not self.has_linked_style:
            return
        self._read_levels(linked.element)

    @property
    def has_linked_style(self) -> bool:
        """Whether this definition carries ``w:numStyleLink``."""
        return bool(self.linked_style_id)

    @property
    def level_count(self) -> int:
        """How many levels are defined; 0 for an unresolved linked definition."""
        return len(self.levels)

    def __repr__(self) -> str:
        """``<AbstractListNumberingDefinition 2, 9 levels>``."""
        return f"<AbstractListNumberingDefinition {self.id}, {self.level_count} levels>"


class ListNumberingDefinition:
    """One ``w:num``. docx4j ``ListNumberingDefinition``."""

    __slots__ = ("abstract_definition", "element", "levels", "num_id")

    def __init__(
        self,
        element: Any,
        abstract_definitions: dict[str, AbstractListNumberingDefinition],
        resolve_linked_style: bool = False,
    ) -> None:
        """Build the instance from its ``w:num`` and the abstract definitions."""
        #: The ``w:num``.
        self.element = element
        #: ``w:numId``.
        self.num_id = element.get(_w("numId")) or ""
        #: The abstract definition this names, or None.
        self.abstract_definition: AbstractListNumberingDefinition | None = None
        #: The levels, with this instance's overrides applied.
        self.levels: dict[str, ListLevel] = {}

        abstract_id = _val(element, "abstractNumId")
        if abstract_id is None:
            return
        abstract = abstract_definitions.get(abstract_id)
        self.abstract_definition = abstract
        if abstract is None:
            return
        if abstract.level_count == 0 and abstract.has_linked_style and not resolve_linked_style:
            # resolved on the second pass, once the numbering style is known
            return

        for ilvl, level in abstract.levels.items():
            instance = level.copy()
            instance.owner_num_id = self.num_id
            self.levels[ilvl] = instance

        for override in element.findall(_w("lvlOverride")):
            ilvl = override.get(_w("ilvl"))
            if ilvl is None:
                continue
            level = self.levels.get(ilvl)
            start_override = override.find(_w("startOverride"))
            start = None if start_override is None else _int(start_override.get(_w("val")))
            if start is not None and level is not None:
                level.set_start_value(start - 1)
            lvl = override.find(_w("lvl"))
            if lvl is not None and level is not None:
                level.set_overrides(lvl)

    # -- reading -----------------------------------------------------------

    def level(self, ilvl: str | int) -> ListLevel | None:
        """The level with this ``w:ilvl``, or None."""
        return self.levels.get(str(ilvl))

    def level_exists(self, ilvl: str | int) -> bool:
        """Whether the level exists: Word writes ``w:num``\\ s whose levels do not."""
        return str(ilvl) in self.levels

    def is_bullet(self, ilvl: str | int) -> bool:
        """Whether the level's ``w:numFmt`` is ``bullet``."""
        level = self.level(ilvl)
        return bool(level is not None and level.is_bullet)

    def font(self, ilvl: str | int) -> str | None:
        """The ``w:hAnsi`` font the level's ``w:rPr`` names, or None."""
        level = self.level(ilvl)
        return None if level is None else level.font

    # -- counting ----------------------------------------------------------

    def increment_counter(self, ilvl: str | int, state: NumberingState) -> None:
        """One more item at a level, and the deeper levels that restart reset."""
        level_int = int(ilvl)
        this_level = self.level(level_int)
        if this_level is None:  # pragma: no cover - guarded by level_exists
            return

        if not this_level.counter(state).encountered_already:
            # make sure the shallower levels have been initialised
            for shallower in range(level_int - 1, -1, -1):
                level = self.level(shallower)
                if level is None or level.counter(state).encountered_already:
                    break
                level.increment_counter(state)

        this_level.increment_counter(state)

        for deeper_ilvl in range(level_int + 1, LEVELS):
            deeper = self.level(deeper_ilvl)
            if deeper is None:
                break
            if deeper.restarts_after(level_int):
                deeper.reset_counter(state)

    def current_number_string(self, ilvl: str | int, state: NumberingState) -> str:
        """The label of a level: its ``w:lvlText`` with the counters filled in.

        ``w:isLgl`` at any level shows the levels this number *inherits* in
        decimal whatever their own ``w:numFmt``, the level carrying it keeping
        its own: Word's Article / Section numbering prints "Section 1.01".
        """
        controlling = self.level(ilvl)
        if controlling is None:  # pragma: no cover - guarded by level_exists
            return ""
        is_legal = controlling.is_lgl
        try:
            this_level = int(ilvl)
        except (TypeError, ValueError):  # pragma: no cover - a malformed w:ilvl
            this_level = -1

        format_string = controlling.level_text or ""
        out: list[str] = []
        index = 0
        while index < len(format_string):
            char = format_string[index]
            if char == "%" and index < len(format_string) - 1:
                digit = format_string[index + 1]
                if digit.isdigit():
                    level_id = int(digit) - 1  # the format string is 1-based
                    level = self.level(level_id)
                    if level is not None:
                        if is_legal and level_id < this_level:
                            out.append(level.current_value_unformatted(state))
                        else:
                            out.append(
                                level.current_value_formatted(state, f"numId {self.num_id}")
                            )
                    index += 2
                    continue
            out.append(char)
            index += 1
        return "".join(out)

    def __repr__(self) -> str:
        """``<ListNumberingDefinition 3 over abstract 2>``."""
        abstract = None if self.abstract_definition is None else self.abstract_definition.id
        return f"<ListNumberingDefinition {self.num_id} over abstract {abstract}>"


# ---------------------------------------------------------------------------
# reading a whole numbering part
# ---------------------------------------------------------------------------


def read_definitions(
    root: Any, styles: dict[str, StyleNumbering] | None = None
) -> tuple[dict[str, AbstractListNumberingDefinition], dict[str, ListNumberingDefinition]]:
    """The two maps of a ``w:numbering``. docx4j ``initialiseMaps``.

    Two passes, as docx4j's: the first reads every ``w:abstractNum`` and every
    ``w:num``; when an abstract definition carries ``w:numStyleLink`` the second
    resolves it (``resolveLinkedAbstractNum``) --- through the numbering style,
    to the ``w:num`` its ``w:numPr`` names, to *that* ``w:num``'s abstract
    definition, whose levels it takes while counting under its own id.

    Args:
        root: the ``w:numbering`` element, as lxml.
        styles: the paragraph styles' numbering, for ``w:numStyleLink``.

    Returns:
        ``(abstract definitions by w:abstractNumId, instances by w:numId)``.
    """
    abstracts: dict[str, AbstractListNumberingDefinition] = {}
    instances: dict[str, ListNumberingDefinition] = {}
    if root is None:
        return abstracts, instances

    needs_second_pass = False
    for element in root.findall(_w("abstractNum")):
        definition = AbstractListNumberingDefinition(element)
        abstracts[definition.id] = definition
        if definition.has_linked_style:
            needs_second_pass = True

    for element in root.findall(_w("num")):
        definition = ListNumberingDefinition(element, abstracts, False)
        instances[definition.num_id] = definition

    if not needs_second_pass:
        return abstracts, instances

    for definition in abstracts.values():
        _resolve_linked(definition, abstracts, instances, styles or {})
    for element in root.findall(_w("num")):
        definition = ListNumberingDefinition(element, abstracts, True)
        instances[definition.num_id] = definition
    return abstracts, instances


def _resolve_linked(
    definition: AbstractListNumberingDefinition,
    abstracts: dict[str, AbstractListNumberingDefinition],
    instances: dict[str, ListNumberingDefinition],
    styles: dict[str, StyleNumbering],
) -> None:
    """docx4j ``resolveLinkedAbstractNum``: the style, its ``w:num``, its abstract."""
    if not definition.has_linked_style:
        return
    style = styles.get(definition.linked_style_id or "")
    if style is None or style.num_id is None:
        return
    instance = instances.get(style.num_id)
    linked = None if instance is None else instance.abstract_definition
    if linked is None or linked is definition:
        return
    definition.update_from_linked_style(linked)
