"""The numbering emulator: the label Word would paint. docx4j ``Emulator``.

CR-002 section 6.2, delivered early through CR-003 Phase H because the list
views and the markdown exporter both need labels (CR-003 section 18). Given a
paragraph's ``w:pPr`` and a :class:`~docx4j_py.model.listnumbering.NumberingState`
it answers with a :class:`NumberingResult`: the label (``"1."``, ``"a)"``,
``"2.1.4."``, a bullet glyph), the font the level names, whether it is a bullet
and the indent the level contributes.

docx4j's ``Emulator`` resolves a paragraph's numbering through
``PropertyResolver``, which is CR-002 Phase B and not yet here. This resolves it
directly instead, over the rules CR-014 measured against Word:

* the paragraph's own ``w:numPr`` first, and a ``w:numId`` of **0** turns
  numbering off (Word's convention);
* otherwise the paragraph style's, following ``w:basedOn`` attribute by
  attribute, so a style stating only ``w:ilvl`` takes its ``w:numId`` from the
  style it is based on (probe ``styles-numpr-ilvl-only``, measured);
* a paragraph naming **no** style is resolved against the ``w:default="1"``
  paragraph style, which may itself be numbered (probe
  ``numbering-default-style-numbered``, measured);
* a ``w:ilvl`` the paragraph states directly is kept even when the ``w:numId``
  comes from its style;
* a level whose ``w:pStyle`` names a *different* paragraph style than the one
  that brought the numbering paints no label and is not counted (probe
  ``numbering-label-ilvl0``), unless the ``w:numPr`` is the paragraph's own,
  because direct formatting always applies.

When ``PropertyResolver`` lands this resolution moves onto it: the style walk
here is the one thing in the module that is not docx4j's own code.

Nothing here unmarshals a part. The numbering definitions and the paragraph
styles are read with lxml from the bytes the parts would be saved as, so a
document whose list labels are only *read* is still written back byte for byte
(CR-003 section 13 made that promise for the markdown exporter, which now reads
its markers from here).
"""

from __future__ import annotations

import dataclasses
from typing import Any

from docx4j_py.model.listnumbering.definitions import (
    AbstractListNumberingDefinition,
    Indent,
    ListLevel,
    ListNumberingDefinition,
    NumberingState,
    NumberingStates,
    StyleNumbering,
    read_definitions,
)

__all__ = [
    "Emulator",
    "ItemLabel",
    "NumRef",
    "NumberingResult",
    "ResultTriple",
    "emulator_of",
    "invalidate",
    "invalidate_labels",
    "labels_for",
]

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _w(local: str) -> str:
    return f"{_W}{local}"


def _root_of(part: Any) -> Any:
    """A part's root element as lxml, from the bytes it would be saved as.

    The trick :mod:`docx4j_py.model.content.describe` and the markdown exporter
    both use: ``get_xml()`` is the part's own marshalling, so an untouched part
    is its source bytes and a part the caller has edited is what it now holds,
    and neither is unmarshalled by being read here.
    """
    if part is None:
        return None
    from lxml import etree

    try:
        return etree.fromstring(part.get_xml().encode("utf-8"))
    except Exception:  # noqa: BLE001 - a part that will not parse defines nothing
        return None


# ---------------------------------------------------------------------------
# what a paragraph is numbered with
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True, slots=True)
class NumberingResult:
    """The label a paragraph gets. docx4j ``Emulator.NumberingResult``.

    docx4j's four (``numString``, ``numFont``, ``bullet``, ``ind``) with three
    extensions this package needs: which list and level the label came from, and
    the item's index among its siblings, which Office JS's ``ListItem`` reports
    and docx4j has no member for.
    """

    num_string: str
    """The label: ``"1."``, ``"a)"``, ``"2.1.4."``, or the bullet's glyph."""
    num_font: str | None = None
    """The ``w:hAnsi`` font the level's ``w:rPr`` names, or None."""
    is_bullet: bool = False
    """Whether the level's ``w:numFmt`` is ``bullet``."""
    ind: Indent | None = None
    """The level's own ``w:ind``: the instance's where it states one, else the
    abstract level's."""
    num_id: str = ""
    """The ``w:numId`` the label was counted in. Extension."""
    ilvl: int = 0
    """The ``w:ilvl``. Extension."""
    count: int = 0
    """The counter's value at this level, before formatting. Extension."""
    sibling_index: int = 0
    """0-based index among the items at this level under the same parent item,
    which is the counter less the level's start value. Extension (Office JS
    ``ListItem.siblingIndex``)."""
    level: ListLevel | None = None
    """The level definition the label came from (docx4j ``getLvl``). Extension."""

    @property
    def bullet(self) -> str | None:
        """docx4j's ``getBullet``: the ``w:lvlText`` of a bullet level, else None."""
        return self.num_string if self.is_bullet else None

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view."""
        return {
            "num_string": self.num_string,
            "num_font": self.num_font,
            "is_bullet": self.is_bullet,
            "ind": None if self.ind is None else self.ind.to_dict(),
            "num_id": self.num_id,
            "ilvl": self.ilvl,
            "sibling_index": self.sibling_index,
        }


#: docx4j's older name for :class:`NumberingResult`, kept as CR-002 section 6.2
#: names it.
ResultTriple = NumberingResult


@dataclasses.dataclass(frozen=True, slots=True)
class NumRef:
    """Where a paragraph's numbering comes from. docx4j ``Emulator.NumRef``."""

    num_id: str | None = None
    """The list, or None when not numbered."""
    ilvl: str | None = None
    """The level; ``"0"`` where nothing states one."""
    direct: bool = False
    """Whether the ``w:numId`` is the paragraph's own, not its style's."""
    not_numbered: bool = False
    """Word paints no label and does not count the paragraph."""
    reason: str | None = None
    """Why not; None otherwise."""

    def __str__(self) -> str:
        """``numId 3 ilvl 0 (direct)``, or ``not numbered: ...``."""
        if self.not_numbered:
            return f"not numbered: {self.reason}"
        where = "direct" if self.direct else "from style"
        return f"numId {self.num_id} ilvl {self.ilvl} ({where})"


@dataclasses.dataclass(frozen=True, slots=True)
class ItemLabel:
    """One list paragraph of a story, as :func:`labels_for` counted it."""

    element: Any
    """The ``w:p``; held so that the cache keeps its key alive."""
    result: NumberingResult

    @property
    def num_id(self) -> str:
        """The ``w:numId`` the item counts in."""
        return self.result.num_id

    @property
    def ilvl(self) -> int:
        """The level."""
        return self.result.ilvl


# ---------------------------------------------------------------------------
# the emulator
# ---------------------------------------------------------------------------


class Emulator:
    """The numbering of one package: the definitions, and the counting.

    docx4j's ``Emulator`` is static over maps the numbering part holds; here one
    object holds both, because the style resolution it does without
    ``PropertyResolver`` needs the package too. The four entry points are
    docx4j's, as static methods over a package:
    :meth:`get_number`, :meth:`get_number_of`, :meth:`peek` and
    :meth:`Emulator.of`.
    """

    __slots__ = (
        "_abstract",
        "_default_style_id",
        "_instances",
        "_labels",
        "_numbering_state",
        "_package",
        "_read",
        "_styles",
    )

    def __init__(self, package: Any) -> None:
        """Build the emulator over a package; the definitions are read on demand."""
        self._package = package
        self._abstract: dict[str, AbstractListNumberingDefinition] = {}
        self._instances: dict[str, ListNumberingDefinition] = {}
        self._styles: dict[str, StyleNumbering] | None = None
        self._default_style_id: str | None = None
        self._read = False
        self._numbering_state = NumberingState()
        #: The story label caches, by ``id`` of the story's container.
        self._labels: dict[int, dict[int, ItemLabel]] = {}

    # -- the definitions ---------------------------------------------------

    @property
    def package(self) -> Any:
        """The package this emulator is over."""
        return self._package

    @property
    def part(self) -> Any:
        """``/word/numbering.xml``, or None when the document has no lists."""
        return getattr(self._package, "numbering_definitions_part", None)

    def refresh(self) -> Emulator:
        """Read the definitions again: what a change to the numbering part needs."""
        self._read = False
        self._abstract = {}
        self._instances = {}
        self._styles = None
        self._default_style_id = None
        self._labels.clear()
        return self

    def _ensure(self) -> None:
        if self._read:
            return
        self._read = True
        root = _root_of(self.part)
        self._abstract, self._instances = read_definitions(root, self.styles)

    @property
    def abstract_list_definitions(self) -> dict[str, AbstractListNumberingDefinition]:
        """``w:abstractNumId`` -> the definition. docx4j ``getAbstractListDefinitions``."""
        self._ensure()
        return self._abstract

    @property
    def instance_list_definitions(self) -> dict[str, ListNumberingDefinition]:
        """``w:numId`` -> the definition. docx4j ``getInstanceListDefinitions``."""
        self._ensure()
        return self._instances

    @property
    def numbering_state(self) -> NumberingState:
        """The counters the state-less calls use. docx4j ``getNumberingState``."""
        return self._numbering_state

    # -- the styles, read without unmarshalling ----------------------------

    @property
    def styles(self) -> dict[str, StyleNumbering]:
        """Every paragraph style's ``w:numPr`` and ``w:basedOn``, by style id."""
        if self._styles is None:
            self._styles, self._default_style_id = _read_styles(self._package)
        return self._styles

    @property
    def default_paragraph_style_id(self) -> str | None:
        """The ``w:default="1"`` paragraph style's id, which may be numbered."""
        if self._styles is None:
            self._styles, self._default_style_id = _read_styles(self._package)
        return self._default_style_id

    def effective_numbering(self, style_id: str) -> StyleNumbering:
        """A style's ``w:numPr``, following ``w:basedOn`` attribute by attribute.

        What ``PropertyResolver.getEffectivePPr(styleId).getNumPr()`` answers in
        docx4j: a style stating only ``w:ilvl`` takes the ``w:numId`` of the
        style it is based on (probe ``styles-numpr-ilvl-only``, case (d)).
        """
        styles = self.styles
        seen: set[str] = set()
        num_id: str | None = None
        ilvl: str | None = None
        current: str | None = style_id
        while current is not None and current not in seen:
            seen.add(current)
            style = styles.get(current)
            if style is None:
                break
            if num_id is None:
                num_id = style.num_id
            if ilvl is None:
                ilvl = style.ilvl
            current = style.based_on
        return StyleNumbering(style_id, None, num_id, ilvl)

    def linked_style_id(self, num_id: str, ilvl: str | None) -> str | None:
        """The ``w:pStyle`` a level names (ECMA-376 17.9.24), or None.

        docx4j ``NumberingDefinitionsPart.getLinkedStyleId``: the instance's
        ``w:lvlOverride/w:lvl`` where it has one, the abstract level otherwise.
        """
        definition = self.instance_list_definitions.get(num_id)
        if definition is None:
            return None
        level = definition.level(ilvl if ilvl else "0")
        return None if level is None else level.p_style

    def style_linked_elsewhere(self, num_id: str, ilvl: str, style_id: str | None) -> bool:
        """Whether the level belongs to a different style than brought the numbering.

        docx4j ``Emulator.styleLinkedElsewhere``, measured on probe
        ``numbering-label-ilvl0``: Word paints no label and does not count such
        a paragraph. The style itself or any style it is based on counts as the
        level's own.
        """
        linked = self.linked_style_id(num_id, ilvl)
        if not linked:
            return False
        if style_id is None:
            return True
        styles = self.styles
        seen: set[str] = set()
        current: str | None = style_id
        while current is not None and current not in seen:
            seen.add(current)
            if linked == current:
                return False
            style = styles.get(current)
            current = None if style is None else style.based_on
        return True

    # -- resolution --------------------------------------------------------

    def resolve(
        self,
        p_style_val: str | None,
        num_id: str | None,
        ilvl: str | None,
        direct_num_pr: bool,
    ) -> NumRef:
        """Which list and level to count a paragraph in. docx4j ``Emulator.resolve``."""
        style_id = p_style_val or None

        if not num_id:
            direct_num_pr = False
            if style_id is None:
                style_id = self.default_paragraph_style_id
                if style_id is None:
                    return NumRef(
                        not_numbered=True,
                        reason="no numId, no paragraph style and no default paragraph style",
                    )
            effective = self.effective_numbering(style_id)
            if effective.num_id is None:
                return NumRef(
                    not_numbered=True,
                    reason=f"no numId, and style {style_id!r} is not numbered",
                )
            num_id = effective.num_id
            if not ilvl:
                # w:ilvl (and its w:val) is optional; its absence means level 0
                ilvl = effective.ilvl or "0"

        if num_id == "0":
            # Word's convention: w:numId 0 turns numbering off for the paragraph
            return NumRef(not_numbered=True, reason="w:numId 0 turns numbering off")

        if not ilvl:
            ilvl = "0"

        if not direct_num_pr and self.style_linked_elsewhere(num_id, ilvl, style_id):
            return NumRef(
                not_numbered=True,
                reason=(
                    f"level {ilvl} of numId {num_id} is linked to a paragraph style "
                    f"other than {style_id!r}"
                ),
            )
        return NumRef(num_id=num_id, ilvl=ilvl, direct=direct_num_pr)

    # -- counting ----------------------------------------------------------

    def number(self, p_pr: Any, state: NumberingState | None = None) -> NumberingResult | None:
        """The next number for a paragraph, from its ``w:pPr``. Increments `state`."""
        if p_pr is None:
            return None
        p_style_val = _value(getattr(p_pr, "p_style", None))
        num_pr = getattr(p_pr, "num_pr", None)
        num_id = None
        ilvl = None
        if num_pr is not None:
            num_id = _value(getattr(num_pr, "num_id", None))
            ilvl = _value(getattr(num_pr, "ilvl", None))
        num_id = None if num_id is None else str(num_id)
        ilvl = None if ilvl is None else str(ilvl)
        return self.number_of(
            p_style_val, num_id, ilvl, bool(num_id), state
        )

    def number_of(
        self,
        p_style_val: str | None,
        num_id: str | None,
        ilvl: str | None,
        direct_num_pr: bool = False,
        state: NumberingState | None = None,
    ) -> NumberingResult | None:
        """The next number, given the style and the ``w:numPr``'s two values."""
        if self.part is None:
            return None
        if state is None:
            state = self._numbering_state
        reference = self.resolve(p_style_val, num_id, ilvl, direct_num_pr)
        if reference.not_numbered:
            return None
        definition = self.instance_list_definitions.get(reference.num_id or "")
        if definition is None or not definition.level_exists(reference.ilvl or "0"):
            # Word writes a w:num whose abstract definition is missing, or lacks
            # the level; docx4j logs and returns an empty result, this says None
            return None

        level_id = reference.ilvl or "0"
        definition.increment_counter(level_id, state)
        label = definition.current_number_string(level_id, state)
        level = definition.level(level_id)
        assert level is not None  # level_exists said so
        counter = level.counter(state)
        return NumberingResult(
            num_string=label,
            num_font=level.font or None,
            is_bullet=level.is_bullet,
            ind=level.ind,
            num_id=definition.num_id,
            ilvl=level.ilvl,
            count=counter.value,
            sibling_index=max(0, counter.value - level.start_value - 1),
            level=level,
        )

    def peek_number(
        self, p_pr: Any, state: NumberingState | None = None
    ) -> NumberingResult | None:
        """What :meth:`number` would return, leaving the state as it was."""
        if state is None:
            state = self._numbering_state
        return self.number(p_pr, state.copy())

    def indent_for(self, num_id: str, ilvl: str | int | None) -> Indent | None:
        """The indent a level contributes. docx4j ``NumberingDefinitionsPart.getInd``.

        The level's own ``w:pPr/w:ind`` first --- ECMA-376 17.9.24, and measured
        (docx4j CR-014 claim 7) --- and the ``w:pStyle`` it is linked to
        afterwards, following ``w:basedOn`` as every other style property does.
        """
        definition = self.instance_list_definitions.get(num_id)
        if definition is None:
            return None
        level = definition.level("0" if ilvl is None else str(ilvl))
        if level is None:
            return None
        if level.ind is not None:
            return level.ind
        if not level.p_style:
            return None
        styles = self.styles
        seen: set[str] = set()
        current: str | None = level.p_style
        while current is not None and current not in seen:
            seen.add(current)
            style = styles.get(current)
            if style is None:
                return None
            if style.ind is not None:
                return style.ind
            current = style.based_on
        return None

    # -- docx4j's entry points ---------------------------------------------

    @staticmethod
    def of(package: Any) -> Emulator | None:
        """The package's emulator. docx4j ``NumberingDefinitionsPart.getEmulator``."""
        return emulator_of(package)

    @staticmethod
    def get_number(
        package: Any, p_pr: Any, state: NumberingState | None = None
    ) -> NumberingResult | None:
        """docx4j ``Emulator.getNumber(wmlPackage, pPr, state)``."""
        emulator = emulator_of(package)
        return None if emulator is None else emulator.number(p_pr, state)

    @staticmethod
    def get_number_of(
        package: Any,
        p_style_val: str | None,
        num_id: str | None,
        ilvl: str | None,
        direct_num_pr: bool | None = None,
        state: NumberingState | None = None,
    ) -> NumberingResult | None:
        """docx4j's ``getNumber(pkg, pStyleVal, numId, levelId, directNumPr, state)``.

        Python has no overloads, so docx4j's second ``getNumber`` is a second
        name (CR-003 section 18). `direct_num_pr` defaults to "the numId given
        here is the paragraph's own", which is what docx4j's four-argument
        overload assumes.
        """
        emulator = emulator_of(package)
        if emulator is None:
            return None
        if direct_num_pr is None:
            direct_num_pr = bool(num_id)
        return emulator.number_of(p_style_val, num_id, ilvl, direct_num_pr, state)

    @staticmethod
    def peek(
        package: Any, p_pr: Any, state: NumberingState | None = None
    ) -> NumberingResult | None:
        """docx4j ``Emulator.peek``: the number, without taking it."""
        emulator = emulator_of(package)
        return None if emulator is None else emulator.peek_number(p_pr, state)

    def __repr__(self) -> str:
        """``<Emulator 3 lists, 2 definitions>``."""
        return (
            f"<Emulator {len(self.instance_list_definitions)} lists, "
            f"{len(self.abstract_list_definitions)} definitions>"
        )


def _value(obj: Any) -> Any:
    """``w:val`` as a plain value, whether the model made it an enum or not."""
    val = getattr(obj, "val", None)
    return getattr(val, "value", val)


def _read_styles(package: Any) -> tuple[dict[str, StyleNumbering], str | None]:
    """Every paragraph style's numbering, and the default style's id.

    Read from ``/word/styles.xml`` with lxml, as
    :func:`docx4j_py.model.content.styles.style_ids_of` reads it, so that
    reading a document's list labels leaves the styles part untouched.
    """
    part = getattr(package, "style_definitions_part", None)
    root = _root_of(part)
    styles: dict[str, StyleNumbering] = {}
    default_id: str | None = None
    if root is None:
        return styles, default_id
    for style in root.findall(_w("style")):
        kind = style.get(_w("type")) or "paragraph"
        # a numbering style is here too: ``w:numStyleLink`` names one, and its
        # own ``w:pPr/w:numPr`` is how the linked list is reached (docx4j
        # ``resolveLinkedAbstractNum``). A paragraph never names one, so one map
        # serves both.
        if kind not in ("paragraph", "numbering"):
            continue
        style_id = style.get(_w("styleId"))
        if not style_id:
            continue
        based_on = None
        num_id = None
        ilvl = None
        ind = None
        p_pr = style.find(_w("pPr"))
        if p_pr is not None:
            num_pr = p_pr.find(_w("numPr"))
            if num_pr is not None:
                num_id_element = num_pr.find(_w("numId"))
                if num_id_element is not None:
                    num_id = num_id_element.get(_w("val"))
                ilvl_element = num_pr.find(_w("ilvl"))
                if ilvl_element is not None:
                    ilvl = ilvl_element.get(_w("val"))
            ind_element = p_pr.find(_w("ind"))
            if ind_element is not None:
                ind = Indent(
                    left=_as_int(ind_element.get(_w("left")) or ind_element.get(_w("start"))),
                    right=_as_int(ind_element.get(_w("right")) or ind_element.get(_w("end"))),
                    hanging=_as_int(ind_element.get(_w("hanging"))),
                    first_line=_as_int(ind_element.get(_w("firstLine"))),
                )
        based_on_element = style.find(_w("basedOn"))
        if based_on_element is not None:
            based_on = based_on_element.get(_w("val"))
        default = (style.get(_w("default")) or "") in ("1", "true", "on")
        styles[style_id] = StyleNumbering(style_id, based_on, num_id, ilvl, default, ind)
        if default and default_id is None and kind == "paragraph":
            default_id = style_id
    return styles, default_id


def _as_int(text: str | None) -> int | None:
    if text is None:
        return None
    try:
        return int(text)
    except ValueError:  # pragma: no cover - a malformed indent
        return None


# ---------------------------------------------------------------------------
# one emulator per package, and the story label caches
# ---------------------------------------------------------------------------


def emulator_of(package: Any) -> Emulator | None:
    """The package's :class:`Emulator`, made on first use and kept.

    None for a package with no numbering part *and* no way to hold one --- a
    detached body --- so that a caller can ask any body about its lists.
    """
    if package is None:
        return None
    try:
        emulator = package._numbering_emulator
    except AttributeError:
        return Emulator(package)
    if emulator is None:
        emulator = Emulator(package)
        package._numbering_emulator = emulator
    return emulator


def invalidate(package: Any) -> None:
    """Forget the definitions and the labels: a change to the numbering part.

    What every writer of ``/word/numbering.xml`` calls --- the ``List`` level
    setters, ``start_new_list``, the markdown importer's ``_Numbering`` --- so
    that the next label read builds the definitions again.
    """
    if package is None:
        return
    emulator = getattr(package, "_numbering_emulator", None)
    if emulator is not None:
        emulator.refresh()


def invalidate_labels(package: Any) -> None:
    """Forget the counted labels, keeping the definitions.

    What :class:`docx4j_py.model.content.reports.recording` calls after every
    mutation: a paragraph inserted, deleted or renumbered changes what every
    item after it is numbered, and nothing else can know that centrally.
    """
    if package is None:
        return
    emulator = getattr(package, "_numbering_emulator", None)
    if emulator is not None:
        emulator._labels.clear()


def labels_for(body: Any) -> dict[int, ItemLabel]:
    """Every list paragraph of a story, by ``id`` of its ``w:p``.

    The story is walked **from its start** in document order with a fresh
    :class:`NumberingState`, which is the only way to know what Word would paint
    in front of a paragraph in the middle of it; the answer is memoised per
    story and thrown away by :func:`invalidate_labels` on the next mutation.

    A *story* is docx4j's ``NumberingStates`` (CR-014 probe P7, measured): the
    body is one, every header and footer of the document share one (headers
    first, then footers, which is the order Word's own counts run in), and the
    footnotes, endnotes and comments parts each have their own. A body inside
    the story --- a table cell's, a content control's --- is counted with the
    part's own body, so ``cell.body.lists`` sees the numbers the page shows.
    """
    package = getattr(body, "package", None)
    emulator = emulator_of(package)
    if emulator is None or emulator.part is None:
        return {}
    root = _story_root(body)
    key = id(root.container)
    found = emulator._labels.get(key)
    if found is not None:
        return found

    state = NumberingState()
    for story_body in _story_bodies(root):
        labels: dict[int, ItemLabel] = {}
        for paragraph in story_body.iter_paragraphs():
            element = paragraph.element
            result = emulator.number(getattr(element, "p_pr", None), state)
            if result is not None:
                labels[id(element)] = ItemLabel(element, result)
        emulator._labels[id(story_body.container)] = labels
    return emulator._labels.get(key, {})


def _story_root(body: Any) -> Any:
    """The body of the part `body` belongs to: a cell's story is its part's."""
    part = getattr(body, "part", None)
    if part is None:
        return body
    from docx4j_py.model.content.body import body_of

    try:
        return body_of(part)
    except Exception:  # noqa: BLE001 - a part with no block content is its own story
        return body


def _story_bodies(root: Any) -> list[Any]:
    """The bodies counted together with `root`, in the order Word counts them."""
    part = getattr(root, "part", None)
    name = type(getattr(part, "_wrapped", part)).__name__
    if name not in ("HeaderPart", "FooterPart"):
        return [root]
    from docx4j_py.model.content.body import body_of

    package = getattr(root, "package", None)
    main = getattr(package, "main_document_part", None)
    if main is None:
        return [root]
    out: list[Any] = []
    seen: set[int] = set()
    for other in list(main.header_parts()) + list(main.footer_parts()):
        wrapped = getattr(package, "trial_part", None)
        other = wrapped(other) if wrapped is not None else other
        try:
            story = body_of(other)
        except Exception:  # noqa: BLE001, S112 - a header with no content counts nothing
            continue
        if id(story.container) in seen:
            continue
        seen.add(id(story.container))
        out.append(story)
    if id(root.container) not in seen:
        out.append(root)
    return out


def states_for(package: Any) -> NumberingStates:
    """A fresh set of per-story counters. docx4j ``NumberingStates``."""
    return NumberingStates()
