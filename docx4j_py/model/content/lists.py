"""``List`` and ``ListItem``: Office JS ``Word.List`` over ``w:num``.

CR-003 section 3.10, Phase H. A :class:`List` is a view over one ``w:num`` and
the ``w:abstractNum`` it points at; a :class:`ListItem` is what a paragraph in
one is. The labels are the numbering emulator's
(:mod:`docx4j_py.model.listnumbering`, CR-002 section 6.2), so
``list_item.list_string`` is what **Word** paints in front of the paragraph ---
restarts, ``w:startOverride``, ``w:lvlRestart``, a numbered paragraph style and
all --- and not this module's own count.

The conventions of section 3.1 hold: ``snake_case`` members, Office JS's string
values (``"Number"``, ``"Bullet"``, ``"Arabic"``, ``"Solid"``), points rather
than twips, ``None`` where Office JS would give a null object, and one
:class:`~docx4j_py.model.content.reports.ChangeReport` per mutation.

Three rules this module keeps, each recorded in CR-003 section 18:

* **``w:numPr`` is a paragraph property.** :meth:`Paragraph.attach_to_list`,
  :meth:`Paragraph.detach_from_list` and ``ListItem.level`` all go through
  ``Paragraph._p_pr()``, so with tracking on they record a ``w:pPrChange``,
  which is what Word writes for a numbering change (CR-003 section 16.7).
  ``w:numberingChange`` is never written, and the numbering part is never
  tracked.
* **A level setter never changes another list.** Where a second ``w:num``
  shares the ``w:abstractNum``, the definition is copied first and this list
  points at the copy, so the change stays local (docx4j-core-ts CR-002 3.9).
* **Definitions are created in one place.** ``start_new_list`` and the markdown
  importer's ``_Numbering`` both go through :func:`numbering_part_of` and the
  two id allocators here, which is what keeps section 3.4's determinism and
  lets a :func:`~docx4j_py.model.content.trial.dry_run` un-add a numbering part
  it created (CR-003 section 14.4).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Literal

from docx4j_py.child import ChildList, deep_copy, link_parents
from docx4j_py.model.content.errors import ContentError
from docx4j_py.model.content.reports import recording
from docx4j_py.model.listnumbering import (
    LEVELS,
    Emulator,
    ItemLabel,
    emulator_of,
    invalidate,
    labels_for,
)
from docx4j_py.wml import el

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.body import Body
    from docx4j_py.model.content.paragraph import Paragraph

__all__ = [
    "BULLETS",
    "LIST_PARAGRAPH_STYLE",
    "NUMBERING",
    "List",
    "ListItem",
    "add_abstract_definition",
    "attach_to_list",
    "detach_from_list",
    "list_item_of",
    "list_of",
    "lists_of",
    "next_abstract_num_id",
    "next_num_id",
    "numbering_part_of",
    "start_new_list",
]

#: The style Word applies when it makes a list item from the ribbon. Applied
#: only to a paragraph with no style of its own (CR-003 section 18).
LIST_PARAGRAPH_STYLE = "ListParagraph"

#: Office JS ``Word.ListNumbering`` -> ``w:numFmt``.
NUMBERING: dict[str, str] = {
    "None": "none",
    "Arabic": "decimal",
    "UpperRoman": "upperRoman",
    "LowerRoman": "lowerRoman",
    "UpperLetter": "upperLetter",
    "LowerLetter": "lowerLetter",
}

#: Office JS ``Word.ListBullet`` -> ``(the glyph, the font)`` Word writes for it.
#: The four Wingdings characters and the Symbol one are in the private use area,
#: where Word puts a symbol font's glyphs; ``Hollow`` is a plain letter "o" in
#: Courier New, exactly as docx4j's default ``numbering.xml`` has it.
BULLETS: dict[str, tuple[str, str]] = {
    "Solid": ("", "Symbol"),
    "Hollow": ("o", "Courier New"),
    "Square": ("", "Wingdings"),
    "Diamonds": ("", "Wingdings"),
    "Arrow": ("", "Wingdings"),
    "Checkmark": ("", "Wingdings"),
}

#: A symbol font's private-use bullet -> the Unicode character it looks like.
#: ``list_string`` reports these, because ``""`` is not a bullet to
#: anything but Word (CR-003 section 18); the raw glyph is the emulator's
#: ``NumberingResult.num_string``.
BULLET_GLYPHS: dict[str, str] = {
    "": "•",
    "": "▪",
    "": "❖",
    "": "➢",
    "": "✓",
    "": "■",
    "": "◆",
}

#: Twips per point, and the indents a new level gets.
TWIPS_PER_POINT = 20
TWIPS_PER_LEVEL = 720
HANGING = 360

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _value(obj: Any) -> Any:
    """``w:val`` as a plain value, whether the model made it an enum or not."""
    val = getattr(obj, "val", None)
    return getattr(val, "value", val)


# ---------------------------------------------------------------------------
# the numbering part: one place that finds it, creates it and allocates ids
# ---------------------------------------------------------------------------


def numbering_part_of(
    package: Any, *, create: bool = False, touched: set[str] | None = None
) -> Any:
    """``/word/numbering.xml``, created with its relationship when asked for.

    The **one** place a numbering part is made: ``start_new_list`` and the
    markdown importer's ``_Numbering`` both come here (CR-003 section 13.6).
    A part created during a :func:`~docx4j_py.model.content.trial.dry_run` is
    noted in the trial's undo log, so the trial un-adds it on the way out ---
    which is the gap CR-003 section 14.4 left open.

    Args:
        package: the ``WordprocessingMLPackage``, or a trial one.
        create: make the part when the document has none.
        touched: a set the part name is added to.

    Returns:
        The part (the trial's copy of it inside a trial), or None.
    """
    part = getattr(package, "numbering_definitions_part", None)
    if part is not None:
        if touched is not None:
            touched.add(str(part.part_name))
        return part
    if not create:
        return None

    from docx4j_py.model.content.picture import note_added_part
    from docx4j_py.openpackaging.parts.wml import NumberingDefinitionsPart
    from docx4j_py.wml import Numbering

    new = NumberingDefinitionsPart()
    new.set_contents(Numbering())
    main = package.get_main_document_part()
    had_override = package.content_type_manager.get_override_content_type(new.part_name)
    relationship = main.add_target_part(new)
    note_added_part(package, new, relationship, main, added_content_type=had_override is None)
    if touched is not None:
        touched.add(str(new.part_name))
    invalidate(package)
    wrap = getattr(package, "trial_part", None)
    return wrap(new) if wrap is not None else new


def next_abstract_num_id(numbering: Any) -> int:
    """The next free ``w:abstractNumId``: one above the document's own."""
    return (
        max(
            (int(getattr(a, "abstract_num_id", 0) or 0) for a in (numbering.abstract_num or ())),
            default=-1,
        )
        + 1
    )


def next_num_id(numbering: Any) -> int:
    """The next free ``w:numId``: one above the document's own, never 0.

    ``w:numId`` 0 is Word's "no numbering", so a list never takes it.
    """
    return max((int(getattr(n, "num_id", 0) or 0) for n in (numbering.num or ())), default=0) + 1


def add_abstract_definition(
    package: Any,
    kind: Literal["Number", "Bullet"] = "Number",
    *,
    touched: set[str] | None = None,
) -> Any:
    """A new ``w:abstractNum`` from docx4j's defaults, and the ``w:num`` for it.

    docx4j ``NumberingDefinitionsPart.addAbstractListNumberingDefinition``, over
    ``unmarshalDefaultNumbering``: its ``w:abstractNum`` 1 is the decimal set
    Word's numbered-list gallery writes and its 0 the bullet set.

    Returns:
        The new ``w:num``.
    """
    part = numbering_part_of(package, create=True, touched=touched)
    numbering = part.contents
    if touched is not None:
        touched.add(str(part.part_name))

    abstract = deep_copy(_default_abstract(kind))
    abstract.abstract_num_id = next_abstract_num_id(numbering)
    # the template's w:nsid and w:tmpl identify the gallery entry the list came
    # from; two lists in one document may share them, as Word's own do
    _append(numbering, "abstract_num", abstract)

    num = el.num(
        num_id=next_num_id(numbering),
        abstract_num_id=el.abstractNumId(val=abstract.abstract_num_id),
    )
    _append(numbering, "num", num)
    invalidate(package)
    return num


def restart_over(package: Any, num_id: int, *, touched: set[str] | None = None) -> Any:
    """A new ``w:num`` over the same ``w:abstractNum``, restarting at 1.

    docx4j ``NumberingDefinitionsPart.restart``: Word's *Restart numbering at 1*
    is a second instance of the same definition carrying a ``w:lvlOverride``
    with a ``w:startOverride`` of 1 at level 0 --- which the emulator spends the
    first time the new ``w:num`` is met, exactly as Word does.
    """
    part = numbering_part_of(package, create=False, touched=touched)
    if part is None:
        raise _no_list(package, num_id)
    numbering = part.contents
    existing = _typed_num(numbering, num_id)
    if existing is None:
        raise _no_list(package, num_id)
    if touched is not None:
        touched.add(str(part.part_name))

    num = el.num(
        num_id=next_num_id(numbering),
        abstract_num_id=el.abstractNumId(val=_value(existing.abstract_num_id)),
    )
    override = el.lvlOverride(ilvl=0, start_override=el.startOverride(val=1))
    num.lvl_override = ChildList([override], owner=num)
    _append(numbering, "num", num)
    invalidate(package)
    return num


def _append(numbering: Any, field: str, value: Any) -> None:
    """Append to one of ``w:numbering``'s lists, making it a ``ChildList`` first."""
    items = getattr(numbering, field, None)
    if items is None:
        items = ChildList([], owner=numbering)
        setattr(numbering, field, items)
    items.append(value)
    link_parents(value)
    value.parent = numbering


def _default_abstract(kind: str) -> Any:
    """docx4j's own default definition: the decimal set, or the bullet set."""
    from docx4j_py.openpackaging.parts.wml import NumberingDefinitionsPart

    part = NumberingDefinitionsPart()
    numbering = part.unmarshal_default_numbering()
    wanted = "bullet" if str(kind) == "Bullet" else "decimal"
    for abstract in numbering.abstract_num or ():
        first = next((lvl for lvl in (abstract.lvl or ()) if int(lvl.ilvl or 0) == 0), None)
        fmt = _value(getattr(first, "num_fmt", None)) if first is not None else None
        if (fmt == "bullet") == (wanted == "bullet"):
            return abstract
    raise ContentError(  # pragma: no cover - the resource carries both
        f"docx4j's default numbering has no {wanted} definition",
        code="list.no_default_definition",
        hint="pass kind='Number' or kind='Bullet'",
    )


def _typed_num(numbering: Any, num_id: int | str) -> Any:
    """The ``w:num`` with this ``w:numId``, from the typed tree, or None."""
    for num in numbering.num or ():
        if str(num.num_id) == str(num_id):
            return num
    return None


def _typed_abstract(numbering: Any, abstract_id: Any) -> Any:
    """The ``w:abstractNum`` with this ``w:abstractNumId``, or None."""
    for abstract in numbering.abstract_num or ():
        if str(abstract.abstract_num_id) == str(abstract_id):
            return abstract
    return None


def _no_list(package: Any, num_id: Any) -> ContentError:
    """The error for a ``w:numId`` the document does not have, listing the ones it does."""
    emulator = emulator_of(package)
    ids = sorted(emulator.instance_list_definitions) if emulator is not None else []
    known = ", ".join(ids) if ids else "none"
    return ContentError(
        f"this document has no list with w:numId {num_id}",
        code="list.not_found",
        hint=f"the lists it has are: {known}; body.lists gives them as List views",
    )


# ---------------------------------------------------------------------------
# the views
# ---------------------------------------------------------------------------


class List:
    """A subset of Office JS ``Word.List`` over a ``w:num`` and its definition."""

    __slots__ = ("body", "id")

    def __init__(self, list_id: int, body: Body) -> None:
        """Build the view over a ``w:numId`` and the body its items are read in."""
        #: Office JS ``id``: the ``w:numId``.
        self.id = int(list_id)
        #: The body whose paragraphs this list's items are looked for in.
        self.body = body

    # -- identity ----------------------------------------------------------

    def __eq__(self, other: object) -> bool:
        """Two views of the same ``w:numId`` in the same body are equal."""
        return (
            isinstance(other, List)
            and other.id == self.id
            and other.body.container is self.body.container
        )

    def __hash__(self) -> int:
        """Hashes by the id and the body's container."""
        return hash((self.id, id(self.body.container)))

    def __repr__(self) -> str:
        """``<List 3 Number, 4 items>``."""
        kinds = self.level_types
        kind = kinds[0] if kinds else "?"
        return f"<List {self.id} {kind}, {len(self.paragraphs)} items>"

    @property
    def address(self) -> str:
        """``"list:3"``: what a :class:`ChangeReport` names this list. Extension."""
        return f"list:{self.id}"

    @property
    def package(self) -> Any:
        """The package the list's definition lives in."""
        return self.body.package

    @property
    def part(self) -> Any:
        """``/word/numbering.xml``, or None when the document has lost it."""
        return numbering_part_of(self.package)

    @property
    def definition(self) -> Any:
        """The emulator's :class:`ListNumberingDefinition` for this ``w:num``."""
        emulator = emulator_of(self.package)
        if emulator is None:
            return None
        return emulator.instance_list_definitions.get(str(self.id))

    # -- reading -----------------------------------------------------------

    @property
    def level_types(self) -> list[str]:
        """Office JS ``levelTypes``: ``"Bullet"`` or ``"Number"`` per level.

        Nine entries, one per ``w:ilvl``; a level the definition does not state
        is reported as ``"Number"``, which is what Word shows for one.
        """
        definition = self.definition
        out: list[str] = []
        for ilvl in range(LEVELS):
            level = None if definition is None else definition.level(ilvl)
            out.append("Bullet" if level is not None and level.is_bullet else "Number")
        return out

    @property
    def paragraphs(self) -> list[Paragraph]:
        """This list's items in the body, in document order."""
        return [item for item, _label in self._items()]

    def _items(self) -> list[tuple[Paragraph, ItemLabel]]:
        labels = labels_for(self.body)
        out: list[tuple[Paragraph, ItemLabel]] = []
        for paragraph in self.body.iter_paragraphs():
            label = labels.get(id(paragraph.element))
            if label is not None and label.num_id == str(self.id):
                out.append((paragraph, label))
        return out

    def level_exists(self, level: int) -> bool:
        """Office JS ``levelExists``: whether the definition states that level."""
        definition = self.definition
        return bool(definition is not None and definition.level_exists(int(level)))

    def get_level_paragraphs(self, level: int) -> list[Paragraph]:
        """Office JS ``getLevelParagraphs``: this list's items at one level."""
        return [item for item, label in self._items() if label.ilvl == int(level)]

    def get_level_string(self, level: int) -> str:
        """Office JS ``getLevelString``: the level's ``w:lvlText``.

        The pattern, not a counted label: ``"%1."``, ``"%1.%2."``, ``"%1)"``, or
        a bullet level's glyph. ``list_item.list_string`` is the counted one.
        """
        definition = self.definition
        found = None if definition is None else definition.level(int(level))
        return "" if found is None or found.level_text is None else found.level_text

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary: what a tool result says about a list."""
        return {
            "id": self.id,
            "address": self.address,
            "level_types": self.level_types,
            "levels": [level for level in range(LEVELS) if self.level_exists(level)],
            "paragraphs": [paragraph.address for paragraph in self.paragraphs],
        }

    # -- writing -----------------------------------------------------------

    def _writable_level(self, level: int, change: Any) -> Any:
        """The typed ``w:lvl`` to edit, the definition copied first if shared.

        Office JS's level setters change *this* list; a ``w:abstractNum`` that a
        second ``w:num`` points at is therefore copied and this ``w:num``
        repointed at the copy before anything is written (docx4j-core-ts CR-002
        section 3.9).
        """
        ilvl = int(level)
        if not 0 <= ilvl < LEVELS:
            raise ContentError(
                f"a list level is 0 to {LEVELS - 1}, not {level}",
                code="list.level_invalid",
                hint=f"pass a level between 0 and {LEVELS - 1}",
            )
        package = self.package
        part = numbering_part_of(package, create=False)
        if part is None:
            raise _no_list(package, self.id)
        numbering = part.contents
        num = _typed_num(numbering, self.id)
        if num is None:
            raise _no_list(package, self.id)
        change.part(part)

        abstract_id = _value(num.abstract_num_id)
        abstract = _typed_abstract(numbering, abstract_id)
        if abstract is None:
            raise ContentError(
                f"w:num {self.id} names w:abstractNum {abstract_id}, which is not there",
                code="list.definition_missing",
                hint="start a new list with paragraph.start_new_list() instead",
            )
        shared = sum(
            1 for other in numbering.num or () if str(_value(other.abstract_num_id)) == str(abstract_id)
        )
        if shared > 1:
            copy = deep_copy(abstract)
            copy.abstract_num_id = next_abstract_num_id(numbering)
            _append(numbering, "abstract_num", copy)
            num.abstract_num_id = el.abstractNumId(val=copy.abstract_num_id)
            num.abstract_num_id.parent = num
            abstract = copy

        found = next((lvl for lvl in (abstract.lvl or ()) if int(lvl.ilvl or 0) == ilvl), None)
        if found is None:
            found = el.lvl(ilvl=ilvl, lvl_jc=el.lvlJc(val="left"))
            found.p_pr = el.pPr(
                ind=el.ind(left=TWIPS_PER_LEVEL * (ilvl + 1), hanging=HANGING)
            )
            _append_level(abstract, found)
        invalidate(package)
        return found

    def set_level_numbering(
        self,
        level: int,
        numbering_type: str,
        format_string: str | None = None,
    ) -> None:
        """Office JS ``setLevelNumbering``: the level's ``w:numFmt`` and ``w:lvlText``.

        Args:
            level: ``w:ilvl``, 0 to 8.
            numbering_type: a ``Word.ListNumbering`` value --- ``"None"``,
                ``"Arabic"``, ``"UpperRoman"``, ``"LowerRoman"``,
                ``"UpperLetter"`` or ``"LowerLetter"``.
            format_string: the level text, ``"%1)"`` or ``"Article %1."``.
                Office JS takes an array of level substitutions here; this takes
                the ``w:lvlText`` itself, which is the same thing said once
                (CR-003 section 18). ``None`` leaves the level's own alone,
                writing ``"%<level+1>."`` when it has none.
        """
        num_fmt = NUMBERING.get(str(numbering_type))
        if num_fmt is None:
            raise ContentError(
                f"{numbering_type!r} is not a Word.ListNumbering value",
                code="list.numbering_invalid",
                hint=f"one of: {', '.join(NUMBERING)}",
            )
        with recording(self.body, "list.set_level_numbering") as change:
            lvl = self._writable_level(level, change)
            lvl.num_fmt = el.numFmt(val=num_fmt)
            lvl.num_fmt.parent = lvl
            text = format_string
            if text is None and lvl.lvl_text is None:
                text = f"%{int(level) + 1}."
            if text is not None:
                lvl.lvl_text = el.lvlText(val=text)
                lvl.lvl_text.parent = lvl
            # a numbered level carries no bullet font
            if lvl.r_pr is not None and getattr(lvl.r_pr, "r_fonts", None) is not None:
                lvl.r_pr.r_fonts = None
            change.touched(self.address)
            change.text(after=f"level {int(level)} = {numbering_type}")

    def set_level_bullet(
        self,
        level: int,
        list_bullet: str,
        char_code: int | None = None,
        font_name: str | None = None,
    ) -> None:
        """Office JS ``setLevelBullet``: the level's glyph and its font.

        Args:
            level: ``w:ilvl``, 0 to 8.
            list_bullet: a ``Word.ListBullet`` value --- ``"Solid"``,
                ``"Hollow"``, ``"Square"``, ``"Diamonds"``, ``"Arrow"``,
                ``"Checkmark"`` or ``"Custom"``.
            char_code: for ``"Custom"``, the character's code point.
            font_name: for ``"Custom"``, the font it is a glyph of.
        """
        kind = str(list_bullet)
        if kind == "Custom":
            if char_code is None:
                raise ContentError(
                    "a Custom bullet needs a char_code",
                    code="list.bullet_invalid",
                    hint="pass char_code=0xF0B7 and font_name='Symbol', or a named ListBullet",
                )
            glyph, font = chr(int(char_code)), font_name
        else:
            found = BULLETS.get(kind)
            if found is None:
                raise ContentError(
                    f"{list_bullet!r} is not a Word.ListBullet value",
                    code="list.bullet_invalid",
                    hint=f"one of: {', '.join(BULLETS)}, or 'Custom' with a char_code",
                )
            glyph, font = found
            font = font_name or font
        with recording(self.body, "list.set_level_bullet") as change:
            lvl = self._writable_level(level, change)
            lvl.num_fmt = el.numFmt(val="bullet")
            lvl.num_fmt.parent = lvl
            lvl.lvl_text = el.lvlText(val=glyph)
            lvl.lvl_text.parent = lvl
            if font:
                lvl.r_pr = el.rPr(r_fonts=el.rFonts(ascii=font, h_ansi=font, hint="default"))
                link_parents(lvl.r_pr)
                lvl.r_pr.parent = lvl
            change.touched(self.address)
            change.text(after=f"level {int(level)} = {kind}")

    def set_level_indents(
        self, level: int, text_indent: float, bullet_number_picker_indentation: float
    ) -> None:
        """Office JS ``setLevelIndents``: where the text and the marker sit.

        Both in **points**, as every other measurement in this API is:
        `text_indent` is the level's ``w:ind/@w:left`` (Office JS: "the same as
        paragraph indent") and `bullet_number_picker_indentation` its
        ``@w:hanging`` ("the same as paragraph hanging indent").
        """
        with recording(self.body, "list.set_level_indents") as change:
            lvl = self._writable_level(level, change)
            if lvl.p_pr is None:
                lvl.p_pr = el.pPr()
                lvl.p_pr.parent = lvl
            lvl.p_pr.ind = el.ind(
                left=round(float(text_indent) * TWIPS_PER_POINT),
                hanging=round(float(bullet_number_picker_indentation) * TWIPS_PER_POINT),
            )
            lvl.p_pr.ind.parent = lvl.p_pr
            change.touched(self.address)
            change.text(
                after=f"level {int(level)} indents {text_indent}pt / "
                f"{bullet_number_picker_indentation}pt"
            )

    def insert_paragraph(self, text: str, location: Literal["Start", "End"] = "End") -> Paragraph:
        """Office JS ``insertParagraph``: a new item at level 0, first or last.

        The new paragraph goes immediately before this list's first item or
        after its last, in the same story, and is attached to this list at level
        0 whatever level the item beside it is at.
        """
        where = str(location)
        if where not in ("Start", "End"):
            raise ContentError(
                f"a list paragraph goes at 'Start' or 'End', not {location!r}",
                code="list.location_invalid",
                hint="pass location='Start' or location='End'",
            )
        items = self.paragraphs
        if not items:
            raise ContentError(
                f"list {self.id} has no items in this body to insert beside",
                code="list.empty",
                hint="use body.insert_paragraph(...) then paragraph.attach_to_list(...)",
            )
        with recording(self.body, "list.insert_paragraph") as change:
            anchor = items[0] if where == "Start" else items[-1]
            new = anchor.insert_paragraph(text, location="Before" if where == "Start" else "After")
            new.attach_to_list(self.id, 0)
            change.touched(new)
            change.text(after=text)
            return new


class ListItem:
    """A subset of Office JS ``Word.ListItem``: what a paragraph in a list is."""

    __slots__ = ("paragraph",)

    def __init__(self, paragraph: Paragraph) -> None:
        """Build the view over a list paragraph."""
        #: The :class:`~docx4j_py.model.content.paragraph.Paragraph` this is of.
        self.paragraph = paragraph

    # -- identity ----------------------------------------------------------

    def __eq__(self, other: object) -> bool:
        """Two views of the same paragraph are equal."""
        return isinstance(other, ListItem) and other.paragraph == self.paragraph

    def __hash__(self) -> int:
        """Hashes by the paragraph."""
        return hash(self.paragraph)

    def __repr__(self) -> str:
        """``<ListItem '1.' level 0>``."""
        return f"<ListItem {self.list_string!r} level {self.level}>"

    @property
    def _label(self) -> ItemLabel | None:
        return labels_for(self.paragraph.parent_body).get(id(self.paragraph.element))

    # -- reading -----------------------------------------------------------

    @property
    def level(self) -> int:
        """Office JS ``level``: the ``w:ilvl`` this item is counted at."""
        label = self._label
        return 0 if label is None else label.ilvl

    @level.setter
    def level(self, value: int) -> None:
        ilvl = int(value)
        if not 0 <= ilvl < LEVELS:
            raise ContentError(
                f"a list level is 0 to {LEVELS - 1}, not {value}",
                code="list.level_invalid",
                hint=f"pass a level between 0 and {LEVELS - 1}",
            )
        paragraph = self.paragraph
        with recording(paragraph.parent_body, "list_item.level") as change:
            change.text(before=f"level={self.level}")
            # through _p_pr, so that tracking records a w:pPrChange (section 16.7)
            p_pr = paragraph._p_pr()
            num_pr = p_pr.num_pr
            if num_pr is None:
                num_pr = el.numPr()
                p_pr.num_pr = num_pr
                num_pr.parent = p_pr
            num_pr.ilvl = el.ilvl(val=ilvl)
            num_pr.ilvl.parent = num_pr
            change.touched(paragraph)
            change.text(after=f"level={ilvl}")

    @property
    def list_string(self) -> str:
        """Office JS ``listString``: the label Word paints in front of the item.

        ``"1."``, ``"a)"``, ``"2.1.4."``; a bullet's glyph as the Unicode
        character it looks like (:data:`BULLET_GLYPHS`), because Word's own
        ``\\uf0b7`` is a Symbol-font code point and not a bullet anywhere else;
        and ``""`` for a level whose ``w:numFmt`` is ``none``.
        """
        label = self._label
        if label is None:
            return ""
        text = label.result.num_string
        if label.result.is_bullet:
            return "".join(BULLET_GLYPHS.get(char, char) for char in text)
        return text

    @property
    def sibling_index(self) -> int:
        """Office JS ``siblingIndex``: 0-based place among the items at this level.

        From the emulator's counters, not a count of paragraphs: an item that
        Word numbers 3 under a parent that started at 1 has sibling index 2,
        and a ``w:startOverride`` moves both together.
        """
        label = self._label
        return 0 if label is None else label.result.sibling_index

    def get_ancestor(self, parent_only: bool = False) -> Paragraph | None:
        """Office JS ``getAncestor``: the item this one sits under.

        The nearest preceding item of the same list at a shallower level;
        with `parent_only`, only when that level is exactly one shallower.
        None where there is none (CR-003 section 3.1: no null objects).
        """
        ordered = _ordered_items(self.paragraph.parent_body)
        label = self._label
        if label is None:
            return None
        position = _position_of(ordered, self.paragraph.element)
        if position < 0:
            return None
        for other in reversed(ordered[:position]):
            if other.num_id != label.num_id:
                continue
            if other.ilvl < label.ilvl:
                if parent_only and other.ilvl != label.ilvl - 1:
                    return None
                return self.paragraph.parent_body.paragraph_for(other.element)
        return None

    def get_descendants(self, direct_children_only: bool = False) -> list[Paragraph]:
        """Office JS ``getDescendants``: the items nested under this one.

        Every following item of the same list at a deeper level, up to the next
        item at this level or shallower; with `direct_children_only`, only those
        exactly one level deeper.
        """
        ordered = _ordered_items(self.paragraph.parent_body)
        label = self._label
        if label is None:
            return []
        position = _position_of(ordered, self.paragraph.element)
        if position < 0:
            return []
        body = self.paragraph.parent_body
        out: list[Paragraph] = []
        for other in ordered[position + 1 :]:
            if other.num_id != label.num_id:
                continue
            if other.ilvl <= label.ilvl:
                break
            if direct_children_only and other.ilvl != label.ilvl + 1:
                continue
            out.append(body.paragraph_for(other.element))
        return out

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary: what a tool result says about a list item."""
        return {
            "level": self.level,
            "list_string": self.list_string,
            "sibling_index": self.sibling_index,
        }


def _ordered_items(body: Any) -> list[ItemLabel]:
    """Every list item of the body's story, in document order."""
    return list(labels_for(body).values())


def _position_of(ordered: list[ItemLabel], element: Any) -> int:
    for index, label in enumerate(ordered):
        if label.element is element:
            return index
    return -1


# ---------------------------------------------------------------------------
# what Paragraph and Body call
# ---------------------------------------------------------------------------


def lists_of(body: Body) -> list[List]:
    """Office JS ``Body.lists``: one :class:`List` per ``w:numId`` in this body.

    In the document order of each list's first item, each ``w:numId`` once.
    """
    labels = labels_for(body)
    out: list[List] = []
    seen: set[str] = set()
    for paragraph in body.iter_paragraphs():
        label = labels.get(id(paragraph.element))
        if label is None or label.num_id in seen:
            continue
        seen.add(label.num_id)
        out.append(List(int(label.num_id), body))
    return out


def list_of(paragraph: Paragraph) -> List | None:
    """Office JS ``Paragraph.list``: the list this paragraph is an item of."""
    label = labels_for(paragraph.parent_body).get(id(paragraph.element))
    if label is None:
        return None
    return List(int(label.num_id), paragraph.parent_body)


def list_item_of(paragraph: Paragraph) -> ListItem | None:
    """Office JS ``Paragraph.listItem``: None when the paragraph is not one."""
    label = labels_for(paragraph.parent_body).get(id(paragraph.element))
    return None if label is None else ListItem(paragraph)


def start_new_list(
    paragraph: Paragraph,
    *,
    kind: Literal["Number", "Bullet"] = "Number",
    like: List | None = None,
) -> List:
    """Office JS ``Paragraph.startNewList``: a new list, this paragraph its first item.

    A new ``w:abstractNum`` copied from docx4j's default numbering --- the
    decimal set for ``"Number"``, the bullet set for ``"Bullet"`` --- and a new
    ``w:num`` over it, with ``/word/numbering.xml`` created when the document
    has none. `like` is Word's *Restart numbering at 1*: a new ``w:num`` over
    **the same** definition as `like`, with a ``w:startOverride`` of 1 at level 0.

    Raises:
        ContentError: the paragraph is already a list item, as Office JS does.
    """
    if is_list_item(paragraph):
        raise ContentError(
            "this paragraph is already a list item",
            code="list.already_a_list_item",
            hint=(
                "call paragraph.detach_from_list() first, or "
                "paragraph.attach_to_list(list_id) to move it to another list"
            ),
        )
    body = paragraph.parent_body
    package = body.package
    touched: set[str] = set()
    with recording(body, "start_new_list") as change:
        if like is not None:
            num = restart_over(package, like.id, touched=touched)
        else:
            num = add_abstract_definition(package, kind, touched=touched)
        num_id = int(num.num_id)
        _set_num_pr(paragraph, num_id, 0)
        _apply_list_style(paragraph, touched)
        change.part(*sorted(touched))
        change.touched(paragraph, f"list:{num_id}")
        change.text(after=f"list {num_id}")
        return List(num_id, body)


def attach_to_list(paragraph: Paragraph, list_id: int, level: int = 0) -> None:
    """Office JS ``Paragraph.attachToList``: put this paragraph in an existing list.

    Raises:
        ContentError: the document has no such ``w:numId``; the message lists
            the ones it has.
    """
    package = paragraph.parent_body.package
    emulator = emulator_of(package)
    if emulator is None or str(list_id) not in emulator.instance_list_definitions:
        raise _no_list(package, list_id)
    ilvl = int(level)
    if not 0 <= ilvl < LEVELS:
        raise ContentError(
            f"a list level is 0 to {LEVELS - 1}, not {level}",
            code="list.level_invalid",
            hint=f"pass a level between 0 and {LEVELS - 1}",
        )
    touched: set[str] = set()
    with recording(paragraph.parent_body, "attach_to_list") as change:
        _set_num_pr(paragraph, int(list_id), ilvl)
        _apply_list_style(paragraph, touched)
        change.part(*sorted(touched))
        change.touched(paragraph, f"list:{int(list_id)}")
        change.text(after=f"list {int(list_id)} level {ilvl}")


def detach_from_list(paragraph: Paragraph) -> None:
    """Office JS ``Paragraph.detachFromList``: this paragraph is no longer an item.

    A ``w:numPr`` of the paragraph's own goes; numbering a **style** contributes
    is turned off the way Word turns it off, with ``w:numPr/w:numId w:val="0"``
    (ECMA-376 17.9.18: "no numbering"). The paragraph's style is left alone ---
    Word leaves *List Paragraph* on a paragraph whose numbering it removes.
    Not a list item: nothing happens.
    """
    package = paragraph.parent_body.package
    emulator = emulator_of(package)
    if emulator is None:
        return
    p_pr = paragraph.element.p_pr
    num_pr = getattr(p_pr, "num_pr", None) if p_pr is not None else None
    reference = _resolve(paragraph, emulator)
    if reference.not_numbered and num_pr is None:
        return
    with recording(paragraph.parent_body, "detach_from_list") as change:
        change.text(before=f"list {reference.num_id}" if reference.num_id else "not a list item")
        p_pr = paragraph._p_pr()
        if reference.direct or reference.not_numbered:
            p_pr.num_pr = None
        else:
            # the numbering comes from the style: Word writes w:numId 0
            p_pr.num_pr = el.numPr(num_id=el.numId(val=0))
            link_parents(p_pr.num_pr)
            p_pr.num_pr.parent = p_pr
        change.touched(paragraph)
        change.text(after="not a list item")


def is_list_item(paragraph: Paragraph) -> bool:
    """Office JS ``Paragraph.isListItem``: whether Word paints a label for it."""
    return id(paragraph.element) in labels_for(paragraph.parent_body)


def _resolve(paragraph: Paragraph, emulator: Emulator) -> Any:
    """Where this paragraph's numbering comes from, as a ``NumRef``."""
    p_pr = paragraph.element.p_pr
    p_style = _value(getattr(p_pr, "p_style", None)) if p_pr is not None else None
    num_pr = getattr(p_pr, "num_pr", None) if p_pr is not None else None
    num_id = _value(getattr(num_pr, "num_id", None)) if num_pr is not None else None
    ilvl = _value(getattr(num_pr, "ilvl", None)) if num_pr is not None else None
    num_id = None if num_id is None else str(num_id)
    ilvl = None if ilvl is None else str(ilvl)
    return emulator.resolve(p_style, num_id, ilvl, bool(num_id))


def _set_num_pr(paragraph: Paragraph, num_id: int, ilvl: int) -> None:
    """Write the paragraph's ``w:numPr``, through ``_p_pr`` so tracking sees it."""
    p_pr = paragraph._p_pr()
    num_pr = el.numPr(ilvl=el.ilvl(val=ilvl), num_id=el.numId(val=num_id))
    p_pr.num_pr = num_pr
    link_parents(num_pr)
    num_pr.parent = p_pr


def _apply_list_style(paragraph: Paragraph, touched: set[str]) -> None:
    """Word's *List Paragraph*, on a paragraph that has no style of its own.

    Word applies ``ListParagraph`` when the ribbon makes a list item, and the
    markdown importer writes it; a paragraph that already names a style ---
    a heading, a quote --- keeps it, which is also what Word does when numbering
    is applied to one (CR-003 section 18).
    """
    from docx4j_py.model.content.styles import ensure_style

    if paragraph.style_id not in ("", "Normal"):
        return
    package = paragraph.parent_body.package
    if package is None:
        return
    ensure_style(package, LIST_PARAGRAPH_STYLE, touched=touched)
    paragraph.element.p_pr.p_style = el.pStyle(val=LIST_PARAGRAPH_STYLE)
    paragraph.element.p_pr.p_style.parent = paragraph.element.p_pr


def _append_level(abstract: Any, lvl: Any) -> None:
    """Put a new ``w:lvl`` in the definition, in ``w:ilvl`` order."""
    levels = abstract.lvl
    if levels is None:
        levels = ChildList([], owner=abstract)
        abstract.lvl = levels
    position = len(levels)
    for index, existing in enumerate(levels):
        if int(existing.ilvl or 0) > int(lvl.ilvl or 0):
            position = index
            break
    levels.insert(position, lvl)
    link_parents(lvl)
    lvl.parent = abstract
