"""A paragraph's text as segments, and the block structure under a container.

This is the piece every view rests on. A paragraph's text is not a string in the
tree: it is spread over runs, and over the hyperlinks, content controls, smart
tags and tracked insertions those runs sit in. :func:`segments_of` flattens it
into one list of :class:`Segment`, each mapping a slice of the accepted-view
text to the item that carries it and to the lists that hold that item and its
run --- which is everything an edit needs to change the text in place, split a
run at an offset, or apply formatting to exactly one span.

What counts as text (docx4j ``TextUtils``, and the same rules as
:func:`docx4j_py.traversal.text_of`): ``w:t``, ``w:tab`` (a tab), ``w:br`` and
``w:cr`` (a newline), ``w:noBreakHyphen`` (U+2011), ``w:softHyphen`` (U+00AD)
and ``w:sym`` (the character its ``w:char`` names). ``w:delText``,
``w:instrText`` and ``w:delInstrText`` are not text --- except that a
``w:delText`` *is* the text of the original view, inside the ``w:del`` or
``w:moveFrom`` that removed it, which is the only place it is read.

What is descended into: ``w:hyperlink``, ``w:smartTag``, ``w:customXml``, a
run-level ``w:sdt``, ``w:fldSimple``, ``w:dir``, ``w:bdo``, and the revision
holder the view shows --- ``w:ins`` and ``w:moveTo`` in the accepted view,
``w:del`` and ``w:moveFrom`` in the original one. What is **not**:
``mc:AlternateContent``, a ``w:drawing`` and therefore a text box.
``Body.text`` is Word's main story (CR-003 sections 3.13 and decided question 8).

Offsets are **code points** (CR-003 decided question 3), and a split never falls
inside a grapheme cluster (section 3.12): :func:`is_grapheme_boundary` is
``unicodedata``'s combining classes plus the zero-width joiner, the variation
selectors, the emoji modifiers, regional-indicator pairs and the Indic
virama conjuncts, and the ``regex`` module's ``\\X`` is used for
:func:`grapheme_clusters` when it happens to be installed.
"""

from __future__ import annotations

import dataclasses
import re
import sys
import unicodedata
from typing import Any

from docx4j_py.child import iter_children
from docx4j_py.namespaces import WML_NS
from docx4j_py.traversal import element_name, run_items_of

__all__ = [
    "RUN_HOLDER_NAMES",
    "TEXT_ITEM_NAMES",
    "Segment",
    "block_children_of",
    "block_list_of",
    "find_all",
    "grapheme_clusters",
    "is_grapheme_boundary",
    "item_text",
    "runs_of",
    "search_pattern",
    "segments_of",
    "set_text",
    "snap_back",
    "snap_forward",
    "split_at",
    "text_of_view",
]


def _w(local: str) -> str:
    return f"{{{WML_NS}}}{local}"


W_P = _w("p")
W_R = _w("r")
W_T = _w("t")
W_SYM = _w("sym")
W_SDT = _w("sdt")
W_DEL_TEXT = _w("delText")
W_TBL = _w("tbl")

#: A run's children that contribute a constant string.
_CONSTANT_TEXT: dict[str, str] = {
    _w("tab"): "\t",
    _w("br"): "\n",
    _w("cr"): "\n",
    _w("noBreakHyphen"): "‑",
    _w("softHyphen"): "­",
}

#: Every run child that carries text. ``w:delText``, ``w:instrText`` and
#: ``w:delInstrText`` are deliberately absent: deleted text and field
#: instructions are not document text.
TEXT_ITEM_NAMES: frozenset[str] = frozenset({W_T, W_SYM, *_CONSTANT_TEXT})

#: Run-level containers whose items are read as part of the paragraph's text.
RUN_HOLDER_NAMES: frozenset[str] = frozenset(
    {
        _w("hyperlink"),
        _w("smartTag"),
        _w("customXml"),
        _w("fldSimple"),
        _w("dir"),
        _w("bdo"),
        W_SDT,
        _w("sdtContent"),
    }
)

#: The four run-level revision holders, and the view each one shows in.
_REVISIONS: dict[str, str] = {
    _w("ins"): "accepted",
    _w("moveTo"): "accepted",
    _w("del"): "original",
    _w("moveFrom"): "original",
}


def _sym_text(node: Any) -> str:
    """``w:sym`` as a character, from its hexadecimal ``w:char``."""
    for attribute in ("char_value", "char", "char_attribute"):
        value = getattr(node, attribute, None)
        if isinstance(value, str) and len(value) == 4:
            try:
                return chr(int(value, 16))
            except ValueError:
                return ""
    return ""


def item_text(qname: str | None, item: Any) -> str | None:
    """The text a run child contributes, or None when it contributes none."""
    if qname == W_T:
        value = getattr(item, "value", None)
        return value if isinstance(value, str) else ""
    if qname == W_SYM:
        return _sym_text(item)
    if qname is None:
        return None
    return _CONSTANT_TEXT.get(qname)


@dataclasses.dataclass(frozen=True, slots=True)
class Segment:
    """One text-bearing item of a run, with its place in the paragraph's text.

    Segments are computed fresh on every use --- nothing is cached, and an edit
    invalidates the indices --- so a caller that mutates the tree recomputes
    them before the next step.
    """

    item: Any
    """The item: a ``w:t``, ``w:tab``, ``w:br``, ``w:cr``, ``w:sym``, ..."""
    owner: list
    """The run's ``content`` list, which holds :attr:`item`."""
    index: int
    """The item's index in :attr:`owner`."""
    run: Any
    """The ``w:r``."""
    run_owner: list
    """The list holding the run: a paragraph's, a hyperlink's, an insertion's."""
    run_index: int
    """The run's index in :attr:`run_owner`."""
    text: str
    """The characters this item contributes."""
    start: int
    """The offset of the first character, in code points from the paragraph's start."""
    end: int
    """One past the last: ``start + len(text)``."""
    editable: bool
    """True for a ``w:t``: text that can be changed in place."""
    revision: str | None = None
    """``"ins"``, ``"del"``, ``"moveFrom"``, ``"moveTo"`` when the run is in one."""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view: the text, the offsets and the item's name."""
        return {
            "text": self.text,
            "start": self.start,
            "end": self.end,
            "item": element_name(self.item),
            "editable": self.editable,
            "revision": self.revision,
        }


def segments_of(container: Any, *, view: str = "accepted") -> list[Segment]:
    """The text segments of a paragraph (or any run holder), in order.

    Args:
        container: a ``w:p``, or anything :func:`run_items_of` knows.
        view: ``"accepted"`` (the default: ``w:ins`` and ``w:moveTo`` count,
            ``w:del`` and ``w:moveFrom`` do not) or ``"original"``, the other
            way round.

    Returns:
        One :class:`Segment` per text-bearing run child, in document order,
        with ``start``/``end`` running over the whole text.
    """
    out: list[Segment] = []
    position = 0

    def visit(items: list | None, revision: str | None) -> None:
        nonlocal position
        if not items:
            return
        for run_index, element in enumerate(items):
            name = element_name(element)
            if name == W_R:
                content = getattr(element, "content", None)
                if not isinstance(content, list):
                    continue
                for index, item in enumerate(content):
                    qname = element_name(item)
                    if qname == W_DEL_TEXT:
                        # deleted text is text in the original view only, and
                        # only inside the revision that deleted it
                        if revision not in ("del", "moveFrom"):
                            continue
                        value = getattr(item, "value", None)
                        text: str | None = value if isinstance(value, str) else ""
                    else:
                        text = item_text(qname, item)
                    if text is None:
                        continue
                    out.append(
                        Segment(
                            item=item,
                            owner=content,
                            index=index,
                            run=element,
                            run_owner=items,
                            run_index=run_index,
                            text=text,
                            start=position,
                            end=position + len(text),
                            editable=qname in (W_T, W_DEL_TEXT),
                            revision=revision,
                        )
                    )
                    position += len(text)
                continue
            shown = _REVISIONS.get(name or "")
            if shown is not None:
                if shown != view:
                    continue
                kind = (name or "").rpartition("}")[2]
                visit(run_items_of(element), kind)
            elif name in RUN_HOLDER_NAMES:
                visit(run_items_of(element), revision)

    visit(run_items_of(container), None)
    return out


def text_of_view(container: Any, *, view: str = "accepted") -> str:
    """The text of one run holder in one of the two views."""
    return "".join(segment.text for segment in segments_of(container, view=view))


def runs_of(container: Any, *, view: str = "accepted") -> list[Any]:
    """Every ``w:r`` of a run holder, nested ones included, in document order.

    Runs with no text are included --- a run holding only a ``w:drawing`` is a
    run --- which is why this is not derived from :func:`segments_of`.
    """
    out: list[Any] = []

    def visit(items: list | None) -> None:
        if not items:
            return
        for element in items:
            name = element_name(element)
            if name == W_R:
                out.append(element)
                continue
            shown = _REVISIONS.get(name or "")
            if shown is not None:
                if shown == view:
                    visit(run_items_of(element))
            elif name in RUN_HOLDER_NAMES:
                visit(run_items_of(element))

    visit(run_items_of(container))
    return out


# ---------------------------------------------------------------------------
# block structure
# ---------------------------------------------------------------------------

_sole_list_cache: dict[type, str | None] = {}


def _holds_blocks(cls: type) -> bool:
    """True when instances of `cls` are themselves containers of block content."""
    try:
        fields = dataclasses.fields(cls)
    except TypeError:
        return False
    for field in fields:
        if field.name != "content":
            continue
        factory = field.default_factory
        return factory is not dataclasses.MISSING and isinstance(factory, type)
    return False


def _sole_list_field(value: Any) -> str | None:
    """The one repeated element field of a class that has no ``content`` list.

    ``w:footnotes`` keeps its notes in ``footnote`` and ``w:comments`` its
    comments in ``comment``; both are a container of blocks all the same, so a
    ``Body`` over such a part works without naming either field here.

    The field's own items have to be containers of blocks --- a ``w:footnote``
    and a ``w:comment`` have a ``content`` list, a ``w:style`` does not --- so
    that ``w:styles``, which likewise keeps one repeated element, is correctly
    reported as holding no block content at all.
    """
    cls = value.__class__
    try:
        return _sole_list_cache[cls]
    except KeyError:
        pass
    found: str | None = None
    try:
        fields = dataclasses.fields(cls)
    except TypeError:
        fields = ()  # type: ignore[assignment]
    for field in fields:
        if field.metadata.get("type") not in (None, "Element"):
            continue
        factory = field.default_factory
        if factory is dataclasses.MISSING or not isinstance(factory, type):
            continue
        if not issubclass(factory, list):
            continue
        if found is not None:  # more than one: no answer
            found = None
            break
        if not all(_holds_blocks(item) for item in _field_types(field)):
            continue
        found = field.name
    _sole_list_cache[cls] = found
    return found


def _field_types(field: Any) -> list[type]:
    """The classes a repeated element field holds, from its type annotation."""
    import typing

    annotation = field.type
    if isinstance(annotation, str):
        try:
            module = sys.modules["docx4j_py.wml"]
            annotation = eval(annotation, vars(module))
        except Exception:  # noqa: BLE001 - an annotation we cannot resolve holds nothing
            return []
    args = typing.get_args(annotation)
    return [arg for arg in args if isinstance(arg, type)]


def block_list_of(value: Any) -> tuple[Any, str] | None:
    """The object and field name that hold a container's block-level list.

    The pair rather than the list itself, because a caller that inserts needs
    the owner for the parent pointers, and a caller that validates needs the
    class to ask the metadata about (:func:`docx4j_py.model.content.body`).
    """
    inner = getattr(value, "sdt_content", None)
    if inner is not None:
        value = inner
    if isinstance(getattr(value, "content", None), list):
        return value, "content"
    name = _sole_list_field(value)
    if name is not None and isinstance(getattr(value, name, None), list):
        return value, name
    return None


def block_children_of(value: Any) -> list | None:
    """The block-level list a container keeps, or None when it keeps none.

    A content control of any form answers with its ``w:sdtContent``'s list, so
    a traversal never has to know the four ``w:sdt`` classes apart (CR-003
    section 4: "the four ``sdtContent`` levels are skipped in paths"). A table
    answers with its rows, a row with its cells and a cell with its blocks,
    because in this model they are all the same ``content`` field --- which is
    what makes descending into a table one line rather than three helpers.
    """
    found = block_list_of(value)
    return None if found is None else getattr(found[0], found[1])


def child_qname(parent: Any, child: Any) -> str | None:
    """The element name `child` is written under inside `parent`."""
    for qname, item in iter_children(parent, mce="all"):
        if item is child:
            return qname
    return element_name(child)


# ---------------------------------------------------------------------------
# grapheme clusters (CR-003 section 3.12, decided question 10)
# ---------------------------------------------------------------------------

_ZWJ = "‍"
_MARKS = frozenset({"Mn", "Mc", "Me"})


def _is_variation_selector(ch: str) -> bool:
    code = ord(ch)
    return 0xFE00 <= code <= 0xFE0F or 0xE0100 <= code <= 0xE01EF


def _is_regional_indicator(ch: str) -> bool:
    return 0x1F1E6 <= ord(ch) <= 0x1F1FF


def _is_emoji_modifier(ch: str) -> bool:
    return 0x1F3FB <= ord(ch) <= 0x1F3FF


def _is_extend(ch: str) -> bool:
    return unicodedata.category(ch) in _MARKS or ch == _ZWJ or _is_variation_selector(ch)


def is_grapheme_boundary(text: str, index: int) -> bool:
    """True when `index` may be split without cutting a grapheme cluster in two.

    CR-003 section 3.12 and decided question 10: ``unicodedata``'s combining
    classes and general categories, plus the cases they do not cover ---

    * a combining mark, a variation selector or an emoji skin-tone modifier
      never begins a cluster;
    * the zero-width joiner binds both ways, which is what holds an emoji
      sequence such as U+1F468 ZWJ U+1F469 ZWJ U+1F467 together;
    * two regional indicators make one flag, so only every second one is a
      boundary;
    * a virama (canonical combining class 9) binds the consonant after it, so
      a Devanagari conjunct such as ``स्ते`` is one cluster;
    * CR LF is one cluster.

    ``0`` and ``len(text)`` are always boundaries.
    """
    if index <= 0 or index >= len(text):
        return True
    previous = text[index - 1]
    current = text[index]
    if previous == "\r" and current == "\n":
        return False
    if current == _ZWJ or previous == _ZWJ:
        return False
    if unicodedata.category(current) in _MARKS or unicodedata.combining(current):
        return False
    if _is_variation_selector(current) or _is_emoji_modifier(current):
        return False
    if _is_regional_indicator(current) and _is_regional_indicator(previous):
        run = 0
        position = index
        while position > 0 and _is_regional_indicator(text[position - 1]):
            run += 1
            position -= 1
        return run % 2 == 0
    # Indic conjunct: consonant, extenders, virama, extenders, consonant.
    position = index - 1
    while position >= 0 and _is_extend(text[position]):
        if unicodedata.combining(text[position]) == 9:
            return False
        position -= 1
    return True


def snap_back(text: str, index: int) -> int:
    """The largest grapheme boundary at or before `index`."""
    index = max(0, min(index, len(text)))
    while not is_grapheme_boundary(text, index):
        index -= 1
    return index


def snap_forward(text: str, index: int) -> int:
    """The smallest grapheme boundary at or after `index`."""
    index = max(0, min(index, len(text)))
    while not is_grapheme_boundary(text, index):
        index += 1
    return index


def grapheme_clusters(text: str) -> list[str]:
    """`text` split into grapheme clusters.

    The ``regex`` module's ``\\X`` when that module is installed --- it carries
    the whole of UAX #29 --- and :func:`is_grapheme_boundary` otherwise, which
    is CR-003 decided question 10: no dependency is added for this.
    """
    try:
        import regex  # type: ignore[import-not-found]
    except ImportError:
        out: list[str] = []
        start = 0
        for index in range(1, len(text) + 1):
            if is_grapheme_boundary(text, index):
                out.append(text[start:index])
                start = index
        return out
    return regex.findall(r"\X", text)


# ---------------------------------------------------------------------------
# splitting a run at an offset
# ---------------------------------------------------------------------------


def split_at(paragraph_element: Any, offset: int, *, prefer: str = "back") -> int:
    """Split the run at a text offset so that ``[offset, ...)`` begins a run.

    The offset is snapped to a grapheme boundary first (CR-003 section 3.12):
    ``prefer="back"`` to the start of the cluster it falls in, ``"forward"`` to
    the end of it. An offset that is already a run boundary splits nothing.

    Args:
        paragraph_element: the ``w:p``.
        offset: a code-point offset into the paragraph's accepted-view text.
        prefer: ``"back"`` or ``"forward"``, which way to snap.

    Returns:
        The offset actually used, which is the one a caller must go on with.
    """
    from docx4j_py.child import ChildList, deep_copy, link_parents
    from docx4j_py.wml import R, el

    text = text_of_view(paragraph_element)
    offset = snap_back(text, offset) if prefer == "back" else snap_forward(text, offset)

    segment = next(
        (s for s in segments_of(paragraph_element) if s.editable and s.start < offset < s.end),
        None,
    )
    if segment is None:
        return offset

    at = offset - segment.start
    set_text(segment.item, segment.text[:at])
    tail = segment.text[at:]
    rest = list(segment.owner[segment.index + 1 :])
    del segment.owner[segment.index + 1 :]
    r_pr = getattr(segment.run, "r_pr", None)
    second = R(content=ChildList([el.t(tail), *rest]))
    if r_pr is not None:
        second.r_pr = deep_copy(r_pr, second)
    segment.run_owner.insert(segment.run_index + 1, second)
    link_parents(second)
    second.parent = getattr(segment.run, "parent", None) or paragraph_element
    return offset


def set_text(item: Any, value: str) -> None:
    """Write a ``w:t``'s characters, with ``xml:space="preserve"`` where needed.

    Word writes the attribute whenever the value begins or ends with a space or
    holds two in a row; ``el.t`` does the same for a new one, and this keeps an
    edited one in step.
    """
    item.value = value
    if value != value.strip() or "  " in value:
        item.space = "preserve"


# ---------------------------------------------------------------------------
# search (Office JS Word.SearchOptions)
# ---------------------------------------------------------------------------


def search_pattern(
    text: str,
    *,
    match_case: bool = False,
    match_whole_word: bool = False,
    match_wildcards: bool = False,
) -> re.Pattern[str]:
    """A compiled pattern for a search text under Office JS's search options.

    Args:
        text: what to look for; a literal unless `match_wildcards` is on.
        match_case: ``False`` (the default) searches case-insensitively, which
            is ``re.IGNORECASE`` (CR-003 section 3.12).
        match_whole_word: the match must begin and end on a Unicode word
            boundary.
        match_wildcards: Word's wildcards --- ``?`` one character, ``*`` any
            run, ``[a-z]`` a class, ``[!x]`` a negated one, ``<`` and ``>`` the
            start and end of a word, ``@`` one or more of what precedes,
            ``{n,m}`` a repeat, and ``\\`` to escape any of them.
    """
    if match_wildcards:
        out: list[str] = []
        index = 0
        while index < len(text):
            char = text[index]
            if char == "?":
                out.append(".")
            elif char == "*":
                out.append(".*?")
            elif char == "<":
                out.append(r"\b(?=\w)")
            elif char == ">":
                out.append(r"\b(?<=\w)")
            elif char == "@":
                out.append("+")
            elif char == "[":
                close = text.find("]", index)
                if close < 0:
                    out.append(r"\[")
                else:
                    group = text[index + 1 : close]
                    if group.startswith("!"):
                        group = "^" + group[1:]
                    out.append(f"[{group.replace(chr(92), chr(92) * 2)}]")
                    index = close
            elif char == "{":
                close = text.find("}", index)
                if close < 0:
                    out.append(r"\{")
                else:
                    out.append("{" + text[index + 1 : close] + "}")
                    index = close
            elif char == "\\" and index + 1 < len(text):
                index += 1
                out.append(re.escape(text[index]))
            else:
                out.append(re.escape(char))
            index += 1
        source = "".join(out)
    else:
        source = re.escape(text)
    if match_whole_word:
        source = rf"(?<!\w){source}(?!\w)"
    return re.compile(source, 0 if match_case else re.IGNORECASE)


def find_all(text: str, pattern: re.Pattern[str]) -> list[tuple[int, int]]:
    """Every ``[start, end)`` match of a pattern, empty matches skipped."""
    out: list[tuple[int, int]] = []
    position = 0
    while position <= len(text):
        match = pattern.search(text, position)
        if match is None:
            break
        if match.end() == match.start():
            position = match.start() + 1
            continue
        out.append((match.start(), match.end()))
        position = match.end()
    return out
